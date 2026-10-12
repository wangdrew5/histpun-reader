"""Talking to Claude: image downscaling plus one streaming `claude -p` call.

No Excel here -- the standalone export uses this module too.
"""
import hashlib
import json
import os
import queue
import re
import signal
import subprocess
import threading
import time
from pathlib import Path

from PIL import Image

DOWNSCALE_LONG_EDGE = 2200  # full-res scans (6000px+) don't add legibility, only latency
CLI_TIMEOUT_SECONDS = 600  # hard deadline for one whole CLI call, start to exit
STDERR_TAIL_CHARS = 20000

RATE_LIMIT_RE = re.compile(
    r"usage limit|rate.?limit|quota exceeded|too many requests|429|"
    r"overloaded|please try again later|limit reached|exceeded your",
    re.IGNORECASE,
)

ROW_START_RE = re.compile(r'"row_number"\s*:\s*\d+')

RESIZE_CACHE_DIR = Path(__file__).parent / ".resize_cache"


def prepare_image_for_reading(image_path: Path) -> Path:
    """Downscales huge scans before sending to Claude. Full 6000px+ originals
    add real minutes of vision-processing latency without adding legibility --
    verified readable by eye at ~2200px on the long edge."""
    try:
        with Image.open(image_path) as im:
            w, h = im.size
            if max(w, h) <= DOWNSCALE_LONG_EDGE:
                return image_path
            scale = DOWNSCALE_LONG_EDGE / max(w, h)
            new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
            resized = im.convert("RGB").resize(new_size, Image.LANCZOS)
            RESIZE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
            # Keyed by the full source path (hashed), not just the filename stem --
            # different image folders can reuse the same numbers (e.g. two
            # independent batches both containing a "45.jpg"), and a stem-only
            # key would silently serve one pipeline's cached resize to the other.
            digest = hashlib.sha1(str(image_path.resolve()).encode()).hexdigest()[:16]
            out_path = RESIZE_CACHE_DIR / f"{image_path.stem}_{digest}.jpg"
            resized.save(out_path, quality=88)
            return out_path
    except Exception:
        return image_path  # fall back to the original if downscaling fails for any reason


def build_cli_cmd(prompt: str, schema: dict, add_dir: Path) -> list:
    return [
        "claude", "-p", prompt,
        "--output-format", "stream-json",
        "--include-partial-messages",
        "--verbose",
        "--json-schema", json.dumps(schema),
        "--allowedTools", "Read",
        "--add-dir", str(add_dir), str(RESIZE_CACHE_DIR),
        "--no-session-persistence",
        "--strict-mcp-config",       # skip unrelated project MCP servers (Gmail/Drive/etc.) -- pure latency for this task
        "--disable-slash-commands",  # skip skill loading -- not needed for a fixed extraction prompt
        "--effort", "low",           # straightforward transcription; low effort cut latency ~5x in testing with no accuracy loss observed
        "--exclude-dynamic-system-prompt-sections",  # keeps the cached system-prompt prefix byte-identical across calls -- more cache hits, zero behavior change
    ]


def _pump(stream, sink):
    """Reader-thread body: forward every line, then None at EOF."""
    try:
        for line in stream:
            sink(line)
    except (OSError, ValueError):
        pass  # pipe closed underneath us during cleanup
    finally:
        sink(None)


def _kill_process_tree(proc):
    """The CLI runs in its own process group (start_new_session) so it and any
    helpers it spawned can be stopped together, and so a terminal Ctrl+C
    reaches only Python -- which then cleans up here, deliberately."""
    if proc.poll() is not None:
        return
    for sig, grace in ((signal.SIGTERM, 5), (signal.SIGKILL, 5)):
        try:
            os.killpg(proc.pid, sig)
        except (ProcessLookupError, PermissionError):
            if proc.poll() is None:
                proc.kill()
        try:
            proc.wait(timeout=grace)
            return
        except subprocess.TimeoutExpired:
            continue


def run_cli_stream(cmd: list, timeout: float, on_line) -> dict:
    """Runs cmd, handing each stdout line to on_line(line) as it arrives.

    One deadline covers the whole call -- launch, streaming, and exit -- so a
    CLI that stalls mid-stream can't hang the batch. stderr is drained on its
    own thread so a chatty CLI can't fill the pipe and deadlock. On timeout,
    an exception from on_line, or Ctrl+C, the child's whole process group is
    terminated and reaped before returning/raising.

    Returns {"timed_out", "returncode", "stdout_tail", "stderr"}.
    """
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, stdin=subprocess.DEVNULL,
            text=True, bufsize=1, start_new_session=True,
        )
    except OSError as exc:
        raise RuntimeError(f"Could not launch claude CLI: {exc}") from exc

    deadline = time.monotonic() + timeout
    lines = queue.Queue()
    stderr_chunks = []

    def keep_stderr(line):
        if line is not None:
            stderr_chunks.append(line)
            if len(stderr_chunks) > 2000:
                del stderr_chunks[:1000]

    threads = [
        threading.Thread(target=_pump, args=(proc.stdout, lines.put), daemon=True),
        threading.Thread(target=_pump, args=(proc.stderr, keep_stderr), daemon=True),
    ]
    for t in threads:
        t.start()

    stdout_tail = []
    timed_out = False
    try:
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                timed_out = True
                break
            try:
                line = lines.get(timeout=min(remaining, 1.0))
            except queue.Empty:
                continue
            if line is None:
                break  # stdout closed
            stdout_tail.append(line)
            del stdout_tail[:-200]
            on_line(line)
        if not timed_out:
            try:
                proc.wait(timeout=max(0.0, deadline - time.monotonic()))
            except subprocess.TimeoutExpired:
                timed_out = True
    finally:
        # Runs on normal exit, timeout, an on_line error, and KeyboardInterrupt
        _kill_process_tree(proc)
        for t in threads:
            t.join(timeout=5)
        for stream in (proc.stdout, proc.stderr):
            try:
                stream.close()
            except Exception:
                pass

    return {
        "timed_out": timed_out,
        "returncode": proc.returncode,
        "stdout_tail": stdout_tail,
        "stderr": "".join(stderr_chunks)[-STDERR_TAIL_CHARS:],
    }


def call_claude_vision(image_path: Path, add_dir: Path, schema: dict, prompt_template: str,
                       on_row=None, timeout: float = CLI_TIMEOUT_SECONDS) -> dict:
    """Runs the CLI in streaming mode and calls on_row(n) each time a new
    row starts arriving in the model's output, so the caller can print live
    per-row progress instead of waiting silently for the whole image."""
    read_path = prepare_image_for_reading(image_path)
    prompt = prompt_template.format(image_path=read_path)
    cmd = build_cli_cmd(prompt, schema, add_dir)

    state = {"buffer": "", "rows_seen": 0, "result": None}

    def on_line(line):
        line = line.strip()
        if not line:
            return
        try:
            evt = json.loads(line)
        except json.JSONDecodeError:
            return

        if evt.get("type") == "result":
            state["result"] = evt
            return

        if evt.get("type") != "stream_event":
            return
        ev = evt.get("event", {})
        ev_type = ev.get("type")
        if ev_type == "content_block_start":
            state["buffer"] = ""
        elif ev_type == "content_block_delta":
            delta = ev.get("delta", {})
            if delta.get("type") == "input_json_delta":
                state["buffer"] += delta.get("partial_json") or ""
                matches = len(ROW_START_RE.findall(state["buffer"]))
                if matches > state["rows_seen"]:
                    state["rows_seen"] = matches
                    if on_row:
                        on_row(matches)

    try:
        run = run_cli_stream(cmd, timeout, on_line)
    except RuntimeError as exc:
        return {"kind": "error", "message": str(exc)}
    except Exception as exc:
        return {"kind": "error", "message": f"Error reading CLI stream: {exc}"}

    if run["timed_out"]:
        return {"kind": "error", "message": f"CLI call timed out after {timeout:g}s (process stopped)"}

    result_envelope = state["result"]
    stderr_text = run["stderr"]
    combined = "".join(run["stdout_tail"]) + "\n" + stderr_text

    returncode = run["returncode"]

    # A rate/usage limit can surface either as a captured "result" stream event
    # with is_error=true, or as a nonzero process exit with no clean envelope.
    # Prefer the CLI's own human-readable "result" text (e.g. "You've hit your
    # session limit · resets 8:40pm (America/Los_Angeles)") whenever we have it,
    # regardless of which path triggered, and only fall back to a raw dump when
    # we truly have nothing structured to work with.
    if result_envelope is not None and result_envelope.get("is_error"):
        dumped = json.dumps(result_envelope)[:2000]
        if RATE_LIMIT_RE.search(dumped):
            reset_msg = result_envelope.get("result")
            return {"kind": "rate_limit", "message": reset_msg if isinstance(reset_msg, str) and reset_msg else dumped}
        # Lead with the CLI's own human-readable reason (e.g. "Not logged in"),
        # which otherwise gets cut off behind the usage/token metadata.
        reason = result_envelope.get("result") or result_envelope.get("errors")
        return {"kind": "error", "message": f"{reason} | {dumped}" if reason else dumped}

    if returncode != 0:
        if RATE_LIMIT_RE.search(combined):
            return {"kind": "rate_limit", "message": combined.strip()[-2000:]}
        return {"kind": "error", "message": combined.strip()[-2000:]}

    if result_envelope is None:
        return {"kind": "error", "message": "No result event received: " + combined.strip()[-2000:]}

    structured = result_envelope.get("structured_output")
    if structured is None:
        try:
            structured = json.loads(result_envelope.get("result", ""))
        except (json.JSONDecodeError, TypeError):
            return {"kind": "error", "message": "No structured_output and 'result' was not parseable JSON"}

    rows = structured.get("rows") if isinstance(structured, dict) else None
    if not isinstance(rows, list):
        return {"kind": "error", "message": "Schema response missing a 'rows' list"}

    return {"kind": "ok", "rows": rows, "cost_usd": result_envelope.get("total_cost_usd")}
