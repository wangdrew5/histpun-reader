"""Issue 4: whole-call deadline, stderr draining, and child cleanup.

Every "CLI" here is a tiny Python child process standing in for `claude`;
nothing contacts Claude or uses any usage.
"""
import json
import os
import sys
import textwrap
import time
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

import claude_cli  # noqa: E402
from claude_cli import call_claude_vision, run_cli_stream  # noqa: E402


def fake_cli(code: str) -> list:
    return [sys.executable, "-c", textwrap.dedent(code)]


def gone(pid: int, within: float = 5.0) -> bool:
    end = time.monotonic() + within
    while time.monotonic() < end:
        try:
            os.kill(pid, 0)
        except ProcessLookupError:
            return True
        except PermissionError:
            return False
        try:  # reap it if it's our own zombie child
            os.waitpid(pid, os.WNOHANG)
        except ChildProcessError:
            pass
        time.sleep(0.05)
    return False


class RunCliStreamTest(unittest.TestCase):
    def test_silent_hang_times_out_and_kills_child(self):
        lines = []
        start = time.monotonic()
        run = run_cli_stream(fake_cli("""
            import os, time
            print(os.getpid(), flush=True)
            time.sleep(60)
        """), timeout=1.5, on_line=lines.append)
        self.assertTrue(run["timed_out"])
        self.assertLess(time.monotonic() - start, 15)
        self.assertTrue(gone(int(lines[0])))

    def test_stall_mid_stream_times_out(self):
        # The old code only timed out *after* stdout closed; a stall while
        # streaming blocked forever.
        lines = []
        run = run_cli_stream(fake_cli("""
            import time
            for i in range(3):
                print('{"type": "noise"}', flush=True)
            time.sleep(60)
        """), timeout=1.5, on_line=lines.append)
        self.assertTrue(run["timed_out"])
        self.assertEqual(len(lines), 3)

    def test_lingering_after_closing_stdout_times_out(self):
        run = run_cli_stream(fake_cli("""
            import os, sys, time
            os.close(1)
            time.sleep(60)
        """), timeout=1.5, on_line=lambda _l: None)
        self.assertTrue(run["timed_out"])

    def test_heavy_stderr_does_not_deadlock(self):
        # 2 MB on stderr before any stdout: far beyond a pipe buffer, so an
        # undrained stderr pipe would block the child forever.
        lines = []
        run = run_cli_stream(fake_cli("""
            import sys
            sys.stderr.write("e" * 2_000_000 + "TAIL")
            sys.stderr.flush()
            print("done", flush=True)
        """), timeout=20, on_line=lines.append)
        self.assertFalse(run["timed_out"])
        self.assertEqual(run["returncode"], 0)
        self.assertEqual([l.strip() for l in lines], ["done"])
        self.assertTrue(run["stderr"].endswith("TAIL"))
        self.assertLessEqual(len(run["stderr"]), claude_cli.STDERR_TAIL_CHARS)

    def test_grandchildren_are_stopped_on_timeout(self):
        lines = []
        run_cli_stream(fake_cli("""
            import subprocess, sys, time
            p = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(60)"])
            print(p.pid, flush=True)
            time.sleep(60)
        """), timeout=1.5, on_line=lines.append)
        self.assertTrue(gone(int(lines[0])))

    def test_ctrl_c_cleans_up_child_and_propagates(self):
        seen = []

        def on_line(line):
            seen.append(int(line))
            raise KeyboardInterrupt

        with self.assertRaises(KeyboardInterrupt):
            run_cli_stream(fake_cli("""
                import os, time
                print(os.getpid(), flush=True)
                time.sleep(60)
            """), timeout=30, on_line=on_line)
        self.assertTrue(gone(seen[0]))

    def test_launch_failure_raises(self):
        with self.assertRaises(RuntimeError):
            run_cli_stream(["/definitely/not/a/real/binary"], timeout=5, on_line=lambda _l: None)


def stream_script(events, exit_code=0, stderr=""):
    return f"""
        import json, sys
        for e in {events!r}:
            print(json.dumps(e), flush=True)
        sys.stderr.write({stderr!r})
        sys.exit({exit_code})
    """


def delta(text):
    return {"type": "stream_event", "event": {"type": "content_block_delta",
                                              "delta": {"type": "input_json_delta", "partial_json": text}}}


class CallClaudeVisionTest(unittest.TestCase):
    """End to end through call_claude_vision with the command swapped out."""

    def call(self, code, timeout=20):
        rows_seen = []
        with mock.patch.object(claude_cli, "build_cli_cmd", return_value=fake_cli(code)):
            result = call_claude_vision(Path("no-such-image.jpg"), Path("."), {}, "{image_path}",
                                        on_row=rows_seen.append, timeout=timeout)
        return result, rows_seen

    def test_ok_result_with_live_row_progress(self):
        rows = [{"row_number": 1, "complexion": "Black"}, {"row_number": 2, "complexion": "White"}]
        events = [
            {"type": "stream_event", "event": {"type": "content_block_start"}},
            delta('{"rows": [{"row_number": 1, '),
            delta('"complexion": "Black"}, {"row_number": 2, "complexion": "White"}]}'),
            {"type": "result", "is_error": False, "structured_output": {"rows": rows}},
        ]
        result, seen = self.call(stream_script(events))
        self.assertEqual(result["kind"], "ok")
        self.assertEqual(result["rows"], rows)
        self.assertEqual(seen, [1, 2])

    def test_timeout_is_reported_as_error(self):
        result, _ = self.call("import time; time.sleep(60)", timeout=1.5)
        self.assertEqual(result["kind"], "error")
        self.assertIn("timed out", result["message"])

    def test_rate_limit_from_stderr_and_nonzero_exit(self):
        result, _ = self.call(stream_script([], exit_code=1, stderr="Error: usage limit reached"))
        self.assertEqual(result["kind"], "rate_limit")

    def test_error_envelope_leads_with_reason(self):
        events = [{"type": "result", "is_error": True, "result": "Not logged in"}]
        result, _ = self.call(stream_script(events, exit_code=1))
        self.assertEqual(result["kind"], "error")
        self.assertTrue(result["message"].startswith("Not logged in"))

    def test_on_row_error_becomes_error_result_and_stops_child(self):
        events = [delta('{"rows": [{"row_number": 1')]
        code = stream_script(events).replace("sys.exit(0)", "import time; time.sleep(60)")
        with mock.patch.object(claude_cli, "build_cli_cmd", return_value=fake_cli(code)):
            start = time.monotonic()
            result = call_claude_vision(Path("x.jpg"), Path("."), {}, "{image_path}",
                                        on_row=lambda n: 1 / 0, timeout=30)
        self.assertEqual(result["kind"], "error")
        self.assertIn("Error reading CLI stream", result["message"])
        self.assertLess(time.monotonic() - start, 15)


if __name__ == "__main__":
    unittest.main()
