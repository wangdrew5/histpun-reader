"""What every entry point shares: stdout setup, config, image selection, and
the progress files. No Excel and no Claude here."""
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent
IMAGE_EXTS = {".jpg", ".jpeg", ".png"}
NUM_RE = re.compile(r"\d+")


def configure_stdout():
    """Line-buffer stdout so progress shows live even when redirected to a
    file or pipe. Called from each main(), never at import. Skipped when
    stdout has been swapped for something that can't be reconfigured (a test
    capture, a closed stream)."""
    reconfigure = getattr(sys.stdout, "reconfigure", None)
    if reconfigure is None:
        return
    try:
        reconfigure(line_buffering=True)
    except (ValueError, OSError):
        pass


def now_iso():
    return datetime.now().isoformat(timespec="seconds")


def natural_key(path: Path):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", path.name)]


def extract_number(stem: str):
    m = NUM_RE.search(stem)
    return int(m.group()) if m else None


def load_config(path) -> dict:
    return json.loads(Path(path).read_text())


def add_selection_args(ap, default_config: str, only_example: str):
    ap.add_argument("--limit", type=int, default=None, help="Only process this many new images (for validation runs)")
    ap.add_argument("--only", default=None,
                    help=f"Process only this one filename (e.g. {only_example}), ignoring everything else pending")
    ap.add_argument("--config", default=str(SCRIPTS_DIR / default_config))


def list_images(folder: Path) -> list:
    return sorted((p for p in folder.iterdir() if p.suffix.lower() in IMAGE_EXTS), key=natural_key)


def select_images(images: list, is_done, only, limit, folder: Path, already_done_message: str) -> list:
    """--only / --limit selection over images not yet done. Exits with a
    message when --only names a missing or already-done image."""
    if only:
        match = next((p for p in images if p.name == only), None)
        if match is None:
            sys.exit(f"'{only}' was not found in {folder}")
        if is_done(match):
            sys.exit(already_done_message.format(name=only))
        return [match]
    todo = [p for p in images if not is_done(p)]
    return todo[:limit] if limit else todo


def load_progress(path: Path) -> dict:
    if path.is_file():
        return json.loads(path.read_text())
    return {}


def save_progress(path: Path, data: dict):
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False))
    tmp.replace(path)


def backup_workbook(workbook_path: Path, backup_dir: Path) -> Path:
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"{workbook_path.stem}_backup_{stamp}{workbook_path.suffix}"
    shutil.copy2(workbook_path, dest)
    return dest


def record_rate_limit(progress: dict, progress_path: Path, name: str, message: str, processed: int):
    """Shared stop-on-usage-limit handling for the two workbook pipelines."""
    reset_text = message.strip()
    print("SUBSCRIPTION LIMIT REACHED.")
    print(f"  {reset_text}")
    progress[name] = {
        "status": "pending",
        "note": f"stopped before processing: {reset_text}",
        "timestamp": now_iso(),
    }
    save_progress(progress_path, progress)
    print(f"\nStopped after {processed} image(s) this run. Progress saved to:\n  {progress_path}")
    print("Re-run the exact same command after the reset time above -- "
          "already-done images will be skipped automatically.")
