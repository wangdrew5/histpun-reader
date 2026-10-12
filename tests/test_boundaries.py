"""Module boundaries from the refactor: no xlwings or stdout changes at import,
no script importing another script, launchers still answer --help, and the
shared image selection behaves like the per-script code it replaced."""
import ast
import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

from common import configure_stdout, list_images, select_images  # noqa: E402

RUNNABLE = {"process_reports", "process_complexion", "export_complexion"}


def run_python(code: str) -> str:
    out = subprocess.run([sys.executable, "-c", code], cwd=SCRIPTS, capture_output=True,
                         text=True, timeout=60)
    if out.returncode != 0:
        raise AssertionError(out.stderr)
    return out.stdout.strip()


class ImportSideEffectsTest(unittest.TestCase):
    def test_no_module_imports_xlwings_at_import_time(self):
        for module in ["export_complexion", "process_complexion", "process_reports",
                       "excel_sheet", "claude_cli", "common", "prompts"]:
            with self.subTest(module=module):
                self.assertEqual(
                    run_python(f"import sys, {module}; print('xlwings' in sys.modules)"), "False")

    def test_import_leaves_stdout_alone(self):
        # stdout is a pipe here, so it starts out not line-buffered.
        self.assertEqual(run_python(
            "import sys; before = sys.stdout.line_buffering\n"
            "import process_reports, process_complexion, export_complexion\n"
            "print(before, sys.stdout.line_buffering)"), "False False")

    def test_configure_stdout_line_buffers_a_redirected_stream(self):
        self.assertEqual(run_python(
            "import sys, common; common.configure_stdout(); print(sys.stdout.line_buffering)"), "True")

    def test_configure_stdout_tolerates_a_replaced_stream(self):
        real = sys.stdout
        try:
            sys.stdout = io.StringIO()  # no reconfigure()
            configure_stdout()
        finally:
            sys.stdout = real


class LayoutTest(unittest.TestCase):
    def test_nothing_imports_a_runnable_script(self):
        for path in SCRIPTS.glob("*.py"):
            tree = ast.parse(path.read_text())
            imported = set()
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    imported |= {a.name.split(".")[0] for a in node.names}
                elif isinstance(node, ast.ImportFrom) and node.module:
                    imported.add(node.module.split(".")[0])
            with self.subTest(file=path.name):
                self.assertEqual(imported & RUNNABLE, set())

    def test_launchers_answer_help(self):
        expected = {
            "run_batch.sh": ["--limit", "--only", "--config"],
            "run_complexion_batch.sh": ["--limit", "--only", "--config"],
            "run_complexion_export.sh": ["--limit", "--only", "--config", "--out"],
        }
        for launcher, flags in expected.items():
            with self.subTest(launcher=launcher):
                out = subprocess.run([str(ROOT / launcher), "--help"], capture_output=True,
                                     text=True, timeout=60)
                self.assertEqual(out.returncode, 0, out.stderr)
                for flag in flags:
                    self.assertIn(flag, out.stdout)


class ImageSelectionTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)
        for name in ["10.jpg", "9.JPG", "100.png", "notes.txt", "2.jpeg", ".DS_Store"]:
            (self.folder / name).write_bytes(b"")
        self.images = list_images(self.folder)

    def tearDown(self):
        self.tmp.cleanup()

    def select(self, done=(), only=None, limit=None):
        return [p.name for p in select_images(self.images, lambda p: p.name in done, only, limit,
                                              self.folder, "'{name}' already done")]

    def test_natural_order_and_image_extensions_only(self):
        self.assertEqual([p.name for p in self.images], ["2.jpeg", "9.JPG", "10.jpg", "100.png"])

    def test_skips_done_and_applies_limit(self):
        self.assertEqual(self.select(done={"9.JPG"}), ["2.jpeg", "10.jpg", "100.png"])
        self.assertEqual(self.select(done={"9.JPG"}, limit=2), ["2.jpeg", "10.jpg"])

    def test_only(self):
        self.assertEqual(self.select(only="10.jpg", limit=0), ["10.jpg"])
        with self.assertRaises(SystemExit) as cm:
            self.select(only="11.jpg")
        self.assertIn("'11.jpg' was not found", str(cm.exception.code))
        with self.assertRaises(SystemExit) as cm:
            self.select(done={"10.jpg"}, only="10.jpg")
        self.assertEqual(cm.exception.code, "'10.jpg' already done")


if __name__ == "__main__":
    unittest.main()
