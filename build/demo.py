"""Rebuild demo/northfork-precision from the test kit. Everything in it is invented."""
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests" / "kit"))
import rolling  # noqa: E402

OUT = ROOT / "demo" / "northfork-precision"

if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as t:
        rolling.build(date(2026, 11, 25), Path(t) / "plant")
        for part in ("plant", "linked"):
            shutil.rmtree(OUT / part, ignore_errors=True)
            shutil.copytree(Path(t) / part, OUT / part)
