"""Build dist/plant-watch.zip, the file people upload in the Claude desktop app."""
import json
import subprocess
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
INCLUDE = ("plugin", "demo", "README.md", "SECURITY.md")
STAMP = (2026, 12, 1, 0, 0, 0)           # fixed, so the same source gives the same file


def files():
    listed = subprocess.run(["git", "ls-files", "-z", "--", *INCLUDE], cwd=ROOT, capture_output=True, check=True).stdout
    return sorted(x for x in listed.decode("utf-8").split("\0") if x)


def build(out):
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in files():
            inside = name[len("plugin/"):] if name.startswith("plugin/") else name
            if name == "plugin/README.md":         # the download carries the full one
                continue
            info = zipfile.ZipInfo(inside, STAMP)
            info.compress_type, info.external_attr = zipfile.ZIP_DEFLATED, 0o644 << 16
            z.writestr(info, (ROOT / name).read_bytes())
    return out


if __name__ == "__main__":
    path = build(ROOT / "dist" / "plant-watch.zip")
    version = json.loads((ROOT / "plugin" / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))["version"]
    print(f"{path} ({path.stat().st_size:,} bytes, version {version})")
