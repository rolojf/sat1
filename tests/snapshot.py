"""Record a baseline of every XML/PDF under facturas/.

Run before refiling so `test_conservation.py` can verify that no file is
dropped or duplicated when we move things around.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FACTURAS = ROOT / "facturas"
BASELINE_PATH = Path(__file__).resolve().parent / "baseline.json"


def iter_files(root: Path, exts: tuple[str, ...]) -> list[Path]:
    lowered = tuple(e.lower() for e in exts)
    return sorted(p for p in root.rglob("*") if p.is_file() and p.suffix.lower() in lowered)


def sha1(path: Path) -> str:
    h = hashlib.sha1()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def build_snapshot() -> dict:
    files = iter_files(FACTURAS, (".xml", ".pdf"))
    entries = []
    xml_count = pdf_count = 0
    for p in files:
        ext = p.suffix.lower()
        if ext == ".xml":
            xml_count += 1
        elif ext == ".pdf":
            pdf_count += 1
        entries.append({
            "path": str(p.relative_to(ROOT)),
            "ext": ext,
            "size": p.stat().st_size,
            "sha1": sha1(p),
        })
    return {
        "facturas_root": str(FACTURAS.relative_to(ROOT)),
        "xml_count": xml_count,
        "pdf_count": pdf_count,
        "total": len(entries),
        "files": entries,
    }


def main() -> None:
    snapshot = build_snapshot()
    BASELINE_PATH.write_text(json.dumps(snapshot, indent=2, sort_keys=False))
    print(
        f"baseline written: {snapshot['xml_count']} XML, "
        f"{snapshot['pdf_count']} PDF, {snapshot['total']} total -> {BASELINE_PATH}"
    )


if __name__ == "__main__":
    main()
