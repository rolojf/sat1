"""No-loss invariant: every file in the baseline still exists after refiling.

Run `python tests/snapshot.py` first to record a baseline. Then run pytest
before and after `clasifica.py --apply`; both runs must stay green.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
FACTURAS = ROOT / "facturas"
BASELINE_PATH = Path(__file__).resolve().parent / "baseline.json"


def _load_baseline() -> dict:
    if not BASELINE_PATH.exists():
        pytest.skip(
            f"no baseline at {BASELINE_PATH}; run `uv run python tests/snapshot.py` first"
        )
    return json.loads(BASELINE_PATH.read_text())


def _iter_current(exts: tuple[str, ...]) -> list[Path]:
    lowered = tuple(e.lower() for e in exts)
    return [
        p for p in FACTURAS.rglob("*")
        if p.is_file() and p.suffix.lower() in lowered
    ]


def _sha1(path: Path) -> str:
    h = hashlib.sha1()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


@pytest.fixture(scope="module")
def baseline() -> dict:
    return _load_baseline()


@pytest.fixture(scope="module")
def current_hashes() -> set[str]:
    return {_sha1(p) for p in _iter_current((".xml", ".pdf"))}


def test_xml_count_unchanged(baseline: dict) -> None:
    current = sum(1 for _ in _iter_current((".xml",)))
    assert current == baseline["xml_count"], (
        f"XML count drifted: baseline={baseline['xml_count']} now={current}"
    )


def test_pdf_count_unchanged(baseline: dict) -> None:
    current = sum(1 for _ in _iter_current((".pdf",)))
    assert current == baseline["pdf_count"], (
        f"PDF count drifted: baseline={baseline['pdf_count']} now={current}"
    )


def test_every_baseline_file_still_present(baseline: dict, current_hashes: set[str]) -> None:
    """Every baseline sha1 must still exist somewhere under facturas/."""
    missing = []
    for entry in baseline["files"]:
        if entry["sha1"] not in current_hashes:
            missing.append(entry["path"])
    assert not missing, f"{len(missing)} file(s) lost: {missing[:5]}{'…' if len(missing) > 5 else ''}"


def test_no_unexpected_new_files(baseline: dict, current_hashes: set[str]) -> None:
    """No new XML/PDF appeared (would mean an accidental copy / corruption)."""
    baseline_hashes = {e["sha1"] for e in baseline["files"]}
    extra = current_hashes - baseline_hashes
    assert not extra, f"{len(extra)} unexpected new file hash(es) appeared"
