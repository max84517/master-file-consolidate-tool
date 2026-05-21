"""Paths configuration — resolves project root regardless of frozen/dev mode."""
import sys
from pathlib import Path


def get_project_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[3]


PROJECT_ROOT = get_project_root()
DATA_DIR = PROJECT_ROOT / "data"
SOURCE_DIR = DATA_DIR / "source"
RESULT_BY_SEGMENT_DIR = DATA_DIR / "result_by_segment"
CONSOLIDATE_ALL_DIR = DATA_DIR / "consolidate_all"

# Ensure dirs exist
for _d in [
    SOURCE_DIR / "NB",
    SOURCE_DIR / "DT",
    SOURCE_DIR / "Peripheral",
    RESULT_BY_SEGMENT_DIR / "NB",
    RESULT_BY_SEGMENT_DIR / "DT",
    RESULT_BY_SEGMENT_DIR / "Peripheral",
    CONSOLIDATE_ALL_DIR,
]:
    _d.mkdir(parents=True, exist_ok=True)
