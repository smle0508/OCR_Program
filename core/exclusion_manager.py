# core/exclusion_manager.py
"""
ExclusionManager (단순화 버전)
────────────────────────────
• exclude_strings  : Set[str]
• JSON 직렬화 only  {"exclude": ["텍스트1", "문구2", ...]}
"""

from __future__ import annotations
import json, threading
from pathlib import Path
from typing import Set, List

APP_DIR   = Path.home() / "AppData" / "Roaming" / "PdfOcrExcel"
APP_DIR.mkdir(parents=True, exist_ok=True)
DATA_PATH = APP_DIR / "exclusions.json"


class ExclusionManager:
    _lock = threading.RLock()

    def __init__(self) -> None:
        self.exclude_strings: Set[str] = set()
        self._load()

    # ─────────── API ───────────
    def list_all(self) -> List[str]:
        return sorted(self.exclude_strings)

    def add_many(self, items: List[str]) -> None:
        with self._lock:
            self.exclude_strings.update(s.strip() for s in items if s.strip())
            self._save()

    def remove(self, item: str) -> None:
        with self._lock:
            self.exclude_strings.discard(item)
            self._save()

    # ─────────── I/O ───────────
    def _load(self):
        if DATA_PATH.exists():
            try:
                data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
                self.exclude_strings = set(data.get("exclude", []))
            except Exception as e:
                print("[ExclusionManager] Load error:", e)

    def _save(self):
        data = {"exclude": sorted(self.exclude_strings)}
        DATA_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
