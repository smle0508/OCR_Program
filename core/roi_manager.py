# core/roi_manager.py
"""
ROIManager
──────────
• ROI 세트(ROISet)를 JSON 파일 하나로 영속화
• CRUD(생성·조회·수정·삭제) API 제공
• 스레드 안전성을 위해 모든 I/O 는 내부 Lock 사용
"""

from __future__ import annotations

import json
import threading
from pathlib import Path
from typing import List, Dict

from .models import ROI, ROISet

# ─────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────
APP_DIR = Path.home() / "AppData" / "Roaming" / "PdfOcrExcel"
APP_DIR.mkdir(parents=True, exist_ok=True)

DATA_PATH = APP_DIR / "roi_sets.json"      # 모든 세트 저장 파일


# ─────────────────────────────────────────────
# ROIManager
# ─────────────────────────────────────────────
class ROIManager:
    """좌표 세트 관리 + JSON 직렬화"""

    _lock = threading.RLock()      # 파일 I/O 동시 접근 보호

    def __init__(self) -> None:
        self._sets: Dict[str, ROISet] = {}
        self._load()

    # ─── Public API ──────────────────────────
    def list_sets(self) -> List[str]:
        """세트 이름 목록 반환"""
        return list(self._sets.keys())

    def get_set(self, set_name: str) -> ROISet | None:
        """세트 이름으로 ROISet 반환 (없으면 None)"""
        return self._sets.get(set_name)

    def upsert_set(self, roi_set: ROISet) -> None:
        """세트 추가 또는 덮어쓰기 후 저장"""
        with self._lock:
            self._sets[roi_set.set_name] = roi_set
            self._save()

    def delete_set(self, set_name: str) -> None:
        """세트 삭제 후 저장"""
        with self._lock:
            if set_name in self._sets:
                del self._sets[set_name]
                self._save()

    # ─── Internal: File I/O ──────────────────
    def _load(self) -> None:
        """JSON → 메모리"""
        if not DATA_PATH.exists():
            return
        try:
            data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
            for d in data.get("sets", []):
                rs = ROISet.from_dict(d)
                self._sets[rs.set_name] = rs
        except Exception as e:
            print(f"[ROIManager] JSON 로드 오류: {e}")

    def _save(self) -> None:
        """메모리 → JSON (pretty-print, 한글 보존)"""
        data = {
            "sets": [rs.to_dict() for rs in self._sets.values()]
        }
        DATA_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
