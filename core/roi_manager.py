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
# 설정: 프로젝트 폴더에 `roi_sets.json` 저장
# ─────────────────────────────────────────────

# 프로젝트 최상위 디렉터리를 기준으로 경로 설정
ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = ROOT_DIR / "roi_sets.json"
# 파일이 없으면 생성
DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
if not DATA_PATH.exists():
    DATA_PATH.write_text(json.dumps({"sets": []}, ensure_ascii=False, indent=2), encoding="utf-8")

class ROIManager:
    """
    ROIManager
    ──────────
    • ROISet을 관리하고, JSON으로 영속화
    """
    _lock = threading.Lock()
    
    def __init__(self) -> None:
        self._sets: Dict[str, ROISet] = {}
        self._load()

    def _load(self) -> None:
        """JSON → 메모리"""
        # 디버그: 실제 로드 경로 출력
        print(f"[ROIManager] Loading ROI sets from: {DATA_PATH}")
        try:
            data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
            # 데이터 파싱 및 메모리에 저장
            for d in data.get("sets", []):
                rs = ROISet.from_dict(d)
                self._sets[rs.set_name] = rs
            # 디버그: 로드된 세트와 ROI 출력
            print(f"[ROIManager] Loaded sets:")
            for name, rs in self._sets.items():
                print(f"  - Set: {name}")
                for roi in rs.rois:
                    print(f"    ROI '{roi.name}': x={roi.x}, y={roi.y}, w={roi.w}, h={roi.h}")
        except Exception as e:
            print(f"[ROIManager] JSON load error: {e}")

    def _save(self) -> None:
        """메모리 → JSON (pretty-print, 한글 보존)"""
        data = {"sets": [rs.to_dict() for rs in self._sets.values()]}
        DATA_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )

    # CRUD API
    def list_sets(self) -> List[str]:
        return list(self._sets.keys())

    def get_set(self, name: str) -> ROISet | None:
        return self._sets.get(name)

    def upsert_set(self, rs: ROISet) -> None:
        with self._lock:
            self._sets[rs.set_name] = rs
            self._save()

    def delete_set(self, name: str) -> None:
        with self._lock:
            if name in self._sets:
                del self._sets[name]
                self._save()