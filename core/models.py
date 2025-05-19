# core/models.py
"""
데이터 계층의 기본 모델 정의
─ ROI(좌표) · ROISet(좌표 세트) · ExclusionRule(제외 텍스트 규칙)
직렬화 편의를 위해 dataclass + dict 변환 메서드 제공
"""

from __future__ import annotations
from dataclasses import dataclass, asdict
from typing import List, Dict


# ─────────────────────────────────────────────
# 1. ROI (Region Of Interest) : 사각 좌표 하나
# ─────────────────────────────────────────────
@dataclass
class ROI:
    name: str               # “품명”, “수량” …
    x: int                  # left   (PDF 픽셀 기준)
    y: int                  # top
    w: int                  # width
    h: int                  # height
    tolerance: int = 0      # 오차(px)
    field_type: str = "single"      # "single" | "table"

    # JSON 직렬화용
    def to_dict(self) -> Dict:
        return asdict(self)

    @staticmethod
    def from_dict(d: Dict) -> "ROI":
        return ROI(**d)


# ─────────────────────────────────────────────
# 2. ROISet : 좌표 여러 개를 이름으로 묶은 단위
# ─────────────────────────────────────────────
@dataclass
class ROISet:
    set_name: str
    rois: List[ROI]

    def to_dict(self) -> Dict:
        return {
            "set_name": self.set_name,
            "rois": [r.to_dict() for r in self.rois],
        }

    @staticmethod
    def from_dict(d: Dict) -> "ROISet":
        return ROISet(
            set_name=d["set_name"],
            rois=[ROI.from_dict(r) for r in d["rois"]],
        )


# ─────────────────────────────────────────────
# 3. ExclusionRule : 특정 좌표 필드에서 제외할 텍스트
#    - linked_fields : 같은 순서로 동시 제외할 필드 목록
# ─────────────────────────────────────────────
@dataclass
class ExclusionRule:
    target_field: str               # 예) "품명"
    exclude_text: str               # 예) "견적"
    linked_fields: List[str] = None # 예) ["수량"]

    def to_dict(self) -> Dict:
        return {
            "target_field": self.target_field,
            "exclude_text": self.exclude_text,
            "linked_fields": self.linked_fields or [],
        }

    @staticmethod
    def from_dict(d: Dict) -> "ExclusionRule":
        return ExclusionRule(
            target_field=d["target_field"],
            exclude_text=d["exclude_text"],
            linked_fields=d.get("linked_fields", []),
        )
