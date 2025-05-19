# core/models.py
"""
모델 정의: ROI, ROISet
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Dict, Any

@dataclass
class ROI:
    name: str
    x: int
    y: int
    w: int
    h: int
    tolerance: int = 0
    field_type: str = 'single'

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> ROI:
        # JSON 키 호환성: 'w' 또는 'width', 'h' 또는 'height'
        w = d.get('w', d.get('width'))
        h = d.get('h', d.get('height'))
        return ROI(
            name=d['name'],
            x=int(d['x']),
            y=int(d['y']),
            w=int(w),
            h=int(h),
            tolerance=int(d.get('tolerance', 0)),
            field_type=d.get('field_type', 'single')
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            'name': self.name,
            'x': self.x,
            'y': self.y,
            'w': self.w,
            'h': self.h,
            'tolerance': self.tolerance,
            'field_type': self.field_type
        }

@dataclass
class ROISet:
    set_name: str
    rois: List[ROI] = field(default_factory=list)

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> ROISet:
        rois = [ROI.from_dict(rd) for rd in d.get('rois', [])]
        return ROISet(set_name=d['set_name'], rois=rois)

    def to_dict(self) -> Dict[str, Any]:
        return {
            'set_name': self.set_name,
            'rois': [roi.to_dict() for roi in self.rois]
        }
