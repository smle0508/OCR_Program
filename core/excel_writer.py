# core/excel_writer.py
"""
ExcelWriter
───────────
• 모든 템플릿(.xls/.xlsx/.xlsm) openpyxl로 안전하게 처리
• .xls 템플릿은 xlrd로 임시 .xlsx 변환 후 openpyxl 로드
• append_row_by_header / append_row_by_column_letter
  - column_letter가 'A81'처럼 숫자 포함 시, 문자만 추출
• save_as_new → 항상 .xlsx로 저장하여 포맷/확장자 불일치 방지
"""
from __future__ import annotations

from pathlib import Path
from datetime import datetime
from typing import Dict
import tempfile
import re

import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook

import xlrd


class ExcelWriter:
    """엑셀 행 추가 헬퍼"""

    def __init__(self, template: str | Path):
        self.template = Path(template)
        ext = self.template.suffix.lower()

        if ext == '.xls':
            # .xls 템플릿 → 임시 .xlsx 생성
            book_xls = xlrd.open_workbook(str(self.template))
            wb_new = Workbook()
            # 기본 시트 제거
            default = wb_new.active
            wb_new.remove(default)
            # 시트 복사
            for sheet_name in book_xls.sheet_names():
                sh = book_xls.sheet_by_name(sheet_name)
                ws = wb_new.create_sheet(title=sheet_name)
                for r in range(sh.nrows):
                    for c in range(sh.ncols):
                        ws.cell(row=r+1, column=c+1, value=sh.cell_value(r, c))
            # 임시 파일로 저장 및 로드
            tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
            wb_new.save(tmp.name)
            self.wb = openpyxl.load_workbook(tmp.name)
        else:
            # .xlsx, .xlsm 템플릿
            self.wb = openpyxl.load_workbook(str(self.template))

    def append_row_by_header(
        self,
        sheet: str,
        header_value: Dict[str, str],
        header_row: int = 1
    ):
        ws = self.wb[sheet]
        headers = {cell.value: cell.column for cell in ws[header_row] if cell.value}
        new_row = ws.max_row + 1
        for key, val in header_value.items():
            col = headers.get(key)
            if col:
                ws.cell(row=new_row, column=col, value=val)

    def append_row_by_column_letter(
        self,
        sheet: str,
        col_value: Dict[str, str]
    ):
        """
        col_value keys may include row numbers, e.g. 'A81'.
        Extract column letters and ignore digits.
        """
        ws = self.wb[sheet]
        new_row = ws.max_row + 1
        for col_letter, val in col_value.items():
            match = re.match(r"^([A-Za-z]+)", col_letter)
            if not match:
                continue
            letters = match.group(1).upper()
            try:
                idx = column_index_from_string(letters)
            except ValueError:
                continue
            ws.cell(row=new_row, column=idx, value=val)

    def save_as_new(self, out_dir: str | Path | None = None) -> Path:
        out_dir = Path(out_dir) if out_dir else self.template.parent
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        # 항상 .xlsx로 저장
        new_name = f"{self.template.stem}_{ts}.xlsx"
        out_path = out_dir / new_name
        self.wb.save(str(out_path))
        return out_path
