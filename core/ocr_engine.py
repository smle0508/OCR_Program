# core/ocr_engine.py
"""
OCR 엔진 래퍼
────────────
• PyMuPDF로 PDF 페이지를 이미지로 변환
• 영역(Crop) → 전처리 → Tesseract 인식
• 단일 필드(extract_roi) 및 테이블(extract_table) 지원
"""

from __future__ import annotations
import io
from pathlib import Path
from typing import Tuple, List

import fitz           # PyMuPDF
import numpy as np
from PIL import Image, ImageFilter, ImageOps
import pytesseract


class OCREngine:
    def __init__(
        self,
        dpi: int = 600,
        lang: str = "kor+eng",
        psm: int = 6,
        oem: int = 3,
        whitelist: str = ""
    ) -> None:
        self.dpi = dpi
        self.lang = lang
        self.psm = psm
        self.oem = oem
        self.whitelist = whitelist

    def _load_image(
        self,
        pdf_path: str | Path,
        page_num: int
    ) -> Image.Image:
        # PDF 페이지를 PIL 이미지로 변환
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        mat = fitz.Matrix(self.dpi / 72, self.dpi / 72)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes()))
        return img

    def _preprocess(self, img: Image.Image, tol: int = 0) -> Image.Image:
        # 그레이스케일 변환
        gray = img.convert("L")
        # 가우시안 블러
        blurred = gray.filter(ImageFilter.GaussianBlur(radius=1))
        # NumPy 배열로 thresholding
        np_img = np.array(blurred)
        thresh = np.mean(np_img)
        bin_arr = (np_img > thresh).astype(np.uint8) * 255
        bin_img = Image.fromarray(bin_arr)
        # 팽창(Dilation)으로 획 강화
        bin_img = bin_img.filter(ImageFilter.MaxFilter(3))
        return bin_img

    def extract_roi(
        self,
        pdf_path: str | Path,
        page_num: int,
        roi: Tuple[int, int, int, int],
        tolerance: int = 0
    ) -> str:
        """
        단일 필드 OCR
        roi=(x,y,w,h), tolerance 픽셀 여유 영역
        """
        img = self._load_image(pdf_path, page_num)
        x, y, w, h = roi
        x0 = max(0, x - tolerance)
        y0 = max(0, y - tolerance)
        x1 = min(img.width, x + w + tolerance)
        y1 = min(img.height, y + h + tolerance)
        cropped = img.crop((x0, y0, x1, y1))
        proc = self._preprocess(cropped, tolerance)
        # Tesseract config
        config = f"--psm {self.psm} --oem {self.oem}"
        if self.whitelist:
            config += f" -c tessedit_char_whitelist={self.whitelist}"
        text = pytesseract.image_to_string(
            proc,
            lang=self.lang,
            config=config
        )
        return text.strip()

    def extract_table(
        self,
        pdf_path: str | Path,
        page_num: int,
        roi: Tuple[int, int, int, int]
    ) -> List[List[str]]:
        """
        테이블 영역 OCR → 2D 리스트 반환
        pytesseract.image_to_data로 셀 감지
        """
        img = self._load_image(pdf_path, page_num)
        x, y, w, h = roi
        cropped = img.crop((x, y, x + w, y + h))
        proc = self._preprocess(cropped)
        config = f"--psm 6 --oem {self.oem}"
        data = pytesseract.image_to_data(
            proc,
            lang=self.lang,
            config=config,
            output_type=pytesseract.Output.DICT
        )
        # 행 단위 그룹화
        n_boxes = len(data['level'])
        rows: dict[int, List[tuple[int, str]]] = {}
        for i in range(n_boxes):
            text = data['text'][i].strip()
            if not text:
                continue
            top = data['top'][i]
            left = data['left'][i]
            # row key based on top coordinate
            rows.setdefault(top, []).append((left, text))
        table: List[List[str]] = []
        for top in sorted(rows.keys()):
            cells = [t for _, t in sorted(rows[top], key=lambda x: x[0])]
            table.append(cells)
        return table
