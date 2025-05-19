# core/ocr_engine.py
"""
OCR Engine 모듈 (ROI 좌표 픽셀 단위로 처리)
PyMuPDF로 PDF에서 지정 영역을 렌더링하고, OpenCV 전처리 후 Tesseract로 텍스트 추출
- PDF 렌더링 DPI = 600으로 고정
- ROI 좌표는 UI에서 저장된 픽셀 좌표를 그대로 사용
- 전처리: 그레이스케일→노이즈 제거(중간값 필터)→샤프닝
"""
import cv2
import numpy as np
import pytesseract
from PIL import Image
import fitz  # PyMuPDF

class OCREngine:
    def __init__(self, dpi: int = 600, lang: str = 'kor+eng') -> None:
        self.dpi = dpi
        self.lang = lang
        # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    def extract_roi(
        self,
        pdf_path: str | fitz.Document,
        page_num: int,
        rect: tuple[float, float, float, float],  # x_px, y_px, width_px, height_px
        tol: int = 0  # 픽셀 단위
    ) -> str:
        # PDF 페이지 로드 후 렌더링 (600dpi)
        close_doc = False
        if not isinstance(pdf_path, fitz.Document):
            doc = fitz.open(pdf_path)
            close_doc = True
        else:
            doc = pdf_path
        page = doc.load_page(page_num)
        zoom = self.dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

        # ROI 좌표는 픽셀 단위 → 직접 사용
        x, y, w, h = rect
        x1 = max(int(x) - tol, 0)
        y1 = max(int(y) - tol, 0)
        x2 = min(int(x + w) + tol, img_cv.shape[1])
        y2 = min(int(y + h) + tol, img_cv.shape[0])
        roi = img_cv[y1:y2, x1:x2]

        # 빈 영역 방지
        if roi.size == 0:
            return ""

                # 전처리: 그레이스케일 → 샤프닝 (노이즈 제거 단계 제거)
        gray = cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)
        kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
        processed = cv2.filter2D(gray, -1, kernel)

        # OCR 추출
        config = r"--oem 1 --psm 6"
        text = pytesseract.image_to_string(processed, lang=self.lang, config=config)

        # 문서 닫기
        if close_doc:
            doc.close()
        return text
