from pathlib import Path
from openpyxl import Workbook

class ExcelWriter:
    """
    엑셀 파일을 생성하고, 지정된 셀에 값을 기록한 후 저장하는 유틸리티 클래스
    """
    def __init__(self, file_path: Path, sheet_name: str = "Sheet1"):
        # 저장할 파일 경로 및 워크북/워크시트 초기화
        self.file_path = Path(file_path)
        self.wb = Workbook()
        self.ws = self.wb.active
        self.ws.title = sheet_name

    def write_values(self, data: dict[tuple[int, int], str]) -> None:
        """
        셀 좌표와 값을 매핑하여 워크시트에 기록합니다.
        data: {(row, col): value, ...}
        """
        for (row, col), value in data.items():
            self.ws.cell(row=row, column=col, value=value)

    def save(self) -> None:
        """
        지정된 경로에 워크북을 저장합니다. 필요한 경우 디렉터리를 생성합니다.
        """
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(self.file_path)
