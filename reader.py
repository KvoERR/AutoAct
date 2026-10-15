from dataclasses import dataclass
import pandas as pd


@dataclass
class ReadSpec:
    """Что читать. Не зависит от режима."""
    sheet_name: str


class ExcelReader:
    """Читает Excel. Не знает про режимы."""
    
    def __init__(self, book_path: str):
        self.book_path = book_path
    
    def read(self, spec: ReadSpec) -> dict:
        df = pd.read_excel(
            self.book_path,
            sheet_name=spec.sheet_name,
            header=None,
        )
        return dict(zip(df[0], df[1]))