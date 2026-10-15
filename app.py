import pandas as pd
from abc import ABC, abstractmethod

from autoact import AutoAct, ManualMode, AutoMode


class BaseParser(ABC):
    def __init__(self, book_path: str):
        self.book_path = book_path

    @abstractmethod
    def read(self, sheet_name: str) -> dict:
        raise


class AutoParser(BaseParser):
    def read(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name)
        return df.to_dict(orient="list")
    
class ManualParser(BaseParser):
    def read(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name)
        df = df.iloc[:, 1:]
        return df.to_dict(orient="list")


class ExcelReader:
    def __init__(self, book_path: str):
        self.book_path = book_path

    def get_mode_name(self, sheet_name: str) -> str:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name, header=None)
        return df[1][0]


MODES = {
    "Ручной": {
        "sheet": "Ручной режим",
        "parser": ManualParser,
        "mode": ManualMode,
    },
    "Автоматический": {
        "sheet": "Автоматический режим",
        "parser": AutoParser,
        "mode": AutoMode,
    },
}


class App:
    def __init__(self, sheet: str, parser: BaseParser, mode: AutoAct):
        self.sheet = sheet
        self.Parser = parser
        self.mode = mode

    def run(self):
        data = self.Parser.read(self.sheet)
        return self.mode.run(data)


def main():
    book_path = "ui.xlsm"
    program_sheet_name = "Программа"

    mode_name = ExcelReader(book_path).get_mode_name(program_sheet_name)

    if mode_name not in MODES:
        raise ValueError(
            f"Неизвестный режим: {mode_name!r}. "
            f"Доступные: {list(MODES)}"
        )

    entry = MODES[mode_name]
    app = App(
        sheet=entry["sheet"],
        Parser=entry["parser"](book_path),
        mode=entry["mode"](),
    )
    app.run()


if __name__ == "__main__":
    main()