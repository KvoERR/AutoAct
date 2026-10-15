import pandas as pd

from autoact import AutoAct, ManualMode, AutoMode

class ExcelReader:
    def __init__(self, book_path: str):
        self.book_path = book_path

    def _read_params(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name, header=None)
        return dict(zip(df[0], df[1]))

    def get_mode_name(self, sheet_name: str) -> str:
        params = self._read_params(sheet_name)
        return params["Режим:"]

    def get_import(self, sheet_name: str) -> tuple[str, str]:
        params = self._read_params(sheet_name)
        return params["Исходные данные:"], params["Шаблоны:"]

    def get_changes(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name)
        df = df.iloc[:, 1:]
        return df.to_dict(orient="list")

    def get_features(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name)
        return df.to_dict(orient="records")

MODES = {
    "Ручной": {
        "sheet": "Ручной режим",
        "mode": ManualMode,
    },
    "Автоматический": {
        "sheet": "Автоматический режим",
        "mode": AutoMode,
    },
}

def main():
    book_path = "ui.xlsm"
    reader = ExcelReader(book_path)

    mode_name = reader.get_mode_name("Программа")
    if mode_name not in MODES:
        raise ValueError(
            f"Неизвестный режим: {mode_name!r}. "
            f"Доступные: {list(MODES)}"
        )

    entry = MODES[mode_name]

    data,templates = reader.get_import("Данные")
    changes = reader.get_changes(entry["sheet"])
    features = reader.get_features("Функции")
    autoact = entry["mode"](data, templates, changes, features)


if __name__ == "__main__":
    main()