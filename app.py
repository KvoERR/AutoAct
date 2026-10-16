import pandas as pd
import sys
import traceback
from pathlib import Path

from autoact import AutoAct, ManualMode, AutoMode

class ExcelReader:
    def __init__(self, book_path: str):
        self.book_path = book_path

    def _read_params(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name, header=None)
        return {
            str(k).strip(): ("" if pd.isna(v) else v)
            for k, v in zip(df[0], df[1])
        }

    def get_mode_name(self, sheet_name: str) -> str:
        return self._read_params(sheet_name)["Режим:"]

    def get_file_name(self, sheet_name: str) -> str:
        return self._read_params(sheet_name)["Имя файла:"]

    def get_import(self, sheet_name: str) -> tuple[str, str]:
        params = self._read_params(sheet_name)
        return params["Исходные данные:"], params["Шаблоны:"]

    def get_changes(self, sheet_name: str) -> dict:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name)
        df = df.iloc[:, 1:]
        df = df.fillna("")
        return df.to_dict(orient="records")

    def get_features(self, sheet_name: str) -> list:
        df = pd.read_excel(self.book_path, sheet_name=sheet_name)
        df = df.fillna("")
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
    book_path = sys.argv[1] if len(sys.argv) > 1 else "ui.xlsm"
    reader = ExcelReader(book_path)

    mode_name = reader.get_mode_name("Программа")
    if mode_name not in MODES:
        raise ValueError(
            f"Неизвестный режим: {mode_name!r}. "
            f"Доступные: {list(MODES)}"
        )

    entry = MODES[mode_name]

    filename = reader.get_file_name("Программа")
    data,templates = reader.get_import("Данные")
    changes = reader.get_changes(entry["sheet"])
    features = reader.get_features("Функции")
    autoact = entry["mode"](filename, data, templates, changes, features)
    autoact.run()


if __name__ == "__main__":
    log_path = Path(__file__).parent / "error.log"
    
    try:
        main()
    except Exception:
        tb = traceback.format_exc()
        log_path.write_text(tb, encoding="utf-8")
        print(tb, file=sys.stderr)
        sys.exit(1)