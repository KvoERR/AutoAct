import sys

from autoact import AutoAct
from reader import ExcelReader, ReadSpec
from config import MODES


class AutoActApp:
    """Связывает reader, mode и spec. Не знает, как они работают внутри."""
    
    def __init__(self, reader: ExcelReader, mode: AutoAct, spec: ReadSpec):
        self.reader = reader
        self.mode = mode
        self.spec = spec
    
    def run(self):
        data = self.reader.read(self.spec)
        return self.mode.run(data)


def create_app(book_path: str, mode_name: str) -> AutoActApp:
    if mode_name not in MODES:
        raise ValueError(
            f"Неизвестный режим: {mode_name!r}. "
            f"Доступные: {list(MODES)}"
        )
    
    entry = MODES[mode_name]
    
    return AutoActApp(
        reader=ExcelReader(book_path),
        mode=entry["mode"](),
        spec=entry["spec"],
    )


def main():
    book_path = sys.argv[1] if len(sys.argv) > 1 else "ui.xlsm"
    
    # Читаем лист «Программа», чтобы узнать режим
    bootstrap = ExcelReader(book_path)
    params = bootstrap.read(ReadSpec(sheet_name="Программа"))
    
    app = create_app(book_path, params["Режим:"])
    app.run()


if __name__ == "__main__":
    main()