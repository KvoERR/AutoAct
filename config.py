from autoact import ManualMode, AutoMode
from reader import ReadSpec


MODES = {
    "Ручной": {
        "mode": ManualMode,
        "spec": ReadSpec(sheet_name="Ручной"),
    },
    "Автоматический": {
        "mode": AutoMode,
        "spec": ReadSpec(sheet_name="Авто"),
    },
}