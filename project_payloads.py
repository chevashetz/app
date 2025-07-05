from utils import extract_number


MODEL_MAP = {
    "Гершель-Балкли": "HerschelBulkley",
    "Степенная": "powerLaw",
    "Бингама": "Bingam"
}

def countFilledRows(table, fluid_params: dict) -> int:
    count = 0
    for row in range(table.rowCount()):
        # Проверяем основные поля в таблице
        depth_from = table.item(row, 1)
        depth_to = table.item(row, 2)
        density = table.item(row, 3)
        # Проверяем, что в fluid_params есть параметры для этой строки
        param = fluid_params.get(row)
        if (depth_from and depth_from.text().strip() and
            depth_to and depth_to.text().strip() and
            density and density.text().strip() and
            param and param.get("model") and param.get("params") and len(param.get("params")) > 0):
            count += 1
    return count

def build_fluids_payloads(fluids_table, fluid_params: dict, MODEL_MAP, extract_number):
    """Собирает список payload для всех заполненных строк таблицы fluids."""
    payloads = []
    for row in range(fluids_table.rowCount()):
        # Берём выбранную модель (по-русски) из таблицы
        combo = fluids_table.cellWidget(row, 4)
        if combo:
            model_rus = combo.currentText()
            transport_model = MODEL_MAP.get(model_rus, "")
        else:
            transport_model = ""

        # rheology_params
        param = fluid_params.get(row, {})
        rheology_params = param.get("params", {})

        # Собираем payload
        payload = {
            "name": fluids_table.item(row, 0).text() if fluids_table.item(row, 0) else "",
            "depth_from": extract_number(fluids_table.item(row, 1).text()),
            "depth_to": extract_number(fluids_table.item(row, 2).text()),
            "density": extract_number(fluids_table.item(row, 3).text()),
            "transport_model": transport_model,
            "rheology_params": rheology_params,
        }
        # Проверка заполненности
        if not payload["transport_model"]:
            continue
        payloads.append(payload)
    return payloads
