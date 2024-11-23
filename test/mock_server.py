import asyncio

from fastapi import FastAPI, Request
from pydantic import BaseModel
import random
import uvicorn
from datetime import datetime

# Создаем модели данных
class FluidCalculationRequest(BaseModel):
    name: str | None = None
    depth_from: float | None = None
    depth_to: float | None = None
    transport_model: str | None = None
    density: float | None = None
    viscosity: float | None = None
    dns: float | None = None

class FluidCalculationResponse(BaseModel):
    calculation_id: str
    timestamp: str
    status: str
    results: dict
    warnings: list[str] | None = None

app = FastAPI(title="Mock Calculation Server")

@app.post("/calculate/")
async def calculate(request: FluidCalculationRequest):
    # Генерируем моковые данные для ответа
    mock_results = {
        "velocity_profile": {
            "radiuses": [round(random.uniform(0, 1), 3) for _ in range(10)],
            "velocities": [round(random.uniform(0, 2), 3) for _ in range(10)]
        },
        "pressure_loss": {
            "total": round(random.uniform(10, 100), 2),
            "friction": round(random.uniform(5, 50), 2),
            "gravity": round(random.uniform(5, 50), 2)
        },
        "reynolds_number": round(random.uniform(2000, 4000)),
        "flow_regime": random.choice(["ламинарный", "турбулентный"]),
        "effective_viscosity": round(random.uniform(10, 30), 2)
    }

    await asyncio.sleep(random.randint(1, 3))  # эмуляция расчетов

    # Формируем полный ответ
    response = FluidCalculationResponse(
        calculation_id=f"calc_{random.randint(1000, 9999)}",
        timestamp=datetime.now().isoformat(),
        status="completed",
        results=mock_results,
        warnings=["Тестовый расчет - данные сгенерированы случайным образом"]
    )

    return response

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000)