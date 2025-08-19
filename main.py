from fastapi import FastAPI, Body
from src.routes.test_routes import test_router
app = FastAPI()

app.title = "Asistente RENOBO para diligenciamiento SIVICOF"

app.version = "0.0.1"

app.include_router(prefix="/test",router=test_router)