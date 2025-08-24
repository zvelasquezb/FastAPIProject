import os
from fastapi import Path, APIRouter
import src.controllers.test_controller as test_controller
from fastapi.responses import FileResponse

test_router = APIRouter()

@test_router.get("/",tags=["pruebas"])
async  def root_from_controller():
    return await test_controller.root()

@test_router.get("/hello/{name}",tags=["pruebas"])
async def say_hello(name: str):
    return {"message": f"Hello {name}"}

@test_router.get("/download-excel/")
async def download_excel():
    file_path = os.path.abspath("src/template_excel/Inf SIVICOF_PLANTILLA.xlsx")
    if os.path.exists(file_path):
        return FileResponse(
            path=file_path,
            filename="resultado.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return {"error": "El archivo Excel no existe. Primero procesa un PDF."}
