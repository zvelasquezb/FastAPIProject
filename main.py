from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates
from typing import List
import shutil
import os
from src.util.process_pdf import process_pdf
from src.util.fill_excel import fill_excel
from src.routes import test_routes
from fastapi.staticfiles import StaticFiles

app = FastAPI()

app.mount("/static", StaticFiles(directory="src/static"), name="static")
app.include_router(test_routes.test_router, prefix="/test", tags=["test"])

app.title = "Asistente RENOBO para diligenciamiento SIVICOF"
app.version = "0.0.1"

templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload/")
async def upload_files(files: List[UploadFile] = File(...)):
    results = []
    temp_dir = "temp_pdf"
    os.makedirs(temp_dir, exist_ok=True)

    for file in files:
        file_path = os.path.join(temp_dir, file.filename)
        try:
            # Guardar archivo temporal
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

            # Procesar PDF
            data = process_pdf(pdf_path=file_path)
            results.append({"filename": file.filename, "data": data})

        except Exception as e:
            results.append({"filename": file.filename, "error": str(e)})
        
        finally:
            # Eliminar archivo temporal si existe
            if os.path.exists(file_path):
                os.remove(file_path)

    # Llenar el Excel con los resultados procesados
    fill_excel(results)
    output_excel = fill_excel(results)
    print("Excel generado en:", output_excel)

    # Eliminar directorio temporal solo si está vacío
    if os.path.exists(temp_dir) and not os.listdir(temp_dir):
        os.rmdir(temp_dir)

    return JSONResponse(content={"results": results})


@app.get("/download-excel/")
async def download_excel():
    file_path = os.path.abspath("src/template_excel/Inf SIVICOF_PLANTILLA.xlsx")
    if os.path.exists(file_path):
        return FileResponse(
            path=file_path,
            filename="resultado.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return JSONResponse(content={"error": f"El archivo Excel no existe en {file_path}. Primero procesa un PDF."}, status_code=404)