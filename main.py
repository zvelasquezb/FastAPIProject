from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from typing import List
import shutil
import os
from src.util.process_pdf import process_pdf
from src.util.fill_excel import fill_excel

app = FastAPI()

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
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            data = process_pdf(pdf_path=file_path)
            results.append({"filename": file.filename, "data": data})

        except Exception as e:
            results.append({"filename": file.filename, "error": str(e)})
        finally:
            # Move the file to a temporary location to be deleted later
            if os.path.exists(file_path):
                temp_file_path = os.path.join("temp_pdf", f"deletable_{file.filename}")
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)
        shutil.move(file_path, temp_file_path)

    print(len(results),"&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&")
    fill_excel(results)
    # Eliminar el directorio temporal solo si está vacío
    if os.path.exists(temp_dir) and not os.listdir(temp_dir):
        os.rmdir(temp_dir)
        
    return JSONResponse(content={"results": results})