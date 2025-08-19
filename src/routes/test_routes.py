from fastapi import Path, APIRouter
import src.controllers.test_controller as test_controller

test_router = APIRouter()

@test_router.get("/",tags=["pruebas"])
async  def root_from_controller():
    return await test_controller.root()

@test_router.get("/hello/{name}",tags=["pruebas"])
async def say_hello(name: str):
    return {"message": f"Hello {name}"}
