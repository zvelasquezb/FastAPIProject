from fastapi import Path



async def root():
    return {"message": "Hello World"}


async def say_hello(name: str):
    return {"message": f"Hello {name}"}
