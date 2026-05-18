from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from openpyxl import Workbook
from openpyxl.drawing.image import Image

import requests
import uuid
import os

app = FastAPI()

class Item(BaseModel):
    title: str
    price: float
    image_url: str

@app.post("/xlsx")
async def create_xlsx(items: list[Item]):

    wb = Workbook()
    ws = wb.active

    ws.append(["Title", "Price", "Image"])

    for idx, item in enumerate(items, start=2):

        ws.cell(idx, 1, item.title)
        ws.cell(idx, 2, item.price)

        img_data = requests.get(item.image_url).content

        img_path = f"/tmp/{uuid.uuid4()}.jpg"

        with open(img_path, "wb") as f:
            f.write(img_data)

        img = Image(img_path)

        img.width = 100
        img.height = 100

        ws.add_image(img, f"C{idx}")

        ws.row_dimensions[idx].height = 80

    file_path = "/tmp/products.xlsx"

    wb.save(file_path)

    return FileResponse(
        file_path,
        filename="products.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )