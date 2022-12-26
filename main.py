from fastapi import FastAPI, File
from pydantic import BaseModel
import openpyxl

app = FastAPI()

class Data(BaseModel):
    name: str
    age: int

def create_excel_file(data: Data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Name"
    ws["B1"] = "Age"
    ws["A2"] = data.name
    ws["B2"] = data.age
    return wb

@app.post("/data")
def create_data(data: Data):
    wb = create_excel_file(data)
    return File(
        wb,
        filename="data.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        content_disposition="attachment",
    )
