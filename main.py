from fastapi import FastAPI, File, HTTPException, Depends
from fastapi.security import APIKeyHeader
from pydantic import BaseModel
from starlette import status
import openpyxl
from dotenv import dotenv_values

# load config from .env to get X-API-KEY list
config = dotenv_values(".env")
api_keys = config['X_API_KEY']
X_API_KEY = APIKeyHeader(name='X-API-Key')

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

def api_key_auth(x_api_key: str = Depends(X_API_KEY)):
    # this function is used to validate X-API-KEY in request header
    # if the sent X-API-KEY in header is not existed in the config file
    #   reject access
    if x_api_key not in api_keys:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Forbidden"
        )

@app.post("/data", dependencies=[Depends(api_key_auth)])
def create_data(data: Data):
    wb = create_excel_file(data)
    return File(
        wb,
        filename="data.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        content_disposition="attachment",
    )
