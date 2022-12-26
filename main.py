from fastapi import FastAPI, File, HTTPException, Depends
from fastapi.security import APIKeyHeader
from typing import List
from pydantic import BaseModel
from starlette import status
import openpyxl
from dotenv import dotenv_values

# load config from .env to get X-API-KEY list
config = dotenv_values(".env")
api_keys = config['X_API_KEY']
X_API_KEY = APIKeyHeader(name='X-API-Key')

app = FastAPI()


class InitialPCapabilityData(BaseModel):
    x_bar: str
    cpk: str

class ControlMethodData(BaseModel):
    sample_no: int
    interval: str
    method_100: str
    in_charge: str
    calibration_interval: str

class SCSymbolData(BaseModel):
    character: str
    shape: str

class ProcessItemParameterData(BaseModel):
    parameter: str
    master_value: int
    limit_type: str
    prefix: str
    main: str
    suffix: str
    tolerance_up: str
    tolerance_down: str
    upper_limit: str
    lower_limit: str
    unit: str

class RemarkData(BaseModel):
    remark: str
    ws_no: str
    related_std: str

class ProcessItemData(BaseModel):
    control_item_no: int
    control_item_type: str
    parameter: ProcessItemParameterData
    sc_symbols: List[SCSymbolData]
    check_timing: str
    control_method: ControlMethodData
    initial_p_capability: InitialPCapabilityData
    remark: RemarkData
    measurement: str
    readability: str
    start_effective: str

class ProcessData(BaseModel):
    name: str
    items: List[ProcessItemData]

class JsonData(BaseModel):
    pcs_no: str
    date: str
    status: str
    line: str
    assy_name: str
    part_name: str
    customer: str
    processes: List[ProcessData]

def create_excel_file(data: JsonData):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Name"
    ws["B1"] = "Age"
    ws["A2"] = "Natapon"
    ws["B2"] = "41"
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

@app.post("/convert_json_to_xlsx", dependencies=[Depends(api_key_auth)])
def create_data(data: JsonData):
    wb = create_excel_file(data)
    return File(
        wb,
        filename="data.xlsx",
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        content_disposition="attachment",
    )
