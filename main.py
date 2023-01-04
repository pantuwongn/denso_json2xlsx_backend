from fastapi import FastAPI, File, HTTPException, Depends
from fastapi.responses import FileResponse
from fastapi.security import APIKeyHeader
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Dict, Union
from pydantic import BaseModel
from starlette import status
import uuid
import openpyxl
import json
from dotenv import dotenv_values
from e_pcs_form import PCSForm

# load config from .env to get X-API-KEY list
config = dotenv_values(".env")
api_keys = config['X_API_KEY']
X_API_KEY = APIKeyHeader(name='X-API-Key')

app = FastAPI()
origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)



def api_key_auth(x_api_key: str = Depends(X_API_KEY)):
    # this function is used to validate X-API-KEY in request header
    # if the sent X-API-KEY in header is not existed in the config file
    #   reject access
    if x_api_key not in api_keys:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Forbidden"
        )

@app.get("/get_mock_data", dependencies=[Depends(api_key_auth)])
def get_mock_data():
    with open('pcs_controlitem.json') as f:
        data = json.load(f)

    return data

@app.post("/convert_json_to_xlsx", dependencies=[Depends(api_key_auth)])
def create_data(data: Dict[str, Union[str, List]]):
    templateFilePath = './templates/e-pcs-control-item-form-template.xlsx'
    random_name = str(uuid.uuid4())
    PCSForm(templateFilePath, data).generate(random_name)
    return FileResponse(f"./output/{random_name}.xlsx", media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',filename='e-pcs.xlsx')
