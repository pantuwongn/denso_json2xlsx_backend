FROM python:3.8-slim

COPY . /app
WORKDIR /app

RUN pip install fastapi openpyxl uvicorn

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]