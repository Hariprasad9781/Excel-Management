from fastapi import FastAPI

from api.auth import router as auth_router
from api.files import router as files_router
from api.excel import router as excel_router
from api.workbooks import router as workbooks_router


app = FastAPI(
    title="Excel Management API",
    version="1.0.0",
)


app.include_router(auth_router)
app.include_router(files_router)
app.include_router(excel_router)
app.include_router(workbooks_router)


@app.get("/")
def root():
    return {
        "message": "Excel Management API is running"
    }


@app.get("/health")
def health_check():
    return {
        "status": "healthy"
    }