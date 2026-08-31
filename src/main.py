from fastapi import FastAPI

app = FastAPI(
    title="Excel Management API",
    version="1.0.0",
)


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