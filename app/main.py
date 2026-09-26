from fastapi import FastAPI
from app.routers import attendance

# --- IMPORT DATABASE LAYER CONFIGS ---
from app.database import engine, Base
import app.models  # Imports models so SQLAlchemy recognizes the table layout

# 1. Force SQLAlchemy to create the database tables if they do not exist
Base.metadata.create_all(bind=engine)

# 2. Initialize the app instance
app = FastAPI(title="Maintenance and Attendance System")

app.include_router(attendance.router)

@app.get("/")
def read_root():
    return {"message": "Attendance API is running smoothly!"}
