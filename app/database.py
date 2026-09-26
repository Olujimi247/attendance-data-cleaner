import os
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# If running on Render, it will look for DATABASE_URL. Otherwise, defaults to local SQLite.
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./attendance_system.db")

# PostgreSQL doesn't need check_same_thread, but SQLite does
if DATABASE_URL.startswith("sqlite"):
    engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
else:
    engine = create_engine(DATABASE_URL)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

