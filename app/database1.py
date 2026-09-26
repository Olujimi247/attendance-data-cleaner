from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# 1. Create local SQLite database engine configuration
SQLALCHEMY_DATABASE_URL = "sqlite:///./attendance_system.db"

engine = create_engine(
    SQLALCHEMY_DATABASE_URL, connect_args={"check_same_thread": False}
)

# 2. Establish Session factory and declarative base tracking layers
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# 3. Create the db session dependency context used by the attendance router
def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
