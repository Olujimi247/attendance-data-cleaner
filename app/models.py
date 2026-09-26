from sqlalchemy import Column, Integer, String, Date, Time, Float
from app.database import Base  # Adjust import based on where your Base instance is defined

class AttendanceRecord(Base):
    __tablename__ = "attendance_records"

    id = Column(Integer, primary_key=True, index=True)
    employee_id = Column(String(50), index=True, nullable=False)
    resumption_date = Column(Date, index=True, nullable=False)
    resumption_time = Column(Time, nullable=False)
    closing_date = Column(String(20), nullable=False)  # Stored as string to handle "MISSING" safely
    closing_time = Column(String(20), nullable=False)  # Stored as string to handle "MISSING" safely
    hours_worked = Column(Float, default=0.0)
    overtime_hours = Column(Float, default=0.0)
    status = Column(String(50), default="OK")
