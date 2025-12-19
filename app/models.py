from sqlalchemy import (
    Boolean, Column, Integer, String, Float,
    ForeignKey, Date, DateTime, Time
)
from sqlalchemy.orm import relationship
from datetime import datetime
from .database import Base


class Employee(Base):
    __tablename__ = "employees"

    id = Column(Integer, primary_key=True)
    emp_code = Column(String, unique=True)
    name = Column(String)
    monthly_salary = Column(Float)
    per_day_salary = Column(Float)
    salary_cycle_start_day = Column(Integer)
    is_active = Column(Boolean, default=True)


class Penalty(Base):
    __tablename__ = "penalties"

    id = Column(Integer, primary_key=True)
    employee_id = Column(Integer, ForeignKey("employees.id"))
    date = Column(Date)
    reason = Column(String)
    amount = Column(Float)
    created_at = Column(DateTime, default=datetime.utcnow)

    employee = relationship("Employee")


class Attendance(Base):
    __tablename__ = "attendance_logs"

    id = Column(Integer, primary_key=True)
    employee_id = Column(Integer, ForeignKey("employees.id"))
    date = Column(Date)
    in_time = Column(Time, nullable=True)
    out_time = Column(Time, nullable=True)
    status = Column(String)
    payable_value = Column(Float)

    employee = relationship("Employee")


class SalaryRun(Base):
    __tablename__ = "salary_runs"

    id = Column(Integer, primary_key=True)
    employee_id = Column(Integer, ForeignKey("employees.id"))

    cycle_start = Column(Date)
    cycle_end = Column(Date)

    present_days = Column(Float)
    half_days = Column(Float)
    absent_days = Column(Float)
    sunday_deductions = Column(Integer)

    penalty_amount = Column(Float)
    gross_salary = Column(Float)
    net_salary = Column(Float)

    created_at = Column(DateTime, default=datetime.utcnow)

    employee = relationship("Employee")