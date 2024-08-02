from sqlalchemy import create_engine, Column, Integer, String, Date, Time, ForeignKey, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship, sessionmaker

DATABASE_URL = "sqlite:///./test.db"

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class Employee(Base):
    __tablename__ = 'employees'

    id = Column(Integer, primary_key=True, index=True)
    telegram_id = Column(Integer, unique=True, index=True)
    fio = Column(String, index=True)
    full_name = Column(String, index=True)
    attendances = relationship("Attendance", back_populates="employee")

class Attendance(Base):
    __tablename__ = 'attendances'

    id = Column(Integer, primary_key=True, index=True)
    employee_id = Column(Integer, ForeignKey('employees.id'), nullable=False)
    date = Column(Date, nullable=False)
    arrival_time = Column(Time, nullable=True)
    departure_time = Column(Time, nullable=True)
    employee = relationship("Employee", back_populates="attendances")
    late = Column(Boolean, default=False)
    departure_type = Column(String, nullable=True)
    departure_reason = Column(String, nullable=True)
    supervisor = Column(String, nullable=True)
    departure_time_actual = Column(Time, nullable=True)
    return_time = Column(Time, nullable=True)
    check = Column(Boolean, default=False)

def init_db():
    Base.metadata.create_all(bind=engine)
