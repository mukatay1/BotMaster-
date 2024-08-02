from datetime import date

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from database import Base, Employee, Attendance

engine = create_engine('sqlite:///test.db')
Session = sessionmaker(bind=engine)
db = Session()


existing_attendance = db.query(Attendance).filter(
            Attendance.date == date.today(),
            Attendance.employee_id == 1).first()

db.delete(existing_attendance)
db.commit()
db.close()
