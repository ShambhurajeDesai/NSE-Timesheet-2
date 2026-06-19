from sqlalchemy import (
    Column,
    String,
    Date,
    Time,
    Interval,
    ForeignKey,
    BigInteger
)

from sqlalchemy.orm import relationship
from database import Base


class BiometricEmployee(Base):

    __tablename__ = "biometric_employee"

    emp_code = Column(String, primary_key=True)

    emp_name = Column(String)

    po_number = Column(String)

    attendance = relationship(
        "BiometricAttendance",
        back_populates="employee",
        cascade="all, delete"
    )


class BiometricAttendance(Base):

    __tablename__ = "biometric_attendance"

    attendance_id = Column(BigInteger, primary_key=True, autoincrement=True)

    emp_code = Column(
        String,
        ForeignKey("biometric_employee.emp_code")
    )

    attendance_date = Column(Date)

    shift_code = Column(String)

    login_time = Column(Time)

    logout_time = Column(Time)

    work_hours = Column(Interval)

    status = Column(String)

    employee = relationship(
        "BiometricEmployee",
        back_populates="attendance"
    )