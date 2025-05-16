# backend/database.py

import numpy as np
import pandas as pd
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker
from datetime import date, datetime

DATABASE_URL = "postgresql://postgres:gmEASxLQBrqaYptnTpjHYDEuXBxlpDkm@centerbeam.proxy.rlwy.net:42161/railway"

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()


class Order(Base):
    __tablename__ = "orders"

    id = Column(Integer, primary_key=True, index=True)
    order_id = Column(String, index=True)
    order_date = Column(Date)
    total_area = Column(Float)
    upload_date = Column(DateTime, default=lambda: datetime.now())

    area_post_1 = Column(Float)
    area_post_2 = Column(Float)
    area_post_3 = Column(Float)
    area_post_4 = Column(Float)
    area_post_5 = Column(Float)
    area_post_6 = Column(Float)
    area_post_7 = Column(Float)
    area_post_8 = Column(Float)
    area_post_9 = Column(Float)
    area_post_10 = Column(Float)

    def as_dict(self):
        return {
            "ID": self.id,
            "Номер заказа": self.order_id,
            "Дата заказа": self.order_date,
            "Дата загрузки": self.upload_date,
            "Площадь всего": self.total_area,
            **{f"Пост {i}": getattr(self, f"area_post_{i}") for i in range(1, 11)},
        }


def get_orders():
    with SessionLocal() as session:
        return session.query(Order).order_by(Order.upload_date.desc()).all()


def get_order_by_id(order_id: str):
    with SessionLocal() as session:
        return session.query(Order).filter(Order.order_id == order_id).first()


def delete_order_by_id(order_id: str):
    with SessionLocal() as session:
        session.query(Order).filter(Order.order_id == order_id).delete()
        session.commit()


def insert_order(data: dict):
    data["upload_date"] = datetime.utcnow()
    with SessionLocal() as session:
        order = Order(**data)
        session.add(order)
        session.commit()


def clean_data(data: dict) -> dict:
    result = {}
    for k, v in data.items():
        if isinstance(v, (np.float64, np.int64)):
            result[k] = float(v)
        elif isinstance(v, pd.Timestamp):
            result[k] = v.date()
        else:
            result[k] = v
    return result


# def init_db():
#     Base.metadata.drop_all(bind=engine)
#     Base.metadata.create_all(bind=engine)
