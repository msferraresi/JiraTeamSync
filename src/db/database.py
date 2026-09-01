import os
import sys
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from src.db.models import Base, AppSettings


def get_db_path():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        )
    return os.path.join(base_dir, "app.db")


engine = create_engine(f"sqlite:///{get_db_path()}", echo=False)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


def init_db():
    Base.metadata.create_all(bind=engine)
    session = SessionLocal()
    if not session.query(AppSettings).first():
        session.add(AppSettings(server_port=8765, interval_minutes=15))
        session.commit()
    session.close()
