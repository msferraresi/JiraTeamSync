from sqlalchemy import Column, Integer, String, Boolean, Text, ForeignKey
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()


class Client(Base):
    __tablename__ = "clients"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(100), nullable=False, unique=True)
    domain = Column(String(200), nullable=False)
    email = Column(String(200), nullable=False)
    api_token = Column(String(300), nullable=False)

    is_active = Column(Boolean, default=True)
    time_window_start = Column(String(5), default="07:00")
    time_window_end = Column(String(5), default="20:00")
    active_weekdays_only = Column(Boolean, default=True)
    vpn_check_host = Column(String(200), nullable=True)

    boards = relationship(
        "BoardConfig", back_populates="client", cascade="all, delete-orphan"
    )


class BoardConfig(Base):
    __tablename__ = "board_configs"

    id = Column(Integer, primary_key=True, autoincrement=True)
    client_id = Column(Integer, ForeignKey("clients.id"), nullable=False)

    board_id = Column(Integer, nullable=True)
    board_name = Column(String(150), nullable=False)
    is_enabled = Column(Boolean, default=True)
    custom_jql = Column(
        Text, default="assignee = currentUser() AND statusCategory != Done"
    )
    hours_per_sp = Column(Integer, default=4)

    # Configuración dinámica en formato JSON:
    # [{"field_id": "customfield_10015", "field_name": "Start date", "role": "real_start", "include_in_body": True}, ...]
    fields_mapping_json = Column(Text, default="[]")

    client = relationship("Client", back_populates="boards")


class AppSettings(Base):
    __tablename__ = "app_settings"
    id = Column(Integer, primary_key=True)
    server_port = Column(Integer, default=8765)
    interval_minutes = Column(Integer, default=15)
