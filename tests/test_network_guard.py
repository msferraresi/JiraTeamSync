from datetime import datetime
from src.services.network_guard import NetworkGuard
from src.db.models import Client


def test_time_window_allowed(monkeypatch):
    client = Client(
        name="FedPat",
        time_window_start="07:00",
        time_window_end="20:00",
        active_weekdays_only=True,
    )

    class MockDateTime(datetime):
        @classmethod
        def now(cls):
            return datetime(2026, 9, 2, 10, 30)  # Miércoles

    monkeypatch.setattr("src.services.network_guard.datetime", MockDateTime)
    assert NetworkGuard.is_within_time_window(client) is True


def test_time_window_rejected_weekend(monkeypatch):
    client = Client(
        name="FedPat",
        time_window_start="07:00",
        time_window_end="20:00",
        active_weekdays_only=True,
    )

    class MockDateTime(datetime):
        @classmethod
        def now(cls):
            return datetime(2026, 9, 5, 11, 0)  # Sábado

    monkeypatch.setattr("src.services.network_guard.datetime", MockDateTime)
    assert NetworkGuard.is_within_time_window(client) is False


def test_can_sync_client_disabled():
    client = Client(name="Test", is_active=False)
    can_sync, reason = NetworkGuard.can_sync_client(client)
    assert can_sync is False
    assert "deshabilitado" in reason.lower()
