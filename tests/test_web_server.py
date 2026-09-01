from fastapi.testclient import TestClient
from src.server.web_server import app, get_db


def test_create_and_list_clients(db_session):
    def override_get_db():
        try:
            yield db_session
        finally:
            pass

    app.dependency_overrides[get_db] = override_get_db
    client = TestClient(app)

    payload = {
        "name": "Test",
        "domain": "test.atlassian.net",
        "email": "user@test.com.ar",
        "api_token": "token123",
        "time_window_start": "08:00",
        "time_window_end": "18:00",
        "active_weekdays_only": True,
    }

    res_post = client.post("/api/clients", json=payload)
    assert res_post.status_code == 200

    res_get = client.get("/api/clients")
    assert res_get.status_code == 200
    data = res_get.json()
    assert len(data) == 1
    assert data[0]["name"] == "Test"
