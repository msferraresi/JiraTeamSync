from unittest.mock import MagicMock
from src.services.jira_service import JiraService
from src.db.models import Client


def test_fetch_board_sprints(monkeypatch):
    client = Client(
        domain="fpatronal.atlassian.net", email="test@test.com", api_token="dummy_token"
    )
    service = JiraService(client)

    mock_response = MagicMock()
    mock_response.status_code = 200
    mock_response.json.return_value = {
        "values": [
            {"id": 1, "name": "Sprint 16", "state": "closed"},
            {"id": 2, "name": "Sprint 17", "state": "closed"},
            {"id": 3, "name": "Sprint 18", "state": "active"},
        ]
    }

    monkeypatch.setattr("requests.get", lambda *args, **kwargs: mock_response)

    sprints = service.fetch_board_sprints(284)
    assert len(sprints) == 3
    assert sprints[-1]["name"] == "Sprint 18"
