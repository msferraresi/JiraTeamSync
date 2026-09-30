from unittest.mock import MagicMock
from src.services.jira_service import JiraService
from src.db.models import Client
from src.utils.jira_utils import (
    sanitize_and_prepare_jql,
    extract_story_points,
    extract_user_logged_hours,
)
from src.utils.date_utils import format_all_day_range
from datetime import date


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


def test_sanitize_and_prepare_jql():
    jql = "assignee = currentUser() ORDER BY created DESC"
    result = sanitize_and_prepare_jql(jql, force=False)
    assert (
        "(assignee = currentUser()) AND (statusCategory != Done OR resolutiondate >= -30d) ORDER BY created DESC"
        == result
    )

    result_force = sanitize_and_prepare_jql(jql, force=True)
    assert "assignee = currentUser() ORDER BY created DESC" == result_force


def test_extract_user_logged_hours():
    fields = {
        "worklog": {
            "worklogs": [
                {"author": {"emailAddress": "user@test.com"}, "timeSpentSeconds": 7200},
                {
                    "author": {"emailAddress": "other@test.com"},
                    "timeSpentSeconds": 3600,
                },
            ]
        }
    }
    hours = extract_user_logged_hours(fields, current_user_email="user@test.com")
    assert hours == 2.0


def test_format_all_day_range():
    d_start = date(2026, 9, 1)
    d_end = date(2026, 9, 3)
    start_str, end_str = format_all_day_range(d_start, d_end)
    assert start_str == "2026-09-01 00:00"
    assert end_str == "2026-09-04 00:00"
