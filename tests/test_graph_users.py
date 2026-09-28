from types import SimpleNamespace
from unittest.mock import Mock

from ms_graph_toolbox import graph_users


def make_graph_object():
    return SimpleNamespace(
        access_token="test-token",
        logger=Mock(),
    )


def test_get_users_builds_filters_and_selects_fields(monkeypatch):
    graph = make_graph_object()
    captured = {}

    def fake_get(url, headers, params=None):
        captured["url"] = url
        captured["headers"] = headers
        captured["params"] = params
        return SimpleNamespace(
            status_code=200,
            json=lambda: {
                "value": [
                    {
                        "displayName": "Alex Smith",
                        "companyName": "Example Corp",
                    }
                ]
            },
        )

    monkeypatch.setattr(graph_users.requests, "get", fake_get)

    result = graph_users.get_users(
        graph,
        select_data=["displayName", "mail"],
        search_name="Alex",
        search_company="example",
    )

    assert result == [
        {
            "displayName": "Alex Smith",
            "companyName": "Example Corp",
        }
    ]

    assert captured["url"] == "https://graph.microsoft.com/v1.0/users"
    assert captured["headers"] == {
        "Authorization": "Bearer test-token",
        "ConsistencyLevel": "eventual",
    }
    assert captured["params"]["$select"] == "displayName,mail"
    assert captured["params"]["$count"] == "true"
    assert captured["params"]["$filter"] == "startswith(displayName,'Alex')"


def test_get_users_applies_company_filter_case_insensitively(monkeypatch):
    graph = make_graph_object()

    def fake_get(url, headers, params=None):
        return SimpleNamespace(
            status_code=200,
            json=lambda: {
                "value": [
                    {"displayName": "Alex", "companyName": "Example Corporation"},
                    {"displayName": "Sam", "companyName": "Other Company"},
                    {"displayName": "Taylor"},
                ]
            },
        )

    monkeypatch.setattr(graph_users.requests, "get", fake_get)

    result = graph_users.get_users(
        graph,
        search_company="EXAMPLE",
    )

    assert result == [
        {
            "displayName": "Alex",
            "companyName": "Example Corporation",
        }
    ]


def test_get_users_follows_next_link(monkeypatch):
    graph = make_graph_object()
    calls = []

    def fake_get(url, headers, params=None):
        calls.append((url, params))

        if len(calls) == 1:
            return SimpleNamespace(
                status_code=200,
                json=lambda: {
                    "value": [{"displayName": "First User"}],
                    "@odata.nextLink": "https://graph.microsoft.com/next-page",
                },
            )

        return SimpleNamespace(
            status_code=200,
            json=lambda: {
                "value": [{"displayName": "Second User"}],
            },
        )

    monkeypatch.setattr(graph_users.requests, "get", fake_get)

    result = graph_users.get_users(graph)

    assert result == [
        {"displayName": "First User"},
        {"displayName": "Second User"},
    ]
    assert calls[0][1]["$count"] == "true"
    assert calls[1] == (
        "https://graph.microsoft.com/next-page",
        None,
    )


def test_get_users_returns_none_for_http_error(monkeypatch):
    graph = make_graph_object()

    def fake_get(url, headers, params=None):
        return SimpleNamespace(
            status_code=403,
            text="Forbidden",
        )

    monkeypatch.setattr(graph_users.requests, "get", fake_get)

    result = graph_users.get_users(
        graph,
        search_name="Alex",
    )

    assert result is None
    graph.logger.error.assert_called_once()