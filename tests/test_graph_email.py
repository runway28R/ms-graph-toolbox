import base64
from types import SimpleNamespace
from unittest.mock import Mock

from ms_graph_toolbox import graph_email


def test_parse_recipients():
    result = graph_email.parse_recipients(
        " first@example.com, second@example.com "
    )

    assert result == [
        {"emailAddress": {"address": "first@example.com"}},
        {"emailAddress": {"address": "second@example.com"}},
    ]


def test_parse_recipients_with_empty_value():
    assert graph_email.parse_recipients(None) == []
    assert graph_email.parse_recipients("") == []


def test_build_attachment_from_bytes():
    logger = Mock()

    result = graph_email.build_attachment(
        {
            "content_bytes": b"hello",
            "name": "hello.txt",
            "content_type": "text/plain",
        },
        logger,
    )

    assert result == {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": "hello.txt",
        "contentType": "text/plain",
        "contentBytes": base64.b64encode(b"hello").decode("utf-8"),
    }


def test_build_inline_attachment():
    logger = Mock()

    result = graph_email.build_attachment(
        {
            "content_bytes": b"image data",
            "name": "image.png",
            "content_type": "image/png",
            "inline": True,
            "content_id": "logo",
        },
        logger,
    )

    assert result["isInline"] is True
    assert result["contentId"] == "logo"


def test_build_attachment_from_file(tmp_path):
    logger = Mock()
    file_path = tmp_path / "document.txt"
    file_path.write_text("hello", encoding="utf-8")

    result = graph_email.build_attachment(
        {"path": str(file_path)},
        logger,
    )

    assert result["name"] == "document.txt"
    assert result["contentBytes"] == base64.b64encode(
        b"hello"
    ).decode("utf-8")


def test_build_attachment_without_source_returns_none():
    logger = Mock()

    result = graph_email.build_attachment(
        {"name": "missing.txt"},
        logger,
    )

    assert result is None
    logger.error.assert_called_once()


def test_send_email_success(monkeypatch):
    logger = Mock()
    graph = SimpleNamespace(access_token="test-token", logger=logger)
    captured_request = {}

    def fake_post(endpoint, headers, json):
        captured_request["endpoint"] = endpoint
        captured_request["headers"] = headers
        captured_request["json"] = json
        return SimpleNamespace(status_code=202, text="")

    monkeypatch.setattr(graph_email.requests, "post", fake_post)

    result = graph_email.send_email(
        graph,
        subject="Test subject",
        content_type="text",
        body="Test body",
        sender="sender@example.com",
        to_field="recipient@example.com",
    )

    assert result == 0
    assert captured_request["endpoint"] == (
        "https://graph.microsoft.com/v1.0/users/"
        "sender@example.com/sendMail"
    )
    assert captured_request["headers"]["Authorization"] == "Bearer test-token"

    message = captured_request["json"]["message"]
    assert message["subject"] == "Test subject"
    assert message["body"] == {
        "contentType": "Text",
        "content": "Test body",
    }
    assert message["toRecipients"] == [
        {"emailAddress": {"address": "recipient@example.com"}}
    ]


def test_send_email_without_access_token():
    logger = Mock()
    graph = SimpleNamespace(access_token=None, logger=logger)

    result = graph_email.send_email(
        graph,
        "Subject",
        "Text",
        "Body",
        "sender@example.com",
        "recipient@example.com",
    )

    assert result == 2
    logger.error.assert_called_once()