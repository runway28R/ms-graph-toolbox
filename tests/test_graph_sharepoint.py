from unittest.mock import Mock

from ms_graph_toolbox import graph_sharepoint


def make_sharepoint():
    return graph_sharepoint.graph_sharepoint(
        access_token="test-token",
        logger=Mock(),
    )


def test_get_site_id(monkeypatch):
    sharepoint = make_sharepoint()
    captured = {}

    def fake_get(url, headers):
        captured["url"] = url
        captured["headers"] = headers
        return type(
            "Response",
            (),
            {
                "status_code": 200,
                "text": "OK",
                "json": lambda self: {"id": "site-123"},
            },
        )()

    monkeypatch.setattr(graph_sharepoint.requests, "get", fake_get)

    result = sharepoint.get_site_id("contoso.sharepoint.com:/sites/demo")

    assert result == "site-123"
    assert captured["url"] == (
        "https://graph.microsoft.com/v1.0/sites/"
        "contoso.sharepoint.com:/sites/demo"
    )
    assert captured["headers"] == {
        "Authorization": "Bearer test-token",
    }


def test_get_document_libraries(monkeypatch):
    sharepoint = make_sharepoint()

    def fake_get(url, headers):
        return type(
            "Response",
            (),
            {
                "json": lambda self: {
                    "value": [
                        {"id": "drive-1", "name": "Documents"},
                        {"id": "drive-2", "name": "Images"},
                    ]
                }
            },
        )()

    monkeypatch.setattr(graph_sharepoint.requests, "get", fake_get)

    result = sharepoint.get_document_libraries("site-123")

    assert result == [
        ("drive-1", "Documents"),
        ("drive-2", "Images"),
    ]


def test_get_folder_content(monkeypatch):
    sharepoint = make_sharepoint()

    def fake_get(url, headers):
        return type(
            "Response",
            (),
            {
                "json": lambda self: {
                    "value": [
                        {"name": "report.pdf", "file": {}},
                        {"name": "Archive", "folder": {}},
                    ]
                }
            },
        )()

    monkeypatch.setattr(graph_sharepoint.requests, "get", fake_get)

    result = sharepoint.get_folder_content("site-123", "drive-456")

    assert result == [
        {"name": "report.pdf", "file": {}},
        {"name": "Archive", "folder": {}},
    ]


def test_print_folder_content(capsys):
    sharepoint = make_sharepoint()

    sharepoint.print_folder_content(
        [
            {"name": "z-folder", "folder": {}},
            {"name": "a-folder", "folder": {}},
            {"name": "z-file.txt", "file": {}},
            {"name": "a-file.txt", "file": {}},
        ]
    )

    output = capsys.readouterr().out

    assert "Folders: 2" in output
    assert "Files: 2" in output
    assert output.index("a-folder") < output.index("z-folder")
    assert output.index("a-file.txt") < output.index("z-file.txt")


def test_upload_file_graph_success(monkeypatch, tmp_path):
    sharepoint = make_sharepoint()
    local_file = tmp_path / "report.txt"
    local_file.write_bytes(b"report contents")

    captured = {}

    def fake_put(url, headers, data):
        captured["url"] = url
        captured["headers"] = headers
        captured["data"] = data

        return type(
            "Response",
            (),
            {
                "status_code": 201,
                "json": lambda self: {
                    "webUrl": "https://contoso.sharepoint.com/report.txt"
                },
            },
        )()

    monkeypatch.setattr(graph_sharepoint.requests, "put", fake_put)

    result = sharepoint.upload_file_graph(
        site_id="site-123",
        drive_id="drive-456",
        folder_path="Reports/2026",
        local_file_path=str(local_file),
    )

    assert result == (
        "https://contoso.sharepoint.com/report.txt",
        1,
    )
    assert captured["data"] == b"report contents"
    assert captured["headers"] == {
        "Authorization": "Bearer test-token",
        "Content-Type": "application/octet-stream",
    }
    assert captured["url"].endswith(
        "/root:/Reports/2026/report.txt:/content"
    )


def test_upload_file_graph_missing_file(tmp_path):
    sharepoint = make_sharepoint()
    missing_file = tmp_path / "missing.txt"

    result = sharepoint.upload_file_graph(
        site_id="site-123",
        drive_id="drive-456",
        folder_path="Reports",
        local_file_path=str(missing_file),
    )

    assert result[1] == 0
    assert "File not found" in result[0]
