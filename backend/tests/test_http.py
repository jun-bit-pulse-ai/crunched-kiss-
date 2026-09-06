from fastapi.testclient import TestClient

from app.limits import MAX_CHAT_MESSAGES, MAX_MESSAGE_CHARS
from app.main import app


client = TestClient(app)


def test_health() -> None:
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_api_health() -> None:
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_chat_without_api_key_returns_error_payload(monkeypatch) -> None:
    monkeypatch.setattr("app.agent.settings.anthropic_api_key", "")
    for path in ("/chat", "/api/chat"):
        response = client.post(path, json={"messages": [{"role": "user", "content": "hi"}]})
        assert response.status_code == 200
        body = response.json()
        assert body["type"] == "error"
        assert "ANTHROPIC_API_KEY" in body["message"]


def test_chat_rejects_empty_messages() -> None:
    response = client.post("/api/chat", json={"messages": []})
    assert response.status_code == 422


def test_chat_rejects_too_many_messages() -> None:
    messages = [{"role": "user", "content": "x"} for _ in range(MAX_CHAT_MESSAGES + 1)]
    response = client.post("/api/chat", json={"messages": messages})
    assert response.status_code == 422


def test_chat_rejects_oversized_message_content() -> None:
    response = client.post(
        "/api/chat",
        json={"messages": [{"role": "user", "content": "x" * (MAX_MESSAGE_CHARS + 1)}]},
    )
    assert response.status_code == 422


def test_chat_does_not_grant_cors_to_other_origins() -> None:
    response = client.options(
        "/api/chat",
        headers={
            "Origin": "https://evil.example",
            "Access-Control-Request-Method": "POST",
        },
    )
    assert "access-control-allow-origin" not in {key.lower() for key in response.headers}
