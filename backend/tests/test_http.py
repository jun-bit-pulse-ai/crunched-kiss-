from fastapi.testclient import TestClient

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
