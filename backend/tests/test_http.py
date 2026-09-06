from fastapi.testclient import TestClient

from app.main import app


client = TestClient(app)


def test_health() -> None:
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json() == {"ok": True}


def test_chat_without_api_key_returns_error_payload(monkeypatch) -> None:
    monkeypatch.setattr("app.agent.settings.anthropic_api_key", "")
    response = client.post("/chat", json={"messages": [{"role": "user", "content": "hi"}]})
    assert response.status_code == 200
    body = response.json()
    assert body["type"] == "error"
    assert "ANTHROPIC_API_KEY" in body["message"]
