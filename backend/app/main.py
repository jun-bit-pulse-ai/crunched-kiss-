import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.agent import run_turn
from app.config import settings
from app.models import ChatRequest, ChatResponse

logger = logging.getLogger(__name__)

app = FastAPI(title="Crunched KISS")
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_methods=["POST", "GET", "OPTIONS"],
    allow_headers=["Content-Type"],
)


@app.get("/health")
@app.get("/api/health")
def health() -> dict[str, bool]:
    return {"ok": True}


@app.post("/chat", response_model=ChatResponse)
@app.post("/api/chat", response_model=ChatResponse)
def chat(request: ChatRequest) -> ChatResponse:
    try:
        return run_turn(request)
    except Exception:
        logger.exception("chat turn failed")
        return ChatResponse(type="error", message="The model request failed. Check server logs.")
