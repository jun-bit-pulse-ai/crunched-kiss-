import logging

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

from app.agent import run_turn
from app.limits import reject_content_length
from app.models import ChatRequest, ChatResponse

logger = logging.getLogger(__name__)

app = FastAPI(title="Crunched KISS")


@app.middleware("http")
async def reject_oversized_body(request: Request, call_next):
    status = reject_content_length(request.headers.get("content-length"))
    if status == 413:
        return JSONResponse(status_code=413, content={"detail": "Request body too large"})
    if status == 400:
        return JSONResponse(status_code=400, content={"detail": "Invalid Content-Length"})
    return await call_next(request)


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
