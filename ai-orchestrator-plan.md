# Multi-Tier AI Agent System — Implementation Plan

## Overview
Build a lightweight orchestration layer that routes tasks between lower-tier models (fast/cheap inference) and Claude Code agents (complex implementation), with a 4-hour implementation target.

---

## Architecture (5-Min Read)

```
┌─────────────────────────────────────────────────────────────────┐
│                        API Gateway                               │
│              (FastAPI / Express — single endpoint)              │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        ▼                     ▼                     ▼
┌───────────────┐    ┌─────────────────┐    ┌───────────────┐
│  Task Router  │    │  Lower-Tier     │    │ Claude Code   │
│  (Classifier) │───▶│  Model Worker   │    │ Agent Worker  │
│               │◄───│  (Fast/ Cheap)  │    │ (Complex Code)│
└───────────────┘    └─────────────────┘    └───────────────┘
        │                                              │
        └─────────────────────┬────────────────────────┘
                              ▼
                    ┌─────────────────┐
                    │  Result Merge   │
                    │  & Response     │
                    └─────────────────┘
```

### Core Components

| Component | Tech | Purpose |
|-----------|------|---------|
| **API Gateway** | FastAPI (Python) | Single `/execute` endpoint, request validation |
| **Task Router** | Small classifier (GPT-4o-mini / local Llama) | Determines task tier in <500ms |
| **Lower-Tier Worker** | OpenRouter / Ollama / vLLM | Handles: summarization, classification, Q&A, simple extraction |
| **Claude Code Agent** | `claude-code` CLI via subprocess | Handles: code generation, refactoring, debugging, architecture |
| **Result Merge** | Python merge layer | Combines outputs, formats final response |
| **State Store** | SQLite (local) or Redis | Tracks job status, caches tier decisions |

---

## Task Classification Rules (Router Logic)

| Tier | Task Types | Model | Latency |
|------|-----------|-------|---------|
| **L1 — Lower-Tier** | Text summarization, sentiment analysis, keyword extraction, simple Q&A, data formatting, translation | GPT-4o-mini / Llama 3.1 8B | <2s |
| **L2 — Claude Code** | Code generation, debugging, refactoring, architecture decisions, multi-file operations, test writing | Claude 3.5 Sonnet (via `claude-code`) | 10-60s |

### Router Decision Tree
```python
def route_task(task: Task) -> Tier:
    if task.type in [CODE_GEN, DEBUG, REFACTOR, ARCHITECTURE]:
        return CLAUDE_CODE
    if task.has_code_context or task.files_affected > 1:
        return CLAUDE_CODE
    if task.complexity_score > 0.7:  # Determined by mini-classifier
        return CLAUDE_CODE
    return LOWER_TIER
```

---

## File Structure (30 min setup)

```
ai-orchestrator/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI app
│   ├── router.py            # Task classification logic
│   ├── models.py            # Pydantic schemas
│   ├── config.py            # Env vars, API keys
│   ├── workers/
│   │   ├── __init__.py
│   │   ├── lower_tier.py    # OpenRouter/Ollama client
│   │   └── claude_code.py   # Claude Code subprocess wrapper
│   └── utils/
│       └── cache.py         # SQLite state tracking
├── scripts/
│   └── setup_claude.sh      # One-time Claude Code setup
├── tests/
│   └── test_router.py
├── requirements.txt
├── .env.example
└── README.md
```

---

## Implementation Phases (4-Hour Breakdown)

### Hour 1: Scaffold & Router
- [ ] **0:00–0:15** — Project setup: `uv init`, install FastAPI, uvicorn, pydantic, httpx, sqlalchemy
- [ ] **0:15–0:30** — Define Pydantic models: `TaskRequest`, `TaskResponse`, `TierDecision`
- [ ] **0:30–0:50** — Implement Task Router: mini-classifier prompt using GPT-4o-mini
- [ ] **0:50–1:00** — Test router with 10 sample tasks, tune thresholds

### Hour 2: Lower-Tier Worker
- [ ] **1:00–1:20** — Build `LowerTierWorker`: OpenRouter client with async httpx
- [ ] **1:20–1:40** — Add caching: SQLite store for tier decisions (avoid re-classifying)
- [ ] **1:40–1:50** — Error handling, retries, fallback to Claude Code on failure
- [ ] **1:50–2:00** — Test with 5 L1 tasks, verify <2s response time

### Hour 3: Claude Code Integration
- [ ] **2:00–2:15** — Install Claude Code CLI (`npm install -g @anthropics/claude-code`)
- [ ] **2:15–2:45** — Build `ClaudeCodeWorker`: subprocess wrapper with:
  - Temp file for context passing
  - `--allowed-tools` flag restriction
  - Timeout handling (30s default, 120s max)
  - Output capture and parsing
- [ ] **2:45–2:55** — Safety: read-only mode option, sandbox directory restriction
- [ ] **2:55–3:00** — Test with 3 L2 tasks (code gen, debug, refactor)

### Hour 4: Integration & Polish
- [ ] **3:00–3:20** — Wire everything in `main.py`: `/execute` endpoint, async queue
- [ ] **3:20–3:35** — Result merger: format responses consistently, add metadata (tier used, latency, cost)
- [ ] **3:35–3:50** — Health check endpoint, basic tests, README
- [ ] **3:50–4:00** — Run end-to-end: 5 mixed tasks, verify routing accuracy >90%

---

## Key Implementation Details

### 1. Router Prompt (GPT-4o-mini)
```
You are a task classifier. Given a user request, classify it as either:
- "lower_tier": Simple tasks (summarize, Q&A, extract keywords, translate)
- "claude_code": Complex tasks (write code, debug, refactor, architecture)

Respond ONLY with {"tier": "lower_tier|claude_code", "confidence": 0.0-1.0}

User request: {task_description}
```

### 2. Claude Code Subprocess Wrapper
```python
import subprocess, tempfile, os

async def run_claude_code(prompt: str, context_files: list = None):
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False) as f:
        f.write(prompt)
        prompt_file = f.name
    
    cmd = [
        "claude-code",
        "--prompt", prompt_file,
        "--allowed-tools", "Edit,Bash,Read,Glob",  # Restrict tools
        "--max-turns", "10",
        "--cwd", "/sandbox/project"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    os.unlink(prompt_file)
    return result.stdout
```

### 3. Cost Optimization
| Strategy | Implementation |
|----------|---------------|
| **Router Caching** | Cache tier decisions in SQLite keyed by task hash |
| **Lower-Tier Default** | Default to lower-tier, escalate only on confidence |
| **Batching** | Group small L1 tasks into single API call |
| **Timeout Fallback** | If Claude Code >30s, return partial + async continuation |

### 4. Safety Guardrails
- Claude Code runs in `--read-only` mode unless explicitly enabled
- Restricted to project directory via `--cwd`
- Tool allowlist: only `Read`, `Edit`, `Bash`, `Glob`
- Max 10 turns, 120s timeout
- Lower-tier models: no file system access

---

## API Contract

### Request
```json
POST /execute
{
  "task_id": "uuid",
  "description": "Refactor this Python function to use async",
  "context": {
    "files": ["/path/to/file.py"],
    "code_snippet": "def old_func(): ..."
  },
  "options": {
    "tier_override": null,
    "read_only": false,
    "timeout_seconds": 60
  }
}
```

### Response
```json
{
  "task_id": "uuid",
  "tier_used": "claude_code",
  "result": "Refactored code...",
  "metadata": {
    "routing_confidence": 0.95,
    "latency_ms": 15420,
    "estimated_cost_usd": 0.08,
    "tools_used": ["Read", "Edit"]
  }
}
```

---

## Post-4-Hour Enhancements (Backlog)

1. **Streaming responses** for long-running Claude Code tasks
2. **Multi-agent collaboration**: Lower-tier pre-process → Claude Code → Lower-tier review
3. **Feedback loop**: Log routing accuracy, retrain classifier
4. **Web UI**: Simple dashboard for monitoring task distribution
5. **Cost tracking**: Per-project billing, budget caps

---

## Success Metrics

| Metric | Target |
|--------|--------|
| Router accuracy | >90% |
| L1 latency | <2s p95 |
| L2 latency | <60s p95 |
| System uptime | 100% (local) |
| End-to-end test pass | 5/5 mixed tasks |

---

## Handoff Checklist for Claude Code Agents

1. ✅ Architecture diagram approved
2. ✅ File structure defined
3. ✅ API contracts documented
4. ✅ Implementation phases with timeboxes
5. ✅ Key code snippets provided
6. ✅ Test criteria defined
7. ✅ Safety guardrails specified

**Start with Hour 1 scaffold. Each phase has defined done criteria. Good luck.**
