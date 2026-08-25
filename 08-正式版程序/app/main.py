# -*- coding: utf-8 -*-
"""FastAPI 主入口：分库总览 Web 服务。"""
from __future__ import annotations

import os
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .routers import import_search, libraries

BASE_DIR = Path(__file__).resolve().parent.parent  # 08-正式版程序

app = FastAPI(title="MSDS 分库总览", version="0.1.0")

# 开发期允许 Vite 前端跨域访问
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173",
                   "http://localhost:8765", "http://127.0.0.1:8765"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(libraries.router)
app.include_router(import_search.router)


@app.get("/api/health")
def health():
    return {"ok": True, "service": "msds-library-overview"}


# 若前端已构建（frontend/dist 存在），托管静态资源
_dist = BASE_DIR / "frontend" / "dist"
if _dist.is_dir():
    app.mount("/", StaticFiles(directory=str(_dist), html=True), name="frontend")


def main():
    import uvicorn

    port = int(os.environ.get("PORT", "8766"))
    uvicorn.run(app, host="127.0.0.1", port=port)


if __name__ == "__main__":
    main()
