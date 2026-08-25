# -*- coding: utf-8 -*-
"""FastAPI 路由：导入检索（Word 上传解析 + 17 节父子级展示）。"""
from __future__ import annotations

import json
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

from fastapi import APIRouter, File, HTTPException, UploadFile

ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
sys.path.insert(0, str(ROOT / "02-检索系统"))

router = APIRouter(prefix="/api/import", tags=["import"])

MAX_UPLOAD_MB = 20


def _node_to_dict(node: Any) -> dict[str, Any]:
    """把 SectionNode / BigTitleNode / FieldNode 转成可 JSON 序列化字典。"""
    out: dict[str, Any] = {}
    for attr in ("number", "title", "full_title", "seq", "label", "value",
                 "std_name", "kind", "editable", "index", "unit"):
        if hasattr(node, attr):
            v = getattr(node, attr)
            if v is not None and not callable(v):
                out[attr] = v
    if hasattr(node, "big_titles"):
        out["big_titles"] = [_node_to_dict(x) for x in node.big_titles]
    if hasattr(node, "children"):
        out["children"] = [_node_to_dict(x) for x in node.children]
    if hasattr(node, "direct_fields"):
        out["direct_fields"] = [_node_to_dict(x) for x in node.direct_fields]
    if hasattr(node, "rows") and isinstance(getattr(node, "rows", None), list):
        rows = []
        for r in node.rows:
            if hasattr(r, "label") and hasattr(r, "value"):
                rows.append({"label": r.label, "value": r.value,
                             "std_name": getattr(r, "std_name", ""),
                             "kind": getattr(r, "kind", "field"),
                             "editable": getattr(r, "editable", 0)})
            elif isinstance(r, dict):
                rows.append(r)
        out["rows"] = rows
    if hasattr(node, "values") and isinstance(getattr(node, "values", None), dict):
        out["values"] = dict(node.values)
    return out


@router.post("/upload")
async def import_upload(file: UploadFile = File(...)) -> dict[str, Any]:
    """上传 MSDS Word 文档，解析出 17 节父子级结构。"""
    name = (file.filename or "").strip()
    if not name.lower().endswith(".docx"):
        raise HTTPException(400, "仅支持 .docx 格式的 MSDS Word 文档")

    tmpdir = Path(tempfile.mkdtemp(prefix="msds_import_"))
    tmp = tmpdir / name
    try:
        with tmp.open("wb") as fh:
            shutil.copyfileobj(file.file, fh)
        if tmp.stat().st_size > MAX_UPLOAD_MB * 1024 * 1024:
            raise HTTPException(413, f"文件超过 {MAX_UPLOAD_MB}MB 限制")

        from core.docx_reader import read_msds
        from core.extract import build_hierarchy

        result = read_msds(tmp)
        sections = build_hierarchy(result)
        payload = {
            "ok": True,
            "filename": name,
            "model": getattr(result, "model", ""),
            "sections": [_node_to_dict(s) for s in sections],
        }
        return payload
    except HTTPException:
        raise
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(500, f"解析失败: {exc}") from exc
    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:  # noqa: BLE001
            pass
