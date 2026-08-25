# -*- coding: utf-8 -*-
"""FastAPI 路由：分库总览（库列表、树下钻、通用表查询）。"""
from __future__ import annotations

import sqlite3
from typing import Any

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from .. import business_tree, lib_discovery

router = APIRouter(prefix="/api/libraries", tags=["libraries"])


class LibRow(BaseModel):
    columns: list[str]
    total: int
    page: int
    page_size: int
    rows: list[dict[str, Any]]


def _open(key: str) -> tuple[sqlite3.Connection, dict[str, Any]]:
    lib = business_tree.resolve_library(key)
    if not lib:
        raise HTTPException(404, f"未知库: {key}")
    try:
        conn = lib_discovery.open_readonly(lib["path"])
    except sqlite3.Error as e:
        raise HTTPException(500, f"打开库失败: {e}") from e
    return conn, lib


@router.get("")
@router.get("/")
def list_libraries() -> list[dict[str, Any]]:
    """库清单（自动发现 + 配置，含表与行数）。"""
    return lib_discovery.discover_libraries()


@router.get("/{key}/tables")
def library_tables(key: str) -> dict[str, Any]:
    conn, lib = _open(key)
    conn.close()
    return {"key": key, "name": lib["name"], "tables": lib["tables"],
            "business_tree": lib["business_tree"]}


# ---------------------------------------------------------------------------
# 通用表查询（库 → 表 → 行 → 全字段）
# ---------------------------------------------------------------------------
@router.get("/{key}/tables/{table}/rows", response_model=LibRow)
def generic_rows(key: str, table: str, page: int = 1, page_size: int = 50,
                 keyword: str = "") -> dict[str, Any]:
    conn, _ = _open(key)
    try:
        return business_tree.table_rows(conn, table, page, page_size, keyword)
    finally:
        conn.close()


@router.get("/{key}/tables/{table}/rows/{rowid}")
def generic_row_detail(key: str, table: str, rowid: int) -> dict[str, Any]:
    conn, _ = _open(key)
    try:
        r = business_tree.row_detail(conn, table, rowid)
        if r is None:
            raise HTTPException(404, "行不存在")
        return r
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# 总型号库业务树：大类 → 型号 → 节 → 字段
# ---------------------------------------------------------------------------
@router.get("/models/categories")
def models_categories() -> list[dict[str, Any]]:
    conn, _ = _open("msds_standard.db")
    try:
        return business_tree.categories(conn)
    finally:
        conn.close()


@router.get("/models/categories/{parent}/models")
def models_by_category(parent: str) -> list[dict[str, Any]]:
    conn, _ = _open("msds_standard.db")
    try:
        return business_tree.models_by_category(conn, parent)
    finally:
        conn.close()


@router.get("/models/{model_id}/summary")
def models_summary(model_id: int) -> dict[str, Any]:
    conn, _ = _open("msds_standard.db")
    try:
        s = business_tree.model_summary(conn, model_id)
        if s is None:
            raise HTTPException(404, "型号不存在")
        return s
    finally:
        conn.close()


@router.get("/models/{model_id}/sections")
def models_sections(model_id: int) -> list[dict[str, Any]]:
    conn, _ = _open("msds_standard.db")
    try:
        return business_tree.model_sections(conn, model_id)
    finally:
        conn.close()


@router.get("/models/{model_id}/sections/{section}")
def models_section_fields(model_id: int, section: int) -> list[dict[str, Any]]:
    conn, _ = _open("msds_standard.db")
    try:
        return business_tree.section_fields(conn, model_id, section)
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# CAS 库业务树：物质 → 别名 / 关联型号
# ---------------------------------------------------------------------------
@router.get("/cas/substances")
def cas_list(keyword: str = "") -> list[dict[str, Any]]:
    conn, _ = _open("cas_library.db")
    try:
        return business_tree.cas_substances(conn, keyword)
    finally:
        conn.close()


@router.get("/cas/substances/{cas_id}")
def cas_detail(cas_id: int) -> dict[str, Any]:
    conn, _ = _open("cas_library.db")
    msds_conn, _ = _open("msds_standard.db")
    try:
        d = business_tree.cas_substance_detail(conn, cas_id, msds_conn)
        if d is None:
            raise HTTPException(404, "物质不存在")
        return d
    finally:
        conn.close()
        msds_conn.close()
