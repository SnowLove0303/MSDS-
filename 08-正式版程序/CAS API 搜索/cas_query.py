#!/usr/bin/env python3
"""Lightweight CAS/name lookup client for PubChem PUG REST.

Examples:
    python cas_query.py "丙烯酸"
    python cas_query.py "9003-01-4"
    python cas_query.py "丙烯酸" --json

The returned CAS values are API matches, not a guarantee of CAS Registry
authority. PubChem documents that CAS coverage is incomplete and may contain
third-party supplied identifiers.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen


API_ROOT = "https://pubchem.ncbi.nlm.nih.gov/rest/pug"
WIKIDATA_API = "https://www.wikidata.org/w/api.php"
WIKIDATA_SPARQL = "https://query.wikidata.org/sparql"
EPA_SRS_SEARCH = "https://cdxapps.epa.gov/oms-substance-registry-services/search/results"
NCI_RESOLVER = "https://cactus.nci.nih.gov/chemical/structure"
CAS_RE = re.compile(r"^\d{2,7}-\d{2}-\d$")
CAS_ANY_RE = re.compile(r"^\d{2,7}-\d{2}-\d$")
DEFAULT_CACHE = Path(__file__).with_name(".cas_cache")
DEFAULT_TIMEOUT = 12
DEFAULT_RETRIES = 2
USER_AGENT = "cas-query/1.0 (local materials research tool)"


class QueryError(RuntimeError):
    """A user-facing query or network error."""


def is_cas(value: str) -> bool:
    return bool(CAS_RE.fullmatch(value.strip()))


def cache_key(url: str) -> Path:
    digest = hashlib.sha256(url.encode("utf-8")).hexdigest()
    return Path(digest + ".json")


def get_json(url: str, cache_dir: Path, timeout: int, use_cache: bool) -> Any:
    cache_dir.mkdir(parents=True, exist_ok=True)
    cache_path = cache_dir / cache_key(url)
    if use_cache and cache_path.exists():
        try:
            cached = json.loads(cache_path.read_text(encoding="utf-8"))
            if isinstance(cached, dict):
                return cached
            cache_path.unlink(missing_ok=True)
        except (OSError, json.JSONDecodeError):
            cache_path.unlink(missing_ok=True)

    # On some Windows/Anaconda environments urllib's TLS handshake to the
    # Wikidata query host times out while the system curl path works. Keep the
    # transport free and keyless, but use curl.exe as a narrow fallback.
    if "query.wikidata.org" in url:
        powershell = shutil.which("powershell.exe") or shutil.which("pwsh")
        if powershell:
            ps_script = (
                "$ProgressPreference='SilentlyContinue'; "
                "$u=$env:CAS_QUERY_URL; "
                "$r=Invoke-WebRequest -UseBasicParsing -TimeoutSec 20 -Uri $u; "
                "[Text.Encoding]::UTF8.GetString([byte[]]$r.Content)"
            )
            try:
                ps_env = dict(__import__("os").environ)
                ps_env["CAS_QUERY_URL"] = url
                completed = subprocess.run(
                    [powershell, "-NoProfile", "-NonInteractive", "-Command", ps_script],
                    capture_output=True,
                    text=True,
                    env=ps_env,
                    timeout=timeout + 5,
                    check=True,
                )
                data = json.loads(completed.stdout)
                try:
                    cache_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
                except OSError:
                    pass
                return data
            except (OSError, subprocess.SubprocessError, json.JSONDecodeError, TypeError):
                pass
        curl = shutil.which("curl.exe") or shutil.which("curl")
        if curl:
            try:
                completed = subprocess.run(
                    [curl, "-fsSL", "--compressed", "--max-time", str(timeout), "-A", USER_AGENT, url],
                    capture_output=True,
                    text=True,
                    timeout=timeout + 3,
                    check=True,
                )
                data = json.loads(completed.stdout)
                try:
                    cache_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
                except OSError:
                    pass
                return data
            except (OSError, subprocess.SubprocessError, json.JSONDecodeError, TypeError):
                pass

    accept = "application/sparql-results+json" if "query.wikidata.org" in url else "application/json"
    request = Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Accept": accept,
            "Accept-Encoding": "identity",
        },
    )
    last_error: Exception | None = None
    for attempt in range(DEFAULT_RETRIES + 1):
        try:
            with urlopen(request, timeout=timeout) as response:
                data = json.loads(response.read().decode("utf-8"))
            break
        except HTTPError as exc:
            if exc.code == 404:
                raise QueryError("未找到匹配物质") from exc
            last_error = exc
        except (URLError, TimeoutError, json.JSONDecodeError) as exc:
            last_error = exc
        if attempt < DEFAULT_RETRIES:
            time.sleep(0.5 * (attempt + 1))
    else:
        if isinstance(last_error, HTTPError):
            raise QueryError(f"免费 API HTTP {last_error.code}") from last_error
        raise QueryError(f"无法访问免费 API：{last_error}") from last_error

    try:
        cache_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    except OSError:
        # Cache is an optimization; a read-only directory must not break lookup.
        pass
    return data


def get_cids(query: str, cache_dir: Path, timeout: int, use_cache: bool) -> list[int]:
    encoded = quote(query, safe="")
    if is_cas(query):
        # Do not fall back from an exact CAS lookup to a fuzzy name lookup:
        # that can turn an unknown polymer/CAS into a misleading monomer hit.
        urls = [f"{API_ROOT}/compound/identifier/{encoded}/cids/JSON?identifier_type=CAS"]
    else:
        urls = [f"{API_ROOT}/compound/name/{encoded}/cids/JSON"]

    for url in urls:
        try:
            data = get_json(url, cache_dir, timeout, use_cache)
        except QueryError:
            continue
        cids = data.get("IdentifierList", {}).get("CID", [])
        if cids:
            return [int(cid) for cid in cids[:10]]
    raise QueryError(f"未找到“{query}”对应的 PubChem 化合物")


def name_variants(name: str) -> list[str]:
    """Generate conservative naming variants without a substance mapping table."""
    value = name.strip()
    value = value.replace("－", "-").replace("–", "-").replace("—", "-").replace("﹣", "-")
    value = re.sub(r"\s+", "", value)
    variants = [value]
    # Chemical locants are often omitted in Chinese common names, e.g.
    # N-甲基-2-吡咯烷酮 vs N-甲基吡咯烷酮.
    without_locant = re.sub(r"(?<=[\u4e00-\u9fff])-?\d+-?(?=[\u4e00-\u9fff])", "", value)
    if without_locant and without_locant not in variants:
        variants.append(without_locant)
    return variants


def resolve_chinese_sparql(name: str, cache_dir: Path, timeout: int, use_cache: bool) -> dict[str, Any] | None:
    """Resolve an exact Chinese label/alias using Wikidata's free SPARQL endpoint."""
    for candidate in name_variants(name):
        escaped = json.dumps(candidate, ensure_ascii=False)
        sparql = f"""
SELECT ?item ?cas ?cid WHERE {{
  {{ ?item rdfs:label {escaped}@zh }}
  UNION
  {{ ?item skos:altLabel {escaped}@zh }}
  OPTIONAL {{ ?item wdt:P231 ?cas . }}
  OPTIONAL {{ ?item wdt:P662 ?cid . }}
}}
LIMIT 10
"""
        url = f"{WIKIDATA_SPARQL}?format=json&query={quote(sparql, safe='')}"
        try:
            data = get_json(url, cache_dir, timeout, use_cache)
        except QueryError:
            continue
        for binding in data.get("results", {}).get("bindings", []):
            cas = binding.get("cas", {}).get("value")
            cid = binding.get("cid", {}).get("value")
            if (isinstance(cas, str) and is_cas(cas)) or (isinstance(cid, str) and cid.isdigit()):
                return {
                    "cas": cas if isinstance(cas, str) and is_cas(cas) else None,
                    "pubchem_cid": int(cid) if isinstance(cid, str) and cid.isdigit() else None,
                    "wikidata_id": binding.get("item", {}).get("value", "").rsplit("/", 1)[-1],
                    "matched_name": candidate,
                    "source": "Wikidata SPARQL",
                }
    return None


def resolve_cas_chinese(cas: str, cache_dir: Path, timeout: int, use_cache: bool) -> list[str]:
    """Return free Chinese labels attached to an exact CAS in Wikidata."""
    escaped = json.dumps(cas, ensure_ascii=False)
    sparql = f"""
SELECT ?label WHERE {{
  ?item wdt:P231 {escaped} .
  ?item rdfs:label ?label .
  FILTER(LANG(?label) = "zh" || LANG(?label) = "zh-hans" || LANG(?label) = "zh-cn")
}}
LIMIT 10
"""
    url = f"{WIKIDATA_SPARQL}?format=json&query={quote(sparql, safe='')}"
    try:
        data = get_json(url, cache_dir, timeout, use_cache)
    except QueryError:
        return []
    values = []
    for binding in data.get("results", {}).get("bindings", []):
        value = binding.get("label", {}).get("value", "").strip()
        if value and value not in values:
            values.append(value)
    return values


def resolve_chinese_name(name: str, cache_dir: Path, timeout: int, use_cache: bool) -> dict[str, Any] | None:
    """Resolve a Chinese label to a CAS via Wikidata, then query PubChem.

    This is a best-effort name bridge. The final result still comes from
    PubChem and the returned CAS remains a candidate requiring verification.
    """
    params = (
        "action=wbsearchentities&search=" + quote(name, safe="")
        + "&language=zh&uselang=zh&format=json&limit=5"
    )
    search_url = f"{WIKIDATA_API}?{params}"
    try:
        search = get_json(search_url, cache_dir, timeout, use_cache)
    except QueryError:
        return None
    for item in search.get("search", []):
        entity_id = item.get("id", "")
        if not entity_id:
            continue
        entity_url = f"{WIKIDATA_API}?action=wbgetentities&ids={quote(entity_id, safe='')}&format=json"
        try:
            entity = get_json(entity_url, cache_dir, timeout, use_cache)
        except QueryError:
            continue
        claims = entity.get("entities", {}).get(entity_id, {}).get("claims", {})
        cas_value = None
        for claim in claims.get("P231", []):  # CAS Registry Number
            value = claim.get("mainsnak", {}).get("datavalue", {}).get("value", "")
            if isinstance(value, str) and is_cas(value):
                cas_value = value
                break
        cid_value = None
        for claim in claims.get("P662", []):  # PubChem CID
            value = claim.get("mainsnak", {}).get("datavalue", {}).get("value", "")
            if str(value).isdigit():
                cid_value = int(value)
                break
        if cas_value or cid_value:
            return {"cas": cas_value, "pubchem_cid": cid_value, "wikidata_id": entity_id}
    return None


def properties_for(cid: int, cache_dir: Path, timeout: int, use_cache: bool) -> dict[str, Any]:
    fields = "Title,MolecularFormula,MolecularWeight,CanonicalSMILES,IsomericSMILES,InChI,InChIKey"
    url = f"{API_ROOT}/compound/cid/{cid}/property/{fields}/JSON"
    data = get_json(url, cache_dir, timeout, use_cache)
    properties = data.get("PropertyTable", {}).get("Properties", [])
    return properties[0] if properties else {"CID": cid}


def synonyms_for(cid: int, cache_dir: Path, timeout: int, use_cache: bool) -> list[str]:
    url = f"{API_ROOT}/compound/cid/{cid}/synonyms/JSON"
    try:
        data = get_json(url, cache_dir, timeout, use_cache)
    except QueryError:
        return []
    info = data.get("InformationList", {}).get("Information", [])
    if not info:
        return []
    return [str(x) for x in info[0].get("Synonym", [])]


def find_cas(synonyms: list[str]) -> list[str]:
    values = []
    for synonym in synonyms:
        value = synonym.strip()
        if CAS_ANY_RE.fullmatch(value) and value not in values:
            values.append(value)
    return values


def cas_only_fallback(cas: str, query: str) -> dict[str, Any]:
    """Return a transparent CAS-only record when no structure database matches.

    This is deliberately not a name mapping table.  A CAS can identify a
    polymer, UVCB, mixture, or trade-related substance for which a normalized
    PubChem structure does not exist.  Returning the exact input with public
    registry links is safer than silently substituting a monomer or product.
    """
    encoded = quote(cas, safe="")
    links = [
        f"https://pubchem.ncbi.nlm.nih.gov/#query={encoded}",
        f"{EPA_SRS_SEARCH}?search=basic&query={encoded}",
        f"{NCI_RESOLVER}/{encoded}/names",
    ]
    title = query if not is_cas(query) else ""
    return {
        "cid": None,
        "title": title,
        "cas_candidates": [cas],
        "chinese_names": [query] if not is_cas(query) else [],
        "synonyms": [],
        "molecular_formula": "",
        "molecular_weight": "",
        "canonical_smiles": "",
        "isomeric_smiles": "",
        "inchi": "",
        "inchikey": "",
        "record_type": "CAS-only / polymer-UVCB candidate",
        "structure_available": False,
        "source_urls": links,
        "source_url": links[0],
        "warning": "已识别输入 CAS 格式，但免费结构化数据库未返回化合物结构；请以供应商 SDS、规格书和法规清单核实中文名称、组成及法规属性。",
    }


def lookup(query: str, cache_dir: Path, timeout: int, use_cache: bool) -> dict[str, Any]:
    query = query.strip()
    if not query:
        raise QueryError("查询内容不能为空")

    resolved_cas = None
    resolution = None
    try:
        cids = get_cids(query, cache_dir, timeout, use_cache)
    except QueryError:
        if is_cas(query):
            return {
                "query": query,
                "query_type": "cas",
                "lookup_status": "identified_without_structure",
                "name_resolution": None,
                "source": "PubChem + public registry links",
                "source_note": "该 CAS 未返回 PubChem 化合物结构；它可能属于聚合物、UVCB、混合物或尚未建立公开结构记录的物质。",
                "queried_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "results": [cas_only_fallback(query, query)],
            }
        resolution = resolve_chinese_sparql(query, cache_dir, timeout, use_cache)
        if not resolution:
            resolution = resolve_chinese_name(query, cache_dir, timeout, use_cache)
        if not resolution:
            raise
        resolved_cas = resolution.get("cas")
        if resolved_cas:
            cids = get_cids(resolved_cas, cache_dir, timeout, use_cache)
        elif resolution.get("pubchem_cid"):
            cids = [int(resolution["pubchem_cid"])]
        else:
            raise QueryError(f"Wikidata 找到“{query}”，但没有可用的 PubChem/CAS 标识")
    records = []
    for cid in cids:
        props = properties_for(cid, cache_dir, timeout, use_cache)
        synonyms = synonyms_for(cid, cache_dir, timeout, use_cache)
        records.append(
            {
                "cid": cid,
                "title": props.get("Title", ""),
                "cas_candidates": find_cas(synonyms),
                "chinese_names": [],
                "synonyms": synonyms[:30],
                "molecular_formula": props.get("MolecularFormula", ""),
                "molecular_weight": props.get("MolecularWeight", ""),
                "canonical_smiles": props.get("ConnectivitySMILES", props.get("CanonicalSMILES", "")),
                "isomeric_smiles": props.get("SMILES", props.get("IsomericSMILES", "")),
                "inchi": props.get("InChI", ""),
                "inchikey": props.get("InChIKey", ""),
                "source_url": f"https://pubchem.ncbi.nlm.nih.gov/compound/{cid}",
            }
        )
        for cas in records[-1]["cas_candidates"][:3]:
            records[-1]["chinese_names"].extend(
                name for name in resolve_cas_chinese(cas, cache_dir, timeout, use_cache)
                if name not in records[-1]["chinese_names"]
            )
    if is_cas(query) and not any(query in item["cas_candidates"] for item in records):
        raise QueryError(f"PubChem 返回结果未包含输入 CAS：{query}，已拒绝输出可能的错误匹配")
    return {
        "query": query,
        "query_type": "cas" if is_cas(query) else "name",
        "lookup_status": "complete",
        "name_resolution": resolution,
        "source": "PubChem PUG REST",
        "source_note": "CAS 候选来自 PubChem 的同义词/第三方标识，需用 SDS、供应商资料或 CAS 官方服务复核。",
        "queried_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "results": records,
    }


def print_text(result: dict[str, Any]) -> None:
    print(f"查询：{result['query']}  ({result['query_type']})")
    print(f"来源：{result['source']}")
    print(f"状态：{result.get('lookup_status', 'complete')}")
    print(f"提示：{result['source_note']}")
    print()
    for index, item in enumerate(result["results"], 1):
        print(f"[{index}] {item['title'] or '(无标题)'}")
        print(f"  中文名称：{', '.join(item['chinese_names']) or '(未找到)'}")
        print(f"  CAS候选：{', '.join(item['cas_candidates']) or '(未返回)'}")
        print(f"  PubChem CID：{item['cid']}")
        if item.get("record_type"):
            print(f"  记录类型：{item['record_type']}")
        if "structure_available" in item:
            print(f"  结构可用：{'是' if item['structure_available'] else '否'}")
        print(f"  分子式：{item['molecular_formula'] or '(未返回)'}")
        print(f"  分子量：{item['molecular_weight'] or '(未返回)'}")
        print(f"  Canonical SMILES：{item['canonical_smiles'] or '(未返回)'}")
        print(f"  Isomeric SMILES：{item['isomeric_smiles'] or '(未返回)'}")
        print(f"  InChIKey：{item['inchikey'] or '(未返回)'}")
        print(f"  同义名：{'; '.join(item['synonyms'][:10]) or '(未返回)'}")
        print(f"  链接：{item['source_url']}")
        if item.get("source_urls"):
            print(f"  公共查询入口：{' | '.join(item['source_urls'])}")
        if item.get("warning"):
            print(f"  警告：{item['warning']}")
        if index != len(result["results"]):
            print()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="通过 PubChem 免费 API 查询 CAS/化学物质信息")
    parser.add_argument("query", help="中文名称、英文名称或 CAS 号，例如：丙烯酸、9003-01-4")
    parser.add_argument("--json", action="store_true", dest="as_json", help="以 JSON 输出")
    parser.add_argument("--no-cache", action="store_true", help="不读取或写入本地缓存")
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help=f"网络超时秒数，默认 {DEFAULT_TIMEOUT}")
    parser.add_argument("--cache-dir", type=Path, default=DEFAULT_CACHE, help="缓存目录")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        result = lookup(args.query, args.cache_dir, max(1, args.timeout), not args.no_cache)
    except QueryError as exc:
        if args.as_json:
            print(json.dumps({"query": args.query, "error": str(exc)}, ensure_ascii=False))
        else:
            print(f"查询失败：{exc}", file=sys.stderr)
        return 2

    if args.as_json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print_text(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
