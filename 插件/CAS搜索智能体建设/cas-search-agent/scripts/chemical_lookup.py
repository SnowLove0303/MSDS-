#!/usr/bin/env python3
"""Evidence-oriented base lookup for Chinese chemical names.

The script supplies candidate identifiers and official-source entry points.
It deliberately does not label PubChem CAS values as authoritative.
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import time
import unicodedata
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote
from urllib.request import Request, urlopen

CAS_RE = re.compile(r"^\d{2,7}-\d{2}-\d$")
EC_RE = re.compile(r"^\d{3}-\d{3}-\d$")
PUBCHEM = "https://pubchem.ncbi.nlm.nih.gov/rest/pug"
WIKIDATA = "https://www.wikidata.org/w/api.php"
CHEBI_OLS = "https://www.ebi.ac.uk/ols4/api/search"
USER_AGENT = "cas-search-agent/0.1 (official-source-first chemical identity lookup)"

SOURCE_REGISTRY = [
    {"name": "ECHA", "tier": "A", "domain": "echa.europa.eu", "url": "https://echa.europa.eu/substance-information"},
    {"name": "ECHA CHEM", "tier": "A", "domain": "chem.echa.europa.eu", "url": "https://chem.echa.europa.eu/"},
    {"name": "EPA CompTox", "tier": "A", "domain": "comptox.epa.gov", "url": "https://comptox.epa.gov/dashboard/"},
    {"name": "OECD eChemPortal", "tier": "A", "domain": "echemportal.org", "url": "https://www.echemportal.org/echemportal/substance-search"},
    {"name": "EU EUR-Lex", "tier": "A", "domain": "eur-lex.europa.eu", "url": "https://eur-lex.europa.eu/"},
    {"name": "NITE-CHRIP", "tier": "A", "domain": "nite.go.jp", "url": "https://www.nite.go.jp/en/chem/chrip/chrip_search/systemTop"},
    {"name": "GESTIS", "tier": "A", "domain": "gestis.dguv.de", "url": "https://gestis.dguv.de/"},
    {"name": "ChEBI/EMBL-EBI", "tier": "B", "domain": "ebi.ac.uk", "url": "https://www.ebi.ac.uk/chebi/"},
    {"name": "NLM PubChem", "tier": "B", "domain": "pubchem.ncbi.nlm.nih.gov", "url": "https://pubchem.ncbi.nlm.nih.gov/"},
    {"name": "CAS Common Chemistry", "tier": "B", "domain": "commonchemistry.cas.org", "url": "https://commonchemistry.cas.org/"},
    {"name": "ChemSpider", "tier": "B", "domain": "chemspider.com", "url": "https://www.chemspider.com/"},
]

COMMON_TRADITIONAL = str.maketrans({
    "亞": "亚", "劑": "剂",
    "鈉": "钠", "鉀": "钾", "銨": "铵", "鋁": "铝", "矽": "硅",
    "鈣": "钙", "鄰": "邻", "對": "对",
})


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    value = value.translate(COMMON_TRADITIONAL)
    value = value.replace("（", "(").replace("）", ")").replace("【", "[").replace("】", "]")
    value = re.sub(r"[\u00a0\s]+", " ", value).strip()
    return value


def normalized_identifier(value: str) -> str:
    value = normalize_text(value)
    return re.sub(r"\s+", "", value)


def is_cas(value: str) -> bool:
    return bool(CAS_RE.fullmatch(normalized_identifier(value)))


def is_ec(value: str) -> bool:
    return bool(EC_RE.fullmatch(normalized_identifier(value)))


def query_variants(query: str) -> list[str]:
    """Create conservative spelling variants to improve recall without fuzzy guessing."""
    base = normalize_text(query)
    variants = [base]
    variants.extend([
        base.replace("(", "").replace(")", ""),
        base.replace("[", "").replace("]", ""),
        base.replace("-", " "),
        base.replace("／", "/"),
        base.replace("、", ","),
    ])
    if base.lower().startswith("cas "):
        variants.append(base[4:])
    if re.search(r"[\u4e00-\u9fff]", base):
        variants.extend([base.replace("酸", " acid"), base.replace("钠", " sodium"), base.replace("钾", " potassium")])
    return list(dict.fromkeys(x.strip() for x in variants if x.strip()))


def get_json(url: str, timeout: int) -> dict:
    request = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    try:
        with urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except (URLError, TimeoutError) as first_error:
        if "wikidata.org" not in url:
            raise
        curl = shutil.which("curl.exe") or shutil.which("curl")
        if curl:
            try:
                completed = subprocess.run(
                    [curl, "-fsSL", "--compressed", "--max-time", str(timeout), "-A", USER_AGENT, url],
                    capture_output=True, text=True, timeout=timeout + 3, check=True,
                )
                return json.loads(completed.stdout)
            except (OSError, subprocess.SubprocessError, json.JSONDecodeError):
                pass
        raise first_error


def pubchem_cids_by_identifier(identifier: str, identifier_types: list[str], timeout: int) -> tuple[list[int], list[str]]:
    """Try explicit PubChem identifier types and retain the successful paths."""
    encoded = quote(normalized_identifier(identifier), safe="")
    cids: list[int] = []
    paths: list[str] = []
    for identifier_type in identifier_types:
        url = f"{PUBCHEM}/compound/identifier/{encoded}/cids/JSON?identifier_type={quote(identifier_type)}"
        try:
            data = get_json(url, timeout)
        except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError):
            continue
        values = [int(x) for x in data.get("IdentifierList", {}).get("CID", [])[:25]]
        if values:
            cids.extend(values)
            paths.append(f"PubChem identifier_type={identifier_type}")
    return list(dict.fromkeys(cids)), paths


def pubchem_cids_by_name(query: str, timeout: int) -> tuple[list[int], list[str]]:
    encoded = quote(normalize_text(query), safe="")
    url = f"{PUBCHEM}/compound/name/{encoded}/cids/JSON"
    data = get_json(url, timeout)
    return [int(x) for x in data.get("IdentifierList", {}).get("CID", [])[:25]], ["PubChem compound/name"]


def pubchem_sids_by_name(query: str, timeout: int) -> tuple[list[int], list[str]]:
    """Search deposited substances as a fallback; SIDs are later mapped to CIDs."""
    encoded = quote(normalize_text(query), safe="")
    url = f"{PUBCHEM}/substance/name/{encoded}/sids/JSON"
    data = get_json(url, timeout)
    sids = [int(x) for x in data.get("IdentifierList", {}).get("SID", [])[:25]]
    return sids, (["PubChem substance/name"] if sids else [])


def pubchem_cids_from_sids(sids: list[int], timeout: int) -> list[int]:
    if not sids:
        return []
    sid_text = ",".join(str(x) for x in sids)
    url = f"{PUBCHEM}/substance/sid/{sid_text}/cids/JSON?cids_type=all"
    try:
        data = get_json(url, timeout)
    except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError):
        return []
    return [int(x) for x in data.get("IdentifierList", {}).get("CID", [])[:25]]


def wikidata_bridge(query: str, timeout: int) -> tuple[list[str], list[int], list[str]]:
    """Resolve multilingual labels/aliases to CAS and PubChem CID claims."""
    cas_values: list[str] = []
    cid_values: list[int] = []
    paths: list[str] = []
    entity_ids: list[str] = []
    for variant in query_variants(query):
        for language in ("zh", "en"):
            params = (
                "action=wbsearchentities&search=" + quote(variant, safe="")
                + f"&language={language}&uselang={language}&format=json&limit=10"
            )
            search = get_json(f"{WIKIDATA}?{params}", timeout)
            entity_ids.extend(item.get("id", "") for item in search.get("search", []))
            if search.get("search"):
                paths.append(f"Wikidata {language} label/alias: {variant}")
    for entity_id in list(dict.fromkeys(x for x in entity_ids if x))[:30]:
        entity_url = f"{WIKIDATA}?action=wbgetentities&ids={quote(entity_id, safe='')}&format=json"
        entity = get_json(entity_url, timeout)
        claims = entity.get("entities", {}).get(entity_id, {}).get("claims", {})
        for claim in claims.get("P231", []):
            value = claim.get("mainsnak", {}).get("datavalue", {}).get("value", "")
            if isinstance(value, str) and is_cas(value):
                cas_values.append(normalized_identifier(value))
        for claim in claims.get("P662", []):
            value = claim.get("mainsnak", {}).get("datavalue", {}).get("value", "")
            if str(value).isdigit():
                cid_values.append(int(value))
    return list(dict.fromkeys(cas_values)), list(dict.fromkeys(cid_values)), list(dict.fromkeys(paths))


def chebi_bridge(query: str, timeout: int) -> tuple[list[str], list[str]]:
    """Use EMBL-EBI OLS/ChEBI search as an additional name bridge."""
    values: list[str] = []
    paths: list[str] = []
    for variant in query_variants(query)[:6]:
        url = f"{CHEBI_OLS}?q={quote(variant, safe='')}&ontology=chebi&rows=10"
        try:
            data = get_json(url, timeout)
        except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError):
            continue
        for doc in data.get("response", {}).get("docs", [])[:10]:
            label = doc.get("label") or doc.get("suggestible") or doc.get("obo_id")
            if label:
                values.append(str(label))
        if data.get("response", {}).get("docs"):
            paths.append(f"ChEBI/OLS search: {variant}")
    return list(dict.fromkeys(values)), list(dict.fromkeys(paths))


def pubchem_record(cid: int, timeout: int) -> dict:
    fields = "Title,MolecularFormula,MolecularWeight,InChIKey"
    props = get_json(f"{PUBCHEM}/compound/cid/{cid}/property/{fields}/JSON", timeout)
    syns = get_json(f"{PUBCHEM}/compound/cid/{cid}/synonyms/JSON", timeout)
    prop = (props.get("PropertyTable", {}).get("Properties") or [{}])[0]
    synonyms = ((syns.get("InformationList", {}).get("Information") or [{}])[0].get("Synonym") or [])
    cas = []
    ec = []
    for synonym in synonyms:
        if not isinstance(synonym, str):
            continue
        cas.extend(re.findall(r"(?<!\d)\d{2,7}-\d{2}-\d(?!\d)", synonym))
        ec.extend(re.findall(r"(?<!\d)\d{3}-\d{3}-\d(?!\d)", synonym))
    return {
        "cid": cid,
        "title": prop.get("Title", ""),
        "molecular_formula": prop.get("MolecularFormula", ""),
        "molecular_weight": prop.get("MolecularWeight", ""),
        "inchikey": prop.get("InChIKey", ""),
        "cas_candidates": list(dict.fromkeys(cas)),
        "ec_candidates": list(dict.fromkeys(ec)),
        "source": "PubChem",
        "source_url": f"https://pubchem.ncbi.nlm.nih.gov/compound/{cid}",
    }


def official_entry_points(query: str, candidates: list[str], mode: str) -> list[dict]:
    terms = " ".join([query, *candidates]).strip()
    encoded = quote(terms, safe="")
    entries = [
        {"authority": "ECHA", "tier": "A", "url": f"https://echa.europa.eu/search-for-chemicals?search_api_fulltext={encoded}", "status": "search-entry"},
        {"authority": "ECHA CHEM", "tier": "A", "url": f"https://chem.echa.europa.eu/search-for-chemicals?search_api_fulltext={encoded}", "status": "search-entry"},
        {"authority": "EPA CompTox", "tier": "A", "url": f"https://comptox.epa.gov/dashboard/search?query={encoded}", "status": "search-entry"},
        {"authority": "OECD eChemPortal", "tier": "B", "url": f"https://www.echemportal.org/echemportal/substance-search?search={encoded}", "status": "search-entry"},
        {"authority": "ChEBI/EMBL-EBI", "tier": "B", "url": f"https://www.ebi.ac.uk/chebi/searchId.do?chebiId={encoded}", "status": "search-entry"},
        {"authority": "NLM PubChem", "tier": "B", "url": f"https://pubchem.ncbi.nlm.nih.gov/#query={encoded}", "status": "search-entry"},
    ]
    if mode == "deep":
        entries.extend([
            {"authority": "EU EUR-Lex", "tier": "A", "url": f"https://eur-lex.europa.eu/search.html?text={encoded}", "status": "search-entry"},
            {"authority": "NITE-CHRIP", "tier": "A", "url": "https://www.nite.go.jp/en/chem/chrip/chrip_search/systemTop", "status": "manual-search-entry"},
            {"authority": "GESTIS", "tier": "A", "url": "https://gestis.dguv.de/", "status": "manual-search-entry"},
            {"authority": "CAS Common Chemistry", "tier": "B", "url": f"https://commonchemistry.cas.org/results?query={encoded}", "status": "search-entry"},
            {"authority": "ChemSpider", "tier": "B", "url": f"https://www.chemspider.com/Search.aspx?q={encoded}", "status": "search-entry"},
        ])
    return entries


def build_search_plan(query: str, terms: list[str], mode: str) -> list[dict]:
    """Build URLs for the Agent's parallel web-search round."""
    selected = SOURCE_REGISTRY if mode == "deep" else [x for x in SOURCE_REGISTRY if x["name"] in {"ECHA", "ECHA CHEM", "EPA CompTox", "OECD eChemPortal", "ChEBI/EMBL-EBI", "NLM PubChem"}]
    plan = []
    for source in selected:
        source_queries = []
        for term in terms[:8]:
            source_queries.append(f'site:{source["domain"]} "{term}" CAS EC')
        plan.append({"source": source["name"], "tier": source["tier"], "domain": source["domain"], "queries": source_queries, "entry_url": source["url"], "status": "ready-for-agent-web-search"})
    return plan


def lookup(query: str, timeout: int, mode: str) -> dict:
    query = normalize_text(query)
    if not query:
        raise ValueError("查询内容不能为空")
    candidates_by_cid: dict[int, dict] = {}
    errors: list[str] = []
    paths: list[str] = []
    cas_bridge: list[str] = []
    chebi_names: list[str] = []

    def add_cids(cids: list[int], path_names: list[str]) -> None:
        paths.extend(path_names)
        for cid in cids:
            if cid in candidates_by_cid:
                continue
            try:
                candidates_by_cid[cid] = pubchem_record(cid, timeout)
            except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError) as exc:
                errors.append(f"PubChem record CID {cid}: {exc}")

    if is_cas(query):
        cids, identifier_paths = pubchem_cids_by_identifier(query, ["CAS", "CAS Registry Number"], timeout)
        add_cids(cids, identifier_paths)
    elif is_ec(query):
        cids, identifier_paths = pubchem_cids_by_identifier(query, ["EINECS", "EC Number", "EINECS/ELINCS"], timeout)
        add_cids(cids, identifier_paths)
    else:
        for variant in query_variants(query):
            try:
                cids, name_paths = pubchem_cids_by_name(variant, timeout)
                add_cids(cids, [f"{path}: {variant}" for path in name_paths])
            except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError):
                continue
        direct_has_cas = any(item.get("cas_candidates") for item in candidates_by_cid.values())
        if not candidates_by_cid or not direct_has_cas:
            try:
                sids, sid_paths = pubchem_sids_by_name(query, timeout)
                add_cids(pubchem_cids_from_sids(sids, timeout), sid_paths)
            except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError):
                pass

            try:
                cas_bridge, wikidata_cids, wikidata_paths = wikidata_bridge(query, timeout)
                paths.extend(wikidata_paths)
                add_cids(wikidata_cids, ["Wikidata PubChem CID claim"] if wikidata_cids else [])
                for cas in cas_bridge:
                    cids, identifier_paths = pubchem_cids_by_identifier(cas, ["CAS", "CAS Registry Number"], timeout)
                    add_cids(cids, [f"Wikidata CAS bridge {cas}", *identifier_paths])
            except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError) as exc:
                errors.append(f"Wikidata multilingual bridge: {exc}")

            try:
                chebi_names, chebi_paths = chebi_bridge(query, timeout)
                paths.extend(chebi_paths)
                for name in chebi_names[:20]:
                    try:
                        cids, name_paths = pubchem_cids_by_name(name, timeout)
                        add_cids(cids, [f"ChEBI name bridge: {name}", *name_paths])
                    except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError):
                        continue
            except (HTTPError, URLError, TimeoutError, ValueError, json.JSONDecodeError) as exc:
                errors.append(f"ChEBI bridge: {exc}")

    candidates = list(candidates_by_cid.values())
    cas_values = [cas for item in candidates for cas in item.get("cas_candidates", [])]
    ec_values = [ec for item in candidates for ec in item.get("ec_candidates", [])]
    all_terms = list(dict.fromkeys([query, *query_variants(query), *cas_bridge, *chebi_names[:10], *cas_values, *ec_values]))
    return {
        "query": query,
        "query_type": "cas" if is_cas(query) else ("ec" if is_ec(query) else "name"),
        "mode": mode,
        "queried_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "candidates": candidates,
        "query_paths": list(dict.fromkeys(paths)),
        "search_terms_used": all_terms,
        "official_source_entry_points": official_entry_points(query, [*cas_values, *ec_values], mode),
        "source_registry": SOURCE_REGISTRY,
        "parallel_search_plan": build_search_plan(query, all_terms, mode),
        "evidence_policy": "PubChem, Wikidata and ChEBI values are candidate evidence; official page verification is required.",
        "errors": errors,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="官方来源优先的中文化学物质基础检索")
    parser.add_argument("query", help="中文名、英文名或 CAS")
    parser.add_argument("--timeout", type=int, default=15)
    parser.add_argument("--mode", choices=("fast", "deep"), default="fast", help="fast=核心来源并行；deep=增加法规、职业健康和专业数据库")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    try:
        result = lookup(args.query, max(1, args.timeout), args.mode)
    except (ValueError, OSError) as exc:
        print(f"查询失败：{exc}", file=sys.stderr)
        return 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"查询：{result['query']}")
        print(f"模式：{result['mode'] if 'mode' in result else 'fast'}")
        print(f"查询路径：{len(result['query_paths'])} 条")
        for item in result["candidates"]:
            print(f"- {item['title']} | CAS候选: {', '.join(item['cas_candidates']) or '无'} | EC候选: {', '.join(item['ec_candidates']) or '无'} | {item['source_url']}")
        if result["query_paths"]:
            print("路径明细：")
            for path in result["query_paths"]:
                print(f"- {path}")
        print("官方来源入口：")
        for source in result["official_source_entry_points"]:
            print(f"- [{source['authority']}] {source['url']}")
        if result["errors"]:
            print("错误：" + "; ".join(result["errors"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
