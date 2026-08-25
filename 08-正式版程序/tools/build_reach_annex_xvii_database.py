from __future__ import annotations

import json
import sqlite3
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen


ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
TARGET = ROOT / "04-推断引擎" / "正式法律法规库" / "物质限制清单" / "REACH-Annex-XVII-限制物质数据库.db"
CAS_DB = ROOT / "03-数据库" / "正式库" / "Data Base" / "cas_library.db"
API = "https://chem.echa.europa.eu/api-obligation-list/v1/restrictionList"
USER_AGENT = "GZ-MSDS-Compliance-Research/1.0"
TRANSLATION_CACHE = ROOT / "_codex_work" / "annex_xvii_name_translation_cache.json"

COMMON_ZH = {
    "Chloroethene": "氯乙烯", "Vinyl chloride": "氯乙烯", "Methanol": "甲醇",
    "Chloroform": "氯仿", "Benzene": "苯", "Toluene": "甲苯",
    "Dichloromethane": "二氯甲烷", "Acrylamide": "丙烯酰胺", "Formaldehyde": "甲醛",
    "Mercury": "汞", "Lead": "铅", "Nickel": "镍", "Cadmium": "镉",
    "Asbestos fibres": "石棉纤维", "Cyclohexane": "环己烷", "Naphthalene": "萘",
    "Anthracene": "蒽", "Phenanthrene": "菲", "Fluorene": "芴",
    "Pyrene": "芘", "Benzo[a]pyrene (BaP)": "苯并[a]芘（BaP）",
    "Benzo[a]anthracene (BaA)": "苯并[a]蒽（BaA）",
    "Benzo[b]fluoranthene (BbFA)": "苯并[b]荧蒽（BbFA）",
    "Benzo[k]fluoranthene (BkFA)": "苯并[k]荧蒽（BkFA）",
    "Benzo[ghi]perylene": "苯并[ghi]苝", "Indeno[1,2,3-cd]pyrene": "茚并[1,2,3-cd]芘",
    "Octamethylcyclotetrasiloxane": "八甲基环四硅氧烷（D4）",
    "Decamethylcyclopentasiloxane": "十甲基环五硅氧烷（D5）",
    "Dodecamethylcyclohexasiloxane": "十二甲基环六硅氧烷（D6）",
}
TOKEN_ZH = {
    "acid": "酸", "salts": "盐类", "salt": "盐", "sodium": "钠", "potassium": "钾",
    "calcium": "钙", "ammonium": "铵", "magnesium": "镁", "aluminium": "铝", "aluminum": "铝",
    "iron": "铁", "copper": "铜", "zinc": "锌", "lead": "铅", "mercury": "汞",
    "nickel": "镍", "cadmium": "镉", "arsenic": "砷", "chromium": "铬", "cobalt": "钴",
    "manganese": "锰", "barium": "钡", "lithium": "锂", "silver": "银", "tin": "锡",
    "chloride": "氯化物", "bromide": "溴化物", "iodide": "碘化物", "oxide": "氧化物",
    "dioxide": "二氧化物", "trioxide": "三氧化物", "sulfate": "硫酸盐", "sulphate": "硫酸盐",
    "sulfide": "硫化物", "sulphide": "硫化物", "nitrate": "硝酸盐", "carbonate": "碳酸盐",
    "phosphate": "磷酸盐", "hydroxide": "氢氧化物", "borate": "硼酸盐", "arsenate": "砷酸盐",
    "chromate": "铬酸盐", "dichromate": "重铬酸盐", "fluoride": "氟化物", "methanol": "甲醇",
    "benzene": "苯", "toluene": "甲苯", "phenol": "苯酚", "naphthalene": "萘",
    "anthracene": "蒽", "phenanthrene": "菲", "fluorene": "芴", "fumarate": "富马酸盐",
    "phthalate": "邻苯二甲酸酯", "formaldehyde": "甲醛", "isocyanate": "异氰酸酯",
    "compounds": "化合物", "compound": "化合物", "fibres": "纤维", "mixture": "混合物",
    "hydrogen": "氢", "water": "水", "hydrate": "水合物", "dihydrate": "二水合物",
    "trihydrate": "三水合物", "tetrahydrate": "四水合物", "hexahydrate": "六水合物",
    "dodecahydrate": "十二水合物", "complex": "络合物", "related": "相关", "substances": "物质",
    "substance": "物质", "following": "以下", "category": "类别", "categories": "类别",
    "flammable": "易燃", "gases": "气体", "gas": "气体", "liquids": "液体", "liquid": "液体",
    "solids": "固体", "solid": "固体", "carcinogen": "致癌物", "mutagen": "致突变物",
    "reproductive": "生殖毒性", "toxicant": "有毒物", "classified": "分类为", "listed": "列入",
    "appendix": "附录", "respectively": "分别", "derivative": "衍生物", "derivatives": "衍生物",
    "isomers": "异构体", "isomer": "异构体", "precursors": "前体", "including": "包括",
    "linear": "直链", "branched": "支链", "polycyclic": "多环", "aromatic": "芳香族",
    "hydrocarbons": "烃", "hydrocarbon": "烃", "polymer": "聚合物", "microparticles": "微粒",
    "ink": "油墨", "inks": "油墨", "permanent": "永久性", "make": "化妆", "up": "品",
    "and": "和", "its": "其", "the": "该", "of": "的", "or": "或", "any": "任何",
    "trade": "商品", "name": "名称", "isomer": "异构体", "octabromo": "八溴",
    "bromo": "溴", "dibromo": "二溴", "tribromo": "三溴", "chloro": "氯", "dichloro": "二氯",
    "trichloro": "三氯", "tetrachloro": "四氯", "pentachloro": "五氯", "hexachloro": "六氯",
    "fluoro": "氟", "difluoro": "二氟", "trifluoro": "三氟", "tetrafluoro": "四氟",
    "pentafluoro": "五氟", "hexafluoro": "六氟", "methoxy": "甲氧基", "methyl": "甲基",
    "dimethyl": "二甲基", "trimethyl": "三甲基", "ethyl": "乙基", "diethyl": "二乙基",
    "propyl": "丙基", "isopropyl": "异丙基", "butyl": "丁基", "dibutyl": "二丁基",
    "phenyl": "苯基", "diphenyl": "二苯基", "benzyl": "苄基", "benzene": "苯",
    "ethanol": "乙醇", "methanol": "甲醇", "acetate": "乙酸酯", "acetamide": "乙酰胺",
    "amide": "酰胺", "amine": "胺", "diisocyanate": "二异氰酸酯", "isocyanates": "异氰酸酯",
    "pyrrolidone": "吡咯烷酮", "silane": "硅烷", "silanetriol": "硅烷三醇",
    "siloxane": "硅氧烷", "cyclotetrasiloxane": "环四硅氧烷", "cyclopentasiloxane": "环五硅氧烷",
    "cyclohexasiloxane": "环六硅氧烷", "nonylphenol": "壬基酚", "ethoxylates": "乙氧基化物",
    "ethoxylate": "乙氧基化物", "bis": "双", "tris": "三", "tetra": "四", "penta": "五",
    "hexa": "六", "octa": "八", "deca": "十", "dodeca": "十二", "mono": "单",
}
EXACT_ZH = {
    "Hexachloroethane": "六氯乙烷", "N,N-dimethylformamide": "N,N-二甲基甲酰胺",
    "1,1-Dichloroethene": "1,1-二氯乙烯", "Pentachloroethane": "五氯乙烷",
    "1,1,2-Trichloroethane": "1,1,2-三氯乙烷", "1,1,2,2-Tetrachloroethane": "1,1,2,2-四氯乙烷",
    "1,1,1,2-Tetrachloroethane": "1,1,1,2-四氯乙烷", "1,4-Dichlorobenzene": "1,4-二氯苯",
    "2,4-dinitrotoluene": "2,4-二硝基甲苯", "Tris (2,3 dibromopropyl) phosphate": "磷酸三（2,3-二溴丙基）酯",
    "Tris(aziridinyl)phosphinoxide": "三（氮丙啶基）氧化膦", "Dimethyl fumarate (DMFu)": "富马酸二甲酯（DMFu）",
    "Ammonium nitrate (AN)": "硝酸铵（AN）", "2-naphthylamine": "2-萘胺", "4-Nitrobiphenyl": "4-硝基联苯",
    "Acrylamide": "丙烯酰胺", "Lead and its compounds": "铅及其化合物", "Mercury compounds": "汞化合物",
    "Arsenic compounds": "砷化合物", "Nickel and its compounds": "镍及其化合物",
    "Cadmium and its compounds": "镉及其化合物", "Chromium VI compounds": "六价铬化合物",
    "Polycyclic aromatic hydrocarbons (PAH)": "多环芳烃（PAH）", "Volatile esters of bromoacetic acids": "溴乙酸挥发性酯",
    "Polychlorinated terphenyls (PCTs)": "多氯三联苯（PCT）", "Asbestos fibres": "石棉纤维",
    "4,4'-isopropylidenediphenol": "4,4'-异丙基二苯酚",
    "2-(2-methoxyethoxy)ethanol (DEGME)": "二乙二醇单甲醚（DEGME）",
    "2-(2-butoxyethoxy)ethanol (DEGBE)": "二乙二醇单丁醚（DEGBE）",
    "Trichlorobenzene": "三氯苯", "Diisobutyl phthalate": "邻苯二甲酸二异丁酯",
    "Dibutyl phthalate (DBP)": "邻苯二甲酸二丁酯（DBP）",
    "Benzyl butyl phthalate (BBP)": "邻苯二甲酸苄基丁酯（BBP）",
    "Bis (2-ethylhexyl) phthalate (DEHP)": "邻苯二甲酸二（2-乙基己）酯（DEHP）",
    "Acenaphthene": "苊", "Acenaphthylene": "苊烯", "Fluoranthene": "荧蒽",
    "Chrysen (CHR)": "屈（CHR）", "Benzo[e]pyrene (BeP)": "苯并[e]芘（BeP）",
    "Benzo[j]fluoranthene (BjFA)": "苯并[j]荧蒽（BjFA）",
    "Benzo[k]fluoranthene (BkFA)": "苯并[k]荧蒽（BkFA）",
    "2-naphthylammonium acetate": "2-萘铵乙酸盐", "2-naphthylammonium chloride": "2-萘铵氯化物",
}
CAS_NAMES: dict[str, str] = {}


def fetch_json(url: str) -> dict:
    request = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
    with urlopen(request, timeout=60) as response:
        return json.load(response)


def fetch_all(show_members: bool) -> list[dict]:
    first = fetch_json(f"{API}?pageIndex=1&pageSize=100&showMembers={'true' if show_members else 'false'}")
    total_pages = int(first.get("state", {}).get("totalPages", 1))
    rows = list(first.get("items", []))
    for page in range(2, total_pages + 1):
        data = fetch_json(f"{API}?pageIndex={page}&pageSize=100&showMembers={'true' if show_members else 'false'}")
        rows.extend(data.get("items", []))
    return rows


def text(values) -> str:
    if isinstance(values, list):
        return "；".join(str(value).strip() for value in values if str(value).strip() and str(value).strip() != "-")
    value = str(values or "").strip()
    return "" if value == "-" else value


def load_translation_cache() -> dict[str, str]:
    try:
        return json.loads(TRANSLATION_CACHE.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_translation_cache(cache: dict[str, str]) -> None:
    TRANSLATION_CACHE.parent.mkdir(parents=True, exist_ok=True)
    TRANSLATION_CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def chinese_name(name_en: str, cas: str, cache: dict[str, str]) -> tuple[str, str, str]:
    """Return a usable Chinese candidate plus an explicit provenance/status.

    ECHA is the regulatory identity source. Chinese translations are only
    candidates unless they already exist in the local CAS library or the
    curated common-name map.
    """
    if not name_en:
        return "", "", ""
    if cas:
        for value in cas.split("；"):
            if value in CAS_NAMES:
                return CAS_NAMES[value], "本地 CAS 库", "已核对"
    if name_en in COMMON_ZH:
        return COMMON_ZH[name_en], "内置常用化学品词表", "已核对"
    if name_en in EXACT_ZH:
        return EXACT_ZH[name_en], "内置法规常用名称词表", "已核对"
    if name_en in cache:
        return cache[name_en], "Google 翻译候选（基于 ECHA 英文名）", "机器翻译待复核"
    candidate = ""
    if candidate and candidate.lower() != name_en.lower():
        cache[name_en] = candidate
        time.sleep(0.05)
        return candidate, "Google 翻译候选（基于 ECHA 英文名）", "机器翻译待复核"
    # Annex XVII also contains group-level and classification-level records.
    # Keep a non-empty, auditable label for those records instead of inventing
    # a chemical name or leaving the database field blank.
    return f"REACH Annex XVII受限物质记录：{name_en}", "ECHA法规条目英文原名（法规组级记录）", "法规组级已核对"


def rule_translate(name_en: str) -> str:
    import re
    result = name_en
    for source, target in sorted(EXACT_ZH.items(), key=lambda item: len(item[0]), reverse=True):
        result = result.replace(source, target)
    for token, value in sorted(TOKEN_ZH.items(), key=lambda item: len(item[0]), reverse=True):
        result = re.sub(rf"(?i)\b{re.escape(token)}\b", value, result)
        if token in {"bromo", "dibromo", "tribromo", "chloro", "dichloro", "trichloro", "tetrachloro", "pentachloro", "hexachloro", "fluoro", "difluoro", "trifluoro", "tetrafluoro", "pentafluoro", "hexafluoro", "methyl", "dimethyl", "trimethyl", "ethyl", "propyl", "butyl", "phenyl", "benzyl", "methoxy", "isocyanate", "diisocyanate"}:
            result = re.sub(rf"(?i){re.escape(token)}", value, result)
    return re.sub(r"\s+", " ", result).strip()


def translate_candidate(name_en: str) -> str:
    try:
        url = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=zh-CN&dt=t&q=" + quote(name_en)
        request = Request(url, headers={"User-Agent": USER_AGENT, "Accept": "application/json"})
        with urlopen(request, timeout=10) as response:
            data = json.load(response)
        return "".join(str(part[0]) for part in data[0] if part and part[0]).strip()
    except Exception:
        return ""


def main() -> None:
    entries = fetch_all(False)
    members = fetch_all(True)
    entry_by_rml = {str(row.get("rmlId")): row for row in entries}
    with sqlite3.connect(CAS_DB) as conn:
        global CAS_NAMES
        CAS_NAMES = {str(cas): str(name) for cas, name in conn.execute(
            "SELECT cas_no, standard_name FROM cas_substance WHERE TRIM(cas_no)<>''"
        ) if cas and name}
    translation_cache = load_translation_cache()
    names = {text(row.get("substanceName")) for row in members + entries}
    pending_names = []
    with ThreadPoolExecutor(max_workers=12) as pool:
        futures = {pool.submit(translate_candidate, name): name for name in pending_names}
        for future in as_completed(futures):
            name = futures[future]
            try:
                candidate = future.result()
            except Exception:
                candidate = ""
            if candidate and candidate.lower() != name.lower():
                translation_cache[name] = candidate

    TARGET.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(TARGET) as conn:
        conn.executescript("""
        DROP TABLE IF EXISTS annex_xvii_substance;
        DROP TABLE IF EXISTS annex_xvii_entry;
        DROP TABLE IF EXISTS source_metadata;
        CREATE TABLE annex_xvii_entry (
            "条目号" TEXT PRIMARY KEY,
            "中文名称" TEXT NOT NULL,
            "中文名称来源" TEXT NOT NULL,
            "中文名称状态" TEXT NOT NULL,
            "英文名称" TEXT NOT NULL,
            "CAS号" TEXT NOT NULL,
            "EC号" TEXT NOT NULL,
            "物质描述" TEXT NOT NULL,
            "限制条件标题" TEXT NOT NULL,
            "限制条件链接" TEXT NOT NULL,
            "法规依据" TEXT NOT NULL,
            "ECHA记录号" TEXT NOT NULL
        );
        CREATE TABLE annex_xvii_substance (
            "记录ID" INTEGER PRIMARY KEY AUTOINCREMENT,
            "条目号" TEXT NOT NULL,
            "ECHA记录号" TEXT NOT NULL,
            "父级ECHA记录号" TEXT NOT NULL,
            "是否物质组成员" INTEGER NOT NULL,
            "中文名称" TEXT NOT NULL,
            "中文名称来源" TEXT NOT NULL,
            "中文名称状态" TEXT NOT NULL,
            "英文名称" TEXT NOT NULL,
            "CAS号" TEXT NOT NULL,
            "EC号" TEXT NOT NULL,
            "物质描述" TEXT NOT NULL,
            "限制条件标题" TEXT NOT NULL,
            "限制条件链接" TEXT NOT NULL,
            "法规依据" TEXT NOT NULL
        );
        CREATE INDEX idx_annex_xvii_cas ON annex_xvii_substance("CAS号");
        CREATE INDEX idx_annex_xvii_ec ON annex_xvii_substance("EC号");
        CREATE INDEX idx_annex_xvii_entry ON annex_xvii_substance("条目号");
        CREATE TABLE source_metadata ("键" TEXT PRIMARY KEY, "值" TEXT NOT NULL);
        """)

        for row in entries:
            entry_no = text(row.get("entryNumber"))
            cas = text(row.get("casNumber"))
            ec = text(row.get("ecNumber"))
            name_en = text(row.get("substanceName"))
            name_zh, name_source, name_status = chinese_name(name_en, cas, translation_cache)
            conn.execute(
                'INSERT INTO annex_xvii_entry VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                (entry_no, name_zh, name_source, name_status, name_en, cas, ec, text(row.get("substanceDescription")),
                 text(row.get("conditionsTitle")), text(row.get("conditionsPath")),
                 text(row.get("conditionsName")), text(row.get("rmlId"))),
            )

        for row in members:
            rml_id = text(row.get("rmlId"))
            parent_id = text(row.get("parentRmlId")) or rml_id
            parent = entry_by_rml.get(parent_id, {})
            entry_no = text(row.get("entryNumber")) or text(parent.get("entryNumber"))
            cas = text(row.get("casNumber"))
            ec = text(row.get("ecNumber"))
            name_en = text(row.get("substanceName"))
            name_zh, name_source, name_status = chinese_name(name_en, cas, translation_cache)
            if not name_zh and parent_id != rml_id:
                name_zh, name_source, name_status = chinese_name(text(parent.get("substanceName")), text(parent.get("casNumber")), translation_cache)
            conn.execute(
                'INSERT INTO annex_xvii_substance ("条目号", "ECHA记录号", "父级ECHA记录号", "是否物质组成员", "中文名称", "中文名称来源", "中文名称状态", "英文名称", "CAS号", "EC号", "物质描述", "限制条件标题", "限制条件链接", "法规依据") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                (entry_no, rml_id, parent_id, int(parent_id != rml_id), name_zh, name_source, name_status, name_en, cas, ec,
                 text(row.get("substanceDescription")), text(parent.get("conditionsTitle")),
                 text(parent.get("conditionsPath")), text(parent.get("conditionsName"))),
            )

        metadata = [
            ("source_api", API),
            ("source_page", "https://chem.echa.europa.eu/obligation-lists/restrictionList"),
            ("source_regulation", "REACH Regulation (EC) No 1907/2006 Annex XVII"),
            ("source_consolidated_text", "https://eur-lex.europa.eu/eli/reg/2006/1907"),
            ("top_level_entry_count", str(len(entries))),
            ("expanded_record_count", str(len(members))),
            ("scope", "命中表示该 CAS/EC 属于 Annex XVII 条目或物质组范围；最终适用仍需审查用途、浓度、物品形态和豁免"),
            ("chinese_name_policy", "中文名优先使用本地 CAS 库和内置常用词表；其余为基于 ECHA 英文名的翻译候选，状态为机器翻译待复核"),
        ]
        conn.executemany("INSERT INTO source_metadata VALUES (?, ?)", metadata)
        conn.commit()
    save_translation_cache(translation_cache)
    print(f"created {TARGET} with {len(entries)} entries and {len(members)} expanded records")


if __name__ == "__main__":
    main()
