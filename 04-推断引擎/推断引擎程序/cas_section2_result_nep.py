"""Populate the confirmed CAS Section 2 result skeleton for NEP.

NEP (CAS 2687-91-4) is the second sample profile after NMP.  The module
deliberately keeps the identity library read-only and records the distinction
between a traceable substance classification, Chinese mixture-disclosure
thresholds, and supplier SDS observations.
"""

from __future__ import annotations

import argparse
import json
import re
import sqlite3
from pathlib import Path
from typing import Any

from cas_section2_result_db import (
    DEFAULT_DB,
    STANDARD_SKELETON_FIELDS,
    connect,
    initialize,
    upsert_condition,
    upsert_fact,
    upsert_source,
    utc_now,
    refresh_standard_result_fields,
    validate_standard_result_text,
)
from cas_section2_result_enrichment import local_source, remote_source


PROJECT_ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
CAS_LIBRARY_DB = PROJECT_ROOT / "03-数据库" / "正式库" / "Data Base" / "cas_library.db"
TEMPLATE_PATH = PROJECT_ROOT / "04-推断引擎" / "数据库模板" / "N-甲基-2-吡咯烷酮-NMP-安全数据表.xlsx"
NEP_CAS = "2687-91-4"
NEP_NAME = "N-乙基吡咯烷酮"
NEP_EC = "220-250-6"
NEP_INDEX = "616-208-00-5"


def _model_key(model: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", model).strip("_")
    return f"original_sds_{value or 'unknown'}"


def read_identity_sources(identity_db: Path = CAS_LIBRARY_DB) -> list[dict[str, Any]]:
    """Read NEP model source paths from the identity DB without opening it writable."""
    if not identity_db.is_file():
        raise FileNotFoundError(f"找不到总 CAS 身份库：{identity_db}")
    uri = f"file:{identity_db.as_posix()}?mode=ro"
    with sqlite3.connect(uri, uri=True) as connection:
        connection.row_factory = sqlite3.Row
        row = connection.execute(
            "SELECT cas_id, standard_name FROM cas_substance WHERE cas_no = ?",
            (NEP_CAS,),
        ).fetchone()
        if row is None:
            raise ValueError(f"CAS 身份库没有 {NEP_CAS} 的唯一身份记录。")
        usages = connection.execute(
            """
            SELECT model, source_file, concentration
            FROM cas_model_usage WHERE cas_id = ?
            ORDER BY model, source_file
            """,
            (int(row["cas_id"]),),
        ).fetchall()
    if not usages:
        raise ValueError(f"CAS 身份库没有 {NEP_CAS} 的型号来源记录。")
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for usage in usages:
        key = (str(usage["model"]), str(usage["source_file"]))
        if key in seen:
            continue
        seen.add(key)
        result.append({"model": key[0], "source_file": key[1], "concentration": usage["concentration"]})
    return result


def source_specs(identity_db: Path = CAS_LIBRARY_DB) -> list[dict[str, Any]]:
    common = [
        local_source(
            "GB30000.1-2024", "化学品分类和标签规范 第1部分：通则",
            r"08_未分类_待核验\GB30000.1-2024.pdf", source_type="mandatory_national_standard",
            version="2024", effective="2025-08-01", clause="第1部分：通则",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C008AEDBFD9A16F3C5BEB671C20618DD",
            evidence="中国 GHS 分类、标签和象形图总则；本来源不包含 NEP 专属条目。",
        ),
        local_source(
            "GB30000.19-2013", "化学品分类和标签规范 第19部分：皮肤腐蚀/刺激",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.19-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第19部分：皮肤腐蚀/刺激",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C5B87B85478148F2AD9FBC80991F2D92",
            evidence="皮肤腐蚀/刺激分类框架；不据此直接把产品混合物的 H318 推成 NEP 物质级结论。",
        ),
        local_source(
            "GB30000.20-2013", "化学品分类和标签规范 第20部分：严重眼损伤/眼刺激",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.20-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第20部分：严重眼损伤/眼刺激",
            uri="https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB+30000&p.p90=circulation_date&p.p91=desc&page=1&pageSize=50&r=0.7951120857180122",
            evidence="眼损伤/眼刺激分类框架；本档案保留中文产品 SDS 的 H318 作为引用事实，未提升为 NEP 固有分类。",
        ),
        local_source(
            "GB30000.24-2013", "化学品分类和标签规范 第24部分：生殖毒性",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.24-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第24部分：生殖毒性",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=71308BF717E07CF1FD39875FA52F1BBB",
            evidence="生殖毒性分类框架；用于后续中国本地化复核。",
        ),
        local_source(
            "GB30000.25-2013", "化学品分类和标签规范 第25部分：特定靶器官毒性——单次接触",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.25-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第25部分：特定靶器官毒性——单次接触",
            uri="https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB+30000&p.p90=circulation_date&p.p91=desc&page=1&pageSize=50&r=0.7951120857180122",
            evidence="STOT 单次接触分类框架；NEP 本地原始 SDS 对该项记录为无数据，不新增固有分类。",
        ),
        local_source(
            "GB30000.28-2013", "化学品分类和标签规范 第28部分：对水生环境的危害",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.28-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第28部分：对水生环境的危害",
            uri="https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB+30000&p.p90=circulation_date&p.p91=desc&page=1&pageSize=50&r=0.7951120857180122",
            evidence="水生环境危害分类框架；本档案只登记供应商生态毒理事实，不据此自动生成水生危害分类。",
        ),
        local_source(
            "GB-T17519-2013", "化学品安全技术说明书编写指南",
            r"08_未分类_待核验\GB-T17519-2013.pdf", source_type="recommended_national_standard",
            version="2013", effective="2014-01-31", clause="3.3.2、表1",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=0DF24B277B29E2F0B39D13D0F2B7BA07",
            evidence="GB/T 17519-2013 现行；用于组分名称、浓度/浓度范围披露门槛，不替代成品 GHS 分类 SCL。",
        ),
        local_source(
            "GB-T16483-2008", "化学品安全技术说明书 内容和项目顺序",
            r"08_未分类_待核验\GB-T16483-2008.pdf", source_type="recommended_national_standard",
            version="2008", effective="2009-02-01", clause="安全技术说明书内容和项目顺序",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=67992CA972A4CF9222095CA06064724A&refer=outter",
            evidence="SDS 分节边界和项目顺序参考。",
        ),
        local_source(
            "GB15258-2009", "化学品安全标签编写规定",
            r"01_法律法规_国内\03_部门规章\GB 15258-2009 化学品安全标签编写规定.txt",
            source_type="mandatory_national_standard", version="2009", effective="2010-05-01",
            clause="标签要素、信号词、危险性说明和防范说明",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=4D487D68BF0BD87E68CE0EA68183DAD6",
            evidence="标签要素编排和中文表述格式来源；不单独决定 NEP 的物质级危害类别。",
        ),
        local_source(
            "GBZ2.1-2019", "工作场所有害因素职业接触限值 第1部分：化学有害因素（含修改单）",
            r"04_技术标准_防护与应急\02_职业卫生与个体防护\GBZ 2.1-2019 工作场所有害因素职业接触限值 第1部分：化学有害因素(含修改单).pdf",
            source_type="occupational_exposure_standard", version="2019（含修改单）", effective="2020-04-01",
            clause="表1、表4",
            evidence="本地 PDF 表1/表4检索未发现 2687-91-4、N-乙基吡咯烷酮或 N-ethyl-2-pyrrolidone；这是本地标准未列出的证据，不等同于全球不存在职业接触限值。",
        ),
    ]
    sources: list[dict[str, Any]] = [
        {
            "key": "nmp_template_skeleton",
            "source_type": "template_workbook",
            "source_title": "N-甲基-2-吡咯烷酮-NMP-安全数据表（已确认结果骨架参考）",
            "source_file_path": str(TEMPLATE_PATH),
            "source_file_name": TEMPLATE_PATH.name,
            "jurisdiction": "CN",
            "clause_reference": "Sheet1：字段、结果表和条件列结构",
            "source_priority": "internal_reference_only",
            "source_status": "imported",
            "evidence_text": "仅作为已确认 CAS Section 2 结果骨架参考，不作为 NEP 的法规或物质事实来源。",
        },
        *common,
        remote_source(
            "EU-2013-944-NEP", "Commission Regulation (EU) No 944/2013：CLP Annex VI NEP 条目",
            "https://eur-lex.europa.eu/legal-content/EN/TXT/?uri=CELEX:02013R0944-20131003",
            source_type="foreign_regulatory_classification", version="2013/944 consolidated entry",
            clause="Annex VI entry 616-208-00-5", jurisdiction="EU",
            priority="traceable_substance_classification",
            evidence="NEP：EC 220-250-6、CAS 2687-91-4、Repr. 1B、H360D、GHS08、Danger；该条目未列出 NEP 专属混合物浓度限值。",
        ),
        remote_source(
            "ECHA-NEP-RAC-2011", "ECHA RAC 对 N-ethyl-2-pyrrolidone 的协调分类意见",
            "https://echa.europa.eu/documents/10162/1c8a16a7-d82e-108a-3451-1c826a164a99",
            source_type="foreign_regulatory_classification", version="ECHA/RAC/CLH-O-0000002192-83-01/F",
            clause="resulting harmonised classification", jurisdiction="EU",
            priority="traceable_substance_classification",
            evidence="ECHA RAC 意见确认协调分类为 Repr. 1B–H360D；用于与 EUR-Lex 条目交叉核对。",
        ),
        remote_source(
            "ECHA-NEP-INFO", "ECHA Substance Information：1-ethylpyrrolidin-2-one",
            "https://www.echa.europa.eu/web/guest/substance-information/-/substanceinfo/100.018.409?_disssubsinfo_WAR_disssubsinfoportlet_substanceId=100.018.409",
            source_type="identity_reference", version="online", clause="EC/CAS、分子式、法规上下文",
            jurisdiction="EU", priority="reference_only",
            evidence="ECHA 物质卡确认 EC 220-250-6、CAS 2687-91-4、分子式 C6H11NO，并显示其监管上下文。",
        ),
        remote_source(
            "NIST-2687-91-4", "NIST Chemistry WebBook：1-Ethyl-2-pyrrolidinone",
            "https://webbook.nist.gov/cgi/cbook.cgi?ID=2687-91-4&Units=SI",
            source_type="identity_reference", version="SRD 69 online", clause="formula、molecular weight、synonyms",
            jurisdiction="US", priority="internal_reference_only",
            evidence="NIST 确认 C6H11NO、相对分子质量 113.1576 及 NEP 同义名。",
        ),
    ]
    identity_sources = read_identity_sources(identity_db)
    for item in identity_sources:
        source_path = Path(item["source_file"])
        sources.append({
            "key": _model_key(item["model"]),
            "source_type": "historical_original_sds",
            "source_title": f"{item['model']} 中文 MSDS（NEP 原始证据）",
            "source_file_path": str(source_path),
            "source_file_name": source_path.name,
            "source_version": "本地中文 MSDS",
            "jurisdiction": "CN",
            "clause_reference": "Section 2、3、11、12、15（以原文件实际内容为准）",
            "source_priority": "historical_original_sds",
            "source_status": "verified",
            "evidence_text": f"CAS 身份库关联型号来源；Section 3 中记录 NEP/CAS {NEP_CAS}，原始文件含量记录为 {item['concentration']}；用于保留历史原文，不替代现行法规结论。",
        })
    return sources


def _upsert_profile(connection: sqlite3.Connection) -> int:
    row = connection.execute(
        "SELECT profile_id, standard_name_zh, ec_no, index_no FROM cas_s2_profile WHERE cas_no = ? AND is_current = 1",
        (NEP_CAS,),
    ).fetchone()
    now = utc_now()
    if row:
        if (row["standard_name_zh"], row["ec_no"], row["index_no"]) != (NEP_NAME, NEP_EC, NEP_INDEX):
            raise ValueError("结果库已有 NEP 当前档案，但身份字段与本次标准值不一致，拒绝覆盖。")
        profile_id = int(row["profile_id"])
        connection.execute(
            """
            UPDATE cas_s2_profile
            SET profile_version = '1.0.0', regulatory_version = ?, profile_status = 'manual_review',
                remarks = ?, updated_at = ? WHERE profile_id = ?
            """,
            (
                "GB 30000.1-2024（2025-08-01生效）；GB 30000.19/20/24/25/28-2013（现行）；GB/T 17519-2013；EU CLP Annex VI entry 616-208-00-5；ECHA RAC 2011",
                "NEP Section 2 标准结果档案。",
                now,
                profile_id,
            ),
        )
        return profile_id
    cursor = connection.execute(
        """
        INSERT INTO cas_s2_profile (
            cas_no, standard_name_zh, ec_no, index_no, profile_version,
            regulatory_version, profile_status, is_current, remarks, created_at, updated_at
        ) VALUES (?, ?, ?, ?, '1.0.0', ?, 'manual_review', 1, ?, ?, ?)
        """,
        (
            NEP_CAS,
            NEP_NAME,
            NEP_EC,
            NEP_INDEX,
            "GB 30000.1-2024（2025-08-01生效）；GB 30000.19/20/24/25/28-2013（现行）；GB/T 17519-2013；EU CLP Annex VI entry 616-208-00-5；ECHA RAC 2011",
            "NEP Section 2 标准结果档案。",
            now,
            now,
        ),
    )
    return int(cursor.lastrowid)


def _upsert_classification(connection: sqlite3.Connection, profile_id: int, row: dict[str, Any]) -> None:
    connection.execute(
        """
        INSERT INTO cas_s2_classification (
            profile_id, hazard_domain, hazard_class, category, classification_text, h_code,
            condition_id, condition_key, value_status, display_order, source_id, raw_text
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(profile_id, hazard_domain, classification_text, h_code, condition_key)
        DO UPDATE SET hazard_class=excluded.hazard_class, category=excluded.category,
            condition_id=excluded.condition_id, value_status=excluded.value_status,
            display_order=excluded.display_order, source_id=excluded.source_id,
            raw_text=excluded.raw_text
        """,
        (
            profile_id, row["hazard_domain"], row["hazard_class"], row["category"],
            row["classification_text"], row["h_code"], row["condition_id"],
            row["condition_key"], row["value_status"], row["display_order"],
            row["source_id"], row["raw_text"],
        ),
    )


def _upsert_label(connection: sqlite3.Connection, profile_id: int, row: dict[str, Any]) -> None:
    connection.execute(
        """
        INSERT INTO cas_s2_label_element (
            profile_id, element_type, element_code, element_text_zh, condition_id,
            condition_key, value_status, display_order, source_id, raw_text
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(profile_id, element_type, element_code, condition_key)
        DO UPDATE SET element_text_zh=excluded.element_text_zh, condition_id=excluded.condition_id,
            value_status=excluded.value_status, display_order=excluded.display_order,
            source_id=excluded.source_id, raw_text=excluded.raw_text
        """,
        (
            profile_id, row["element_type"], row["element_code"], row["element_text_zh"],
            row["condition_id"], row["condition_key"], row["value_status"],
            row["display_order"], row["source_id"], row["raw_text"],
        ),
    )


def _upsert_summary(connection: sqlite3.Connection, profile_id: int, row: dict[str, Any]) -> None:
    connection.execute(
        """
        INSERT INTO cas_s2_hazard_summary (
            profile_id, hazard_domain, summary_text, value_status, condition_id, source_id, raw_text
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(profile_id, hazard_domain)
        DO UPDATE SET summary_text=excluded.summary_text, value_status=excluded.value_status,
            condition_id=excluded.condition_id, source_id=excluded.source_id, raw_text=excluded.raw_text
        """,
        (profile_id, row["hazard_domain"], row["summary_text"], row["value_status"],
         row["condition_id"], row["source_id"], row["raw_text"]),
    )


def enrich_nep(db_path: Path = DEFAULT_DB, identity_db: Path = CAS_LIBRARY_DB) -> dict[str, Any]:
    initialize(db_path)
    with connect(db_path) as connection:
        profile_id = _upsert_profile(connection)
        source_ids = {item["key"]: upsert_source(connection, profile_id, item) for item in source_specs(identity_db)}

        def condition(**data: Any) -> int:
            data.setdefault("status", "manual_review")
            data.setdefault("unit", "%")
            data.setdefault("priority", 100)
            return upsert_condition(connection, profile_id, data)

        intrinsic_id = condition(
            condition_key="substance_intrinsic", applies_to_type="profile", applies_to_code="*",
            condition_type="unconditional", scope="traceable_substance_classification",
            raw_text="物质级分类结果（不以混合物含量阈值为条件）", status="confirmed", priority=10,
            source_id=source_ids["EU-2013-944-NEP"],
            notes="EU CLP Annex VI/ECHA RAC 的可追溯物质级分类；中国市场正式发布仍需按 GB 30000.1-2024 及对应分册完成本地化复核。",
        )
        condition(
            condition_key="cn_disclosure_reproductive_ge_0_1", applies_to_type="hazard_domain",
            applies_to_code="reproductive_toxicity", condition_type="lower_bound", lower_operator=">=",
            lower_value=0.1, scope="CN_mixture_component_disclosure", raw_text="生殖毒性危险组分含量 ≥ 0.1%",
            status="confirmed", priority=20, source_id=source_ids["GB-T17519-2013"],
            notes="GB/T 17519-2013 3.3.2表1；是混合物组分披露门槛，不是成品 GHS 分类 SCL。",
        )
        condition(
            condition_key="cn_disclosure_skin_ge_1", applies_to_type="hazard_domain",
            applies_to_code="skin_irritation", condition_type="lower_bound", lower_operator=">=",
            lower_value=1.0, scope="CN_mixture_component_disclosure", raw_text="皮肤腐蚀/刺激危险组分含量 ≥ 1.0%",
            status="confirmed", priority=20, source_id=source_ids["GB-T17519-2013"],
            notes="GB/T 17519-2013 3.3.2表1；是混合物组分披露门槛，不是成品 GHS 分类 SCL。",
        )
        condition(
            condition_key="cn_disclosure_eye_ge_1", applies_to_type="hazard_domain",
            applies_to_code="eye_irritation", condition_type="lower_bound", lower_operator=">=",
            lower_value=1.0, scope="CN_mixture_component_disclosure", raw_text="严重眼损伤/眼刺激危险组分含量 ≥ 1.0%",
            status="confirmed", priority=20, source_id=source_ids["GB-T17519-2013"],
            notes="GB/T 17519-2013 3.3.2表1；是混合物组分披露门槛，不是成品 GHS 分类 SCL。",
        )
        condition(
            condition_key="cn_disclosure_stot_se_ge_1", applies_to_type="hazard_domain",
            applies_to_code="specific_target_organ_toxicity_single_exposure", condition_type="lower_bound",
            lower_operator=">=", lower_value=1.0, scope="CN_mixture_component_disclosure",
            raw_text="特定靶器官毒性（单次接触）危险组分含量 ≥ 1.0%", status="confirmed", priority=20,
            source_id=source_ids["GB-T17519-2013"],
            notes="GB/T 17519-2013 3.3.2表1；NEP 原始 MSDS 对该项记录为无数据，不据此新增固有分类。",
        )
        condition(
            condition_key="eu_clp_nep_scl_not_provided", applies_to_type="profile", applies_to_code=NEP_CAS,
            condition_type="not_provided", scope="EU_CLP_mixture_classification",
            raw_text="EU CLP Annex VI entry 616-208-00-5 未列出 NEP 专属混合物浓度限值；不得用中国组分披露阈值替代。",
            status="manual_review", priority=50, source_id=source_ids["EU-2013-944-NEP"],
            notes="混合物须另按适用地区的 GHS/CLP 混合物分类规则计算；当前档案不生成一个未经来源支持的 NEP 专属 SCL。",
        )
        default_id = condition(
            condition_key="profile_default", applies_to_type="profile", applies_to_code="*",
            condition_type="not_provided", scope="mixture_component", raw_text="含量触及条件：",
            status="unknown", priority=100, source_id=source_ids["nmp_template_skeleton"],
            notes="沿用已确认 NMP 骨架的占位条件；无数值、比较符或适用范围时禁止参与自动分类。",
        )

        _upsert_classification(connection, profile_id, {
            "hazard_domain": "reproductive_toxicity", "hazard_class": "生殖毒性", "category": "1B",
            "classification_text": "生殖毒性，类别1B", "h_code": "H360D", "condition_id": intrinsic_id,
            "condition_key": "substance_intrinsic", "value_status": "confirmed", "display_order": 1,
            "source_id": source_ids["EU-2013-944-NEP"], "raw_text": "生殖毒性，类别1B；H360D",
        })
        labels = [
            ("pictogram", "GHS08", "", 1, "GHS08"),
            ("signal_word", "", "危险", 2, "危险"),
            ("hazard_statement", "H360D", "可能损害未出生的婴儿。", 3, "H360D 可能损害未出生的婴儿。"),
            ("precautionary_statement", "P201", "使用前获取特殊说明。", 4, "P201 使用前获取特殊说明。"),
            ("precautionary_statement", "P202", "在阅读并了解所有安全预防措施之前，切勿操作。", 5, "P202 在阅读并了解所有安全预防措施之前，切勿操作。"),
            ("precautionary_statement", "P280", "戴防护手套/防护服/护目镜/面部防护。", 6, "P280 戴防护手套/防护服/护目镜/面部防护。"),
            ("precautionary_statement", "P308+P313", "如果接触或有疑虑：求医/就诊。", 7, "P308+P313 如果接触或有疑虑：求医/就诊。"),
        ]
        for element_type, code, text, order, raw in labels:
            _upsert_label(connection, profile_id, {
                "element_type": element_type, "element_code": code, "element_text_zh": text,
                "condition_id": intrinsic_id, "condition_key": "substance_intrinsic",
                "value_status": "confirmed", "display_order": order,
                "source_id": source_ids["EU-2013-944-NEP"], "raw_text": raw,
            })

        representative_key = "original_sds_PU_2188"
        corroborating_key = "original_sds_PU_2835"

        def fact(domain: str, code: str, label: str, value: str, *, source_key: str,
                 clause: str, status: str = "reference_only", numeric: float | None = None,
                 lower: float | None = None, upper: float | None = None, unit: str = "",
                 qualifier: str = "", scope: str = "supplier_sds_component_fact", order: int = 100) -> None:
            upsert_fact(connection, profile_id, {
                "fact_domain": domain, "fact_code": code, "fact_label_zh": label,
                "value_text": value, "numeric_value": numeric, "lower_value": lower,
                "upper_value": upper, "unit": unit, "qualifier": qualifier,
                "applicability_scope": scope, "value_status": status,
                "source_id": source_ids[source_key], "clause_reference": clause,
                "raw_text": value, "display_order": order,
            })

        fact("identity", "molecular_formula", "分子式", "C6H11NO", source_key="NIST-2687-91-4", clause="identity", status="confirmed", scope="identity_reference")
        fact("identity", "molecular_weight", "相对分子质量", "113.1576", source_key="NIST-2687-91-4", clause="identity", status="confirmed", numeric=113.1576, unit="g/mol", scope="identity_reference")
        fact("health", "acute_oral_ld50", "急性经口毒性 LD50", "3,200", source_key=representative_key, clause="Section 11：急性毒性", numeric=3200, unit="mg/kg（大鼠）")
        fact("health", "acute_inhalation_lc50", "急性吸入毒性 LC50", ">5.1", source_key=representative_key, clause="Section 11：急性毒性", lower=5.1, unit="mg/L（4 h，大鼠）", qualifier="greater_than")
        fact("health", "acute_dermal_ld50", "急性经皮毒性 LD50", ">2,000", source_key=representative_key, clause="Section 11：急性毒性", lower=2000, unit="mg/kg（大鼠）", qualifier="greater_than")
        fact("health", "skin_irritation_test", "皮肤刺激性试验", "兔：未引起实验室动物皮肤刺激", source_key=representative_key, clause="Section 11：主要皮肤刺激性")
        fact("health", "eye_damage_test", "眼损伤试验", "兔：可对眼睛造成严重损伤", source_key=representative_key, clause="Section 11：主要粘膜刺激性")
        fact("health", "skin_sensitization_test", "皮肤致敏性试验", "小鼠：未引起实验室动物过敏", source_key=representative_key, clause="Section 11：致敏性")
        fact("health", "genotoxicity_evidence", "致突变性事实", "细菌和哺乳类细胞培养中未见基因损害", source_key=representative_key, clause="Section 11：致突变性")
        fact("health", "carcinogenicity_evidence", "致癌性事实", "动物试验中未见致癌影响", source_key=representative_key, clause="Section 11：致癌性")
        fact("health", "reproductive_toxicity_evidence", "生殖毒性事实", "动物试验提示对生长发育有影响；高剂量下可能降低生育能力", source_key=representative_key, clause="Section 11：生殖毒性")
        fact("health", "stot_se_data", "特异性靶器官毒性（单次接触）", "无数据资料", source_key=representative_key, clause="Section 11：特异性靶器官系统毒性", status="unknown")
        fact("health", "inhalation_hazard_data", "吸入危险", "无数据资料", source_key=representative_key, clause="Section 11：吸入危险", status="unknown")
        fact("health", "local_sds_section2_reference", "中文原始 MSDS Section 2 引用", "代表性产品 MSDS 对混合物标注：严重眼睛损伤/眼睛刺激性，类别1（H318）；生殖毒性物质，类别1B（H360）", source_key=representative_key, clause="Section 2.1/2.2", scope="CN_original_product_sds_mixture")
        fact("environment", "fish_lc50_96h", "鱼类 LC50", ">464–999", source_key=representative_key, clause="Section 12：生态毒性", lower=464, upper=999, unit="mg/L（96 h，斑马鱼）")
        fact("environment", "algae_ec50_72h", "藻类 EC50", ">101", source_key=representative_key, clause="Section 12：生态毒性", lower=101, unit="mg/L（72 h）", qualifier="greater_than")
        fact("environment", "daphnia_ec50_48h", "大型溞 EC50", ">104", source_key=representative_key, clause="Section 12：生态毒性", lower=104, unit="mg/L（48 h）", qualifier="greater_than")
        fact("environment", "biodegradation_28d", "生物降解性", "90–100%，28 d；快速生物降解", source_key=representative_key, clause="Section 12：持久性和降解性", lower=90, upper=100, unit="%（28 d）")
        fact("regulatory_status", "eu_clp_harmonized_classification", "EU CLP 协调分类", "Repr. 1B；H360D；GHS08；危险", source_key="EU-2013-944-NEP", clause="Annex VI entry 616-208-00-5", status="confirmed", scope="EU_CLP")
        fact("regulatory_status", "eu_clp_nep_scl", "EU CLP NEP 专属浓度限值", "未列出；需按适用的混合物分类规则另行计算", source_key="EU-2013-944-NEP", clause="Annex VI entry 616-208-00-5", status="not_listed", scope="EU_CLP_mixture_classification")
        fact("occupational_exposure", "cn_gbz2_1_nep_listing", "GBZ 2.1-2019 中国职业接触限值条目", "本地标准 PDF 表1/表4未检出 CAS 2687-91-4、N-乙基吡咯烷酮或 N-ethyl-2-pyrrolidone", source_key="GBZ2.1-2019", clause="表1、表4", status="not_listed", scope="CN_GBZ2.1-2019")
        fact("identity", "source_cross_check", "原始 MSDS 交叉核验", "CAS 身份库登记 22 个型号来源；PU-2188 与 PU-2835 的 Section 11/12 NEP 数据表述一致", source_key=corroborating_key, clause="Section 2、3、11、12", status="reference_only", scope="CN_identity_library_source_cross_check")

        summary_rows = [
            ("physical_chemical", "已登记身份与物化基础事实；本次不将未检索到的沸点、闪点等数值填成推断值。", "NIST-2687-91-4"),
            ("health", "EU 可追溯物质级分类为 Repr. 1B/H360D；中文产品 MSDS 的 H318 只保留为混合物引用事实。", "EU-2013-944-NEP"),
            ("environment", "本地中文 MSDS 提供鱼类、藻类、大型溞和生物降解性参考数据；未据此自动生成水生环境危害分类。", representative_key),
            ("other", "暂无可直接升级为 NEP 固有分类的其他危害结论；输出需保留来源范围并进入人工复核。", representative_key),
        ]
        for domain, summary, source_key in summary_rows:
            _upsert_summary(connection, profile_id, {
                "hazard_domain": domain, "summary_text": summary, "value_status": "derived_candidate",
                "condition_id": intrinsic_id, "source_id": source_ids[source_key], "raw_text": summary,
            })
        connection.execute(
            "UPDATE cas_s2_hazard_summary SET condition_id = ? WHERE profile_id = ? AND hazard_domain IN ('physical_chemical','health','environment','other')",
            (intrinsic_id, profile_id),
        )
        review_note = "NEP Section 2 标准结果档案。"
        existing_review = connection.execute(
            "SELECT review_id FROM cas_s2_review WHERE profile_id = ? AND review_scope = 'profile'", (profile_id,)
        ).fetchone()
        if existing_review:
            connection.execute(
                "UPDATE cas_s2_review SET review_status='manual_review', reviewer='', reviewed_at='', review_note=? WHERE review_id=?",
                (review_note, int(existing_review["review_id"])),
            )
        else:
            connection.execute(
                "INSERT INTO cas_s2_review (profile_id, review_scope, review_status, review_note) VALUES (?, 'profile', 'manual_review', ?)",
                (profile_id, review_note),
            )

        refresh_standard_result_fields(connection, profile_id)

        local_count = connection.execute(
            """
            SELECT COUNT(DISTINCT source_file_path || '|' || source_hash_sha256)
            FROM cas_s2_source WHERE profile_id=? AND source_file_path<>'' AND source_file_path NOT LIKE 'remote://%'
            """,
            (profile_id,),
        ).fetchone()[0]
        return {
            "profile_id": profile_id,
            "cas_no": NEP_CAS,
            "standard_name_zh": NEP_NAME,
            "identity_source_model_count": len(read_identity_sources(identity_db)),
            "source_count": connection.execute("SELECT COUNT(*) FROM cas_s2_source WHERE profile_id=?", (profile_id,)).fetchone()[0],
            "local_source_file_count": local_count,
            "condition_count": connection.execute("SELECT COUNT(*) FROM cas_s2_condition WHERE profile_id=?", (profile_id,)).fetchone()[0],
            "classification_count": connection.execute("SELECT COUNT(*) FROM cas_s2_classification WHERE profile_id=?", (profile_id,)).fetchone()[0],
            "label_element_count": connection.execute("SELECT COUNT(*) FROM cas_s2_label_element WHERE profile_id=?", (profile_id,)).fetchone()[0],
            "hazard_summary_count": connection.execute("SELECT COUNT(*) FROM cas_s2_hazard_summary WHERE profile_id=?", (profile_id,)).fetchone()[0],
            "fact_count": connection.execute("SELECT COUNT(*) FROM cas_s2_fact WHERE profile_id=?", (profile_id,)).fetchone()[0],
        }


def validate_nep(db_path: Path = DEFAULT_DB, identity_db: Path = CAS_LIBRARY_DB) -> dict[str, Any]:
    report: dict[str, Any] = {"db": str(db_path), "cas_no": NEP_CAS, "checks": []}

    def check(name: str, passed: bool, detail: Any) -> None:
        report["checks"].append({"name": name, "passed": bool(passed), "detail": detail})

    with connect(db_path) as connection:
        fk_errors = [tuple(row) for row in connection.execute("PRAGMA foreign_key_check")]
        check("foreign_key_check", not fk_errors, fk_errors)
        profiles = connection.execute(
            "SELECT * FROM cas_s2_profile WHERE cas_no=? AND is_current=1", (NEP_CAS,)
        ).fetchall()
        check("one_current_profile", len(profiles) == 1, len(profiles))
        profile = profiles[0] if profiles else None
        if profile:
            profile_id = int(profile["profile_id"])
            standard_rows = connection.execute(
                """
                SELECT field_key, field_label_zh, value_text, value_status, display_order
                FROM v_cas_s2_current_standard_result
                WHERE profile_id = ? ORDER BY display_order
                """,
                (profile_id,),
            ).fetchall()
            expected_field_keys = [key for key, _ in STANDARD_SKELETON_FIELDS]
            actual_field_keys = [row["field_key"] for row in standard_rows]
            check("fixed_standard_skeleton_13_fields", actual_field_keys == expected_field_keys, {
                "actual_count": len(actual_field_keys), "actual": actual_field_keys, "expected": expected_field_keys,
            })
            check("standard_field_values_are_explicit", all(str(row["value_text"] or "").strip() for row in standard_rows), [dict(row) for row in standard_rows])
            result_text_errors = []
            for row in standard_rows:
                try:
                    validate_standard_result_text(row["field_key"], row["value_text"])
                except ValueError as exc:
                    result_text_errors.append(str(exc))
            check("standard_fields_contain_direct_results_only", not result_text_errors, result_text_errors)
            summary_values = {
                row["field_key"]: row["value_text"]
                for row in standard_rows
                if row["field_key"] in {"physical_chemical_hazard", "health_hazard", "environmental_hazard", "other_hazard"}
            }
            check("standard_summary_fields_default_to_no_data", all(value == "无数据" for value in summary_values.values()), summary_values)
            check("standard_field_identity", {
                row["field_key"]: row["value_text"] for row in standard_rows if row["field_key"] in {"standard_name", "cas_no", "ec_no", "index_no"}
            } == {
                "standard_name": NEP_NAME, "cas_no": NEP_CAS, "ec_no": NEP_EC, "index_no": NEP_INDEX,
            }, [dict(row) for row in standard_rows[:4]])
            check("identity_fields", (profile["standard_name_zh"], profile["ec_no"], profile["index_no"]) == (NEP_NAME, NEP_EC, NEP_INDEX), dict(profile))
            check("manual_review_status", profile["profile_status"] == "manual_review", profile["profile_status"])
            rows = connection.execute("SELECT * FROM cas_s2_classification WHERE profile_id=? ORDER BY display_order", (profile_id,)).fetchall()
            classifications = [(row["classification_text"], row["h_code"], row["value_status"]) for row in rows]
            check("only_traceable_intrinsic_classification", classifications == [("生殖毒性，类别1B", "H360D", "confirmed")], classifications)
            labels = connection.execute("SELECT element_type,element_code,value_status FROM cas_s2_label_element WHERE profile_id=? ORDER BY display_order", (profile_id,)).fetchall()
            label_codes = [(row["element_type"], row["element_code"], row["value_status"]) for row in labels]
            check("required_label_elements", {"GHS08", "H360D", "P201", "P202", "P280", "P308+P313"} <= {row[1] for row in label_codes}, label_codes)
            check("no_h318_promoted_to_intrinsic", not connection.execute("SELECT 1 FROM cas_s2_classification WHERE profile_id=? AND h_code='H318'", (profile_id,)).fetchone(), "H318 only retained as a reference fact")
            source_count = connection.execute("SELECT source_count FROM v_cas_s2_current_profile WHERE profile_id=?", (profile_id,)).fetchone()[0]
            distinct_count = connection.execute(
                "SELECT COUNT(DISTINCT source_file_path || '|' || source_hash_sha256) FROM cas_s2_source WHERE profile_id=? AND source_file_path<>'' AND source_file_path NOT LIKE 'remote://%'",
                (profile_id,),
            ).fetchone()[0]
            check("source_count_is_distinct_local_file_count", source_count == distinct_count, {"view": source_count, "distinct_local": distinct_count})
            check("expected_fact_domains", {"identity", "health", "environment", "regulatory_status", "occupational_exposure"} <= {row[0] for row in connection.execute("SELECT DISTINCT fact_domain FROM cas_s2_fact WHERE profile_id=?", (profile_id,))}, "fact domains")
            missing_files = [row[0] for row in connection.execute("SELECT source_file_path FROM cas_s2_source WHERE profile_id=? AND source_file_path NOT LIKE 'remote://%'", (profile_id,)) if not Path(row[0]).is_file()]
            check("local_source_files_exist", not missing_files, missing_files)
    try:
        identity_sources = read_identity_sources(identity_db)
        check("identity_source_read_only_cross_check", len(identity_sources) == 22, {"count": len(identity_sources)})
    except Exception as exc:  # pragma: no cover - diagnostic report
        check("identity_source_read_only_cross_check", False, str(exc))
    report["passed"] = all(item["passed"] for item in report["checks"])
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="NEP CAS Section 2 结果骨架补充工具")
    sub = parser.add_subparsers(dest="command", required=True)
    enrich = sub.add_parser("enrich-nep", help="按 NMP 已确认骨架写入 NEP")
    enrich.add_argument("--db", type=Path, default=DEFAULT_DB)
    enrich.add_argument("--identity-db", type=Path, default=CAS_LIBRARY_DB)
    enrich.add_argument("--report", type=Path)
    validate = sub.add_parser("validate-nep", help="校验 NEP 结果档案")
    validate.add_argument("--db", type=Path, default=DEFAULT_DB)
    validate.add_argument("--identity-db", type=Path, default=CAS_LIBRARY_DB)
    validate.add_argument("--report", type=Path)
    args = parser.parse_args()
    if args.command == "enrich-nep":
        result = enrich_nep(args.db, args.identity_db)
        if args.report:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"command": args.command, "passed": True, **result}, ensure_ascii=False, indent=2))
        return 0
    report = validate_nep(args.db, args.identity_db)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
