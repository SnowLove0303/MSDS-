"""Reusable CAS Section 2 evidence enrichment, with NMP as the first sample.

This module keeps legal sources, structured facts, executable conditions and
output labels separate.  Later CAS imports can reuse the same upsert helpers
and replace only the source/fact/condition specifications.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from cas_section2_result_db import (
    CAS_EXPECTED,
    DEFAULT_DB,
    DEFAULT_TEMPLATE,
    connect,
    initialize,
    read_template,
    upsert_condition,
    upsert_fact,
    upsert_source,
    utc_now,
    refresh_standard_result_fields,
)


PROJECT_ROOT = Path(r"F:\正式项目与模块化内容\冠志\MSDS")
LEGAL_ROOT = PROJECT_ROOT / "04-推断引擎" / "法规匹配库"


def local_source(
    key: str,
    title: str,
    relative_path: str,
    *,
    source_type: str,
    version: str = "",
    effective: str = "",
    clause: str = "",
    priority: str = "current_official_standard",
    uri: str = "",
    jurisdiction: str = "CN",
    evidence: str = "",
) -> dict[str, Any]:
    return {
        "key": key,
        "source_type": source_type,
        "source_title": title,
        "source_file_path": str(LEGAL_ROOT / relative_path),
        "source_version": version,
        "effective_date": effective,
        "jurisdiction": jurisdiction,
        "source_uri": uri,
        "clause_reference": clause,
        "source_priority": priority,
        "source_status": "verified",
        "evidence_text": evidence,
    }


def remote_source(
    key: str,
    title: str,
    uri: str,
    *,
    source_type: str,
    version: str = "",
    effective: str = "",
    clause: str = "",
    priority: str = "reference_only",
    jurisdiction: str = "INT",
    evidence: str = "",
) -> dict[str, Any]:
    return {
        "key": key,
        "source_type": source_type,
        "source_title": title,
        "source_version": version,
        "effective_date": effective,
        "jurisdiction": jurisdiction,
        "source_uri": uri,
        "clause_reference": clause,
        "source_priority": priority,
        "source_status": "verified" if priority != "internal_reference_only" else "reference_only",
        "evidence_text": evidence,
    }


def source_specs(template_path: Path) -> list[dict[str, Any]]:
    return [
        {
            "key": "nmp_template_workbook",
            "source_type": "template_workbook",
            "source_title": "N-甲基-2-吡咯烷酮-NMP-安全数据表（用户模板）",
            "source_file_path": str(template_path),
            "source_file_name": template_path.name,
            "jurisdiction": "CN",
            "clause_reference": "Sheet1",
            "source_priority": "internal_reference_only",
            "source_status": "imported",
            "evidence_text": "模板原始结构与标签文字；空白项不解释为无危害或无数据。",
        },
        local_source(
            "GB30000.1-2024", "化学品分类和标签规范 第1部分：通则",
            r"08_未分类_待核验\GB30000.1-2024.pdf", source_type="mandatory_national_standard",
            version="2024", effective="2025-08-01", clause="第1部分：通则",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C008AEDBFD9A16F3C5BEB671C20618DD",
            evidence="中国现行 GHS 分类、标签和象形图总则；不包含 NMP 的专属成分条目。",
        ),
        local_source(
            "GB30000.19-2013", "化学品分类和标签规范 第19部分：皮肤腐蚀/刺激",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.19-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第19部分：皮肤腐蚀/刺激",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=C5B87B85478148F2AD9FBC80991F2D92",
            evidence="皮肤腐蚀/刺激分类方法和标签要素来源。",
        ),
        local_source(
            "GB30000.20-2013", "化学品分类和标签规范 第20部分：严重眼损伤/眼刺激",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.20-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第20部分：严重眼损伤/眼刺激",
            uri="https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB+30000&p.p90=circulation_date&p.p91=desc&page=1&pageSize=50&r=0.7951120857180122",
            evidence="中国现行 GHS 眼损伤/眼刺激分类分册；官方目录列为现行。",
        ),
        local_source(
            "GB30000.24-2013", "化学品分类和标签规范 第24部分：生殖毒性",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.24-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第24部分：生殖毒性",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=71308BF717E07CF1FD39875FA52F1BBB",
            evidence="中国现行 GHS 生殖毒性分类分册。",
        ),
        local_source(
            "GB30000.25-2013", "化学品分类和标签规范 第25部分：特定靶器官毒性——单次接触",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.25-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第25部分：特定靶器官毒性——单次接触",
            uri="https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB+30000&p.p90=circulation_date&p.p91=desc&page=1&pageSize=50&r=0.7951120857180122",
            evidence="中国现行 GHS STOT 单次接触分类分册；官方目录列为现行。",
        ),
        local_source(
            "GB30000.28-2013", "化学品分类和标签规范 第28部分：对水生环境的危害",
            r"03_技术标准_分类与框架\01_GHS分类_GB30000系列\GB 30000.28-2013.pdf",
            source_type="mandatory_national_standard", version="2013", effective="2014-11-01",
            clause="第28部分：对水生环境的危害",
            uri="https://openstd.samr.gov.cn/bzgk/std/std_list?p.p1=0&p.p2=GB+30000&p.p90=circulation_date&p.p91=desc&page=1&pageSize=50&r=0.7951120857180122",
            evidence="中国现行 GHS 水生环境危害分类分册；本 NMP 样本未据此新增水生危害分类。",
        ),
        local_source(
            "GB-T17519-2013", "化学品安全技术说明书编写指南",
            r"08_未分类_待核验\GB-T17519-2013.pdf", source_type="recommended_national_standard",
            version="2013", effective="2014-01-31", clause="3.3.2、表1",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=0DF24B277B29E2F0B39D13D0F2B7BA07",
            evidence="3.3.2及表1规定危险组分名称、浓度或浓度范围的混合物组分披露门槛；该门槛不等同于成品 GHS 分类 SCL。",
        ),
        local_source(
            "GB-T16483-2008", "化学品安全技术说明书 内容和项目顺序",
            r"08_未分类_待核验\GB-T16483-2008.pdf", source_type="recommended_national_standard",
            version="2008", effective="2009-02-01", clause="安全技术说明书内容和项目顺序",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=67992CA972A4CF9222095CA06064724A&refer=outter",
            evidence="SDS 分节边界和项目顺序参考，不直接推导 NMP 分类。",
        ),
        local_source(
            "GB15258-2009", "化学品安全标签编写规定",
            r"01_法律法规_国内\03_部门规章\GB 15258-2009 化学品安全标签编写规定.txt",
            source_type="mandatory_national_standard", version="2009", effective="2010-05-01",
            clause="标签要素、信号词、危险性说明和防范说明",
            uri="https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=4D487D68BF0BD87E68CE0EA68183DAD6",
            evidence="标签要素编排和中文表述格式来源；不单独决定 NMP 的危害类别。",
        ),
        local_source(
            "GBZ2.1-2019", "工作场所有害因素职业接触限值 第1部分：化学有害因素（含修改单）",
            r"04_技术标准_防护与应急\02_职业卫生与个体防护\GBZ 2.1-2019 工作场所有害因素职业接触限值 第1部分：化学有害因素(含修改单).pdf",
            source_type="occupational_exposure_standard", version="2019（含修改单）",
            effective="2020-04-01", clause="表1、表4", evidence="本地官方 PDF 的表1/表4文本检索未发现 NMP、CAS 872-50-4 或 N-甲基-2-吡咯烷酮条目；这只是本地标准未列出的证据，不等同于全球不存在任何职业接触限值。",
        ),
        remote_source(
            "GB-T27563-2011", "工业用N-甲基-2-吡咯烷酮",
            "https://std.samr.gov.cn/gb/search/gbDetailed?id=71F772D7DDD4D3A7E05397BE0A0AB82A",
            source_type="product_standard", version="2011", effective="2012-05-01",
            clause="产品标准身份和质量指标参考", jurisdiction="CN",
            priority="current_official_standard", evidence="官方标准查询显示 GB/T 27563-2011 为现行产品标准；不直接替代 GHS 分类规则。",
        ),
        remote_source(
            "EU-2016-1179", "Commission Regulation (EU) 2016/1179，CLP Annex VI NMP entry",
            "https://eur-lex.europa.eu/eli/reg/2016/1179/oj/eng", source_type="foreign_regulatory_classification",
            version="2016/1179", effective="2018-03-01", clause="Annex VI, entry 606-021-00-7",
            jurisdiction="EU", priority="traceable_substance_classification",
            evidence="现行 EU 协调分类：Repr. 1B、STOT SE 3、Skin Irrit. 2、Eye Irrit. 2；H360D/H335/H315/H319；GHS08/GHS07；Dgr；专属浓度限值仅登记 STOT SE 3，C≥10%。",
        ),
        remote_source(
            "EU-2018-588", "Commission Regulation (EU) 2018/588，REACH Annex XVII entry 71",
            "https://eur-lex.europa.eu/legal-content/EN/TXT/PDF/?uri=CELEX%3A32018R0588", source_type="foreign_regulatory_restriction",
            version="2018/588", effective="2020-05-09", clause="Annex XVII, entry 71", jurisdiction="EU",
            priority="current_official_standard", evidence="EU REACH 限制：NMP 浓度达到或超过0.3%时触发特定限制义务；这是欧盟监管限制，不是中国 GHS 成品分类阈值。",
        ),
        remote_source(
            "MEM-NMP-2022-QA", "应急管理部关于 NMP/GBL 是否列入危险化学品目录的公开答复",
            "https://www.mem.gov.cn/hd/gzly/lyhf/202207/t20220711_418070.shtml", source_type="official_regulatory_qa",
            version="2022-06-13", clause="2022-06-13公开答复", jurisdiction="CN",
            priority="current_official_standard", evidence="公开答复称 NMP 和 GBL 未列入《危险化学品目录（2015版）》；按日期和问题范围保存，不能冒充当前目录全文或直接推导 GHS 分类。",
        ),
        remote_source(
            "ICSC-0513", "ILO/WHO IPCS ICSC 0513：N-甲基-2-吡咯烷酮",
            "https://www.inchem.org/documents/icsc/icsc/eics0513.htm", source_type="international_safety_card",
            version="ICSC 0513", clause="identity、physical properties、health hazards、fire/explosion",
            priority="internal_reference_only", evidence="国际安全卡事实源，用于补充物化、暴露途径和危害事实；不替代中国法规分类。",
        ),
        remote_source(
            "CICAD-35", "WHO/IPCS CICAD 35：N-Methyl-2-pyrrolidone",
            "https://www.inchem.org/documents/cicads/cicads/cicad35.htm", source_type="international_assessment",
            version="CICAD 35", clause="environmental fate、health-based guidance、physical properties",
            priority="internal_reference_only", evidence="国际评估事实源，用于环境归趋、吸收/渗透和参考健康值；不得当作中国法定职业接触限值。",
        ),
        remote_source(
            "NIST-872-50-4", "NIST Chemistry WebBook：CAS 872-50-4",
            "https://webbook.nist.gov/cgi/cbook.cgi?ID=872-50-4&Units=SI", source_type="identity_reference",
            version="online", clause="identity、formula、molecular weight", jurisdiction="US",
            priority="internal_reference_only", evidence="身份和基础物化参考源。",
        ),
        remote_source(
            "SIGMA-270458", "Merck/Sigma-Aldrich NMP product data",
            "https://www.sigmaaldrich.cn/CN/zh/product/sigald/270458", source_type="supplier_reference",
            version="online", clause="identity、physical properties、supplier GHS reference",
            priority="direct_product_test_or_supplier_confirmed_fact", evidence="供应商产品资料，仅用于补充方法/批次相关物化参考；法规分类以可追溯法规条目为准。",
        ),
    ]


def enrich_nmp(db_path: Path, template_path: Path) -> dict[str, Any]:
    expected = read_template(template_path)
    initialize(db_path)
    now = utc_now()
    with connect(db_path) as connection:
        profile = connection.execute(
            "SELECT * FROM cas_s2_profile WHERE cas_no = ? AND is_current = 1", (CAS_EXPECTED,)
        ).fetchone()
        if profile is None:
            raise ValueError("当前结果库没有 NMP 当前档案，请先执行 seed-nmp。")
        profile_id = int(profile["profile_id"])

        source_ids = {
            spec["key"]: upsert_source(connection, profile_id, spec)
            for spec in source_specs(template_path)
        }

        def condition(**data: Any) -> int:
            data.setdefault("status", "manual_review")
            data.setdefault("unit", "%")
            data.setdefault("priority", 100)
            return upsert_condition(connection, profile_id, data)

        intrinsic_id = condition(
            condition_key="substance_intrinsic", applies_to_type="profile", applies_to_code="*",
            condition_type="unconditional", scope="traceable_substance_classification",
            raw_text="物质级分类结果（不以混合物含量阈值为条件）", status="confirmed", priority=10,
            source_id=source_ids["EU-2016-1179"],
            notes="现行 EU CLP Annex VI 的可追溯物质级交叉证据；中国市场正式发布仍需结合 GB 30000.1-2024 及对应分册完成本地化复核。",
        )
        threshold_rows = [
            ("cn_disclosure_reproductive_ge_0_1", "reproductive_toxicity", 0.1, "生殖毒性危险组分含量 ≥ 0.1%"),
            ("cn_disclosure_skin_ge_1", "skin_irritation", 1.0, "皮肤腐蚀/刺激危险组分含量 ≥ 1.0%"),
            ("cn_disclosure_eye_ge_1", "eye_irritation", 1.0, "严重眼损伤/眼刺激危险组分含量 ≥ 1.0%"),
            ("cn_disclosure_stot_se_ge_1", "specific_target_organ_toxicity_single_exposure", 1.0, "特定靶器官毒性（单次接触）危险组分含量 ≥ 1.0%"),
        ]
        for key, applies_code, value, raw in threshold_rows:
            condition(
                condition_key=key, applies_to_type="hazard_domain", applies_to_code=applies_code,
                condition_type="lower_bound", lower_operator=">=", lower_value=value,
                scope="CN_mixture_component_disclosure", raw_text=raw, status="confirmed", priority=20,
                source_id=source_ids["GB-T17519-2013"],
                notes="GB/T 17519-2013 3.3.2表1；这是混合物组分名称/浓度披露门槛，不等同于成品 GHS 分类 SCL。",
            )
        condition(
            condition_key="eu_clp_nmp_stot_se_ge_10", applies_to_type="h_code", applies_to_code="H335",
            condition_type="lower_bound", lower_operator=">=", lower_value=10.0,
            scope="EU_CLP_mixture_classification", raw_text="STOT SE 3；H335：C ≥ 10%",
            status="confirmed", priority=30, source_id=source_ids["EU-2016-1179"],
            notes="EU CLP Annex VI entry 606-021-00-7 的现行 NMP 专属浓度限值；仅适用于 EU CLP，不得写成中国 GHS 阈值。",
        )
        condition(
            condition_key="eu_clp_nmp_repr_ge_5_historical", applies_to_type="h_code", applies_to_code="H360D",
            condition_type="manual_review", lower_operator=">=", lower_value=5.0,
            scope="EU_CLP_historical_superseded", raw_text="历史 EU 条目曾记录 Repr. 1B；H360D：C ≥ 5%",
            status="manual_review", priority=900, source_id=source_ids["EU-2016-1179"],
            notes="历史阈值来自旧法规版本，现行 2016/1179 条目不再登记该 5% 生殖毒性专属限值；保留用于版本审计，禁止执行。",
        )
        condition(
            condition_key="eu_reach_nmp_ge_0_3_restriction", applies_to_type="profile", applies_to_code=CAS_EXPECTED,
            condition_type="lower_bound", lower_operator=">=", lower_value=0.3,
            scope="EU_REACH_regulatory_restriction", raw_text="NMP 浓度 ≥ 0.3% 时触发 EU REACH Annex XVII entry 71 限制义务",
            status="confirmed", priority=40, source_id=source_ids["EU-2018-588"],
            notes="这是 EU REACH 监管限制触发点，不是 GHS 成品分类或中国 SDS 组分披露阈值。",
        )

        connection.execute(
            """
            UPDATE cas_s2_profile
            SET profile_version = '1.1.0', regulatory_version = ?, profile_status = 'manual_review',
                remarks = ?, updated_at = ? WHERE profile_id = ?
            """,
            (
                "GB 30000.1-2024（2025-08-01生效）；GB 30000.19/20/24/25/28-2013（现行）；GB/T 17519-2013；EU CLP 2016/1179（交叉证据）；EU REACH 2018/588（限制，非中国分类）",
                "NMP Section 2 标准结果档案。",
                now,
                profile_id,
            ),
        )
        connection.execute(
            "UPDATE cas_s2_condition SET notes = ? WHERE profile_id = ? AND condition_key = 'profile_default'",
            ("模板占位条件：原始表单未提供数值或比较符；保留用于追溯，禁止作为执行条件。", profile_id),
        )
        connection.execute(
            "UPDATE cas_s2_classification SET condition_id = ?, condition_key = 'substance_intrinsic', value_status = 'confirmed', source_id = ? WHERE profile_id = ?",
            (intrinsic_id, source_ids["EU-2016-1179"], profile_id),
        )
        connection.execute(
            "UPDATE cas_s2_label_element SET condition_id = ?, condition_key = 'substance_intrinsic', value_status = 'confirmed', source_id = ? WHERE profile_id = ?",
            (intrinsic_id, source_ids["EU-2016-1179"], profile_id),
        )

        def fact(domain: str, code: str, label: str, *, source_key: str, clause: str,
                 value_text: str = "", numeric: float | None = None, lower: float | None = None,
                 upper: float | None = None, unit: str = "", qualifier: str = "",
                 scope: str = "", status: str = "reference_only", order: int = 100) -> None:
            upsert_fact(connection, profile_id, {
                "fact_domain": domain, "fact_code": code, "fact_label_zh": label,
                "value_text": value_text, "numeric_value": numeric, "lower_value": lower,
                "upper_value": upper, "unit": unit, "qualifier": qualifier,
                "applicability_scope": scope, "value_status": status,
                "source_id": source_ids[source_key], "clause_reference": clause,
                "raw_text": value_text, "display_order": order,
            })

        fact("identity", "molecular_formula", "分子式", source_key="NIST-872-50-4", clause="identity", value_text="C5H9NO", status="confirmed")
        fact("identity", "molecular_weight", "相对分子质量", source_key="NIST-872-50-4", clause="identity", value_text="99.13", numeric=99.13, unit="g/mol", status="confirmed")
        fact("physical_chemical", "appearance", "外观", source_key="ICSC-0513", clause="physical properties", value_text="无色、吸湿性液体，有特征气味", status="confirmed")
        fact("physical_chemical", "boiling_point", "沸点", source_key="ICSC-0513", clause="physical properties", value_text="202", numeric=202, unit="℃", status="confirmed")
        fact("physical_chemical", "melting_point", "熔点", source_key="ICSC-0513", clause="physical properties", value_text="-24.4", numeric=-24.4, unit="℃", status="confirmed")
        fact("physical_chemical", "relative_density", "相对密度", source_key="ICSC-0513", clause="physical properties", value_text="1.03", numeric=1.03, unit="无量纲", status="confirmed")
        fact("physical_chemical", "water_solubility", "水溶性", source_key="ICSC-0513", clause="physical properties", value_text="与水完全混溶", status="confirmed")
        fact("physical_chemical", "vapor_pressure", "蒸气压", source_key="ICSC-0513", clause="physical properties", value_text="39", numeric=39, unit="Pa（25℃）", qualifier="25℃", status="confirmed")
        fact("physical_chemical", "vapor_pressure", "蒸气压", source_key="SIGMA-270458", clause="supplier physical data", value_text="0.29", numeric=0.29, unit="mmHg（20℃）", qualifier="20℃", status="reference_only")
        fact("physical_chemical", "vapor_density", "蒸气相对密度", source_key="ICSC-0513", clause="physical properties", value_text="3.4", numeric=3.4, unit="空气=1", status="confirmed")
        fact("physical_chemical", "flash_point", "闪点", source_key="ICSC-0513", clause="fire and explosion", value_text="86", numeric=86, unit="℃", qualifier="闭杯", status="confirmed")
        fact("physical_chemical", "flash_point", "闪点", source_key="SIGMA-270458", clause="supplier physical data", value_text="91", numeric=91, unit="℃", qualifier="闭杯，供应商数据", status="reference_only")
        fact("physical_chemical", "autoignition_temperature", "自燃温度", source_key="ICSC-0513", clause="fire and explosion", value_text="245", numeric=245, unit="℃", qualifier="ICSC", status="confirmed")
        fact("physical_chemical", "explosive_limits", "爆炸极限", source_key="ICSC-0513", clause="fire and explosion", value_text="1.3–9.5", lower=1.3, upper=9.5, unit="体积分数%", status="confirmed")
        fact("physical_chemical", "log_pow", "正辛醇/水分配系数 log Pow", source_key="CICAD-35", clause="physical properties", value_text="-0.38", numeric=-0.38, unit="无量纲", status="confirmed")
        fact("physical_chemical", "kinematic_viscosity", "运动黏度", source_key="ICSC-0513", clause="physical properties", value_text="1.62", numeric=1.62, unit="mm²/s（25℃）", status="reference_only")
        fact("health", "exposure_routes", "主要暴露途径", source_key="ICSC-0513", clause="routes of exposure", value_text="吸入、皮肤接触、吞咽", status="confirmed")
        fact("health", "irritation_effects", "刺激性危害事实", source_key="ICSC-0513", clause="health effects", value_text="可引起眼睛和呼吸道刺激；可引起轻度皮肤刺激，长期皮肤接触可致皮炎", status="confirmed")
        fact("health", "dermal_absorption", "皮肤吸收/渗透", source_key="ICSC-0513", clause="health effects", value_text="可经皮吸收，并可增强其他物质经皮渗透", status="confirmed")
        fact("health", "reproductive_toxicity_evidence", "生殖毒性依据", source_key="EU-2016-1179", clause="Annex VI entry 606-021-00-7", value_text="现行 EU CLP Annex VI 追溯分类为 Repr. 1B，H360D", scope="EU_CLP", status="confirmed")
        fact("occupational_exposure", "cn_gbz2_1_nmp_listing", "GBZ 2.1-2019 中国职业接触限值条目", source_key="GBZ2.1-2019", clause="表1、表4", value_text="本地标准 PDF 表1/表4未检出 NMP 条目", scope="CN_GBZ2.1-2019", status="not_listed")
        fact("occupational_exposure", "eu_oel_twa", "EU-OEL 时间加权平均值", source_key="ICSC-0513", clause="occupational exposure limits", value_text="40", numeric=40, unit="mg/m³（8小时）", scope="EU foreign reference", status="reference_only")
        fact("occupational_exposure", "eu_oel_stel", "EU-OEL 短时间接触限值", source_key="ICSC-0513", clause="occupational exposure limits", value_text="80", numeric=80, unit="mg/m³（短时间）", scope="EU foreign reference", status="reference_only")
        fact("occupational_exposure", "eu_reach_dnel_inhalation", "EU REACH DNEL 吸入", source_key="EU-2018-588", clause="Annex XVII entry 71", value_text="14.4", numeric=14.4, unit="mg/m³", scope="EU_REACH_entry_71", status="reference_only")
        fact("occupational_exposure", "eu_reach_dnel_dermal", "EU REACH DNEL 皮肤", source_key="EU-2018-588", clause="Annex XVII entry 71", value_text="4.8", numeric=4.8, unit="mg/kg体重/日", scope="EU_REACH_entry_71", status="reference_only")
        fact("environment", "environmental_fate", "环境归趋", source_key="CICAD-35", clause="environmental fate", value_text="与水完全混溶；在土壤中具有迁移性；不预期明显生物富集；可快速生物降解", status="reference_only")
        fact("environment", "bioaccumulation_potential", "生物富集潜势", source_key="CICAD-35", clause="environmental fate", value_text="不预期明显生物富集（log Kow -0.38）", status="reference_only")
        fact("regulatory_status", "cn_hazardous_chemicals_catalogue_2015", "中国危险化学品目录状态", source_key="MEM-NMP-2022-QA", clause="2022-06-13公开答复", value_text="2022-06-13官方公开答复称 NMP 未列入《危险化学品目录（2015版）》", scope="CN_dated_official_qa", status="reference_only")
        fact("regulatory_status", "eu_reach_restriction", "EU REACH 限制", source_key="EU-2018-588", clause="Annex XVII entry 71", value_text="浓度 ≥0.3% 时触发 Annex XVII entry 71 特定义务", numeric=0.3, unit="%", scope="EU_REACH_entry_71", status="confirmed")
        fact("other_hazard", "skin_permeation_enhancement", "其他危害", source_key="ICSC-0513", clause="health effects", value_text="可增强其他物质经皮渗透", status="confirmed")

        summaries = {
            "physical_chemical": ("可燃性物化事实已登记；加热或燃烧可产生有毒烟气，闪点因来源/方法为86–91℃，输出应保留来源和方法。", "ICSC-0513"),
            "health": ("吸入、皮肤接触或吞咽均需控制；可引起眼睛、呼吸道及皮肤刺激；存在 Repr. 1B/H360D 追溯分类证据。", "ICSC-0513"),
            "environment": ("与水完全混溶；不预期明显生物富集；环境归趋和快速生物降解为国际参考事实。", "CICAD-35"),
            "other": ("可增强其他物质经皮渗透；需结合具体配方和暴露场景人工复核。", "ICSC-0513"),
        }
        for domain, (summary, source_key) in summaries.items():
            connection.execute(
                "UPDATE cas_s2_hazard_summary SET summary_text = ?, value_status = 'derived_candidate', condition_id = ?, source_id = ?, raw_text = ? WHERE profile_id = ? AND hazard_domain = ?",
                (summary, intrinsic_id, source_ids[source_key], summary, profile_id, domain),
            )
        connection.execute(
            "UPDATE cas_s2_review SET review_status = 'manual_review', reviewed_at = '', review_note = ? WHERE profile_id = ? AND review_scope = 'profile'",
            ("NMP Section 2 标准结果档案。", profile_id),
        )
        refresh_standard_result_fields(connection, profile_id)

        return {
            "profile_id": profile_id,
            "cas_no": expected["identity"]["cas_no"],
            "source_count": connection.execute("SELECT COUNT(*) FROM cas_s2_source WHERE profile_id = ?", (profile_id,)).fetchone()[0],
            "local_source_file_count": connection.execute("SELECT COUNT(DISTINCT source_file_path || '|' || source_hash_sha256) FROM cas_s2_source WHERE profile_id = ? AND source_file_path <> '' AND source_file_path NOT LIKE 'remote://%'", (profile_id,)).fetchone()[0],
            "condition_count": connection.execute("SELECT COUNT(*) FROM cas_s2_condition WHERE profile_id = ?", (profile_id,)).fetchone()[0],
            "fact_count": connection.execute("SELECT COUNT(*) FROM cas_s2_fact WHERE profile_id = ?", (profile_id,)).fetchone()[0],
            "classification_count": connection.execute("SELECT COUNT(*) FROM cas_s2_classification WHERE profile_id = ?", (profile_id,)).fetchone()[0],
            "label_element_count": connection.execute("SELECT COUNT(*) FROM cas_s2_label_element WHERE profile_id = ?", (profile_id,)).fetchone()[0],
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="NMP CAS Section 2 法规事实和阈值补充工具")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    result = enrich_nmp(args.db, args.template)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"command": "enrich-nmp", "passed": True, **result}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
