# -*- coding: utf-8 -*-
"""多语言/公司品牌产出支持。

本模块不把机器翻译结果伪装成正式合规译文：
1. 从项目已有中英 MSDS 样本建立只读翻译记忆与公司配置索引；
2. 保护 CAS、EC、UN、GHS、H/P 语句、数字和单位；
3. 先执行已批准术语和样本记忆替换；
4. 未解决的中文内容标记为待复核，交给后续本地模型或人工确认。
"""
from __future__ import annotations

import re
import os
import subprocess
from collections import defaultdict
from pathlib import Path

try:
    from docx import Document
except Exception:  # pragma: no cover - 运行环境缺少 python-docx 时仍可提供配置接口
    Document = None


SAMPLE_ROOT = Path(r"F:\冠志工作空间\产品\TDS MSDS\TDS MSDS")
SAMPLE_DOC_ROOT = SAMPLE_ROOT / "产品 TDS MSDS -- WORD版本 - 副本"

COMPANY_PROFILES = {
    "guanzhi": {
        "label": "冠志",
        "english_label": "Guanzhi",
        "company_name": "广州冠志新材料科技有限公司",
        "company_name_en": "Guangzhou Guanzhi New Material Technology Co., Ltd.",
        "supplier_name": "广州冠志新材料科技有限公司",
        "supplier_name_en": "Guangzhou Guanzhi New Material Technology Co., Ltd.",
        "supplier_address": "广州市萝岗区科学城掬泉路3号广州国际企业孵化器A区1106室",
        "supplier_address_en": "Room 1106, Block A of Guangzhou International Business Incubator, No. 3 Juquan Road, Huangpu District, Guangzhou City",
        "telephone": "86-20-82567990",
        "fax": "86-20-32214789",
    },
    "guocai": {
        "label": "国彩",
        "english_label": "Guocai",
        "company_name": "英德市国彩精细化工有限公司",
        "company_name_en": "Yingde Guocai Fine Chemical Co., Ltd",
        "supplier_name": "英德市国彩精细化工有限公司",
        "supplier_name_en": "Yingde Guocai Fine Chemical Co., Ltd",
        "supplier_address": "广东省英德市白沙镇太平村更古坑凯迪工业园区",
        "supplier_address_en": "Kaidi Industrial Park, genggukeng, Taiping Village, Baisha Town, Yingde City, Guangdong Province",
        "telephone": "86-763-2811205",
        "fax": "86-763-2811024",
    },
}

# 这些术语来自现有英文 MSDS 样本中反复出现的 Section 1~16 固定表达。
# 键按最长优先替换，避免“危害”先替换后破坏长术语。
APPROVED_TERMS = {
    "物料及供应商标识": "Identification",
    "物料安全数据表": "SAFETY DATA SHEET",
    "修订日期": "Revision date",
    "产品名称": "Product name",
    "中文名称": "Chinese name",
    "化学品分类": "Chemical category",
    "产品使用建议和使用限制": "Product use suggestions and restrictions",
    "供应商信息": "Supplier information",
    "供应商名称": "Name of supplier",
    "供应商地址": "Supplier address",
    "危险性概述": "Hazards Identification",
    "GHS危险性类别": "GHS hazard category",
    "标签要素": "Tag element",
    "危险象形标记": "Dangerous pictographic symbol",
    "警告词": "Warning word",
    "信号词": "Signal word",
    "危害性说明": "Hazard statement",
    "防范说明": "Precautionary statement",
    "预防措施": "Preventive measures",
    "事故响应": "Accident response",
    "安全存储": "Secure storage",
    "废弃处置": "Disposal of waste",
    "物理化学危险": "Physical and chemical hazards",
    "健康危害": "Health hazards",
    "环境危害": "Environmental hazards",
    "成分/组成信息": "Composition/Information on Ingredients",
    "成分/组成资料": "Composition/Information on Ingredients",
    "产品类型": "Product type",
    "急救措施": "First aid measures",
    "消防措施": "Fire-fighting measures",
    "泄漏应急处理": "Accidental release measures",
    "操作处置与储存": "Handling and storage",
    "接触控制/个体防护": "Exposure controls/personal protection",
    "物理和化学性质": "Physical and chemical properties",
    "稳定性和反应性": "Stability and reactivity",
    "毒理学信息": "Toxicological information",
    "生态学信息": "Ecological information",
    "运输信息": "Transport information",
    "法规信息": "Regulatory information",
    "其他信息": "Other information",
    "其它的规定": "Other provisions",
    "符合下列法规要求": "Complies with the following regulatory requirements",
    "危险化学品安全管理条例，国务院令591号": "Regulations on the Safety Management of Hazardous Chemicals, State Council Decree No. 591",
    "GB/T 16483 化学品安全技术说明书内容和项目顺序": "GB/T 16483 Safety data sheet for chemicals — Content and order of sections",
    "GB 13690 化学品分类和危险性公示通则": "GB 13690 General rules for classification and hazard communication of chemicals",
    "GB 30000.2-29 化学品分类和标签规范": "GB 30000.2-29 Rules for classification and labelling of chemicals",
    "GB 15258 化学品安全标签编写规定": "GB 15258 Rules for the preparation of chemical safety labels",
    "该产品无可用的毒理学研究。": "No toxicological studies are available for this product.",
    "该产品无可用的生态毒理学研究。": "No ecotoxicological studies are available for this product.",
    "无数据": "No data available",
    "无": "None",
    "警告": "Warning",
    "易燃": "Flammable",
    "氧化剂": "Oxidizing",
    "加压气体": "Gases under pressure",
    "腐蚀": "Corrosive",
    "急性毒性": "Acute toxicity",
    "感叹号": "Exclamation mark",
    "健康危害": "Health hazard",
    "环境危害": "Environmental hazard",
}

_PROTECTED_RE = re.compile(
    r"(?:CAS\s*(?:No\.?|号)?\s*[:：]?\s*)?\d{2,7}-\d{2,7}-\d|"
    r"\b(?:GHS0[1-9]|H\d{3}|P\d{3}(?:\+P\d{3})?|UN\s*\d{4}|EC\s*\d+)\b|"
    r"\b\d+(?:[.,]\d+)?(?:\s*[~～–-]\s*\d+(?:[.,]\d+)?)?\s*(?:%|°C|℃|kPa|MPa|Pa|mg/L|g/L|ppm|mmHg)\b",
    re.IGNORECASE,
)
_PRODUCT_RE = re.compile(r"\b[A-Z]{1,8}-\d+[A-Z0-9+]*\b", re.IGNORECASE)
_COMPANY_RE = re.compile(r"(?i)(?:msds|sds).*?(?:cn|en|中文|英文).*?(?:guanzhi|guocai|冠志|国彩)")


def _clean(text: object) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _language(name: str) -> str | None:
    n = name.lower()
    if re.search(r"(?:^|[_\s-])(?:en|english)(?:[_\s.-]|$)", n) or "英文" in n:
        return "en"
    if re.search(r"(?:^|[_\s-])(?:cn|chinese)(?:[_\s.-]|$)", n) or "中文" in n:
        return "zh"
    return None


def _company(name: str) -> str | None:
    n = name.lower()
    if "guanzhi" in n or "冠志" in n:
        return "guanzhi"
    if "guocai" in n or "国彩" in n:
        return "guocai"
    return None


def _model_key(name: str) -> str:
    match = _PRODUCT_RE.search(Path(name).stem)
    return match.group(0).upper() if match else Path(name).stem.upper()


def _doc_table_cells(path: Path) -> list[tuple[int, int, int, str]]:
    if Document is None or path.suffix.lower() != ".docx":
        return []
    doc = Document(str(path))
    chunks = []
    for table_index, table in enumerate(doc.tables):
        for row_index, row in enumerate(table.rows):
            for cell_index, cell in enumerate(row.cells):
                text = _clean(cell.text)
                if text:
                    chunks.append((table_index, row_index, cell_index, text))
    return chunks


def sample_index(root: Path = SAMPLE_DOC_ROOT) -> dict[str, dict[str, list[Path]]]:
    """建立 {型号: {语言/公司: [样本]}} 索引；仅读取文件名。"""
    result: dict[str, dict[str, list[Path]]] = defaultdict(lambda: defaultdict(list))
    if not root.is_dir():
        return result
    for path in root.rglob("*.docx"):
        if path.name.startswith("~$") or "msds" not in path.name.lower():
            continue
        lang, company = _language(path.name), _company(path.name)
        if lang and company:
            result[_model_key(path.name)][f"{lang}_{company}"].append(path)
    return result


def _pair_memory(zh: list[tuple[int, int, int, str]], en: list[tuple[int, int, int, str]]) -> dict[str, str]:
    memory = {}
    english_by_position = {(table, row, cell): text for table, row, cell, text in en}
    for table, row, cell, source in zh:
        # 位置对齐只用于较长的正文块；短标签/地址在中英样本中经常因
        # “中文名称”等字段增删发生位移，禁止将它们误配成别的字段。
        if len(source) < 40 or not re.search(r"[\u3400-\u9fff]", source):
            continue
        target = english_by_position.get((table, row, cell), "")
        if len(target) >= 40 and not re.search(r"[\u3400-\u9fff]", target):
            memory[source] = target
    return memory


def build_memory(model: str = "", company: str = "guanzhi") -> dict[str, object]:
    """从已有中英文成对样本构建本次翻译预览所需的记忆。"""
    index = sample_index()
    selected = []
    for key, variants in index.items():
        if model and key.upper() != model.upper():
            continue
        zh = variants.get(f"zh_{company}", [])
        en = variants.get(f"en_{company}", [])
        if zh and en:
            selected.append((key, zh[0], en[0]))
    pairs = {}
    sources = []
    for key, zh_path, en_path in selected:
        mapping = _pair_memory(_doc_table_cells(zh_path), _doc_table_cells(en_path))
        pairs[key] = {"source": str(zh_path), "target": str(en_path), "segments": len(mapping)}
        sources.append(mapping)
    merged = {}
    for item in sources:
        merged.update(item)
    return {"model": model or "*", "company": company, "sample_count": len(selected), "pairs": pairs, "memory": merged}


def _protect(text: str):
    saved = {}
    def repl(match):
        token = f"__MSDS_TOKEN_{len(saved)}__"
        saved[token] = match.group(0)
        return token
    return _PROTECTED_RE.sub(repl, text), saved


def _restore(text: str, saved: dict[str, str]) -> str:
    for token, value in saved.items():
        text = text.replace(token, value)
    return text


def translate_text(text: object, target_language: str = "zh", memory: dict[str, str] | None = None) -> dict[str, object]:
    source = str(text or "")
    if target_language != "en" or not source:
        return {"text": source, "status": "source", "unresolved": []}
    protected, saved = _protect(source)
    translated = (memory or {}).get(source, protected)
    if translated == protected:
        for term in sorted(APPROVED_TERMS, key=len, reverse=True):
            translated = translated.replace(term, APPROVED_TERMS[term])
    translated = _restore(translated, saved)
    unresolved = sorted(set(re.findall(r"[\u3400-\u9fff]+", translated)))
    status = "approved_memory" if source in (memory or {}) else ("terminology_only" if not unresolved else "needs_review")
    return {"text": translated, "status": status, "unresolved": unresolved}


def translate_values(values: dict[str, object], target_language: str = "zh", company: str = "guanzhi", model: str = "") -> dict[str, object]:
    if not model:
        joined = " ".join(str(value or "") for value in (values or {}).values())
        product = _PRODUCT_RE.search(joined)
        model = product.group(0).upper() if product else ""
    memory_result = build_memory(model, company)
    memory = memory_result["memory"]
    translated = {}
    review = []
    for key, value in (values or {}).items():
        result = translate_text(value, target_language, memory)
        translated[key] = result["text"]
        if result["status"] == "needs_review":
            review.append({"key": key, "source": str(value or ""), "unresolved": result["unresolved"]})
    return {"values": translated, "review": review, "memory": {k: v for k, v in memory_result.items() if k != "memory"}, "source_language": "zh", "target_language": target_language, "company": company}


def company_profile(company: str, language: str = "zh") -> dict[str, str]:
    profile = dict(COMPANY_PROFILES.get(company, COMPANY_PROFILES["guanzhi"]))
    if language == "en":
        return {"company_name": profile["company_name_en"], "supplier_name": profile["supplier_name_en"], "supplier_address": profile["supplier_address_en"], "telephone": profile["telephone"], "fax": profile["fax"]}
    return {"company_name": profile["company_name"], "supplier_name": profile["supplier_name"], "supplier_address": profile["supplier_address"], "telephone": profile["telephone"], "fax": profile["fax"]}


def profile_options() -> list[dict[str, object]]:
    index = sample_index()
    return [{"id": key, "label": profile["label"], "english_label": profile["english_label"], "sample_models": sorted(k for k, v in index.items() if f"zh_{key}" in v and f"en_{key}" in v)} for key, profile in COMPANY_PROFILES.items()]


def convert_docx_to_pdf(source: Path, target: Path) -> Path:
    """Use the installed WPS CLI to convert the verified DOCX to PDF."""
    source = Path(source).resolve()
    target = Path(target).resolve()
    cli_root = Path(r"C:\Codex\plugins\cache\wps-cli-local\wps-cli\0.2.0\src")
    if not source.is_file():
        raise FileNotFoundError(f"待转换 Word 不存在：{source}")
    target.parent.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    old_pythonpath = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = str(cli_root) + (os.pathsep + old_pythonpath if old_pythonpath else "")
    cmd = ["python", "-m", "wps_cli.main", "export", "convert", str(source), "pdf", "--output", str(target), "--json"]
    completed = subprocess.run(cmd, env=env, capture_output=True, text=True, timeout=180)
    if completed.returncode != 0 or not target.is_file():
        detail = (completed.stderr or completed.stdout or "WPS 转换未生成文件").strip()
        raise RuntimeError(f"Word 转 PDF 失败：{detail[-1200:]}")
    return target


def apply_company_branding(path: Path, company: str = "guanzhi", language: str = "zh") -> dict[str, object]:
    """Replace the company identity in S0/header/footer without changing Sections 1-16."""
    if Document is None:
        raise RuntimeError("公司标识覆写需要 python-docx")
    path = Path(path)
    profile = COMPANY_PROFILES.get(company, COMPANY_PROFILES["guanzhi"])
    target = profile["company_name_en"] if language == "en" else profile["company_name"]
    sources = {
        profile["company_name"], profile["company_name_en"],
        COMPANY_PROFILES["guanzhi"]["company_name"], COMPANY_PROFILES["guanzhi"]["company_name_en"],
        COMPANY_PROFILES["guocai"]["company_name"], COMPANY_PROFILES["guocai"]["company_name_en"],
    }
    sources.discard(target)
    doc = Document(str(path))
    replacements = 0

    def replace_paragraph(paragraph):
        nonlocal replacements
        original = paragraph.text
        if not original:
            return
        for source in sorted(sources, key=len, reverse=True):
            if source in original:
                changed = False
                for run in paragraph.runs:
                    if source in run.text:
                        run.text = run.text.replace(source, target)
                        replacements += 1
                        changed = True
                if not changed:
                    paragraph.text = original.replace(source, target)
                    replacements += 1
                original = paragraph.text

    def visit_story(story):
        for paragraph in story.paragraphs:
            replace_paragraph(paragraph)
        for table in story.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_paragraph(paragraph)

    visit_story(doc)
    for section in doc.sections:
        for story in (section.header, section.first_page_header, section.even_page_header,
                      section.footer, section.first_page_footer, section.even_page_footer):
            visit_story(story)
    doc.save(str(path))
    return {"target": target, "replacements": replacements}


def translate_docx_in_place(path: Path, language: str, company: str = "guanzhi", model: str = "") -> dict[str, object]:
    """Translate an already verified DOCX while retaining run formatting where possible."""
    if language != "en":
        return {"review": [], "segments": 0}
    if Document is None:
        raise RuntimeError("英文产出需要 python-docx")
    if not model:
        probe = Document(str(path))
        probe_text = " ".join([p.text for p in probe.paragraphs[:20]])
        probe_text += " " + " ".join(
            cell.text for table in probe.tables[:4] for row in table.rows for cell in row.cells
        )
        product = _PRODUCT_RE.search(probe_text)
        model = product.group(0).upper() if product else ""
    memory_result = build_memory(model, company)
    memory = memory_result["memory"]
    doc = Document(str(path))
    review = []
    translated_segments = 0

    def visit(paragraph):
        nonlocal translated_segments
        original = _clean(paragraph.text)
        if not original:
            return
        result = translate_text(original, "en", memory)
        if result["text"] == original:
            # 对分 run 的标签执行术语级翻译，避免覆盖模板的字号和字体。
            for run in paragraph.runs:
                converted = translate_text(run.text, "en", memory)["text"]
                if converted != run.text:
                    run.text = converted
                    translated_segments += 1
        else:
            if len(paragraph.runs) == 1:
                paragraph.runs[0].text = result["text"]
            else:
                # 多 run 正文保留首 run 的字符样式，清空其余文字，避免破坏段落结构。
                paragraph.runs[0].text = result["text"]
                for run in paragraph.runs[1:]:
                    run.text = ""
            translated_segments += 1
        if result["unresolved"]:
            review.append({"source": original, "unresolved": result["unresolved"]})

    for paragraph in doc.paragraphs:
        visit(paragraph)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    visit(paragraph)
    for section in doc.sections:
        for story in (section.header, section.first_page_header, section.even_page_header,
                      section.footer, section.first_page_footer, section.even_page_footer):
            for paragraph in story.paragraphs:
                visit(paragraph)
            for table in story.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            visit(paragraph)
    doc.save(str(path))
    return {"review": review, "segments": translated_segments, "sample_count": memory_result["sample_count"]}
