# -*- coding: utf-8 -*-
"""核心读取器回归测试: 以 MSDS_CN 国彩 模板为标准."""
from __future__ import annotations

from pathlib import Path

from core.docx_reader import TEMPLATE_PATH, read_msds

# 默认模板 = 内化副本 (templates/), 字节级一致, 不依赖外部源路径
TEMPLATE = TEMPLATE_PATH
PU1034 = Path(r"F:\正式项目与模块化内容\MSDS 数据清理模块\标准模板\产品详情\TDS MSDS\产品 TDS MSDS -- WORD版本\1-1 单组份水性聚氨酯树脂 PU\PU-1034 msds_CN 国彩.docx")


def test_default_template_is_embedded():
    """默认模板应为程序内化副本 (templates/ 目录), 且可读."""
    from core.docx_reader import _TEMPLATE_RESOURCE
    assert str(_TEMPLATE_RESOURCE).endswith("templates\\MSDS_CN 国彩 模板.docx")
    assert _TEMPLATE_RESOURCE.exists(), "内化模板文件缺失"
    assert TEMPLATE == _TEMPLATE_RESOURCE, "TEMPLATE_PATH 应指向内化副本"
    r = read_msds(TEMPLATE)
    assert r.sections_count == 17
    assert "PEA-4139" in r.header


def test_template_reads_16_sections():
    r = read_msds(TEMPLATE)
    assert r.sections_count == 17, f"应读出0页眉页脚+16节, 实际 {r.sections_count}"
    assert r.tables_count == 16
    for n in range(0, 17):
        assert r.has_section(n), f"缺失第{n}节"
    assert "物料安全数据表" in r.header


def test_section0_page_fields():
    """页眉/页脚字段化纳入 section 0, 父子级: 0→页眉/页脚(sub)→字段."""
    r = read_msds(TEMPLATE)
    s0 = r.section(0)
    assert s0 is not None
    fm = s0.field_map()
    # 父子级: 页眉/页脚为 sub 子标题 (0.1 页眉 / 0.2 页脚)
    subs = [row for row in s0.iter_rows() if row.kind == "sub"]
    assert any("页眉" in (x.label or "") for x in subs), f"页眉应为子标题: {subs}"
    assert any("页脚" in (x.label or "") for x in subs), f"页脚应为子标题: {subs}"
    # 页眉表产品名 + 页脚公司/修订日期
    assert any("产品名称" in k for k in fm), f"页眉应含产品名字段: {list(fm)}"
    assert any("公司" in k for k in fm), f"页脚应含公司字段: {list(fm)}"
    assert any("页码" in k for k in fm), f"页脚应含页码字段: {list(fm)}"
    # 不可编辑默认: 页码/版本/公司
    for f in s0.fields:
        if "版本" in f.label or "Version" in f.label:
            assert f.editable is False
    # 页眉产品名默认可编辑 (随产品变)
    prod = [f for f in s0.fields if "产品名称" in f.label]
    assert prod and prod[0].editable is True
    assert any("修订" in f.label for f in s0.fields), "页脚应含修订日期字段"


def test_lines_split_into_three_columns():
    """序号|标题|内容 分拆为三列 (iter_rows)."""
    from core.structure import split_line, split_seq
    # split_seq: 序号/标题 拆列
    assert split_seq("8.1 暴露控制") == ("8.1", "暴露控制")
    assert split_seq("9.23 其他信息") == ("9.23", "其他信息")
    assert split_seq("1.1产品名称") == ("1.1", "产品名称")
    assert split_seq("手部防护") == ("", "手部防护")
    assert split_seq("3 / 5") == ("", "3 / 5"), "页码不应被拆"
    assert split_seq("GB/T 16483") == ("", "GB/T 16483"), "法规号不应被拆"
    # split_line 兼容入口: label 已去序号
    assert split_line("8.1 暴露控制") == ("sub", "暴露控制", "")
    assert split_line("氟化橡胶 –FKM:厚度≧0.4mm；穿透时间≧480min.")[1] == "氟化橡胶 –FKM"
    assert split_line("· 该产品无可用的毒理学研究。") == ("note", "", "· 该产品无可用的毒理学研究。")
    # 模板 S8: 材质行拆为 序号|标题|内容 三列
    r = read_msds(TEMPLATE)
    s8 = r.section(8)
    rows = [x for x in s8.iter_rows() if x.kind == "field" and x.label.startswith("氟化橡胶")]
    assert rows, "S8 氟化橡胶行应拆为 field 三列"
    assert rows[0].value, f"氟化橡胶行应有内容, 实际: {rows[0]!r}"
    # S9 字段: 序号列应有值
    s9 = r.section(9)
    rows9 = [x for x in s9.iter_rows() if x.kind == "field" and x.label == "外观"]
    assert rows9 and rows9[0].seq == "9.1", f"S9 外观 序号应拆为 9.1: {rows9}"


def test_multiline_headers_pair_with_content():
    """'加粗标题：' 与下一行内容 配对成 标题|内容 双列 (标题列受保护)."""
    # S2 多行通栏 (PU-1034): GHS分类 → 根据GHS不属于危险物
    if PU1034.exists():
        r = read_msds(PU1034)
        s2 = r.section(2)
        assert s2 is not None
        fields = {x.label: x.value for x in s2.iter_rows() if x.kind == "field"}
        assert fields.get("GHS分类") == "根据GHS不属于危险物", f"S2 拆分失败: {fields}"
        assert fields.get("GHS象形图") == "根据GHS不属于危险物"
    # S15 跨行标题 (模板): 符合下列法规要求 → 法规条目并入内容列
    t = read_msds(TEMPLATE)
    s15 = t.section(15)
    assert s15 is not None
    f15 = [x for x in s15.iter_rows() if x.kind == "field" and x.label == "符合下列法规要求"]
    assert f15, "S15 符合下列法规要求 应为 标题列"
    assert "GB/T 16483" in f15[0].value and "GB 15258" in f15[0].value, \
        f"法规条目应并入内容列: {f15[0].value!r}"


def test_s8_order_matches_doc():
    """S8 顺序与 Word 原始行序一致: 8.1 子标题在前, 材质行在建议之前."""
    r = read_msds(TEMPLATE)
    s8 = r.section(8)
    rows = [x for x in s8.iter_rows() if x.kind != "section"]
    labels = [x.label for x in rows]
    assert labels.index("暴露控制") < labels.index("呼吸系统防护"), \
        f"8.1 子标题应在呼吸系统防护之前: {labels}"
    assert labels.index("氟化橡胶 –FKM") < labels.index("建议"), \
        f"材质行应在建议之前: {labels}"
    assert labels.index("建议") < labels.index("眼睛防护") < labels.index("身体防护"), \
        f"建议/眼睛/身体防护顺序应保持: {labels}"
    assert labels[-1] == "身体防护", f"S8 末行应为身体防护: {labels}"


def test_s3_component_table_cells():
    """S3 成分表渲染为 cell 结构: 表头4cell + 每成分行4cell (含列间内框)."""
    r = read_msds(TEMPLATE)
    s3 = r.section(3)
    assert s3 is not None and s3.is_component_table
    # 成分行按序: 丙烯酸共聚物 在 去离子水 前 (与文档一致)
    names = [c.name for c in s3.components]
    assert names == ["丙烯酸共聚物", "去离子水", "二乙二醇单丁醚"], names
    # 产品类型字段在 order 中为 field
    assert "field" in s3.order


def test_extract_hierarchy():
    """分层提取: section → 大标题 → 小标题 → 字段+内容."""
    from core.extract import extract_doc, get_field, search_fields, render_text
    r = read_msds(TEMPLATE)
    entries = extract_doc(r)
    assert entries, "应提取出分层条目"
    # 大标题正确 (S1 节标题)
    s1 = [e for e in entries if e.section == 1]
    assert s1 and s1[0].big_title == "1.物料及供应商标识"
    # S8 小标题上下文: 呼吸系统防护 归属 8.1 暴露控制
    s8 = [e for e in entries if e.section == 8]
    breath = [e for e in s8 if e.label == "呼吸系统防护"]
    assert breath and breath[0].sub_title == "8.1 暴露控制", f"呼吸系统防护应归属 8.1: {breath[0]!r}"
    # S3 成分
    s3 = [e for e in entries if e.section == 3 and e.kind == "component"]
    assert any("丙烯酸共聚物" in e.label for e in s3), "S3 成分应提取"
    # 精确定位 API
    f = get_field(entries, 1, "中文名称")
    assert f and "PEA-4139" in f.value
    f2 = get_field(entries, 9, "密度")
    assert f2 and f2.value
    # 检索 API
    hit = search_fields(entries, "供应商", scope="label")
    assert any(e.label == "供应商名称" for e in hit)
    assert all("供应商" in e.full_label() for e in hit), "label 检索应只匹配标签"
    # 文本渲染含分层结构
    text = render_text(entries)
    assert "[1] 1.物料及供应商标识" in text
    assert "─ 8.1 暴露控制" in text


def test_extract_many_batch():
    """批量提取: 多文件统一按节/关键字过滤."""
    from core.extract import extract_many
    if not PU1034.exists():
        return
    many = extract_many([TEMPLATE, PU1034], sections={1}, query="供应商名称", scope="label")
    assert set(many) == {TEMPLATE.name, PU1034.name}
    for fn, es in many.items():
        assert es and all(e.section == 1 for e in es)
        assert all(e.label == "供应商名称" for e in es), f"{fn}: {[(e.label) for e in es]}"


def test_title_column_protected():
    """序号子标题/标题列 默认不可覆写; 内容列可编辑."""
    t = read_msds(TEMPLATE)
    s8 = t.section(8)
    assert s8 is not None
    for x in s8.iter_rows():
        if x.kind == "sub":
            assert x.editable is False, f"子标题应锁定不可覆写: {x.label}"
    # 材质行: 标题(材质)列可编辑为内容列, 材质名是标题
    rubber = [x for x in s8.iter_rows() if x.kind == "field" and x.label.startswith("氟化橡胶")]
    assert rubber and rubber[0].editable is True and rubber[0].value, \
        f"材质行应拆为 标题|厚度 两列: {rubber}"


def test_editable_defaults():
    """默认字段权限: 节标题/子标题不可编辑, 字段/内容可编辑."""
    r = read_msds(TEMPLATE)
    s8 = r.section(8)
    rows = s8.iter_rows()
    sec_row = [x for x in rows if x.kind == "section"]
    assert sec_row and sec_row[0].editable is False
    sub_rows = [x for x in rows if x.kind == "sub"]
    for sr in sub_rows:
        assert sr.editable is False, f"子标题应不可编辑: {sr.label}"
    field_rows = [x for x in rows if x.kind == "field"]
    assert field_rows and all(x.editable for x in field_rows), "字段应默认可编辑"


def test_s1_fields():
    r = read_msds(TEMPLATE)
    s1 = r.section(1)
    assert s1 is not None
    fm = s1.field_map()
    assert "中文名称" in fm
    assert "供应商名称" in fm
    assert fm["供应商名称"] == "英德市国彩精细化工有限公司"
    assert "供应商地址" in fm


def test_header_footer_includes_tables():
    """页眉/页脚含表格: 页眉表存产品名, 页脚表存公司/修订日期/页号."""
    r = read_msds(TEMPLATE)
    assert "物料安全数据表" in r.header
    assert "PEA-4139" in r.header, f"页眉应含页眉表产品名, 实际: {r.header!r}"
    assert "英德市国彩精细化工有限公司" in r.footer, f"页脚应含公司名, 实际: {r.footer!r}"
    assert "修订日期" in r.footer
    assert "3 / 5" in r.footer, f"页脚应含页码, 实际: {r.footer!r}"


def test_s8_cross_column_tab_split():
    """S8 一行跨多列: 标签格含制表符时, tab 前为标签、tab 后内容并入值."""
    r = read_msds(TEMPLATE)
    s8 = r.section(8)
    assert s8 is not None
    fm = s8.field_map()
    assert "手部防护" in fm, f"S8 手部防护标签应正确拆分, 实际字段: {list(fm)}"
    assert "喷涂过程中要求有呼吸防护设备。" in fm["手部防护"]
    assert "建议戴上防护手套。" in fm["手部防护"]


def test_s3_components():
    r = read_msds(TEMPLATE)
    s3 = r.section(3)
    assert s3 is not None
    assert s3.is_component_table
    assert len(s3.components) >= 3
    first = s3.components[0]
    assert first.name and first.cas and first.conc
    assert "3.1 产品类型" in s3.field_map()


def test_s9_fields_and_numbering_anomaly():
    r = read_msds(TEMPLATE)
    s9 = r.section(9)
    assert s9 is not None
    fm = s9.field_map()
    assert "9.1 外观" in fm
    assert "9.23 其他信息" in fm
    # 模板 S9 编号应连续 (9.1..9.23)
    nums = [int(f.label.split(".")[1].split()[0]) for f in s9.fields if f.label.startswith("9.")]
    assert nums == list(range(1, 24)), f"S9 编号不连续: {nums}"


def test_s11_has_component_reference():
    r = read_msds(TEMPLATE)
    s11 = r.section(11)
    assert s11 is not None
    text = "\n".join(s11.lines)
    assert "二乙二醇单丁醚" in text or any("二乙二醇" in f.value for f in s11.fields)


def test_real_pu1034_inline_section():
    """真实文件: S3 表内嵌 S4 节标题, 读取器应正确切分."""
    if not PU1034.exists():
        return  # 源文件缺失时跳过
    r = read_msds(PU1034)
    s3 = r.section(3)
    assert s3 is not None and s3.is_component_table
    # 成分: 前 3 个应为 聚氨酯聚合物/去离子水/三乙胺
    names = [c.name for c in s3.components]
    assert "聚氨酯聚合物" in names and "去离子水" in names and "三乙胺" in names
    assert len(s3.components) == 3, f"S3 成分数异常: {names}"
    # S4 被正确切出 (内嵌节)
    s4 = r.section(4)
    assert s4 is not None, "S4 应被识别"
    assert any("误服" in f.label for f in s4.fields), "S4 字段应归属 S4"
    # S4 内容不应混入 S3
    assert not any("误服" in c.name for c in s3.components)
    # 无第4节缺失异常
    assert not any(a.message == "缺失第4节" for a in r.anomalies)


# ============================================================
# S3 成分兼容性: 一行多成分拆分 + 全角符号归一化
# ============================================================

def test_component_split_multi_line():
    """一行多成分 (三列各含 \\n) 拆分为多个 ComponentData, 按行对齐."""
    from core.structure import split_component_cells
    comps = split_component_cells(
        "封闭型脂肪族聚异氰酸酯\n二丙二醇二甲醚\n水\nN,N-二甲基乙醇胺",
        "商业机密\n872-50-4\n7732-18-5\n108-01-0",
        "45-55\n5-10\n35-45\n0.5-2",
    )
    assert len(comps) == 4, f"应拆 4 个成分: {[c.name for c in comps]}"
    assert comps[0].name == "封闭型脂肪族聚异氰酸酯" and comps[0].cas == "商业机密"
    assert comps[1].name == "二丙二醇二甲醚" and comps[1].cas == "872-50-4"
    assert comps[2].name == "水" and comps[2].cas == "7732-18-5"
    assert comps[3].name == "N,N-二甲基乙醇胺" and comps[3].cas == "108-01-0"
    assert [c.conc for c in comps] == ["45-55", "5-10", "35-45", "0.5-2"]


def test_component_fullwidth_normalize():
    """全角符号 ＞＜～％－ 归一为半角 (含量/名称)."""
    from core.structure import normalize_component_conc, normalize_component_name
    assert normalize_component_conc("＞45") == ">45"
    assert normalize_component_conc("＜1") == "<1"
    assert normalize_component_conc("30.0～40.0") == "30.0~40.0"
    assert normalize_component_conc("> 30") == ">30"
    assert normalize_component_conc("95±1") == "95±1"      # ± 保留
    assert normalize_component_conc("无数据") == "无数据"     # 文字保留
    assert normalize_component_name("N,N－二甲基乙酰胺(DMAC)") == "N,N-二甲基乙酰胺(DMAC)"
    assert normalize_component_name("2-正辛基-4异噻唑啉--3酮") == "2-正辛基-4异噻唑啉-3酮"
    assert normalize_component_name("N,N二甲基乙醇胺") == "N,N二甲基乙醇胺"  # 缺连字符保留


def test_component_cas_placeholder_kept():
    """占位 CAS (商业机密/待确认) 保留语义; 纯符号 CAS 置空."""
    from core.structure import normalize_component_cas
    assert normalize_component_cas("商业机密") == "商业机密"
    assert normalize_component_cas("待确认") == "待确认"
    assert normalize_component_cas("无单一") == "无单一"
    assert normalize_component_cas("——") == ""       # 格式错误 → 空
    assert normalize_component_cas("－") == ""
    assert normalize_component_cas(" 7732-18-5 ") == "7732-18-5"  # 去空白保留
    assert normalize_component_cas("55965-84-9") == "55965-84-9"


def test_s3_real_multi_component_file():
    """真实文件: '丙烯酸共聚物\\n去离子水' 应拆成两个成分, 且记录拆分异常."""
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\冠志 guanzhi\PA-3615 MSDS（冠志）.docx")
    if not p.exists():
        return  # 文件缺失时跳过 (环境无关)
    r = read_msds(p)
    s3 = r.section(3)
    assert s3 is not None and s3.is_component_table
    names = [c.name for c in s3.components]
    assert "丙烯酸共聚物" in names, f"拆分后应含丙烯酸共聚物: {names}"
    assert "去离子水" in names, f"拆分后应含去离子水: {names}"
    # 无换行符残留
    assert all("\n" not in c.name for c in s3.components)
    assert any("成分拆分" in a.message for a in r.anomalies), "应记录成分拆分异常"


# ============================================================
# 覆写基座: 格式保留覆写 (内化模板为默认覆写底本)
# ============================================================

# ============================================================
# 覆写基座: 格式保留覆写 (内化模板为默认覆写底本)
# ============================================================

def test_section_title_letter_prefix():
    """节标题兼容: v1./l1. 前缀 (冠志/手误) 剥除后识别; 英文 SECTION 长标题可识别."""
    from core.detectors import is_section_title
    # v1. 前缀 (全库 233 文件 S1)
    assert is_section_title("v1.物料及供应商标识") == (1, "物料及供应商标识")
    # l1. 前缀 (定稿模板手误)
    assert is_section_title("l1.物料及供应商标识") == (1, "物料及供应商标识")
    # 英文 SECTION 长标题 (>60 字符)
    t = "SECTION 1: Identification of the substance/mixture and of the company/undertaking"
    assert len(t) > 60
    assert is_section_title(t) == (1, "Identification of the substance/mixture and of the company/undertaking")
    # 非节标题不受影响
    assert is_section_title("pH值（1%水溶液）") == (None, None)
    assert is_section_title("Version：V1.0") == (None, None)
    assert is_section_title("8.1 暴露控制") == (None, None)


def test_section1_letter_prefix_real_file():
    """真实文件: v1. 前缀的 S1 应被识别, 不缺失第1节 (BL-8085 国彩)."""
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\BL-8085 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    assert r.has_section(1), f"BL-8085 不应缺失第1节: {sorted(r.sections)}"
    assert not any("缺失第1节" in a.message for a in r.anomalies)



if __name__ == "__main__":
    import sys
    r = read_msds(TEMPLATE)
    print("解析摘要:", r.summary())
    print("异常:", [(a.level, a.section, a.message) for a in r.anomalies])
    for n in (1, 3, 9):
        sec = r.section(n)
        print(f"\n=== S{n} {sec.title} ===")
        if sec.is_component_table:
            for c in sec.components:
                print(f"  成分: {c.name} | {c.cas} | {c.conc}")
        for f in sec.fields:
            print(f"  {f.label}: {f.value[:40]}")


# ============================================================
# 逐文件真实导入实测的修复: 空值告警消除 + S9 一行/行内多编号拆分
# ============================================================

def test_no_empty_value_warning():
    """模板字段留空是常态 (供应商信息:/产品名称: 子标题行), 不再产生"字段值为空"告警.

    全库实测: 646 文件原 590 个触发空值 warn 弹窗, 修复后 0.
    """
    from core.docx_reader import TEMPLATE_PATH, read_msds
    r = read_msds(TEMPLATE_PATH)
    assert not any("字段值为空" in a.message for a in r.anomalies), \
        [a.message for a in r.anomalies]
    # 空值字段仍保留 (可编辑待填), 只是不告警
    s1 = r.section(1)
    fm = s1.field_map()
    assert "1.3供应商信息" in fm and not fm["1.3供应商信息"], "S1 供应商信息 模板留空"
    assert "1.1产品名称" in fm and not fm["1.1产品名称"], "S1 产品名称 模板留空"


def test_s9_multi_number_row_split():
    """S9 一行多编号 (标签格换行分隔, 值格多行对齐) 拆分为多个独立字段."""
    from core.docx_reader import _parse_field_row
    from core.structure import SectionData
    sec = SectionData(number=9, title="x", full_title="9. x")
    _parse_field_row(
        ["9.19 Dynamic viscosity:\n9.20 Explosion characteristics\n9.21 Dust explosion level",
         "<250 mPa·s\nNo data\nNot applicable"], sec)
    assert [f.label for f in sec.fields] == [
        "9.19 Dynamic viscosity", "9.20 Explosion characteristics", "9.21 Dust explosion level"]
    assert [f.value for f in sec.fields] == ["<250 mPa·s", "No data", "Not applicable"]


def test_s9_inline_multi_number_split():
    """S9 行内多编号 (两编号挤同一行无换行, 如 RA-15000) 断行后拆分, 值对齐不错位."""
    from core.docx_reader import _parse_field_row
    from core.structure import SectionData
    sec = SectionData(number=9, title="x", full_title="9. x")
    _parse_field_row(
        ["9.19 最低成膜温度MFFT/℃9.20 玻璃化温度Tg/℃：", "40\n45"], sec)
    assert [f.label for f in sec.fields] == ["9.19 最低成膜温度MFFT/℃", "9.20 玻璃化温度Tg/℃"]
    assert [f.value for f in sec.fields] == ["40", "45"]


def test_s9_composite_label_not_split():
    """复合标签 (无编号多行, 如 'Combustion value:\nSaturated vapor pressure:') 不误拆."""
    from core.docx_reader import _parse_field_row
    from core.structure import SectionData
    sec = SectionData(number=9, title="x", full_title="9. x")
    _parse_field_row(["Combustion value:\nSaturated vapor pressure:", "Not applicable\nNo data"], sec)
    assert len(sec.fields) == 1, f"复合标签应保持单字段: {[f.label for f in sec.fields]}"
    assert "\n" in sec.fields[0].label or "Saturated" in sec.fields[0].label


# ============================================================
# 无标题行兼容 (BL-8085 S5): 首格空、次列有内容 → 内容放文本列
# ============================================================

def test_no_title_row_blank_first_cell():
    """无标题行: 首格空、次列有内容 → 独立字段 (label 空, value 在内容列)."""
    from core.docx_reader import _parse_field_row
    from core.structure import SectionData
    sec = SectionData(number=5, title="消防措施", full_title="5.消防措施")
    _parse_field_row(["物质或混合物的特殊危害：", "燃烧时释放一氧化碳、二氧化碳、氮氧化物和少量的氰化氢"], sec)
    _parse_field_row(["", "在着火或爆炸情况下，不要吸进烟尘。"], sec)
    _parse_field_row(["消防预防措施和保护设备：", "消防人员必须佩戴自供气式呼吸器。"], sec)
    labels = [f.label for f in sec.fields]
    values = [f.value for f in sec.fields]
    assert labels == ["物质或混合物的特殊危害", "", "消防预防措施和保护设备"], labels
    assert values[1] == "在着火或爆炸情况下，不要吸进烟尘。", values
    assert sec.order == ["field", "field", "field"]


def test_no_title_row_continuous_merge():
    """连续空首格行 (PA-4408 S7 跨行续行) 合并为一个块, 不碎片化."""
    from core.docx_reader import _parse_field_row
    from core.structure import SectionData
    sec = SectionData(number=7, title="操作和储存", full_title="7.操作和储存")
    _parse_field_row(["7.1 安全操作防范：", "操作时遵守化学品的常见预防措施。"], sec)
    _parse_field_row(["", "避免与皮肤和眼睛接"], sec)
    _parse_field_row(["", "触。"], sec)
    _parse_field_row(["", "远离食物、饮料和烟草。"], sec)
    assert len(sec.fields) == 2, f"连续空首格应合并进首个无标题行: {[f.label for f in sec.fields]}"
    assert sec.fields[1].label == ""
    assert "避免与皮肤和眼睛接\n触。\n远离食物、饮料和烟草。" in sec.fields[1].value, sec.fields[1].value


def test_no_title_row_real_bl8085():
    """真实文件: BL-8085 S5 特殊危害后无标题行, 内容按原文位置放文本列."""
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\BL-8085 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    s5 = r.section(5)
    rows = [x for x in s5.iter_rows() if x.kind == "field"]
    nohead = [x for x in rows if not x.label]
    assert nohead, "BL-8085 S5 应有无标题行"
    assert nohead[0].value == "在着火或爆炸情况下，不要吸进烟尘。", nohead[0]
    assert nohead[0].editable is True, "无标题行内容列应默认可编辑 (可覆写)"


# ============================================================
# 成分表头可检索 (BL-8085 表头词 '化学品名称/CAS编号/含量' 检索不到)
# ============================================================

def test_component_header_searchable():
    """S3 成分表头 (化学品名称|CAS编号|含量) 作为可检索条目纳入 extract.

    BL-8085 实测: 修复前检索 '化学品名称'/'CAS编号'/'w/w' 0 命中.
    """
    from core.docx_reader import read_msds
    from core.extract import extract_doc, search_fields, get_field
    r = read_msds(TEMPLATE)
    entries = extract_doc(r)
    headers = [e for e in entries if e.kind == "component_header"]
    assert headers, "模板 S3 应有成分表头条目"
    assert "化学品名称" in headers[0].value and "CAS编号" in headers[0].value, headers[0].value
    # 表头词可检索
    for q in ("化学品名称", "CAS编号", "w/w"):
        hits = search_fields(entries, q, scope="all")
        assert hits, f"表头词 {q} 应可检索"
        assert any(e.kind == "component_header" for e in hits), f"{q} 应命中表头条目"
    # 表头不作为数据字段精确定位 (get_field 只返回 field/component)
    assert get_field(entries, 3, "成分表头") is None


def test_component_header_real_bl8085():
    """真实文件: BL-8085 表头词可检索 (修复前 化学品名称/CAS编号/%（w/w） 均 0 命中)."""
    from core.docx_reader import read_msds
    from core.extract import extract_doc, search_fields
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\BL-8085 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    entries = extract_doc(r)
    for q in ("化学品名称", "CAS编号"):
        assert search_fields(entries, q, scope="all"), f"BL-8085 表头词 {q} 应可检索"


# ============================================================
# 单列(整格)成分表解析 + 英文成分表头识别
#   全库 646 文件 token 级对照: 修复前 43 文件 S3 遗漏 token,
#   修复后 0 遗漏. 单列表格 (tab/空格分隔) + 三列英文表头均覆盖.
# ============================================================

def test_flat_component_tab_pa4757():
    """单列 tab 成分表: 'Ingredient\\tCAS\\tNo. %(w/w)' (PA-4757 格式).

    名称/CAS/含量挤一格 ('Acrylic polymer\t9003-01-4 40±1') 也要拆对.
    """
    from core.structure import split_flat_component_text
    text = ("3.1 Mixtures\nDescription: Polymer\n"
            "Dangerous components: Void\n3.2 Composition\n"
            "Ingredient\tCAS\tNo. %(w/w)\n"
            "Acrylic polymer\t9003-01-4 40±1\n"
            "Deionized water 7732-18-5\t60±1")
    flat = split_flat_component_text(text)
    assert flat is not None, "应识别出 tab 成分表"
    assert len(flat["rows"]) == 2, flat["rows"]
    assert flat["rows"][0] == ("Acrylic polymer", "9003-01-4", "40±1"), flat["rows"][0]
    assert flat["rows"][1] == ("Deionized water", "7732-18-5", "60±1"), flat["rows"][1]
    assert "3.1 Mixtures" in flat["pre_lines"][0]
    assert "Dangerous components: Void" in " ".join(flat["pre_lines"])


def test_flat_component_tab_hpu():
    """单列 tab 成分表: 列序 浓度|名称|CAS (HPU-7651 格式), 尾注不作成分."""
    from core.structure import split_flat_component_text
    text = ("Hazardous Components\nThere are no hazardous components above the "
            "relevant concentration limits.\nOTHER INGREDIENTS\n"
            "Concentration\t \t Components\t \t CAS-No.\n"
            "0.5 - 5%\t Dimethylethanolamine\t 108-01-0\n"
            "This product contains an amine neutralizing agent which is bound "
            "in the matrix of this product as a salt.")
    flat = split_flat_component_text(text)
    assert flat is not None
    assert len(flat["rows"]) == 1, flat["rows"]
    assert flat["rows"][0] == ("Dimethylethanolamine", "108-01-0", "0.5-5%"), flat["rows"][0]
    assert any("This product contains" in l for l in flat["post_lines"]), \
        "尾注应保留为说明行, 不作成分"


def test_flat_component_space_separated():
    """单列空格分隔成分表: 'Chemical Name CAS No. % (w/w)' (OS-8030 格式)."""
    from core.structure import split_flat_component_text
    text = ("3.1 Product Type: Mixture Mixture\n3.2 Ingredients\n"
            "Chemical Name CAS No. % (w/w)\n"
            "Aliphatic Polycarbodiimide Proprietary 35-45\n"
            "Water 7732-18-5 55-65")
    flat = split_flat_component_text(text)
    assert flat is not None
    assert len(flat["rows"]) == 2, flat["rows"]
    assert flat["rows"][0] == ("Aliphatic Polycarbodiimide", "Proprietary", "35-45")
    assert flat["rows"][1] == ("Water", "7732-18-5", "55-65")


def test_flat_component_s4_overflow_not_component():
    """单列 tab 中文成分表 + S4 溢出通栏 (PA-4408): 溢出文字带数字含量但无 CAS,
    不作成分, 保留为说明行."""
    from core.structure import split_flat_component_text
    text = ("产品类型：\t混合物\n成分\n"
            "化学品名称\tCAS 编号\t %（w/w）\n"
            "羟基丙烯酸酯聚合物\t商业机密\t＞40\n水\t7732-18-5\t＞40\n"
            "乙二醇丁醚\t111-76-2\t6.5\nN,N 二甲基乙醇胺\t108-01-0\t1-2\n"
            "N,N-二甲基乙醇胺，中和剂，已键合为盐，质量浓度小于 2.0%\n"
            "GHS 分类：依然液体 3 H226；急性毒性 4 吸入性 H332\n"
            "特定阈值浓度≥5%")
    flat = split_flat_component_text(text)
    assert flat is not None
    assert len(flat["rows"]) == 4, f"仅 4 个成分, S4 溢出不作成分: {flat['rows']}"
    assert flat["rows"][0] == ("羟基丙烯酸酯聚合物", "商业机密", "＞40")
    assert any("质量浓度小于 2.0%" in l for l in flat["post_lines"]), "S4 溢出应保留"
    assert any("GHS 分类" in l for l in flat["post_lines"])
    assert any("特定阈值浓度" in l for l in flat["post_lines"])


def test_component_header_row_english_and_guard():
    """英文三列表头识别 + 单行守卫.

    - 'Chemical Name | CAS Number | %（w/w）' 应识别 (PA-4771/BL-8146 格式)
    - 含换行的整格成分表大文本不得误判为表头 (PA-4757/HPU-7651 丢数据根因)
    """
    from core.detectors import is_component_header_row
    assert is_component_header_row(["Chemical Name", "CAS Number", "%（w/w）"]) is True
    assert is_component_header_row(["Chemical name", "CAS NO.", "%（w/w）"]) is True
    assert is_component_header_row(["Concentration", "Components", "CAS-No."]) is True
    # 单行守卫: 整格大文本 (含 \n) 不判表头
    big = ("Dangerous components: Void\n3.2 Composition\n"
           "Ingredient\tCAS\tNo. %(w/w)\nAcrylic polymer\t9003-01-4 40±1\n"
           "Deionized water 7732-18-5\t60±1")
    assert is_component_header_row([big]) is False, "整格成分表不得误判为表头"


def test_component_header_saved_actual():
    """识别到的表头保存到 SectionData.component_header, extract 检索用实际表头
    (英文表头 'Chemical Name | CAS Number | %（w/w）' 可检索到 'Number')."""
    from core.docx_reader import read_msds
    from core.extract import extract_doc, search_fields
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\英文\国彩 guocai\PA-4771 msds_EN Guocai.docx")
    if not p.exists():
        return
    r = read_msds(p)
    sec3 = r.section(3)
    assert sec3.component_header, "应保存实际表头"
    assert "Chemical Name" in sec3.component_header, sec3.component_header
    entries = extract_doc(r)
    headers = [e for e in entries if e.kind == "component_header"]
    assert headers and "Chemical Name" in headers[0].value, headers
    # 英文表头词可检索
    assert search_fields(entries, "Number", scope="all"), "英文表头 'Number' 应可检索"


def test_flat_component_real_pa4757():
    """真实文件: PA-4757 单列 tab 成分表 2 成分 (修复前 0 成分, 内容整格丢弃)."""
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\英文\国彩 guocai\PA-4757 msds_EN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    sec3 = r.section(3)
    assert sec3.is_component_table and len(sec3.components) == 2, sec3.components
    assert (sec3.components[0].name, sec3.components[0].cas, sec3.components[0].conc) \
        == ("Acrylic polymer", "9003-01-4", "40±1")
    assert sec3.components[1].name == "Deionized water"


def test_flat_component_real_pa4408():
    """真实文件: PA-4408 中文单列 tab 成分表 4 成分 + S4 溢出通栏保留."""
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\PA-4408 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    sec3 = r.section(3)
    assert sec3.is_component_table and len(sec3.components) == 4, sec3.components
    names = [c.name for c in sec3.components]
    assert names == ["羟基丙烯酸酯聚合物", "水", "乙二醇丁醚", "N,N 二甲基乙醇胺"], names
    # 产品类型 tab 两列 → 字段 (自动编号前缀 3.1 已恢复)
    assert sec3.field_map().get("3.1 产品类型") == "混合物"
    # S4 溢出通栏保留为说明行
    joined_lines = "\n".join(sec3.lines)
    assert "GHS 分类" in joined_lines and "特定阈值浓度" in joined_lines


# ============================================================
# GUI 恢复显示默认模板 (导入产品后一键切回内化模板)
# ============================================================

def test_gui_restore_default_template():
    """GUI『恢复默认模板』: 导入产品后显示源切回内化模板, 比对基准重置,
    字段权限标注清空. 无 tk 显示环境自动跳过."""
    try:
        import tkinter as tk
        import tkinter.messagebox as mb
        _ = tk.Tk()   # 探测显示环境
        _.destroy()
    except Exception:
        return
    from gui.main_window import MainWindow
    from core.docx_reader import read_msds, TEMPLATE_PATH
    orig_err = mb.showerror
    mb.showerror = lambda *a, **k: None   # 防止异常弹窗卡住测试
    try:
        w = MainWindow()
        w.withdraw()
        # 启动即加载默认模板, 显示源为 template
        assert w.template is not None
        assert w._display_source == "template"
        assert w._display_source_of() is w.template
        # 模拟导入产品 → 显示源切换为 product
        p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\英文\国彩 guocai\PA-4771 msds_EN Guocai.docx")
        if p.exists():
            w.product = read_msds(p)
            w._display_source = "product"
            assert w._display_source_of() is w.product
            # 恢复默认模板 → 显示源/比对基准切回内化模板, 标注清空
            w._restore_default_template()
            assert w._display_source == "template"
            assert w._display_source_of() is w.template
            assert w.template.file_name == TEMPLATE_PATH.name
            assert w._editable_overrides == {}
        w.destroy()
    finally:
        mb.showerror = orig_err


# ============================================================
# 文本格式段落式 MSDS 统一呈现 (BL-8085 S11 等单列段落文本)
#   段落式文本无 表格格式, 而是 段落+空格+分行标志 ("请参阅以下数据："、
#   "11.1 毒理学效应" 标题) → 必须以 Reader 标准三列表格统一呈现.
# ============================================================

def test_paragraph_heading_break_pending():
    """短标题行打断上一 pending (8085 S11: 急性毒性，经皮 → 吸入 独立标题)."""
    from core.structure import split_text_block
    rows = split_text_block(
        "急性毒性，经皮\n科学地研究，而不仅仅是合理地研究。\n急性毒性，吸入\n"
        "半数致死浓度（LC50） 大鼠: > 1.12 mg/l, 4 h\n试验环境: 粉尘/烟雾")
    labels = [r.label for r in rows]
    assert "急性毒性，经皮" in labels, labels
    assert "急性毒性，吸入" in labels, labels
    # 急性毒性，经皮 只带走它的内容, 不吞掉后面的标题
    ep = [r for r in rows if r.label == "急性毒性，经皮"][0]
    assert ep.value == "科学地研究，而不仅仅是合理地研究。", ep.value
    xs = [r for r in rows if r.label == "急性毒性，吸入"][0]
    assert xs.value == "", f"吸入标题应空值 (后续为冒号行): {xs.value!r}"


def test_paragraph_heading_after_previous():
    """下一行是行尾冒号标题 → 当前短行是内容 (PU-1034: 根据GHS不属于危险物)."""
    from core.structure import split_text_block
    rows = split_text_block("GHS分类：\n根据GHS不属于危险物\nGHS象形图：\n根据GHS不属于危险物")
    fields = {r.label: r.value for r in rows if r.kind == "field"}
    assert fields.get("GHS分类") == "根据GHS不属于危险物", fields
    assert fields.get("GHS象形图") == "根据GHS不属于危险物", fields


def test_paragraph_continuation_merge():
    """无冒号续行 (对产品的研究.) 并入上一 field 内容 (8085 S11 方法行)."""
    from core.structure import split_text_block
    rows = split_text_block("方法: OECD化学品测试指南423\n对产品的研究.")
    m = [r for r in rows if r.label == "方法"][0]
    assert m.value == "OECD化学品测试指南423\n对产品的研究.", m.value


def test_paragraph_real_bl8085_s11():
    """真实文件: BL-8085 S11 段落式文本 → 统一三列表格.

    标题 (急性毒性，经口/原发性皮肤刺激/致癌性/CMR评估) 独立行,
    内容 (无数据资料/科学地研究...) 正确配对, 序号子标题 11.1 独立.
    """
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\BL-8085 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    s11 = r.section(11)
    rows = [x for x in s11.iter_rows() if x.kind != "section"]
    labels = [x.label for x in rows]
    # 标题行识别
    for h in ("急性毒性，经口", "急性毒性，经皮", "原发性皮肤刺激",
              "原发性粘膜刺激", "致癌性", "亚急性，亚慢性和延迟毒性",
              "CMR 评估", "STOT 评估 – 一次性接触"):
        assert h in labels, f"标题应识别: {h}"
    # 序号子标题独立
    subs = [x for x in rows if x.kind == "sub"]
    assert any(x.seq == "11.1" and x.label == "毒理学效应" for x in subs), subs
    # 内容配对
    fm = {x.label: x.value for x in rows if x.kind == "field"}
    assert fm.get("致癌性") == "无数据资料", fm.get("致癌性")
    assert fm.get("亚急性，亚慢性和延迟毒性") == "无数据资料"
    # 续行并入 (方法: OECD... + 对产品的研究.)
    assert any(x.label == "方法" and "对产品的研究." in x.value for x in rows), \
        "对产品的研究. 应并入上一 方法 field"
    # 无内容丢失: 每行都有归属 (无 note 通栏堆积整段)
    notes = [x for x in rows if x.kind == "note"]
    assert not notes, f"S11 不应有通栏 note 堆积 (应结构化): {notes}"


# ============================================================
# 自动编号序号恢复 (Word <w:numPr> 列表生成序号 → 拼接为文本)
# ============================================================

def test_resolver_lvltext_literal():
    """自动编号解析核心: lvlText 字面量前缀/后缀保留 + 快照时序.

    - "9.%1" → "9.1" (字面量 '9.' 保留, 不能只留 %n 引用)
    - 先取计数器快照再递增 (否则 "9.%1" 会把 9.1 算成 2)
    - %n 引用第 n 级**计数器数字** (Word 标准): lvl1 '%1.%2' 在 lvl0=3
      下显示 "3.1", 不继承 lvl0 的字面量前缀
    - 更深级引用 lvl0 当前值; 升回浅级重置更深级
    """
    from core.docx_reader import _NumberingResolver
    r = _NumberingResolver.__new__(_NumberingResolver)
    r._numdefs = {"1": {0: (1, "9.%1"), 1: (1, "%1.%2")}}
    r._counters = {}
    # lvlText 字面量前缀保留 + 快照时序 (先取值后递增)
    assert r.resolve("1", 0) == "9.1"
    assert r.resolve("1", 0) == "9.2"
    # 更深级引用 lvl0 计数器 (=3), %2=本级 1 → 3.1 (不继承字面前缀)
    assert r.resolve("1", 1) == "3.1"
    assert r.resolve("1", 1) == "3.2"
    # 升回浅级 → lvl0 续 9.3, 重置更深级
    assert r.resolve("1", 0) == "9.3"
    assert r.resolve("1", 1) == "4.1"


def test_auto_numbering_seq_restored():
    """真实文件: BL-8085 S9 自动编号序号恢复.

    原文 S9 外观/嗅觉阀值/.../引燃温度 15 行由 Word 列表生成 (9.1~9.15),
    后接显式 9.17~9.19. 缺失 9.16 是文档真实情况 (原文无此行), 如实告警.
    """
    import re
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\BL-8085 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    s9 = r.section(9)
    labels = [f.label for f in s9.fields]
    # 自动编号 9.1~9.15
    assert "9.1 外观" in labels, labels[:6]
    assert "9.10 密度" in labels
    assert "9.15 引燃温度" in labels
    n9 = [int(m.group(1)) for f in s9.fields
          if (m := re.match(r"^9\.(\d+)", f.label))]
    assert n9[:15] == list(range(1, 16)), f"自动编号应连续 9.1..9.15: {n9}"
    # 显式 9.17~9.19 衔接
    assert "9.17动力粘度" in labels
    assert "9.19其他信息" in labels


# ============================================================
# 加粗 → 标题列归类 (无冒号英文标签放宽为字段)
# ============================================================

def test_bold_english_label_field():
    """加粗英文标签 (Waste Disposal Method) 归为字段标题列, 非通栏 note."""
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\英文\国彩 guocai\RU-10130 msds_EN Guocai.docx")
    if not p.exists():
        return
    r = read_msds(p)
    found = False
    for num, sec in r.sections.items():
        for row in sec.iter_rows():
            if row.label == "Waste Disposal Method":
                assert row.kind == "field", (p.name, num, row)
                assert row.value, f"Waste Disposal Method 应有值: {row}"
                found = True
    assert found, "应找到 Waste Disposal Method 字段"


def test_bold_heading_note_not_misclassified():
    """加粗内容句不被误判为标题: '根据EC指令...接触限值信息' 仍是内容."""
    from core.docx_reader import read_msds
    p = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS\中文\国彩 guocai\BL-8085 msds_CN 国彩.docx")
    if not p.exists():
        return
    r = read_msds(p)
    for num, sec in r.sections.items():
        for row in sec.iter_rows():
            if "接触限值信息" in row.label:
                # 含逗号+长数字的完整陈述句 → 不归为标题行
                assert row.kind == "note" or row.kind == "field", (num, row)


# ============================================================
# 英文占位符 CAS (Trade secret / Business secret) 保留语义
# ============================================================

def test_english_cas_placeholder_kept():
    """英文占位符 CAS 原样保留, 词不拆散 ('Trade secret'/'Business secret')."""
    from core.docx_reader import read_msds
    base = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS")
    for fname in ("PA-4835 msds_EN 国彩.docx", "EC-1804 msds_EN Guocai.docx"):
        p = next((x for x in base.rglob("*.docx") if x.name == fname), None)
        if not p:
            continue
        r = read_msds(p)
        casses = [c.cas for sec in r.sections.values() for c in sec.components]
        assert any("secret" in c.lower() for c in casses), (fname, casses)


def test_trade_secrets_plural_kept():
    """复数 'Trade Secrets' 占位符 CAS 应保留空格 (PU-3210 英文模板).

    回归: 原 _CAS_PLACEHOLDER_EN_SPACE_RE 只匹配单数 secret, 复数
    'Trade Secrets' 被去空格成 'TradeSecrets' — 占位语义被破坏, token
    完整性判定报 'Trade'/'Secrets' 遗漏.
    """
    from core.structure import normalize_component_cas
    assert normalize_component_cas("Trade Secrets") == "Trade Secrets"
    assert normalize_component_cas("Business Secrets") == "Business Secrets"
    assert normalize_component_cas("Trade secret") == "Trade secret"
    assert normalize_component_cas("Business secret") == "Business secret"


def test_mix_header_kept():
    """S3 单列 'Mixtures' 声明应保留 (OS/PTF 英文模板).

    回归: 原 _is_mix_header 命中后整格 continue, 'Mixtures' 标题本身
    被丢弃 (对应中文 '产品类型：混合物'); 现在作为 sub 保留, 成分 note
    仍正常提升.
    """
    from core.docx_reader import read_msds
    base = Path(r"F:\正式项目与模块化内容\Word 覆写模块\数据库\MSDS")
    for fname in ("OS-1400 msds_EN 冠志.docx", "PA-4006 msds_EN 国彩.docx"):
        p = next((x for x in base.rglob("*.docx") if x.name == fname), None)
        if not p:
            continue
        r = read_msds(p)
        s3 = r.sections.get(3)
        assert s3 is not None, fname
        # 任一 sub/field 标签含 'Mixtures' (大小写不敏感)
        labels = [row.label for row in s3.iter_rows()]
        assert any("mixtures" in (x or "").lower() for x in labels), (fname, labels)
        # 成分不丢: 至少 1 个成分被提升
        assert len(s3.components) >= 1, (fname, s3.components)
