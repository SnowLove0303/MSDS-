PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS schema_meta (
    meta_key TEXT PRIMARY KEY,
    meta_value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS cas_s2_profile (
    profile_id INTEGER PRIMARY KEY,
    cas_no TEXT NOT NULL,
    standard_name_zh TEXT NOT NULL,
    ec_no TEXT NOT NULL DEFAULT '',
    index_no TEXT NOT NULL DEFAULT '',
    language_code TEXT NOT NULL DEFAULT 'zh-CN',
    jurisdiction TEXT NOT NULL DEFAULT 'CN',
    profile_version TEXT NOT NULL DEFAULT '1.0.0',
    regulatory_version TEXT NOT NULL DEFAULT '',
    profile_status TEXT NOT NULL DEFAULT 'draft'
        CHECK (profile_status IN ('draft', 'confirmed', 'manual_review', 'retired')),
    is_current INTEGER NOT NULL DEFAULT 1 CHECK (is_current IN (0, 1)),
    remarks TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE (cas_no, profile_version, language_code, jurisdiction)
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_cas_s2_current_profile
    ON cas_s2_profile (cas_no)
    WHERE is_current = 1;

CREATE TABLE IF NOT EXISTS cas_s2_source (
    source_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    source_type TEXT NOT NULL,
    source_title TEXT NOT NULL DEFAULT '',
    source_file_path TEXT NOT NULL,
    source_file_name TEXT NOT NULL,
    source_hash_sha256 TEXT NOT NULL DEFAULT '',
    source_version TEXT NOT NULL DEFAULT '',
    effective_date TEXT NOT NULL DEFAULT '',
    collected_at TEXT NOT NULL DEFAULT '',
    jurisdiction TEXT NOT NULL DEFAULT 'CN',
    source_uri TEXT NOT NULL DEFAULT '',
    clause_reference TEXT NOT NULL DEFAULT '',
    source_priority TEXT NOT NULL DEFAULT 'reference_only'
        CHECK (source_priority IN ('current_official_standard', 'direct_product_test_or_supplier_confirmed_fact', 'traceable_substance_classification', 'historical_original_sds', 'internal_reference_only', 'reference_only')),
    source_status TEXT NOT NULL DEFAULT 'reference_only'
        CHECK (source_status IN ('reference_only', 'imported', 'verified', 'superseded')),
    evidence_text TEXT NOT NULL DEFAULT '',
    UNIQUE (profile_id, source_file_path, source_hash_sha256)
);

CREATE INDEX IF NOT EXISTS idx_cas_s2_source_profile_status
    ON cas_s2_source (profile_id, source_status, source_priority);

CREATE TABLE IF NOT EXISTS cas_s2_condition (
    condition_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    condition_key TEXT NOT NULL,
    applies_to_type TEXT NOT NULL DEFAULT 'profile',
    applies_to_code TEXT NOT NULL DEFAULT '*',
    condition_type TEXT NOT NULL
        CHECK (condition_type IN ('unconditional', 'lower_bound', 'upper_bound', 'range', 'not_provided', 'manual_review')),
    lower_operator TEXT NOT NULL DEFAULT '',
    lower_value REAL,
    upper_operator TEXT NOT NULL DEFAULT '',
    upper_value REAL,
    unit TEXT NOT NULL DEFAULT '%',
    scope TEXT NOT NULL DEFAULT 'mixture_component',
    raw_text TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'unknown'
        CHECK (status IN ('confirmed', 'unknown', 'conflicting', 'manual_review')),
    priority INTEGER NOT NULL DEFAULT 100,
    source_id INTEGER REFERENCES cas_s2_source(source_id) ON DELETE SET NULL,
    notes TEXT NOT NULL DEFAULT '',
    UNIQUE (profile_id, condition_key)
);

CREATE TABLE IF NOT EXISTS cas_s2_classification (
    classification_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    hazard_domain TEXT NOT NULL,
    hazard_class TEXT NOT NULL,
    category TEXT NOT NULL DEFAULT '',
    classification_text TEXT NOT NULL,
    h_code TEXT NOT NULL DEFAULT '',
    condition_id INTEGER REFERENCES cas_s2_condition(condition_id) ON DELETE SET NULL,
    condition_key TEXT NOT NULL DEFAULT 'profile_default',
    value_status TEXT NOT NULL DEFAULT 'derived_candidate'
        CHECK (value_status IN ('confirmed', 'range', 'unknown', 'not_applicable', 'confidential', 'conflicting', 'derived_candidate')),
    display_order INTEGER NOT NULL DEFAULT 100,
    source_id INTEGER REFERENCES cas_s2_source(source_id) ON DELETE SET NULL,
    raw_text TEXT NOT NULL DEFAULT '',
    UNIQUE (profile_id, hazard_domain, classification_text, h_code, condition_key)
);

CREATE INDEX IF NOT EXISTS idx_cas_s2_classification_h_code
    ON cas_s2_classification (h_code);

CREATE TABLE IF NOT EXISTS cas_s2_label_element (
    element_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    element_type TEXT NOT NULL
        CHECK (element_type IN ('pictogram', 'signal_word', 'hazard_statement', 'precautionary_statement')),
    element_code TEXT NOT NULL DEFAULT '',
    element_text_zh TEXT NOT NULL DEFAULT '',
    condition_id INTEGER REFERENCES cas_s2_condition(condition_id) ON DELETE SET NULL,
    condition_key TEXT NOT NULL DEFAULT 'profile_default',
    value_status TEXT NOT NULL DEFAULT 'derived_candidate'
        CHECK (value_status IN ('confirmed', 'range', 'unknown', 'not_applicable', 'confidential', 'conflicting', 'derived_candidate')),
    display_order INTEGER NOT NULL DEFAULT 100,
    source_id INTEGER REFERENCES cas_s2_source(source_id) ON DELETE SET NULL,
    raw_text TEXT NOT NULL DEFAULT '',
    UNIQUE (profile_id, element_type, element_code, condition_key)
);

CREATE INDEX IF NOT EXISTS idx_cas_s2_label_code
    ON cas_s2_label_element (element_type, element_code);

CREATE TABLE IF NOT EXISTS cas_s2_hazard_summary (
    summary_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    hazard_domain TEXT NOT NULL
        CHECK (hazard_domain IN ('physical_chemical', 'health', 'environment', 'other')),
    summary_text TEXT NOT NULL DEFAULT '',
    value_status TEXT NOT NULL DEFAULT 'unknown'
        CHECK (value_status IN ('confirmed', 'range', 'unknown', 'not_applicable', 'confidential', 'conflicting', 'derived_candidate')),
    condition_id INTEGER REFERENCES cas_s2_condition(condition_id) ON DELETE SET NULL,
    source_id INTEGER REFERENCES cas_s2_source(source_id) ON DELETE SET NULL,
    raw_text TEXT NOT NULL DEFAULT '',
    UNIQUE (profile_id, hazard_domain)
);

CREATE TABLE IF NOT EXISTS cas_s2_fact (
    fact_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    fact_domain TEXT NOT NULL,
    fact_code TEXT NOT NULL,
    fact_label_zh TEXT NOT NULL DEFAULT '',
    value_text TEXT NOT NULL DEFAULT '',
    numeric_value REAL,
    lower_value REAL,
    upper_value REAL,
    unit TEXT NOT NULL DEFAULT '',
    qualifier TEXT NOT NULL DEFAULT '',
    applicability_scope TEXT NOT NULL DEFAULT '',
    value_status TEXT NOT NULL DEFAULT 'reference_only'
        CHECK (value_status IN ('confirmed', 'unknown', 'not_listed', 'not_applicable', 'conflicting', 'reference_only', 'derived_candidate', 'superseded')),
    source_id INTEGER REFERENCES cas_s2_source(source_id) ON DELETE SET NULL,
    clause_reference TEXT NOT NULL DEFAULT '',
    raw_text TEXT NOT NULL DEFAULT '',
    display_order INTEGER NOT NULL DEFAULT 100,
    UNIQUE (profile_id, fact_code, source_id, qualifier)
);

CREATE INDEX IF NOT EXISTS idx_cas_s2_fact_domain_code
    ON cas_s2_fact (profile_id, fact_domain, fact_code, value_status);

CREATE TABLE IF NOT EXISTS cas_s2_review (
    review_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    review_scope TEXT NOT NULL DEFAULT 'profile',
    review_status TEXT NOT NULL
        CHECK (review_status IN ('pending', 'manual_review', 'verified', 'rejected')),
    reviewer TEXT NOT NULL DEFAULT '',
    reviewed_at TEXT NOT NULL DEFAULT '',
    review_note TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_cas_s2_review_profile
    ON cas_s2_review (profile_id, review_status);

-- Fixed output contract from the confirmed NMP workbook.  Every current CAS
-- profile must expose exactly these 13 fields to the Web result layer.
CREATE TABLE IF NOT EXISTS cas_s2_skeleton_field (
    field_key TEXT PRIMARY KEY,
    field_label_zh TEXT NOT NULL,
    field_type TEXT NOT NULL
        CHECK (field_type IN ('identity', 'classification', 'label', 'summary')),
    display_order INTEGER NOT NULL UNIQUE,
    required INTEGER NOT NULL DEFAULT 1 CHECK (required IN (0, 1))
);

CREATE TABLE IF NOT EXISTS cas_s2_result_field (
    result_field_id INTEGER PRIMARY KEY,
    profile_id INTEGER NOT NULL REFERENCES cas_s2_profile(profile_id) ON DELETE CASCADE,
    field_key TEXT NOT NULL REFERENCES cas_s2_skeleton_field(field_key),
    value_text TEXT NOT NULL DEFAULT '无数据',
    value_status TEXT NOT NULL DEFAULT 'unknown'
        CHECK (value_status IN ('confirmed', 'unknown', 'not_applicable', 'conflicting', 'reference_only', 'derived_candidate', 'superseded')),
    condition_id INTEGER REFERENCES cas_s2_condition(condition_id) ON DELETE SET NULL,
    source_id INTEGER REFERENCES cas_s2_source(source_id) ON DELETE SET NULL,
    raw_text TEXT NOT NULL DEFAULT '',
    display_order INTEGER NOT NULL,
    UNIQUE (profile_id, field_key)
);

CREATE INDEX IF NOT EXISTS idx_cas_s2_result_field_profile_order
    ON cas_s2_result_field (profile_id, display_order);

CREATE TRIGGER IF NOT EXISTS trg_cas_s2_result_field_no_explanation_insert
BEFORE INSERT ON cas_s2_result_field
WHEN NEW.value_text LIKE '%来源和方法%'
  OR NEW.value_text LIKE '%追溯分类证据%'
  OR NEW.value_text LIKE '%国际参考事实%'
  OR NEW.value_text LIKE '%人工复核%'
  OR NEW.value_text LIKE '%需结合具体配方%'
  OR NEW.value_text LIKE '%参考数据%'
  OR NEW.value_text LIKE '%参考事实%'
  OR NEW.value_text LIKE '%未据此自动%'
  OR NEW.value_text LIKE '%暂无可直接升级%'
BEGIN
    SELECT RAISE(ABORT, 'cas_s2_result_field 只允许直接 Section 2 标准结果');
END;

CREATE TRIGGER IF NOT EXISTS trg_cas_s2_result_field_no_explanation_update
BEFORE UPDATE OF value_text ON cas_s2_result_field
WHEN NEW.value_text LIKE '%来源和方法%'
  OR NEW.value_text LIKE '%追溯分类证据%'
  OR NEW.value_text LIKE '%国际参考事实%'
  OR NEW.value_text LIKE '%人工复核%'
  OR NEW.value_text LIKE '%需结合具体配方%'
  OR NEW.value_text LIKE '%参考数据%'
  OR NEW.value_text LIKE '%参考事实%'
  OR NEW.value_text LIKE '%未据此自动%'
  OR NEW.value_text LIKE '%暂无可直接升级%'
BEGIN
    SELECT RAISE(ABORT, 'cas_s2_result_field 只允许直接 Section 2 标准结果');
END;

INSERT INTO cas_s2_skeleton_field (field_key, field_label_zh, field_type, display_order, required) VALUES
    ('standard_name', '标准名称', 'identity', 1, 1),
    ('cas_no', 'CAS', 'identity', 2, 1),
    ('ec_no', 'EC', 'identity', 3, 1),
    ('index_no', 'Index No', 'identity', 4, 1),
    ('ghs_classification', 'GHS 危险性类别', 'classification', 5, 1),
    ('pictogram', '象形图', 'label', 6, 1),
    ('signal_word', '信号词', 'label', 7, 1),
    ('hazard_statement', '危险性说明', 'label', 8, 1),
    ('precautionary_statement', '防范说明', 'label', 9, 1),
    ('physical_chemical_hazard', '物理与化学危险', 'summary', 10, 1),
    ('health_hazard', '健康危害', 'summary', 11, 1),
    ('environmental_hazard', '环境危害', 'summary', 12, 1),
    ('other_hazard', '其他危害', 'summary', 13, 1)
ON CONFLICT(field_key) DO NOTHING;

CREATE TRIGGER IF NOT EXISTS trg_cas_s2_skeleton_field_no_extra_insert
BEFORE INSERT ON cas_s2_skeleton_field
WHEN NOT (
    (NEW.field_key = 'standard_name' AND NEW.field_label_zh = '标准名称' AND NEW.field_type = 'identity' AND NEW.display_order = 1 AND NEW.required = 1) OR
    (NEW.field_key = 'cas_no' AND NEW.field_label_zh = 'CAS' AND NEW.field_type = 'identity' AND NEW.display_order = 2 AND NEW.required = 1) OR
    (NEW.field_key = 'ec_no' AND NEW.field_label_zh = 'EC' AND NEW.field_type = 'identity' AND NEW.display_order = 3 AND NEW.required = 1) OR
    (NEW.field_key = 'index_no' AND NEW.field_label_zh = 'Index No' AND NEW.field_type = 'identity' AND NEW.display_order = 4 AND NEW.required = 1) OR
    (NEW.field_key = 'ghs_classification' AND NEW.field_label_zh = 'GHS 危险性类别' AND NEW.field_type = 'classification' AND NEW.display_order = 5 AND NEW.required = 1) OR
    (NEW.field_key = 'pictogram' AND NEW.field_label_zh = '象形图' AND NEW.field_type = 'label' AND NEW.display_order = 6 AND NEW.required = 1) OR
    (NEW.field_key = 'signal_word' AND NEW.field_label_zh = '信号词' AND NEW.field_type = 'label' AND NEW.display_order = 7 AND NEW.required = 1) OR
    (NEW.field_key = 'hazard_statement' AND NEW.field_label_zh = '危险性说明' AND NEW.field_type = 'label' AND NEW.display_order = 8 AND NEW.required = 1) OR
    (NEW.field_key = 'precautionary_statement' AND NEW.field_label_zh = '防范说明' AND NEW.field_type = 'label' AND NEW.display_order = 9 AND NEW.required = 1) OR
    (NEW.field_key = 'physical_chemical_hazard' AND NEW.field_label_zh = '物理与化学危险' AND NEW.field_type = 'summary' AND NEW.display_order = 10 AND NEW.required = 1) OR
    (NEW.field_key = 'health_hazard' AND NEW.field_label_zh = '健康危害' AND NEW.field_type = 'summary' AND NEW.display_order = 11 AND NEW.required = 1) OR
    (NEW.field_key = 'environmental_hazard' AND NEW.field_label_zh = '环境危害' AND NEW.field_type = 'summary' AND NEW.display_order = 12 AND NEW.required = 1) OR
    (NEW.field_key = 'other_hazard' AND NEW.field_label_zh = '其他危害' AND NEW.field_type = 'summary' AND NEW.display_order = 13 AND NEW.required = 1)
)
BEGIN
    SELECT RAISE(ABORT, 'cas_s2_skeleton_field 仅允许固定 13 个标准字段');
END;

CREATE TRIGGER IF NOT EXISTS trg_cas_s2_skeleton_field_no_update
BEFORE UPDATE ON cas_s2_skeleton_field
WHEN OLD.field_key <> NEW.field_key
  OR OLD.field_label_zh <> NEW.field_label_zh
  OR OLD.field_type <> NEW.field_type
  OR OLD.display_order <> NEW.display_order
  OR OLD.required <> NEW.required
BEGIN
    SELECT RAISE(ABORT, 'cas_s2_skeleton_field 标准骨架不可修改');
END;

CREATE TRIGGER IF NOT EXISTS trg_cas_s2_skeleton_field_no_delete
BEFORE DELETE ON cas_s2_skeleton_field
BEGIN
    SELECT RAISE(ABORT, 'cas_s2_skeleton_field 标准骨架不可删除');
END;

CREATE VIEW IF NOT EXISTS v_cas_s2_current_profile AS
SELECT
    p.profile_id,
    p.cas_no,
    p.standard_name_zh,
    p.ec_no,
    p.index_no,
    p.language_code,
    p.jurisdiction,
    p.profile_version,
    p.regulatory_version,
    p.profile_status,
    p.is_current,
    COUNT(DISTINCT CASE
        WHEN s.source_file_path <> '' AND s.source_file_path NOT LIKE 'remote://%'
        THEN s.source_file_path || '|' || s.source_hash_sha256
    END) AS source_count,
    p.created_at,
    p.updated_at
FROM cas_s2_profile AS p
LEFT JOIN cas_s2_source AS s ON s.profile_id = p.profile_id
WHERE p.is_current = 1
GROUP BY p.profile_id;

CREATE VIEW IF NOT EXISTS v_cas_s2_current_standard_result AS
SELECT
    p.profile_id,
    p.cas_no,
    p.standard_name_zh,
    p.profile_version,
    p.profile_status,
    s.field_key,
    s.field_label_zh,
    s.field_type,
    s.display_order,
    COALESCE(r.value_text, '无数据') AS value_text,
    COALESCE(r.value_status, 'unknown') AS value_status,
    r.condition_id,
    r.source_id,
    COALESCE(r.raw_text, '') AS raw_text
FROM cas_s2_profile AS p
CROSS JOIN cas_s2_skeleton_field AS s
LEFT JOIN cas_s2_result_field AS r
    ON r.profile_id = p.profile_id AND r.field_key = s.field_key
WHERE p.is_current = 1;

INSERT OR REPLACE INTO schema_meta (meta_key, meta_value) VALUES
    ('schema_id', 'cas_section2_result_library'),
    ('schema_version', '1.3.0'),
    ('language', 'zh-CN'),
    ('purpose', 'CAS物质级Section 2标准结果、法规证据、结构化事实与含量条件库'),
    ('identity_scope', 'CAS身份库独立维护；本库只保存CAS Section 2结果事实'),
    ('current_profile_rule', '同一CAS仅允许一个当前有效档案；历史版本可保留但不得并列参与默认推导'),
    ('facts_supported', 'true'),
    ('standard_skeleton_id', 'nmp_sheet1_13_fields'),
    ('standard_skeleton_field_count', '13'),
    ('standard_skeleton_missing_value', '无数据'),
    ('source_locator_rule', '本地来源使用真实文件路径；远程法规来源使用 remote:// 标识并将正式 URL 保存于 source_uri'),
    ('threshold_scope_rule', '法规含量阈值必须区分中国混合物组分披露、成品分类、境外限制等适用范围，不得混用');

-- GHS 象形图标准字典：图片作为数据库 BLOB 保存，供 CAS 检索和其他输出模块复用。
CREATE TABLE IF NOT EXISTS cas_s2_pictogram_catalog (
    pictogram_code TEXT PRIMARY KEY,
    label_zh TEXT NOT NULL,
    image_filename TEXT NOT NULL,
    mime_type TEXT NOT NULL,
    image_blob BLOB NOT NULL,
    image_sha256 TEXT NOT NULL,
    source_path TEXT NOT NULL,
    source_status TEXT NOT NULL DEFAULT 'verified'
);
