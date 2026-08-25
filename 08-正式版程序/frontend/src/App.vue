<script setup>
import { ref, onMounted, computed } from 'vue'
import { api } from './api'

const libraries = ref([])
const libMeta = ref(null)      // 当前库元信息（含表）
const selected = ref(null)     // 当前选中树节点 {type, ...}
const loading = ref(false)
const error = ref('')
const expanded = ref(new Set()) // 展开的节点 key

// ---- 节点类型: lib / table / row / category / model / section / cas
function nodeKey(n) { return n.type + ':' + (n.id ?? '') }

function toggle(key) {
  const s = expanded.value
  s.has(key) ? s.delete(key) : s.add(key)
  expanded.value = new Set(s)
}
function isOpen(key) { return expanded.value.has(key) }

// ---- 加载库列表
async function loadLibraries() {
  try {
    libraries.value = await api.libraries()
    // 型号库: 预加载大类列表挂到 lib.categories
    const ml = getLib('msds_standard.db')
    if (ml) {
      try { ml.categories = await api.categories() } catch (e) { ml.categories = [] }
    }
  } catch (e) { error.value = e.message }
}
function getLib(key) { return libraries.value.find(l => l.key === key) }

// ---- 选择节点
function select(n) { selected.value = n }
function isSelected(n) { return selected.value && selected.value.id === n.id && selected.value.type === n.type && (selected.value.lib === n.lib) }

// ---- 加载表行
const rowsState = ref({})  // key: lib:table -> {rows, columns, total, page}
async function loadTableRows(libKey, table, page = 1) {
  const k = `${libKey}:${table}`
  try {
    const r = await api.tableRows(libKey, table, page, 50)
    rowsState.value[k] = { rows: r.rows, columns: r.columns, total: r.total, page: r.page, page_size: r.page_size }
  } catch (e) { error.value = e.message }
}

// ---- 行详情
const rowDetailState = ref(null)
async function showRowDetail(libKey, table, rowid) {
  try { rowDetailState.value = await api.rowDetail(libKey, table, rowid) }
  catch (e) { error.value = e.message }
}

// ---- 大类
const catModels = ref(null)
async function loadCatModels(parent) {
  try { catModels.value = await api.modelsByCategory(parent) }
  catch (e) { error.value = e.message }
}

// ---- 型号节/字段
const modelSections = ref(null)
const sectionFields = ref(null)
async function loadModelSections(modelId) {
  modelSections.value = null; sectionFields.value = null
  try { modelSections.value = await api.modelSections(modelId) }
  catch (e) { error.value = e.message }
}
async function loadSectionFields(modelId, section) {
  try { sectionFields.value = await api.sectionFields(modelId, section) }
  catch (e) { error.value = e.message }
}

// ---- CAS
const casList = ref([])
const casDetailState = ref(null)
async function loadCasList() {
  try { casList.value = await api.casSubstances() }
  catch (e) { error.value = e.message }
}
async function loadCasDetail(casId) {
  try { casDetailState.value = await api.casDetail(casId) }
  catch (e) { error.value = e.message }
}

onMounted(loadLibraries)
</script>

<template>
  <div class="app">
    <header class="topbar">
      <div class="brand"><span class="logo">▦</span> MSDS 分库总览</div>
      <div class="topbar-right">
        <span class="badge">{{ libraries.length }} 个库</span>
        <span class="badge muted">只读浏览</span>
      </div>
    </header>

    <div class="layout">
      <!-- 左侧树 -->
      <aside class="tree-panel">
        <div class="tree-head">数据库</div>
        <div class="tree-body">
          <div v-for="lib in libraries" :key="lib.key" class="tree-node">
            <div class="node-row lib"
                 :class="{ open: isOpen('lib:' + lib.key) }"
                 @click="toggle('lib:' + lib.key); select({ type: 'lib', id: lib.key, lib: lib.key, name: lib.name })">
              <span class="caret">{{ isOpen('lib:' + lib.key) ? '▾' : '▸' }}</span>
              <span class="icon">📚</span>
              <span class="label">{{ lib.name }}</span>
              <span class="count">{{ lib.tables.length }} 表</span>
            </div>
            <div v-if="isOpen('lib:' + lib.key)" class="tree-children">
              <!-- 型号库业务树 -->
              <template v-if="lib.key === 'msds_standard.db'">
                <div class="node-row group" @click="toggle('cats:' + lib.key)">
                  <span class="caret">{{ isOpen('cats:' + lib.key) ? '▾' : '▸' }}</span>
                  <span class="icon">🗂️</span><span class="label">型号分类（业务树）</span>
                </div>
                <div v-if="isOpen('cats:' + lib.key)" class="tree-children">
                  <div v-for="c in lib.categories" :key="c.parent" class="tree-node">
                    <div class="node-row cat"
                         @click="toggle('cat:' + c.parent); loadCatModels(c.parent); select({ type: 'category', id: c.parent, lib: lib.key, name: c.parent })">
                      <span class="caret">{{ isOpen('cat:' + c.parent) ? '▾' : '▸' }}</span>
                      <span class="icon">📁</span>
                      <span class="label">{{ c.parent }}</span>
                      <span class="count">{{ c.model_count }}</span>
                    </div>
                    <div v-if="isOpen('cat:' + c.parent)" class="tree-children">
                      <div v-for="m in catModels" :key="m.model_id" class="tree-node">
                        <div class="node-row model"
                             @click="toggle('model:' + m.model_id); loadModelSections(m.model_id); select({ type: 'model', id: m.model_id, lib: lib.key, name: m.model })">
                          <span class="caret">{{ isOpen('model:' + m.model_id) ? '▾' : '▸' }}</span>
                          <span class="icon">📄</span>
                          <span class="label">{{ m.model }}</span>
                          <span class="count">{{ m.fields_count }}</span>
                        </div>
                        <div v-if="isOpen('model:' + m.model_id)" class="tree-children">
                          <div v-for="s in modelSections" :key="s.section" class="node-row section"
                               @click="loadSectionFields(m.model_id, s.section); select({ type: 'section', id: `${m.model_id}-${s.section}`, lib: lib.key, model_id: m.model_id, section: s.section, name: `S${s.section}`, model_name: m.model })">
                            <span class="icon">§</span>
                            <span class="label">S{{ s.section }}</span>
                            <span class="count">{{ s.field_count }}</span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </template>
              <!-- CAS 库业务树 -->
              <template v-else-if="lib.key === 'cas_library.db'">
                <div class="node-row group" @click="toggle('cas:' + lib.key); loadCasList()">
                  <span class="caret">{{ isOpen('cas:' + lib.key) ? '▾' : '▸' }}</span>
                  <span class="icon">🧪</span><span class="label">物质列表（业务树）</span>
                </div>
                <div v-if="isOpen('cas:' + lib.key)" class="tree-children">
                  <div v-for="s in casList" :key="s.cas_id" class="node-row cas"
                       @click="loadCasDetail(s.cas_id); select({ type: 'cas', id: s.cas_id, lib: lib.key, name: s.standard_name })">
                    <span class="icon">🧪</span>
                    <span class="label">{{ s.standard_name }}</span>
                    <span class="count">{{ s.cas_no }}</span>
                  </div>
                </div>
              </template>
              <!-- 通用表 -->
              <div v-for="t in lib.tables" :key="t.name" class="tree-node">
                <div class="node-row table"
                     @click="toggle('table:' + lib.key + ':' + t.name); loadTableRows(lib.key, t.name); select({ type: 'table', id: t.name, lib: lib.key, name: t.name })">
                  <span class="caret">{{ isOpen('table:' + lib.key + ':' + t.name) ? '▾' : '▸' }}</span>
                  <span class="icon">🗄️</span>
                  <span class="label">{{ t.name }}</span>
                  <span class="count">{{ t.rows }}</span>
                </div>
                <div v-if="isOpen('table:' + lib.key + ':' + t.name)" class="tree-children">
                  <div v-for="r in (rowsState[lib.key + ':' + t.name]?.rows || [])" :key="r.__rowid__ ?? r[Object.keys(r)[0]]"
                       class="node-row row" @click="showRowDetail(lib.key, t.name, r.__rowid__); select({ type: 'row', id: r.__rowid__, lib: lib.key, table: t.name, name: String(r[Object.keys(r)[0]] ?? r.__rowid__) })">
                    <span class="icon">▸</span>
                    <span class="label mono">{{ r[Object.keys(r)[0]] ?? r.__rowid__ }}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </aside>

      <!-- 右侧内容 -->
      <main class="content-panel">
        <div v-if="error" class="error-box">{{ error }}</div>
        <div v-if="loading" class="loading">加载中…</div>

        <!-- 未选中 -->
        <div v-if="!selected && !error" class="empty">
          <div class="empty-icon">▦</div>
          <p>从左侧选择数据库 / 分类 / 型号 / 表，查看真实数据</p>
        </div>

        <!-- 库概览 -->
        <template v-else-if="selected.type === 'lib'">
          <h2 class="panel-title">{{ selected.name }}</h2>
          <p class="panel-sub mono">{{ getLib(selected.id)?.path }}</p>
          <div class="table-grid">
            <table class="data-table">
              <thead><tr><th>表</th><th>行数</th></tr></thead>
              <tbody>
                <tr v-for="t in getLib(selected.id)?.tables || []" :key="t.name">
                  <td class="mono">{{ t.name }}</td><td>{{ t.rows }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </template>

        <!-- 大类 -->
        <template v-else-if="selected.type === 'category'">
          <h2 class="panel-title">📁 {{ selected.name }}</h2>
          <div class="table-grid">
            <table class="data-table">
              <thead><tr><th>型号</th><th>来源</th><th>字段数</th></tr></thead>
              <tbody>
                <tr v-for="m in catModels" :key="m.model_id" @click="loadModelSections(m.model_id); select({ type: 'model', id: m.model_id, lib: selected.lib, name: m.model })">
                  <td><a>{{ m.model }}</a></td><td>{{ m.source }}</td><td>{{ m.fields_count }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </template>

        <!-- 型号概览 -->
        <template v-else-if="selected.type === 'model'">
          <h2 class="panel-title">📄 {{ selected.name }}</h2>
          <div class="sections-bar">
            <span v-for="s in modelSections" :key="s.section" class="chip"
                  @click="loadSectionFields(selected.id, s.section); select({ type: 'section', id: `${selected.id}-${s.section}`, lib: selected.lib, model_id: selected.id, section: s.section, name: `S${s.section}`, model_name: selected.name })">
              S{{ s.section }} <i>{{ s.field_count }}</i>
            </span>
          </div>
          <div class="panel-hint">点击上方节徽章查看该节字段明细</div>
        </template>

        <!-- 节字段 -->
        <template v-else-if="selected.type === 'section'">
          <h2 class="panel-title">§ {{ selected.model_name }} / S{{ selected.section }}</h2>
          <div class="table-grid">
            <table class="data-table">
              <thead><tr><th>序号</th><th>标签</th><th>标准名</th><th>值</th><th>类型</th><th>可编辑</th></tr></thead>
              <tbody>
                <tr v-for="f in sectionFields" :key="f.seq" class="field-row">
                  <td class="mono">{{ f.seq }}</td>
                  <td>{{ f.label }}</td>
                  <td class="muted">{{ f.std_name }}</td>
                  <td class="value-cell">{{ f.value }}</td>
                  <td>{{ f.kind }}</td>
                  <td>{{ f.editable ? '✓' : '—' }}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </template>

        <!-- CAS 详情 -->
        <template v-else-if="selected.type === 'cas'">
          <h2 class="panel-title">🧪 {{ casDetailState?.standard_name }}</h2>
          <div class="detail-grid" v-if="casDetailState">
            <div class="detail-item"><label>CAS 号</label><span class="mono">{{ casDetailState.cas_no }}</span></div>
            <div class="detail-item"><label>分类</label><span>{{ casDetailState.category }}</span></div>
            <div class="detail-item"><label>登记状态</label><span>{{ casDetailState.registry_status }}</span></div>
            <div class="detail-item"><label>来源数/型号数</label><span>{{ casDetailState.source_count }} / {{ casDetailState.model_count }}</span></div>
          </div>
          <h3 class="sub-title">别名</h3>
          <table class="data-table" v-if="casDetailState?.aliases?.length">
            <thead><tr><th>别名</th><th>类型</th><th>出现次数</th></tr></thead>
            <tbody>
              <tr v-for="a in casDetailState.aliases" :key="a.alias"><td>{{ a.alias }}</td><td>{{ a.alias_type }}</td><td>{{ a.occurrence_count }}</td></tr>
            </tbody>
          </table>
          <h3 class="sub-title">关联型号</h3>
          <table class="data-table" v-if="casDetailState?.usage?.length">
            <thead><tr><th>型号</th><th>原始名称</th><th>原始CAS</th><th>浓度</th><th>操作</th></tr></thead>
            <tbody>
              <tr v-for="u in casDetailState.usage" :key="u.model + u.raw_name">
                <td>{{ u.model }}</td><td>{{ u.raw_name }}</td><td class="mono">{{ u.raw_cas }}</td><td>{{ u.concentration }}</td>
                <td><a v-if="u.model_id" @click="loadModelSections(u.model_id); select({ type: 'model', id: u.model_id, lib: 'msds_standard.db', name: u.model })">查看型号 →</a></td>
              </tr>
            </tbody>
          </table>
        </template>

        <!-- 通用表 -->
        <template v-else-if="selected.type === 'table'">
          <h2 class="panel-title">🗄️ {{ selected.name }}</h2>
          <p class="panel-sub muted">共 {{ rowsState[selected.lib + ':' + selected.name]?.total ?? 0 }} 行</p>
          <div class="table-grid wide">
            <table class="data-table">
              <thead>
                <tr><th v-for="c in rowsState[selected.lib + ':' + selected.name]?.columns || []" :key="c">{{ c }}</th><th>操作</th></tr>
              </thead>
              <tbody>
                <tr v-for="(r, idx) in rowsState[selected.lib + ':' + selected.name]?.rows || []" :key="idx">
                  <td v-for="c in rowsState[selected.lib + ':' + selected.name]?.columns || []" :key="c" class="value-cell">{{ r[c] }}</td>
                  <td><a @click="showRowDetail(selected.lib, selected.name, r.__rowid__)">详情</a></td>
                </tr>
              </tbody>
            </table>
          </div>
        </template>

        <!-- 行详情 -->
        <template v-else-if="selected.type === 'row' && rowDetailState">
          <h2 class="panel-title">▸ {{ selected.name }} <span class="muted mono">#{{ selected.id }}</span></h2>
          <div class="detail-grid">
            <div class="detail-item" v-for="(v, k) in rowDetailState" :key="k">
              <label>{{ k }}</label><span class="value-cell">{{ v }}</span>
            </div>
          </div>
        </template>
      </main>
    </div>
  </div>
</template>