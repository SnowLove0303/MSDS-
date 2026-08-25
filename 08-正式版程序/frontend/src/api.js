// API 封装：分库总览
const BASE = '/api/libraries'

async function jfetch(url) {
  const r = await fetch(url)
  if (!r.ok) {
    let detail = r.statusText
    try { detail = (await r.json()).detail || detail } catch {}
    throw new Error(`${r.status} ${detail}`)
  }
  return r.json()
}

export const api = {
  health: () => jfetch('/api/health'),
  libraries: () => jfetch(`${BASE}`),
  // 型号库业务树
  categories: () => jfetch(`${BASE}/models/categories`),
  modelsByCategory: (parent) => jfetch(`${BASE}/models/categories/${encodeURIComponent(parent)}/models`),
  modelSections: (modelId) => jfetch(`${BASE}/models/${modelId}/sections`),
  sectionFields: (modelId, section) => jfetch(`${BASE}/models/${modelId}/sections/${section}`),
  modelSummary: (modelId) => jfetch(`${BASE}/models/${modelId}/summary`),
  // CAS 库业务树
  casSubstances: (keyword = '') => jfetch(`${BASE}/cas/substances?keyword=${encodeURIComponent(keyword)}`),
  casDetail: (casId) => jfetch(`${BASE}/cas/substances/${casId}`),
  // 通用表
  tableRows: (libKey, table, page = 1, pageSize = 20, keyword = '') =>
    jfetch(`${BASE}/${libKey}/tables/${table}/rows?page=${page}&page_size=${pageSize}&keyword=${encodeURIComponent(keyword)}`),
  rowDetail: (libKey, table, rowid) => jfetch(`${BASE}/${libKey}/tables/${table}/${rowid}`),
}