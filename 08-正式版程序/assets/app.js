// GZ 应用工作台 — 桌面外壳 + 检索工作台
(function () {
  "use strict";

  // ===== 配置 =====
  const API_OLD = window.location.origin;        // 旧 server.py
  const API_NEW = API_OLD;                         // 与当前正式入口保持同源
  const SECTION_NAMES = { 0:"页眉页脚",1:"物料及供应商标识",2:"危险性概述",3:"成分/组成信息",4:"急救措施",5:"消防措施",6:"泄漏应急处理",7:"操作处置与储存",8:"接触控制/个体防护",9:"物理和化学性质",10:"稳定性和反应性",11:"毒理学信息",12:"生态学信息",13:"废弃处置",14:"运输信息",15:"法规信息",16:"其他信息" };

  // ===== 工具函数 =====
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" })[c]);
  let ghsCatalog = {};
  const ghsCodes = value => Array.from(new Set((String(value || "").match(/GHS0[1-9]/gi) || []).map(x => x.toUpperCase())));
  function renderGhsValue(value, emptyText = "未识别 GHS 编号") {
    const codes = ghsCodes(value);
    if (!codes.length) return `<div class="ow-ghs-preview is-empty">${esc(emptyText)}</div>`;
    return `<div class="ow-ghs-preview">${codes.map(code => {
      const item = ghsCatalog[code] || { code, label: "未登记含义", data_uri: "" };
      const image = item.data_uri ? `<img src="${esc(item.data_uri)}" alt="${esc(code)}" loading="lazy">` : `<span class="ow-ghs-missing">?</span>`;
      return `<span class="ow-ghs-item">${image}<b>${esc(code)}</b><em>${esc(item.label)}</em></span>`;
    }).join("")}</div>`;
  }
  function loadGhsCatalog() {
    return api(API_OLD + "/api/ghs/pictograms").then(data => { ghsCatalog = data.items || {}; return ghsCatalog; });
  }
  function api(url, opt) {
    return fetch(url, opt || {}).then(r => r.status === 204 ? null : r.json());
  }
  let toastTimer;
  function toast(msg, sec) {
    const el = document.getElementById("toast") || (() => {
      const t = document.createElement("div"); t.id = "toast"; t.style.cssText = "position:fixed;bottom:60px;left:50%;transform:translateX(-50%);background:#1a2a3a;color:#fff;padding:10px 24px;border-radius:8px;font-size:13px;z-index:99999;opacity:0;transition:opacity .3s;pointer-events:none";
      document.body.appendChild(t); return t;
    })();
    el.textContent = msg; el.style.opacity = "1";
    clearTimeout(toastTimer); toastTimer = setTimeout(() => el.style.opacity = "0", (sec || 3) * 1000);
  }

  // ===== 窗口管理器 =====
  const windows = {};
  let winZIndex = 100;

  function createWindow(id, cfg) {
    const el = document.createElement("div");
    el.className = `gzwindow ${cfg.className || ""}`;
    el.style.cssText = `left:${cfg.x || 100}px;top:${cfg.y || 80}px;width:${cfg.w || 1000}px;height:${cfg.h || 700}px;z-index:${++winZIndex}`;
    const windowActions = cfg.hideMinimize
      ? `<button class="gzw-btn" data-act="close">✕</button>`
      : `<button class="gzw-btn" data-act="min">─</button><button class="gzw-btn" data-act="close">✕</button>`;
    el.innerHTML = `
      <div class="gzwindow-titlebar" data-win="${id}">
        <span class="gzwindow-title">${esc(cfg.title || "应用")}</span>
        <span class="gzwindow-actions">${windowActions}</span>
      </div>
      <div class="gzwindow-body" data-win="${id}"></div>
      <div class="gzw-resize gzw-resize-n" data-dir="n"></div>
      <div class="gzw-resize gzw-resize-s" data-dir="s"></div>
      <div class="gzw-resize gzw-resize-e" data-dir="e"></div>
      <div class="gzw-resize gzw-resize-w" data-dir="w"></div>
      <div class="gzw-resize gzw-resize-ne" data-dir="ne"></div>
      <div class="gzw-resize gzw-resize-nw" data-dir="nw"></div>
      <div class="gzw-resize gzw-resize-se" data-dir="se"></div>
      <div class="gzw-resize gzw-resize-sw" data-dir="sw"></div>`;
    document.body.appendChild(el);
    const body = el.querySelector(".gzwindow-body");

    // 标题栏拖动
    const bar = el.querySelector(".gzwindow-titlebar");
    let offX, offY, dragging = false;
    bar.onmousedown = e => {
      if (e.target.tagName === "BUTTON" || e.target.closest(".gzwindow-actions")) return;
      dragging = true; offX = e.clientX - el.offsetLeft; offY = e.clientY - el.offsetTop;
      el.style.cursor = "move";
    };
    document.addEventListener("mousemove", e => {
      if (!dragging) return;
      el.style.left = Math.max(0, e.clientX - offX) + "px";
      el.style.top = Math.max(0, e.clientY - offY) + "px";
    });
    document.addEventListener("mouseup", () => { if (dragging) { dragging = false; el.style.cursor = ""; } });

    // 8 方向缩放
    let resizing = false, rDir = "", rStart = {}, rRect = {};
    const MIN_W = 600, MIN_H = 400;
    el.querySelectorAll(".gzw-resize").forEach(h => {
      h.onmousedown = e => {
        e.stopPropagation();
        resizing = true; rDir = h.dataset.dir; rRect = el.getBoundingClientRect();
        rStart = { x: e.clientX, y: e.clientY };
      };
    });
    document.addEventListener("mousemove", e => {
      if (!resizing) return;
      const dx = e.clientX - rStart.x, dy = e.clientY - rStart.y;
      let { left, top, width, height } = rRect;
      const d = rDir;
      if (d.includes("e")) width = Math.max(MIN_W, rRect.width + dx);
      if (d.includes("w")) { const w = Math.max(MIN_W, rRect.width - dx); left = rRect.right - w; width = w; }
      if (d.includes("s")) height = Math.max(MIN_H, rRect.height + dy);
      if (d.includes("n")) { const h = Math.max(MIN_H, rRect.height - dy); top = rRect.bottom - h; height = h; }
      el.style.left = left + "px"; el.style.top = top + "px";
      el.style.width = width + "px"; el.style.height = height + "px";
    });
    document.addEventListener("mouseup", () => { resizing = false; });

    // 按钮
    el.querySelector("[data-act=close]").onclick = () => closeWindow(id);
    el.querySelector("[data-act=min]")?.addEventListener("click", () => {
      el.style.display = "none";
      updateTaskbar();
    });

    el.onmousedown = () => { el.style.zIndex = ++winZIndex; };
    const win = { id, el, body, cfg };
    windows[id] = win;
    updateTaskbar();
    setTimeout(() => { el.style.opacity = "1"; }, 10);
    return win;
  }

  function closeWindow(id) {
    const w = windows[id]; if (!w) return;
    w.el.remove(); delete windows[id]; updateTaskbar();
  }

  function updateTaskbar() {
    const space = $(".taskbar-space");
    if (!space) return;
    $$(".taskbar-window").forEach(e => e.remove());
    Object.values(windows).forEach(w => {
      if (w.el.style.display === "none") return;
      const btn = document.createElement("button");
      btn.className = "taskbar-window";
      btn.textContent = w.cfg.title || "应用";
      btn.onclick = () => {
        w.el.style.display = "flex"; w.el.style.zIndex = ++winZIndex; updateTaskbar();
      };
      space.before(btn);
    });
  }

  // ===== 桌面图标 =====
  function addDesktopIcon(label, icon, onclick) {
    const area = $(".desktop-icons");
    if (!area) return;
    const d = document.createElement("div");
    d.className = "desktop-app-icon";
    d.innerHTML = `<div class="app-icon-glyph" aria-hidden="true">${icon}</div><div class="app-icon-name">${esc(label)}</div>`;
    d.ondblclick = onclick;
    d.onclick = () => { $$(".desktop-app-icon").forEach(x => x.classList.remove("selected")); d.classList.add("selected"); };
    area.appendChild(d);
  }

  // 点击桌面取消选中
  document.addEventListener("click", e => {
    if (!e.target.closest(".desktop-icon")) $$(".desktop-icon").forEach(x => x.classList.remove("active"));
  });

  // ===== 检索工作台 =====
  function openDbSearchApp() {
    const workW = Math.max(1180, Math.min(1540, window.innerWidth - 44));
    const workH = Math.max(720, Math.min(920, window.innerHeight - 74));
    const workX = Math.max(12, Math.round((window.innerWidth - workW) / 2));
    const workY = Math.max(12, Math.round((window.innerHeight - workH) / 2));
    if (windows["dbsearch"]) {
      const w = windows["dbsearch"];
      w.el.style.display = "flex";
      w.el.style.width = workW + "px"; w.el.style.height = workH + "px";
      w.el.style.left = workX + "px"; w.el.style.top = workY + "px";
      w.el.style.zIndex = ++winZIndex; updateTaskbar();
      return;
    }
    const win = createWindow("dbsearch", {
      title: "检索工作台", className: "dbsearch-window",
      x: workX, y: workY, w: workW, h: workH, hideMinimize: true
    });
    win.body.innerHTML = `
      <div class="wk-toolbar">
        <select class="wk-mode"><option value="import">📥 导入检索</option><option value="knowledge">📚 知识库检索</option><option value="cas">🧪 CAS 库检索</option></select>
        <button class="wk-import-btn" id="wk-import-btn" type="button">导入 MSDS</button>
        <input id="wk-import-file" type="file" accept=".docx" hidden>
        <div class="wk-kb-tools" id="wk-kb-tools" style="display:none">
          <select class="kb-inp kb-category" id="kb-category"><option value="">全部大类</option></select>
          <input type="text" class="kb-inp kb-model-query" id="kb-inp" placeholder="搜型号名…">
          <input type="text" class="kb-inp kb-kw" id="kb-kw" placeholder="关键词检索（标签/值）…">
          <button class="kb-toolbar-btn" id="kb-search-btn" type="button">检索</button>
          <button class="kb-toolbar-clear" id="kb-clear" type="button">清除</button>
        </div>
        <div class="wk-cas-tools" id="wk-cas-tools" style="display:none">
          <select class="kb-inp cas-status-filter" id="cas-status-filter">
            <option value="">全部</option>
            <option value="has_cas">有 CAS 号</option>
            <option value="no_cas">无 CAS 号</option>
            <option value="secret">商业机密</option>
          </select>
          <input class="kb-inp" id="cas-query" type="text" placeholder="搜索 CAS号、物质名称、别名、型号…">
          <button class="kb-toolbar-btn" id="cas-search-btn" type="button">检索</button>
          <button class="kb-toolbar-clear" id="cas-clear-btn" type="button">清除</button>
        </div>
      </div>
      <div class="wk-main">
        <div class="wk-left" id="wk-left"></div>
        <div class="wk-split" id="wk-split"></div>
        <div class="wk-right" id="wk-right">
          <div class="wk-empty">选择左侧节点查看内容</div>
        </div>
      </div>`;
    initWorkbench(win.body);
  }

  function openBlankApp(id, title, icon) {
    const workW = Math.max(900, Math.min(1280, window.innerWidth - 120));
    const workH = Math.max(560, Math.min(760, window.innerHeight - 120));
    const workX = Math.max(40, Math.round((window.innerWidth - workW) / 2));
    const workY = Math.max(40, Math.round((window.innerHeight - workH) / 2));
    if (windows[id]) {
      const w = windows[id];
      w.el.style.display = "flex";
      w.el.style.width = workW + "px"; w.el.style.height = workH + "px";
      w.el.style.left = workX + "px"; w.el.style.top = workY + "px";
      w.el.style.zIndex = ++winZIndex; updateTaskbar();
      return;
    }
    const win = createWindow(id, { title, className: "blank-app-window", x: workX, y: workY, w: workW, h: workH });
    win.body.innerHTML = `<div class="gz-empty-app" aria-label="${esc(title)}空白应用">
      <div class="gz-empty-app-icon">${esc(icon)}</div>
      <h2>${esc(title)}</h2>
      <p>空白应用，待配置。</p>
    </div>`;
  }

  // ===== MSDS 覆写工具：完整 Python 表单桥接 =====
  function openOverwriteApp() {
    const workW = Math.max(1040, Math.min(1480, window.innerWidth - 44));
    const workH = Math.max(700, Math.min(930, window.innerHeight - 74));
    const workX = Math.max(12, Math.round((window.innerWidth - workW) / 2));
    const workY = Math.max(12, Math.round((window.innerHeight - workH) / 2));
    if (windows["msds-overwrite"]) {
      const w = windows["msds-overwrite"];
      w.el.style.display = "flex"; w.el.style.width = workW + "px"; w.el.style.height = workH + "px";
      w.el.style.left = workX + "px"; w.el.style.top = workY + "px";
      w.el.style.zIndex = ++winZIndex; updateTaskbar(); return;
    }
    const win = createWindow("msds-overwrite", { title: "MSDS 覆写工具", className: "overwrite-window", x: workX, y: workY, w: workW, h: workH });
    win.body.innerHTML = `<div class="overwrite-app">
      <header class="overwrite-toolbar"><div><span class="overwrite-kicker">MSDS OVERWRITE</span><h2>表单工具 · 多语言 · 多品牌 · 多格式产出</h2></div>
        <label class="overwrite-template">模板<select id="ow-template"></select></label>
        <label class="ow-choice">语言<select id="ow-language"><option value="zh">中文</option><option value="en">English</option></select></label>
        <label class="ow-choice">公司<select id="ow-company"><option value="guanzhi">冠志</option><option value="guocai">国彩</option></select></label>
        <label class="ow-choice">格式<select id="ow-format"><option value="docx">Word</option><option value="pdf">PDF</option></select></label>
        <button class="kb-toolbar-btn" id="ow-load" type="button">加载表单</button>
      </header>
      <div class="overwrite-main"><aside class="overwrite-sections" id="ow-sections"></aside><section class="overwrite-editor" id="ow-editor"><div class="wk-empty">正在加载固定冠志模板…</div></section></div>
      <footer class="overwrite-footer"><input id="ow-output" class="kb-inp" value="" placeholder="输出名称（批量产出会自动追加公司和格式）"><button class="kb-toolbar-btn" id="ow-preview" type="button">预览写入项</button><button class="kb-toolbar-btn ow-primary" id="ow-run" type="button">覆写并产出</button><button class="kb-toolbar-btn ow-primary" id="ow-run-batch" type="button">中文四件套产出</button><button class="kb-toolbar-btn" id="ow-download-batch" type="button" disabled>一键下载产出项</button><span id="ow-status">等待加载</span></footer>
      <div class="overwrite-result" id="ow-result"></div></div>`;
    const $template = win.body.querySelector("#ow-template");
    const $sections = win.body.querySelector("#ow-sections");
    const $editor = win.body.querySelector("#ow-editor");
    const $status = win.body.querySelector("#ow-status");
    const $result = win.body.querySelector("#ow-result");
    const $language = win.body.querySelector("#ow-language");
    const $company = win.body.querySelector("#ow-company");
    const $format = win.body.querySelector("#ow-format");
    const $batchRun = win.body.querySelector("#ow-run-batch");
    const $batchDownload = win.body.querySelector("#ow-download-batch");
    const state = { template: "", fields: [], values: {}, components: [], product_type: "混合物", bio_rows: [], laws: [], s15_special: {}, s16_special: {}, label_overrides: {}, touched_sections: new Set(), cleared_keys: new Set(), section: 0, language: "zh", company: "guanzhi", output_format: "docx" };
    const sectionNames = {0:"页眉页脚",1:"物料及供应商标识",2:"危险性概述",3:"成分/组成信息",4:"急救措施",5:"消防措施",6:"泄漏应急处理",7:"操作处置与储存",8:"接触控制/个体防护",9:"物理和化学性质",10:"稳定性和反应性",11:"毒理学信息",12:"生态学信息",13:"废弃处置",14:"运输信息",15:"法规信息",16:"其他信息"};
    const fieldValue = key => String(state.values[key] ?? "");
    function mark(sec) { state.touched_sections.add(Number(sec)); $status.textContent = `已编辑 S${sec}`; }
    function rowInput(field) {
      const value = fieldValue(field.key); const multi = field.kind === "note" || field.section === 2 || field.section === 4 || field.section === 5 || field.section === 6 || field.section === 7 || field.section === 10 || field.section === 11 || field.section === 12 || field.section === 13 || field.section === 14 || field.section === 15 || field.section === 16;
      const control = multi ? `<textarea data-ow-key="${esc(field.key)}" rows="3">${esc(value)}</textarea>` : `<input data-ow-key="${esc(field.key)}" value="${esc(value)}">`;
      const pictogram = field.label === "GHS象形图" ? renderGhsValue(value) : "";
      return `<label class="ow-field"><span><b>${esc(field.seq || "")}</b> ${esc(field.label)}<small>${esc(field.parent || "")}</small></span><div class="ow-field-control">${control}<button type="button" class="ow-clear" data-ow-clear="${esc(field.key)}">清空</button>${pictogram}</div></label>`;
    }
    function renderSection(sec) {
      state.section = Number(sec); $sections.querySelectorAll("button").forEach(x => x.classList.toggle("active", x.dataset.sec === String(sec)));
      const fields = state.fields.filter(x => Number(x.section) === Number(sec));
      let html = `<div class="ow-section-head"><span>S${sec}</span><div><h3>${esc(sectionNames[sec] || "MSDS 节")}</h3><p>空白字段不进入正式呈现；S11/S12 无可用研究时使用规定兜底行，S9 只呈现有值项。</p></div></div>`;
      if (Number(sec) === 3) {
        html += `<label class="ow-field"><span><b>3.1</b> 产品类型</span><select id="ow-product-type"><option>混合物</option><option>单质/化合物</option><option>未知</option></select></label><div class="ow-special-head"><b>3.2 成分</b><button type="button" id="ow-add-component">＋添加成分</button></div><div id="ow-components"></div>`;
      } else if (Number(sec) === 8) {
        html += fields.map(rowInput).join("");
        html += `<div class="ow-special-head"><b>8.2 生物限值</b><button type="button" id="ow-add-bio">＋添加行</button></div><div id="ow-bio"></div>`;
      } else if (Number(sec) === 15) {
        html += fields.map(rowInput).join("");
        html += `<div class="ow-special-head"><b>S15 其它规定/合规说明</b></div><label class="ow-field"><span><b></b> 其它的规定<small>可编辑正文槽位，结构标题保留</small></span><div class="ow-field-control"><textarea data-ow-s15-special="其它的规定" rows="3">${esc(state.s15_special["其它的规定"] || "")}</textarea></div></label><label class="ow-field"><span><b></b> 符合下列法规要求<small>可编辑正文槽位，结构标题保留</small></span><div class="ow-field-control"><textarea data-ow-s15-special="符合下列法规要求" rows="3">${esc(state.s15_special["符合下列法规要求"] || "")}</textarea></div></label>`;
        html += `<div class="ow-special-head"><b>法规条目</b><button type="button" id="ow-add-law">＋添加法规</button></div><div id="ow-laws"></div>`;
      } else if (Number(sec) === 16) {
        html += fields.map(rowInput).join("") || `<label class="ow-field"><span><b></b> 免责声明<small>正式 S16 免责声明输入槽位</small></span><div class="ow-field-control"><textarea data-ow-s16-special="免责声明" rows="5">${esc(state.s16_special["免责声明"] || "")}</textarea></div></label>`;
      } else html += fields.map(rowInput).join("") || `<div class="wk-empty">本节没有可编辑字段</div>`;
      $editor.innerHTML = html;
      $editor.querySelectorAll("[data-ow-key]").forEach(el => { el.oninput = () => { state.values[el.dataset.owKey] = el.value; const field = fields.find(x => x.key === el.dataset.owKey); if (field?.section === 16) state.s16_special[field.label] = el.value; const preview = el.closest(".ow-field")?.querySelector(".ow-ghs-preview"); if (preview) preview.outerHTML = renderGhsValue(el.value); mark(sec); }; });
      $editor.querySelectorAll("[data-ow-clear]").forEach(el => { el.onclick = () => { const key = el.dataset.owClear; const input = $editor.querySelector(`[data-ow-key="${CSS.escape(key)}"]`); if (input) input.value = ""; state.values[key] = ""; state.cleared_keys.add(key); const preview = el.closest(".ow-field")?.querySelector(".ow-ghs-preview"); if (preview) preview.outerHTML = renderGhsValue(""); mark(sec); }; });
      $editor.querySelectorAll("[data-ow-s15-special]").forEach(el => { el.oninput = () => { state.s15_special[el.dataset.owS15Special] = el.value; mark(15); }; });
      $editor.querySelectorAll("[data-ow-s16-special]").forEach(el => { el.oninput = () => { state.s16_special[el.dataset.owS16Special] = el.value; const field = fields.find(x => x.label === el.dataset.owS16Special); if (field) state.values[field.key] = el.value; mark(16); }; });
      if (Number(sec) === 3) renderComponents();
      if (Number(sec) === 8) { renderBio(); $editor.querySelector("#ow-add-bio").onclick = () => { state.bio_rows.push(["","","","",""]); renderBio(); mark(8); }; }
      if (Number(sec) === 15) { renderLaws(); $editor.querySelector("#ow-add-law").onclick = () => { state.laws.push(""); renderLaws(); mark(15); }; }
    }
    function renderComponents() {
      const box = $editor.querySelector("#ow-components"); if (!box) return;
      const select = $editor.querySelector("#ow-product-type"); if (select) { select.value = state.product_type; select.onchange = () => { state.product_type = select.value; mark(3); }; }
      box.innerHTML = state.components.map((x, i) => `<div class="ow-array-row"><input data-c="name" data-i="${i}" value="${esc(x.name || "")}" placeholder="成分名称"><input data-c="cas" data-i="${i}" value="${esc(x.cas || "")}" placeholder="CAS 编号"><input data-c="conc" data-i="${i}" value="${esc(x.conc || "")}" placeholder="含量"><button type="button" data-del-c="${i}">删除</button></div>`).join("") || `<div class="ow-muted">暂无成分</div>`;
      box.querySelectorAll("[data-c]").forEach(el => { el.oninput = () => { state.components[Number(el.dataset.i)][el.dataset.c] = el.value; mark(3); }; });
      box.querySelectorAll("[data-del-c]").forEach(el => { el.onclick = () => { state.components.splice(Number(el.dataset.delC), 1); renderComponents(); mark(3); }; });
      const add = $editor.querySelector("#ow-add-component"); if (add) add.onclick = () => { state.components.push({name:"",cas:"",conc:""}); renderComponents(); mark(3); };
    }
    function renderBio() { const box = $editor.querySelector("#ow-bio"); if (!box) return; box.innerHTML = state.bio_rows.map((r, i) => `<div class="ow-array-row ow-bio-row">${r.map((v,j) => `<input data-b="${i}" data-j="${j}" value="${esc(v || "")}" placeholder="${["组分名称","标准来源","生物监测指标","生物限值","采样时间"][j]}">`).join("")}<button type="button" data-del-b="${i}">删除</button></div>`).join("") || `<div class="ow-muted">暂无生物限值</div>`; box.querySelectorAll("[data-b]").forEach(el => { el.oninput = () => { state.bio_rows[Number(el.dataset.b)][Number(el.dataset.j)] = el.value; mark(8); }; }); box.querySelectorAll("[data-del-b]").forEach(el => { el.onclick = () => { state.bio_rows.splice(Number(el.dataset.delB),1); renderBio(); mark(8); }; }); }
    function renderLaws() { const box = $editor.querySelector("#ow-laws"); if (!box) return; box.innerHTML = state.laws.map((v,i) => `<div class="ow-array-row"><input data-law="${i}" value="${esc(v)}" placeholder="法规条目"><button type="button" data-del-law="${i}">删除</button></div>`).join("") || `<div class="ow-muted">暂无法规条目</div>`; box.querySelectorAll("[data-law]").forEach(el => { el.oninput = () => { state.laws[Number(el.dataset.law)] = el.value; mark(15); }; }); box.querySelectorAll("[data-del-law]").forEach(el => { el.onclick = () => { state.laws.splice(Number(el.dataset.delLaw),1); renderLaws(); mark(15); }; }); }
    function requestBody() { const model = state.values["S1||Product name"] || state.values["S1||中文名称"] || state.values["S0||产品型号"] || state.values["S0||产品名称"] || ""; return { template_path: state.template, model: model, language: state.language, company: state.company, output_format: state.output_format, form_state: { values: state.values, cleared_keys: Array.from(state.cleared_keys), touched_sections: Array.from(state.touched_sections), components: state.components, product_type: state.product_type, bio_rows: state.bio_rows, laws: state.laws, s15_special: state.s15_special, s16_special: state.s16_special, label_overrides: state.label_overrides } }; }
    async function loadForm() { $status.textContent = "加载模板表单…"; $batchDownload.disabled = true; $batchDownload.onclick = null; const data = await api(API_OLD + "/api/overwrite/form?template=" + encodeURIComponent(state.template) + "&language=" + encodeURIComponent(state.language) + "&company=" + encodeURIComponent(state.company)); if (data.company) state.company = data.company; $company.value = state.company; Object.assign(state, data.form); state.touched_sections = new Set(); state.cleared_keys = new Set(); state.section = 0; $sections.innerHTML = Array.from({length:17}, (_,i) => `<button type="button" data-sec="${i}">S${i}<span>${esc(sectionNames[i])}</span></button>`).join(""); $sections.querySelectorAll("button").forEach(b => b.onclick = () => renderSection(b.dataset.sec)); renderSection(0); $status.textContent = `${state.company === "guocai" ? "国彩" : "冠志"}固定信息已加载`; }
    async function showPreview() { try { $status.textContent = "生成预览…"; const body = requestBody(); const r = await fetch(API_OLD + "/api/overwrite/preview", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)}); const d = await r.json(); if (!r.ok || !d.ok) throw new Error(d.error || "预览失败"); let extra = ""; if (state.language === "en") { const tr = await fetch(API_OLD + "/api/multiformat/translate-preview", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({values: state.values, language: state.language, company: state.company})}).then(x => x.json()); const review = tr.result?.review || []; extra = ` · 样本记忆 ${tr.result?.memory?.sample_count || 0} 组 · 待复核 ${review.length} 项`; } $result.innerHTML = `<div class="ow-result-ok">预览成功：${Object.entries(d.summary).map(([k,v]) => `S${k} ${v}`).join(" · ")}${extra}</div>`; $status.textContent = "预览完成"; } catch(e) { $result.innerHTML = `<div class="ow-result-error">${esc(e.message)}</div>`; $status.textContent = "预览失败"; } }
    async function runOverwrite() { try { $status.textContent = `正在覆写并产出 ${state.output_format.toUpperCase()}…`; const body = requestBody(); body.output_name = win.body.querySelector("#ow-output").value.trim(); const r = await fetch(API_OLD + "/api/overwrite", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)}); const d = await r.json(); if (!r.ok) throw new Error(d.error || "产出失败"); const warn = (d.problems || []).length ? `<div class="ow-result-warn">产出完成，但校验有 ${d.problems.length} 项提示。</div>` : `<div class="ow-result-ok">产出完成，结构闭环校验通过。</div>`; const reviewCount = (d.translation?.review?.length || 0) + (d.translation?.document_review?.length || 0); const review = reviewCount ? `<div class="ow-result-warn">英文内容仍有 ${reviewCount} 项待人工复核，当前为工作稿，不能直接视为正式合规译文。</div>` : ""; $result.innerHTML = `${warn}${review}<div>文件：${esc(d.output_name || d.output)}</div><a class="ow-download" href="${esc(d.download_url || "#")}">下载 ${esc(state.output_format.toUpperCase())}</a><details><summary>查看覆写日志（${(d.logs || []).length} 条）</summary><pre>${esc((d.logs || []).join("\n"))}</pre></details>`; $status.textContent = d.formal_ready ? "产出成功" : (d.ok ? "产出成功，待人工复核" : "产出完成，有校验提示"); } catch(e) { $result.innerHTML = `<div class="ow-result-error">${esc(e.message)}</div>`; $status.textContent = "产出失败"; } }
    async function runBatch() { if (state.language !== "zh") { $result.innerHTML = `<div class="ow-result-warn">中文四件套产出只支持中文表单，请先切换到中文。</div>`; $status.textContent = "请切换到中文"; return; } try { $batchRun.disabled = true; $batchDownload.disabled = true; $status.textContent = "正在产出中文四件套…"; const body = requestBody(); body.output_name = win.body.querySelector("#ow-output").value.trim(); const r = await fetch(API_OLD + "/api/overwrite/batch", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)}); const d = await r.json(); const rows = (d.outputs || []).map(item => item.ok ? `<div class="ow-batch-row ow-result-ok">${esc(item.company === "guocai" ? "国彩" : "冠志")} · ${esc(item.output_format === "pdf" ? "PDF" : "Word")}：<a class="ow-download" href="${esc(item.download_url || "#")}">下载 ${esc(item.output_name || "文件")}</a></div>` : `<div class="ow-batch-row ow-result-error">${esc(item.company === "guocai" ? "国彩" : "冠志")} · ${esc(item.output_format === "pdf" ? "PDF" : "Word")}：${esc(item.error || "产出失败")}</div>`).join(""); const archive = d.download_url ? `<div class="ow-result-ok">四份文件已打包：${esc(d.archive_name || "中文MSDS四件套.zip")}，可点击右侧“一键下载产出项”。</div>` : `<div class="ow-result-warn">未生成完整四件套下载包，请查看各文件状态。</div>`; $result.innerHTML = `${d.ok ? `<div class="ow-result-ok">中文四件套产出完成。</div>` : `<div class="ow-result-warn">中文四件套未全部完成：${esc(d.error || "请查看失败项")}</div>`}${archive}${rows}`; if (d.download_url && d.ok) { $batchDownload.disabled = false; $batchDownload.onclick = () => { const anchor = document.createElement("a"); anchor.href = d.download_url; anchor.download = d.archive_name || "中文MSDS四件套.zip"; document.body.appendChild(anchor); anchor.click(); anchor.remove(); }; } $status.textContent = d.ok ? "四件套产出完成" : "四件套部分失败"; } catch(e) { $result.innerHTML = `<div class="ow-result-error">${esc(e.message)}</div>`; $status.textContent = "四件套产出失败"; } finally { $batchRun.disabled = false; } }
    $language.onchange = () => { state.language = $language.value; const option = Array.from($template.options).find(x => x.dataset.language === state.language); if (option) { state.template = option.value; $template.value = option.value; loadForm().catch(e => { $status.textContent = e.message; }); } else mark(state.section); };
    $company.onchange = () => { state.company = $company.value; loadForm().catch(e => { $status.textContent = e.message; }); };
    $format.onchange = () => { state.output_format = $format.value; };
    $template.onchange = () => { state.template = $template.value; const option = $template.options[$template.selectedIndex]; if (option?.dataset.language) { state.language = option.dataset.language; $language.value = state.language; } loadForm().catch(e => { $status.textContent = e.message; }); };
    win.body.querySelector("#ow-load").onclick = () => loadForm().catch(e => { $status.textContent = e.message; });
    win.body.querySelector("#ow-preview").onclick = showPreview; win.body.querySelector("#ow-run").onclick = runOverwrite; $batchRun.onclick = runBatch;
    api(API_OLD + "/api/overwrite/templates").then(data => { $template.innerHTML = (data.templates || []).map(x => `<option value="${esc(x.path)}" data-language="${esc(x.language || "zh")}">${esc(x.name)}${x.approved ? "（固定批准模板）" : ""}</option>`).join(""); state.template = data.default || $template.value; $template.value = state.template; return loadForm(); }).catch(e => { $status.textContent = e.message || "模板加载失败"; });
  }

  // ===== TDS 覆写工具：全源目录检索 + 中英文/品牌/格式产出 =====
  function openTdsOverwriteApp() {
    const workW = Math.max(1120, Math.min(1540, window.innerWidth - 30));
    const workH = Math.max(720, Math.min(940, window.innerHeight - 48));
    const workX = Math.max(8, Math.round((window.innerWidth - workW) / 2));
    const workY = Math.max(8, Math.round((window.innerHeight - workH) / 2));
    if (windows["tds-overwrite"]) {
      const w = windows["tds-overwrite"];
      w.el.style.display = "flex"; w.el.style.width = workW + "px"; w.el.style.height = workH + "px";
      w.el.style.left = workX + "px"; w.el.style.top = workY + "px";
      w.el.style.zIndex = ++winZIndex; updateTaskbar(); return;
    }
    const win = createWindow("tds-overwrite", { title: "TDS 覆写工具", className: "tds-window", x: workX, y: workY, w: workW, h: workH });
    win.body.innerHTML = `<div class="tds-app">
      <header class="tds-toolbar"><div class="tds-brand"><span class="tds-kicker">TDS OVERWRITE</span><h2>产品 TDS · 统一骨架多格式产出</h2><small>源目录只读检索 · 固定模板写入 · 语言/品牌信息独立路由</small></div>
        <label class="tds-choice">目标语言<select id="tds-language"><option value="zh">中文</option><option value="en">English</option></select></label>
        <label class="tds-choice">目标公司<select id="tds-company"><option value="guanzhi">冠志</option><option value="guocai">国彩</option></select></label>
        <button class="kb-toolbar-btn" id="tds-refresh" type="button">刷新源目录</button>
      </header>
      <div class="tds-main"><aside class="tds-sources"><div class="tds-source-head"><input id="tds-query" class="kb-inp" placeholder="检索型号、文件名、产品类别"><button class="kb-toolbar-btn" id="tds-search" type="button">检索</button></div><div id="tds-catalog-status" class="tds-catalog-status">正在读取 TDS 源目录…</div><div id="tds-catalog" class="tds-catalog"></div></aside>
        <section class="tds-editor-wrap"><div id="tds-editor" class="tds-editor"><div class="wk-empty">从左侧选择一个 TDS 源文件开始</div></div></section>
        <aside class="tds-output"><div class="tds-output-head"><span class="tds-kicker">OUTPUT MATRIX</span><h3>选择产出项</h3><p>中文/英文 × 冠志/国彩 × Word/PDF，共 8 项。</p></div><div id="tds-targets" class="tds-targets"></div><label class="tds-output-name">输出名称<input id="tds-output-name" class="kb-inp" placeholder="默认使用产品型号"></label><div class="tds-output-actions"><button class="kb-toolbar-btn" id="tds-preview" type="button">预览字段</button><button class="kb-toolbar-btn ow-primary" id="tds-generate" type="button">产出选定项</button><button class="kb-toolbar-btn ow-primary" id="tds-batch" type="button">一键产出八件套</button><button class="kb-toolbar-btn" id="tds-download" type="button" disabled>一键下载产出项</button></div><div id="tds-status" class="tds-status">等待选择源文件</div><div id="tds-result" class="tds-result"></div></aside>
      </div></div>`;
    const $language = win.body.querySelector("#tds-language");
    const $company = win.body.querySelector("#tds-company");
    const $query = win.body.querySelector("#tds-query");
    const $catalog = win.body.querySelector("#tds-catalog");
    const $catalogStatus = win.body.querySelector("#tds-catalog-status");
    const $editor = win.body.querySelector("#tds-editor");
    const $targets = win.body.querySelector("#tds-targets");
    const $status = win.body.querySelector("#tds-status");
    const $result = win.body.querySelector("#tds-result");
    const $download = win.body.querySelector("#tds-download");
    const state = { items: [], selected: null, form: null, language: "zh", company: "guanzhi", archive: "" };
    const targetDefs = [
      ["zh", "guanzhi", "docx", "中文 · 冠志 · Word"], ["zh", "guanzhi", "pdf", "中文 · 冠志 · PDF"],
      ["zh", "guocai", "docx", "中文 · 国彩 · Word"], ["zh", "guocai", "pdf", "中文 · 国彩 · PDF"],
      ["en", "guanzhi", "docx", "English · Guanzhi · Word"], ["en", "guanzhi", "pdf", "English · Guanzhi · PDF"],
      ["en", "guocai", "docx", "English · Guocai · Word"], ["en", "guocai", "pdf", "English · Guocai · PDF"]
    ];
    $targets.innerHTML = targetDefs.map((x, i) => `<label class="tds-target"><input type="checkbox" data-tds-target="${i}" checked><span>${esc(x[3])}</span></label>`).join("");
    function targetValues(index) { const x = targetDefs[index]; return { language: x[0], company: x[1], output_format: x[2], label: x[3] }; }
    function renderCatalog() {
      const rows = state.items || [];
      $catalog.innerHTML = rows.map(item => `<button type="button" class="tds-source-row ${state.selected?.source_id === item.source_id ? "active" : ""}" data-tds-id="${esc(item.source_id)}"><span class="tds-source-model">${esc(item.model || "未识别型号")}</span><span class="tds-source-name">${esc(item.name)}</span><small>${esc(item.category)} · ${esc(item.language === "en" ? "EN" : item.language === "zh" ? "CN" : "未知")} · ${esc(item.company === "guocai" ? "国彩" : item.company === "guanzhi" ? "冠志" : "品牌未知")}${item.counterparts ? ` · 配对 ${item.counterparts}` : ""}</small></button>`).join("") || `<div class="wk-empty">未找到 TDS 文件</div>`;
      $catalog.querySelectorAll("[data-tds-id]").forEach(el => el.onclick = () => loadSource(el.dataset.tdsId));
    }
    function renderEditor() {
      const f = state.form;
      if (!f) { $editor.innerHTML = `<div class="wk-empty">从左侧选择一个 TDS 源文件开始</div>`; return; }
      const text = key => esc(f[key] || "");
      const list = key => (f[key] || []).map(x => String(x || ""));
      $editor.innerHTML = `<div class="tds-editor-head"><div><span class="tds-kicker">SOURCE FORM</span><h3>${esc(f.model || "TDS")} · ${esc(f.source_name || "")}</h3><p>${esc(f.category || "")} · 当前目标：${state.language === "en" ? "English" : "中文"} / ${state.company === "guocai" ? "国彩" : "冠志"} · 配对源：${esc(f.paired_source_name || "未使用")}</p></div><span class="tds-readonly-badge">源文件只读</span></div>
        <div class="tds-form-grid"><label><span>型号</span><input data-tds-field="model" value="${text("model")}"></label><label><span>产品名称/标题</span><input data-tds-field="product_name" value="${text("product_name")}"></label>
        <label class="full"><span>产品描述</span><textarea data-tds-field="description" rows="4">${text("description")}</textarea></label>
        <label><span>供货形式</span><textarea data-tds-field="supply" rows="2">${text("supply")}</textarea></label><label><span>包装</span><textarea data-tds-field="packaging" rows="2">${text("packaging")}</textarea></label>
        <label class="full"><span>注意事项</span><textarea data-tds-field="precaution" rows="2">${text("precaution")}</textarea></label>
        <label class="full"><span>储存</span><textarea data-tds-field="storage" rows="3">${text("storage")}</textarea></label></div>
        <div class="tds-array-block"><div class="tds-array-head"><h4>性能指标</h4><button type="button" class="kb-toolbar-btn" id="tds-add-performance">＋添加行</button></div><div id="tds-performance"></div></div>
        <div class="tds-array-block"><div class="tds-array-head"><h4>产品特性</h4><button type="button" class="kb-toolbar-btn" data-tds-add-list="features">＋添加行</button></div><div id="tds-features"></div></div>
        <div class="tds-array-block"><div class="tds-array-head"><h4>应用</h4><button type="button" class="kb-toolbar-btn" data-tds-add-list="application">＋添加行</button></div><div id="tds-application"></div></div>`;
      $editor.querySelectorAll("[data-tds-field]").forEach(el => el.oninput = () => { f[el.dataset.tdsField] = el.value; });
      function renderPerformance() { const box = $editor.querySelector("#tds-performance"); box.innerHTML = (f.performance || []).map((row, i) => `<div class="tds-performance-row"><input data-tds-p="item" data-i="${i}" value="${esc(row.item || "")}" placeholder="项目"><input data-tds-p="index" data-i="${i}" value="${esc(row.index || "")}" placeholder="指标"><input data-tds-p="unit" data-i="${i}" value="${esc(row.unit || "")}" placeholder="单位"><input data-tds-p="method" data-i="${i}" value="${esc(row.method || "")}" placeholder="测试方法"><button type="button" data-tds-del-p="${i}">删</button></div>`).join("") || `<div class="tds-muted">暂无性能指标</div>`; box.querySelectorAll("[data-tds-p]").forEach(el => el.oninput = () => { f.performance[Number(el.dataset.i)][el.dataset.tdsP] = el.value; }); box.querySelectorAll("[data-tds-del-p]").forEach(el => el.onclick = () => { f.performance.splice(Number(el.dataset.tdsDelP), 1); renderPerformance(); }); }
      function renderList(key) { const box = $editor.querySelector(`#tds-${key}`); box.innerHTML = list(key).map((value, i) => `<div class="tds-list-row"><textarea data-tds-list="${key}" data-i="${i}" rows="2">${esc(value)}</textarea><button type="button" data-tds-del-list="${key}:${i}">删</button></div>`).join("") || `<div class="tds-muted">暂无内容</div>`; box.querySelectorAll("[data-tds-list]").forEach(el => el.oninput = () => { f[key][Number(el.dataset.i)] = el.value; }); box.querySelectorAll("[data-tds-del-list]").forEach(el => el.onclick = () => { const [kind, i] = el.dataset.tdsDelList.split(":"); f[kind].splice(Number(i), 1); renderList(kind); }); }
      $editor.querySelector("#tds-add-performance").onclick = () => { f.performance = f.performance || []; f.performance.push({ item: "", index: "", unit: "", method: "" }); renderPerformance(); };
      $editor.querySelectorAll("[data-tds-add-list]").forEach(el => el.onclick = () => { const key = el.dataset.tdsAddList; f[key] = f[key] || []; f[key].push(""); renderList(key); });
      renderPerformance(); renderList("features"); renderList("application");
    }
    async function loadCatalog() { try { $catalogStatus.textContent = "正在读取 TDS 源目录…"; const q = encodeURIComponent($query.value.trim()); const data = await api(API_OLD + "/api/tds/catalog?q=" + q + "&limit=1000"); state.items = data.items || []; $catalogStatus.textContent = `源文件 ${data.total || 0} 个 · 当前显示 ${state.items.length} 个 · 目录只读`; renderCatalog(); } catch (e) { $catalogStatus.textContent = e.message || "源目录读取失败"; $catalog.innerHTML = `<div class="tds-error">${esc(e.message || "源目录读取失败")}</div>`; } }
    async function loadSource(id) { try { $status.textContent = "正在读取选中 TDS…"; const data = await api(API_OLD + "/api/tds/source?id=" + encodeURIComponent(id) + "&language=" + encodeURIComponent(state.language) + "&company=" + encodeURIComponent(state.company)); state.selected = state.items.find(x => x.source_id === id) || { source_id: id }; state.form = data.form; renderCatalog(); renderEditor(); const review = (state.form.translation_review || []).length; $status.textContent = `已加载 ${state.form.model || state.form.source_name || "TDS"}${review ? ` · ${review} 项待复核` : ""}`; } catch (e) { $status.textContent = e.message || "TDS 读取失败"; $result.innerHTML = `<div class="tds-error">${esc(e.message || "TDS 读取失败")}</div>`; } }
    function requestBody(target) { return { source_id: state.selected?.source_id || "", language: target.language, form_language: state.language, company: target.company, output_format: target.output_format, output_name: win.body.querySelector("#tds-output-name").value.trim(), form: state.form }; }
    function renderResult(items, archive) { const rows = (items || []).map(item => item.ok ? `<div class="tds-result-row ok">${esc(item.language === "en" ? "English" : "中文")} · ${esc(item.company === "guocai" ? "国彩" : "冠志")} · ${esc(item.output_format === "pdf" ? "PDF" : "Word")}：<a class="ow-download" href="${esc(item.download_url || "#")}">下载</a></div>` : `<div class="tds-result-row error">${esc(item.language || "")} · ${esc(item.company || "")} · ${esc(item.output_format || "")}：${esc(item.error || "产出失败")}</div>`).join(""); $result.innerHTML = `${archive ? `<div class="ow-result-ok">八件套已完成：${esc(archive.archive_name)}，可点击“一键下载产出项”。</div>` : ""}${rows || `<div class="tds-muted">等待产出</div>`}`; if (archive?.download_url) { state.archive = archive.download_url; $download.disabled = false; $download.onclick = () => { const a = document.createElement("a"); a.href = state.archive; a.download = archive.archive_name || "TDS八件套.zip"; document.body.appendChild(a); a.click(); a.remove(); }; } else { state.archive = ""; $download.disabled = true; $download.onclick = null; } }
    async function preview() { if (!state.selected || !state.form) { $status.textContent = "请先选择 TDS 源文件"; return; } try { $status.textContent = "正在生成字段预览…"; const r = await fetch(API_OLD + "/api/tds/preview", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({source_id: state.selected.source_id, language: state.language, company: state.company, form: state.form})}); const d = await r.json(); if (!r.ok || !d.ok) throw new Error(d.error || "预览失败"); const n = (d.translation_review || []).length; $result.innerHTML = `<div class="ow-result-ok">字段预览成功 · 固定模板：${esc(d.template || "")} · 待复核 ${n} 项</div>`; $status.textContent = n ? "预览完成，存在待复核翻译" : "预览完成"; } catch (e) { $result.innerHTML = `<div class="tds-error">${esc(e.message || "预览失败")}</div>`; $status.textContent = "预览失败"; } }
    async function generateSelected() { if (!state.selected || !state.form) { $status.textContent = "请先选择 TDS 源文件"; return; } const indexes = Array.from($targets.querySelectorAll("[data-tds-target]:checked")).map(x => Number(x.dataset.tdsTarget)); if (!indexes.length) { $status.textContent = "请至少选择一个产出项"; return; } try { $status.textContent = `正在产出 ${indexes.length} 个文件…`; $download.disabled = true; const outputs = []; for (const index of indexes) { const target = targetValues(index); const response = await fetch(API_OLD + "/api/tds/generate", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(requestBody(target))}); const data = await response.json(); outputs.push(response.ok ? data : {ok:false, ...target, error:data.error || "产出失败"}); } renderResult(outputs, null); $status.textContent = outputs.every(x => x.ok) ? `已完成 ${outputs.length} 个文件` : "部分产出失败"; } catch (e) { $result.innerHTML = `<div class="tds-error">${esc(e.message || "产出失败")}</div>`; $status.textContent = "产出失败"; } }
    async function generateBatch() { if (!state.selected || !state.form) { $status.textContent = "请先选择 TDS 源文件"; return; } try { $status.textContent = "正在产出中文/英文 × 冠志/国彩 × Word/PDF 八件套…"; $download.disabled = true; const body = {source_id: state.selected.source_id, form_language: state.language, output_name: win.body.querySelector("#tds-output-name").value.trim(), form: state.form}; const response = await fetch(API_OLD + "/api/tds/batch", {method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(body)}); const data = await response.json(); if (!response.ok) throw new Error(data.error || "八件套产出失败"); renderResult(data.outputs || [], data); $status.textContent = data.ok ? "八件套产出完成" : (data.error || "八件套部分失败"); } catch (e) { $result.innerHTML = `<div class="tds-error">${esc(e.message || "八件套产出失败")}</div>`; $status.textContent = "八件套产出失败"; } }
    $language.onchange = () => { state.language = $language.value; if (state.selected) loadSource(state.selected.source_id); };
    $company.onchange = () => { state.company = $company.value; if (state.selected) loadSource(state.selected.source_id); };
    win.body.querySelector("#tds-search").onclick = loadCatalog; win.body.querySelector("#tds-refresh").onclick = loadCatalog; $query.onkeydown = e => { if (e.key === "Enter") loadCatalog(); };
    win.body.querySelector("#tds-preview").onclick = preview; win.body.querySelector("#tds-generate").onclick = generateSelected; win.body.querySelector("#tds-batch").onclick = generateBatch;
    loadCatalog();
  }

  // ===== CAS 联网搜索 =====
  function openCasOnlineApp() {
    const workW = Math.max(980, Math.min(1320, window.innerWidth - 100));
    const workH = Math.max(600, Math.min(820, window.innerHeight - 100));
    const workX = Math.max(30, Math.round((window.innerWidth - workW) / 2));
    const workY = Math.max(30, Math.round((window.innerHeight - workH) / 2));
    if (windows["casonline"]) {
      const w = windows["casonline"];
      w.el.style.display = "flex";
      w.el.style.width = workW + "px"; w.el.style.height = workH + "px";
      w.el.style.left = workX + "px"; w.el.style.top = workY + "px";
      w.el.style.zIndex = ++winZIndex; updateTaskbar();
      return;
    }
    const win = createWindow("casonline", { title: "CAS 联网搜索", className: "cas-online-window", x: workX, y: workY, w: workW, h: workH });
    win.body.innerHTML = `<div class="cas-online-app">
      <form class="cas-online-toolbar" id="cas-online-form">
        <span class="cas-online-kicker">CAS API</span>
        <input id="cas-online-query" class="kb-inp" type="text" placeholder="输入 CAS 号或单个物质名称，例如 872-50-4 / N-methyl-2-pyrrolidone" autocomplete="off">
        <button class="kb-toolbar-btn" type="submit">联网搜索</button>
      </form>
      <div id="cas-online-status" class="cas-online-status">输入 CAS 号或物质名称开始搜索</div>
      <div id="cas-online-result" class="cas-online-result"><div class="wk-empty">等待查询</div></div>
    </div>`;
    const form = win.body.querySelector("#cas-online-form");
    const query = win.body.querySelector("#cas-online-query");
    const status = win.body.querySelector("#cas-online-status");
    const resultBox = win.body.querySelector("#cas-online-result");

    function resultTable(item) {
      const rows = [
        ["PubChem CID", item.cid], ["CAS 候选", (item.cas_candidates || []).join("；") || "未返回"],
        ["分子式", item.molecular_formula], ["分子量", item.molecular_weight],
        ["Canonical SMILES", item.canonical_smiles], ["Isomeric SMILES", item.isomeric_smiles],
        ["InChIKey", item.inchikey], ["同义名", (item.synonyms || []).join("；") || "未返回"]
      ];
      return `<section class="cas-online-card"><h3>${esc(item.title || "未命名物质")}</h3>
        <table class="cas-online-table"><tbody>${rows.map(row => `<tr><th>${esc(row[0])}</th><td>${esc(row[1] || "未返回")}</td></tr>`).join("")}
        <tr><th>来源链接</th><td><a href="${esc(item.source_url || "#")}" target="_blank" rel="noreferrer">${esc(item.source_url || "未返回")}</a></td></tr>
        </tbody></table></section>`;
    }

    form.onsubmit = async e => {
      e.preventDefault();
      const value = query.value.trim();
      if (!value) { status.textContent = "请输入 CAS 号或物质名称"; return; }
      status.textContent = "联网查询中…"; resultBox.innerHTML = "<div class='wk-empty'>正在查询 PubChem / Wikidata</div>";
      try {
        const response = await fetch(API_OLD + "/api/cas-online?q=" + encodeURIComponent(value));
        const data = await response.json();
        if (!response.ok || !data.ok) throw new Error(data.error || "联网查询失败");
        const result = data.result || {};
        status.textContent = `查询：${result.query} · 类型：${result.query_type} · 来源：${result.source} · 结果 ${result.results.length} 条`;
        resultBox.innerHTML = `<div class="cas-online-meta"><div><b>来源提示：</b>${esc(result.source_note || "")}</div>${result.name_resolution ? `<div><b>名称解析：</b>${esc(JSON.stringify(result.name_resolution, null, 2))}</div>` : ""}</div>${(result.results || []).map(resultTable).join("") || "<div class='wk-empty'>没有返回结果</div>"}`;
      } catch (error) {
        status.textContent = "查询失败";
        resultBox.innerHTML = `<div class="cas-online-error">${esc(error.message || "联网查询失败")}</div>`;
      }
    };
  }

  // ===== 合规性判断（REACH SVHC CAS 判断） =====
  function openComplianceApp() {
    const workW = Math.max(980, Math.min(1380, window.innerWidth - 70));
    const workH = Math.max(620, Math.min(860, window.innerHeight - 80));
    const workX = Math.max(20, Math.round((window.innerWidth - workW) / 2));
    const workY = Math.max(20, Math.round((window.innerHeight - workH) / 2));
    if (windows["compliance"]) {
      const w = windows["compliance"];
      w.el.style.display = "flex";
      w.el.style.width = workW + "px"; w.el.style.height = workH + "px";
      w.el.style.left = workX + "px"; w.el.style.top = workY + "px";
      w.el.style.zIndex = ++winZIndex; updateTaskbar();
      return;
    }
    const win = createWindow("compliance", {
      title: "合规性判断", className: "compliance-window",
      x: workX, y: workY, w: workW, h: workH
    });
    win.body.innerHTML = `<div class="regulatory-shell" aria-label="合规性判断与法律法规物质清单">
      <aside class="regulatory-sidebar">
        <div class="regulatory-sidebar-head"><span class="regulatory-mark">GZ</span><div><b>法规判断</b><small>Compliance Desk</small></div></div>
        <nav class="regulatory-nav" aria-label="法规功能导航">
          <button class="regulatory-nav-item is-active" type="button" data-reg-page="compliance"><span>⚖</span><b>合规性检测</b></button>
          <button class="regulatory-nav-item regulatory-nav-parent" type="button" data-reg-toggle="laws"><span>▤</span><b>法律法规物质清单</b><i>▾</i></button>
          <div class="regulatory-subnav is-open" data-reg-subnav="laws">
            <button class="regulatory-subitem is-active" type="button" data-reg-page="reach-list"><span>REACH</span><b>SVHC 253项</b></button>
            <button class="regulatory-subitem" type="button" data-reg-page="annex-xvii-list"><span>REACH</span><b>Annex XVII</b></button>
            <button class="regulatory-subitem" type="button" data-reg-page="rohs-list"><span>RoHS</span><b>限用物质</b></button>
            <button class="regulatory-subitem" type="button" data-reg-page="hsf-001-list"><span>HSF</span><b>HSF 001</b></button>
          </div>
        </nav>
        <div class="regulatory-sidebar-foot">只读法规数据 · SVHC 2026-02-04 · Annex XVII ECHA CHEM · RoHS 2011/65/EU + 2015/863 · HSF 001 · 飞书 Wiki 全量明细</div>
      </aside>
      <main class="regulatory-main" id="regulatory-main"></main>
    </div>`;

    const main = win.body.querySelector("#regulatory-main");
    const nav = win.body.querySelector(".regulatory-nav");

    function setActive(page) {
      win.body.querySelectorAll("[data-reg-page]").forEach(item => item.classList.toggle("is-active", item.dataset.regPage === page));
    }

    function renderPending(items) {
      const dataRows = (items || []).map(item => `<tr><td>${esc(item["中文名称"] || "（空）")}</td><td>${esc(item["英文名称"] || "（空）")}</td></tr>`).join("");
      return `<section class="compliance-table-card compliance-pending-card is-collapsed"><table class="compliance-result-table pending-result-table"><thead>
          <tr><th colspan="2"><button class="reach-section-title" type="button" aria-expanded="false"><span>无CAS待审查物质</span><span class="reach-section-tools"><b>${items.length}</b><i aria-hidden="true">▸</i></span></button></th></tr>
          <tr class="pending-columns"><th>中文名称</th><th>英文名称</th></tr></thead><tbody><tr class="pending-hint"><td colspan="2">以下物质没有 CAS 号，不能通过单一 CAS 直接完成判断。</td></tr>${dataRows || "<tr><td colspan=\"2\">（无）</td></tr>"}</tbody></table></section>`;
    }

    function renderAnnexPending(items) {
      const rows = (items || []).map(item => `<tr><td>${esc(item["条目号"] || "（空）")}</td><td>${esc(item["中文名称"] || "（空）")}</td><td>${esc(item["英文名称"] || "（空）")}</td><td>${item["限制条件链接"] ? `<a class="regulation-source-link" href="${esc(item["限制条件链接"])}" target="_blank" rel="noopener">查看条件</a>` : "（无链接）"}</td></tr>`).join("");
      return `<section class="compliance-table-card compliance-pending-card is-collapsed annex-pending-card"><table class="compliance-result-table pending-result-table"><thead>
          <tr><th colspan="4"><button class="reach-section-title" type="button" aria-expanded="false"><span>Annex XVII 无 CAS/EC 待审查条目</span><span class="reach-section-tools"><b>${(items || []).length}</b><i aria-hidden="true">▸</i></span></button></th></tr>
          <tr class="pending-columns"><th>条目号</th><th>中文名称</th><th>英文名称</th><th>限制条件</th></tr></thead><tbody><tr class="pending-hint"><td colspan="4">这些条目没有 CAS 号或 EC 号，不能通过单一编号直接完成物质身份判断，需要结合用途、组成、浓度和物质组成员继续审查。</td></tr>${rows || "<tr><td colspan=\"4\">（无）</td></tr>"}</tbody></table></section>`;
    }

    function renderComplianceTable(title, headers, records, renderRow, matched, query, emptyText, extraClass = "") {
      const columns = headers.length;
      const rows = matched ? records.map(renderRow).join("") : `<tr><td colspan="${columns}">${esc(emptyText)}：标识 ${esc(query)}</td></tr>`;
      return `<section class="compliance-table-card ${matched ? "is-warning" : "is-neutral"} ${extraClass}"><table class="compliance-result-table"><thead>
        <tr><th colspan="${columns}">${title}</th></tr><tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr></thead><tbody>${rows}</tbody></table></section>`;
    }

    function renderMatch(result) {
      return renderComplianceTable(result.matched ? "⚠ REACH SVHC 警告" : "REACH SVHC 检测通过", ["中文名称", "英文名称", "物质描述", "EC号", "CAS号"], result.records, item => `<tr><td>${esc(item["中文名称"] || "（空）")}</td><td>${esc(item["英文名称"] || "（空）")}</td><td>${esc(item["物质描述"] || "（空）")}</td><td>${esc(item["EC号"] || "（空）")}</td><td>${esc(item["CAS号"] || "（空）")}</td></tr>`, result.matched, result.query, "REACH SVHC 检测通过", "reach-result-table");
    }

    function renderRohsMatch(result) {
      return renderComplianceTable(result.matched ? "⚠ RoHS 限制物质警告" : "RoHS 检测通过", ["中文名称", "英文名称", "CAS号", "EC号", "均质材料限值", "限值ppm", "筛查说明", "法规依据"], result.records, item => `<tr><td>${esc(item["中文名称"] || "（空）")}</td><td>${esc(item["英文名称"] || "（空）")}</td><td>${esc(item["CAS号"] || "（空）")}</td><td>${esc(item["EC号"] || "（空）")}</td><td>${esc(item["均质材料限值"] || "（空）")}</td><td>${esc(item["限值ppm"] || "（空）")}</td><td>${esc(item["材料筛查说明"] || "（空）")}</td><td>${esc(item["法规依据"] || "（空）")}</td></tr>`, result.matched, result.query, "RoHS 检测通过", "rohs-result-table");
    }

    function renderHsf001Match(result) {
      const title = result.matched ? "⚠ HSF 001 有害物质警告" : (result.status || "HSF 001 检测通过");
      return renderComplianceTable(title, ["类别", "具体相关物质", "CAS号", "源表CAS号", "物质级判定", "CAS映射状态", "源表行号", "法规依据", "备注"], result.records, item => `<tr><td>${esc(item["类别"] || "（空）")}</td><td>${esc(item["中文名称"] || "（空）")}</td><td class="reach-id mono">${esc(item["CAS号"] || "（空）")}</td><td class="reach-id mono">${esc(item["源表CAS号"] || "（空）")}</td><td>${esc(item["物质级判定"] || "（空）")}</td><td>${esc(item["CAS映射状态"] || "（空）")}</td><td>${esc(item["源表行号"] || "（空）")}</td><td>${esc(item["法规依据"] || "（空）")}</td><td>${esc(item["备注"] || "（空）")}</td></tr>`, result.matched, result.query, result.status || "HSF 001 当前CAS未命中", "hsf-001-result-table");
    }

    function renderAnnexXviiMatch(result) {
      return renderComplianceTable(result.matched ? "⚠ REACH Annex XVII 限制警告" : "REACH Annex XVII 检测通过", ["条目号", "中文名称", "英文名称", "CAS号", "EC号", "物质描述", "限制条件", "法规依据"], result.records, item => `<tr><td>${esc(item["条目号"] || "（空）")}</td><td>${esc(item["中文名称"] || "（空）")}</td><td>${esc(item["英文名称"] || "（空）")}</td><td>${esc(item["CAS号"] || "（空）")}</td><td>${esc(item["EC号"] || "（空）")}</td><td>${esc(item["物质描述"] || "（空）")}</td><td>${esc(item["限制条件标题"] || "（空）")}<br><a class="regulation-source-link" href="${esc(item["限制条件链接"] || "#")}" target="_blank" rel="noopener">查看官方条件</a></td><td>${esc(item["法规依据"] || "（空）")}</td></tr>`, result.matched, result.query, "REACH Annex XVII 检测通过", "annex-xvii-result-table");
    }

    function showCompliance() {
      setActive("compliance");
      main.innerHTML = `<section class="regulatory-page reach-check-page"><header class="regulatory-page-head"><div><span class="regulatory-eyebrow">COMPLIANCE CHECK</span><h2>合规性检测</h2><p>输入一个 CAS 号或 EC 号，同时检测 REACH SVHC、Annex XVII、RoHS 与 HSF 001 限制物质清单。</p></div><span class="regulatory-page-chip">SVHC + XVII + RoHS + HSF</span></header>
        <form class="reach-check-toolbar" id="reach-check-form"><span class="reach-check-kicker">CAS / EC</span><input id="reach-check-cas" class="kb-inp" type="text" inputmode="numeric" placeholder="输入 CAS 号或 EC 号，例如 110-71-4 / 201-963-1" autocomplete="off"><button class="kb-toolbar-btn" type="submit">判断</button></form>
        <div id="reach-check-status" class="reach-check-status">输入 CAS 号或 EC 号后同时进行四套清单判断；HSF 001 使用 CAS 号</div><div class="reach-check-content"><div id="reach-check-result" class="reach-check-result"><div class="wk-empty">等待输入 CAS 号或 EC 号</div></div><aside id="reach-check-pending" class="reach-check-pending"><div class="wk-empty">查询后显示无 CAS 待审查物质</div></aside></div></section>`;
      const form = main.querySelector("#reach-check-form"); const input = main.querySelector("#reach-check-cas"); const status = main.querySelector("#reach-check-status"); const resultBox = main.querySelector("#reach-check-result"); const pendingBox = main.querySelector("#reach-check-pending"); const casPattern = /^\d{2,7}-\d{2,7}-\d$/;
      pendingBox.addEventListener("click", e => { const toggle = e.target.closest(".reach-section-title"); if (!toggle) return; const card = toggle.closest(".compliance-pending-card"); const collapsed = card.classList.toggle("is-collapsed"); toggle.setAttribute("aria-expanded", String(!collapsed)); const arrow = toggle.querySelector("i"); if (arrow) arrow.textContent = collapsed ? "▸" : "▾"; });
      form.onsubmit = async e => { e.preventDefault(); const cas = input.value.trim(); if (!casPattern.test(cas)) { status.textContent = "请输入格式正确的 CAS 号或 EC 号，例如 110-71-4 / 201-963-1"; resultBox.innerHTML = `<div class="reach-check-error">CAS 号或 EC 号格式不正确</div>`; pendingBox.innerHTML = renderPending([]); return; }
        status.textContent = "正在同时读取 REACH、RoHS 与 HSF 001 数据库…"; resultBox.innerHTML = "<div class='wk-empty'>判断中…</div>"; pendingBox.innerHTML = "<div class='wk-empty'>加载无 CAS 待审查物质…</div>";
        try { const response = await fetch(API_OLD + "/api/compliance/check?identifier=" + encodeURIComponent(cas)); const data = await response.json(); if (!response.ok || !data.ok) throw new Error(data.error || "综合判断失败"); const result = data.result; status.textContent = `标识 ${result.query} · SVHC ${result.reach.matched ? "命中" : "检测通过"} · Annex XVII ${result.annex_xvii.matched ? "命中" : "检测通过"} · RoHS ${result.rohs.matched ? "命中" : "检测通过"} · HSF 001 ${result.hsf_001.matched ? "命中" : (result.hsf_001.status || "检测通过")} · 无 CAS 待审查 ${result.reach.no_cas_count} 条 · Annex XVII 无 CAS/EC 待审查 ${result.annex_xvii.no_identifier_count || 0} 条`; resultBox.innerHTML = `<div class="compliance-result-grid">${renderMatch(result.reach)}${renderAnnexXviiMatch(result.annex_xvii)}${renderRohsMatch(result.rohs)}${renderHsf001Match(result.hsf_001)}</div>`; pendingBox.innerHTML = renderPending(result.reach.no_cas_pending || []) + renderAnnexPending(result.annex_xvii.no_identifier_entries || []); } catch (error) { status.textContent = "判断失败"; resultBox.innerHTML = `<div class="reach-check-error">${esc(error.message || "综合判断失败")}</div>`; pendingBox.innerHTML = renderPending([]) + renderAnnexPending([]); }
      };
    }

    function showReachList() {
      setActive("reach-list");
      main.innerHTML = `<section class="regulatory-page reach-list-page"><header class="regulatory-page-head"><div><span class="regulatory-eyebrow">LEGAL MATERIALS REGISTER</span><h2>REACH SVHC 253项</h2><p>完整查看候选清单中的中文名称、英文名称、CAS号、EC号和物质描述。</p></div><span class="regulatory-page-chip">253 ENTRIES</span></header><div id="reach-list-status" class="reach-list-status">正在读取清单…</div><div class="reach-table-wrap" id="reach-table-wrap"><div class="wk-empty">加载中…</div></div></section>`;
      const status = main.querySelector("#reach-list-status"); const box = main.querySelector("#reach-table-wrap");
      fetch(API_OLD + "/api/reach/list").then(r => r.json().then(data => ({ ok: r.ok, data }))).then(({ ok, data }) => { if (!ok || !data.ok) throw new Error(data.error || "清单读取失败"); const result = data.result; status.textContent = `共 ${result.count} 项 · 5 个物质字段 · 只读查看`; const rows = result.rows.map(row => `<tr><td>${esc(row["中文名称"])}</td><td class="reach-en">${esc(row["英文名称"])}</td><td class="reach-id mono">${esc(row["CAS号"])}</td><td class="reach-id mono">${esc(row["EC号"])}</td><td>${esc(row["物质描述"])}</td></tr>`).join(""); box.innerHTML = `<table class="reach-register-table"><thead><tr><th>中文名称</th><th>英文名称</th><th>CAS号</th><th>EC号</th><th>物质描述</th></tr></thead><tbody>${rows}</tbody></table>`; }).catch(error => { status.textContent = "清单读取失败"; box.innerHTML = `<div class="reach-check-error">${esc(error.message || "清单读取失败")}</div>`; });
    }

    function showRohsList() {
      setActive("rohs-list");
      main.innerHTML = `<section class="regulatory-page reach-list-page"><header class="regulatory-page-head"><div><span class="regulatory-eyebrow">LEGAL MATERIALS REGISTER</span><h2>RoHS 限用物质</h2><p>EU RoHS 2011/65/EU Annex II 与 2015/863 修订的 10 项材料级限制物质。</p></div><span class="regulatory-page-chip">10 ENTRIES</span></header><div id="reach-list-status" class="reach-list-status">正在读取清单…</div><div class="reach-table-wrap" id="reach-table-wrap"><div class="wk-empty">加载中…</div></div></section>`;
      const status = main.querySelector("#reach-list-status"); const box = main.querySelector("#reach-table-wrap");
      fetch(API_OLD + "/api/rohs/list").then(r => r.json().then(data => ({ ok: r.ok, data }))).then(({ ok, data }) => { if (!ok || !data.ok) throw new Error(data.error || "清单读取失败"); const result = data.result; status.textContent = `共 ${result.count} 项 · 含 CAS/EC、均质材料限值与筛查说明 · 只读查看`; const rows = result.rows.map(row => `<tr><td>${esc(row["物质类别"])}</td><td>${esc(row["中文名称"])}</td><td class="reach-en">${esc(row["英文名称"])}</td><td class="reach-id mono">${esc(row["CAS号"])}</td><td class="reach-id mono">${esc(row["EC号"])}</td><td>${esc(row["均质材料限值"])}<br><small>${esc(row["限值ppm"])} ppm</small></td><td>${esc(row["材料筛查说明"])}</td><td>${esc(row["法规依据"])}</td></tr>`).join(""); box.innerHTML = `<table class="reach-register-table rohs-register-table"><thead><tr><th>物质类别</th><th>中文名称</th><th>英文名称</th><th>CAS号</th><th>EC号</th><th>均质材料限值</th><th>材料筛查说明</th><th>法规依据</th></tr></thead><tbody>${rows}</tbody></table>`; }).catch(error => { status.textContent = "清单读取失败"; box.innerHTML = `<div class="reach-check-error">${esc(error.message || "清单读取失败")}</div>`; });
    }

    function showHsf001List() {
      setActive("hsf-001-list");
      main.innerHTML = `<section class="regulatory-page reach-list-page"><header class="regulatory-page-head"><div><span class="regulatory-eyebrow">LEGAL MATERIALS REGISTER</span><h2>HSF 001 有害物质清单</h2><p>Inventec《无有害物质（HSF）管理规范》HSF 001；来源为飞书 Wiki 内嵌 Sheet，全量展示限制项目、具体相关物质、CAS 和源表行号。</p></div><span class="regulatory-page-chip">HSF 001 · FEISHU</span></header><div id="reach-list-status" class="reach-list-status">正在读取清单…</div><div class="regulatory-source-note" id="hsf-001-source-note"></div><div class="reach-table-wrap" id="reach-table-wrap"><div class="wk-empty">加载中…</div></div></section>`;
      const status = main.querySelector("#reach-list-status"); const sourceNote = main.querySelector("#hsf-001-source-note"); const box = main.querySelector("#reach-table-wrap");
      fetch(API_OLD + "/api/hsf-001/list").then(r => r.json().then(data => ({ ok: r.ok, data }))).then(({ ok, data }) => { if (!ok || !data.ok) throw new Error(data.error || "清单读取失败"); const result = data.result; const mapped = result.rows.filter(row => String(row["CAS号"] || "").trim()).length; const projects = new Set(result.rows.map(row => row["类别"]).filter(Boolean)).size; status.textContent = `共 ${result.count} 条物质明细 · ${projects} 个限制项目 · ${mapped} 条有效CAS映射 · 仅做物质级筛查`; sourceNote.innerHTML = `<div>${esc(result.meta["来源说明"] || "")}</div><div>${esc(result.meta["CAS规范化说明"] || "")}</div><div>${esc(result.meta["判定范围"] || "")}</div>`; const rows = result.rows.map(row => `<tr><td>${esc(row["序号"])}</td><td>${esc(row["源表行号"])}</td><td>${esc(row["类别"])}</td><td>${esc(row["中文名称"])}</td><td class="reach-id mono">${esc(row["CAS号"] || "（未建立有效CAS映射）")}</td><td class="reach-id mono">${esc(row["源表CAS号"] || "（空）")}</td><td>${esc(row["物质级判定"])}</td><td>${esc(row["CAS映射状态"])}</td><td>${esc(row["法规依据"])}</td><td>${esc(row["备注"])}</td></tr>`).join(""); box.innerHTML = `<table class="reach-register-table hsf-001-register-table"><thead><tr><th>序号</th><th>源表行号</th><th>类别</th><th>具体相关物质</th><th>CAS号</th><th>源表CAS号</th><th>物质级判定</th><th>CAS映射状态</th><th>法规依据</th><th>备注</th></tr></thead><tbody>${rows}</tbody></table>`; }).catch(error => { status.textContent = "清单读取失败"; box.innerHTML = `<div class="reach-check-error">${esc(error.message || "清单读取失败")}</div>`; });
    }

    function showAnnexXviiList() {
      setActive("annex-xvii-list");
      main.innerHTML = `<section class="regulatory-page reach-list-page"><header class="regulatory-page-head"><div><span class="regulatory-eyebrow">LEGAL MATERIALS REGISTER</span><h2>REACH Annex XVII</h2><p>REACH 限制条目及物质组的顶层限制记录；CAS/EC 组成员用于合规性检测命中。</p></div><span class="regulatory-page-chip">ANNEX XVII</span></header><div id="reach-list-status" class="reach-list-status">正在读取清单…</div><div class="reach-table-wrap" id="reach-table-wrap"><div class="wk-empty">加载中…</div></div></section>`;
      const status = main.querySelector("#reach-list-status"); const box = main.querySelector("#reach-table-wrap");
      fetch(API_OLD + "/api/annex-xvii/list").then(r => r.json().then(data => ({ ok: r.ok, data }))).then(({ ok, data }) => {
        if (!ok || !data.ok) throw new Error(data.error || "清单读取失败");
        const result = data.result;
        status.textContent = `共 ${result.count} 个现行顶层条目 · 组级条目可展开查看完整物质成员 · 只读查看`;
        const rows = result.rows.map((row, index) => {
          const entryKey = `annex-entry-${index}`;
          const members = (row["物质成员"] || []).filter(member => Number(member["是否物质组成员"]) === 1);
          const memberRows = members.map(member => `<tr class="annex-member-row" data-parent-entry="${entryKey}" hidden><td class="reach-id mono">↳ ${esc(row["条目号"])}</td><td>${esc(member["中文名称"] || "")}</td><td class="reach-en">${esc(member["英文名称"] || "")}</td><td class="reach-id mono">${esc(member["CAS号"] || "")}</td><td class="reach-id mono">${esc(member["EC号"] || "")}</td><td>${esc(member["物质描述"] || "")}</td><td>${esc(row["法规依据"] || "")}</td><td></td></tr>`).join("");
          const toggle = members.length ? `<button type="button" class="annex-member-toggle" data-annex-toggle="${entryKey}" aria-expanded="false">查看 ${members.length} 条成员</button>` : "无展开成员";
          const mainRow = `<tr><td class="reach-id mono">${esc(row["条目号"])}</td><td>${esc(row["中文名称"])}</td><td class="reach-en">${esc(row["英文名称"])}</td><td class="reach-id mono">${esc(row["CAS号"])}</td><td class="reach-id mono">${esc(row["EC号"])}</td><td>${esc(row["物质描述"])}</td><td>${esc(row["法规依据"])}<br><a class="regulation-source-link" href="${esc(row["限制条件链接"] || "#")}" target="_blank" rel="noopener">官方条件</a></td><td>${toggle}</td></tr>`;
          return mainRow + memberRows;
        }).join("");
        box.innerHTML = `<table class="reach-register-table annex-xvii-register-table"><thead><tr><th>条目号</th><th>中文名称</th><th>英文名称</th><th>CAS号</th><th>EC号</th><th>物质描述</th><th>法规依据</th><th>完整物质清单</th></tr></thead><tbody>${rows}</tbody></table>`;
        box.querySelectorAll("[data-annex-toggle]").forEach(button => button.addEventListener("click", () => {
          const key = button.dataset.annexToggle;
          const open = button.getAttribute("aria-expanded") !== "true";
          button.setAttribute("aria-expanded", String(open));
          button.textContent = open ? "收起成员" : `查看 ${box.querySelectorAll(`[data-parent-entry="${key}"]`).length} 条成员`;
          box.querySelectorAll(`[data-parent-entry="${key}"]`).forEach(row => { row.hidden = !open; });
        }));
      }).catch(error => { status.textContent = "清单读取失败"; box.innerHTML = `<div class="reach-check-error">${esc(error.message || "清单读取失败")}</div>`; });
    }

    nav.addEventListener("click", e => { const page = e.target.closest("[data-reg-page]"); if (page && !page.disabled) { if (page.dataset.regPage === "compliance") showCompliance(); else if (page.dataset.regPage === "reach-list") showReachList(); else if (page.dataset.regPage === "rohs-list") showRohsList(); else if (page.dataset.regPage === "annex-xvii-list") showAnnexXviiList(); else if (page.dataset.regPage === "hsf-001-list") showHsf001List(); return; } const toggle = e.target.closest("[data-reg-toggle]"); if (toggle) { const sub = win.body.querySelector(`[data-reg-subnav="${toggle.dataset.regToggle}"]`); const open = sub.classList.toggle("is-open"); toggle.classList.toggle("is-expanded", open); toggle.querySelector("i").textContent = open ? "▾" : "▸"; } });
    showCompliance();
  }

  // ===== 工作台逻辑 =====
  function initWorkbench(container) {
    const $mode = container.querySelector(".wk-mode");
    const $importBtn = container.querySelector("#wk-import-btn");
    const $importFile = container.querySelector("#wk-import-file");
    const $kbTools = container.querySelector("#wk-kb-tools");
    const $casTools = container.querySelector("#wk-cas-tools");
    const $left = container.querySelector("#wk-left");
    const $right = container.querySelector("#wk-right");
    const $split = container.querySelector("#wk-split");
    let currentMode = "import";
    let importData = null;    // 导入检索的解析结果
    let kbModels = [];       // 知识库的全部型号
    let kbFiltered = [];
    let kbCurrentModel = null; // 当前选中的型号数据 ({detail, wide, rows})
    let kbCurrentSec = null;

    // ---- 分割线拖动 ----
    let dragging = false, startX = 0, startW = 0;
    $split.onmousedown = e => { dragging = true; startX = e.clientX; startW = $left.offsetWidth; document.body.style.cursor = "col-resize"; };
    document.addEventListener("mousemove", e => {
      if (!dragging) return;
      const w = Math.max(200, Math.min(500, startW + (e.clientX - startX)));
      $left.style.width = w + "px";
    });
    document.addEventListener("mouseup", () => { dragging = false; document.body.style.cursor = ""; });

    // ---- 模式切换 ----
    $mode.onchange = () => {
      currentMode = $mode.value;
      $importBtn.style.display = currentMode === "import" ? "inline-flex" : "none";
      $kbTools.style.display = currentMode === "knowledge" ? "flex" : "none";
      $casTools.style.display = currentMode === "cas" ? "flex" : "none";
      renderLeft();
      renderRight(null);
    };

    // ---- 渲染左侧 ----
    function renderLeft() {
      if (currentMode === "import") renderImportLeft();
      else if (currentMode === "knowledge") renderKnowledgeLeft();
      else renderCasLeft();
    }

    // ==================== 导入检索 ====================
    function renderImportLeft() {
      $importBtn.style.display = "inline-flex";
      $importBtn.onclick = () => $importFile.click();
      $importFile.onchange = () => { if ($importFile.files[0]) uploadFile($importFile.files[0]); };
      renderImportSkeleton();

      function uploadFile(f) {
        if (!f.name.toLowerCase().endsWith(".docx")) { toast("仅支持 .docx 格式"); return; }
        $importBtn.textContent = "解析中..."; $importBtn.disabled = true;
        const fd = new FormData(); fd.append("file", f);
        api(API_NEW + "/api/import/upload", { method: "POST", body: fd }).then(data => {
          if (!data || !data.ok) { toast("解析失败"); $importBtn.textContent = "导入 MSDS"; $importBtn.disabled = false; return; }
          importData = data;
          renderImportResult();
          $importBtn.textContent = "导入 MSDS"; $importBtn.disabled = false;
          toast("解析完成：" + (data.model || data.filename));
        }).catch(e => { toast("上传失败：" + e.message); $importBtn.textContent = "导入 MSDS"; $importBtn.disabled = false; });
      }

      function renderImportSkeleton() {
        const treeHtml = Array.from({length: 17}, (_, n) =>
          `<div class="kb-tree-node" data-sec="${n}">
            <span class="kb-tree-num">S${n}</span>
            <span class="kb-tree-title">${esc(SECTION_NAMES[n] || "")}</span>
          </div>`).join("");
        $left.innerHTML = `<div class="gui-nav-title">17 节目录</div>
          <div class="gui-section-tree" id="import-tree-list">${treeHtml}</div>`;
        const tree = document.getElementById("import-tree-list");
        tree?.querySelectorAll(".kb-tree-node").forEach(el => {
          el.onclick = () => {
            tree.querySelectorAll(".kb-tree-node").forEach(x => x.classList.remove("active"));
            el.classList.add("active");
            if (importData) showImportSection(parseInt(el.dataset.sec));
          };
        });
      }

      function renderImportResult() {
        renderImportSkeleton();
        renderRight(`<div class="gui-content" id="import-sec-content"><div class="wk-empty">选择 Section 查看内容</div></div>`);
        const tree = document.getElementById("import-tree-list");
        tree?.querySelectorAll(".kb-tree-node").forEach(el => {
          el.onclick = () => {
            tree.querySelectorAll(".kb-tree-node").forEach(x => x.classList.remove("active"));
            el.classList.add("active"); showImportSection(parseInt(el.dataset.sec));
          };
        });
        const first = tree?.querySelector(".kb-tree-node");
        if (first) { first.classList.add("active"); showImportSection(parseInt(first.dataset.sec)); }
      }

      function showImportSection(num) {
        const sec = importData.sections.find(s => s.number === num);
        const target = document.getElementById("import-sec-content");
        if (!sec) { if (target) target.innerHTML = "<div class='wk-empty'>该节无数据</div>"; return; }
        let html = `<h4 class="sec-title">S${sec.number} ${esc(SECTION_NAMES[sec.number] || sec.title)}</h4>`;
        // 直接渲染 GUI 标准骨架行；不再从原始 big_titles/direct_fields 自行 flatten。
        const allFields = sec.rows || [];

        if (!allFields.length) { if (target) target.innerHTML = "<div class='wk-empty'>该节无数据</div>"; return; }
        html += renderGuiTable(allFields);
        if (target) target.innerHTML = `<div class="il-sec-view">${html}</div>`;
      }
    }

    // ==================== CAS 库检索 ====================
    function renderCasLeft() {
      $left.innerHTML = `
        <div class="db-model-title">CAS 物质列表</div>
        <div class="cas-list" id="cas-list"></div>
        <div class="kb-status" id="cas-status"></div>`;
      const $query = container.querySelector("#cas-query");
      const $statusFilter = container.querySelector("#cas-status-filter");
      const $search = container.querySelector("#cas-search-btn");
      const $clear = container.querySelector("#cas-clear-btn");
      const $list = $left.querySelector("#cas-list");
      const $status = $left.querySelector("#cas-status");

      function loadCas(q = "", status = $statusFilter.value) {
        $status.textContent = "检索中…";
        const params = [];
        if (q) params.push("q=" + encodeURIComponent(q));
        if (status) params.push("status=" + encodeURIComponent(status));
        api(API_OLD + "/api/cas" + (params.length ? "?" + params.join("&") : ""))
          .then(data => {
            const items = data.items || [];
            $status.textContent = `共 ${items.length} 个物质`;
            $list.innerHTML = items.map(item => `
              <div class="cas-item" data-id="${item.cas_id}">
                <div class="cas-item-name">${esc(item.standard_name || "未命名物质")}</div>
                <div class="cas-item-meta">${esc(item.cas_no || "无 CAS 号")} · ${esc(item.registry_status || "待确认")}</div>
              </div>`).join("") || "<div class='wk-empty' style='padding:20px'>无匹配 CAS 物质</div>";
            $list.querySelectorAll(".cas-item").forEach(el => {
              el.onclick = () => {
                $list.querySelectorAll(".cas-item").forEach(x => x.classList.remove("active"));
                el.classList.add("active");
                const item = items.find(x => String(x.cas_id) === el.dataset.id);
                if (item) renderCasDetail(item);
              };
            });
            const first = $list.querySelector(".cas-item");
            if (first) { first.classList.add("active"); renderCasDetail(items[0]); }
          })
          .catch(() => { $list.innerHTML = "<div class='wk-empty'>CAS 库加载失败</div>"; $status.textContent = ""; });
      }

      $statusFilter.onchange = () => loadCas($query.value.trim(), $statusFilter.value);
      $search.onclick = () => loadCas($query.value.trim(), $statusFilter.value);
      $query.onkeydown = e => { if (e.key === "Enter") loadCas($query.value.trim(), $statusFilter.value); };
      $clear.onclick = () => { $query.value = ""; $statusFilter.value = ""; loadCas(); };
      loadCas();
    }

    function renderCasDetail(item) {
      const usage = item.usage || [];
      const aliases = item.aliases || [];
      const s2 = item.s2_result;
      const usageRows = usage.map(row => `<tr>
        <td>${esc(row.model || "")}</td><td>${esc(row.raw_name || "")}</td>
        <td>${esc(row.raw_cas || "")}</td><td>${esc(row.concentration || "")}</td>
      </tr>`).join("") || `<tr><td colspan="4">暂无型号使用关系</td></tr>`;
      let resultHtml = `<section class="cas-result-section"><h4 class="cas-usage-title">Section 2 标准结果</h4><div class="cas-result-empty">该 CAS 暂无标准骨架结果</div></section>`;
      if (s2) {
        const standardRows = (s2.standard_fields || []).slice().sort((a, b) => Number(a.display_order || 0) - Number(b.display_order || 0));
        const rowsHtml = standardRows.map(row => { const label = row.field_label_zh || ""; const value = row.value_text || "无数据"; const isPictogram = label === "象形图" || label === "GHS象形图"; const content = isPictogram ? `<div class="cas-pictogram-value">${esc(value)}</div>${renderGhsValue(value, "无 GHS 象形图")}` : esc(value); return `<tr><th>${esc(label)}</th><td>${content}</td></tr>`; }).join("");
        resultHtml = `<section class="cas-result-section"><h4 class="cas-usage-title">Section 2 标准结果</h4>
          <table class="cas-result-table cas-standard-table"><tbody>${rowsHtml || `<tr><td colspan="2">暂无 Section 2 标准结果</td></tr>`}</tbody></table></section>`;
      }
      renderRight(`<div class="gui-content cas-content">
        <h4 class="sec-title">CAS 物质详情</h4>
        <table class="cas-detail-table"><tbody>
          <tr><th>CAS 号</th><td>${esc(item.cas_no || "无 CAS 号")}</td><th>标准名称</th><td>${esc(item.standard_name || "")}</td></tr>
          <tr><th>类别</th><td>${esc(item.category || "")}</td><th>登记状态</th><td>${esc(item.registry_status || "")}</td></tr>
          <tr><th>别名</th><td colspan="3">${esc(aliases.join("；") || "无")}</td></tr>
          <tr><th>备注</th><td colspan="3">${esc(item.remarks || "")}</td></tr>
        </tbody></table>
        <h4 class="cas-usage-title">型号使用关系（${usage.length}）</h4>
        <table class="cas-usage-table"><thead><tr><th>型号</th><th>原始名称</th><th>原始 CAS</th><th>浓度</th></tr></thead>
          <tbody>${usageRows}</tbody></table>
        ${resultHtml}
      </div>`);
    }

    // ==================== 知识库检索 ====================
    function renderKnowledgeLeft() {
      $left.innerHTML = `
        <div class="db-model-title">型号列表</div>
        <div class="kb-list" id="kb-list"></div>
        <div class="kb-status" id="kb-status"></div>`;
      const $inp = container.querySelector("#kb-inp");
      const $category = container.querySelector("#kb-category");
      const $kw = container.querySelector("#kb-kw");
      const $search = container.querySelector("#kb-search-btn");
      const $clear = container.querySelector("#kb-clear");
      const $list = $left.querySelector("#kb-list");
      const $status = $left.querySelector("#kb-status");

      let currentCategory = "";
      $category.onchange = () => { currentCategory = $category.value; loadModels(); };
      $inp.oninput = () => filterModels($inp.value.trim());
      $search.onclick = () => {
        const kw = $kw.value.trim();
        if (kw) keywordSearch(kw);
        else filterModels($inp.value.trim());
      };
      $clear.onclick = () => { $inp.value = ""; $kw.value = ""; $category.value = ""; currentCategory = ""; kwHits = {}; loadModels(); };
      $inp.onkeydown = e => { if (e.key === "Enter") filterModels($inp.value.trim()); };
      $kw.onkeydown = e => { if (e.key === "Enter") keywordSearch($kw.value.trim()); };

      let kwHits = {};

      function filterModels(q) {
        q = q.toLowerCase();
        if (!q) { kbFiltered = kbModels; }
        else { kbFiltered = kbModels.filter(m => m.model.toLowerCase().includes(q)); }
        renderModelList();
      }

      function keywordSearch(kw) {
        kw = kw.trim();
        if (!kw) { filterModels($inp.value.trim()); return; }
        $status.textContent = "检索中…";
        const qs = "?q=" + encodeURIComponent(kw) + (currentCategory ? "&category=" + encodeURIComponent(currentCategory) : "");
        api(API_OLD + "/api/models" + qs).then(list => {
          const items = list.items || [];
          kwHits = {};
          const ids = new Set(items.map(m => m.id));
          kbFiltered = kbModels.filter(m => ids.has(m.id) || ids.has(m.model_id));
          items.forEach(m => { kwHits[m.id] = m.hits || ""; });
          renderModelList();
        }).catch(() => { toast("检索失败"); $status.textContent = ""; });
      }

      function renderModelList() {
        $list.innerHTML = kbFiltered.map(m => {
          const hit = kwHits[m.id || m.model_id] || "";
          const hitHtml = hit ? `<span class="import-hit">${esc(hit)}</span>` : "";
          return `<div class="kb-item" data-id="${m.id || m.model_id}">
            <div class="kb-item-name">${esc(m.model)}</div>
            <div class="kb-item-meta">${esc(m.source || "")} · ${m.fields || m.fields_count || 0} 字段 ${hitHtml}</div>
          </div>`;
        }).join("") || "<div class='wk-empty' style='padding:20px'>无匹配型号</div>";
        $status.textContent = `共 ${kbFiltered.length} 个型号`;
        $list.querySelectorAll(".kb-item").forEach(el => {
          el.onclick = () => {
            $list.querySelectorAll(".kb-item").forEach(x => x.classList.remove("active"));
            el.classList.add("active");
            loadModel(parseInt(el.dataset.id));
          };
        });
      }

      function loadModel(id) {
        kbCurrentSec = null;
        api(API_OLD + "/api/models/" + id).then(data => {
          if (!data || !data.rows) { toast("加载型号详情失败"); return; }
          kbCurrentModel = data;
          // 渲染右侧 17 节目录 + 内容
          // Keep the GUI's fixed S0~S16 navigation tree, including sections
          // whose current record has no populated value.
          const secs = Array.from({length: 17}, (_, n) => n);
          let treeHtml = secs.map(n =>
            `<div class="kb-tree-node" data-sec="${n}">
              <span class="kb-tree-num">S${n}</span>
              <span class="kb-tree-title">${esc(SECTION_NAMES[n] || "")}</span>
            </div>`
          ).join("");
          let rightHtml = `<div class="kb-right-split">
            <div class="kb-tree-list" id="kb-tree-list">${treeHtml}</div>
            <div class="kb-sec-content" id="kb-sec-content"><div class="wk-empty">选择节查看内容</div></div>
          </div>`;
          renderRight(rightHtml);
          // 绑定节点击
          const $tree = document.getElementById("kb-tree-list");
          if ($tree) {
            $tree.querySelectorAll(".kb-tree-node").forEach(el => {
              el.onclick = () => {
                $tree.querySelectorAll(".kb-tree-node").forEach(x => x.classList.remove("active"));
                el.classList.add("active");
                showSection(parseInt(el.dataset.sec));
              };
            });
          }
          // 自动选中第一节
          if (secs.length > 0) {
            const first = $tree?.querySelector(".kb-tree-node");
            if (first) { first.classList.add("active"); showSection(secs[0]); }
          }
        }).catch(() => toast("加载型号详情失败"));
      }

      function showSection(n) {
        kbCurrentSec = n;
        const rows = kbCurrentModel?.rows?.[n];
        if (!rows || !rows.length) {
          const $sc = document.getElementById("kb-sec-content");
          if ($sc) $sc.innerHTML = "<div class='wk-empty'>该节无数据</div>";
          return;
        }
        let html = `<h4 class="sec-title">S${n} ${esc(SECTION_NAMES[n] || "")}</h4>`;
        html += renderGuiTable(rows);
        const $sc = document.getElementById("kb-sec-content");
        if ($sc) $sc.innerHTML = html;
      }

      function loadModels() {
        const suffix = currentCategory ? "?category=" + encodeURIComponent(currentCategory) : "";
        api(API_OLD + "/api/models" + suffix).then(list => {
          kbModels = list.items || [];
          kbFiltered = kbModels;
          renderModelList();
        }).catch(() => { $list.innerHTML = "<div class='wk-empty'>加载型号列表失败</div>"; });
      }

      // 大类是独立筛选条件，型号列表只显示该大类下的型号。
      $category.innerHTML = '<option value="">全部大类</option>';
      api(API_OLD + "/api/categories").then(data => {
        (data.items || []).forEach(group => {
          const option = document.createElement("option");
          option.value = group.name; option.textContent = group.name;
          $category.appendChild(option);
        });
      }).catch(() => {});
      loadModels();
    }

    // ===== 渲染右侧 =====
    function renderRight(html) {
      if (typeof html === "string") {
        $right.innerHTML = html;
      } else {
        $right.innerHTML = `<div class="wk-empty">选择左侧节点查看内容</div>`;
      }
    }

    // 与 gui/section_tree.py 的 SectionView 对齐：权限徽章 | 序号 | 标签 | 字段。
    // Web 只负责渲染，editable 仍来自既有解析器/数据库字段权限结果。
    function renderGuiTable(rows) {
      let html = `<div class="gui-table-wrap"><table class="gui-table"><thead><tr>`;
      html += `<th class="gui-badge-col">状态</th><th class="gui-seq-col">序号</th><th class="gui-label-col">标签</th><th>字段</th></tr></thead><tbody>`;
      (rows || []).forEach((r, i) => {
        const editable = !!r.editable;
        if (r.kind === "subtable") {
          const headers = Array.isArray(r.sub_header) ? r.sub_header : [];
          const subRows = Array.isArray(r.sub_rows) ? r.sub_rows : [];
          // GUI already renders the preceding `sub` row as the section title.
          // The nested table is the following content block; rendering its label
          // again here creates duplicate skeleton rows such as “成分/生物限值”.
          html += `<tr class="gui-nested-row"><td colspan="4"><table class="gui-nested-table"><thead><tr>`;
          headers.forEach(h => { html += `<th>${esc(h)}</th>`; });
          html += `</tr></thead><tbody>`;
          (subRows.length ? subRows : [headers.map(() => "")]).forEach(row => {
            html += `<tr>`;
            headers.forEach((_, index) => { html += `<td>${esc(row?.[index] ?? "")}</td>`; });
            html += `</tr>`;
          });
          html += `</tbody></table></td></tr>`;
          return;
        }
        const status = editable ? "可编辑" : "不可编辑";
        const badgeClass = editable ? "gui-badge-editable" : "gui-badge-fixed";
        if (r.span) {
          html += `<tr><td><span class="gui-badge ${badgeClass}">${status}</span></td>`;
          html += `<td colspan="3" class="gui-span-value">${esc(r.value || r.label || "")}</td></tr>`;
          return;
        }
        html += `<tr><td><span class="gui-badge ${badgeClass}">${status}</span></td>`;
        html += `<td class="gui-seq">${esc(r.seq || "")}</td>`;
        html += `<td class="gui-label">${r.label ? "🔒 " : ""}${esc(r.label || "")}</td>`;
        html += `<td class="gui-value ${editable ? "" : "gui-value-muted"}">${esc(r.value || "")}</td></tr>`;
      });
      html += `</tbody></table></div>`;
      return html;
    }

    // 初始渲染
    renderLeft();
    renderRight(null);
  }

  // ===== 初始化 =====
  function init() {
    loadGhsCatalog().catch(() => {});
    addDesktopIcon("检索工作台", "📚", openDbSearchApp);
    addDesktopIcon("CAS 联网搜索", "🧪", openCasOnlineApp);
    addDesktopIcon("合规性判断", "⚖️", openComplianceApp);
    addDesktopIcon("MSDS 覆写工具", "📝", openOverwriteApp);
    addDesktopIcon("TDS 覆写工具", "📄", openTdsOverwriteApp);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
