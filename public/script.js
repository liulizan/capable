/**
 * Capable（可培）— 前端脚本 v2
 * 新增：marked渲染、Toast、localStorage草稿、Tab切换、导出下拉
 */

// ==================== marked 配置 ====================
const useMarked = typeof marked !== 'undefined';
if (useMarked && marked.setOptions) {
  marked.setOptions({ breaks: true, gfm: true });
}

// ==================== 主题 ====================
const THEME_KEY = 'capable_theme';
const $html = document.documentElement;
const $btnTheme = document.getElementById('btnTheme');
const $themeIcon = $btnTheme.querySelector('.theme-icon');
const $themeLabel = $btnTheme.querySelector('.theme-label');

function applyTheme(theme) {
  $html.setAttribute('data-theme', theme);
  $themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
  $themeLabel.textContent = theme === 'dark' ? '亮色模式' : '深色模式';
}
function toggleTheme() {
  const next = $html.getAttribute('data-theme') === 'light' ? 'dark' : 'light';
  applyTheme(next);
  localStorage.setItem(THEME_KEY, next);
}
$btnTheme.addEventListener('click', toggleTheme);
applyTheme(localStorage.getItem(THEME_KEY) === 'dark' ? 'dark' : 'light');

// ==================== Toast ====================
const $toastContainer = document.getElementById('toastContainer');
function showToast(msg, type = 'info', duration = 3000) {
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  $toastContainer.appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => { el.classList.remove('show'); setTimeout(() => el.remove(), 300); }, duration);
}

// ==================== Markdown 渲染 ====================
function renderMarkdown(text) {
  if (useMarked) {
    return marked.parse(text);
  }
  // 降级：简版渲染
  let html = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/^---$/gm, '<hr>');
  html = html.replace(/\n\n/g, '</p><p>');
  html = '<p>' + html + '</p>';
  return html;
}

// ==================== DOM ====================
const $identityPreset = document.getElementById('identityPreset');
const $userProfile = document.getElementById('userProfile');
const $content = document.getElementById('content');
const $btnTransform = document.getElementById('btnTransform');
const $transformSpinner = $btnTransform.querySelector('.spinner');
const $loadingSkeleton = document.getElementById('loadingSkeleton');
const $outputArea = document.getElementById('outputArea');
const $outputContent = document.getElementById('outputContent');
const $tabBar = document.getElementById('tabBar');
const $outputToolbar = document.getElementById('outputToolbar');
const $actionArea = document.getElementById('actionArea');
const $roleplaySection = document.getElementById('roleplaySection');
const $rpScenario = document.getElementById('rpScenario');
const $userReply = document.getElementById('userReply');
const $btnRoleplay = document.getElementById('btnRoleplay');
const $rpSpinner = $btnRoleplay.querySelector('.spinner');
const $rpResult = document.getElementById('rpResult');
const $emptyState = document.getElementById('emptyState');
const $modeSelector = document.getElementById('modeSelector');
const $feedbackArea = document.getElementById('feedbackArea');
const $feedbackBtns = $feedbackArea.querySelectorAll('.btn-feedback');
const $feedbackThanks = document.getElementById('feedbackThanks');
const $dropZone = document.getElementById('dropZone');
const $fileInput = document.getElementById('fileInput');
const $fileStatus = document.getElementById('fileStatus');
const $urlInput = document.getElementById('urlInput');
const $btnFetchUrl = document.getElementById('btnFetchUrl');
const $fetchSpinner = $btnFetchUrl.querySelector('.spinner');
const $urlStatus = document.getElementById('urlStatus');
const $btnCopyAll = document.getElementById('btnCopyAll');
const $btnExport = document.getElementById('btnExport');
const $exportMenu = document.getElementById('exportMenu');
const $btnResetAll = document.getElementById('btnResetAll');
const $btnSaveTemplate = document.getElementById('btnSaveTemplate');
const $fileList = document.getElementById('fileList');
const $charCount = document.getElementById('charCount');
const $btnPasteClip = document.getElementById('btnPasteClip');
const $btnClearTextFloat = document.getElementById('btnClearTextFloat');
const $btnSaveDraft = document.getElementById('btnSaveDraft');
const $searchInput = document.getElementById('searchInput');
const $outputSearch = document.getElementById('outputSearch');
const $modeCount = document.getElementById('modeCount');
const $toggleEvidence = document.getElementById('toggleEvidence');
const $toggleFallacy = document.getElementById('toggleFallacy');
const $toggleCounter = document.getElementById('toggleCounter');
const $toggleEvidenceCb = $toggleEvidence?.querySelector('input');
const $toggleFallacyCb = $toggleFallacy?.querySelector('input');
const $toggleCounterCb = $toggleCounter?.querySelector('input');

// ==================== 状态 ====================
const DRAFT_KEY = 'capable_draft';
let conversationHistory = [];
let currentSessionId = null;

// ==================== localStorage 草稿 ====================
function saveDraft() {
  const draft = {
    identity: $identityPreset.value,
    customIdentity: $userProfile.value,
    content: $content.value,
    modes: [...$modeSelector.querySelectorAll('input:checked')].map(c => c.value),
    time: Date.now(),
  };
  localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
}
function loadDraft() {
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    if (!raw) return;
    const d = JSON.parse(raw);
    if (d.identity) { $identityPreset.value = d.identity; $identityPreset.dispatchEvent(new Event('change')); }
    if (d.customIdentity) $userProfile.value = d.customIdentity;
    if (d.content) $content.value = d.content;
    if (d.modes) {
      const cbs = $modeSelector.querySelectorAll('input[name="mode"]');
      cbs.forEach(cb => cb.checked = d.modes.includes(cb.value));
    }
  } catch {}
}
// 自动保存（每5秒 + 输入时）
setInterval(saveDraft, 5000);
[$content, $userProfile, $identityPreset].forEach(el => el.addEventListener('input', saveDraft));
$modeSelector.addEventListener('change', saveDraft);

// ==================== 身份下拉 ====================
$identityPreset.addEventListener('change', () => {
  const val = $identityPreset.value;
  if (val === '__custom__') { $userProfile.classList.remove('hidden'); $userProfile.focus(); }
  else { $userProfile.classList.add('hidden'); if (val) $userProfile.value = val; else $userProfile.value = ''; }
  saveDraft();
});
$userProfile.addEventListener('input', () => {
  if (!$userProfile.classList.contains('hidden')) $identityPreset.value = '__custom__';
});

// ==================== 模式选择 ====================
function getSelectedModes() {
  const checked = $modeSelector.querySelectorAll('input[name="mode"]:checked');
  return checked.length > 0 ? [...checked].map(cb => cb.value) : ['capability'];
}

// ==================== 模式卡片排序 + 计数 ====================
function sortModeCards() {
  // 按分组分别排序：每组内勾选的排前面
  ['review', 'deep'].forEach(group => {
    const cards = [...$modeSelector.querySelectorAll(`.mode-card[data-group="${group}"]`)];
    const checked = cards.filter(c => c.querySelector('input').checked);
    const unchecked = cards.filter(c => !c.querySelector('input').checked);
    const label = $modeSelector.querySelector(`.mode-group-label:has(+ [data-group="${group}"])`) ||
                  [...$modeSelector.querySelectorAll('.mode-group-label')].find(l => {
                    let el = l.nextElementSibling;
                    while (el && el.classList.contains('mode-card')) {
                      if (el.dataset.group === group) return true;
                      el = el.nextElementSibling;
                    }
                    return false;
                  });
    // 在组内重新排列
    const anchor = [...$modeSelector.querySelectorAll(`.mode-card[data-group="${group}"]`)][0];
    if (anchor) {
      [...checked, ...unchecked].forEach(c => $modeSelector.insertBefore(c, anchor.nextElementSibling || null));
    }
    // 更新分组计数
    const cnt = $modeSelector.querySelector(`.mode-group-count[data-group="${group}"]`);
    if (cnt) cnt.textContent = `已选 ${checked.length}/2`;
  });
  // 更新总计数
  const total = $modeSelector.querySelectorAll('input[name="mode"]:checked').length;
  $modeCount.textContent = `已选 ${total}/4`;
}

$modeSelector.addEventListener('change', (e) => {
  if (e.target.name === 'mode') {
    const card = e.target.closest('.mode-card');
    if (card) card.classList.toggle('selected', e.target.checked);
    sortModeCards();
    saveDraft();
  }
});
// 初始化
sortModeCards();
$modeSelector.querySelectorAll('input:checked').forEach(cb => cb.closest('.mode-card')?.classList.add('selected'));

// ==================== 转化按钮状态 + 字数统计 ====================
function updateTransformBtn() {
  $btnTransform.disabled = !$content.value.trim();
}
function updateCharCount() {
  $charCount.textContent = `${$content.value.length.toLocaleString()} / 50000 字`;
}
$content.addEventListener('input', () => { updateTransformBtn(); updateCharCount(); saveDraft(); });
updateTransformBtn(); updateCharCount();

// ==================== 抓取按钮状态 ====================
function updateFetchBtn() {
  $btnFetchUrl.disabled = !$urlInput.value.trim();
}
$urlInput.addEventListener('input', updateFetchBtn);
updateFetchBtn();

// ==================== 清空 / 重置 / 保存模板 ====================
$btnResetAll.addEventListener('click', () => {
  $content.value = ''; $urlInput.value = ''; $userProfile.value = ''; $identityPreset.value = '';
  $userProfile.classList.add('hidden'); $urlStatus.classList.add('hidden'); $fileStatus.classList.add('hidden');
  $fileList.innerHTML = ''; $fileList.classList.add('hidden');
  $modeSelector.querySelectorAll('input[name="mode"]').forEach(cb => { cb.checked = cb.value === 'capability'; cb.closest('.mode-card')?.classList.toggle('selected', cb.value === 'capability'); });
  sortModeCards();
  hideResults(); updateTransformBtn(); updateCharCount(); updateFetchBtn(); saveDraft();
  showToast('🔄 已重置全部配置', 'info');
});

$btnSaveTemplate.addEventListener('click', () => {
  const identity = $identityPreset.value === '__custom__' ? $userProfile.value : $identityPreset.value;
  const modes = getSelectedModes();
  const favs = JSON.parse(localStorage.getItem('capable_favorites') || '[]');
  favs.unshift({ identity, modes, time: Date.now() });
  if (favs.length > 10) favs.pop();
  localStorage.setItem('capable_favorites', JSON.stringify(favs));
  showToast('💾 配置模板已保存到收藏', 'success');
});

// ==================== 剪贴板粘贴 ====================
$btnPasteClip.addEventListener('click', async () => {
  try { const t = await navigator.clipboard.readText(); $content.value = t; updateTransformBtn(); updateCharCount(); showToast('✅ 已粘贴', 'success'); }
  catch { showToast('❌ 无法读取剪贴板', 'error'); }
});
$btnClearTextFloat.addEventListener('click', () => { $content.value = ''; updateTransformBtn(); updateCharCount(); });
$btnSaveDraft.addEventListener('click', () => { saveDraft(); showToast('💾 草稿已保存', 'success'); });

// ==================== 文件列表管理 ====================
function addFileToList(name, size) {
  $fileList.classList.remove('hidden');
  const li = document.createElement('li');
  const sizeStr = size > 1024*1024 ? (size/1024/1024).toFixed(1)+'MB' : (size/1024).toFixed(0)+'KB';
  li.innerHTML = `<span class="file-name">📎 ${name} (${sizeStr})</span><button class="file-remove-btn" title="移除">✕</button>`;
  li.querySelector('.file-remove-btn').addEventListener('click', () => { li.remove(); if (!$fileList.querySelector('li')) $fileList.classList.add('hidden'); });
  $fileList.appendChild(li);
}
// 文件列表在 uploadFile 成功的分支中更新（见下方 uploadFile 函数内的 addFileToList 调用）

// ==================== 批判分析折叠 + 搜索 ====================
function applyCritiqueCards() {
  // 将论点标题转为折叠卡片
  const hs = $outputContent.querySelectorAll('h3');
  hs.forEach(h3 => {
    if (h3.closest('.claim-card')) return; // 已转换
    const card = document.createElement('div'); card.className = 'claim-card open';
    const header = document.createElement('div'); header.className = 'claim-card-header';
    header.innerHTML = `<span>${h3.textContent}</span><span class="claim-actions"><button class="btn-copy-claim">📋</button><button class="btn-mark-claim">⭐</button></span>`;
    const body = document.createElement('div'); body.className = 'claim-card-body';
    // 收集 h3 之后到下一个 h3 之间的内容
    let el = h3.nextElementSibling;
    while (el && el.tagName !== 'H3') { body.appendChild(el.cloneNode(true)); const next = el.nextElementSibling; el.remove(); el = next; }
    // 替换
    header.addEventListener('click', (e) => { if (!e.target.closest('button')) card.classList.toggle('open'); });
    header.querySelector('.btn-copy-claim').addEventListener('click', (e) => { e.stopPropagation(); navigator.clipboard.writeText(body.textContent.trim()); showToast('✅ 已复制', 'success'); });
    header.querySelector('.btn-mark-claim').addEventListener('click', (e) => { e.stopPropagation(); card.classList.toggle('marked'); });
    card.appendChild(header); card.appendChild(body);
    h3.replaceWith(card);
  });
}

// 搜索筛选
$searchInput.addEventListener('input', () => {
  const q = $searchInput.value.trim().toLowerCase();
  const cards = $outputContent.querySelectorAll('.claim-card');
  cards.forEach(c => c.classList.toggle('hidden-by-search', q && !c.textContent.toLowerCase().includes(q)));
});

// 为每个 h2/h3 标题添加"复制本段"按钮
function injectSectionCopyBtns() {
  $outputContent.querySelectorAll('h2, h3').forEach(heading => {
    if (heading.querySelector('.section-copy-btn')) return;
    const btn = document.createElement('button');
    btn.className = 'section-copy-btn';
    btn.textContent = '📋';
    btn.title = '复制本段内容';
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      // 收集该标题到下一个同级标题之间的内容
      let text = heading.textContent.replace('📋', '').trim() + '\n';
      let el = heading.nextElementSibling;
      while (el && !['H2','H3'].includes(el.tagName)) {
        if (!el.classList.contains('section-copy-btn')) text += el.textContent + '\n';
        el = el.nextElementSibling;
      }
      navigator.clipboard.writeText(text.trim()).then(() => showToast('✅ 已复制本段', 'success'));
    });
    heading.appendChild(btn);
  });
}

// ==================== 模式卡片选中增强 ====================
$modeSelector.addEventListener('change', (e) => {
  if (e.target.name === 'mode') {
    const card = e.target.closest('.mode-card');
    if (card) card.classList.toggle('selected', e.target.checked);
  }
});

// ==================== 校验 ====================
function validate() {
  if (!$content.value.trim()) { showToast('⚠️ 请先输入要转化的内容', 'warning'); return false; }
  return true;
}

// ==================== 核心转化 ====================
async function transformContent() {
  if ($btnTransform.disabled) return;
  if (!validate()) return;

  const content = $content.value.trim();
  setLoading(true);
  hideResults();
  $loadingSkeleton.classList.remove('hidden');
  $emptyState.classList.add('hidden');
  showToast('⏳ 正在转化中…', 'info', 0);

  try {
    const res = await fetch('/api/transform', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ content, userProfile: $userProfile.value.trim(), identity: $userProfile.value.trim(), modes: getSelectedModes() }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `请求失败 (${res.status})`);

    conversationHistory = data.conversationHistory;
    currentSessionId = data.sessionId || null;

    if (data.results) {
      showMultiResults(data.results);
    } else {
      showResults(data.reply, data.mode || 'capability');
    }
    $toastContainer.querySelectorAll('.toast').forEach(t => t.remove());
    showToast('✅ 转化完成', 'success');

    // 保存历史
    const identity = $identityPreset.value === '__custom__' ? $userProfile.value : $identityPreset.value;
    addToHistory(content, identity, getSelectedModes(), $outputContent.textContent.trim());
  } catch (err) {
    showToast(`❌ ${err.message}`, 'error');
  } finally {
    setLoading(false);
    $loadingSkeleton.classList.add('hidden');
  }
}

function setLoading(loading) {
  $btnTransform.disabled = loading;
  $btnTransform.classList.toggle('loading', loading);
  $btnTransform.querySelector('.btn-text').textContent = loading ? '转化中…' : '开始能力转化';
  $transformSpinner.classList.toggle('hidden', !loading);
}

// ==================== 结果渲染 ====================
function showResults(text, mode) {
  let mainText = text, rpText = '';
  if (mode === 'capability') {
    const idx = text.indexOf('现在开始角色扮演');
    if (idx !== -1) { mainText = text.slice(0, idx).trim(); rpText = text.slice(idx).trim(); }
  }
  $outputContent.innerHTML = renderMarkdown(mainText);
  $outputArea.classList.remove('hidden');
  $tabBar.classList.add('hidden');
  $outputTabs.classList.add('hidden');
  $emptyState.classList.add('hidden');
  $actionArea.classList.remove('hidden');
  $roleplaySection.classList.toggle('hidden', !(mode === 'capability' && rpText));
  if (rpText) $rpScenario.textContent = rpText;
  $rpResult.classList.add('hidden');
  toggleCritiqueToggles(mode === 'critique');
  $outputSearch.classList.toggle('hidden', mode !== 'critique');
  if (mode === 'critique') { setTimeout(applyCritiqueCards, 200); $searchInput.value = ''; }
  setTimeout(injectSectionCopyBtns, 100);
}

/** 多模式结果 Tab 渲染 */
function showMultiResults(results) {
  const modes = Object.keys(results).filter(k => results[k]);
  if (modes.length <= 1) {
    const m = modes[0] || 'capability';
    return showResults(results[m], m);
  }

  const labels = { summarize: '📋 总结', extract: '🔍 提取', capability: '🧠 练习卡片', critique: '💬 批判' };
  // 渲染 Tab 栏
  $tabBar.innerHTML = modes.map((m, i) =>
    `<button class="tab-btn${i === 0 ? ' active' : ''}" data-tab="${m}">${labels[m] || m}</button>`
  ).join('');
  $tabBar.classList.remove('hidden');

  // 渲染全部内容（按 mode 分组）
  let html = '';
  modes.forEach((m, i) => {
    html += `<div class="tab-panel${i === 0 ? ' active' : ''}" data-tab="${m}">${renderMarkdown(results[m])}</div>`;
  });
  $outputContent.innerHTML = html;
  $outputArea.classList.remove('hidden');
  $emptyState.classList.add('hidden');
  $actionArea.classList.remove('hidden');
  $roleplaySection.classList.add('hidden');

  // Tab 切换事件
  $tabBar.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      $tabBar.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const tab = btn.dataset.tab;
      $outputContent.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.dataset.tab === tab));
    });
  });

  toggleCritiqueToggles(modes.includes('critique'));
  const hasCritique = modes.includes('critique');
  $outputSearch.classList.toggle('hidden', !hasCritique);
  if (hasCritique) { setTimeout(applyCritiqueCards, 200); $searchInput.value = ''; }
  $outputTabs.classList.toggle('hidden', modes.length <= 1);
  setTimeout(injectSectionCopyBtns, 100);
}

function toggleCritiqueToggles(visible) {
  [$toggleEvidence, $toggleFallacy, $toggleCounter].forEach(t => t?.classList.toggle('hidden', !visible));
}

function hideResults() {
  $outputArea.classList.add('hidden'); $actionArea.classList.add('hidden'); $tabBar.classList.add('hidden');
  $roleplaySection.classList.add('hidden'); $rpResult.classList.add('hidden'); $loadingSkeleton.classList.add('hidden');
  $emptyState.classList.remove('hidden');
  $outputContent.innerHTML = ''; $rpScenario.textContent = ''; $userReply.value = '';
  $feedbackArea.classList.add('hidden'); $feedbackThanks.classList.add('hidden');
  conversationHistory = []; currentSessionId = null;
}

// ==================== 导出 ====================
function getOutputText() { return $outputContent.textContent.trim(); }

$btnCopyAll.addEventListener('click', async () => {
  await navigator.clipboard.writeText(getOutputText());
  showToast('✅ 已复制到剪贴板', 'success');
});

$btnExport.addEventListener('click', () => $exportMenu.classList.toggle('hidden'));
document.addEventListener('click', (e) => { if (!$btnExport.contains(e.target) && !$exportMenu.contains(e.target)) $exportMenu.classList.add('hidden'); });

$exportMenu.addEventListener('click', (e) => {
  const btn = e.target.closest('button');
  if (!btn) return;
  $exportMenu.classList.add('hidden');
  const fmt = btn.dataset.format;
  if (fmt === 'md') exportFile(getOutputText(), 'text/markdown', 'md');
  else if (fmt === 'word') {
    const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:'Microsoft YaHei',sans-serif;line-height:1.8;max-width:800px;margin:0 auto;padding:2rem}h2{color:#6366f1}h3{color:#4f46e5}table{border-collapse:collapse;width:100%}td,th{padding:.4rem .6rem;border:1px solid #ddd}</style></head><body>${$outputContent.innerHTML}</body></html>`;
    exportFile(html, 'application/msword', 'doc');
  } else if (fmt === 'csv') {
    // 提取表格数据导出 CSV
    const tables = $outputContent.querySelectorAll('table');
    let csv = '';
    tables.forEach(table => {
      table.querySelectorAll('tr').forEach(row => {
        const cells = [...row.querySelectorAll('td,th')].map(c => `"${c.textContent.replace(/"/g,'""')}"`).join(',');
        csv += cells + '\n';
      });
      csv += '\n';
    });
    if (!csv) csv = getOutputText().split('\n').map(l => `"${l.replace(/"/g,'""')}"`).join('\n');
    exportFile(csv || '无数据', 'text/csv', 'csv');
  }
});

function exportFile(content, mime, ext) {
  const blob = new Blob(['﻿' + content], { type: `${mime};charset=utf-8` });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = `capable-${new Date().toISOString().slice(0,10)}.${ext}`;
  a.click(); URL.revokeObjectURL(a.href);
  showToast(`✅ 已导出 .${ext}`, 'success');
}

// ==================== 批判分析开关 ====================
[$toggleEvidenceCb, $toggleFallacyCb, $toggleCounterCb].forEach(cb => { if (cb) cb.addEventListener('change', () => applyCritiqueFilters()); });
function applyCritiqueFilters() {
  // 简化实现：通过 CSS 类控制
  $outputContent.classList.toggle('hide-evidence', !($toggleEvidenceCb?.checked ?? true));
  $outputContent.classList.toggle('hide-fallacy', !($toggleFallacyCb?.checked ?? true));
  $outputContent.classList.toggle('hide-counter', !($toggleCounterCb?.checked ?? true));
}

// 清空功能已拆分为 $btnClearInput + $btnResetAll（见上方）

// ==================== 文件上传 ====================
$dropZone.addEventListener('click', () => $fileInput.click());
$fileInput.addEventListener('change', (e) => { if (e.target.files[0]) uploadFile(e.target.files[0]); });
['dragenter','dragover'].forEach(ev => $dropZone.addEventListener(ev, (e) => { e.preventDefault(); $dropZone.classList.add('drag-over'); }));
['dragleave','drop'].forEach(ev => $dropZone.addEventListener(ev, (e) => { e.preventDefault(); $dropZone.classList.remove('drag-over'); }));
$dropZone.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) uploadFile(f); });

async function uploadFile(file) {
  const ext = file.name.toLowerCase().split('.').pop();
  if (!['docx','xlsx','md','pdf','txt'].includes(ext)) { showFileStatus('不支持 .' + ext + ' 格式', 'error'); return; }
  if (file.size > 4 * 1024 * 1024) { showFileStatus('文件超过 4MB，请压缩后重试', 'error'); return; }
  showFileStatus(`⏳ 解析「${file.name}」…`, 'info');
  try {
    const fd = new FormData(); fd.append('file', file);
    const res = await fetch('/api/parse-file', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);
    $content.value = data.text;
    $content.dispatchEvent(new Event('input', { bubbles: true }));
    addFileToList(data.fileName, file.size);
    $fileStatus.classList.add('hidden'); // 隐藏旧状态栏，改由文件列表展示
  } catch (err) { showFileStatus(`❌ ${err.message}`, 'error'); }
}
function showFileStatus(text, type) {
  $fileStatus.textContent = text; $fileStatus.className = `file-status ${type}`; $fileStatus.classList.remove('hidden');
  const old = $fileStatus.querySelector('.file-remove'); if (old) old.remove();
  if (type === 'success') {
    const btn = document.createElement('span'); btn.className = 'file-remove'; btn.textContent = '✕ 清除';
    btn.addEventListener('click', () => { $content.value = ''; $fileStatus.classList.add('hidden'); $fileStatus.innerHTML = ''; $fileInput.value = ''; });
    $fileStatus.appendChild(btn);
  }
}

// ==================== 链接抓取 ====================
async function fetchUrlContent() {
  if ($btnFetchUrl.disabled) return;
  const url = $urlInput.value.trim();
  if (!url) return showUrlStatus('请输入链接', 'error');
  if (!/^https?:\/\/.+/.test(url)) return showUrlStatus('链接需以 http:// 或 https:// 开头', 'error');
  $btnFetchUrl.disabled = true; $btnFetchUrl.querySelector('.btn-fetch-text').textContent = '抓取中…'; $fetchSpinner.classList.remove('hidden');
  hideUrlStatus();
  try {
    const res = await fetch('/api/fetch-url', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ url }) });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);
    $content.value = data.text;
    $content.dispatchEvent(new Event('input', { bubbles: true }));
    showUrlStatus(`✅「${data.title}」(${data.source}，${data.charCount.toLocaleString()} 字)`, 'success');
  } catch (err) { showUrlStatus(`❌ ${err.message}`, 'error'); }
  finally { $btnFetchUrl.disabled = false; $btnFetchUrl.querySelector('.btn-fetch-text').textContent = '抓取'; $fetchSpinner.classList.add('hidden'); }
}
$btnFetchUrl.addEventListener('click', fetchUrlContent);
$urlInput.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); fetchUrlContent(); } });
function showUrlStatus(text, type) { $urlStatus.textContent = text; $urlStatus.className = `file-status ${type}`; $urlStatus.classList.remove('hidden'); }
function hideUrlStatus() { $urlStatus.classList.add('hidden'); }

// ==================== 角色扮演 ====================
async function submitRoleplay() {
  const reply = $userReply.value.trim();
  if (!reply) { showToast('⚠️ 请输入回复', 'warning'); return; }
  if (!conversationHistory.length) { showToast('⚠️ 请先完成转化', 'warning'); return; }
  setRpLoading(true);
  try {
    const res = await fetch('/api/roleplay', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ conversationHistory, userReply: reply, sessionId: currentSessionId }) });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);
    conversationHistory = data.conversationHistory;
    $rpResult.innerHTML = renderMarkdown(data.reply);
    $rpResult.classList.remove('hidden');
    $userReply.value = '';
    $feedbackArea.classList.remove('hidden');
    $feedbackThanks.classList.add('hidden');
    $feedbackBtns.forEach(b => { b.classList.remove('selected'); b.disabled = false; });
  } catch (err) { showToast(`❌ ${err.message}`, 'error'); }
  finally { setRpLoading(false); }
}
function setRpLoading(l) { $btnRoleplay.disabled = l; $btnRoleplay.querySelector('.btn-text').textContent = l ? '处理中…' : '提交回复'; $rpSpinner.classList.toggle('hidden', !l); }

// ==================== 反馈 ====================
$feedbackBtns.forEach(btn => btn.addEventListener('click', function() {
  const v = this.dataset.value;
  $feedbackBtns.forEach(b => { b.classList.toggle('selected', b.dataset.value === v); b.disabled = true; });
  const responses = { A:'🎉 太好了！继续把更多好内容变成真本事！', B:'🙏 感谢反馈！试试更具体描述场景，结果会更精准。', C:'🤔 能说说哪里没感觉吗？我会持续改进。', D:'📝 谢谢诚实反馈！能说一句原因吗？' };
  $feedbackThanks.textContent = responses[v] || responses.A;
  $feedbackThanks.classList.remove('hidden');
}));

// ==================== 事件绑定 ====================
$btnTransform.addEventListener('click', transformContent);
$btnRoleplay.addEventListener('click', submitRoleplay);
$content.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); transformContent(); } });
$userReply.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); submitRoleplay(); } });

// ==================== 历史记录 ====================
const HISTORY_KEY = 'capable_history';
const MAX_HISTORY = 10;
const $btnHistory = document.getElementById('btnHistory');
const $historyPanel = document.getElementById('historyPanel');
const $btnCloseHistory = document.getElementById('btnCloseHistory');
const $historyList = document.getElementById('historyList');
const $historySearch = document.getElementById('historySearch');
const $btnClearHistory = document.getElementById('btnClearHistory');

function getHistory() {
  try { return JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]'); }
  catch { return []; }
}
function saveHistory(entry) {
  let history = getHistory();
  // 去重：相同原文不重复保存
  history = history.filter(h => h.content !== entry.content);
  history.unshift(entry);
  if (history.length > MAX_HISTORY) history = history.slice(0, MAX_HISTORY);
  localStorage.setItem(HISTORY_KEY, JSON.stringify(history));
}
function addToHistory(content, identity, modes, results) {
  const entry = {
    id: Date.now(),
    content: content.slice(0, 500), // 截取前500字作为摘要
    fullContent: content,
    identity,
    modes,
    results, // 存储结果文本（截取前2000字）
    time: new Date().toLocaleString('zh-CN'),
  };
  saveHistory(entry);
}

function renderHistory(filter = '') {
  const history = getHistory();
  const q = filter.toLowerCase();
  const filtered = q ? history.filter(h => h.content.toLowerCase().includes(q) || h.identity.toLowerCase().includes(q)) : history;

  if (filtered.length === 0) {
    $historyList.innerHTML = '<div class="history-empty">暂无历史记录</div>';
    return;
  }
  $historyList.innerHTML = filtered.map(h => `
    <div class="history-item" data-id="${h.id}">
      <div class="history-item-time">${h.time}</div>
      <div class="history-item-preview">${h.content.slice(0, 80)}${h.content.length > 80 ? '…' : ''}</div>
      <div class="history-item-meta">
        ${h.identity ? `<span>👤 ${h.identity}</span>` : ''}
        <span>${h.modes?.join(', ') || ''}</span>
      </div>
    </div>
  `).join('');

  // 点击恢复
  $historyList.querySelectorAll('.history-item').forEach(item => {
    item.addEventListener('click', () => {
      const h = history.find(e => e.id === +item.dataset.id);
      if (!h) return;
      // 回填
      $content.value = h.fullContent || h.content;
      updateTransformBtn(); updateCharCount();
      if (h.identity) {
        $identityPreset.value = '__custom__';
        $identityPreset.dispatchEvent(new Event('change'));
        $userProfile.value = h.identity;
      }
      if (h.modes) {
        $modeSelector.querySelectorAll('input[name="mode"]').forEach(cb => cb.checked = h.modes.includes(cb.value));
        sortModeCards();
      }
      if (h.results) {
        // 尝试恢复结果展示
        try {
          const parsed = JSON.parse(h.results);
          if (typeof parsed === 'object' && parsed.summarize) showMultiResults(parsed);
          else showResults(h.results, h.modes[0] || 'capability');
        } catch {
          showResults(h.results, h.modes[0] || 'capability');
        }
      }
      $historyPanel.classList.add('hidden');
      showToast('✅ 已恢复历史记录', 'success');
    });
  });
}

$btnHistory.addEventListener('click', () => {
  $historyPanel.classList.toggle('hidden');
  if (!$historyPanel.classList.contains('hidden')) renderHistory();
});
$btnCloseHistory.addEventListener('click', () => $historyPanel.classList.add('hidden'));
$btnClearHistory.addEventListener('click', () => {
  localStorage.removeItem(HISTORY_KEY); renderHistory(); showToast('🗑️ 历史已清空', 'info');
});
$historySearch.addEventListener('input', () => renderHistory($historySearch.value));

// ==================== 快捷标签 ====================
const quickPresets = {
  '文档解析': { identity: '', modes: ['summarize', 'extract'] },
  '书本精读': { identity: '在校生', modes: ['summarize', 'capability'] },
  '课程刷题': { identity: '在校生', modes: ['capability'] },
  '报告复盘': { identity: '产品经理', modes: ['critique', 'summarize'] },
  '播客转练': { identity: '自媒体', modes: ['summarize', 'capability'] },
};
document.querySelectorAll('.quick-tag').forEach(btn => {
  btn.addEventListener('click', () => {
    const preset = quickPresets[btn.dataset.preset];
    if (!preset) return;
    // 设置身份
    if (preset.identity) {
      $identityPreset.value = preset.identity;
      $identityPreset.dispatchEvent(new Event('change'));
    }
    // 设置模式
    $modeSelector.querySelectorAll('input[name="mode"]').forEach(cb => cb.checked = preset.modes.includes(cb.value));
    sortModeCards();
    $modeSelector.querySelectorAll('.mode-card').forEach(c => {
      const cb = c.querySelector('input');
      c.classList.toggle('selected', cb.checked);
    });
    saveDraft();
    showToast(`✅ 已切换「${btn.dataset.preset}」模板`, 'info');
  });
});

// ==================== 收藏 ====================
const $btnFavorites = document.getElementById('btnFavorites');
$btnFavorites.addEventListener('click', () => {
  const identity = $identityPreset.value === '__custom__' ? $userProfile.value : $identityPreset.value;
  const modes = getSelectedModes();
  const fav = { identity, modes, time: Date.now() };
  const favs = JSON.parse(localStorage.getItem('capable_favorites') || '[]');
  // 去重
  const exists = favs.findIndex(f => f.identity === identity && f.modes.join(',') === modes.join(','));
  if (exists >= 0) {
    favs.splice(exists, 1);
    showToast('⭐ 已取消收藏', 'info');
  } else {
    favs.unshift(fav);
    if (favs.length > 10) favs.pop();
    showToast('⭐ 已收藏当前配置', 'success');
  }
  localStorage.setItem('capable_favorites', JSON.stringify(favs));
});

// ==================== 全选 / 清空模式 ====================
document.getElementById('btnSelectAll').addEventListener('click', () => {
  $modeSelector.querySelectorAll('input[name="mode"]').forEach(cb => { cb.checked = true; cb.closest('.mode-card')?.classList.add('selected'); });
  sortModeCards(); saveDraft();
});
document.getElementById('btnClearAll').addEventListener('click', () => {
  $modeSelector.querySelectorAll('input[name="mode"]').forEach(cb => { cb.checked = false; cb.closest('.mode-card')?.classList.remove('selected'); });
  sortModeCards(); saveDraft();
});

// ==================== 新建角色 / 导入 Prompt ====================
document.getElementById('btnNewRole').addEventListener('click', () => {
  const name = prompt('输入自定义角色名称：');
  if (!name) return;
  const prompt_ = prompt('输入该角色的专属 Prompt 指令（可选）：');
  $identityPreset.value = '__custom__';
  $identityPreset.dispatchEvent(new Event('change'));
  $userProfile.value = prompt_ ? `${name}：${prompt_}` : name;
  showToast(`✅ 已创建角色「${name}」`, 'success');
});
document.getElementById('btnImportPrompt').addEventListener('click', () => {
  const text = prompt('粘贴专属 Prompt 模板内容：');
  if (!text) return;
  $identityPreset.value = '__custom__';
  $identityPreset.dispatchEvent(new Event('change'));
  $userProfile.value = text;
  showToast('✅ Prompt 已导入', 'success');
});

// ==================== 自定义输出模板 ====================
document.getElementById('btnAddTemplate').addEventListener('click', () => {
  const name = prompt('输入自定义输出模板名称：');
  if (!name) return;
  showToast(`💡 自定义模板「${name}」功能开发中，敬请期待`, 'info');
});

// ==================== 空状态示例填充 ====================
const samples = {
  book: `高效沟通不是天赋，而是可以刻意练习的技能。\n\n真正的倾听不是被动地等对方说完，而是主动理解对方的立场和情绪。研究表明，在冲突沟通中，先花2分钟复述对方的观点再做回应，冲突解决率提升40%。\n\n沟通的三个核心原则：第一，先倾听再表达。第二，用具体事例代替抽象评价。不要说"你做事不靠谱"，而是说"上周三的报告中缺少了市场数据部分"。第三，在表达不同意见时，先肯定再补充。\n\n总结：高效沟通 = 深度倾听 + 具体反馈 + 合作姿态。`,
  report: `2025年Q2移动互联网行业报告摘要：\n\n1. 短视频用户规模达10.2亿，同比增长8.3%，日均使用时长128分钟。\n2. AI应用渗透率从Q1的12%跃升至Q2的34%，其中AIGC工具类增速最快。\n3. 电商直播GMV突破2.8万亿，品牌自播占比首次超过达人直播（52% vs 48%）。\n4. 本地生活服务订单量同比增长67%，三线及以下城市贡献了58%的增量。\n5. 用户付费意愿持续提升，内容付费用户规模达5.8亿，人均月消费47元。`,
  course: `费曼学习法的核心是以教代学。当你能用最简单的语言向一个完全不懂的人解释清楚一个概念时，你才真正理解了它。\n\n具体步骤：第一步选择要学习的概念，第二步想象你要把它教给一个孩子，第三步如果卡住了就回去重新学习，第四步用类比和简单语言简化。\n\n这个方法的关键在于暴露你的知识盲区。很多人以为自己懂了，但一开口就说不清楚，这就是没学透的信号。\n\n建议每天花15分钟用费曼学习法复习一个知识点，坚持21天，知识留存率可以从5%提升到90%。`,
};
document.querySelectorAll('.example-chip').forEach(chip => {
  chip.addEventListener('click', () => {
    const text = samples[chip.dataset.sample];
    if (text) { $content.value = text; $content.dispatchEvent(new Event('input', { bubbles: true })); showToast('✅ 示例内容已填充', 'success'); }
  });
});

// ==================== 输出标签栏切换 ====================
const $outputTabs = document.getElementById('outputTabs');
$outputTabs.addEventListener('click', (e) => {
  const tab = e.target.closest('.out-tab');
  if (!tab) return;
  $outputTabs.querySelectorAll('.out-tab').forEach(t => t.classList.remove('active'));
  tab.classList.add('active');
  const filter = tab.dataset.outtab;
  // 显示/隐藏对应模式的内容
  $outputContent.querySelectorAll('.multi-section, .tab-panel').forEach(el => {
    if (filter === 'all') { el.style.display = ''; }
    else { el.style.display = el.dataset.tab === filter || el.dataset.mode === filter ? '' : 'none'; }
  });
});

// ==================== 初始化 ====================
console.log('🧠 Capable v2 已就绪');

// ==================== QR Code 分享 ====================
const $btnQrCode = document.getElementById('btnQrCode');
const $qrModal = document.getElementById('qrModal');
const $qrImage = document.getElementById('qrImage');
const $qrUrlText = document.getElementById('qrUrlText');
const $btnCopyUrl = document.getElementById('btnCopyUrl');
const $qrModalClose = $qrModal.querySelector('.qr-modal-close');
const $qrModalBackdrop = $qrModal.querySelector('.qr-modal-backdrop');

function openQrModal() {
  const url = window.location.href;
  $qrImage.src = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(url)}`;
  $qrUrlText.textContent = url;
  $qrModal.classList.remove('hidden');
}

function closeQrModal() {
  $qrModal.classList.add('hidden');
  $qrImage.src = '';
}

$btnQrCode.addEventListener('click', openQrModal);
$qrModalClose.addEventListener('click', closeQrModal);
$qrModalBackdrop.addEventListener('click', closeQrModal);

$btnCopyUrl.addEventListener('click', async () => {
  try {
    await navigator.clipboard.writeText(window.location.href);
    showToast('✅ 链接已复制到剪贴板', 'success');
  } catch {
    // 降级方案
    $qrUrlText.select();
    document.execCommand('copy');
    showToast('✅ 链接已复制到剪贴板', 'success');
  }
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && !$qrModal.classList.contains('hidden')) {
    closeQrModal();
  }
});
