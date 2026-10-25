const $ = (selector) => document.querySelector(selector);
const state = { file: null, base64: null, sheetNames: [], result: null, busy: false };

const fields = [
  ['number', '序号'], ['platform', '平台', 'platform'], ['product', '产品'], ['title', '标题'],
  ['link', '链接'], ['createdAt', '创建时间', 'date'], ['owner', '帖子负责人'],
  ['likeScreenshot', '点赞截图'], ['likeCount', '点赞个数'],
];

function showToast(message, duration = 4200) {
  const toast = $('#toast');
  toast.textContent = message;
  toast.hidden = false;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => { toast.hidden = true; }, duration);
}

function setBusy(busy, label = '') {
  state.busy = busy;
  $('#processButton').disabled = busy || !state.base64;
  $('#processButton').textContent = busy ? label : '开始整理评论';
  $('#downloadButton').disabled = busy;
}

async function fileToBase64(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let binary = '';
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

async function postJSON(path, body) {
  const response = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.error || `请求失败（${response.status}）`);
  }
  return response;
}

async function chooseFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    showToast('请选择 .xlsx 格式的工作簿。');
    return;
  }
  if (file.size > 20 * 1024 * 1024) {
    showToast('文件超过 20 MB，请先压缩或拆分工作簿。');
    return;
  }
  state.file = file;
  state.result = null;
  $('#fileName').textContent = file.name;
  $('#emptyState').hidden = false;
  $('#resultView').hidden = true;
  $('#sheetPickers').hidden = true;
  setBusy(true, '读取工作表…');
  try {
    state.base64 = await fileToBase64(file);
    const response = await postJSON('/api/sheets', { workbookBase64: state.base64 });
    const info = await response.json();
    state.sheetNames = info.sheetNames;
    populateSelect($('#templateSheet'), info.sheetNames, info.defaultTemplate);
    populateSelect($('#sourceSheet'), info.sheetNames, info.defaultSource);
    $('#sheetPickers').hidden = false;
    $('#processButton').disabled = false;
  } catch (error) {
    state.file = null;
    state.base64 = null;
    $('#fileName').textContent = '还没有选择文件';
    showToast(error.message);
  } finally {
    setBusy(false);
  }
}

function populateSelect(select, names, chosen) {
  select.replaceChildren();
  names.forEach((name) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    option.selected = name === chosen;
    select.append(option);
  });
}

$('#fileInput').addEventListener('change', (event) => chooseFile(event.target.files[0]));
const dropzone = $('#dropzone');
for (const eventName of ['dragenter', 'dragover']) {
  dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.add('drag'); });
}
for (const eventName of ['dragleave', 'drop']) {
  dropzone.addEventListener(eventName, (event) => { event.preventDefault(); dropzone.classList.remove('drag'); });
}
dropzone.addEventListener('drop', (event) => chooseFile(event.dataTransfer.files[0]));

async function processFile() {
  if (!state.base64 || state.busy) return;
  setBusy(true, 'DeepSeek 正在拆分…');
  try {
    const response = await postJSON('/api/process', {
      workbookBase64: state.base64,
      templateSheet: $('#templateSheet').value,
      sourceSheet: $('#sourceSheet').value,
    });
    state.result = await response.json();
    renderResult();
    $('#emptyState').hidden = true;
    $('#resultView').hidden = false;
    const fallbackCount = state.result.posts.filter((post) => post.splitMode === 'local').length;
    if (fallbackCount) showToast(`有 ${fallbackCount} 篇笔记使用了本地规则拆分，请检查评论标签。`, 7000);
  } catch (error) {
    showToast(error.message, 7000);
  } finally {
    setBusy(false);
  }
}
$('#processButton').addEventListener('click', processFile);

function createMetaField(post, postIndex, key, label, type = '') {
  const wrapper = document.createElement('div');
  wrapper.className = 'meta-field';
  const title = document.createElement('label');
  title.textContent = label;
  const input = type === 'platform' ? document.createElement('select') : document.createElement('input');
  if (type === 'platform') {
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = '选择平台';
    input.append(placeholder);
    ['抖音', '小红书'].forEach((platform) => {
      const option = document.createElement('option');
      option.value = platform;
      option.textContent = platform;
      input.append(option);
    });
  } else if (type === 'date') {
    input.type = 'date';
  }
  input.value = type === 'date' ? String(post[key] ?? '').slice(0, 10) : (post[key] ?? '');
  input.setAttribute('aria-label', `${label}，第 ${postIndex + 1} 篇笔记`);
  input.addEventListener('input', () => { post[key] = input.value; });
  if (type === 'platform') input.addEventListener('change', () => { post[key] = input.value; });
  wrapper.append(title, input);
  return wrapper;
}

function createCommentRow(post, postIndex, comment, commentIndex) {
  const row = document.createElement('div');
  row.className = 'comment-row';
  const label = document.createElement('input');
  label.placeholder = '评论编号';
  label.value = comment.label || '';
  label.setAttribute('aria-label', `第 ${postIndex + 1} 篇笔记，第 ${commentIndex + 1} 条评论编号`);
  label.addEventListener('input', () => { comment.label = label.value; });
  const text = document.createElement('textarea');
  text.placeholder = '评论内容';
  text.value = comment.text || '';
  text.rows = Math.max(1, Math.min(4, Math.ceil(text.value.length / 55)));
  text.setAttribute('aria-label', `第 ${postIndex + 1} 篇笔记，第 ${commentIndex + 1} 条评论内容`);
  text.addEventListener('input', () => { comment.text = text.value; text.rows = Math.max(1, Math.min(5, Math.ceil(text.value.length / 55))); });
  const remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'remove-comment';
  remove.title = '删除这条评论';
  remove.setAttribute('aria-label', '删除评论');
  remove.textContent = '×';
  remove.addEventListener('click', () => {
    post.comments.splice(commentIndex, 1);
    renderResult();
  });
  row.append(label, text, remove);
  return row;
}

function renderResult() {
  const { posts, warnings } = state.result;
  $('#resultSub').textContent = `${state.file.name} · 修改后的内容会写入导出文件。`;
  const commentCount = posts.reduce((sum, post) => sum + post.comments.length, 0);
  const aiCount = posts.filter((post) => post.splitMode === 'deepseek').length;
  const stats = $('#resultStats');
  stats.replaceChildren();
  [[`${posts.length} 篇笔记`, posts.length], [`${commentCount} 条评论`, commentCount], [`AI 拆分`, `${aiCount}/${posts.length}`]].forEach(([label, value]) => {
    const item = document.createElement('span');
    item.className = 'stat';
    const strong = document.createElement('strong');
    strong.textContent = value;
    item.append(strong, document.createTextNode(` ${label}`));
    stats.append(item);
  });
  const notices = $('#notices');
  notices.replaceChildren();
  (warnings || []).forEach((warning) => {
    const notice = document.createElement('div');
    notice.className = 'notice';
    notice.textContent = `${warning} 页面仍提供可编辑的拆分结果。`;
    notices.append(notice);
  });

  const container = $('#posts');
  container.replaceChildren();
  posts.forEach((post, postIndex) => {
    const card = document.createElement('article');
    card.className = 'post-card';
    const heading = document.createElement('div');
    heading.className = 'post-heading';
    const number = document.createElement('span');
    number.className = 'post-number';
    number.textContent = post.number || postIndex + 1;
    const title = document.createElement('h3');
    title.textContent = post.title || '未填写标题';
    heading.append(number, title);
    const metadata = document.createElement('div');
    metadata.className = 'post-meta-grid';
    fields.forEach(([key, label, type]) => metadata.append(createMetaField(post, postIndex, key, label, type)));
    const commentsTitle = document.createElement('div');
    commentsTitle.className = 'comments-title';
    const titleText = document.createElement('span');
    titleText.textContent = '内容端评论';
    const hint = document.createElement('small');
    hint.textContent = `${post.comments.length} 条 · 可直接修改`;
    commentsTitle.append(titleText, hint);
    const commentList = document.createElement('div');
    post.comments.forEach((comment, commentIndex) => commentList.append(createCommentRow(post, postIndex, comment, commentIndex)));
    const add = document.createElement('button');
    add.type = 'button';
    add.className = 'add-comment';
    add.textContent = '+ 添加评论行';
    add.addEventListener('click', () => { post.comments.push({ label: '', text: '' }); renderResult(); });
    card.append(heading, metadata, commentsTitle, commentList, add);
    container.append(card);
  });
}

$('#downloadButton').addEventListener('click', async () => {
  if (!state.result || state.busy) return;
  setBusy(true, '正在生成文件…');
  $('#downloadButton').textContent = '正在生成…';
  try {
    const response = await postJSON('/api/export', {
      workbookBase64: state.base64,
      templateSheet: state.result.templateSheet,
      sourceSheet: state.result.sourceSheet,
      posts: state.result.posts,
    });
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    const base = state.file.name.replace(/\.xlsx$/i, '');
    link.download = `${base}_整理后.xlsx`;
    document.body.append(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showToast('整理后的工作簿已下载。');
  } catch (error) {
    showToast(error.message, 7000);
  } finally {
    $('#downloadButton').textContent = '↓　下载整理后的 Excel';
    setBusy(false);
  }
});

fetch('/api/status').then((response) => response.json()).then((status) => {
  if (status.apiKeyConfigured) {
    $('#modelDot').classList.add('ready');
    $('#modelStatus').textContent = `DeepSeek 已配置 · ${status.model}`;
    $('#statusDot').classList.add('ready');
    $('#statusText').textContent = '本机运行 · 模型已连接';
  } else {
    $('#modelStatus').textContent = '未配置 Key · 可用本地规则整理';
  }
}).catch(() => { $('#modelStatus').textContent = '无法读取模型配置'; });
