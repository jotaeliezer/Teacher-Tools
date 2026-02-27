// ==== Report Comments ====
function loadCommentConfigForFile(name){
  const s = loadSettings();
  const map = s.commentSettings || {};
  const cfg = (name && map[name]) ? map[name] : {};
  const merged = {...COMMENT_DEFAULTS, ...cfg};
  merged.extraColumns = Array.isArray(merged.extraColumns) ? [...merged.extraColumns] : [];
  if (merged.gradeColumn && !allColumns.includes(merged.gradeColumn)){
    merged.gradeColumn = "";
  }
  return merged;
}
  function persistCommentConfig(){
    if (!currentFileName) return;
    const s = loadSettings();
    const all = {...(s.commentSettings || {})};
    all[currentFileName] = {
      ...commentConfig,
      extraColumns: [...(commentConfig.extraColumns || [])]
    };
    saveSettings({ commentSettings: all });
  }
  function openCommentsModal(){
    if (!rows.length){
      status('Load data to open Report Comments.');
      return;
    }
    ensureDefaultCommentGradeColumn();
    const gradeCol = commentConfig.gradeColumn;
    if (!gradeCol){
      status('Select a grade column in Report Comments settings to continue.');
    }
    try{
      activateTab('comments');
    }catch(err){
      console.error('Failed to open Report Comments modal:', err);
      status('Could not open Report Comments. Check console for details.');
      closeCommentsModal();
    }
  }
  function openCommentsModalInternal(){
    populateCommentGradeOptions();
    buildCommentExtrasList();
    applyCommentSettingsToInputs();
    refreshCommentsPreview();
    initializeCommentBuilder();
  }
  function closeCommentsModal(){
    activateTab('data');
  }
  function populateCommentGradeOptions(){
    if (!commentGradeSelect) return;
    ensureDefaultCommentGradeColumn();
    const selected = commentConfig.gradeColumn || "";
    commentGradeSelect.innerHTML = '<option value="">Select column...</option>';
    allColumns.forEach(col => {
      const option = document.createElement('option');
      option.value = col;
      option.textContent = col;
      if (col === selected) option.selected = true;
      commentGradeSelect.appendChild(option);
    });
  }
  function buildCommentExtrasList(){
    if (!commentExtrasList) return;
    commentExtrasList.innerHTML = "";
    const dataColumns = allColumns.filter(col => columnHasData(col));
    if (!dataColumns.length){
      const msg = document.createElement('div');
      msg.className = 'muted';
      msg.textContent = 'Load data with marks to choose columns.';
      commentExtrasList.appendChild(msg);
      return;
    }
    const extrasSet = new Set((commentConfig.extraColumns || []).filter(col => dataColumns.includes(col)));
    if (commentConfig.extraColumns && commentConfig.extraColumns.length !== extrasSet.size){
      commentConfig.extraColumns = Array.from(extrasSet);
      persistCommentConfig();
    }
    dataColumns.forEach(col => {
      const label = document.createElement('label');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = extrasSet.has(col);
      checkbox.addEventListener('change', () => handleExtraColumnToggle(col, checkbox.checked));
      const span = document.createElement('span');
      span.textContent = col;
      label.appendChild(checkbox);
      label.appendChild(span);
      commentExtrasList.appendChild(label);
    });
  }
  function applyCommentSettingsToInputs(){
    if (!commentGradeSelect) return;
    ensureDefaultCommentGradeColumn();
    commentGradeSelect.value = commentConfig.gradeColumn || "";
    if (commentHighThresholdInput) commentHighThresholdInput.value = commentConfig.highThreshold;
    if (commentMidThresholdInput) commentMidThresholdInput.value = commentConfig.midThreshold;
    if (commentHighTemplateInput) commentHighTemplateInput.value = commentConfig.highTemplate;
    if (commentMidTemplateInput) commentMidTemplateInput.value = commentConfig.midTemplate;
    if (commentLowTemplateInput) commentLowTemplateInput.value = commentConfig.lowTemplate;
  }
  function handleCommentSettingsChange(){
    if (commentGradeSelect) commentConfig.gradeColumn = commentGradeSelect.value;
    if (commentHighThresholdInput){
      commentConfig.highThreshold = clampNumber(commentHighThresholdInput.value, COMMENT_DEFAULTS.highThreshold);
    }
    if (commentMidThresholdInput){
      commentConfig.midThreshold = clampNumber(commentMidThresholdInput.value, COMMENT_DEFAULTS.midThreshold);
    }
    ensureCommentThresholds();
    if (commentHighTemplateInput) commentConfig.highTemplate = commentHighTemplateInput.value || COMMENT_DEFAULTS.highTemplate;
    if (commentMidTemplateInput) commentConfig.midTemplate = commentMidTemplateInput.value || COMMENT_DEFAULTS.midTemplate;
    if (commentLowTemplateInput) commentConfig.lowTemplate = commentLowTemplateInput.value || COMMENT_DEFAULTS.lowTemplate;
    persistCommentConfig();
    syncActiveContext();
    refreshCommentsPreview();
  }
  function handleExtraColumnToggle(col, checked){
    if (!col) return;
    const next = new Set(commentConfig.extraColumns || []);
    if (checked) next.add(col);
    else next.delete(col);
    commentConfig.extraColumns = Array.from(next);
    persistCommentConfig();
    syncActiveContext();
    refreshCommentsPreview();
  }
  function ensureDefaultCommentGradeColumn(){
    if (commentConfig.gradeColumn && allColumns.includes(commentConfig.gradeColumn)) return;
    let candidate = allColumns.find(c => /^calculated final /i.test(c));
    if (!candidate){
      candidate = allColumns.find(c => isLikelyMarkColumn(c));
    }
    if (candidate){
      commentConfig.gradeColumn = candidate;
      persistCommentConfig();
      if (commentGradeSelect) commentGradeSelect.value = candidate;
    }
  }
  function restructureTemplate(kind){
    const nextTemplate = buildTemplateVariant(kind);
    if (!nextTemplate) return;
    switch(kind){
      case 'high':
        if (commentHighTemplateInput) commentHighTemplateInput.value = nextTemplate;
        commentConfig.highTemplate = nextTemplate;
        break;
      case 'mid':
        if (commentMidTemplateInput) commentMidTemplateInput.value = nextTemplate;
        commentConfig.midTemplate = nextTemplate;
        break;
      case 'low':
        if (commentLowTemplateInput) commentLowTemplateInput.value = nextTemplate;
        commentConfig.lowTemplate = nextTemplate;
        break;
      default:
        return;
    }
    persistCommentConfig();
    syncActiveContext();
    refreshCommentsPreview();
  }
  function buildTemplateVariant(kind){
    const pool = TEMPLATE_VARIANTS[kind];
    if (!pool || !pool.length) return '';
    if (templateIndices[kind] == null) templateIndices[kind] = 0;
    const idx = templateIndices[kind] % pool.length;
    templateIndices[kind] = (idx + 1) % pool.length;
    const choice = pool[idx];
    if (!Array.isArray(choice)) return String(choice || '');
    return choice.join(' ');
  }
  function ensureCommentThresholds(){
    if (commentConfig.highThreshold < commentConfig.midThreshold){
      commentConfig.highThreshold = commentConfig.midThreshold;
      if (commentHighThresholdInput) commentHighThresholdInput.value = commentConfig.highThreshold;
    }
    commentConfig.highThreshold = clampNumber(commentConfig.highThreshold, COMMENT_DEFAULTS.highThreshold);
    commentConfig.midThreshold = clampNumber(commentConfig.midThreshold, COMMENT_DEFAULTS.midThreshold);
  }
  function clampNumber(value, fallback){
    const num = Number(value);
    if (Number.isNaN(num)) return fallback;
    return Math.min(100, Math.max(0, num));
  }
  function refreshCommentsPreview(){
    if (!commentsPreviewEl) return;
    generatedComments = [];
    commentsPreviewEl.innerHTML = "";
    if (!rows.length){
      if (commentsStatusEl) commentsStatusEl.textContent = 'No data loaded.';
      if (commentsCountEl) commentsCountEl.textContent = '0';
      return;
    }
    ensureDefaultCommentGradeColumn();
    if (!commentConfig.gradeColumn){
      if (commentsStatusEl) commentsStatusEl.textContent = 'Select a mark column to generate comments.';
      if (commentsCountEl) commentsCountEl.textContent = '0';
      return;
    }
    const activeSet = visibleRowSet instanceof Set ? visibleRowSet : new Set(Array.from(rows.keys()));
    const indices = filteredIdx.length ? filteredIdx : Array.from(rows.keys());
    let missingMarks = 0;
    indices.forEach(idx => {
      if (!activeSet.has(idx)) return;
      const row = rows[idx];
      const raw = row?.[commentConfig.gradeColumn];
      const markMeta = deriveMarkMeta(raw, commentConfig.gradeColumn);
      const markValue = markMeta ? Number(markMeta.raw) : null;
      if (markValue == null) missingMarks++;
      const template = selectTemplateForMark(markValue);
      if (!template) return;
      const comment = fillTemplateWithRow(template, row, markValue, raw, idx);
      const name = getRowLabel(idx);
      // record index in generatedComments before pushing so we can reference it reliably
      const gIndex = generatedComments.length;
      generatedComments.push({ name, comment });
      const card = document.createElement('div');
      card.className = 'comment-card';
      const header = document.createElement('header');
      const title = document.createElement('strong');
      title.textContent = name;
      const markBadge = document.createElement('span');
      markBadge.textContent = markMeta ? `${markMeta.raw} %` : formatMarkDisplay(raw);
      const copyBtn = document.createElement('button');
      copyBtn.className = 'btn';
      copyBtn.style.marginLeft = 'auto';
      copyBtn.textContent = 'Copy';
      copyBtn.dataset.copyIndex = String(gIndex);
      header.appendChild(title);
      header.appendChild(markBadge);
      header.appendChild(copyBtn);
      card.appendChild(header);
      const textArea = document.createElement('textarea');
      textArea.value = comment;
      // store generated-comments index so edits update the correct entry
      textArea.dataset.genIndex = String(gIndex);
      textArea.addEventListener('input', () => {
        const gi = Number(textArea.dataset.genIndex);
        if (!Number.isNaN(gi) && generatedComments[gi]) generatedComments[gi].comment = textArea.value;
      });
      card.appendChild(textArea);
      commentsPreviewEl.appendChild(card);
    });
    if (commentsStatusEl){
      if (!generatedComments.length){
        commentsStatusEl.textContent = 'No marks found in the selected column.';
      }else if (missingMarks === generatedComments.length){
        commentsStatusEl.textContent = 'Marks are not in percent format. Continue anyway?';
      }else{
        commentsStatusEl.textContent = '';
      }
    }
    if (commentsCountEl){
      commentsCountEl.textContent = String(generatedComments.length);
    }
  }
  function selectTemplateForMark(mark){
    if (mark == null) return commentConfig.lowTemplate;
    if (mark >= commentConfig.highThreshold) return commentConfig.highTemplate;
    if (mark >= commentConfig.midThreshold) return commentConfig.midTemplate;
    return commentConfig.lowTemplate;
  }
  function fillTemplateWithRow(template, row, markNum, rawMark, rowIndex){
    const first = firstNameKey ? (row?.[firstNameKey] ?? '') : '';
    const last = lastNameKey ? (row?.[lastNameKey] ?? '') : '';
    const name = row?.[studentNameColumn] || buildStudentName(row) || getRowLabel(rowIndex);
    const displayMark = markNum != null ? `${markNum.toFixed(1)}%` : (rawMark != null ? String(rawMark) : '');
    const base = (template || '').replace(/\{(name|first|last|mark)\}/gi, (_, token) => {
      switch(token.toLowerCase()){
        case 'name': return name;
        case 'first': return first;
        case 'last': return last;
        case 'mark': return displayMark;
        default: return '';
      }
    });
    return (base + buildExtraAssessmentsText(row)).trim();
  }
  function buildExtraAssessmentsText(row){
    const cols = Array.isArray(commentConfig.extraColumns) ? commentConfig.extraColumns : [];
    if (!cols.length) return '';
    const parts = [];
    cols.forEach(col => {
      if (!col) return;
      const value = row?.[col];
      if (value == null || String(value).trim() === '') return;
      parts.push(`${col}: ${value}`);
    });
    if (!parts.length) return '';
    return ` Recent assessments: ${parts.join('; ')}.`;
  }
  function copyCommentByIndex(idx){
    if (idx == null || idx < 0 || idx >= generatedComments.length) return;
    const entry = generatedComments[idx];
    copyTextToClipboard(entry.comment);
    markCopyButton(idx);
  }
  function copyAllComments(){
    if (!generatedComments.length){
      if (commentsStatusEl) commentsStatusEl.textContent = 'No comments to copy.';
      return;
    }
    const bundle = generatedComments.map(entry => `${entry.name}: ${entry.comment}`).join('\n\n');
    copyTextToClipboard(bundle);
  }
  function markCopyButton(idx){
    const btn = commentsPreviewEl?.querySelector(`button[data-copy-index="${idx}"]`);
    if (!btn) return;
    btn.textContent = 'Copied';
    btn.classList.add('primary');
    btn.disabled = true;
    setTimeout(() => {
      btn.textContent = 'Copy';
      btn.classList.remove('primary');
      btn.disabled = false;
    }, 1500);
  }
  async function copyTextToClipboard(text){
    if (navigator && navigator.clipboard && navigator.clipboard.writeText){
      try{
        await navigator.clipboard.writeText(text);
        if (commentsStatusEl) commentsStatusEl.textContent = 'Copied to clipboard.';
      }catch{
        fallbackCopy(text);
      }
    }else{
      fallbackCopy(text);
    }
    setTimeout(() => {
      if (commentsStatusEl) commentsStatusEl.textContent = '';
    }, 2000);
  }
  function fallbackCopy(text){
    const helper = document.createElement('textarea');
    helper.value = text;
    document.body.appendChild(helper);
    helper.select();
    document.execCommand('copy');
    document.body.removeChild(helper);
    if (commentsStatusEl) commentsStatusEl.textContent = 'Copied.';
  }

  // Wire up comment-specific UI after functions are defined
  (function initCommentEvents(){
    const inputs = [commentGradeSelect, commentHighThresholdInput, commentMidThresholdInput,
      commentHighTemplateInput, commentMidTemplateInput, commentLowTemplateInput];
    inputs.forEach(el => {
      if (!el) return;
      const eventName = el.tagName === 'SELECT' ? 'change' : 'input';
      el.addEventListener(eventName, handleCommentSettingsChange);
    });
    if (commentsCopyAllBtn){
      commentsCopyAllBtn.addEventListener('click', copyAllComments);
    }
    if (commentHighRestructureBtn){
      commentHighRestructureBtn.addEventListener('click', () => restructureTemplate('high'));
    }
    if (commentMidRestructureBtn){
      commentMidRestructureBtn.addEventListener('click', () => restructureTemplate('mid'));
    }
    if (commentLowRestructureBtn){
      commentLowRestructureBtn.addEventListener('click', () => restructureTemplate('low'));
    }
    if (commentsPreviewEl){
      commentsPreviewEl.addEventListener('click', (e) => {
        const btn = e.target.closest('button[data-copy-index]');
        if (!btn) return;
        const idx = Number(btn.dataset.copyIndex);
        copyCommentByIndex(idx);
      });
    }
  }());
  function formatMarkDisplay(raw){
    if (raw == null || raw === '') return '—';
    const meta = deriveMarkMeta(raw, commentConfig.gradeColumn || '');
    return formatMarkText(meta, raw);
  }
  
