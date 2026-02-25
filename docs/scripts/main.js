// ==== Init from stored settings ====
  const initialSettings = loadSettings();
  setWrapHeaders(!!initialSettings.wrapHeaders, false);
  setDarkMode(!!initialSettings.darkMode, false);
  zoomLevel = typeof initialSettings.zoom === 'number' ? initialSettings.zoom : 1;
  setZoom(zoomLevel, false);
  markColorsEnabled = initialSettings.markColors !== false;
  if (markColorsToggle) markColorsToggle.checked = markColorsEnabled;
  dataColumnsOnly = !!initialSettings.dataColumnsOnly;
  updateColumnDataFilterButton();
  isTransposed = !!(initialSettings.transposed);
  updateTransposeButton();
  setFilesDrawerExpanded(!!initialSettings.filesDrawerExpanded, false);
  selectedTemplateId = resolveSavedPrintTemplate(initialSettings);
  printAllClasses = !!initialSettings.printAllClasses;
  builderAiEndpoint = String(initialSettings.builderAiEndpoint || '').trim();
  saveSettings({ printTemplateId: selectedTemplateId });

  let builderActiveCommentSection = null;
  let builderLookingAheadAutoKey = '';
  let builderLookingAheadManualKey = '';
  let builderLookingAheadAutoUpdating = false;
  let builderOutputAnimationMode = 'instant';
  let builderOutputAnimationToken = 0;
  let builderOutputAnimating = false;
  let builderLastFullOutput = '';
  let builderDiffClearTimer = null;
  let builderRevisedAnimationToken = 0;

// ==== Directory handle persistence ====
  const HANDLE_DB = 'teacher_tools_handles';
  const HANDLE_STORE = 'dirs';
  let savedDirectoryHandle = null;

  function isHandleStoreSupported(){
    return typeof indexedDB !== 'undefined' && isDirectoryPickerSupported();
  }
  function openHandleDb(){
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(HANDLE_DB, 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(HANDLE_STORE)){
          db.createObjectStore(HANDLE_STORE);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  async function saveDirectoryHandle(handle){
    if (!handle || !isHandleStoreSupported()) return;
    const db = await openHandleDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(HANDLE_STORE, 'readwrite');
      tx.objectStore(HANDLE_STORE).put(handle, 'last');
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }
  async function loadDirectoryHandleFromStore(){
    if (!isHandleStoreSupported()) return null;
    const db = await openHandleDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(HANDLE_STORE, 'readonly');
      const req = tx.objectStore(HANDLE_STORE).get('last');
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  }
  async function verifyDirectoryPermission(handle){
    if (!handle) return false;
    const perm = await handle.queryPermission({ mode:'read' });
    if (perm === 'granted') return true;
    if (perm === 'prompt'){
      const res = await handle.requestPermission({ mode:'read' });
      return res === 'granted';
    }
    return false;
  }
  function updateSavedFolderLabel(handle){
    if (!lastFolderLabel) return;
    if (handle){
      lastFolderLabel.textContent = `Saved folder: ${handle.name}`;
      lastFolderLabel.style.color = '';
    }else{
      lastFolderLabel.textContent = 'No saved folder.';
      lastFolderLabel.style.color = '#888';
    }
  }
  async function initSavedDirectoryHandle(){
    if (!isHandleStoreSupported()){
      updateSavedFolderLabel(null);
      return;
    }
    try{
      const handle = await loadDirectoryHandleFromStore();
      if (handle && await verifyDirectoryPermission(handle)){
        savedDirectoryHandle = handle;
        updateSavedFolderLabel(handle);
        return;
      }
    }catch(err){
      console.warn('Could not load saved folder handle', err);
    }
    savedDirectoryHandle = null;
    updateSavedFolderLabel(null);
  }

// ==== Dropzone logic ====
  dropzone.addEventListener('click', () => fileInput.click());
  ;['dragenter','dragover'].forEach(ev =>
    dropzone.addEventListener(ev, (e) => { e.preventDefault(); e.stopPropagation(); dropzone.classList.add('hover'); })
  );
  ;['dragleave','dragend','drop'].forEach(ev =>
    dropzone.addEventListener(ev, (e) => { e.preventDefault(); e.stopPropagation(); dropzone.classList.remove('hover'); })
  );
  dropzone.addEventListener('drop', async (e) => {
    const files = Array.from(e.dataTransfer?.files || []);
    if (!files.length) return;
    await handleFiles(files);
  });
  fileInput.addEventListener('change', async (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    await handleFiles(files);
    e.target.value = "";
  });
  if (loadDirBtn){
    loadDirBtn.addEventListener('click', async () => {
      if (!isDirectoryPickerSupported()){
        status('Folder loading requires a supported browser (Chrome/Edge).');
        return;
      }
      try{
        loadDirBtn.disabled = true;
        if (savedDirectoryHandle){
          const confirmed = await showFolderPrompt(`Load CSV/XLSX files from saved folder "${savedDirectoryHandle.name}"?`);
          if (confirmed){
            status('Reading saved folder...');
            const loaded = await loadFilesFromDirectoryHandle(savedDirectoryHandle);
            if (loaded){
              status('READY');
              return;
            }
            status('Saved folder unavailable, pick a new one.');
          }
        }
        status('Select a folder...');
        await handleDirectoryPicker();
        status('READY');
      }catch(err){
        console.error(err);
        status('Folder load canceled or failed.');
      }finally{
        loadDirBtn.disabled = false;
      }
    });
  }
  initSavedDirectoryHandle();
  wrapHeadersToggle.addEventListener('change', (e) => {
    setWrapHeaders(e.target.checked, true);
    render();
  });
  if (markColorsToggle){
    markColorsToggle.addEventListener('change', (e) => {
      if (e.target.checked && rows.length && !datasetHasPercentSymbol()){
        e.target.checked = false;
        showMarkWarning('Your marks are not in percent format. Continue anyway?', () => applyMarkColorSetting(true));
        return;
      }
      applyMarkColorSetting(!!e.target.checked);
    });
  }
  if (tabDataBtn){
    tabDataBtn.addEventListener('click', () => activateTab('data'));
  }
  if (tabPrintBtn){
    tabPrintBtn.addEventListener('click', () => {
      if (!rows.length){
        status('Load data before printing.');
        return;
      }
      activateTab('print');
    });
  }
  if (tabCommentsBtn){
    tabCommentsBtn.addEventListener('click', () => {
      if (!rows.length){
        status('Load data before generating comments.');
        return;
      }
      activateTab('comments');
    });
  }
  if (commentsCloseBtn){
    commentsCloseBtn.addEventListener('click', () => activateTab('data'));
  }
  [
    [commentTermAllBtn, 'all'],
    [commentTermT1Btn, 'T1'],
    [commentTermT2Btn, 'T2'],
    [commentTermT3Btn, 'T3']
  ].forEach(([btn, term]) => {
    if (btn){
      btn.addEventListener('click', () => setCommentTerm(term));
    }
  });
  if (printCommentsDrawer){
    printCommentsDrawer.addEventListener('click', (e) => e.stopPropagation());
  }
  setCommentTerm(commentTermFilter);
  if (markWarningContinue){
    markWarningContinue.addEventListener('click', () => {
      closeMarkWarningModal();
      const action = pendingMarkWarningAction;
      pendingMarkWarningAction = null;
      if (typeof action === 'function') action();
    });
  }
  if (markWarningCancel){
    markWarningCancel.addEventListener('click', () => {
      closeMarkWarningModal();
      pendingMarkWarningAction = null;
    });
  }
  [printTeacherInput, printClassInput, printTitleInput].forEach(input => {
    if (!input) return;
    input.addEventListener('input', handlePrintMetaInputChange);
  });
  if (builderReportOutput && builderReportOverlay){
    builderReportOutput.addEventListener('scroll', () => {
      builderReportOverlay.scrollTop = builderReportOutput.scrollTop;
      builderReportOverlay.scrollLeft = builderReportOutput.scrollLeft;
    });
  }
  printTermInputs.forEach(radio => {
    radio.addEventListener('change', handlePrintMetaInputChange);
  });
  printTemplateInputs.forEach(input => {
    input.addEventListener('change', handleTemplateSelectionChange);
  });
  if (printAllClassesToggle){
    printAllClassesToggle.addEventListener('change', () => {
      printAllClasses = !!printAllClassesToggle.checked;
      saveSettings({ printAllClasses });
      selectedClassIds = new Set(getEffectivePrintContextIds());
      renderPrintPreview();
      renderPrintMarkingColumnPanel();
    });
  }
  if (printMarkingSelectAllBtn){
    printMarkingSelectAllBtn.addEventListener('click', handlePrintMarkingSelectAll);
  }
  if (printMarkingClearBtn){
    printMarkingClearBtn.addEventListener('click', handlePrintMarkingClear);
  }
  [printPreviewCloseBtn, printPreviewCancelBtn].forEach(btn => {
    if (btn){
      btn.addEventListener('click', () => {
        closePrintPreviewModal();
      });
    }
  });
  [printPreviewPrintTopBtn, printPreviewPrintBtn].forEach(btn => {
    if (!btn) return;
    btn.addEventListener('click', () => {
      launchPrintDialog();
    });
  });
  if (printCommentsBtn){
    printCommentsBtn.addEventListener('click', () => togglePrintCommentsDrawer(true));
  }
  if (printCommentsClose){
    printCommentsClose.addEventListener('click', () => togglePrintCommentsDrawer(false));
  }
  [printCommentsPronounMale, printCommentsPronounFemale].forEach(btn => {
    if (!btn) return;
    btn.addEventListener('change', () => {
      reportCardPronoun = btn.value === 'female' ? 'female' : 'male';
      updatePrintPronounButtons();
      buildPrintCommentsPanel();
      renderPrintPreview();
    });
  });
  if (colSearchInput){
    colSearchInput.addEventListener('input', (e) => {
      columnFilterText = e.target.value || '';
      rebuildColsUI();
    });
  }
  if (columnDataFilterBtn){
    columnDataFilterBtn.addEventListener('click', () => {
      setDataColumnsOnly(!dataColumnsOnly);
    });
  }
  if (rowFilterGoodBtn){
    rowFilterGoodBtn.addEventListener('click', () => handlePerformanceFilterRequest('high'));
  }
  if (rowFilterMidBtn){
    rowFilterMidBtn.addEventListener('click', () => handlePerformanceFilterRequest('mid'));
  }
  if (rowFilterNeedsBtn){
    rowFilterNeedsBtn.addEventListener('click', () => handlePerformanceFilterRequest('low'));
  }
  if (rowFilterClearBtn){
    rowFilterClearBtn.addEventListener('click', () => clearPerformanceFilter());
  }
  if (builderStudentSelect){
    builderStudentSelect.addEventListener('change', () => {
      autoSaveCurrentReport('student-switch');
      const idx = Number(builderStudentSelect.value);
      if (!Number.isNaN(idx)) prefillBuilderFromRow(idx);
    });
  }
  // "Use Data Values" button removed; keep display name/pronouns always synced to selection.
  if (builderCorePerformanceSelect){
    builderCorePerformanceSelect.addEventListener('change', () => {
      updatePerformanceSelectStyle();
      applyRandomLookingAheadSelection();
      if (shouldAutoGenerateBuilderReport()) builderGenerateReport();
    });
  }
  if (builderGenerateBtn){
    builderGenerateBtn.addEventListener('click', () => {
      builderOutputAnimationMode = 'generate';
      builderGenerateReport();
    });
  }
  if (builderGenerateAiBtn){
    builderGenerateAiBtn.addEventListener('click', builderGenerateReportWithAI);
  }
  if (builderCopyBtn){
    builderCopyBtn.addEventListener('click', () => {
      builderCopyReport();
      saveCurrentReport('copy');
    });
  }
  if (builderClearBtn){
    builderClearBtn.addEventListener('click', clearSelectedComments);
  }
  if (builderSaveBtn){
    builderSaveBtn.addEventListener('click', () => saveCurrentReport('manual'));
  }
  if (savedReportsClearBtn){
    savedReportsClearBtn.addEventListener('click', () => {
      savedReports = [];
      persistSavedReports();
      renderSavedReports();
    });
  }
  function toggleSavedReports(open){
    if (!savedReportsPanel || !savedReportsBackdrop) return;
    const show = open != null ? !!open : !savedReportsPanel.classList.contains('open');
    if (show){
      savedReportsPanel.style.display = 'flex';
      // trigger transition
      requestAnimationFrame(() => savedReportsPanel.classList.add('open'));
      savedReportsBackdrop.style.display = 'block';
      renderSavedReports();
    }else{
      savedReportsPanel.classList.remove('open');
      savedReportsBackdrop.style.display = 'none';
      setTimeout(() => {
        savedReportsPanel.style.display = 'none';
      }, 250);
    }
  }
  if (savedReportsOpenBtn){
    savedReportsOpenBtn.addEventListener('click', () => toggleSavedReports(true));
  }
  [savedReportsCloseBtn, savedReportsBackdrop].forEach(el => {
    if (!el) return;
    el.addEventListener('click', () => toggleSavedReports(false));
  });
  if (savedReportsListEl){
    savedReportsListEl.addEventListener('click', (e) => {
      const btn = e.target.closest('button[data-action]');
      if (!btn) return;
      const id = btn.dataset.id;
      const action = btn.dataset.action;
      const entry = savedReports.find(r => r.id === id);
      if (!entry) return;
      if (action === 'copy'){
        copyTextToClipboard(entry.text);
        status('Saved comment copied.');
      }else if (action === 'restore'){
        restoreSavedReport(entry);
        status('Saved comment restored to builder.');
      }else if (action === 'delete'){
        savedReports = savedReports.filter(r => r.id !== id);
        persistSavedReports();
        renderSavedReports();
      }
    });
  }
  function updateGradeGroupControls(){
    if (!builderIncludeFinalGradeInput) return;
    const group = builderGradeGroupSelect?.value || 'middle';
    const isElem = group === 'elem';
    const wrapper = builderIncludeFinalGradeInput.closest('.builder-inline-toggle');
    if (wrapper){
      wrapper.style.display = isElem ? '' : 'none';
    }
    builderIncludeFinalGradeInput.disabled = !isElem;
  }

function setupSelectAnimations(){
    const selects = [builderStudentSelect, builderGradeGroupSelect, builderCorePerformanceSelect, builderTermSelector].filter(Boolean);
    selects.forEach(sel => {
      sel.addEventListener('focus', () => sel.classList.add('is-open'));
      sel.addEventListener('blur', () => sel.classList.remove('is-open'));
      sel.addEventListener('change', () => sel.classList.remove('is-open'));
      sel.addEventListener('mousedown', () => sel.classList.add('is-open'));
    });
  }

  setupSelectAnimations();
  updateGradeGroupControls();
  if (builderCommentOrderToggle){
    builderCommentOrderToggle.addEventListener('click', toggleCommentOrderMode);
  }
  if (builderTermSelector){
    builderTermSelector.addEventListener('change', () => {
      if (shouldAutoGenerateBuilderReport()) builderGenerateReport();
    });
  }
      if (builderIncludeFinalGradeInput){
    builderIncludeFinalGradeInput.addEventListener('change', () => {
      if (shouldAutoGenerateBuilderReport()) builderGenerateReport();
    });
  }
if (builderGradeGroupSelect){
    builderGradeGroupSelect.addEventListener('change', () => {
      updateGradeGroupControls();
      if (shouldAutoGenerateBuilderReport()) builderGenerateReport();
    });
  }
[builderPronounMaleInput, builderPronounFemaleInput].forEach(el => {
    if (!el) return;
    el.addEventListener('change', () => {
      if (shouldAutoGenerateBuilderReport()) builderGenerateReport();
    });
  });
  if (darkModeBtn){
    darkModeBtn.addEventListener('click', () => {
      setDarkMode(!darkModeEnabled);
    });
  }
  if (filesDrawerToggle){
    filesDrawerToggle.addEventListener('click', () => {
      const expanded = !(filesDrawer?.classList.contains('collapsed'));
      setFilesDrawerExpanded(!expanded, true);
    });
  }
  if (zoomInBtn){
    zoomInBtn.addEventListener('click', () => setZoom(zoomLevel + ZOOM_STEP));
  }
  if (zoomOutBtn){
    zoomOutBtn.addEventListener('click', () => setZoom(zoomLevel - ZOOM_STEP));
  }
  renameCancel.addEventListener('click', () => closeRenameModal());
  renameSave.addEventListener('click', saveRename);
  renamePrompt.addEventListener('click', (e) => { if (e.target === renamePrompt) closeRenameModal(); });
  renameInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter'){ e.preventDefault(); saveRename(); }
  });
  transposeBtn.addEventListener('click', () => {
    isTransposed = !isTransposed;
    updateTransposeButton();
    saveSettings({ transposed: isTransposed });
    if (activeContext) activeContext.isTransposed = isTransposed;
    render();
  });
  colsDiv.addEventListener('dragover', (e) => {
    if (!draggingCol) return;
    e.preventDefault();
  });
  colsDiv.addEventListener('drop', (e) => {
    if (!draggingCol) return;
    if (e.target.closest && e.target.closest('.col-row')) return;
    e.preventDefault();
    moveColumnToEnd(draggingCol);
  });

  // ==== Keyboard (Undo/Redo) ====
  document.addEventListener('keydown', (e) => {
    const z = (e.key === 'z' || e.key === 'Z');
    const y = (e.key === 'y' || e.key === 'Y');
    if ((e.ctrlKey || e.metaKey) && z) { e.preventDefault(); doUndo(); }
    if ((e.ctrlKey || e.metaKey) && y) { e.preventDefault(); doRedo(); }
  });

  // ==== Persistence ====
  function loadSettings(){
    try { return JSON.parse(localStorage.getItem(LS_KEY) || "{}"); } catch { return {}; }
  }
  function saveSettings(data){
    localStorage.setItem(LS_KEY, JSON.stringify({...loadSettings(), ...data}));
  }
  function loadSavedReports(){
    const s = loadSettings();
    savedReports = Array.isArray(s.savedReports) ? s.savedReports : [];
    renderSavedReports();
  }
  function persistSavedReports(){
    const trimmed = savedReports.slice(0, SAVED_REPORT_LIMIT);
    saveSettings({ savedReports: trimmed });
  }

  // ==== Main file handler ====
  async function handleFiles(files){
    for (const file of files){
      await handleFile(file);
    }
  }

  function isDirectoryPickerSupported(){
    return typeof window.showDirectoryPicker === 'function';
  }
  function showFolderPrompt(message){
    return new Promise((resolve) => {
      if (!folderPrompt || !folderPromptConfirm || !folderPromptCancel || !folderPromptText){
        const fallback = confirm(message);
        resolve(fallback);
        return;
      }
      folderPromptText.textContent = message;
      folderPrompt.style.display = 'flex';
      const cleanup = () => {
        folderPrompt.style.display = 'none';
        folderPromptConfirm.removeEventListener('click', onConfirm);
        folderPromptCancel.removeEventListener('click', onCancel);
        folderPrompt.removeEventListener('click', backdropHandler);
      };
      const onConfirm = () => { cleanup(); resolve(true); };
      const onCancel = () => { cleanup(); resolve(false); };
      const backdropHandler = (e) => {
        if (e.target === folderPrompt) onCancel();
      };
      folderPromptConfirm.addEventListener('click', onConfirm);
      folderPromptCancel.addEventListener('click', onCancel);
      folderPrompt.addEventListener('click', backdropHandler);
    });
  }
  loadSavedReports();
  async function handleDirectoryPicker(){
    if (!isDirectoryPickerSupported()) throw new Error('Directory picker not supported');
    const dirHandle = await window.showDirectoryPicker();
    await useDirectoryHandle(dirHandle, true);
  }
  async function loadFilesFromDirectoryHandle(handle){
    const files = [];
    for await (const entry of handle.values()){
      if (entry.kind !== 'file') continue;
      const name = (entry.name || '').toLowerCase();
      if (!name.endsWith('.csv') && !name.endsWith('.xlsx') && !name.endsWith('.xls')) continue;
      const file = await entry.getFile();
      files.push(file);
    }
    if (!files.length){
      status('No CSV/XLSX files found in that folder.');
      return false;
    }
    await handleFiles(files);
    return true;
  }
  async function useDirectoryHandle(handle, persist){
    if (!handle) return;
    const ok = await verifyDirectoryPermission(handle);
    if (!ok){
      status('Folder permission denied.');
      return;
    }
    const loaded = await loadFilesFromDirectoryHandle(handle);
    if (loaded && persist){
      savedDirectoryHandle = handle;
      updateSavedFolderLabel(handle);
      await saveDirectoryHandle(handle);
    }
  }

  async function handleFile(f){
    if (activeContext){
      activeContext.searchText = searchEl.value || "";
      activeContext.isTransposed = isTransposed;
    }
    currentFileName = f.name;
    commentConfig = loadCommentConfigForFile(currentFileName);
    status('READING...');
    const name = f.name.toLowerCase();
    try{
      rows = [];
      originalRows = [];
      allColumns = [];
      columnOrder = [];
      visibleColumns = [];
      visibleRowSet = new Set();
      filteredIdx = [];
      sortState = {key:null, dir:null};
      undoStack = [];
      redoStack = [];
      if (name.endsWith('.csv')) await loadCSV(f);
      else await loadXLSX(f);
      originalRows = JSON.parse(JSON.stringify(rows));
      applySavedHeaderMapping();
      initColumns();
      initRows();
      setupStudentNameColumn();
      applyDefaultStudentSort();
      if (!commentConfig.gradeColumn && allColumns.includes('Calculated Final Grade')){
        commentConfig.gradeColumn = 'Calculated Final Grade';
        persistCommentConfig();
      }
      const s = loadSettings();
      isTransposed = !!s.transposed;
      updateTransposeButton();
      if (Object.prototype.hasOwnProperty.call(s, 'wrapHeaders')) {
        setWrapHeaders(!!s.wrapHeaders, false);
      }
      if (Array.isArray(s.columnOrder)){
        applyStoredColumnOrder(s.columnOrder);
      }
      if (Array.isArray(s.visibleColumns)) {
        visibleColumns = s.visibleColumns.filter(c => allColumns.includes(c));
        if (!visibleColumns.length){
          visibleColumns = columnOrder.slice(0, Math.min(8, columnOrder.length));
        }
      }
      syncVisibleOrder();
      const rowPrefs = s.visibleRows || {};
      if (currentFileName && Object.prototype.hasOwnProperty.call(rowPrefs, currentFileName)){
        const stored = Array.isArray(rowPrefs[currentFileName]) ? rowPrefs[currentFileName] : [];
        setVisibleRowsFromList(stored, true);
      }else{
        persistRowVisibility();
      }
      if (s.lastSearch != null){ searchEl.value = s.lastSearch; }
      rebuildColsUI();
      render();
    resetPerformanceFilterState();
    setDataActionButtonsDisabled(false);
    editHeadersBtn.disabled = false;
    transposeBtn.disabled = false;
    metaEl.innerText = `${f.name} - ${rows.length} rows, ${allColumns.length} columns`;
    status('READY');
    registerContext(createContextRecord(f.name));
    // Reset comment/builder UI so it repopulates for the new file
    builderBankRendered = false;
    builderSelectedRowIndex = null;
    if (isCommentsTabActive()){
      openCommentsModalInternal();
    }
  }catch(err){
    status('ERROR: ' + err); console.error(err);
  }
  }

  // ==== Loaders ====
  async function loadCSV(file) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => { rows = res.data; resolve(); },
        error: (err) => reject(err)
      });
    });
  }
  async function loadXLSX(file) {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data);
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  }

  // ==== Columns / Render ====
  function initColumns(){
    allColumns = rows.length ? Object.keys(rows[0]) : [];
    columnOrder = [...allColumns];
    visibleColumns = columnOrder.slice(0, Math.min(8, columnOrder.length));
    enforceEndOfLineColumn();
    ensureDefaultCommentGradeColumn();
    filteredIdx = Array.from(rows.keys());
    rebuildColsUI();
  }
  function getColumnPanelSortRank(name, originalIndex){
    const lower = (name || '').toLowerCase();
    if (name === STUDENT_NAME_LABEL) return { rank: 0, lesson: null, idx: originalIndex };
    if (lower === 'orgid') return { rank: 1, lesson: null, idx: originalIndex };
    const match = /^l\s*(\d+)/i.exec(name || '');
    if (match){
      return { rank: 2, lesson: Number(match[1]) || 0, idx: originalIndex };
    }
    return { rank: 3, lesson: null, idx: originalIndex };
  }
  function rebuildColsUI(){
    colsDiv.innerHTML = "";
    const query = (columnFilterText || '').toLowerCase().trim();
    const filterData = !!dataColumnsOnly;
    let rendered = 0;
    const sortedColumns = columnOrder
      .map((name, idx) => ({ name, idx, sort: getColumnPanelSortRank(name, idx) }))
      .sort((a, b) => {
        if (a.sort.rank !== b.sort.rank) return a.sort.rank - b.sort.rank;
        if (a.sort.rank === 2 && b.sort.rank === 2){
          if (a.sort.lesson !== b.sort.lesson) return a.sort.lesson - b.sort.lesson;
        }
        return a.sort.idx - b.sort.idx;
      })
      .map(entry => entry.name);
    sortedColumns.forEach(c => {
      const name = c || '';
      if (query && !name.toLowerCase().includes(query)) return;
      const isEndIndicator = (c === END_OF_LINE_COL);
      const hasData = (!isEndIndicator && rows.length) ? columnHasData(c) : false;
      if (filterData && !isEndIndicator && !hasData) return;
      const id = 'c_' + c.replace(/[^a-z0-9_]/gi, '_');
      const row = document.createElement('div');
      row.className = 'col-row';
      if (hasData) row.classList.add('has-data');
      row.draggable = !isEndIndicator;
      row.dataset.col = c;
      row.innerHTML = `
        <label style="flex:1; display:flex; align-items:center; gap:6px;">
          <input type="checkbox" id="${id}" ${visibleColumns.includes(c) ? 'checked':''} ${isEndIndicator ? 'disabled':''}/>
          <span>${escapeHtml(c)}</span>
        </label>
        <span class="muted" style="font-size:11px; user-select:none;">&#8645;</span>
      `;
      colsDiv.appendChild(row);

      const checkbox = row.querySelector('input');
      if (!isEndIndicator){
        checkbox.addEventListener('change', (ev) => {
          const before = [...visibleColumns];
          if (ev.target.checked) {
            if (!visibleColumns.includes(c)) visibleColumns.push(c);
          } else {
            visibleColumns = visibleColumns.filter(x => x !== c);
        }
        syncVisibleOrder();
        pushUndo({type:'setVisibleColumns', before, after:[...visibleColumns]});
        saveSettings({ visibleColumns });
        colCountEl.textContent = `${visibleColumns.length}/${allColumns.length}`;
          render();
        });
      }

      if (!isEndIndicator){
        row.addEventListener('dragstart', (e) => {
          draggingCol = c;
          row.classList.add('dragging');
          if (e.dataTransfer){
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', c);
          }
        });
        row.addEventListener('dragend', () => {
          draggingCol = null;
          row.classList.remove('dragging');
        });
        row.addEventListener('dragover', (e) => {
          if (!draggingCol || draggingCol === c) return;
          e.preventDefault();
          e.stopPropagation();
          row.classList.add('drag-over');
        });
        row.addEventListener('dragleave', () => row.classList.remove('drag-over'));
        row.addEventListener('drop', (e) => {
          if (!draggingCol || draggingCol === c) return;
          e.preventDefault();
          e.stopPropagation();
          row.classList.remove('drag-over');
          reorderColumns(draggingCol, c);
        });
      } else {
        row.classList.add('muted');
      }
      rendered++;
    });
    if (!rendered){
      const empty = document.createElement('div');
      empty.className = 'muted';
      empty.style.fontSize = '12px';
      empty.textContent = filterData ? 'No columns contain data yet.' : 'No columns match search.';
      colsDiv.appendChild(empty);
    }
    colCountEl.textContent = `${visibleColumns.length}/${allColumns.length}`;
  }
  function setDataColumnsOnly(enabled){
    const next = !!enabled;
    if (dataColumnsOnly === next) return;
    dataColumnsOnly = next;
    updateColumnDataFilterButton();
    saveSettings({ dataColumnsOnly });
    rebuildColsUI();
  }
  function updateColumnDataFilterButton(){
    if (!columnDataFilterBtn) return;
    columnDataFilterBtn.textContent = `Data Columns: ${dataColumnsOnly ? 'ON' : 'OFF'}`;
    columnDataFilterBtn.classList.toggle('active', dataColumnsOnly);
  }
  function initRows(){
    visibleRowSet = new Set(Array.from(rows.keys()));
    rebuildRowsUI();
  }
  function rebuildRowsUI(){
    if (!(visibleRowSet instanceof Set)){
      visibleRowSet = new Set(Array.from(rows.keys()));
    }
    rowsDiv.innerHTML = "";
    const entries = rows.map((_, idx) => ({
      idx,
      label: (getRowLabel(idx) || '').toLowerCase()
    }));
    entries.sort((a,b) => {
      const la = a.label;
      const lb = b.label;
      if (la < lb) return -1;
      if (la > lb) return 1;
      return a.idx - b.idx;
    });
    entries.forEach(entry => {
      const idx = entry.idx;
      const labelText = getRowLabel(idx);
      const rowEl = document.createElement('div');
      rowEl.className = 'row-item';
      const label = document.createElement('label');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = visibleRowSet.has(idx);
      label.appendChild(checkbox);
      const span = document.createElement('span');
      span.textContent = labelText;
      label.appendChild(span);
      rowEl.appendChild(label);
      rowsDiv.appendChild(rowEl);

      checkbox.addEventListener('change', () => {
        const before = getVisibleRowList();
        if (checkbox.checked) visibleRowSet.add(idx);
        else visibleRowSet.delete(idx);
        const after = getVisibleRowList();
        if (!arraysEqual(before, after)){
          pushUndo({type:'setVisibleRows', beforeRows: before, afterRows: after});
        }
        persistRowVisibility();
        rowCountEl.textContent = `${visibleRowSet.size}/${rows.length}`;
        render();
        performanceFilterLevel = null;
        updatePerformanceFilterButtons();
        updateRowFilterStatus('');
      });
    });
    rowCountEl.textContent = `${visibleRowSet.size}/${rows.length}`;
  }
  function setAllRowsVisible(pushUndoEntry = true){
    if (!(visibleRowSet instanceof Set)) visibleRowSet = new Set();
    const before = pushUndoEntry ? getVisibleRowList() : [];
    visibleRowSet = new Set(Array.from(rows.keys()));
    const after = getVisibleRowList();
    if (pushUndoEntry && !arraysEqual(before, after)){
      pushUndo({type:'setVisibleRows', beforeRows: before, afterRows: after});
    }
    persistRowVisibility();
    rebuildRowsUI();
    render();
  }
  function getCalculatedFinalGradeValue(row){
    if (!row) return null;
    const raw = row[FINAL_GRADE_COLUMN];
    const meta = deriveMarkMeta(raw, FINAL_GRADE_COLUMN);
    return meta ? meta.value : null;
  }
  function collectPerformanceRows(level){
    const results = [];
    ensureCommentThresholds();
    const high = Number(commentConfig.highThreshold) || COMMENT_DEFAULTS.highThreshold;
    const mid = Number(commentConfig.midThreshold) || COMMENT_DEFAULTS.midThreshold;
    rows.forEach((row, idx) => {
      const val = getCalculatedFinalGradeValue(row);
      if (val == null) return;
      if (level === 'high' && val >= high) results.push(idx);
      else if (level === 'mid' && val < high && val >= mid) results.push(idx);
      else if (level === 'low' && val < mid) results.push(idx);
    });
    return results;
  }
  function getPerformanceFilterLabel(level){
    switch(level){
      case 'high': return 'Good';
      case 'mid': return 'Mid';
      case 'low': return 'Needs Support';
      default: return 'All';
    }
  }
  function updateRowFilterStatus(text){
    if (rowFilterStatusEl) rowFilterStatusEl.textContent = text || '';
  }
  function updatePerformanceFilterButtons(){
    const map = [
      [rowFilterGoodBtn, 'high'],
      [rowFilterMidBtn, 'mid'],
      [rowFilterNeedsBtn, 'low']
    ];
    map.forEach(([btn, level]) => {
      if (!btn) return;
      btn.classList.toggle('active', performanceFilterLevel === level);
    });
  }
  function handlePerformanceFilterRequest(level){
    if (!rows.length){
      updateRowFilterStatus('Load data to use performance filters.');
      return;
    }
    if (!allColumns.includes(FINAL_GRADE_COLUMN)){
      updateRowFilterStatus(`"${FINAL_GRADE_COLUMN}" column is required for performance filters.`);
      return;
    }
    const apply = () => applyPerformanceFilter(level);
    if (!columnHasPercentValues(FINAL_GRADE_COLUMN)){
      showMarkWarning(`${FINAL_GRADE_COLUMN} must include % values to filter by performance. Continue anyway?`, apply);
      return;
    }
    apply();
  }
  function applyPerformanceFilter(level){
    const matches = collectPerformanceRows(level);
    performanceFilterLevel = level;
    updatePerformanceFilterButtons();
    if (!matches.length){
      const before = getVisibleRowList();
      visibleRowSet = new Set();
      const after = getVisibleRowList();
      if (!arraysEqual(before, after)){
        pushUndo({type:'setVisibleRows', beforeRows: before, afterRows: after});
      }
      persistRowVisibility();
      rebuildRowsUI();
      render();
      updateRowFilterStatus('No students match this filter.');
      return;
    }
    const before = getVisibleRowList();
    visibleRowSet = new Set(matches);
    const after = getVisibleRowList();
    if (!arraysEqual(before, after)){
      pushUndo({type:'setVisibleRows', beforeRows: before, afterRows: after});
    }
    persistRowVisibility();
    rebuildRowsUI();
    render();
    updateRowFilterStatus(`${matches.length} students match (${getPerformanceFilterLabel(level)}).`);
  }
  function clearPerformanceFilter(){
    performanceFilterLevel = null;
    updatePerformanceFilterButtons();
    updateRowFilterStatus('Showing all students.');
    setAllRowsVisible(true);
  }
  function resetPerformanceFilterState(){
    performanceFilterLevel = null;
    updatePerformanceFilterButtons();
    updateRowFilterStatus('');
  }
  function columnHasData(col){
    if (!col || col === END_OF_LINE_COL || !rows.length) return false;
    for (let i = 0; i < rows.length; i++){
      const value = rows[i]?.[col];
      if (value == null) continue;
      if (typeof value === 'string'){
        if (value.trim() !== '') return true;
      }else{
        return true;
      }
    }
    return false;
  }
  function refreshColumnDataHighlights(targetCol){
    if (!colsDiv) return;
    const rowsToUpdate = targetCol
      ? Array.from(colsDiv.querySelectorAll('.col-row')).filter(row => row.dataset.col === targetCol)
      : Array.from(colsDiv.querySelectorAll('.col-row'));
    rowsToUpdate.forEach(row => {
      const col = row.dataset.col;
      if (!col || col === END_OF_LINE_COL) return;
      const hasData = columnHasData(col);
      row.classList.toggle('has-data', hasData);
    });
  }
  function handleColumnDataMutation(col = null){
    if (dataColumnsOnly){
      rebuildColsUI();
    }else{
      refreshColumnDataHighlights(col);
    }
  }
  function setVisibleRowsFromList(list, allowEmpty = true){
    visibleRowSet = new Set();
    (list || []).forEach(idx => {
      if (Number.isInteger(idx) && idx >= 0 && idx < rows.length){
        visibleRowSet.add(idx);
      }
    });
    if (!visibleRowSet.size && !allowEmpty){
      visibleRowSet = new Set(Array.from(rows.keys()));
    }
    rebuildRowsUI();
  }
  function getVisibleRowList(){
    if (!(visibleRowSet instanceof Set)) return [];
    return Array.from(visibleRowSet).sort((a,b) => a - b);
  }
  function persistRowVisibility(){
    if (!currentFileName || !(visibleRowSet instanceof Set)) return;
    const settings = loadSettings();
    const vr = {...(settings.visibleRows || {})};
    vr[currentFileName] = getVisibleRowList();
    saveSettings({ visibleRows: vr });
  }
  function createContextRecord(name){
    return {
      id: nextFileId++,
      name,
      originalName: name,
      rows,
      originalRows,
      allColumns,
      columnOrder,
      visibleColumns,
      visibleRowSet,
      sortState,
      undoStack,
      redoStack,
      isTransposed,
      searchText: searchEl.value || "",
      firstNameKey,
      lastNameKey,
      studentNameColumn,
      studentNameWarning,
      commentConfig: {
        ...commentConfig,
        extraColumns: [...(commentConfig.extraColumns || [])]
    }
  };
  }
  const REPORT_BUILDER_TEMPLATES = {
    good1: {
      partA: "[Student] has had an excellent first term, demonstrating a strong grasp of the course concepts. [Student]'s overall standing is [AVG%], reflecting solid mastery across units. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "[Student]'s strong work ethic and positive attitude have been evident throughout the term. I am confident [he/she] will continue to excel as we move forward into Term 2."
    },
    good2: {
      partA: "[Student] has demonstrated exceptional understanding of the material this term, holding an overall mark of [AVG%]. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [His/Her] problem-solving assignments showcase careful thinking and clear presentation. [CC_COMMENTARY]",
      partB: "I am pleased with [Student]'s progress and look forward to seeing [him/her] build on this strong foundation in the coming term."
    },
    good3: {
      partA: "[Student] has exceeded expectations this term with outstanding performance across all areas. With an overall [AVG%], [he/she] is performing at a high level across the course. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "I hope to see [Student] continue to be a leader in the classroom as we move forward. [He/She] is on track for an exceptional year ahead."
    },
    average1: {
      partA: "[Student] has built a solid [AVG%] foundation this term, meeting many core expectations while still growing in a few areas. [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "With steady routines and the same positive attitude, I’m confident [he/she] will keep climbing next term."
    },
    average2: {
      partA: "[Student] has made steady gains this term, though results vary across tasks. [SCORE_COMMENTARY_TEST1][RETEST_CLAUSE] With regular review of past work, [he/she] can turn this into consistent success. [CC_COMMENTARY]",
      partB: "[He/She] has shown [he/she] can succeed; next term is about applying those habits every week.",
      retestClause: " However, the retest improvement from [ORIG%] to [RETEST%] on Test 2 indicates initial struggles with preparation.",
      noRetestClause: " [SCORE_COMMENTARY_TEST2]"
    },
    average3: {
      partA: "[Student] has a reasonable start around [AVG%], showing understanding of many concepts and working to connect the tougher ones. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "With a bit more practice and questions early, I expect [he/she] to feel more confident and keep improving."
    },
    newstu1: {
      partA: "[Student] is new to the program and has performed well while adapting to the pace and expectations. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "[Student] shows curiosity and willingness to try new routines; with continued practice and support, [he/she] will keep growing in confidence each week."
    },
    newstu2: {
      partA: "[Student] has recently joined, is learning our structures, and has performed fairly while applying concepts more consistently. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "As [he/she] builds regular study habits and asks questions early, we’ll see steady growth—especially impressive on top of adjusting as a new student."
    },
    newstu3: {
      partA: "[Student] is in the early stages of the program and, while still finding routines for the pace, has shown fair performance given the transition. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "There is clear opportunity to adapt habits and gain confidence—commendable effort so far as a new student. I will continue to guide [him/her] so these changes feel manageable and supportive."
    },
    poor1: {
      partA: "[Student] has faced challenges this term and is still building routines that support learning. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "There’s good opportunity to adjust habits and see progress next term, and I’ll support [him/her] in making those changes."
    },
    poor2: {
      partA: "[Student] has worked hard even as results haven’t matched effort yet. [SCORE_COMMENTARY_TEST1][RETEST_CLAUSE] Building steady routines will help turn that effort into improvement. [CC_COMMENTARY]",
      partB: "Next term is a chance to apply these routines and see growth; with consistent effort and timely support, I believe [he/she] will move forward confidently.",
      retestClause: " While [he/she] improved from [ORIG%] to [RETEST%] on the Test 2 retest, this still reflects difficulty with the material.",
      noRetestClause: " [SCORE_COMMENTARY_TEST2]"
    },
    poor3: {
      partA: "[Student] has had a challenging term, but consistent routines and feedback can open a clear path for growth. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
      partB: "By adopting stronger habits and reaching out early with questions, [Student] can make meaningful progress next term. I’ll continue to guide and support [him/her] through these adjustments."
    }
  };
  const REPORT_BUILDER_TEMPLATE_VARIANTS = {
    good1: {
      partA: [
        "[Student] has had an excellent first term, demonstrating a strong grasp of the course concepts. [Student]'s overall standing is [AVG%], reflecting solid mastery across units. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] is having a very strong term, showing confident understanding across units. [Student] currently holds [AVG%], which reflects reliable mastery and careful work. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "[Student]'s strong work ethic and positive attitude have been evident throughout the term. I am confident [he/she] will continue to excel as we move forward into Term 2.",
        "[Student] combines effort with precision, and I expect continued success as we move into the next term."
      ]
    },
    good2: {
      partA: [
        "[Student] has demonstrated exceptional understanding of the material this term, holding an overall mark of [AVG%]. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [His/Her] problem-solving assignments showcase careful thinking and clear presentation. [CC_COMMENTARY]",
        "[Student] has consistently shown high-level understanding this term, with an overall [AVG%]. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [His/Her] work is organized and thorough. [CC_COMMENTARY]"
      ],
      partB: [
        "I am pleased with [Student]'s progress and look forward to seeing [him/her] build on this strong foundation in the coming term.",
        "[Student] has built a strong foundation and is well positioned for continued success."
      ]
    },
    good3: {
      partA: [
        "[Student] has exceeded expectations this term with outstanding performance across all areas. With an overall [AVG%], [he/she] is performing at a high level across the course. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] has been performing at an exceptional level this term. The overall [AVG%] highlights strong mastery across the curriculum. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "I hope to see [Student] continue to be a leader in the classroom as we move forward. [He/She] is on track for an exceptional year ahead.",
        "[Student] is poised to continue leading by example with consistent effort and high achievement."
      ]
    },
    average1: {
      partA: [
        "[Student] has built a solid [AVG%] foundation this term, meeting many core expectations while still growing in a few areas. [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] holds [AVG%] this term and is meeting many expectations, with a few areas that still need reinforcement. [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "With steady routines and the same positive attitude, I'm confident [he/she] will keep climbing next term.",
        "With consistent routines and continued effort, [he/she] should see steady gains next term."
      ]
    },
    average2: {
      partA: [
        "[Student] has made steady gains this term, though results vary across tasks. [SCORE_COMMENTARY_TEST1][RETEST_CLAUSE] With regular review of past work, [he/she] can turn this into consistent success. [CC_COMMENTARY]",
        "[Student] has shown steady progress, even if outcomes fluctuate across tasks. [SCORE_COMMENTARY_TEST1][RETEST_CLAUSE] Consistent review will help turn this into reliable results. [CC_COMMENTARY]"
      ],
      partB: [
        "[He/She] has shown [he/she] can succeed; next term is about applying those habits every week.",
        "[He/She] has the ability to succeed; next term is about applying those habits consistently."
      ]
    },
    average3: {
      partA: [
        "[Student] has a reasonable start around [AVG%], showing understanding of many concepts and working to connect the tougher ones. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] is working at around [AVG%], showing solid understanding on many outcomes while still building confidence in the more complex ones. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "With a bit more practice and questions early, I expect [he/she] to feel more confident and keep improving.",
        "With more steady practice and early questions, [he/she] should continue to gain confidence."
      ]
    },
    newstu1: {
      partA: [
        "[Student] is new to the program and has performed well while adapting to the pace and expectations. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] is adjusting to the program and has performed well while learning the routines and expectations. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "[Student] shows curiosity and willingness to try new routines; with continued practice and support, [he/she] will keep growing in confidence each week.",
        "With continued practice and support, [Student] will keep building confidence and independence."
      ]
    },
    newstu2: {
      partA: [
        "[Student] has recently joined, is learning our structures, and has performed fairly while applying concepts more consistently. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] is settling into the program and is beginning to apply concepts more consistently. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "As [he/she] builds regular study habits and asks questions early, we'll see steady growth -- especially impressive on top of adjusting as a new student.",
        "As routines become more familiar, I expect steady growth from [Student] each term."
      ]
    },
    newstu3: {
      partA: [
        "[Student] is in the early stages of the program and, while still finding routines for the pace, has shown fair performance given the transition. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] is still adjusting to the pace and routines, but has shown fair performance during the transition. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "There is clear opportunity to adapt habits and gain confidence -- commendable effort so far as a new student. I will continue to guide [him/her] so these changes feel manageable and supportive.",
        "As routines settle, [Student] should gain confidence; I will continue to guide [him/her] through these changes."
      ]
    },
    poor1: {
      partA: [
        "[Student] has faced challenges this term and is still building routines that support learning. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] has encountered difficulty this term and is still developing consistent routines. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "There's good opportunity to adjust habits and see progress next term, and I'll support [him/her] in making those changes.",
        "With structured support and practice, [Student] can make progress next term."
      ]
    },
    poor2: {
      partA: [
        "[Student] has worked hard even as results haven't matched effort yet. [SCORE_COMMENTARY_TEST1][RETEST_CLAUSE] Building steady routines will help turn that effort into improvement. [CC_COMMENTARY]",
        "[Student] has put in effort, though results have been uneven. [SCORE_COMMENTARY_TEST1][RETEST_CLAUSE] Consistent routines will help translate effort into progress. [CC_COMMENTARY]"
      ],
      partB: [
        "Next term is a chance to apply these routines and see growth; with consistent effort and timely support, I believe [he/she] will move forward confidently.",
        "With steady routines and timely support, [Student] can build confidence and make progress next term."
      ]
    },
    poor3: {
      partA: [
        "[Student] has had a challenging term, but consistent routines and feedback can open a clear path for growth. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]",
        "[Student] has faced a difficult term, but steady routines and feedback will support growth over time. [SCORE_COMMENTARY_TEST1] [SCORE_COMMENTARY_TEST2] [CC_COMMENTARY]"
      ],
      partB: [
        "By adopting stronger habits and reaching out early with questions, [Student] can make meaningful progress next term. I'll continue to guide and support [him/her] through these adjustments.",
        "With stronger habits and early questions, [Student] can make meaningful progress; I'll continue to support [him/her]."
      ]
    }
  };
  function getReportBuilderTemplate(coreLevel, context){
    const base = REPORT_BUILDER_TEMPLATES[coreLevel];
    const variants = REPORT_BUILDER_TEMPLATE_VARIANTS[coreLevel];
    if (!base && !variants) return null;
    const seedBase = `${context.studentName || ''}|${coreLevel}|${context.termLabel || ''}`;
    const partA = pickVariant((variants && variants.partA) || (base && base.partA), `${seedBase}|A`);
    const partB = pickVariant((variants && variants.partB) || (base && base.partB), `${seedBase}|B`);
    const retestClause = pickVariant((variants && variants.retestClause) || (base && base.retestClause) || '', `${seedBase}|RC`);
    const noRetestClause = pickVariant((variants && variants.noRetestClause) || (base && base.noRetestClause) || '', `${seedBase}|NRC`);
    return {
      partA,
      partB,
      retestClause,
      noRetestClause
    };
  }
  const REPORT_BUILDER_CC = {
    excellent: "[Student]'s problem-solving tasks are consistently high-quality, showing both mathematical rigor and creative thinking. [He/She] has many strong scores on these tasks, showcasing resilience and insight.",
    good: "[Student]'s problem-solving assignments are well-presented with organized solutions showing good effort and attention to detail.",
    average: "[Student]'s problem-solving work is adequate but would benefit from more detail and careful presentation to fully develop [his/her] understanding.",
    below: "[Student]'s problem-solving assignments are often incomplete and require more effort and detail. [He/She] should start these earlier in the week to allow more time for thoughtful completion.",
    poor: "[Student]'s problem-solving work lacks the detail and quality needed, and the presentation of solutions needs significant improvement. [He/She] needs to make a stronger effort with these assignments."
  };
  const PERFORMANCE_TONE_VARIANTS = {
    good: [
      "With this level of achievement, [Student] is ready for deeper challenges and enrichment.",
      "[Student] is well positioned to extend learning through enrichment and higher-order tasks."
    ],
    average: [
      "With steady routines and consistent review, [he/she] can turn this progress into stronger results.",
      "Regular practice and early questions will help [him/her] move from steady progress to stronger consistency."
    ],
    satisfactory: [
      "With targeted review and steady routines, [Student] can build stronger consistency.",
      "With guided practice and regular check-ins, [Student] can turn partial mastery into steady results."
    ],
    newstu: [
      "As a newer student, continued support and routine will help [him/her] build confidence and consistency.",
      "As [he/she] settles into the program, routines and feedback will help build confidence and momentum."
    ],
    poor: [
      "Targeted practice and clear routines will be important next steps for [Student].",
      "With structured support and step-by-step practice, [Student] can make meaningful gains."
    ]
  };
    const PERFORMANCE_INTRO_BY_GRADE_GROUP = {
    elem: {
      good: [
        "[Student] is doing very well and shows growing confidence.",
        "[Student] is thriving and takes pride in learning.",
        "[Student] shows strong effort and a positive attitude."
      ],
      average: [
        "[Student] is making steady progress and is building confidence.",
        "[Student] is progressing well with growing independence.",
        "[Student] is on track and benefits from steady routines."
      ],
      satisfactory: [
        "[Student] is showing progress but needs steady guidance to build confidence.",
        "[Student] is developing core skills and benefits from regular practice and support.",
        "[Student] is learning steadily, though accuracy and independence are still growing."
      ],
      poor: [
        "[Student] is building core skills and needs steady support.",
        "[Student] is developing foundational skills and needs guidance.",
        "[Student] is gaining confidence with structured support."
      ],
      newstu: [
        "[Student] is adjusting to the program and is building early routines.",
        "[Student] is settling into expectations and learning new routines.",
        "[Student] is new to the program and is beginning to build confidence."
      ]
    },
    middle: {
      good: [
        "[Student] is performing strongly and shows solid reasoning skills.",
        "[Student] is achieving strong results with careful reasoning.",
        "[Student] demonstrates strong understanding and consistent effort."
      ],
      average: [
        "[Student] is meeting expectations and showing steady progress.",
        "[Student] is progressing well and meeting most outcomes.",
        "[Student] is developing steady results with growing confidence."
      ],
      satisfactory: [
        "[Student] is making moderate progress and needs more consistency.",
        "[Student] is meeting some expectations but benefits from guided review.",
        "[Student] is showing partial mastery and needs stronger routines."
      ],
      poor: [
        "[Student] is working to strengthen foundational understanding.",
        "[Student] is building core skills and needs more consistency.",
        "[Student] is improving but still needs guided practice."
      ],
      newstu: [
        "[Student] is adjusting well and beginning to apply concepts.",
        "[Student] is settling in and learning expectations.",
        "[Student] is new to the program and is beginning to build routines."
      ]
    },
    high: {
      good: [
        "[Student] is performing at a high level with consistent reasoning and precision.",
        "[Student] is excelling and demonstrates strong analytical reasoning.",
        "[Student] is achieving top results with disciplined study habits."
      ],
      average: [
        "[Student] is meeting expectations and showing steady growth in reasoning.",
        "[Student] is progressing well and building accuracy in reasoning.",
        "[Student] is on track with steady improvement in performance."
      ],
      satisfactory: [
        "[Student] is achieving at a satisfactory level and needs more consistent accuracy.",
        "[Student] is meeting basic expectations but should refine reasoning and precision.",
        "[Student] is showing developing understanding and would benefit from regular review."
      ],
      poor: [
        "[Student] is developing essential skills and needs more consistent accuracy.",
        "[Student] is rebuilding core skills and needs stronger consistency.",
        "[Student] is working to improve accuracy and reasoning."
      ],
      newstu: [
        "[Student] is adjusting to expectations and building foundational habits.",
        "[Student] is settling into the course and learning expectations.",
        "[Student] is new to the program and is developing steady routines."
      ]
    }
  };
  function getPerformanceIntroLine(coreLevel, context){
    if (!coreLevel) return '';
    let key = '';
    if (coreLevel.startsWith('good')) key = 'good';
    else if (coreLevel.startsWith('average')) key = 'average';
    else if (coreLevel.startsWith('satisfactory')) key = 'satisfactory';
    else if (coreLevel.startsWith('satisfactory')) key = 'satisfactory';
    else if (coreLevel.startsWith('newstu')) key = 'newstu';
    else if (coreLevel.startsWith('poor')) key = 'poor';
    if (!key) return '';
    const groupKey = normalizeGradeGroup(context.gradeGroup) || 'middle';
    const group = PERFORMANCE_INTRO_BY_GRADE_GROUP[groupKey] || PERFORMANCE_INTRO_BY_GRADE_GROUP.middle;
    const variants = group[key] || [];
    const versionIdx = Math.max(0, (parseInt(String(coreLevel).replace(/\D/g, ''), 10) || 1) - 1);
    const line = variants.length ? variants[versionIdx % variants.length] : '';
    return builderReplacePlaceholders(line, context);
  }
  function replaceFirstSentence(text, replacement){
    if (!text || !replacement) return text || '';
    const sentences = splitIntoSentences(text);
    if (!sentences.length) return text;
    sentences[0] = replacement;
    return sentences.join(' ');
  }
  

function getPerformanceToneLine(coreLevel, context){
    if (!coreLevel) return '';
    let key = '';
    if (coreLevel.startsWith('good')) key = 'good';
    else if (coreLevel.startsWith('average')) key = 'average';
    else if (coreLevel.startsWith('satisfactory')) key = 'satisfactory';
    else if (coreLevel.startsWith('newstu')) key = 'newstu';
    else if (coreLevel.startsWith('poor')) key = 'poor';
    if (!key) return '';
    const variants = PERFORMANCE_TONE_VARIANTS[key];
    const line = pickVariant(variants, `${context.studentName || ''}|${coreLevel}|tone`);
    return builderReplacePlaceholders(line, context);
  }
  const REPORT_COMMENT_BANK = [
    {
      id: 'participation',
      title: '🙋 Class Participation',
      options: [
        {id:'cp1', label:'Consistently participates, valuable insights', text:"[Student] consistently raises [his/her] hand and contributes valuable insights to class discussions.", type:'positive'},
        {id:'cp2', label:'Eagerly contributes, actively engages', text:"[Student] eagerly contributes to class discussions and actively engages with the material.", type:'positive'},
        {id:'cp3', label:'Active participant, positive environment', text:"[Student] is an active participant who helps create a positive learning environment for peers.", type:'positive'},
        {id:'cp4', label:'Needs to participate more, be more outgoing', text:"[Student] is encouraged to be more outgoing in class and volunteer answers more frequently.", type:'constructive'},
        {id:'cp5', label:'Adequate participation, could improve', text:"[Student] participates adequately but would benefit from raising [his/her] hand more often.", type:'constructive'},
        {id:'cp6', label:'Rarely participates, needs engagement', text:"[Student] rarely volunteers answers or participates in class discussions and should make a stronger effort to engage.", type:'constructive'},
        {id:'cp7', label:'Asks insightful questions, curious', text:"[Student] often asks insightful questions about alternative solutions and demonstrates genuine curiosity about the material.", type:'positive'},
        {id:'cp8', label:'Good group work and collaboration', text:"[Student] works well in group discussions and contributes constructively to collaborative work.", type:'positive'},
        {id:'cp9', label:'Makes most of class time, enthusiastic', text:"[Student] makes the most of class time, engaging enthusiastically and dedicating [himself/herself] to discussions.", type:'positive'},
        {id:'cp10', label:'Be more present, ask clarifying questions', text:"[Student] is encouraged to be more present during class by asking questions to clarify understanding.", type:'constructive'}
      ]
    },
    {
      id: 'peer',
      title: '🤝 Peer Teaching & Helping',
      options: [
        {id:'peer1', label:'Serves as teacher among peers', text:"[Student] serves as a teacher among [his/her] peers and is always willing to help classmates succeed.", type:'positive'},
        {id:'peer2', label:'Willing to help peers succeed', text:"[Student] is always willing to help [his/her] peers succeed and contributes to a supportive learning environment.", type:'positive'},
        {id:'peer3', label:'Helps classmates, works constructively', text:"[Student] helps out [his/her] classmates and works constructively with peers.", type:'positive'},
        {id:'peer4', label:'Cheerful demeanor enhances classroom morale', text:"[His/Her] cheerful demeanor has enhanced the morale of our classroom.", type:'positive'},
        {id:'peer5', label:'Presence helps all students', text:"[His/Her] presence in the class helps all students and contributes to a positive atmosphere.", type:'positive'}
      ]
    },
    {
      id: 'homeworkQuality',
      title: '📚 Homework Quality',
      options: [
        {id:'hw1', label:'Evident effort, quality and detail', text:"[Student] consistently puts evident effort into [his/her] assignments, completing them with quality and attention to detail.", type:'positive'},
        {id:'hw2', label:'Assignments reflect hard work and detail', text:"[His/Her] assignments consistently reflect [his/her] hard work and attention to detail.", type:'positive'},
        {id:'hw3', label:'Care and detail evident in high-quality work', text:"[Student]'s care and attention to detail are evident in [his/her] high-quality assignments.", type:'positive'},
        {id:'hw4', label:'Diligence shines through in quality', text:"[Student]'s diligence shines through in the high quality of [his/her] assignments.", type:'positive'},
        {id:'hw5', label:'High quality, detailed problem-solving', text:"[Student]'s problem-solving assignments consistently demonstrate high quality, detailed solutions.", type:'positive'},
        {id:'hw6', label:'Needs better presentation/neatness', text:"[Student] needs to improve the presentation and neatness of [his/her] assignments.", type:'constructive'},
        {id:'hw7', label:'Be more detailed for bonus marks', text:"[Student] is encouraged to be more detailed with [his/her] solutions on problem-solving assignments to pick up bonus marks.", type:'constructive'}
      ]
    },
    {
      id: 'homeworkCompletion',
      title: '⏰ Homework Completion',
      options: [
        {id:'hwc1', label:'Start homework earlier in week', text:"[Student] should start homework earlier in the week to allow more time for thoughtful completion.", type:'constructive'},
        {id:'hwc2', label:'Excellent time management, starts early', text:"[Student] demonstrates excellent time management by starting [his/her] homework early in the week.", type:'positive'},
        {id:'hwc3', label:'Submits assignments on time', text:"[Student] demonstrates strong time management by consistently submitting assignments on time.", type:'positive'},
        {id:'hwc4', label:'Needs to complete homework fully, not rush', text:"[Student] should focus on completing homework assignments completely rather than rushing through them.", type:'constructive'},
        {id:'hwc5', label:'Complete fully AND start earlier', text:"[Student] needs to focus on completing homework assignments completely and starting them earlier in the week.", type:'constructive'},
        {id:'hwc6', label:'Prioritize homework for full benefit of class', text:"[Student] is encouraged to prioritize homework completion so that [he/she] can receive the full benefit of in-class instruction.", type:'constructive'},
        {id:'hwc7', label:'Submit homework on time', text:"[Student] should focus on submitting homework on time to stay on track with the curriculum.", type:'constructive'}
      ]
    },
    {
      id: 'studyHabits',
      title: '📖 Study Habits',
      options: [
        {id:'sh1', label:'Redo past assignments before tests', text:"[Student] should focus on test preparation by redoing previous assignments and reviewing errors carefully.", type:'constructive'},
        {id:'sh2', label:'Strong test preparation', text:"[Student] demonstrates strong test preparation skills and consistently performs well on assessments.", type:'positive'},
        {id:'sh3', label:'Lacks adequate test preparation', text:"[Student]'s test scores indicate a lack of adequate preparation and [he/she] needs to dedicate more study time.", type:'constructive'},
        {id:'sh4', label:'Make corrections, learn from mistakes', text:"[Student] should focus on making corrections to tests and learning from [his/her] mistakes.", type:'constructive'},
        {id:'sh5', label:'Review to recognize patterns', text:"[Student] is encouraged to review previous assignments to recognize recurring problem types and patterns.", type:'constructive'},
        {id:'sh6', label:'Avoid careless mistakes', text:"[Student] should be more attentive to avoid careless mistakes on tests and assignments.", type:'constructive'},
        {id:'sh7', label:'Impressive retest improvement (positive)', text:"[Student] showed impressive improvement on [his/her] retest, demonstrating strong ability to learn from mistakes and apply corrections.", type:'positive'},
        {id:'sh8', label:'Set consistent SOM study time', text:"[Student] is encouraged to prioritize a consistent time for Spirit of Math work to engage more deeply with the material.", type:'constructive'}
      ]
    },
    {
      id: 'assignmentTestGap',
      title: '📊 Assignment vs Test Gap',
      options: [
        {id:'gap1', label:'Does well on HW, struggles on tests', text:"Patterns indicate that while [Student] does very well on assignments, [he/she] struggles to translate this knowledge on the corresponding tests.", type:'constructive'},
        {id:'gap2', label:'Good on assignments, review needed for tests', text:"[Student] continues to do very well on assignments but sometimes struggles with tests, suggesting [he/she] needs to focus more on review before assessments.", type:'constructive'},
        {id:'gap3', label:'High quality work, but more test review needed', text:"[His/Her] work continues to be high quality, but test performance suggests more thorough review is needed when preparing for assessments.", type:'constructive'},
        {id:'gap4', label:'Strong in HW, needs better test translation', text:"[Student] demonstrates strong understanding in assignments but needs to translate this more effectively to test situations.", type:'constructive'}
      ]
    },
    {
      id: 'brightspace',
      title: '💻 Brightspace & Resources',
      options: [
        {id:'bs1', label:'Watch Brightspace videos after lessons', text:"[Student] is encouraged to watch the Brightspace videos after each lesson to reinforce [his/her] understanding of new concepts.", type:'constructive'},
        {id:'bs2', label:'Watch videos before starting homework', text:"[Student] should develop the habit of watching Brightspace videos to review lesson content before starting homework.", type:'constructive'},
        {id:'bs3', label:'Use videos to start homework', text:"[Student] should use Brightspace videos as a catalyst to begin [his/her] homework questions.", type:'constructive'},
        {id:'bs4', label:'Reattempt class examples independently', text:"[Student] is encouraged to reattempt examples taken up in class by [himself/herself] to reinforce understanding of new concepts.", type:'constructive'},
        {id:'bs5', label:'Utilize Brightspace resources', text:"[Student] should utilize the resources available on Brightspace to supplement [his/her] learning.", type:'constructive'},
        {id:'bs6', label:'Reinforce concepts with videos (formal)', text:"Moving forward, [Student] should reinforce [his/her] understanding of core concepts by watching the Brightspace videos after each lesson.", type:'constructive'},
        {id:'bs7', label:'Videos help with HW completion', text:"These videos will help [him/her] review the information necessary to complete homework successfully.", type:'constructive'}
      ]
    },
    {
      id: 'helpSeeking',
      title: '🆘 Seeking Help',
      options: [
        {id:'help1', label:'Use Microsoft Teams for help', text:"[Student] is encouraged to reach out for help on Microsoft Teams when needed.", type:'constructive'},
        {id:'help2', label:'Takes initiative to ask questions', text:"[Student] takes initiative to ask questions when [he/she] doesn't understand something.", type:'positive'},
        {id:'help3', label:'Needs to seek help more proactively', text:"[Student] needs to be more proactive about seeking help and should not hesitate to ask questions.", type:'constructive'},
        {id:'help4', label:'Use class time to ask questions', text:"[Student] should actively use class time to ask questions when needed, so that any difficulties can be addressed promptly.", type:'constructive'},
        {id:'help5', label:'Rarely seeks help, should ask more', text:"[Student] rarely seeks help and would benefit greatly from asking questions more frequently.", type:'constructive'},
        {id:'help6', label:'Reach out for extra help', text:"[Student] is encouraged to reach out for extra help as needed and continue practicing diligently.", type:'constructive'},
        {id:'help7', label:'Ready to learn, asks clarifying questions', text:"[Student] consistently comes to class ready to learn, asking insightful questions to clarify any uncertainties.", type:'positive'},
        {id:'help8', label:'Seek help BEFORE class', text:"[Student] is encouraged to seek help from [his/her] teacher prior to class to address difficulties early.", type:'constructive'}
      ]
    },
    {
      id: 'attendance',
      title: '📅 Attendance & Behavior',
      options: [
        {id:'att1', label:'Too many absences/make-ups', text:"[Student] would benefit from attending class more regularly and reducing reliance on make-up sessions.", type:'constructive'},
        {id:'att2', label:'Excellent attendance', text:"[Student] has excellent attendance and is consistently engaged during class time.", type:'positive'},
        {id:'att3', label:'Positive attitude, enthusiastic', text:"[Student] demonstrates a positive attitude and enthusiasm towards learning.", type:'positive'},
        {id:'att4', label:'Listens well, follows directions', text:"[Student] listens well to instructions and follows directions carefully.", type:'positive'},
        {id:'att5', label:'Pleasure to have in class', text:"[Student] is a pleasure to have in class and contributes to a positive classroom atmosphere.", type:'positive'},
        {id:'att6', label:'Pleasure to have in class (formal)', text:"It has been a pleasure having [Student] in the class this term.", type:'positive'},
        {id:'att7', label:'Ready-to-learn attitude', text:"[Student] comes to class each week with a ready-to-learn attitude.", type:'positive'},
        {id:'att8', label:'Regular attendance needed for curriculum', text:"Attending class regularly and reducing reliance on make-ups would help [Student] stay on track with the curriculum.", type:'constructive'},
        {id:'att9', label:'Attendance affecting assessment', text:"[Student]'s attendance in class has made it difficult to accurately assess [his/her] participation and progress.", type:'constructive'}
      ]
    },
    {
      id: 'personalQualities',
      title: '✨ Personal Qualities',
      options: [
        {id:'pq1', label:'Warm and friendly', text:"[Student] is a warm and friendly student with a positive demeanor.", type:'positive'},
        {id:'pq2', label:'Kind-hearted, offers valuable insights', text:"[Student] is a kind-hearted student who consistently offers valuable insights during class discussions.", type:'positive'},
        {id:'pq3', label:'Honest and hardworking', text:"[Student] is an honest and hardworking student whose diligence has yielded remarkable results this term.", type:'positive'},
        {id:'pq4', label:'Demonstrates humility and willingness to learn', text:"[Student] demonstrates humility, effort, and willingness to learn.", type:'positive'},
        {id:'pq5', label:'Pleasant, remains positive through challenges', text:"[Student] is a pleasant student who remains positive while working through challenging problems.", type:'positive'},
        {id:'pq6', label:'Smart and capable, great potential', text:"[Student] is a smart and capable student with great potential.", type:'positive'},
        {id:'pq7', label:'Sincere, committed to overcome challenges', text:"[Student] demonstrates sincerity and a commitment to overcome challenges.", type:'positive'},
        {id:'pq8', label:'Dedicated, diligent in all work', text:"[Student] is a dedicated student whose diligence is evident in all aspects of [his/her] work.", type:'positive'},
        {id:'pq9', label:'Fortitude and commitment to excellence', text:"[Student] demonstrates fortitude and commitment to excellence that are admirable.", type:'positive'},
        {id:'pq10', label:'Takes work seriously, strives for excellence', text:"[Student] takes [his/her] work seriously and strives for excellence in all that [he/she] does.", type:'positive'},
        {id:'pq11', label:'Cheerful and kind', text:"[Student] is cheerful, kind, and contributes to a positive classroom atmosphere.", type:'positive'},
        {id:'pq12', label:'Polite and well-behaved', text:"[Student] is polite, well-behaved, and shows respect to teachers and peers alike.", type:'positive'},
        {id:'pq13', label:'Shows effort despite struggles', text:"Even when struggling with some homework submissions, [Student] remains present, attentive, and gives the impression of doing [his/her] best.", type:'positive'},
        {id:'pq14', label:'Reliable classroom presence', text:"[Student] consistently shows up ready to participate, bringing steady effort and a positive attitude to class.", type:'positive'}
      ]
    },
    {
      id: 'understanding',
      title: '🎓 Understanding & Mastery',
      options: [
        {id:'und1', label:'Demonstrates mastery', text:"[Student] demonstrates mastery of core concepts and applies them skillfully.", type:'positive'},
        {id:'und2', label:'Natural aptitude, impressive problem-solving', text:"[Student] demonstrates a natural aptitude for mathematics and impressive problem-solving skills.", type:'positive'},
        {id:'und3', label:'Propensity and love for mathematics', text:"[Student] demonstrates a propensity and love for mathematics through active engagement with the material.", type:'positive'},
        {id:'und4', label:'Evident skill, asset to class', text:"[His/Her] evident skill in mathematics has made [him/her] a highly engaged student and an asset to the class.", type:'positive'},
        {id:'und5', label:'Advanced understanding', text:"[Student] demonstrates an advanced understanding of the material.", type:'positive'},
        {id:'und6', label:'Understands basics, struggles with difficult', text:"[Student] shows understanding of many concepts but sometimes struggles with more challenging applications.", type:'constructive'},
        {id:'und7', label:'Inconsistent understanding', text:"[Student] demonstrates inconsistent understanding and would benefit from more thorough review of fundamental concepts.", type:'constructive'},
        {id:'und8', label:'Capable, needs more confidence', text:"[Student] is capable of excellent work but needs to develop more confidence in [his/her] abilities.", type:'constructive'},
        {id:'und9', label:'Expertly applies rules, sophisticated techniques', text:"[Student] expertly applies mathematical rules and demonstrates sophisticated problem-solving techniques.", type:'positive'},
        {id:'und10', label:'Creative solutions to higher-order problems', text:"[Student] consistently impresses with [his/her] ability to think of creative solutions to higher-order problems.", type:'positive'}
      ]
    },
    {
      id: 'newStudent',
      title: '🆕 New Student Context',
      options: [
        {id:'new1', label:'Great commitment as new student', text:"[Student] has shown great commitment, perseverance, and dedication as a new student in the rigorous SOM program.", type:'positive'},
        {id:'new2', label:'Impressive adaptability, first year', text:"[Student] has demonstrated impressive adaptability in [his/her] first year in the program.", type:'positive'},
        {id:'new3', label:'Good progress for first-time student', text:"[Student] is showing a lot of progress despite being a first-time SOM student.", type:'positive'},
        {id:'new4', label:'Enthusiastically attends as new student', text:"[Student] continues to enthusiastically attend the Spirit of Math program as a new student.", type:'positive'},
        {id:'new5', label:'Fast-paced curriculum needs HW completion', text:"As our curriculum is sophisticated and fast-paced, [Student] is encouraged to stay on top of homework to receive the full benefit of instruction.", type:'constructive'},
        {id:'new6', label:'Adapting to pace, willing to learn', text:"[Student] is still adapting to the pace of the class but shows genuine willingness to keep learning and applying each concept.", type:'constructive'},
        {id:'new7', label:'Positive attitude despite lower marks', text:"Even though current marks are below target, [Student] stays positive, asks questions, and is beginning to adjust to the program’s expectations.", type:'constructive'},
        {id:'new8', label:'Progressing steadily with support', text:"As a new enrollee, [Student] is progressing steadily; continued practice and feedback will help [him/her] apply concepts more confidently.", type:'constructive'},
        {id:'new9', label:'Building stamina for fast lessons', text:"[Student] is building stamina for our faster lessons; with ongoing practice, [he/she] will translate this effort into stronger results.", type:'constructive'}
      ]
    },
    {
      id: 'future',
      title: '🔮 Looking Ahead',
      options: [
        {id:'future1', label:'On track for successful year', text:"[Student] is well-positioned for a successful year ahead.", type:'positive'},
        {id:'future2', label:'Hope to see continued success', text:"I hope to see [Student] continue [his/her] strong performance in Term 2.", type:'positive'},
        {id:'future3', label:'Confident of better results with effort', text:"With more consistent effort, I am confident [Student] is capable of achieving stronger results in the coming term.", type:'constructive'},
        {id:'future4', label:'Hope to see leadership', text:"I hope to see [Student] being a leader in the class for the rest of the year.", type:'positive'},
        {id:'future5', label:'Next term = opportunity', text:"The next term offers an opportunity to build on [his/her] strengths and address areas for improvement.", type:'constructive'},
        {id:'future6', label:'Confident feedback will bring success', text:"By implementing this feedback, I am confident [Student] will see greater success in Term 2.", type:'constructive'},
        {id:'future7', label:'Certain of success with feedback', text:"As [Student] implements this feedback, I am certain [he/she] will have a successful term 2.", type:'constructive'},
        {id:'future8', label:'Look forward to working next term', text:"I look forward to continuing to work with [Student] next term.", type:'positive'},
        {id:'future9', label:'Will excel with ongoing effort', text:"With ongoing effort, [he/she] will certainly continue to excel.", type:'positive'},
        {id:'future10', label:'Successful year with adjustments', text:"[Student] is on [his/her] way to a successful year if [he/she] makes a few adjustments.", type:'constructive'}
      ]
    },
    {
      id: 'familyCommunication',
      title: '📞 Family Communication',
      options: [
        {id:'fam1', label:'Kept parents informed about progress', text:"I have been in regular contact with [Student]'s family to keep them informed about progress and next steps.", type:'positive'},
        {id:'fam2', label:'Discussed missing assignments with parent', text:"I have spoken with [Student]'s parent about missing assignments and shared a plan to get caught up.", type:'positive'},
        {id:'fam3', label:'Answered homework questions via email', text:"I have responded to family emails with clarifications on homework questions to support [Student]'s learning at home.", type:'positive'},
        {id:'fam4', label:'Collaborating with family on support plan', text:"I am collaborating with [Student]'s family to provide consistent support between class and home.", type:'positive'},
        {id:'fam5', label:'Parent updated on progress checkpoints', text:"[Student]'s parent has been updated on progress checkpoints and understands the upcoming expectations.", type:'positive'},
        {id:'fam6', label:'Strong rapport with parents', text:"There is a strong rapport with [Student]'s parents; our ongoing communication helps keep goals aligned.", type:'positive'},
        {id:'fam7', label:'Parents responsive and supportive', text:"[Student]'s parents have been responsive and supportive in our conversations about progress and next steps.", type:'positive'}
      ]
    }
  ];
  function registerContext(ctx){
    fileContexts.push(ctx);
    activeContext = ctx;
    currentFileName = ctx.name;
    updateFilesUI();
    highlightActiveFile();
    syncActiveContext();
    renderWarning('');
  }
  function setActiveContext(id){
    if (activeContext && activeContext.id === id){
      renderWarning(activeContext.studentNameWarning || '');
      return;
    }
    if (activeContext){
      activeContext.searchText = searchEl.value || "";
      activeContext.isTransposed = isTransposed;
    }
    const ctx = fileContexts.find(c => c.id === id);
    if (!ctx) return;
    rows = ctx.rows;
    originalRows = ctx.originalRows;
    allColumns = ctx.allColumns;
    columnOrder = ctx.columnOrder;
    visibleColumns = ctx.visibleColumns;
    visibleRowSet = ctx.visibleRowSet;
    sortState = ctx.sortState;
    undoStack = ctx.undoStack;
    redoStack = ctx.redoStack;
    isTransposed = ctx.isTransposed;
    firstNameKey = ctx.firstNameKey || null;
    lastNameKey = ctx.lastNameKey || null;
    studentNameColumn = ctx.studentNameColumn || null;
    studentNameWarning = ctx.studentNameWarning || '';
    commentConfig = {...COMMENT_DEFAULTS, ...(ctx.commentConfig || {})};
    commentConfig.extraColumns = Array.isArray(commentConfig.extraColumns) ? [...commentConfig.extraColumns] : [];
    currentFileName = ctx.name;
    searchEl.value = ctx.searchText || "";
    activeContext = ctx;
    setDataActionButtonsDisabled(false);
    editHeadersBtn.disabled = false;
    transposeBtn.disabled = false;
    updateTransposeButton();
    enforceEndOfLineColumn();
    ensureDefaultCommentGradeColumn();
    rebuildColsUI();
    rebuildRowsUI();
    render();
    resetPerformanceFilterState();
    highlightActiveFile();
    renderWarning(ctx.studentNameWarning || '');
    status('READY');
    // Align Print View selection with the clicked file in single-class mode.
    if (!printAllClasses){
      selectedClassIds = new Set([ctx.id]);
    }
    if (isPrintTabActive()){
      renderPrintClassList();
      renderPrintPreview();
      renderPrintMarkingColumnPanel();
    }
    if (isCommentsTabActive()){
      openCommentsModalInternal();
    }
  }
  function syncActiveContext(){
    if (!activeContext || activeContext.name !== currentFileName) return;
    activeContext.rows = rows;
    activeContext.originalRows = originalRows;
    activeContext.allColumns = allColumns;
    activeContext.columnOrder = columnOrder;
    activeContext.visibleColumns = visibleColumns;
    activeContext.visibleRowSet = visibleRowSet;
    activeContext.sortState = sortState;
    activeContext.undoStack = undoStack;
    activeContext.redoStack = redoStack;
    activeContext.isTransposed = isTransposed;
    activeContext.searchText = searchEl.value || "";
    activeContext.firstNameKey = firstNameKey;
    activeContext.lastNameKey = lastNameKey;
    activeContext.studentNameColumn = studentNameColumn;
    activeContext.studentNameWarning = studentNameWarning;
    activeContext.commentConfig = {
      ...commentConfig,
      extraColumns: [...(commentConfig.extraColumns || [])]
    };
  }
  function highlightActiveFile(){
    Array.from(filesList.querySelectorAll('.file-item')).forEach(el => {
      el.classList.toggle('active', activeContext && Number(el.dataset.id) === activeContext.id);
    });
    if (filesDrawerMiniList){
      Array.from(filesDrawerMiniList.querySelectorAll('.files-mini-item')).forEach(el => {
        el.classList.toggle('active', activeContext && Number(el.dataset.id) === activeContext.id);
      });
    }
    renderWarning(activeContext?.studentNameWarning || '');
  }
  function getFileNameWithoutExtension(name){
    const raw = String(name || '').trim();
    if (!raw) return 'file';
    return raw.replace(/\.[^./\\]+$/, '') || raw;
  }
  function updateFilesUI(){
    if (fileCountEl) fileCountEl.textContent = String(fileContexts.length);
    filesList.innerHTML = "";
    if (filesDrawerMiniList) filesDrawerMiniList.innerHTML = "";
    fileContexts.forEach(ctx => {
      const item = document.createElement('div');
      item.className = 'file-item' + (activeContext && ctx.id === activeContext.id ? ' active' : '');
      item.dataset.id = ctx.id;
      const row = document.createElement('div');
      row.className = 'file-row';
      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.value = ctx.name;
      nameInput.style.flex = '1';
      nameInput.style.border = '1px solid #e5e5e5';
      nameInput.style.borderRadius = '8px';
      nameInput.style.padding = '4px 6px';
      nameInput.style.fontWeight = '600';
      nameInput.addEventListener('click', (e) => e.stopPropagation());
      nameInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter'){
          e.preventDefault();
          nameInput.blur();
        }
        if (e.key === 'Escape'){
          e.preventDefault();
          nameInput.value = ctx.name;
          nameInput.blur();
        }
      });
      nameInput.addEventListener('blur', () => {
        const next = nameInput.value.trim();
        if (next && next !== ctx.name){
          renameContext(ctx.id, next);
        }else{
          nameInput.value = ctx.name;
        }
      });
      const metaSpan = document.createElement('small');
      metaSpan.className = 'muted';
      metaSpan.textContent = `${ctx.rows.length}r x ${ctx.allColumns.length}c`;
      row.appendChild(nameInput);
      row.appendChild(metaSpan);
      item.appendChild(row);
      const actions = document.createElement('div');
      actions.className = 'file-actions';
      const renameBtn = document.createElement('button');
      renameBtn.className = 'btn';
      renameBtn.textContent = 'Rename';
      renameBtn.title = 'Edit file name';
      renameBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        nameInput.focus();
        nameInput.select();
      });
      const removeBtn = document.createElement('button');
      removeBtn.className = 'btn warn';
      removeBtn.textContent = 'X';
      removeBtn.title = 'Remove file from session';
      removeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        removeContext(ctx.id);
      });
      actions.appendChild(renameBtn);
      actions.appendChild(removeBtn);
      item.appendChild(actions);
      filesList.appendChild(item);

      if (filesDrawerMiniList){
        const mini = document.createElement('button');
        mini.type = 'button';
        mini.className = 'files-mini-item' + (activeContext && ctx.id === activeContext.id ? ' active' : '');
        mini.dataset.id = ctx.id;
        mini.title = ctx.name;
        mini.textContent = getFileNameWithoutExtension(ctx.name);
        mini.setAttribute('aria-label', `Switch to ${ctx.name}`);
        mini.addEventListener('click', () => setActiveContext(ctx.id));
        filesDrawerMiniList.appendChild(mini);
      }
    });
    highlightActiveFile();
    syncPrintClassSelection();
    renderPrintClassList();
    if (isPrintTabActive()){
      renderPrintPreview();
    }
  }
  function syncVisibleOrder(){
    const set = new Set(visibleColumns);
    visibleColumns = columnOrder.filter(c => set.has(c));
  }
  function applyStoredColumnOrder(order){
    if (!Array.isArray(order)) return;
    const seen = new Set();
    const next = [];
    order.forEach(c => {
      if (allColumns.includes(c) && !seen.has(c)){
        seen.add(c);
        next.push(c);
      }
    });
    allColumns.forEach(c => {
      if (!seen.has(c)){
        seen.add(c);
        next.push(c);
      }
    });
    columnOrder = next;
  }
  function reorderColumns(source, target){
    const before = columnOrder.slice();
    const srcIdx = columnOrder.indexOf(source);
    const tgtIdx = columnOrder.indexOf(target);
    if (srcIdx === -1 || tgtIdx === -1) return;
    if (srcIdx < tgtIdx && srcIdx === tgtIdx - 1) return;
    const [moved] = columnOrder.splice(srcIdx, 1);
    const insertIdx = srcIdx < tgtIdx ? tgtIdx - 1 : tgtIdx;
    columnOrder.splice(insertIdx, 0, moved);
    syncVisibleOrder();
    pushUndo({type:'reorderColumns', beforeOrder: before, afterOrder: columnOrder.slice()});
    saveSettings({ columnOrder });
    rebuildColsUI();
    rebuildRowsUI();
    render();
  }
  function moveColumnToEnd(col){
    const before = columnOrder.slice();
    const idx = columnOrder.indexOf(col);
    if (idx === -1 || idx === columnOrder.length - 1) return;
    const [moved] = columnOrder.splice(idx, 1);
    columnOrder.push(moved);
    syncVisibleOrder();
    pushUndo({type:'reorderColumns', beforeOrder: before, afterOrder: columnOrder.slice()});
    saveSettings({ columnOrder });
    rebuildColsUI();
    render();
    renderWarning(studentNameWarning || '');
  }

  function render(){
    enforceEndOfLineColumn();
    if (!isPrintTabActive()){
      togglePrintCommentsDrawer(false, true);
    }
    filterRows(searchEl.value);
    applySort();
    thead.innerHTML = "";
    if (isTransposed){
      renderTransposedHeader();
      renderBodyTransposed();
    }else{
      renderNormalHeader();
      renderBody();
    }
    syncActiveContext();
    updateMeta();
    updateFilesUI();
    adjustTableSize();
    if (isCommentsTabActive()){
      buildCommentExtrasList();
      refreshCommentsPreview();
    }
  }

  function renderNormalHeader(){
    const tr = document.createElement('tr');
    visibleColumns.forEach(c => {
      const th = document.createElement('th');
      th.textContent = c;
      th.addEventListener('click', () => cycleSort(c, th));
      if (studentNameColumn && c === studentNameColumn) th.classList.add('sticky-student');
      if (sortState.key === c) th.classList.add(sortState.dir === 'asc' ? 'sort-asc' : sortState.dir === 'desc' ? 'sort-desc' : '');
      tr.appendChild(th);
    });
    thead.appendChild(tr);
  }

  function renderBody(){
    tbody.innerHTML = "";
    filteredIdx.forEach((i, rowPos) => {
      const r = rows[i];
      const tr = document.createElement('tr');
      tr.dataset.rowIndex = i;
      visibleColumns.forEach((c, colPos) => {
        const td = document.createElement('td');
        const isStudent = studentNameColumn && c === studentNameColumn;
        const isEndIndicator = (c === END_OF_LINE_COL);
        if (isStudent) td.classList.add('sticky-student');
        td.dataset.col = c;
        td.contentEditable = (isStudent || isEndIndicator) ? "false" : "true";
        if (isStudent || isEndIndicator) td.classList.add('readonly');
        td.dataset.rpos = String(rowPos);
        td.dataset.cpos = String(colPos);
        if (!isEndIndicator){
          td.addEventListener('focus', () => { td.dataset.old = td.textContent; });
          td.addEventListener('blur', () => {
            const oldVal = td.dataset.old ?? "";
            const newVal = td.textContent;
            if (oldVal !== newVal) {
              const normalized = normalizeMarkInput(newVal, c);
              pushUndo({type:'editCell', rowIndex:i, col:c, before:oldVal, after:newVal});
              rows[i][c] = normalized;
              if (!isStudent && isNameColumn(c)) refreshStudentNameForRow(i);
              if (!isStudent){
                const meta = applyMarkStyling(td, normalized, c);
                td.textContent = meta ? formatMarkText(meta, normalized) : normalized ?? '';
              }
              handleColumnDataMutation(c);
            }
          });
          td.addEventListener('keydown', handleCellNav);
        }
        if (isEndIndicator){
          td.textContent = END_OF_LINE_VALUE;
        }else{
          const meta = (!isStudent) ? applyMarkStyling(td, r[c], c) : null;
          td.textContent = meta ? formatMarkText(meta, r[c]) : (r[c] ?? "");
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    updateCounts();
    autoFit();
  }

  function renderTransposedHeader(){
    const tr = document.createElement('tr');
    const corner = document.createElement('th');
    corner.textContent = 'Field';
    corner.className = 'corner-header no-sort';
    tr.appendChild(corner);
    if (!filteredIdx.length){
      const th = document.createElement('th');
      th.textContent = 'No rows match filter';
      th.classList.add('no-sort', 'muted');
      tr.appendChild(th);
    }else{
      for (const idx of filteredIdx){
        const th = document.createElement('th');
        th.textContent = getRowLabel(idx);
        th.classList.add('no-sort');
        tr.appendChild(th);
      }
    }
    thead.appendChild(tr);
  }

  function renderBodyTransposed(){
    tbody.innerHTML = "";
    visibleColumns.forEach((col, rowPos) => {
      const tr = document.createElement('tr');
      tr.dataset.columnKey = col;
      const rowHeader = document.createElement('th');
      rowHeader.textContent = col;
      rowHeader.className = 'row-header no-sort';
      tr.appendChild(rowHeader);
      filteredIdx.forEach((idx, colPos) => {
        const td = document.createElement('td');
        const isStudent = studentNameColumn && col === studentNameColumn;
        const isEndIndicator = (col === END_OF_LINE_COL);
        td.dataset.col = col;
        td.dataset.rowIndex = String(idx);
        td.contentEditable = (isStudent || isEndIndicator) ? "false" : "true";
        if (isStudent || isEndIndicator) td.classList.add('readonly');
        td.dataset.rpos = String(rowPos);
        td.dataset.cpos = String(colPos);
        if (!isEndIndicator){
          td.addEventListener('focus', () => { td.dataset.old = td.textContent; });
          td.addEventListener('blur', () => {
            const oldVal = td.dataset.old ?? "";
            const newVal = td.textContent;
            if (oldVal !== newVal) {
              const normalized = normalizeMarkInput(newVal, col);
              pushUndo({type:'editCell', rowIndex:idx, col, before:oldVal, after:newVal});
              rows[idx] = rows[idx] || {};
              rows[idx][col] = normalized;
              if (!isStudent && isNameColumn(col)) refreshStudentNameForRow(idx);
              if (!isStudent){
                const meta = applyMarkStyling(td, normalized, col);
                td.textContent = meta ? formatMarkText(meta, normalized) : normalized ?? '';
              }
              handleColumnDataMutation(col);
            }
          });
          td.addEventListener('keydown', handleCellNav);
        }
        if (isEndIndicator){
          td.textContent = END_OF_LINE_VALUE;
        }else{
          const meta = (!isStudent) ? applyMarkStyling(td, rows[idx]?.[col], col) : null;
          td.textContent = meta ? formatMarkText(meta, rows[idx]?.[col]) : (rows[idx]?.[col] ?? "");
        }
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    updateCounts();
  }

  function updateCounts(){
    if (isTransposed){
      countsEl.textContent = `${visibleColumns.length} columns x ${filteredIdx.length} rows`;
    }else{
      countsEl.textContent = `${filteredIdx.length} / ${rows.length} rows`;
    }
  }

  function getRowLabel(idx){
    const row = rows[idx] || {};
    if (studentNameColumn){
      const name = String(row[studentNameColumn] ?? '').trim();
      if (name) return name;
    }
    const built = buildStudentName(row);
    if (built) return built;
    const fallbackKey = columnOrder.find(col => {
      if (!col || col === END_OF_LINE_COL) return false;
      const norm = String(col).toLowerCase();
      if (norm.includes('org')) return false;
      return true;
    }) || columnOrder[0];
    const raw = fallbackKey ? row[fallbackKey] : null;
    const text = raw != null && String(raw).trim() ? String(raw) : `Row ${idx + 1}`;
    return text;
  }
  function handleCellNav(e){
    if (e.key !== 'Enter' && e.key !== 'Tab') return;
    const td = e.currentTarget;
    const rowPos = Number(td.dataset.rpos);
    const colPos = Number(td.dataset.cpos);
    if (Number.isNaN(rowPos) || Number.isNaN(colPos)) return;
    const maxRows = isTransposed ? visibleColumns.length : filteredIdx.length;
    const maxCols = isTransposed ? filteredIdx.length : visibleColumns.length;
    let targetRow = rowPos;
    let targetCol = colPos;
    if (e.key === 'Enter'){
      e.preventDefault();
      if (e.shiftKey){
        if (rowPos > 0) targetRow = rowPos - 1;
      }else{
        if (rowPos < maxRows - 1) targetRow = rowPos + 1;
      }
    }else if (e.key === 'Tab'){
      e.preventDefault();
      if (e.shiftKey){
        if (colPos > 0){
          targetCol = colPos - 1;
        }else if (rowPos > 0){
          targetRow = rowPos - 1;
          targetCol = maxCols - 1;
        }
      }else{
        if (colPos < maxCols - 1){
          targetCol = colPos + 1;
        }else if (rowPos < maxRows - 1){
          targetRow = rowPos + 1;
          targetCol = 0;
        }
      }
    }
    if (targetRow === rowPos && targetCol === colPos) return;
    focusCell(targetRow, targetCol);
  }
  function focusCell(rowPos, colPos){
    const cell = tbody.querySelector(`td[data-rpos="${rowPos}"][data-cpos="${colPos}"]`);
    if (!cell) return;
    cell.focus();
    document.getSelection()?.selectAllChildren(cell);
  }
  function buildViewMatrix(options = {}){
    const useTransposed = (typeof options.forceTransposed === 'boolean') ? options.forceTransposed : isTransposed;
    if (useTransposed){
      const headerRow = ['Field', ...filteredIdx.map(getRowLabel)];
      const body = visibleColumns.map(col => {
        const row = [col];
        for (const idx of filteredIdx){
          row.push(rows[idx]?.[col] ?? "");
        }
        return row;
      });
      return [headerRow, ...body];
    }
    const header = [...visibleColumns];
    const body = filteredIdx.map(i => visibleColumns.map(c => rows[i]?.[c] ?? ""));
    return [header, ...body];
  }
  function getTermWeekRange(term){
    if (term === 'T2') return { start:14, end:26 };
    if (term === 'T3') return { start:27, end:39 };
    return { start:1, end:13 };
  }
  function detectLessonTerm(label){
    if (!label) return null;
    const text = String(label).toUpperCase();
    const match = text.match(/L\s*(\d+)/);
    if (!match) return null;
    const num = Number(match[1]);
    if (Number.isNaN(num)) return null;
    if (num >= 1 && num <= 13) return { term:'T1', number:num };
    if (num >= 14 && num <= 26) return { term:'T2', number:num };
    if (num >= 27 && num <= 39) return { term:'T3', number:num };
    return { term:'T1', number:num };
  }
  function getTermLabel(term){
    if (term === 'T2') return 'Term 2';
    if (term === 'T3') return 'Term 3';
    return 'Term 1';
  }
  function getPrintMarkingSelectedColumns(ctx){
    if (!ctx) return [];
    const source = getPrintMarkingSourceColumns(ctx);
    const selected = new Set(ensurePrintMarkingSelection(ctx));
    if (!selected.size) return [];
    const ordered = source.filter(col => selected.has(col));
    return ordered;
  }
  function buildTemplateColumns(ctx, templateId){
    const base = [
      { id:'index', label:'#', className:'index-col' },
      { id:'name', label:'Student Name', className:'student-name-col' }
    ];
    if (templateId === 'attendance'){
      const range = getTermWeekRange(printMeta.term || 'T1');
      for (let week = range.start; week <= range.end; week++){
        base.push({ id:`w${week}`, label:`Week ${week}` });
      }
    }else if (templateId === 'drill'){
      for (let i = 1; i <= 23; i += 2){
        base.push({ id:`drill${i}`, label:`Drill ${i}` });
      }
    }else if (templateId === 'marking'){
      const selectedCols = getPrintMarkingSelectedColumns(ctx);
      selectedCols.forEach((label, idx) => {
        base.push({
          id:`marking-col-${idx}`,
          label,
          term: detectLessonTerm(label)?.term || null
        });
      });
    }
    return base;
  }
  function getContextDisplayName(ctx){
    if (!ctx) return '__________';
    return ctx.name || '__________';
  }
  function deriveContextStudentName(ctx, row, displayIndex){
    if (!ctx || !row) return `Student ${displayIndex + 1}`;
    if (ctx.studentNameColumn){
      const text = String(row[ctx.studentNameColumn] ?? '').trim();
      if (text) return text;
    }
    const firstKey = ctx.firstNameKey;
    const lastKey = ctx.lastNameKey;
    const first = firstKey ? String(row[firstKey] ?? '').trim() : '';
    const last = lastKey ? String(row[lastKey] ?? '').trim() : '';
    const combined = `${first} ${last}`.trim();
    if (combined) return combined;
    const fallbackKey = ctx.columnOrder?.[0];
    if (fallbackKey){
      const fallback = String(row[fallbackKey] ?? '').trim();
      if (fallback) return fallback;
    }
    return `Student ${displayIndex + 1}`;
  }
  function buildPrintRowOrder(ctx){
    if (!ctx || !Array.isArray(ctx.rows) || !ctx.rows.length) return [];
    const total = ctx.rows.length;
    const visibleSet = (ctx.visibleRowSet instanceof Set && ctx.visibleRowSet.size)
      ? ctx.visibleRowSet
      : new Set(Array.from({ length: total }, (_, i) => i));
    const q = (ctx.searchText || '').toLowerCase().trim();
    const searchColumns = (Array.isArray(ctx.visibleColumns) && ctx.visibleColumns.length
      ? ctx.visibleColumns
      : Array.isArray(ctx.columnOrder) && ctx.columnOrder.length
        ? ctx.columnOrder
        : ctx.allColumns || []);
    const order = [];
    for (let i = 0; i < total; i++){
      if (!visibleSet.has(i)) continue;
      if (!q){
        order.push(i);
        continue;
      }
      const row = ctx.rows[i] || {};
      const hit = searchColumns.some(col => String(row?.[col] ?? '').toLowerCase().includes(q));
      if (hit) order.push(i);
    }
    const state = ctx.sortState || {};
    if (state.key && state.dir){
      const key = state.key;
      const dir = state.dir;
      order.sort((a,b) => {
        const va = ctx.rows[a]?.[key];
        const vb = ctx.rows[b]?.[key];
        if (va == null && vb == null) return 0;
        if (va == null) return dir === 'asc' ? -1 : 1;
        if (vb == null) return dir === 'asc' ? 1 : -1;
        const na = Number(va);
        const nb = Number(vb);
        const bothNum = !Number.isNaN(na) && !Number.isNaN(nb);
        const cmp = bothNum ? (na - nb) : String(va).localeCompare(String(vb));
        return dir === 'asc' ? cmp : -cmp;
      });
    }
    return order;
  }
  function getContextRowsForPrint(ctx){
    if (!ctx || !Array.isArray(ctx.rows)) return [];
    const order = buildPrintRowOrder(ctx);
    const rowsForPrint = order.map((rowIdx, displayIndex) => {
      const row = ctx.rows[rowIdx] || {};
      const name = deriveContextStudentName(ctx, row, displayIndex);
      return { rowIdx, name };
    });
    for (let i = 0; i < PRINT_EXTRA_BLANK_ROWS; i++){
      rowsForPrint.push({ rowIdx: null, name: '' });
    }
    return rowsForPrint;
  }
  function createPrintEmptyState(message){
    const wrap = document.createElement('div');
    wrap.className = 'print-preview-empty';
    wrap.textContent = message;
    return wrap;
  }
  function createUnifiedPrintFooter(classDisplay, termLabel, teacherName){
    const footer = document.createElement('div');
    footer.className = 'print-footer';
    const left = document.createElement('div');
    left.className = 'print-footer-left';
    left.textContent = classDisplay || '[class info]';
    const center = document.createElement('div');
    center.className = 'print-footer-center';
    center.textContent = termLabel || '';
    const right = document.createElement('div');
    right.className = 'print-footer-right';
    right.textContent = teacherName || '[teacher name]';
    footer.appendChild(left);
    footer.appendChild(center);
    footer.appendChild(right);
    return footer;
  }
  function normalizePrintTemplateId(value){
    return Object.prototype.hasOwnProperty.call(PRINT_TEMPLATE_DEFS, value) ? value : null;
  }
  function resolveSavedPrintTemplate(settings){
    const direct = normalizePrintTemplateId(settings?.printTemplateId);
    if (direct) return direct;
    const legacyOrder = ['attendance', 'marking', 'drill', 'reportCard'];
    const legacyList = Array.isArray(settings?.printTemplateIds) ? settings.printTemplateIds : [];
    for (const key of legacyOrder){
      if (legacyList.includes(key) && normalizePrintTemplateId(key)) return key;
    }
    return 'attendance';
  }
  function setSelectedPrintTemplate(templateId, persist = true){
    const next = normalizePrintTemplateId(templateId) || 'attendance';
    selectedTemplateId = next;
    printTemplateInputs.forEach(input => {
      input.checked = input.value === next;
    });
    if (persist){
      saveSettings({ printTemplateId: next });
    }
  }
  function getSelectedPrintTemplates(){
    if (!normalizePrintTemplateId(selectedTemplateId)){
      setSelectedPrintTemplate('attendance');
    }else{
      setSelectedPrintTemplate(selectedTemplateId, false);
    }
    return [selectedTemplateId];
  }
  function getEffectivePrintContextIds(){
    if (printAllClasses){
      return fileContexts.map(ctx => ctx.id);
    }
    if (activeContext?.id != null){
      return [activeContext.id];
    }
    if (fileContexts.length){
      return [fileContexts[0].id];
    }
    return [];
  }
  function getSelectedPrintContexts(){
    const valid = new Map(fileContexts.map(ctx => [ctx.id, ctx]));
    const contexts = [];
    getEffectivePrintContextIds().forEach(id => {
      if (valid.has(id)) contexts.push(valid.get(id));
    });
    if (!contexts.length && activeContext && valid.has(activeContext.id)){
      contexts.push(activeContext);
    }
    return contexts;
  }
  function renderTemplatePage(ctx, templateId){
    if (templateId === 'reportCard'){
      return renderReportCards(ctx);
    }
    const columns = buildTemplateColumns(ctx, templateId);
    if (!columns.length) return null;
    const students = getContextRowsForPrint(ctx);
    const columnWidths = computePrintColumnWidths(students);
    const applyColumnWidth = (node, columnId) => {
      if (!node) return;
      const width = columnWidths[columnId];
      if (!width) return;
      node.style.minWidth = `${width}px`;
      node.style.width = `${width}px`;
    };
    const teacherName = (printMeta.teacher || '').trim();
    const classDisplay = (printMeta.classInfo || '').trim();
    const termLabel = getTermLabel(printMeta.term);
    const templateLabel = PRINT_TEMPLATE_DEFS[templateId]?.label || 'Template';

    const page = document.createElement('div');
    page.className = 'class-template-page';

    const header = document.createElement('div');
    header.className = 'page-header';
    const title = document.createElement('div');
    title.className = 'page-title';
    title.textContent = templateLabel;
    header.appendChild(title);

    const meta = document.createElement('div');
    meta.className = 'page-meta';
    const pieces = [`Class: ${classDisplay}`, `Teacher: ${teacherName}`];
    if (printMeta.title){
      pieces.push(printMeta.title);
    }
    meta.textContent = pieces.join('   ');
    header.appendChild(meta);
    page.appendChild(header);

    if (templateId === 'drill'){
      const drillNote = document.createElement('div');
      drillNote.className = 'drill-note';
      drillNote.innerHTML = `<strong>Drill Type(s):</strong> _______________________________`;
      page.appendChild(drillNote);
    }

    const table = document.createElement('table');
    table.className = 'template-table';
    if (templateId === 'marking'){
      table.classList.add('template-table-marking');
      table.style.tableLayout = 'fixed';
    }
    const colgroup = document.createElement('colgroup');
    columns.forEach(col => {
      const colEl = document.createElement('col');
      if (col.id === 'index' && columnWidths.index){
        colEl.style.width = `${columnWidths.index}px`;
      }else if (col.id === 'name' && columnWidths.name){
        colEl.style.width = `${columnWidths.name}px`;
      }else if (templateId === 'marking'){
        colEl.style.width = '26px';
      }
      colgroup.appendChild(colEl);
    });
    table.appendChild(colgroup);
    const thead = document.createElement('thead');
    if (templateId === 'marking') {
      // Top row: rotated assignment headers (visual top)
      const topRow = document.createElement('tr');
      columns.forEach(col => {
        const th = document.createElement('th');
        if (col.id === 'index' || col.id === 'name') {
          if (col.id === 'name') th.className = 'student-name-header';
          th.innerHTML = '&nbsp;';
          applyColumnWidth(th, col.id);
        } else {
          th.className = 'rotate-header';
          const span = document.createElement('div');
          span.className = 'rotate-text';
          span.textContent = col.label;
          th.appendChild(span);
        }
        topRow.appendChild(th);
      });
      thead.appendChild(topRow);

      // Second row: index and Student Name labels (prevents Student Name cell from stretching)
      const secondRow = document.createElement('tr');
      columns.forEach(col => {
        const th = document.createElement('th');
        if (col.id === 'index' || col.id === 'name') {
          th.textContent = col.label;
          if (col.id === 'name') th.classList.add('student-name-header');
          applyColumnWidth(th, col.id);
        } else {
          th.className = 'assignment-spacer';
          th.innerHTML = '&nbsp;';
        }
        secondRow.appendChild(th);
      });
      thead.appendChild(secondRow);
    } else {
      const headRow = document.createElement('tr');
      columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col.label;
        if (col.className) th.classList.add(col.className);
        applyColumnWidth(th, col.id);
        headRow.appendChild(th);
      });
      thead.appendChild(headRow);
    }
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    if (!students.length){
      const emptyRow = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = columns.length;
      cell.textContent = 'No students available.';
      emptyRow.appendChild(cell);
      tbody.appendChild(emptyRow);
    }else{
      students.forEach((entry, idx) => {
        const tr = document.createElement('tr');
        columns.forEach(col => {
          const td = document.createElement('td');
          if (col.id === 'index'){
            td.textContent = String(idx + 1);
            applyColumnWidth(td, 'index');
          }else if (col.id === 'name'){
            td.textContent = entry.name || '';
            td.classList.add('student-name-cell');
            applyColumnWidth(td, 'name');
          }else{
            td.innerHTML = '&nbsp;';
          }
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }
    table.appendChild(tbody);
    page.appendChild(table);

    page.appendChild(createUnifiedPrintFooter(classDisplay, termLabel, teacherName));
    return page;
  }
  function renderReportCards(ctx){
    if (!ctx || !Array.isArray(ctx.rows)) return null;
    const frag = document.createDocumentFragment();
    const indices = getPrintableRowIndices(ctx);
    if (!indices.length){
      const empty = document.createElement('div');
      empty.className = 'print-preview-empty';
      empty.textContent = 'No students available.';
      return empty;
    }
    const markColumns = getReportCardColumns(ctx);
    if (!markColumns.length){
      const empty = document.createElement('div');
      empty.className = 'print-preview-empty';
      empty.textContent = 'No mark columns available for this term.';
      return empty;
    }
    const teacherName = (printMeta.teacher || '').trim();
    const classDisplay = (printMeta.classInfo || '').trim();
    const termLabel = getTermLabel(printMeta.term);
    const title = (printMeta.title || '').trim();

    const targetIdx = (reportStudentIndex != null && indices.includes(reportStudentIndex)) ? reportStudentIndex : indices[0];
    const row = ctx.rows[targetIdx] || {};
    const displayIdx = indices.indexOf(targetIdx);
    const studentName = deriveContextStudentName(ctx, row, displayIdx);

      const page = document.createElement('div');
      page.className = 'report-card-page';

      const header = document.createElement('div');
      header.className = 'report-card-header';
      header.innerHTML = `
        <div class="report-card-meta">
          <div class="rc-name">${escapeHtml(studentName || `Student ${displayIdx + 1}`)}</div>
          <div class="rc-line">${escapeHtml(classDisplay)} · ${escapeHtml(termLabel)}</div>
          ${title ? `<div class="rc-line">${escapeHtml(title)}</div>` : ''}
          <div class="rc-line">Teacher: ${escapeHtml(teacherName)}</div>
        </div>
      `;
      page.appendChild(header);

    const tableWrap = document.createElement('div');
    tableWrap.className = 'report-card-table-wrap';
      const table = document.createElement('table');
      table.className = 'report-card-table report-card-table-transposed';
      const thead = document.createElement('thead');
      const headRow = document.createElement('tr');
      ['Field', 'Value'].forEach(label => {
        const th = document.createElement('th');
        th.textContent = label;
        headRow.appendChild(th);
      });
      thead.appendChild(headRow);
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      markColumns.forEach(col => {
        const tr = document.createElement('tr');
        const labelTd = document.createElement('td');
        labelTd.textContent = cleanAssignmentLabel(col);
        const markTd = document.createElement('td');
        const meta = deriveMarkMeta(row[col], col);
        markTd.textContent = meta ? formatMarkText(meta, meta.raw) : (row[col] || '').toString();
        tr.appendChild(labelTd);
        tr.appendChild(markTd);
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      tableWrap.appendChild(table);
      page.appendChild(tableWrap);

      const commentBox = document.createElement('div');
      commentBox.className = 'report-card-comment';
      const commentsText = getReportCardCommentsText(ctx, studentName);
      commentBox.textContent = commentsText ? commentsText : 'Comments:';
      page.appendChild(commentBox);

      page.appendChild(createUnifiedPrintFooter(classDisplay, termLabel, teacherName));

      frag.appendChild(page);
    return frag;
  }
  function getReportCardColumns(ctx){
    if (!ctx || !Array.isArray(ctx.columnOrder)) return [];
    const reserved = new Set([END_OF_LINE_COL, ctx.studentNameColumn, ctx.firstNameKey, ctx.lastNameKey].filter(Boolean));
    const term = (printMeta.term || '').toUpperCase();
    const range = term === 'T1' ? {min:1,max:13} : term === 'T2' ? {min:14,max:26} : term === 'T3' ? {min:27,max:39} : null;
    const selected = [];
    ctx.columnOrder.forEach(col => {
      if (!col || reserved.has(col)) return;
      const num = getLessonNumber(col);
      if (range){
        if (num == null) return;
        if (num < range.min || num > range.max) return;
        if (!columnHasData(col)) return;
        selected.push(col);
      }else if (num != null){
        if (!columnHasData(col)) return;
        selected.push(col);
      }
    });
    if (selected.length) return selected;
    // Fallback: any likely mark columns with data, optionally term-checked via detectLessonTerm
    ctx.columnOrder.forEach(col => {
      if (!col || reserved.has(col)) return;
      if (!isLikelyMarkColumn(col)) return;
      if (!columnHasData(col)) return;
      if (range){
        const termData = detectLessonTerm(col);
        if (!termData || termData.term !== term) return;
      }
      selected.push(col);
    });
    return selected;
  }
  function getReportCardCommentsText(ctx, studentName){
    const name = studentName || getReportStudentName(ctx);
    const pronouns = reportCardPronoun === 'female'
      ? { he:'she', him:'her', his:'her', He:'She', Him:'Her', His:'Her' }
      : { he:'he', him:'him', his:'his', He:'He', Him:'Him', His:'His' };
    const replaceText = (text) => {
      if (!text) return '';
      return text
        .replace(/\[Student\]/g, name)
        .replace(/\[student\]/g, name)
        .replace(/\[Student Name\]/g, name)
        .replace(/\[he\/she\]/gi, (m) => m === m.toUpperCase() ? pronouns.He : pronouns.he)
        .replace(/\[him\/her\]/gi, (m) => m === m.toUpperCase() ? pronouns.Him : pronouns.him)
        .replace(/\[his\/her\]/gi, (m) => m === m.toUpperCase() ? pronouns.His : pronouns.his);
    };
    const lines = Array.from(reportCardCommentsMap.values()).map(base => replaceText(base));
    return lines.filter(Boolean).join('\n');
  }
  function getReportStudentName(ctx){
    const indices = getPrintableRowIndices(ctx);
    const targetIdx = (reportStudentIndex != null && indices.includes(reportStudentIndex)) ? reportStudentIndex : indices[0];
    const row = ctx?.rows?.[targetIdx];
    const displayIdx = indices.indexOf(targetIdx);
    return deriveContextStudentName(ctx, row, displayIdx);
  }
  function buildPrintCommentsPanel(){
    if (!printCommentsList || !Array.isArray(REPORT_COMMENT_BANK)) return;
    printCommentsList.innerHTML = '';
    REPORT_COMMENT_BANK.forEach(section => {
      const wrap = document.createElement('div');
      wrap.className = 'print-comments-section';
      const title = document.createElement('div');
      title.className = 'print-comments-title';
      title.textContent = section.title;
      wrap.appendChild(title);
      section.options.forEach(opt => {
        const label = document.createElement('label');
        label.className = 'print-comment-row';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = reportCardCommentsMap.has(opt.text);
        cb.addEventListener('change', () => {
          if (cb.checked){
            reportCardCommentsMap.set(opt.text, opt.text);
          }else{
            reportCardCommentsMap.delete(opt.text);
          }
          renderPrintPreview();
        });
        const span = document.createElement('span');
        span.textContent = opt.label;
        label.appendChild(cb);
        label.appendChild(span);
        wrap.appendChild(label);
      });
      printCommentsList.appendChild(wrap);
    });
  }
  function togglePrintCommentsDrawer(show, forceHide = false){
    if (!printCommentsDrawer) return;
    if (forceHide){
      printCommentsDrawer.classList.remove('show');
      return;
    }
    const isShow = !!show;
    printCommentsDrawer.classList.toggle('show', isShow);
    if (isShow){
      updatePrintPronounButtons();
      buildPrintCommentsPanel();
    }
  }
  function updatePrintPronounButtons(){
    [printCommentsPronounMale, printCommentsPronounFemale].forEach(btn => {
      if (!btn) return;
      const isActive = (btn.value === 'female' ? 'female' : 'male') === reportCardPronoun;
      btn.checked = isActive;
    });
  }
  function renderPrintPreview(){
    if (!printPreviewContent || !printContainer) return;
    if (isPrintTabActive()){
      document.body.classList.add('print-preview');
    }
    printPreviewContent.innerHTML = "";
    movePrintContainerToPreview();
    printContainer.innerHTML = "";
    const contexts = getSelectedPrintContexts();
    if (!contexts.length){
      printContainer.appendChild(createPrintEmptyState('Select at least one roster to print.'));
      return;
    }
    const templates = getSelectedPrintTemplates();
    if (!templates.length){
      printContainer.appendChild(createPrintEmptyState('Select at least one template.'));
      return;
    }
    // Sync student selector with current active context
    buildReportStudentOptions(activeContext);
    buildPrintCommentsPanel();
    const stack = document.createElement('div');
    stack.className = 'template-page-stack';
    contexts.forEach(ctx => {
      templates.forEach(templateId => {
        const page = renderTemplatePage(ctx, templateId);
        if (page) stack.appendChild(page);
      });
    });
    if (!stack.children.length){
      stack.appendChild(createPrintEmptyState('No students available to print.'));
    }
    printContainer.appendChild(stack);
  }
  function movePrintContainerToPreview(){
    if (!printPreviewContent || !printContainer) return;
    printContainer.style.display = 'block';
    if (printContainer.parentElement !== printPreviewContent){
      printPreviewContent.appendChild(printContainer);
    }
  }
  function movePrintContainerToPrintRoot(){
    if (!printContainer) return;
    printContainer.style.display = 'none';
    if (printContainer.parentElement !== document.body){
      document.body.appendChild(printContainer);
    }
  }
  function syncPrintClassSelection(){
    if (printAllClasses){
      selectedClassIds = new Set(fileContexts.map(ctx => ctx.id));
      return;
    }
    const validIds = new Set(fileContexts.map(ctx => ctx.id));
    selectedClassIds = new Set([...selectedClassIds].filter(id => validIds.has(id)));
    if (!selectedClassIds.size){
      if (activeContext){
        selectedClassIds.add(activeContext.id);
      }else if (fileContexts.length){
        selectedClassIds.add(fileContexts[0].id);
      }
    }
  }
  function handlePrintClassSelection(id, checked){
    if (checked){
      selectedClassIds.add(id);
    }else{
      selectedClassIds.delete(id);
    }
    if (!selectedClassIds.size && activeContext){
      selectedClassIds.add(activeContext.id);
    }
    renderPrintPreview();
  }
  function renderPrintClassList(){
    if (!printClassList) return;
    syncPrintClassSelection();
    printClassList.innerHTML = "";
    if (!fileContexts.length){
      const msg = document.createElement('div');
      msg.className = 'muted';
      msg.style.fontSize = '12px';
      msg.textContent = 'Load files to enable printing.';
      printClassList.appendChild(msg);
      return;
    }
    fileContexts.forEach(ctx => {
      const label = document.createElement('label');
      label.className = 'print-class-option';
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = String(ctx.id);
      checkbox.checked = selectedClassIds.has(ctx.id);
      checkbox.addEventListener('change', () => handlePrintClassSelection(ctx.id, checkbox.checked));
      const name = document.createElement('span');
      name.textContent = ctx.name;
      const meta = document.createElement('small');
      meta.textContent = `${ctx.rows.length} students`;
      label.appendChild(checkbox);
      label.appendChild(name);
      label.appendChild(meta);
      printClassList.appendChild(label);
    });
    if (isPrintTabActive()) renderPrintMarkingColumnPanel();
  }
  function setDataActionButtonsDisabled(disabled){
    [exportCsv, exportXlsx, resetBtn].forEach(btn => {
      if (btn) btn.disabled = disabled;
    });
  }
  function updateTransposeButton(){
    transposeBtn.textContent = `Transpose: ${isTransposed ? 'ON' : 'OFF'}`;
    transposeBtn.classList.toggle('active', isTransposed);
    document.body.classList.toggle('is-transposed', isTransposed);
    if (transposeBanner){
      transposeBanner.classList.toggle('show', isTransposed);
    }
  }
  function setWrapHeaders(enabled, persist = true){
    wrapHeadersEnabled = !!enabled;
    document.body.classList.toggle('wrap-headers', wrapHeadersEnabled);
    if (wrapHeadersToggle) wrapHeadersToggle.checked = wrapHeadersEnabled;
    if (persist) saveSettings({ wrapHeaders: wrapHeadersEnabled });
  }
  function setFilesDrawerExpanded(expanded, persist = true){
    const isExpanded = !!expanded;
    if (filesDrawer){
      filesDrawer.classList.toggle('collapsed', !isExpanded);
      filesDrawer.setAttribute('aria-expanded', isExpanded ? 'true' : 'false');
    }
    if (filesDrawerToggle){
      filesDrawerToggle.setAttribute('aria-expanded', isExpanded ? 'true' : 'false');
      filesDrawerToggle.textContent = isExpanded ? '📂' : '📁';
      filesDrawerToggle.title = isExpanded ? 'Collapse files panel' : 'Expand files panel';
    }
    if (appRoot){
      appRoot.classList.toggle('files-drawer-collapsed', !isExpanded);
      appRoot.classList.toggle('files-drawer-expanded', isExpanded);
    }
    if (persist) saveSettings({ filesDrawerExpanded: isExpanded });
  }
  function setDarkMode(enabled, persist = true){
    darkModeEnabled = !!enabled;
    document.body.classList.toggle('dark-mode', darkModeEnabled);
    if (darkModeBtn){
      darkModeBtn.textContent = darkModeEnabled ? '☀' : '◐';
      darkModeBtn.title = darkModeEnabled ? 'Switch to light mode' : 'Switch to dark mode';
      darkModeBtn.setAttribute('aria-label', darkModeEnabled ? 'Switch to light mode' : 'Switch to dark mode');
      darkModeBtn.classList.toggle('active', darkModeEnabled);
    }
    if (persist) saveSettings({ darkMode: darkModeEnabled });
  }
  function setZoom(level, persist = true){
    let next = Number(level);
    if (Number.isNaN(next)) next = 1;
    next = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, next));
    zoomLevel = next;
    document.documentElement.style.setProperty('--zoom-factor', zoomLevel);
    if (zoomInBtn) zoomInBtn.disabled = zoomLevel >= MAX_ZOOM - 0.001;
    if (zoomOutBtn) zoomOutBtn.disabled = zoomLevel <= MIN_ZOOM + 0.001;
    if (persist) saveSettings({ zoom: zoomLevel });
  }
  function setupStudentNameColumn(){
    firstNameKey = null;
    lastNameKey = null;
    studentNameColumn = null;
    studentNameWarning = '';
    if (!rows.length){
      return;
    }
    const findByName = (label) => allColumns.find(c => normalizeHeader(c) === label) || null;
    firstNameKey = findByName('first name');
    lastNameKey = findByName('last name');
    if (!firstNameKey || !lastNameKey){
      const label = currentFileName ? ` in ${currentFileName}` : '';
      studentNameWarning = `Missing "First Name" and "Last Name" columns${label}.`;
      return;
    }
    studentNameColumn = STUDENT_NAME_LABEL;
    const prioritizeStudentColumn = (list) => {
      const base = Array.isArray(list) ? list.filter(c => c !== studentNameColumn) : [];
      base.unshift(studentNameColumn);
      return base;
    };
    allColumns = prioritizeStudentColumn(allColumns);
    columnOrder = prioritizeStudentColumn(columnOrder);
    addStudentNames(rows);
    addStudentNames(originalRows);
    if (!visibleColumns.includes(studentNameColumn)){
      visibleColumns = prioritizeStudentColumn(visibleColumns);
      saveSettings({ visibleColumns });
    }else{
      visibleColumns = prioritizeStudentColumn(visibleColumns);
    }
    syncVisibleOrder();
  }
  function applyDefaultStudentSort(){
    if (!rows.length) return;
    if (!studentNameColumn) return;
    sortState = { key: studentNameColumn, dir: 'asc' };
    saveSettings({ sortState });
  }
  function updateMeta(){
    if (!activeContext){
      metaEl.textContent = 'No file loaded';
      return;
    }
    metaEl.textContent = `${activeContext.name} - ${rows.length} rows, ${allColumns.length} columns`;
  }

  // ==== Sorting ====
  function cycleSort(col, th){
    const prev = {...sortState};
    // cycle: null -> asc -> desc -> null
    if (sortState.key !== col){ sortState = {key:col, dir:'asc'}; }
    else if (sortState.dir === 'asc'){ sortState.dir = 'desc'; }
    else if (sortState.dir === 'desc'){ sortState = {key:null, dir:null}; }
    else { sortState = {key:col, dir:'asc'}; }
    pushUndo({type:'sort', before:prev, after:{...sortState}});
    saveSettings({ sortState });
    applySort();
    render(); // re-draw header arrows
  }
  function applySort(){
    if (!sortState.key || !sortState.dir) return;
    const key = sortState.key, dir = sortState.dir;
    filteredIdx.sort((a,b) => {
      const va = rows[a]?.[key]; const vb = rows[b]?.[key];
      if (va == null && vb == null) return 0;
      if (va == null) return dir==='asc'? -1 : 1;
      if (vb == null) return dir==='asc'? 1 : -1;
      const na = Number(va), nb = Number(vb);
      const bothNum = !Number.isNaN(na) && !Number.isNaN(nb);
      const cmp = bothNum ? (na - nb) : String(va).localeCompare(String(vb));
      return dir==='asc' ? cmp : -cmp;
    });
  }

  // ==== Filtering ====
  searchEl.addEventListener('input', () => {
    const prev = loadSettings().lastSearch || "";
    const next = searchEl.value;
    pushUndo({type:'search', beforeText: prev, afterText: next});
    saveSettings({ lastSearch: next });
    if (activeContext) activeContext.searchText = next;
    render();
  });
  function filterRows(query){
    const q = (query || "").toLowerCase().trim();
    filteredIdx = [];
    const activeSet = visibleRowSet instanceof Set ? visibleRowSet : new Set(Array.from(rows.keys()));
    if (!activeSet.size) return; // all rows hidden
    for (let i = 0; i < rows.length; i++) {
      if (!activeSet.has(i)) continue;
      if (!q) { filteredIdx.push(i); continue; }
      const r = rows[i];
      const hit = visibleColumns.some(c => String(r[c] ?? "").toLowerCase().includes(q));
      if (hit) filteredIdx.push(i);
    }
  }

  // ==== Auto-fit widths ====
  function autoFit(){
    if (isTransposed) return;
    const maxChars = 30;
    const headerCells = thead.querySelectorAll('th');
    const widths = visibleColumns.map(c => Math.min(maxChars, Math.max(6, String(c).length)) * 9 + 18);
    const sample = Math.min(filteredIdx.length, 300);
    for (let r = 0; r < sample; r++) {
      const obj = rows[filteredIdx[r]];
      visibleColumns.forEach((c, i) => {
        const w = Math.min(maxChars, String(obj[c] ?? '').length) * 9 + 18;
        widths[i] = Math.max(widths[i], w);
      });
    }
    headerCells.forEach((th, i) => th.style.width = widths[i] + 'px');
    const firstRow = tbody.querySelector('tr');
    if (firstRow) {
      firstRow.querySelectorAll('td').forEach((td, i) => td.style.width = widths[i] + 'px');
    }
  }

  function adjustTableSize(){
    if (!tableWrapper || !table) return;
    if (isTransposed){
      table.style.tableLayout = 'auto';
      const minWidth = Math.max(filteredIdx.length, 1) * 160 + 220;
      table.style.minWidth = `${minWidth}px`;
      table.style.width = `${minWidth}px`;
      tableWrapper.style.overflowX = 'auto';
    }else{
      table.style.tableLayout = 'fixed';
      table.style.minWidth = '';
      table.style.width = '100%';
      tableWrapper.style.overflowX = 'auto';
    }
  }

  // ==== Export / Reset ====
  selectAllColsBtn.addEventListener('click', () => {
    const before = [...visibleColumns];
    visibleColumns = columnOrder.slice();
    pushUndo({type:'setVisibleColumns', before, after:[...visibleColumns]});
    saveSettings({ visibleColumns });
    rebuildColsUI();
    render();
  });
  clearAllColsBtn.addEventListener('click', () => {
    const before = [...visibleColumns];
    visibleColumns = [];
    pushUndo({type:'setVisibleColumns', before, after:[...visibleColumns]});
    saveSettings({ visibleColumns });
    rebuildColsUI();
    render();
  });
  if (termBtn1) termBtn1.addEventListener('click', () => selectLessonRange(1, 13));
  if (termBtn2) termBtn2.addEventListener('click', () => selectLessonRange(14, 26));
  if (termBtn3) termBtn3.addEventListener('click', () => selectLessonRange(27, 39));
  selectAllRowsBtn.addEventListener('click', () => {
    setAllRowsVisible(true);
    performanceFilterLevel = null;
    updatePerformanceFilterButtons();
    updateRowFilterStatus('');
  });
  clearRowsBtn.addEventListener('click', () => {
    if (!(visibleRowSet instanceof Set)) return;
    const before = getVisibleRowList();
    visibleRowSet.clear();
    const after = getVisibleRowList();
    if (!arraysEqual(before, after)){
      pushUndo({type:'setVisibleRows', beforeRows: before, afterRows: after});
    }
    persistRowVisibility();
    rebuildRowsUI();
    render();
    performanceFilterLevel = null;
    updatePerformanceFilterButtons();
    updateRowFilterStatus('');
  });
  filesList.addEventListener('click', (e) => {
    const item = e.target.closest('.file-item');
    if (!item) return;
    const id = Number(item.dataset.id);
    if (Number.isNaN(id)) return;
    setActiveContext(id);
  });

  exportCsv.onclick = () => {
    const matrix = buildViewMatrix();
    const csv = Papa.unparse(matrix);
    const blob = new Blob([csv], {type: 'text/csv;charset=utf-8;'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const safe = (currentFileName || 'visible_export').replace(/[\\/:*?"<>|]+/g, '_');
    a.download = `${safe}.csv`;
    a.click();
  };
  exportXlsx.onclick = () => {
    const arr = buildViewMatrix();
    const ws = XLSX.utils.aoa_to_sheet(arr);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "data");
    const safe = (currentFileName || 'visible_export').replace(/[\\/:*?"<>|]+/g, '_');
    XLSX.writeFile(wb, `${safe}.xlsx`);
  };
  resetBtn.onclick = () => {
    const before = deepClone(rows);
    rows = deepClone(originalRows);
    filteredIdx = Array.from(rows.keys());
    pushUndo({type:'replaceRows', before, after: deepClone(rows)});
    searchEl.value = ""; saveSettings({ lastSearch: "" });
    sortState = {key:null, dir:null}; saveSettings({ sortState });
    const beforeRows = getVisibleRowList();
    visibleRowSet = new Set(Array.from(rows.keys()));
    const afterRows = getVisibleRowList();
    if (!arraysEqual(beforeRows, afterRows)){
      pushUndo({type:'setVisibleRows', beforeRows, afterRows});
    }
    persistRowVisibility();
    rebuildRowsUI();
    render();
    handleColumnDataMutation();
    resetPerformanceFilterState();
  };
  
  async function launchPrintDialog(){
    if (!printContainer) return;
    renderPrintPreview();
    movePrintContainerToPrintRoot();
    printContainer.style.display = 'block';
    document.body.classList.remove('print-preview');
    document.body.classList.add('printing-preview');
    try {
      window.print();
    } finally {
      printContainer.style.display = 'none';
      document.body.classList.remove('printing-preview');
      if (isPrintTabActive()){
        movePrintContainerToPreview();
        document.body.classList.add('print-preview');
      }
    }
  }

  // ==== Undo/Redo ====
  function pushUndo(op){
    undoStack.push(op);
    if (undoStack.length > UNDO_LIMIT) undoStack.shift();
    redoStack.length = 0; // clear redo on new action
  }
  function doUndo(){
    if (!undoStack.length) return;
    const op = undoStack.pop(); redoStack.push(op);
    const full = applyInverse(op);
    if (full) return;
    render();
  }
  function doRedo(){
    if (!redoStack.length) return;
    const op = redoStack.pop(); undoStack.push(op);
    const full = applyForward(op);
    if (full) return;
    render();
  }
  function applyInverse(op){
    switch(op.type){
      case 'editCell':
        rows[op.rowIndex] = rows[op.rowIndex] || {};
        rows[op.rowIndex][op.col] = op.before;
        handleColumnDataMutation(op.col);
        break;
      case 'setVisibleColumns':
        visibleColumns = [...op.before];
        syncVisibleOrder();
        rebuildColsUI();
        saveSettings({ visibleColumns });
        render();
        return true;
      case 'setVisibleRows':
        setVisibleRowsFromList(op.beforeRows);
        persistRowVisibility();
        render();
        return true;
      case 'search':
        searchEl.value = op.beforeText;
        saveSettings({ lastSearch: op.beforeText });
        render();
        return true;
      case 'sort': sortState = {...op.before}; saveSettings({ sortState }); applySort(); break;
      case 'replaceRows':
        rows = deepClone(op.before);
        filteredIdx = Array.from(rows.keys());
        handleColumnDataMutation();
        break;
      case 'renameHeaders': renameColumns(op.beforeMapping, false, op.applyAll); return true;
      case 'reorderColumns':
        columnOrder = [...op.beforeOrder];
        syncVisibleOrder();
        rebuildColsUI();
        saveSettings({ columnOrder });
        render();
        return true;
    }
    return false;
  }
  function applyForward(op){
    switch(op.type){
      case 'editCell':
        rows[op.rowIndex] = rows[op.rowIndex] || {};
        rows[op.rowIndex][op.col] = op.after;
        handleColumnDataMutation(op.col);
        break;
      case 'setVisibleColumns':
        visibleColumns = [...op.after];
        syncVisibleOrder();
        rebuildColsUI();
        saveSettings({ visibleColumns });
        render();
        return true;
      case 'setVisibleRows':
        setVisibleRowsFromList(op.afterRows);
        persistRowVisibility();
        render();
        return true;
      case 'search':
        searchEl.value = op.afterText;
        saveSettings({ lastSearch: op.afterText });
        render();
        return true;
      case 'sort': sortState = {...op.after}; saveSettings({ sortState }); applySort(); break;
      case 'replaceRows':
        rows = deepClone(op.after);
        filteredIdx = Array.from(rows.keys());
        handleColumnDataMutation();
        break;
      case 'renameHeaders': renameColumns(op.afterMapping, false); return true;
      case 'reorderColumns':
        columnOrder = [...op.afterOrder];
        syncVisibleOrder();
        rebuildColsUI();
        saveSettings({ columnOrder });
        render();
        return true;
    }
    return false;
  }

  // ==== Header Tools (Phase 1.5) ====
  editHeadersBtn.addEventListener('click', openHeaderModal);
  [hdrFind,hdrReplace,hdrRmStart,hdrRmEnd,hdrBetweenPre,hdrBetweenSuf,hdrAddPre,hdrAddSuf,hdrCase,hdrTrim,hdrCollapse,hdrLower,hdrInclusive,hdrNormalize,hdrTargetFilter]
    .forEach(el => el.addEventListener('input', refreshHeaderPreview));
  hdrCancel.addEventListener('click', closeHeaderModal);
  hdrApply.addEventListener('click', () => {
    const mapping = computeHeaderMapping();
    if (!Object.keys(mapping).length){ closeHeaderModal(); return; }
    const beforeMap = invertMapping(mapping); // reverse mapping for undo
    pushUndo({type:'renameHeaders', beforeMapping:beforeMap, afterMapping:mapping});
    renameColumns(mapping, true);
    closeHeaderModal();
  });

  let headerModalHome = null;
  let headerModalNext = null;
  function detachHeaderModal(){
    if (!modalBackdrop) return;
    if (modalBackdrop.parentElement !== document.body){
      headerModalHome = modalBackdrop.parentElement;
      headerModalNext = modalBackdrop.nextSibling;
      document.body.appendChild(modalBackdrop);
    }
  }
  function restoreHeaderModal(){
    if (!modalBackdrop || !headerModalHome) return;
    if (headerModalNext){
      headerModalHome.insertBefore(modalBackdrop, headerModalNext);
    }else{
      headerModalHome.appendChild(modalBackdrop);
    }
    headerModalHome = null;
    headerModalNext = null;
  }
  function openHeaderModal(){
    headerCountChip.textContent = `${allColumns.length} headers`;
    refreshHeaderPreview();
    detachHeaderModal();
    modalBackdrop.style.display = 'flex';
  }
  function closeHeaderModal(){
    modalBackdrop.style.display = 'none';
    restoreHeaderModal();
  }

  function computeHeaderMapping(){
    const rules = {
      find: hdrFind.value,
      replace: hdrReplace.value,
      rmStart: parseInt(hdrRmStart.value || "0", 10),
      rmEnd: parseInt(hdrRmEnd.value || "0", 10),
      bPre: hdrBetweenPre.value,
      bSuf: hdrBetweenSuf.value,
      inclusive: hdrInclusive.checked,
      addPre: hdrAddPre.value,
      addSuf: hdrAddSuf.value,
      caseIns: hdrCase.checked,
      trim: hdrTrim.checked,
      collapse: hdrCollapse.checked,
      lower: hdrLower.checked,
      norm: hdrNormalize.value,
      targetTokens: (hdrTargetFilter.value || '').split(',').map(token => token.trim()).filter(Boolean)
    };
    const normalizedTargetTokens = rules.targetTokens.map(token => rules.caseIns ? token.toLowerCase() : token);
    const shouldModifyHeader = (name) => {
      if (!normalizedTargetTokens.length) return true;
      const compareValue = rules.caseIns ? String(name).toLowerCase() : String(name);
      return normalizedTargetTokens.some(token => compareValue.includes(token));
    };
    const map = {};
    for (const name of allColumns){
      if (!shouldModifyHeader(name)) continue;
      let out = name;

      // find/replace (supports /regex/ form)
      if (rules.find){
        if (rules.find.startsWith('/') && rules.find.lastIndexOf('/') > 0){
          const last = rules.find.lastIndexOf('/');
          const pat = rules.find.slice(1, last);
          const flags = rules.caseIns ? 'i' : '';
          try{
            const re = new RegExp(pat, flags);
            out = out.replace(re, rules.replace);
          }catch{}
        }else{
          const src = rules.caseIns ? new RegExp(escapeRegExp(rules.find), 'ig') : new RegExp(escapeRegExp(rules.find), 'g');
          out = out.replace(src, rules.replace);
        }
      }

      // remove first/last N
      if (rules.rmStart>0) out = out.slice(rules.rmStart);
      if (rules.rmEnd>0) out = out.slice(0, Math.max(0, out.length - rules.rmEnd));

      // remove between prefix/suffix
      if (rules.bPre && rules.bSuf){
        const preIdx = out.indexOf(rules.bPre);
        const sufIdx = out.indexOf(rules.bSuf, preIdx + rules.bPre.length);
        if (preIdx >= 0 && sufIdx > preIdx){
          if (rules.inclusive) {
            out = out.slice(0, preIdx) + out.slice(sufIdx + rules.bSuf.length);
          } else {
            out = out.slice(0, preIdx + rules.bPre.length) + out.slice(sufIdx);
          }
        }
      }

      // trim / collapse / lower
      if (rules.trim) out = out.trim();
      if (rules.collapse) out = out.replace(/\s+/g, '_');
      if (rules.lower) out = out.toLowerCase();

      // normalize case
      if (rules.norm === 'snake'){
        out = out.trim().replace(/\s+/g, '_').replace(/[^\w]/g, '_').replace(/_+/g,'_').replace(/^_+|_+$/g,'').toLowerCase();
      } else if (rules.norm === 'title'){
        out = out.replace(/[_\-]+/g,' ').replace(/\s+/g,' ')
                 .split(' ').map(w => w ? (w[0].toUpperCase() + w.slice(1).toLowerCase()) : '').join(' ').trim();
      } else if (rules.norm === 'kebab'){
        out = out.trim().replace(/\s+/g, '-').replace(/[^\w\-]/g,'-').replace(/-+/g,'-').replace(/^-+|-+$/g,'').toLowerCase();
      }

      // add pre/suf
      if (rules.addPre) out = rules.addPre + out;
      if (rules.addSuf) out = out + rules.addSuf;

      if (out !== name) map[name] = out;
    }
    return map;
  }

  function refreshHeaderPreview(){
    const map = computeHeaderMapping();
    hdrPreview.innerHTML = "";
    const keys = allColumns;
    keys.forEach(k => {
      const after = map[k] ?? k;
      const row = document.createElement('div'); row.className = 'preview-row';
      row.innerHTML = `<span class="chip preview-chip">${escapeHtml(k)}</span> ? <span class="chip preview-chip ${after===k ? '' : 'preview-chip-modified'}">${escapeHtml(after)}</span>`;
      hdrPreview.appendChild(row);
    });
  }

  function renameColumns(mapping, persist, applyAll){
    if (!Object.keys(mapping).length) return;
    Object.keys(mapping).forEach(key => {
      if (key === END_OF_LINE_COL || mapping[key] === END_OF_LINE_COL){
        delete mapping[key];
      }
    });
    rows = remapRows(rows, mapping);
    originalRows = remapRows(originalRows, mapping);
    // remap headers
    allColumns = allColumns.map(c => mapping[c] || c);
    visibleColumns = visibleColumns.map(c => mapping[c] || c);
    columnOrder = columnOrder.map(c => mapping[c] || c);
    if (firstNameKey) firstNameKey = mapping[firstNameKey] || firstNameKey;
    if (lastNameKey) lastNameKey = mapping[lastNameKey] || lastNameKey;
    if (studentNameColumn) studentNameColumn = mapping[studentNameColumn] || studentNameColumn;
    if (commentConfig && Array.isArray(commentConfig.extraColumns)){
      commentConfig.extraColumns = commentConfig.extraColumns.map(c => mapping[c] || c);
    }
    syncVisibleOrder();
    const updates = { columnOrder, visibleColumns };
    saveSettings(updates);
    if (persist && currentFileName){
      persistHeaderMapping(currentFileName, mapping);
    }
    // re-evaluate student names if needed
    setupStudentNameColumn();
    enforceEndOfLineColumn();
    // Re-render everything
    rebuildColsUI();
    render();
  }

  function openRenameModal(ctx){
    renameTargetId = ctx.id;
    renameInput.value = ctx.name;
    renamePrompt.style.display = 'flex';
    setTimeout(() => {
      renameInput.focus();
      renameInput.select();
    }, 0);
  }
  function closeRenameModal(){
    renamePrompt.style.display = 'none';
    renameInput.value = '';
    renameTargetId = null;
  }
  function saveRename(){
    if (renameTargetId == null){ closeRenameModal(); return; }
    const name = renameInput.value.trim();
    if (!name){ closeRenameModal(); return; }
    renameContext(renameTargetId, name);
    closeRenameModal();
  }
  function renameContext(id, newName){
    const ctx = fileContexts.find(c => c.id === id);
    if (!ctx) return;
    ctx.name = newName;
    if (activeContext && activeContext.id === id){
      currentFileName = newName;
      metaEl.textContent = `${newName} - ${rows.length} rows, ${allColumns.length} columns`;
    }
    updateFilesUI();
  }
  function showMarkWarning(message, onContinue){
    if (markWarningMessage) markWarningMessage.textContent = message;
    pendingMarkWarningAction = typeof onContinue === 'function' ? onContinue : null;
    if (markWarningModal) markWarningModal.style.display = 'flex';
  }
  function closeMarkWarningModal(){
    if (markWarningModal) markWarningModal.style.display = 'none';
  }
  function activateTab(kind){
    const sections = {
      data: dataTabSection,
      print: printTabSection,
      comments: commentsTabSection
    };
    const buttons = {
      data: tabDataBtn,
      print: tabPrintBtn,
      comments: tabCommentsBtn
    };
    Object.entries(sections).forEach(([key, el]) => {
      if (!el) return;
      const isActive = key === kind;
      el.classList.toggle('active', isActive);
      el.style.display = isActive ? '' : 'none';
      if (isActive){
        el.classList.remove('fade-in');
        void el.offsetWidth; // force reflow
        el.classList.add('fade-in');
      }
    });
    Object.entries(buttons).forEach(([key, btn]) => {
      if (btn) btn.classList.toggle('active', key === kind);
    });
    const printActive = kind === 'print';
    document.body.classList.toggle('print-preview', printActive);
    if (!printActive){
      document.body.classList.remove('printing-preview');
      movePrintContainerToPrintRoot();
    }
    if (kind !== 'print'){
      togglePrintCommentsDrawer(false, true);
    }
    if (kind === 'print'){
      syncPrintMetaInputs();
      renderPrintClassList();
      renderPrintPreview();
      renderPrintMarkingColumnPanel();
    }
    if (kind === 'comments'){
      openCommentsModalInternal();
    }
  }
  function isCommentsTabActive(){
    return !!(commentsTabSection && commentsTabSection.classList.contains('active'));
  }
  function isPrintTabActive(){
    return !!(printTabSection && printTabSection.classList.contains('active'));
  }
  function openPrintPreviewModal(){
    activateTab('print');
  }
  function closePrintPreviewModal(){
    activateTab('data');
  }
  function syncPrintMetaInputs(){
    if (printTeacherInput) printTeacherInput.value = printMeta.teacher || '';
    if (printClassInput) printClassInput.value = printMeta.classInfo || '';
    if (printTitleInput) printTitleInput.value = printMeta.title || '';
    const termValue = printMeta.term || 'T1';
    printTermInputs.forEach(radio => {
      radio.checked = radio.value === termValue;
    });
    setSelectedPrintTemplate(selectedTemplateId, false);
    if (printAllClassesToggle){
      printAllClassesToggle.checked = !!printAllClasses;
    }
    buildReportStudentOptions(activeContext);
  }
  function handlePrintMetaInputChange(){
    const prevTerm = printMeta.term;
    printMeta.teacher = (printTeacherInput?.value || '').trim();
    printMeta.classInfo = (printClassInput?.value || '').trim();
    printMeta.title = (printTitleInput?.value || '').trim();
    const checkedTerm = printTermInputs.find(radio => radio.checked)?.value;
    if (checkedTerm) printMeta.term = checkedTerm;
    if (checkedTerm && checkedTerm !== prevTerm){
      // Reset marking selections to align with the newly chosen term
      fileContexts.forEach(ctx => {
        printMarkingSelections[ctx.id] = new Set(detectDefaultMarkingColumns(ctx));
      });
    }
    renderPrintPreview();
    renderPrintMarkingColumnPanel();
  }
  function handleTemplateSelectionChange(event){
    const target = event?.target;
    const next = normalizePrintTemplateId(target?.value) || selectedTemplateId || 'attendance';
    // Checkbox UI behaves like radio: exactly one template always selected.
    setSelectedPrintTemplate(next);
    renderPrintPreview();
    renderPrintMarkingColumnPanel();
  }
  function isMarkingTemplateSelected(){
    return selectedTemplateId === 'marking';
  }
  function getPrintMarkingSourceColumns(ctx){
    if (!ctx) return [];
    const source = Array.isArray(ctx.columnOrder) && ctx.columnOrder.length
      ? ctx.columnOrder
      : Array.isArray(ctx.allColumns) ? ctx.allColumns : [];
    const skip = new Set([END_OF_LINE_COL, ctx.studentNameColumn, ctx.firstNameKey, ctx.lastNameKey].filter(Boolean));
    return source.filter(col => col && !skip.has(col));
  }
  function detectDefaultMarkingColumns(ctx){
    const source = getPrintMarkingSourceColumns(ctx);
    const selectedTerm = (printMeta.term || '').toUpperCase();
    const seen = new Set();
    const defaults = [];
    source.forEach(colName => {
      const termData = detectLessonTerm(colName);
      if (!termData) return;
      if (selectedTerm && termData.term !== selectedTerm) return;
      const key = `${termData.term}:${termData.number}:${colName}`;
      if (seen.has(key)) return;
      seen.add(key);
      defaults.push(colName);
    });
    return defaults;
  }
  function ensurePrintMarkingSelection(ctx){
    if (!ctx) return [];
    if (!printMarkingSelections[ctx.id]){
      printMarkingSelections[ctx.id] = new Set(detectDefaultMarkingColumns(ctx));
    }
    const set = printMarkingSelections[ctx.id];
    if (!set.size){
      detectDefaultMarkingColumns(ctx).forEach(col => set.add(col));
    }
    return Array.from(set);
  }
  function updatePrintMarkingSelection(ctx, updater){
    if (!ctx) return;
    const current = new Set(ensurePrintMarkingSelection(ctx));
    updater(current);
    printMarkingSelections[ctx.id] = current;
    renderPrintPreview();
    renderPrintMarkingColumnPanel();
  }
  function renderPrintMarkingColumnPanel(){
    if (!printMarkingPanel) return;
    const show = isMarkingTemplateSelected();
    printMarkingPanel.style.display = show ? '' : 'none';
    if (!show){
      return;
    }
    const ctx = activeContext;
    if (!printMarkingColumnsList) return;
    printMarkingColumnsList.innerHTML = '';
    if (!ctx){
      const msg = document.createElement('div');
      msg.className = 'muted';
      msg.style.fontSize = '12px';
      msg.textContent = 'Load a class to choose columns.';
      printMarkingColumnsList.appendChild(msg);
      return;
    }
    const source = getPrintMarkingSourceColumns(ctx);
    const selected = new Set(ensurePrintMarkingSelection(ctx));
    if (!source.length){
      const msg = document.createElement('div');
      msg.className = 'muted';
      msg.style.fontSize = '12px';
      msg.textContent = 'No columns available.';
      printMarkingColumnsList.appendChild(msg);
      return;
    }
    source.forEach(col => {
      const label = document.createElement('label');
      label.className = 'print-marking-row';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = selected.has(col);
      if (cb.checked) label.classList.add('selected');
      cb.addEventListener('change', () => {
        updatePrintMarkingSelection(ctx, set => {
          if (cb.checked) set.add(col); else set.delete(col);
          label.classList.toggle('selected', cb.checked);
        });
      });
      const span = document.createElement('span');
      span.textContent = col;
      label.appendChild(cb);
      label.appendChild(span);
      printMarkingColumnsList.appendChild(label);
    });
  }
  function getPrintableRowIndices(ctx){
    if (!ctx || !Array.isArray(ctx.rows)) return [];
    if (ctx.id === activeContext?.id && Array.isArray(filteredIdx) && filteredIdx.length){
      return [...filteredIdx];
    }
    if (ctx.visibleRowSet instanceof Set && ctx.visibleRowSet.size){
      return Array.from(ctx.visibleRowSet).filter(i => i >= 0 && i < ctx.rows.length);
    }
    return Array.from({ length: ctx.rows.length }, (_, i) => i);
  }
  function buildReportStudentOptions(ctx){
    if (!printReportStudentSelect) return;
    printReportStudentSelect.innerHTML = '';
    if (!ctx || !Array.isArray(ctx.rows) || !ctx.rows.length){
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = 'No students available';
      printReportStudentSelect.appendChild(opt);
      printReportStudentSelect.disabled = true;
      togglePrintCommentsDrawer(false, true);
      return;
    }
    const indices = getPrintableRowIndices(ctx);
    if (!indices.length){
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = 'No students available';
      printReportStudentSelect.appendChild(opt);
      printReportStudentSelect.disabled = true;
      togglePrintCommentsDrawer(false, true);
      return;
    }
    printReportStudentSelect.disabled = false;
    const gradeColumn = commentConfig.gradeColumn || FINAL_GRADE_COLUMN;
    indices.forEach((rowIdx, displayIndex) => {
      const opt = document.createElement('option');
      opt.value = String(rowIdx);
      opt.textContent = deriveContextStudentName(ctx, ctx.rows[rowIdx], displayIndex);
      const meta = gradeColumn ? deriveMarkMeta(ctx.rows[rowIdx]?.[gradeColumn], gradeColumn) : null;
      applyOptionMarkStyle(opt, meta);
      printReportStudentSelect.appendChild(opt);
    });
    const fallback = indices.includes(reportStudentIndex) ? reportStudentIndex : indices[0];
    reportStudentIndex = fallback;
    printReportStudentSelect.value = String(fallback);
  }
  if (printReportStudentSelect){
    printReportStudentSelect.addEventListener('change', () => {
      const val = Number(printReportStudentSelect.value);
      reportStudentIndex = Number.isNaN(val) ? null : val;
      renderPrintPreview();
    });
  }
  function handlePrintMarkingSelectAll(){
    const ctx = activeContext;
    if (!ctx) return;
    updatePrintMarkingSelection(ctx, set => {
      set.clear();
      getPrintMarkingSourceColumns(ctx).forEach(col => set.add(col));
    });
  }
  function handlePrintMarkingClear(){
    const ctx = activeContext;
    if (!ctx) return;
    updatePrintMarkingSelection(ctx, set => set.clear());
  }
  function removeContext(id){
    const idx = fileContexts.findIndex(c => c.id === id);
    if (idx === -1) return;
    fileContexts.splice(idx, 1);
    if (activeContext && activeContext.id === id){
      activeContext = null;
      if (fileContexts.length){
        const fallback = fileContexts[Math.min(idx, fileContexts.length - 1)];
        setActiveContext(fallback.id);
      }else{
        resetToEmptyState();
        updateFilesUI();
      }
    }else{
      updateFilesUI();
    }
  }
  // ==== Report Builder (Grade 10) ====
  const BUILDER_HINTS = {
    trigTest1:[/trig.*test.*1/i,/test\s*1/i],
    trigTest2:[/trig.*test.*2/i,/test\s*2/i],
    retest:[/re[-\s]?test/i],
    original:[/original/i],
    termAverage:[/term.*average/i,/avg/i]
  };
  function initializeCommentBuilder(){
    if (!builderStudentSelect) return;
    buildBuilderStudentOptions();
    if (!builderBankRendered){
      buildBuilderCommentBank();
      builderBankRendered = true;
    }
    updateCommentOrderToggleLabel();
    if (builderSelectedRowIndex == null && rows.length){
      builderSelectedRowIndex = filteredIdx.length ? filteredIdx[0] : 0;
    }
    if (builderSelectedRowIndex != null){
      prefillBuilderFromRow(builderSelectedRowIndex);
    }
    updateBuilderSelectedTags();
    // Populate assignment columns list
    buildBuilderAssignmentsList();
  }
  function shouldAutoGenerateBuilderReport(){
    const gradeGroup = builderGradeGroupSelect?.value || 'middle';
    if (gradeGroup === 'elem'){
      return Boolean(builderStudentNameInput?.value.trim());
    }
    return Boolean(builderStudentNameInput?.value.trim() && builderCorePerformanceSelect?.value);
  }
  function toggleCommentOrderMode(){
    const modes = ['sandwich', 'selection', 'bullet'];
    const idx = modes.indexOf(commentOrderMode);
    commentOrderMode = modes[(idx + 1) % modes.length];
    updateCommentOrderToggleLabel();
    if (shouldAutoGenerateBuilderReport()){
      builderGenerateReport();
    }
  }
  function updateCommentOrderToggleLabel(){
    if (!builderCommentOrderToggle) return;
    const labelMap = {
      sandwich: 'Order: Sandwich',
      selection: 'Order: Selection',
      bullet: 'Order: Bullets'
    };
    builderCommentOrderToggle.textContent = labelMap[commentOrderMode] || 'Order: Selection';
    builderCommentOrderToggle.title = 'Toggle comment order mode (sandwich, selection, bullets)';
  }
  function buildBuilderAssignmentsList(){
    const listEl = document.getElementById('builderAssignmentsList');
    if (!listEl) return;
    
    listEl.innerHTML = '';
    
    // Get columns that have data (assignments)
    const assignmentColumns = allColumns.filter(col => {
      if (!col) return false;
      if (col === END_OF_LINE_COL) return false;
      if (col === studentNameColumn) return false;
      if (col === firstNameKey) return false;
      if (col === lastNameKey) return false;
      return columnHasData(col);
    });
    
    if (!assignmentColumns.length){
      const msg = document.createElement('div');
      msg.className = 'muted';
      msg.style.fontSize = '12px';
      msg.textContent = 'No assignment columns available.';
      listEl.appendChild(msg);
      return;
    }
    
    // Get the current student's row
    const studentRow = (builderSelectedRowIndex != null && rows[builderSelectedRowIndex]) 
      ? rows[builderSelectedRowIndex] 
      : null;
    
    assignmentColumns.forEach(col => {
      const wrapper = document.createElement('div');
      wrapper.style.display = 'flex';
      wrapper.style.alignItems = 'center';
      wrapper.style.gap = '8px';
      wrapper.style.padding = '4px 0';
      wrapper.style.borderBottom = '1px solid #eee';
      
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = col;
      checkbox.dataset.assignmentCol = col;
      checkbox.style.flexShrink = '0';
      
      const labelText = document.createElement('span');
      labelText.textContent = cleanAssignmentLabel(col);
      labelText.style.flex = '1';
      labelText.style.fontSize = '13px';
      labelText.style.textAlign = 'left';
      
      const markSpan = document.createElement('span');
      markSpan.style.fontSize = '12px';
      markSpan.style.fontWeight = '600';
      markSpan.style.color = '#666';
      markSpan.style.minWidth = '60px';
      markSpan.style.textAlign = 'right';
      markSpan.style.flexShrink = '0';
      
      // Get the student's mark for this assignment
      if (studentRow && studentRow[col] != null && String(studentRow[col]).trim() !== '') {
        const meta = deriveMarkMeta(studentRow[col], col);
        markSpan.textContent = meta ? meta.raw : String(studentRow[col]);
        if (meta && markColorsEnabled){
          markSpan.classList.add(meta.className);
        }
      } else {
        markSpan.textContent = '—';
        markSpan.style.color = '#ccc';
      }
      
      wrapper.appendChild(checkbox);
      wrapper.appendChild(labelText);
      wrapper.appendChild(markSpan);
      listEl.appendChild(wrapper);

      const handleToggle = () => {
        if (shouldAutoGenerateBuilderReport()){
          builderOutputAnimationMode = 'selection';
          builderGenerateReport();
        }
      };
      checkbox.addEventListener('change', handleToggle);
      wrapper.addEventListener('click', (e) => {
        if (e.target === checkbox) return;
        checkbox.checked = !checkbox.checked;
        checkbox.dispatchEvent(new Event('change'));
      });
    });
  }
  function buildBuilderStudentOptions(){
    if (!builderStudentSelect) return;
    builderStudentSelect.innerHTML = '';
    const indices = filteredIdx.length ? filteredIdx : rows.map((_, idx) => idx);
    const gradeColumn = commentConfig.gradeColumn || FINAL_GRADE_COLUMN;
    indices.forEach(idx => {
      const option = document.createElement('option');
      option.value = String(idx);
      option.textContent = rows[idx]?.[studentNameColumn] || buildStudentName(rows[idx]) || getRowLabel(idx);
      const meta = deriveMarkMeta(rows[idx]?.[gradeColumn], gradeColumn);
      applyOptionMarkStyle(option, meta);
      builderStudentSelect.appendChild(option);
    });
    if (!indices.length){
      const empty = document.createElement('option');
      empty.value = '';
      empty.textContent = 'No students available';
      builderStudentSelect.appendChild(empty);
      builderStudentSelect.disabled = true;
    }else{
      builderStudentSelect.disabled = false;
      builderStudentSelect.value = builderSelectedRowIndex != null ? String(builderSelectedRowIndex) : String(indices[0]);
    }
  }
  function applyOptionMarkStyle(option, meta){
    if (!option) return;
    option.classList.remove('mark-high','mark-mid','mark-low');
    option.style.backgroundColor = '';
    option.style.color = '';
    option.style.fontWeight = '';
    if (!meta) return;
    option.classList.add(meta.className);
    const palette = {
      'mark-high': { bg:'#cbe592', fg:'#1f3a07' },
      'mark-mid': { bg:'#ffe58d', fg:'#5c4d00' },
      'mark-low': { bg:'#ffc8c8', fg:'#5a0e0e' }
    };
    const colors = palette[meta.className];
    if (colors){
      option.style.backgroundColor = colors.bg;
      option.style.color = colors.fg;
      option.style.fontWeight = '600';
    }
  }
  function prefillBuilderFromRow(rowIndex){
    const row = rows[rowIndex];
    if (!row) return;
    builderSelectedRowIndex = rowIndex;
    if (builderStudentSelect) builderStudentSelect.value = String(rowIndex);
    if (builderStudentNameInput){
      builderStudentNameInput.value = getFirstNameFromRow(row, rowIndex);
    }
    const pronoun = derivePronounFromRow(row);
    if (pronoun === 'female' && builderPronounFemaleInput){
      builderPronounFemaleInput.checked = true;
    }else if (builderPronounMaleInput){
      builderPronounMaleInput.checked = true;
    }
    applyHintValue(builderTrigTest1Input, row, BUILDER_HINTS.trigTest1);
    applyHintValue(builderTrigTest2Input, row, BUILDER_HINTS.trigTest2);
    applyHintValue(builderRetestInput, row, BUILDER_HINTS.retest);
    applyHintValue(builderOriginalInput, row, BUILDER_HINTS.original);
    if (commentConfig.gradeColumn){
      const val = row?.[commentConfig.gradeColumn];
      if (builderTermAverageInput && val != null && val !== ''){
        builderTermAverageInput.value = parseMarkValue(val);
      }
    }
    applyHintValue(builderTermAverageInput, row, BUILDER_HINTS.termAverage, true);
    // Refresh assignment list to show new student's marks
    buildBuilderAssignmentsList();
    autoSelectBuilderTemplate(row);
    applyRandomLookingAheadSelection();
  }
  function derivePronounFromRow(row){
    if (!row) return 'male';
    for (const col of allColumns){
      const header = (col || '').toLowerCase();
      if (!/(gender|pronoun)/.test(header)) continue;
      const value = String(row[col] || '').toLowerCase();
      if (!value) continue;
      if (value.includes('female') || value.includes('she') || value.includes('girl')) return 'female';
      if (value.includes('male') || value.includes('he') || value.includes('boy')) return 'male';
    }
    return 'male';
  }
  function applyHintValue(input, row, hints, override){
    if (!input || !row || !Array.isArray(hints)) return;
    for (const col of allColumns){
      if (!col) continue;
      const match = hints.some(rx => rx.test(col));
      if (!match) continue;
      const value = row[col];
      if (value == null || String(value).trim() === '') continue;
      if (input.value && !override) return;
      input.value = parseMarkValue(value);
      return;
    }
  }
  function parseMarkValue(value){
    if (typeof value === 'number') return String(value);
    const num = Number(String(value).replace(/[^0-9.\-]/g, ''));
    return Number.isNaN(num) ? '' : String(num);
  }
  function parseMarkToNumber(value){
    if (value == null || value === '') return null;
    if (typeof value === 'number' && !Number.isNaN(value)) return value;
    const str = parseMarkValue(value);
    if (!str) return null;
    const num = Number(str);
    return Number.isNaN(num) ? null : num;
  }
  function formatMarkForComment(meta, rawValue, numericValue){
    if (meta){
      return meta.hasPercent ? `${meta.raw} %` : meta.raw;
    }
    if (rawValue != null && String(rawValue).trim()){
      return String(rawValue).trim();
    }
    if (numericValue != null){
      return `${Math.round(numericValue * 10) / 10}`;
    }
    return '';
  }
  function applyPronounPlaceholders(text, pronouns){
    if (!text) return '';
    if (!pronouns) return text;
    return text
      .replace(/\[he\/she\]/g, pronouns.he)
      .replace(/\[He\/She\]/g, pronouns.He)
      .replace(/\[his\/her\]/g, pronouns.his)
      .replace(/\[His\/Her\]/g, pronouns.His)
      .replace(/\[him\/her\]/g, pronouns.him)
      .replace(/\[Him\/Her\]/g, pronouns.Him)
      .replace(/\[himself\/herself\]/g, pronouns.him + 'self')
      .replace(/\[Himself\/Herself\]/g, pronouns.Him + 'self');
  }
  function normalizeGradeGroup(value){
    const v = String(value || '').toLowerCase();
    if (v.startsWith('elem')) return 'elem';
    if (v.startsWith('middle')) return 'middle';
    if (v.startsWith('high')) return 'high';
    return '';
  }
function buildGradeBasedComment(row, studentName, pronouns, gradeGroup, includeFinalGrade){
    if (!row) return '';
    const gradeColumn = commentConfig.gradeColumn || FINAL_GRADE_COLUMN;
    if (!gradeColumn) return '';
    const raw = row?.[gradeColumn];
    if (raw == null || raw === '') return '';
    const meta = deriveMarkMeta(raw, gradeColumn);
    const numeric = meta?.value ?? parseMarkToNumber(raw);
    if (numeric == null) return '';
    const high = Number(commentConfig.highThreshold ?? COMMENT_DEFAULTS.highThreshold);
    const mid = Number(commentConfig.midThreshold ?? COMMENT_DEFAULTS.midThreshold);
    let template = '';
    const groupKey = normalizeGradeGroup(gradeGroup);
    const groupTemplates = (typeof GRADE_GROUP_TEMPLATES !== 'undefined' && groupKey) ? GRADE_GROUP_TEMPLATES[groupKey] : null;
    if (groupTemplates){
      if (numeric >= high){
        template = groupTemplates.high;
      }else if (numeric >= mid){
        template = groupTemplates.mid;
      }else{
        template = groupTemplates.low;
      }
    }else if (numeric >= high){
      template = commentConfig.highTemplate || COMMENT_DEFAULTS.highTemplate;
    }else if (numeric >= mid){
      template = commentConfig.midTemplate || COMMENT_DEFAULTS.midTemplate;
    }else{
      template = commentConfig.lowTemplate || COMMENT_DEFAULTS.lowTemplate;
    }
    if (!template) return '';
    if (template.includes('approaches every lesson with curiosity')){
      template = COMMENT_DEFAULTS.highTemplate;
    }else if (template.includes('demonstrates steady progress and thoughtful engagement')){
      template = COMMENT_DEFAULTS.midTemplate;
    }else if (template.includes('developing foundational skills and currently holds')){
      template = COMMENT_DEFAULTS.lowTemplate;
    }
    const markText = formatMarkForComment(meta, raw, numeric);
    const nameText = studentName || buildStudentName(row) || 'This student';
    let result = applyPronounPlaceholders(
      template
        .replace(/\{name\}/gi, nameText)
        .replace(/\{mark\}/gi, markText),
      pronouns
    );
    if ((normalizeGradeGroup(gradeGroup) === 'elem') && includeFinalGrade && markText){
      result += ` Current overall mark is ${markText}.`;
    }
        if ((normalizeGradeGroup(gradeGroup) !== 'elem') && markText && !result.includes(markText)){
      result += ` Calculated final grade is ${markText}.`;
    }
    return result;
  }
  function autoSelectBuilderTemplate(row){
    if (!builderCorePerformanceSelect || !row) return;
    const gradeColumn = commentConfig.gradeColumn || FINAL_GRADE_COLUMN;
    if (!gradeColumn) return;
    const numeric = parseMarkToNumber(row[gradeColumn]);
    if (numeric == null) return;
    const high = Number(commentConfig.highThreshold ?? COMMENT_DEFAULTS.highThreshold);
    const mid = Number(commentConfig.midThreshold ?? COMMENT_DEFAULTS.midThreshold);
    const satisfactory = Number(commentConfig.satisfactoryThreshold ?? COMMENT_DEFAULTS.satisfactoryThreshold ?? (mid - 5));
    let templateValue = '';
    if (numeric >= high){
      templateValue = 'good1';
    }else if (numeric >= mid){
      templateValue = 'average1';
    }else if (numeric >= satisfactory){
      templateValue = 'satisfactory1';
    }else{
      templateValue = 'poor1';
    }
    if (!templateValue) return;
    builderCorePerformanceSelect.value = templateValue;
    builderCorePerformanceSelect.dispatchEvent(new Event('change', { bubbles:true }));
    updatePerformanceSelectStyle();
    builderGenerateReport();
    const baseComment = buildGradeBasedComment(row, builderStudentNameInput?.value.trim(), getBuilderPronouns(), builderGradeGroupSelect?.value, builderIncludeFinalGradeInput?.checked);
    const introLine = getPerformanceIntroLine(builderCorePerformanceSelect?.value || '', {
      studentName: builderStudentNameInput?.value.trim() || '',
      pronouns: getBuilderPronouns(),
      termAverage: builderTermAverageInput?.value.trim() || '',
      gradeGroup: builderGradeGroupSelect?.value || 'middle'
    });
    const baseWithIntro = introLine ? replaceFirstSentence(baseComment, introLine) : baseComment;
    if (baseWithIntro && builderReportOutput){
      setBuilderReportOutputText(polishGrammar(cleanFluency(baseWithIntro)), 'generate');
    }
  }
  function updatePerformanceSelectStyle(){
    if (!builderCorePerformanceSelect) return;
    const value = String(builderCorePerformanceSelect.value || '');
    builderCorePerformanceSelect.classList.remove('perf-good','perf-average','perf-satisfactory','perf-new','perf-poor');
    if (value.startsWith('good')){
      builderCorePerformanceSelect.classList.add('perf-good');
    }else if (value.startsWith('average')){
      builderCorePerformanceSelect.classList.add('perf-average');
    }else if (value.startsWith('satisfactory')){
      builderCorePerformanceSelect.classList.add('perf-satisfactory');
    }else if (value.startsWith('newstu')){
      builderCorePerformanceSelect.classList.add('perf-new');
    }else if (value.startsWith('poor')){
      builderCorePerformanceSelect.classList.add('perf-poor');
    }
  }
  function buildBuilderCommentBank(){
    if (!builderCommentBankEl) return;
    builderCommentBankEl.innerHTML = '';
    builderActiveCommentSection = null;
    const columnCount = 3;
    const columns = [];
    for (let i = 0; i < columnCount; i++){
      const col = document.createElement('div');
      col.className = 'builder-comment-column';
      builderCommentBankEl.appendChild(col);
      columns.push(col);
    }
    const collapseSection = (sectionEl) => {
      if (!sectionEl) return;
      const content = sectionEl.querySelector('.builder-comment-columns');
      sectionEl.classList.remove('expanded');
      sectionEl.classList.add('collapsed');
      if (content){
        content.style.maxHeight = '0px';
        content.style.opacity = '0';
        content.style.transform = 'translateY(-4px)';
      }
      const headerBtn = sectionEl.querySelector('.builder-comment-toggle');
      if (headerBtn) headerBtn.setAttribute('aria-expanded', 'false');
    };
    const expandSection = (sectionEl) => {
      if (!sectionEl) return;
      const content = sectionEl.querySelector('.builder-comment-columns');
      sectionEl.classList.add('expanded');
      sectionEl.classList.remove('collapsed');
      const headerBtn = sectionEl.querySelector('.builder-comment-toggle');
      if (headerBtn) headerBtn.setAttribute('aria-expanded', 'true');
      if (content){
        content.style.maxHeight = '0px';
        content.style.opacity = '0';
        content.style.transform = 'translateY(-4px)';
        requestAnimationFrame(() => {
          content.style.maxHeight = content.scrollHeight + 'px';
          content.style.opacity = '1';
          content.style.transform = 'translateY(0)';
        });
      }
    };

    REPORT_COMMENT_BANK.forEach((section, sectionIndex) => {
      const sectionEl = document.createElement('div');
      sectionEl.className = 'builder-comment-section collapsed';
      const headerBtn = document.createElement('button');
      headerBtn.type = 'button';
      headerBtn.className = 'builder-comment-toggle';
      headerBtn.setAttribute('aria-expanded', 'false');
      const title = document.createElement('span');
      title.className = 'builder-comment-title';
      title.textContent = section.title;
      const chevron = document.createElement('span');
      chevron.className = 'builder-comment-chevron';
      chevron.textContent = 'v';
      headerBtn.appendChild(title);
      headerBtn.appendChild(chevron);
      headerBtn.addEventListener('click', () => {
        if (builderActiveCommentSection && builderActiveCommentSection !== sectionEl){
          collapseSection(builderActiveCommentSection);
        }
        const isExpanded = sectionEl.classList.contains('expanded');
        if (isExpanded){
          collapseSection(sectionEl);
          builderActiveCommentSection = null;
          return;
        }
        const columnEl = sectionEl.parentElement;
        if (columnEl){
          columnEl.insertBefore(sectionEl, columnEl.firstChild);
        }
        expandSection(sectionEl);
        builderActiveCommentSection = sectionEl;
      });
      sectionEl.appendChild(headerBtn);

      const sortedOptions = [...section.options].sort((a, b) => {
        if (a.type === 'positive' && b.type !== 'positive') return -1;
        if (a.type !== 'positive' && b.type === 'positive') return 1;
        return a.label.localeCompare(b.label);
      });

      const columnsWrap = document.createElement('div');
      columnsWrap.className = 'builder-comment-columns';
      for (let i = 0; i < sortedOptions.length; i += 5){
        const col = document.createElement('div');
        col.className = 'builder-comment-col';
        sortedOptions.slice(i, i + 5).forEach(opt => {
          const wrapper = document.createElement('label');
          wrapper.className = 'builder-comment-item';
          if (opt.type === 'positive') {
            wrapper.classList.add('positive');
          } else if (opt.type === 'constructive') {
            wrapper.classList.add('constructive');
          }
          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.dataset.text = opt.text;
          checkbox.dataset.label = opt.label;
          checkbox.dataset.type = opt.type || '';
          checkbox.dataset.section = section.id || '';
          checkbox.dataset.id = opt.id || '';
          checkbox.id = `builder_${opt.id}`;
          const summary = document.createElement('span');
          summary.textContent = opt.label;
          wrapper.appendChild(checkbox);
          wrapper.appendChild(summary);
          col.appendChild(wrapper);
        });
        columnsWrap.appendChild(col);
      }
      sectionEl.appendChild(columnsWrap);
      // place section into 3-column layout
      const targetCol = columns[sectionIndex % columnCount];
      if (targetCol){
        targetCol.appendChild(sectionEl);
      }else{
        builderCommentBankEl.appendChild(sectionEl);
      }
      // ensure collapsed styles
      columnsWrap.style.maxHeight = '0px';
      columnsWrap.style.opacity = '0';
      columnsWrap.style.transform = 'translateY(-4px)';
    });

    builderCommentCheckboxes = Array.from(builderCommentBankEl.querySelectorAll('input[type="checkbox"]'));
    builderCommentBankEl.addEventListener('change', (e) => {
      const target = e.target;
      if (target && target.matches('input[type="checkbox"]')){
        const section = (target.dataset.section || '').toLowerCase();
        if (section === 'future' && !builderLookingAheadAutoUpdating){
          const key = getBuilderLookingAheadKey();
          if (key) {
            builderLookingAheadManualKey = key;
            builderLookingAheadAutoKey = key;
          }
          if (target.checked){
            builderCommentCheckboxes
              .filter(cb => cb !== target && (cb.dataset.section || '').toLowerCase() === 'future')
              .forEach(cb => {
                cb.checked = false;
                cb.dataset.auto = 'false';
              });
          }
          target.dataset.auto = 'false';
        }
      }
      updateBuilderSelectedTags();
      if (shouldAutoGenerateBuilderReport()){
        builderOutputAnimationMode = 'selection';
        builderGenerateReport();
      }
    });
  }

  function updateBuilderSelectedTags(){
    if (!builderSelectedTagsEl || !builderCommentCheckboxes.length) return;
    const selected = builderCommentCheckboxes.filter(cb => cb.checked);
    builderSelectedTagsEl.innerHTML = '';
    if (!selected.length){
      if (builderSelectedCommentsEl) builderSelectedCommentsEl.style.display = 'none';
      return;
    }
    if (builderSelectedCommentsEl) builderSelectedCommentsEl.style.display = 'block';
    selected.forEach(cb => {
      const tag = document.createElement('button');
      tag.type = 'button';
      tag.className = 'builder-tag';
      const type = (cb.dataset.type || '').toLowerCase();
      if (type === 'positive'){
        tag.classList.add('builder-tag-positive');
      }else if (type){
        tag.classList.add('builder-tag-constructive');
      }
      tag.setAttribute('aria-label', `Remove ${cb.dataset.label || 'selected comment'}`);
      const label = document.createElement('span');
      label.className = 'builder-tag-label';
      label.textContent = cb.dataset.label || 'Selected';
      const remove = document.createElement('span');
      remove.className = 'builder-tag-remove';
      remove.textContent = 'x';
      tag.appendChild(label);
      tag.appendChild(remove);
      tag.addEventListener('click', () => {
        cb.checked = false;
        cb.dataset.auto = 'false';
        updateBuilderSelectedTags();
        if (shouldAutoGenerateBuilderReport()){
          builderOutputAnimationMode = 'selection';
          builderGenerateReport();
        }
      });
      builderSelectedTagsEl.appendChild(tag);
    });
  }
  function updateCommentTermButtons(){
    const map = [
      [commentTermAllBtn, 'all'],
      [commentTermT1Btn, 'T1'],
      [commentTermT2Btn, 'T2'],
      [commentTermT3Btn, 'T3']
    ];
    map.forEach(([btn, term]) => {
      if (!btn) return;
      btn.classList.toggle('active', commentTermFilter === term);
      btn.setAttribute('aria-pressed', commentTermFilter === term ? 'true' : 'false');
    });
  }
  function setCommentTerm(term){
    commentTermFilter = term || 'all';
    updateCommentTermButtons();
  }
  function getBuilderPronouns(){
    if (builderPronounFemaleInput && builderPronounFemaleInput.checked){
      return { he:'she', his:'her', him:'her', He:'She', His:'Her', Him:'Her' };
    }
    return { he:'he', his:'his', him:'him', He:'He', His:'His', Him:'Him' };
  }
  function formatPercentValue(v){
    const s = String(v ?? '').trim();
    if (!s) return '';
    return s.includes('%') ? s : `${s}%`;
  }
  function builderReplacePlaceholders(text, context){
    if (!text) return '';
    const pronouns = context.pronouns;
    return text
      .replace(/\[Student\]/g, context.studentName)
      .replace(/\[he\/she\]/g, pronouns.he)
      .replace(/\[He\/She\]/g, pronouns.He)
      .replace(/\[his\/her\]/g, pronouns.his)
      .replace(/\[His\/Her\]/g, pronouns.His)
      .replace(/\[him\/her\]/g, pronouns.him)
      .replace(/\[Him\/Her\]/g, pronouns.Him)
      .replace(/\[himself\/herself\]/g, pronouns.him + 'self')
      .replace(/\[AVG%\]/g, context.termAverage ? formatPercentValue(context.termAverage) : '[AVG%]')
      .replace(/\[RETEST%\]/g, context.retestScore ? formatPercentValue(context.retestScore) : '[RETEST%]')
      .replace(/\[ORIG%\]/g, context.originalScore ? formatPercentValue(context.originalScore) : '[ORIG%]');
  }

  // ==== Assignment-aware builder helpers ====
  const ASSIGNMENT_BANDS = [
    { min: 90, key: 'high' },
    { min: 80, key: 'solid' },
    { min: 70, key: 'ok' },
    { min: 60, key: 'low' },
    { min: -Infinity, key: 'veryLow' }
  ];
  function hashString(input){
    let hash = 0;
    const str = String(input || '');
    for (let i = 0; i < str.length; i++){
      hash = ((hash << 5) - hash) + str.charCodeAt(i);
      hash |= 0;
    }
    return Math.abs(hash);
  }
  function pickVariant(variants, seed){
    if (Array.isArray(variants)){
      if (!variants.length) return '';
      const idx = hashString(seed) % variants.length;
      return variants[idx];
    }
    return variants || '';
  }
  function pickUniqueVariant(variants, seed, usedSet, scopeKey){
    const list = Array.isArray(variants) ? variants.filter(Boolean) : [variants].filter(Boolean);
    if (!list.length) return '';
    const start = hashString(seed) % list.length;
    for (let offset = 0; offset < list.length; offset += 1){
      const idx = (start + offset) % list.length;
      const id = `${scopeKey}:${idx}`;
      if (!usedSet || !usedSet.has(id)){
        usedSet?.add(id);
        return list[idx];
      }
    }
    return list[start];
  }
  const ASSIGNMENT_TEMPLATES = {
    default: {
      high: [
        "[Student] excelled on {label} with {score}%, showing strong command; [he/she] kept responses precise.",
        "[Student] showed clear mastery on {label}, earning {score}% with confident reasoning.",
        "On {label}, [Student] scored {score}%, demonstrating strong control and accuracy."
      ],
      solid: [
        "[Student] performed well on {label} at {score}%, with organized work and mostly solid reasoning.",
        "[Student] earned {score}% on {label}, showing steady understanding and good structure.",
        "{label} was handled well by [Student], who scored {score}% with mostly accurate steps."
      ],
      ok: [
        "[Student] met expectations on {label} with {score}%, though [he/she] should review the trickier steps.",
        "{label} landed at {score}% for [Student]; a brief review of key errors will help.",
        "[Student] reached {score}% on {label} and would benefit from revisiting the more complex items."
      ],
      low: [
        "[Student] struggled on {label} at {score}%, and would benefit from targeted practice on key outcomes.",
        "{label} was challenging for [Student] ({score}%); more guided practice is recommended.",
        "[Student] scored {score}% on {label}, indicating gaps that need focused support."
      ],
      veryLow: [
        "[Student] faced difficulty on {label} ({score}%), needing close support and guided review.",
        "{label} resulted in {score}% for [Student]; intensive review of fundamentals is needed.",
        "[Student] scored {score}% on {label}, and will benefit from step-by-step practice."
      ]
    },
    quiz: {
      high: [
        "[Student] was sharp on {label}, scoring {score}% and recalling concepts quickly.",
        "[Student] scored {score}% on {label}, showing quick recall and tidy work.",
        "{label} went very well for [Student], who scored {score}% with confident recall."
      ],
      solid: [
        "[Student] handled {label} at {score}%, showing steady recall and clear setups.",
        "{label} came in at {score}% for [Student]; recall was steady and work was organized.",
        "[Student] scored {score}% on {label}, showing dependable recall and structure."
      ],
      ok: [
        "[Student] managed {label} with {score}%; a quick refresh on core items will help.",
        "{label} landed at {score}% for [Student]; a brief review will help tighten recall.",
        "[Student] reached {score}% on {label}; reviewing key items will improve accuracy."
      ],
      low: [
        "[Student] found {label} challenging at {score}%; focused review of missed items is needed.",
        "{label} was difficult for [Student] ({score}%); targeted review will help.",
        "[Student] scored {score}% on {label}, indicating gaps in core recall."
      ],
      veryLow: [
        "[Student] struggled on {label} ({score}%), and needs guided practice on fundamentals.",
        "{label} resulted in {score}% for [Student]; foundational review is required.",
        "[Student] scored {score}% on {label}; close guidance on fundamentals is needed."
      ]
    },
    test: {
      high: [
        "[Student] excelled on {label} with {score}%, maintaining accuracy under time; [he/she] demonstrated mastery.",
        "[Student] earned {score}% on {label}, showing strong accuracy under time pressure.",
        "{label} was a strong performance at {score}%, highlighting solid mastery."
      ],
      solid: [
        "[Student] performed strongly on {label} at {score}%, with clear reasoning across sections.",
        "[Student] scored {score}% on {label}, showing clear steps and mostly accurate reasoning.",
        "{label} came in at {score}% with solid reasoning and organized work from [Student]."
      ],
      ok: [
        "[Student] reached {score}% on {label}; tightening multi-step work will boost results.",
        "{label} landed at {score}% for [Student]; refining multi-step work will help.",
        "[Student] scored {score}% on {label} and should focus on multi-step accuracy."
      ],
      low: [
        "[Student] scored {score}% on {label}; targeted re-practice on weak areas is recommended.",
        "{label} came in at {score}%; focused re-practice is recommended.",
        "[Student] reached {score}% on {label}, indicating specific gaps to target."
      ],
      veryLow: [
        "[Student] had difficulty on {label} ({score}%), and requires close support to rebuild skills.",
        "{label} resulted in {score}% for [Student]; close support is needed to rebuild skills.",
        "[Student] scored {score}% on {label} and will benefit from guided re-teaching."
      ]
    },
    project: {
      high: [
        "[Student] led {label} with {score}%, showing initiative and depth in the work.",
        "[Student] excelled on {label}, earning {score}% with thoughtful depth.",
        "{label} was a standout at {score}%, showing initiative and depth from [Student]."
      ],
      solid: [
        "[Student] produced solid work on {label} at {score}%, meeting expectations.",
        "{label} was completed well at {score}%, meeting expectations with clear effort.",
        "[Student] earned {score}% on {label}, showing solid execution."
      ],
      ok: [
        "[Student] completed {label} with {score}%; clearer organization will strengthen results.",
        "{label} came in at {score}% for [Student]; clearer organization will help.",
        "[Student] scored {score}% on {label} and would benefit from clearer structure."
      ],
      low: [
        "[Student] earned {score}% on {label}; more structure and checkpoints are needed.",
        "{label} landed at {score}% for [Student]; more structure is needed.",
        "[Student] reached {score}% on {label}, suggesting a need for clearer checkpoints."
      ],
      veryLow: [
        "[Student] struggled to complete {label} ({score}%), needing more guidance.",
        "{label} resulted in {score}% for [Student]; more guidance is needed.",
        "[Student] scored {score}% on {label} and needs more structured guidance."
      ]
    },
    homework: {
      high: [
        "[Student] completed {label} at {score}%, reflecting consistent practice.",
        "[Student] scored {score}% on {label}, reflecting consistent practice.",
        "{label} was strong at {score}%, showing steady practice."
      ],
      solid: [
        "[Student] maintained {score}% on {label}, keeping pace with assignments.",
        "[Student] held {score}% on {label}, keeping steady pace with practice.",
        "{label} came in at {score}%, showing consistent practice habits."
      ],
      ok: [
        "[Student] logged {score}% on {label}; regular review would strengthen recall.",
        "{label} came in at {score}% for [Student]; more consistent review will help.",
        "[Student] earned {score}% on {label}; regular review will help."
      ],
      low: [
        "[Student] reached {score}% on {label}; nightly practice will help solidify skills.",
        "{label} landed at {score}% for [Student]; more regular practice will help.",
        "[Student] scored {score}% on {label}, suggesting practice gaps."
      ],
      veryLow: [
        "[Student] often missed {label} ({score}%), needing routine support to build habits.",
        "{label} resulted in {score}% for [Student]; consistent routines are needed.",
        "[Student] scored {score}% on {label}, and needs structured practice habits."
      ]
    }
  };
  function detectAssignmentType(label = ''){
    const t = String(label || '').toLowerCase();
    if (t.includes('quiz')) return 'quiz';
    if (t.includes('test') || t.includes('exam')) return 'test';
    if (t.includes('project')) return 'project';
    if (t.includes('hw') || t.includes('homework') || t.includes('assignment')) return 'homework';
    return 'default';
  }
  function cleanAssignmentLabel(label = ''){
    const m = String(label || '').match(/^L\d+\s*-\s*(.*)$/i);
    return m ? m[1].trim() || String(label || '').trim() : String(label || '');
  }
  function collectAssignmentFacts(selectedAssignments, row){
    if (!Array.isArray(selectedAssignments) || !selectedAssignments.length || !row) return [];
    const facts = [];
    selectedAssignments.slice(0, 2).forEach((col) => {
      const meta = deriveMarkMeta(row[col], col);
      if (!meta) return;
      facts.push({
        label: cleanAssignmentLabel(col),
        score: meta.value,
        scoreText: meta.raw
      });
    });
    return facts;
  }
  function buildAssignmentSentences(selectedAssignments, row, context, overallMeta){
    if (!selectedAssignments.length || !row) return '';
    const overallValue = overallMeta?.value;
    const overallDisplay = overallMeta?.raw ? formatPercentValue(overallMeta.raw) : '';
    const used = new Set();
    const entries = collectAssignmentFacts(selectedAssignments, row);
    if (!entries.length) return '';
    const leadVariants = [
      "Assignment results this term show a mix of strengths across recent tasks.",
      "Recent assignment evidence highlights where [Student] is performing strongly and where refinement is needed.",
      "Across recent assignments, [Student] is showing clear growth with a few targeted focus points."
    ];
    const clauseOpenersA = [
      "On {label}, [Student] scored {score}%",
      "For {label}, [Student] recorded {score}%",
      "In {label}, [Student] earned {score}%",
      "Across {label}, [Student] posted {score}%"
    ];
    const clauseOpenersB = [
      "while on {label}, [he/she] scored {score}%",
      "whereas in {label}, [he/she] recorded {score}%",
      "and on {label}, [he/she] earned {score}%",
      "while in {label}, [he/she] posted {score}%"
    ];
    const aboveNotes = [
      `which sits above the overall ${overallDisplay} and can be leveraged across other tasks`,
      `which is stronger than the overall ${overallDisplay} and highlights an area of confidence`,
      `which outperforms the overall ${overallDisplay} and signals a strength to build from`
    ];
    const belowNotes = [
      `which is below the overall ${overallDisplay} and identifies a targeted area to reinforce`,
      `which trails the overall ${overallDisplay} and would benefit from focused review`,
      `which is weaker than the overall ${overallDisplay} and should be tightened with deliberate practice`
    ];
    const alignedNotes = [
      `which aligns closely with the overall ${overallDisplay}`,
      `which is broadly in line with the overall ${overallDisplay}`,
      `which reflects a level similar to the overall ${overallDisplay}`
    ];
    const closeVariants = [
      "This pattern gives us a clear next step for instruction and practice.",
      "Together, these results outline both a dependable strength and a concrete growth target.",
      "This spread of results gives a practical roadmap for next-step support."
    ];
    const chooseRelation = (entry, idx) => {
      if (overallValue == null || !overallDisplay) return '';
      const diff = entry.score - overallValue;
      if (diff >= 8){
        return pickUniqueVariant(aboveNotes, `${entry.label}|above|${idx}`, used, 'rel-above');
      }
      if (diff <= -8){
        return pickUniqueVariant(belowNotes, `${entry.label}|below|${idx}`, used, 'rel-below');
      }
      return pickUniqueVariant(alignedNotes, `${entry.label}|aligned|${idx}`, used, 'rel-aligned');
    };
    const clauseAOpen = pickUniqueVariant(clauseOpenersA, `${entries[0].label}|openA`, used, 'openA');
    let clauseA = clauseAOpen.replace('{label}', entries[0].label).replace('{score}', entries[0].scoreText);
    const relationA = chooseRelation(entries[0], 0);
    clauseA += relationA ? `, ${relationA}.` : '.';
    clauseA = builderReplacePlaceholders(clauseA, context);
    if (entries.length === 1){
      const lead = builderReplacePlaceholders(pickUniqueVariant(leadVariants, `${entries[0].label}|lead`, used, 'lead'), context);
      const close = builderReplacePlaceholders(pickUniqueVariant(closeVariants, `${entries[0].label}|close`, used, 'close'), context);
      return `${lead} ${clauseA} ${close}`.trim();
    }
    const clauseBOpen = pickUniqueVariant(clauseOpenersB, `${entries[1].label}|openB`, used, 'openB');
    let clauseB = clauseBOpen.replace('{label}', entries[1].label).replace('{score}', entries[1].scoreText);
    const relationB = chooseRelation(entries[1], 1);
    clauseB += relationB ? `, ${relationB}.` : '.';
    clauseB = builderReplacePlaceholders(clauseB, context);
    const lead = builderReplacePlaceholders(pickUniqueVariant(leadVariants, `${entries[0].label}|${entries[1].label}|lead`, used, 'lead'), context);
    const close = builderReplacePlaceholders(pickUniqueVariant(closeVariants, `${entries[0].label}|${entries[1].label}|close`, used, 'close'), context);
    return `${lead} ${clauseA} ${clauseB} ${close}`.trim();
  }
  function builderGetScoreCommentary(score, testNumber, level){
    if (!score && score !== 0) return '';
    const num = Number(score);
    const testName = testNumber === 1 ? 'Test 1' : 'Test 2';
    if (level.startsWith('good')){
      if (num === 100) return `[He/She] achieved a perfect score of 100% on ${testName}, demonstrating complete mastery of the material.`;
      if (num >= 90) return `[He/She] achieved an impressive ${num}% on ${testName}, demonstrating strong mastery.`;
      if (num >= 80) return `[He/She] achieved a solid ${num}% on ${testName}, showing good understanding.`;
      if (num >= 70) return `[He/She] achieved ${num}% on ${testName}, which is respectable but below [his/her] usual standards.`;
      if (num >= 60) return `[He/She] scored ${num}% on ${testName}, which was below expectations given [his/her] capabilities.`;
      if (num >= 55) return `[He/She] scored only ${num}% on ${testName}, which is concerning given [his/her] potential.`;
      if (num >= 50) return `[He/She] scored ${num}% on ${testName}, which is at the passing threshold and well below [his/her] potential.`;
      return `[He/She] unfortunately failed with ${num}% on ${testName}, despite generally strong abilities elsewhere.`;
    }
    if (level.startsWith('average')){
      if (num === 100) return `[He/She] achieved an exceptional 100% on ${testName}, demonstrating [his/her] true potential.`;
      if (num >= 90) return `[He/She] achieved a strong ${num}% on ${testName}, showing what [he/she] is capable of.`;
      if (num >= 80) return `[He/She] achieved ${num}% on ${testName}, which was a positive result.`;
      if (num >= 70) return `[He/She] achieved ${num}% on ${testName}, demonstrating adequate understanding.`;
      if (num >= 60) return `[He/She] scored ${num}% on ${testName}, indicating some gaps in understanding.`;
      if (num >= 55) return `[He/She] scored ${num}% on ${testName}, which is at the lower threshold and requires attention.`;
      if (num >= 50) return `[He/She] achieved ${num}% on ${testName}, just meeting the passing standard and indicating preparation gaps.`;
      return `[He/She] failed with ${num}% on ${testName}, demonstrating inadequate preparation.`;
    }
    if (level.startsWith('poor')){
      if (num === 100) return `[He/She] surprisingly achieved 100% on ${testName}, demonstrating untapped potential.`;
      if (num >= 90) return `[He/She] achieved ${num}% on ${testName}, showing [he/she] is capable of much more.`;
      if (num >= 80) return `[He/She] achieved ${num}% on ${testName}, which shows capability but highlights inconsistency.`;
      if (num >= 70) return `[He/She] achieved ${num}% on ${testName}, demonstrating basic understanding.`;
      if (num >= 60) return `[He/She] scored ${num}% on ${testName}, showing continued struggles.`;
      if (num >= 55) return `[He/She] scored ${num}% on ${testName}, which is at the lower threshold and indicates significant gaps.`;
      if (num >= 50) return `[He/She] scored ${num}% on ${testName}, just meeting the passing standard, which indicates serious preparation concerns.`;
      return `[He/She] failed with ${num}% on ${testName}, highlighting serious concerns about preparation.`;
    }
    return '';
  }
  function builderGetCCCommentary(level){
    if (!level) return '';
    return REPORT_BUILDER_CC[level] || '';
  }
  function getBuilderSelectedComments(){
    return builderCommentCheckboxes
      .filter(cb => cb.checked)
      .map(cb => ({
        id: cb.dataset.id || '',
        text: cb.dataset.text || '',
        type: (cb.dataset.type || '').toLowerCase(),
        section: (cb.dataset.section || '').toLowerCase()
      }));
  }
  function prioritizePersonalQualities(items){
    if (!Array.isArray(items) || !items.length) return [];
    const pq = items.filter(it => it.section === 'personalqualities');
    const rest = items.filter(it => it.section !== 'personalqualities');
    return pq.concat(rest);
  }
  function orderCommentsSandwichStyle(items){
    if (!Array.isArray(items) || !items.length) return [];
    const positives = items.filter(it => it.type === 'positive');
    const constructives = items.filter(it => it.type !== 'positive');
    const ordered = [];
    if (positives.length){
      ordered.push(positives[0]);
    }
    ordered.push(...constructives);
    if (positives.length > 1){
      ordered.push(...positives.slice(1));
    }
    return ordered.length ? ordered : items;
  }
  function buildCommentSnippetText(selectedComments, context){
    if (!Array.isArray(selectedComments) || !selectedComments.length) return '';
    const replace = (item) => builderReplacePlaceholders(item.text, context);
    if (commentOrderMode === 'sandwich'){
      const prioritized = prioritizePersonalQualities(selectedComments);
      const pqFirst = prioritized.filter(it => it.section === 'personalqualities');
      const others = prioritized.filter(it => it.section !== 'personalqualities');
      const ordered = pqFirst.concat(orderCommentsSandwichStyle(others));
      return ordered.map(replace).join(' ');
    }
    if (commentOrderMode === 'bullet'){
      const nameRegex = new RegExp(`\\b${escapeRegExp(context.studentName)}\\b`, 'gi');
      const positives = selectedComments.filter(it => it.type === 'positive').map(replace);
      const constructives = selectedComments.filter(it => it.type !== 'positive').map(replace);
      const bulletify = (items) => items.map((txt) => `- ${txt.replace(nameRegex, context.pronouns.He)}`);
      const blocks = [];
      if (positives.length){
        blocks.push(['Strengths:', ...bulletify(positives)].join('\n'));
      }
      if (constructives.length){
        blocks.push(['Growth focus:', ...bulletify(constructives)].join('\n'));
      }
      return blocks.join('\n\n');
    }
    return selectedComments.map(replace).join(' ');
  }


  function insertSentencesAfterFirst(text, sentencesToInsert, maxCount){
    if (!text) return text || '';
    const sentences = splitIntoSentences(text);
    if (!sentences.length) return text;
    const insert = Array.isArray(sentencesToInsert) ? sentencesToInsert.filter(Boolean) : [];
    if (!insert.length) return text;
    const count = Math.max(1, Math.min(maxCount || insert.length, insert.length));
    sentences.splice(1, 0, ...insert.slice(0, count));
    return sentences.join(' ');
  }
  function ensureGradeSentenceNearTop(text, markText, gradeSentence){
    if (!text || !markText || !gradeSentence) return text || '';
    const sentences = splitIntoSentences(text);
    if (!sentences.length) return text;
    const topTwo = sentences.slice(0, 2).join(' ');
    if (topTwo.includes(markText) || topTwo.includes(gradeSentence)) return text;
    sentences.splice(1, 0, gradeSentence);
    return sentences.join(' ');
  }

function splitIntoSentences(text){
    if (!text) return [];
    return String(text).split(/(?<=[.!?])\s+/).filter(Boolean);
  }
  function normalizeSentenceForDedupe(sentence, context){
    let text = String(sentence || '').toLowerCase();
    const studentName = String(context?.studentName || '').trim();
    if (studentName){
      text = text.replace(new RegExp(`\\b${escapeRegExp(studentName)}\\b`, 'gi'), ' student ');
    }
    text = text
      .replace(/\b(he|she|his|her|him|they|their|them)\b/g, ' pronoun ')
      .replace(/\d+(?:\.\d+)?\s*%/g, ' percent ')
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\b(the|a|an|and|or|but|to|of|in|on|at|with|for|from|this|that|is|are|was|were|be|been|being)\b/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    return text;
  }
  function sentenceSimilarity(a, b, context){
    const aa = normalizeSentenceForDedupe(a, context).split(' ').filter(Boolean);
    const bb = normalizeSentenceForDedupe(b, context).split(' ').filter(Boolean);
    if (!aa.length || !bb.length) return 0;
    const setA = new Set(aa);
    const setB = new Set(bb);
    let intersect = 0;
    setA.forEach(token => {
      if (setB.has(token)) intersect += 1;
    });
    const union = new Set([...setA, ...setB]).size || 1;
    return intersect / union;
  }
  function detectSentenceIntentBucket(sentence){
    const s = String(sentence || '').toLowerCase();
    if (/(next term|coming term|year ahead|look forward|successful year)/.test(s)) return 'future';
    if (/(assignment|test|quiz|exam|challenge|review|topic|\d+(?:\.\d+)?\s*%)/.test(s)) return 'assignment';
    if (/(overall|\babove\b|\bbelow\b|outperform|stronger than|weaker than|align)/.test(s)) return 'comparison';
    if (/(time management|starting .* early|on time|routine|consistent review)/.test(s)) return 'time-management';
    if (/(attendance|engaged during class|morale|demeanor|attitude)/.test(s)) return 'classroom';
    if (/(steady progress|meeting expectations|progressing|growth)/.test(s)) return 'progress';
    return 'general';
  }
  function sentenceInfoScore(sentence, requirements){
    const text = String(sentence || '');
    let score = text.length;
    if (/\d+(?:\.\d+)?\s*%/.test(text)) score += 25;
    const labels = Array.isArray(requirements?.assignmentLabels) ? requirements.assignmentLabels : [];
    labels.forEach(label => {
      if (label && text.toLowerCase().includes(String(label).toLowerCase())) score += 30;
    });
    if (requirements?.overallMark && text.includes(requirements.overallMark)) score += 18;
    if (detectSentenceIntentBucket(text) === 'future') score += 8;
    return score;
  }
  function dedupeGeneratedComment(report, context, requirements = {}){
    const source = String(report || '').trim();
    if (!source) return '';
    if (commentOrderMode === 'bullet') return source;
    const sentences = splitIntoSentences(source);
    const kept = [];
    sentences.forEach(sentence => {
      const current = sentence.trim();
      if (!current) return;
      const bucket = detectSentenceIntentBucket(current);
      let duplicateIndex = -1;
      let duplicateStrength = 0;
      for (let i = 0; i < kept.length; i += 1){
        const existing = kept[i];
        const sim = sentenceSimilarity(current, existing.text, context);
        const sameBucket = bucket !== 'general' && existing.bucket === bucket;
        const exactKey = normalizeSentenceForDedupe(current, context) === normalizeSentenceForDedupe(existing.text, context);
        if (exactKey || sim >= 0.78 || (sameBucket && sim >= 0.58)){
          duplicateIndex = i;
          duplicateStrength = sim;
          break;
        }
      }
      if (duplicateIndex === -1){
        kept.push({ text: current, bucket });
        return;
      }
      const existing = kept[duplicateIndex];
      const keepCurrent = sentenceInfoScore(current, requirements) > sentenceInfoScore(existing.text, requirements);
      if (keepCurrent || duplicateStrength >= 0.9){
        kept[duplicateIndex] = { text: current, bucket };
      }
    });
    return kept.map(item => item.text).join(' ');
  }
  function ensureRequiredFactsInReport(report, requirements = {}){
    const result = String(report || '').trim();
    if (!result || commentOrderMode === 'bullet') return result;
    const sourceSentences = splitIntoSentences(requirements.sourceText || '');
    const outputSentences = splitIntoSentences(result);
    const hasText = (needle) => String(outputSentences.join(' ')).toLowerCase().includes(String(needle || '').toLowerCase());
    const appendFromSource = (matchFn) => {
      const hit = sourceSentences.find(matchFn);
      if (hit && !hasText(hit)){
        outputSentences.push(hit.trim());
      }
    };
    const labels = Array.isArray(requirements.assignmentLabels) ? requirements.assignmentLabels : [];
    labels.forEach(label => {
      if (!label || hasText(label)) return;
      appendFromSource(sentence => sentence.toLowerCase().includes(String(label).toLowerCase()));
    });
    const scores = Array.isArray(requirements.assignmentScores) ? requirements.assignmentScores : [];
    scores.forEach(score => {
      if (!score || hasText(score)) return;
      appendFromSource(sentence => sentence.includes(score));
    });
    if (requirements.overallMark && !hasText(requirements.overallMark)){
      appendFromSource(sentence => sentence.includes(requirements.overallMark));
    }
    if (requirements.requireFutureSentence){
      const hasFuture = outputSentences.some(s => detectSentenceIntentBucket(s) === 'future');
      if (!hasFuture && requirements.futureSentence){
        outputSentences.push(String(requirements.futureSentence).trim());
      }
    }
    return outputSentences.filter(Boolean).join(' ');
  }
  function buildGenerationRequirements({
    sourceText,
    selectedAssignments,
    studentRow,
    overallMeta,
    lookingAheadText
  }){
    const assignmentFacts = collectAssignmentFacts(selectedAssignments, studentRow);
    return {
      sourceText: String(sourceText || ''),
      assignmentLabels: assignmentFacts.map(item => item.label).filter(Boolean),
      assignmentScores: assignmentFacts.map(item => item.scoreText).filter(Boolean),
      overallMark: overallMeta?.raw ? formatPercentValue(overallMeta.raw) : '',
      requireFutureSentence: Boolean(String(lookingAheadText || '').trim()),
      futureSentence: String(lookingAheadText || '').trim()
    };
  }
  function buildUnifiedComment({ baseComment, termLabel, toneLine, commentSnippet, assignmentParagraph, lookingAheadText }){
    let baseText = baseComment || '';
    if (termLabel){
      baseText = `${termLabel}: ${baseText}`;
    }
    let closing = '';
    let core = baseText;
    if (baseText){
      const sentences = splitIntoSentences(baseText);
      if (sentences.length > 1){
        closing = sentences.pop();
        core = sentences.join(' ');
      }
    }
    const blocks = [];
    const appendUniqueBlock = (text) => {
      const blockText = String(text || '').trim();
      if (!blockText) return;
      const existing = splitIntoSentences(blocks.join(' '));
      const incoming = splitIntoSentences(blockText);
      const filtered = incoming.filter(sentence => {
        const bucket = detectSentenceIntentBucket(sentence);
        if (bucket === 'future' && existing.some(s => detectSentenceIntentBucket(s) === 'future')) return false;
        return !existing.some(s => sentenceSimilarity(sentence, s, { studentName: '' }) >= 0.74);
      });
      if (!filtered.length) return;
      blocks.push(filtered.join(' '));
    };
    appendUniqueBlock(core);
    appendUniqueBlock(toneLine);
    appendUniqueBlock(commentSnippet);
    appendUniqueBlock(assignmentParagraph);
    appendUniqueBlock(lookingAheadText);
    appendUniqueBlock(closing);
    return blocks.filter(Boolean).join('\n\n');
  }

  function getBuilderLookingAheadKey(){
    const studentKey = builderSelectedRowIndex != null
      ? `row:${builderSelectedRowIndex}`
      : `name:${builderStudentNameInput?.value.trim() || ''}`;
    const perfKey = builderCorePerformanceSelect?.value || '';
    return `${studentKey}::${perfKey}`;
  }
  function applyRandomLookingAheadSelection(force = false){
    if (!builderCommentBankEl) return null;
    if (!builderCommentCheckboxes.length){
      builderCommentCheckboxes = Array.from(builderCommentBankEl.querySelectorAll('input[type="checkbox"]'));
    }
    const lookingCbs = builderCommentCheckboxes.filter(cb => (cb.dataset.section || '').toLowerCase() === 'future');
    if (!lookingCbs.length) return null;
    const key = getBuilderLookingAheadKey();
    if (!force){
      if (key && key === builderLookingAheadAutoKey) return null;
      if (key && key === builderLookingAheadManualKey) return null;
    }else if (key && builderLookingAheadManualKey === key){
      builderLookingAheadManualKey = '';
    }
    builderLookingAheadAutoUpdating = true;
    // clear all looking ahead selections, then pick one
    lookingCbs.forEach(cb => {
      cb.checked = false;
      cb.dataset.auto = 'false';
    });
    const positiveLooking = lookingCbs.filter(cb => (cb.dataset.type || '').toLowerCase() === 'positive');
    const pool = positiveLooking.length ? positiveLooking : lookingCbs;
    const pick = pool[Math.floor(Math.random() * pool.length)];
    if (pick){
      pick.checked = true;
      pick.dataset.auto = 'true';
    }
    builderLookingAheadAutoUpdating = false;
    builderLookingAheadAutoKey = key;
    updateBuilderSelectedTags();
    return pick ? pick.dataset.text || '' : '';
  }

function buildCommentBlocks(selectedComments, context){
    if (!Array.isArray(selectedComments) || !selectedComments.length){
      return { pqSentences: [], lookingAheadSentences: [], lookingAheadPositiveSentences: [], otherSentences: [], otherText: '' };
    }
    if (commentOrderMode === 'bullet'){
      return { pqSentences: [], lookingAheadSentences: [], lookingAheadPositiveSentences: [], otherSentences: [], otherText: buildCommentSnippetText(selectedComments, context) };
    }
    const replace = (item) => builderReplacePlaceholders(item.text, context);
    let ordered = selectedComments;
    if (commentOrderMode === 'sandwich'){
      const prioritized = prioritizePersonalQualities(selectedComments);
      const pqFirst = prioritized.filter(it => it.section === 'personalqualities');
      const others = prioritized.filter(it => it.section !== 'personalqualities');
      ordered = pqFirst.concat(orderCommentsSandwichStyle(others));
    }
    const pqSentences = [];
    const lookingAheadSentences = [];
    const lookingAheadPositiveSentences = [];
    const otherSentences = [];
    ordered.forEach(item => {
      const txt = replace(item);
      if (item.section === 'personalqualities') pqSentences.push(txt);
      else if (item.section === 'future'){
        lookingAheadSentences.push(txt);
        if ((item.type || '').toLowerCase() === 'positive'){
          lookingAheadPositiveSentences.push(txt);
        }
      }
      else otherSentences.push(txt);
    });
    return { pqSentences, lookingAheadSentences, lookingAheadPositiveSentences, otherSentences, otherText: otherSentences.join(' ') };
  }

function cleanFluency(text){
    if (!text) return '';
    const lines = String(text).split(/\n+/);
    const cleanedLines = lines.map(line => {
      const trimmed = line.trim();
      if (!trimmed) return '';
      const sentences = trimmed.split(/(?<=[.!?])\s+/);
      const seen = new Set();
      const uniq = [];
      sentences.forEach(s => {
        const key = s.trim().toLowerCase();
        if (!key || seen.has(key)) return;
        seen.add(key);
        uniq.push(s.trim());
      });
      return uniq.join(' ');
    }).filter(Boolean);
    return cleanedLines.join('\n\n').trim();
  }
  function polishGrammar(text){
    if (!text) return '';
    let t = String(text);
    t = t.replace(/\s+([,.;:!])/g, '$1');
    t = t.replace(/[ \t]{2,}/g, ' ');
    t = t.replace(/ \n/g, '\n');
    t = t.replace(/\b(a)\s+([aeiou])/gi, (_, a, v) => (a === 'A' ? 'An ' : 'an ') + v);
    t = t.replace(/\b(an)\s+([^aeiou\W])/gi, (_, an, c) => (an === 'An' ? 'A ' : 'a ') + c);
    return t.trim();
  }
  function applyPronounsAfterFirstSentence(text, name, pronouns){
    if (!text || !name) return text || '';
    const safeName = escapeRegExp(name.trim());
    if (!safeName) return text;
    const sentences = String(text).split(/(?<=[.!?])\s+/);
    return sentences.map((s, idx) => {
      if (idx === 0) return s;
      const re = new RegExp(`^\\s*${safeName}(\\b|'s\\b)`, 'i');
      if (!re.test(s)) return s;
      return s.replace(re, (_, possessive) => possessive ? pronouns.His : pronouns.He);
    }).join(' ');
  }
  function clearBuilderOutputAnimation(){
    builderOutputAnimationToken += 1;
    builderOutputAnimating = false;
    builderReportOutput?.classList.remove('builder-output-animating');
  }
  function escapeHtml(text){
    return String(text || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
  function splitWords(text){
    const parts = String(text || '').match(/(\s+|[^\s]+)/g) || [];
    return parts;
  }
  function buildWordDiffMap(prevText, nextText){
    const prevWords = String(prevText || '').trim().split(/\s+/).filter(Boolean).map(w => w.toLowerCase());
    const nextWords = String(nextText || '').trim().split(/\s+/).filter(Boolean);
    const marks = new Array(nextWords.length).fill(false);
    if (!nextWords.length) return marks;
    if (!prevWords.length){
      marks.fill(true);
      return marks;
    }
    const m = prevWords.length;
    const n = nextWords.length;
    const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
    for (let i = 1; i <= m; i += 1){
      const a = prevWords[i - 1];
      for (let j = 1; j <= n; j += 1){
        const b = String(nextWords[j - 1] || '').toLowerCase();
        dp[i][j] = a === b ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
      }
    }
    let i = m;
    let j = n;
    while (i > 0 && j > 0){
      const a = prevWords[i - 1];
      const b = String(nextWords[j - 1] || '').toLowerCase();
      if (a === b){
        marks[j - 1] = false;
        i -= 1;
        j -= 1;
      }else if (dp[i - 1][j] >= dp[i][j - 1]){
        i -= 1;
      }else{
        marks[j - 1] = true;
        j -= 1;
      }
    }
    while (j > 0){
      marks[j - 1] = true;
      j -= 1;
    }
    return marks;
  }
  function renderBuilderDiffOverlay(prevText, nextText){
    if (!builderOutputWrap || !builderReportOverlay || !builderReportOutput) return;
    if (builderDiffClearTimer){
      clearTimeout(builderDiffClearTimer);
      builderDiffClearTimer = null;
    }
    const tokens = splitWords(nextText);
    const marks = buildWordDiffMap(prevText, nextText);
    let wordIdx = 0;
    const html = tokens.map(token => {
      if (/^\s+$/.test(token)){
        return escapeHtml(token);
      }
      const changed = !!marks[wordIdx];
      wordIdx += 1;
      const safe = escapeHtml(token);
      return changed ? `<span class="revise-diff">${safe}</span>` : safe;
    }).join('');
    builderReportOverlay.innerHTML = html;
    builderReportOverlay.scrollTop = builderReportOutput.scrollTop;
    builderReportOverlay.scrollLeft = builderReportOutput.scrollLeft;
    builderOutputWrap.classList.add('overlay-active');
    builderDiffClearTimer = setTimeout(() => {
      builderOutputWrap?.classList.remove('overlay-active');
      if (builderReportOverlay) builderReportOverlay.innerHTML = '';
      builderDiffClearTimer = null;
    }, 2000);
  }
  function clearBuilderDiffOverlay(){
    if (builderDiffClearTimer){
      clearTimeout(builderDiffClearTimer);
      builderDiffClearTimer = null;
    }
    builderOutputWrap?.classList.remove('overlay-active');
    if (builderReportOverlay) builderReportOverlay.innerHTML = '';
  }
  function pulseBuilderOutput(className = 'builder-output-pulse'){
    if (!builderReportOutput) return;
    builderReportOutput.classList.remove('builder-output-pulse');
    void builderReportOutput.offsetWidth;
    builderReportOutput.classList.add(className);
  }
  function countChangedSentences(prev, next){
    const toSentences = (txt) => String(txt || '')
      .split(/(?<=[.!?])\s+/)
      .map(s => s.trim())
      .filter(Boolean);
    const a = toSentences(prev);
    const b = toSentences(next);
    const total = Math.max(a.length, b.length);
    if (!total) return 0;
    let changed = 0;
    for (let i = 0; i < total; i += 1){
      if ((a[i] || '') !== (b[i] || '')) changed += 1;
    }
    return changed;
  }
  function getCommonPrefixLength(a, b){
    const left = String(a || '');
    const right = String(b || '');
    const max = Math.min(left.length, right.length);
    let i = 0;
    while (i < max && left[i] === right[i]) i += 1;
    return i;
  }
  function getCommonSuffixLength(a, b, prefixLen = 0){
    const left = String(a || '');
    const right = String(b || '');
    const max = Math.min(left.length - prefixLen, right.length - prefixLen);
    let i = 0;
    while (i < max && left[left.length - 1 - i] === right[right.length - 1 - i]) i += 1;
    return i;
  }
  function getSelectionChangeMeta(prev, next){
    const left = String(prev || '');
    const right = String(next || '');
    const prefixLen = getCommonPrefixLength(left, right);
    const suffixLen = getCommonSuffixLength(left, right, prefixLen);
    const prevMiddle = left.slice(prefixLen, left.length - suffixLen);
    const nextMiddle = right.slice(prefixLen, right.length - suffixLen);
    return { prefixLen, suffixLen, prevMiddle, nextMiddle };
  }
  function getFirstDiffIndex(a, b){
    const left = String(a || '');
    const right = String(b || '');
    const minLen = Math.min(left.length, right.length);
    for (let i = 0; i < minLen; i += 1){
      if (left[i] !== right[i]) return i;
    }
    if (left.length !== right.length) return minLen;
    return -1;
  }
  function getSelectionAnimationStart(prev, next){
    const target = String(next || '');
    const diff = getFirstDiffIndex(prev, target);
    if (diff <= 0) return 0;
    let start = 0;
    const boundaries = ['\n\n', '. ', '! ', '? ', '\n'];
    boundaries.forEach(marker => {
      const idx = target.lastIndexOf(marker, diff - 1);
      if (idx >= 0){
        start = Math.max(start, idx + marker.length);
      }
    });
    if (start === 0){
      const wordIdx = target.lastIndexOf(' ', diff - 1);
      if (wordIdx >= 0) start = wordIdx + 1;
    }
    return Math.min(start, target.length);
  }
  function animateBuilderOutputWords(prevText, nextText){
    if (!builderReportOutput) return;
    clearBuilderOutputAnimation();
    builderOutputAnimating = true;
    builderReportOutput.classList.add('builder-output-animating');
    const token = builderOutputAnimationToken;
    const next = String(nextText || '');
    const start = getSelectionAnimationStart(prevText, next);
    const prefix = next.slice(0, start);
    const suffix = next.slice(start);
    const parts = suffix.split(/(\s+)/);
    let i = 0;
    builderReportOutput.value = prefix;
    const step = () => {
      if (token !== builderOutputAnimationToken || !builderReportOutput) return;
      const end = Math.min(i + 2, parts.length);
      for (; i < end; i += 1){
        builderReportOutput.value = prefix + parts.slice(0, i + 1).join('');
      }
      if (i < parts.length){
        setTimeout(step, 16);
      }else{
        builderReportOutput.value = next;
        builderOutputAnimating = false;
        builderReportOutput.classList.remove('builder-output-animating');
      }
    };
    step();
  }
  function animateBuilderOutputRemoval(prevText, nextText){
    if (!builderReportOutput) return;
    clearBuilderOutputAnimation();
    builderOutputAnimating = true;
    builderReportOutput.classList.add('builder-output-animating');
    const token = builderOutputAnimationToken;
    const prev = String(prevText || '');
    const next = String(nextText || '');
    const change = getSelectionChangeMeta(prev, next);
    const prefix = prev.slice(0, change.prefixLen);
    const suffix = prev.slice(prev.length - change.suffixLen);
    const removed = change.prevMiddle;
    if (!removed){
      builderReportOutput.value = next;
      builderOutputAnimating = false;
      builderReportOutput.classList.remove('builder-output-animating');
      return;
    }
    let keepLen = removed.length;
    builderReportOutput.value = prev;
    const step = () => {
      if (token !== builderOutputAnimationToken || !builderReportOutput) return;
      builderReportOutput.value = prefix + removed.slice(0, Math.max(keepLen, 0)) + suffix;
      keepLen -= 2;
      if (keepLen >= 0){
        setTimeout(step, 14);
      }else{
        builderReportOutput.value = next;
        builderOutputAnimating = false;
        builderReportOutput.classList.remove('builder-output-animating');
      }
    };
    step();
  }
  function animateBuilderOutputChars(text, onDone = null){
    if (!builderReportOutput) return;
    clearBuilderOutputAnimation();
    builderOutputAnimating = true;
    builderReportOutput.classList.add('builder-output-animating');
    const token = builderOutputAnimationToken;
    const source = String(text || '');
    let i = 0;
    builderReportOutput.value = '';
    builderReportOutput.classList.remove('builder-output-fade');
    void builderReportOutput.offsetWidth;
    builderReportOutput.classList.add('builder-output-fade');
    const step = () => {
      if (token !== builderOutputAnimationToken || !builderReportOutput) return;
      builderReportOutput.value = source.slice(0, i);
      i += 4;
      if (i <= source.length){
        setTimeout(step, 7);
      }else{
        builderReportOutput.value = source;
        builderOutputAnimating = false;
        builderReportOutput.classList.remove('builder-output-animating');
        if (typeof onDone === 'function') onDone();
      }
    };
    step();
  }
  function setBuilderReportOutputText(nextText, mode = 'instant'){
    if (!builderReportOutput) return;
    const previous = String(builderReportOutput.value || '');
    const next = String(nextText || '');
    builderLastFullOutput = next;
    if (mode === 'selection'){
      clearBuilderDiffOverlay();
      const change = getSelectionChangeMeta(previous, next);
      const isPureRemoval = !!change.prevMiddle && !change.nextMiddle;
      if (isPureRemoval){
        animateBuilderOutputRemoval(previous, next);
      }else{
        animateBuilderOutputWords(previous, next);
      }
      return;
    }
    if (mode === 'revise'){
      animateBuilderOutputChars(next, () => renderBuilderDiffOverlay(previous, next));
      return;
    }
    clearBuilderOutputAnimation();
    clearBuilderDiffOverlay();
    builderReportOutput.value = next;
    if (mode === 'generate'){
      const changed = countChangedSentences(previous, next);
      const pulses = Math.min(Math.max(changed, 1), 3);
      for (let p = 0; p < pulses; p += 1){
        setTimeout(() => pulseBuilderOutput('builder-output-pulse'), p * 110);
      }
    }
  }
  function setBuilderRevisedOutputText(text, mode = 'instant'){
    if (!builderRevisedOutput) return;
    const next = String(text || '');
    builderRevisedAnimationToken += 1;
    builderRevisedOutput.value = next;
    if (mode === 'revise'){
      builderRevisedOutput.classList.remove('builder-output-fade');
      void builderRevisedOutput.offsetWidth;
      builderRevisedOutput.classList.add('builder-output-fade');
    }else{
      builderRevisedOutput.classList.remove('builder-output-fade');
    }
  }
  function getSelectedBuilderAssignments(){
    return Array.from(document.querySelectorAll('#builderAssignmentsList input[type="checkbox"]:checked'))
      .map(cb => cb.value)
      .filter(Boolean);
  }
  function normalizeBuilderAiEndpoint(raw){
    const value = String(raw || '').trim();
    if (!value) return '';
    const noSlash = value.replace(/\/+$/, '');
    if (/\/api\/generate-comment$/i.test(noSlash)) return noSlash;
    if (/^https?:\/\//i.test(noSlash)) return `${noSlash}/api/generate-comment`;
    return '';
  }
  function ensureBuilderAiEndpoint(){
    const current = normalizeBuilderAiEndpoint(builderAiEndpoint);
    if (current){
      builderAiEndpoint = current;
      return current;
    }
    const entered = window.prompt(
      'Paste your deployed Revise API URL (example: https://teacher-tools-api.vercel.app).',
      builderAiEndpoint || ''
    );
    if (entered == null) return '';
    const normalized = normalizeBuilderAiEndpoint(entered);
    if (!normalized){
      status('Invalid API URL. Use a full https:// URL.');
      return '';
    }
    builderAiEndpoint = normalized;
    saveSettings({ builderAiEndpoint });
    return normalized;
  }
  function buildBuilderAiPayload(draft){
    const selectedComments = getBuilderSelectedComments();
    const assignments = getSelectedBuilderAssignments();
    const studentRow = rows[builderSelectedRowIndex] || null;
    const gradeColumn = commentConfig.gradeColumn || FINAL_GRADE_COLUMN;
    const finalGradeMeta = studentRow && gradeColumn ? deriveMarkMeta(studentRow[gradeColumn], gradeColumn) : null;
    const finalGrade = finalGradeMeta ? formatMarkText(finalGradeMeta, studentRow[gradeColumn]) : '';
    const orderMode = ['sandwich', 'selection', 'bullet'].includes(commentOrderMode) ? commentOrderMode : 'selection';
    return {
      draft: String(draft || '').trim(),
      reviseMode: 'rewrite',
      orderMode,
      sentenceFormatHint: orderMode === 'bullet' ? 'bullets' : 'paragraph',
      studentName: builderStudentNameInput?.value.trim() || '',
      pronoun: builderPronounFemaleInput?.checked ? 'she' : 'he',
      gradeGroup: builderGradeGroupSelect?.value || 'middle',
      performanceLevel: builderCorePerformanceSelect?.value || '',
      termLabel: builderTermSelector?.value || '',
      finalGrade,
      selectedAssignments: assignments,
      selectedComments: selectedComments.map(item => ({
        id: item.id,
        category: item.section,
        text: item.text,
        type: item.type
      })),
      customComment: builderCustomCommentInput?.value.trim() || ''
    };
  }
  function parseAiCommentResponse(data){
    if (!data || typeof data !== 'object') return '';
    const candidates = [
      data.comment,
      data.text,
      data.output,
      data.result
    ];
    for (const item of candidates){
      const text = String(item || '').trim();
      if (text) return text;
    }
    return '';
  }
  function getWordCount(text){
    return String(text || '').trim().split(/\s+/).filter(Boolean).length;
  }
  function getSentenceCount(text){
    return String(text || '').trim().split(/(?<=[.!?])\s+/).map(s => s.trim()).filter(Boolean).length;
  }
  function shouldKeepOriginalDraft(draft, revised, orderMode){
    const draftText = String(draft || '').trim();
    const revisedText = String(revised || '').trim();
    const draftWords = getWordCount(draftText);
    const revisedWords = getWordCount(revisedText);
    if (!revisedWords) return true;
    if (draftWords >= 28 && revisedWords < Math.max(20, Math.floor(draftWords * 0.6))) return true;
    if (orderMode !== 'bullet' && getSentenceCount(revisedText) < 4) return true;
    return false;
  }
  async function builderGenerateReportWithAI(){
    if (!builderReportOutput) return;
    const draft = String(builderLastFullOutput || builderReportOutput.value || '').trim();
    if (!draft){
      status('Generate a local draft first, then click Revise.');
      return;
    }
    const endpoint = ensureBuilderAiEndpoint();
    if (!endpoint) return;
    const payload = buildBuilderAiPayload(draft);
    const originalLabel = builderGenerateAiBtn?.textContent || 'Revise';
    try{
      if (builderGenerateAiBtn){
        builderGenerateAiBtn.disabled = true;
        builderGenerateAiBtn.textContent = 'Revising...';
      }
      status('Revising comment with AI...');
      const response = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok){
        const detail = String(data?.error || data?.message || response.statusText || 'Request failed');
        status(`AI error: ${detail}`);
        return;
      }
      const aiText = parseAiCommentResponse(data);
      if (!aiText){
        status('AI returned an empty comment.');
        return;
      }
      const modelUsed = String(data?.modelUsed || '').trim();
      const polished = polishGrammar(cleanFluency(aiText));
      setBuilderRevisedOutputText(polished, 'revise');
      if (polished.trim() === draft.trim()){
        status('AI returned the same wording. Try Revise again for a different rewrite.');
      }else{
        status(modelUsed ? `AI revision ready (${modelUsed}).` : 'AI revision ready.');
      }
    }catch(err){
      console.error(err);
      status('Could not reach AI API.');
    }finally{
      if (builderGenerateAiBtn){
        builderGenerateAiBtn.disabled = false;
        builderGenerateAiBtn.textContent = originalLabel;
      }
    }
  }

  function builderGenerateReport(){
    if (!builderReportOutput) return;
    const outputMode = builderOutputAnimationMode || 'instant';
    builderOutputAnimationMode = 'instant';
    setBuilderRevisedOutputText('', 'instant');
    const studentName = builderStudentNameInput?.value.trim();
    const coreLevel = builderCorePerformanceSelect?.value;
    if (!studentName){
      setBuilderReportOutputText('Provide a student name.', 'instant');
      return;
    }
    if (!coreLevel && (builderGradeGroupSelect?.value || 'middle') !== 'elem'){
      setBuilderReportOutputText('Select a core performance level.', 'instant');
      return;
    }
    const studentRow = rows[builderSelectedRowIndex];
    const context = {
      studentName,
      pronouns: getBuilderPronouns(),
      retestScore: builderRetestInput?.value.trim(),
      originalScore: builderOriginalInput?.value.trim(),
      termAverage: builderTermAverageInput?.value.trim(),
      termLabel: builderTermSelector?.value || '',
      customComment: builderCustomCommentInput?.value.trim(),
      gradeGroup: builderGradeGroupSelect?.value || "middle",
      includeFinalGrade: !!builderIncludeFinalGradeInput?.checked
    };

    const gradeColumn = commentConfig.gradeColumn || FINAL_GRADE_COLUMN;
    const gradeMeta = (context.gradeGroup === 'elem') ? null : (studentRow && gradeColumn ? deriveMarkMeta(studentRow[gradeColumn], gradeColumn) : null);
    const derivedGrade = gradeMeta ? gradeMeta.raw : '';
    const providedGrade = (context.termAverage || '').trim();
    const providedMeta = (context.gradeGroup === 'elem' || !providedGrade) ? null : deriveMarkMeta(providedGrade, gradeColumn);
    const finalGradeToUse = derivedGrade || (providedMeta ? providedMeta.raw : providedGrade);
    context.termAverage = (context.gradeGroup === 'elem') ? '' : finalGradeToUse;

    const selectedAssignments = getSelectedBuilderAssignments();
    const overallMeta = (context.gradeGroup === 'elem') ? null : deriveMarkMeta(finalGradeToUse, gradeColumn);
    const assignmentSentences = buildAssignmentSentences(selectedAssignments, studentRow, context, overallMeta);
    const assignmentParagraph = assignmentSentences || '';

    const selectedComments = getBuilderSelectedComments();
    const commentBlocks = buildCommentBlocks(selectedComments, context);
    const lookingAheadSelected = commentBlocks.lookingAheadPositiveSentences?.[0] || '';
    const commentSnippetText = commentBlocks.otherText;
    const toneLine = getPerformanceToneLine(coreLevel, context);

    let report = '';
    const baseComment = studentRow ? buildGradeBasedComment(studentRow, context.studentName, context.pronouns, context.gradeGroup, context.includeFinalGrade) : '';
    if (baseComment){
      const introLine = getPerformanceIntroLine(coreLevel, context);
      let baseWithIntro = introLine ? replaceFirstSentence(baseComment, introLine) : baseComment;
      if (commentBlocks.pqSentences.length){
        baseWithIntro = insertSentencesAfterFirst(baseWithIntro, commentBlocks.pqSentences, 3);
      }
      report = buildUnifiedComment({
        baseComment: baseWithIntro,
        termLabel: context.termLabel,
        toneLine,
        commentSnippet: commentSnippetText,
        assignmentParagraph,
        lookingAheadText: lookingAheadSelected
      });
    }else{
      const template = getReportBuilderTemplate(coreLevel, context);
      if (!template){
        setBuilderReportOutputText('Template not available for this selection.', 'instant');
        return;
      }
      let partA = template.partA;
      let partB = template.partB;
      const score1 = '';
      const score2 = '';
      const ccComment = builderGetCCCommentary(builderCCPerformanceSelect?.value || '');
      partA = partA.replace('[SCORE_COMMENTARY_TEST1]', score1)
                   .replace('[SCORE_COMMENTARY_TEST2]', score2)
                   .replace('[CC_COMMENTARY]', ccComment);
      if (partA.includes('[RETEST_CLAUSE]')){
        if (context.retestScore && context.originalScore && template.retestClause){
          partA = partA.replace('[RETEST_CLAUSE]', template.retestClause);
        }else{
          partA = partA.replace('[RETEST_CLAUSE]', template.noRetestClause || '');
          partA = partA.replace('[SCORE_COMMENTARY_TEST2]', score2);
        }
      }
      partA = builderReplacePlaceholders(partA, context);
      partB = builderReplacePlaceholders(partB, context);
      const blocks = [];
      if (context.termLabel){
        blocks.push(`${context.termLabel}: ${partA}`);
      }else{
        blocks.push(partA);
      }
      if (toneLine) blocks.push(toneLine);
      if (commentSnippetText) blocks.push(commentSnippetText);
      if (assignmentParagraph) blocks.push(assignmentParagraph);
      const tailText = lookingAheadSelected;
      if (tailText) blocks.push(tailText);
      if (partB) blocks.push(partB);
      report = blocks.filter(Boolean).join('\n\n');
    }

    if (context.customComment){
      report += ' ' + builderReplacePlaceholders(context.customComment, context);
    }

    const pronounAdjusted = applyPronounsAfterFirstSentence(report, context.studentName, context.pronouns);
    const trimmedReport = cleanFluency(pronounAdjusted.trim());
    const requirements = buildGenerationRequirements({
      sourceText: trimmedReport,
      selectedAssignments,
      studentRow,
      overallMeta,
      lookingAheadText: lookingAheadSelected
    });
    const dedupedReport = dedupeGeneratedComment(trimmedReport, context, requirements);
    const factSafeReport = ensureRequiredFactsInReport(dedupedReport, requirements);
    const polishedReport = polishGrammar(factSafeReport);
    setBuilderReportOutputText(polishedReport, outputMode);
    builderReportOutput.dataset.lastSig = computeReportSignature(polishedReport, {
      studentName: context.studentName,
      coreLevel,
      termLabel: context.termLabel,
      commentIds: selectedComments.map(c => c.id).filter(Boolean),
      assignments: selectedAssignments
    });
  }

  function builderCopyReport(){
    if (!builderReportOutput || !String(builderReportOutput.value || '').trim()) return;
    copyTextToClipboard(String(builderReportOutput.value || '').trim());
  }
  function getCurrentReportData(){
    const text = String(builderReportOutput?.value || '').trim();
    if (!text) return null;
    const selectedComments = getBuilderSelectedComments();
    const assignments = Array.from(document.querySelectorAll('#builderAssignmentsList input[type="checkbox"]:checked')).map(cb => cb.value);
    return {
      text,
      studentName: builderStudentNameInput?.value.trim() || '',
      coreLevel: builderCorePerformanceSelect?.value || '',
      termLabel: builderTermSelector?.value || '',
      commentIds: selectedComments.map(c => c.id).filter(Boolean),
      commentLabels: selectedComments.map(c => c.text).filter(Boolean),
      assignments
    };
  }
  function computeReportSignature(text, meta){
    const payload = {
      text: text || '',
      studentName: meta?.studentName || '',
      coreLevel: meta?.coreLevel || '',
      termLabel: meta?.termLabel || '',
      commentIds: Array.isArray(meta?.commentIds) ? [...meta.commentIds].sort() : [],
      assignments: Array.isArray(meta?.assignments) ? [...meta.assignments].sort() : []
    };
    return JSON.stringify(payload);
  }
  function saveCurrentReport(reason = 'manual'){
    const data = getCurrentReportData();
    if (!data || !data.text) return;
    const signature = computeReportSignature(data.text, data);
    const exists = savedReports.some(r => r.signature === signature);
    if (exists) return;
    const entry = {
      id: Date.now() + '_' + Math.random().toString(36).slice(2,8),
      ts: Date.now(),
      signature,
      ...data
    };
    savedReports.unshift(entry);
    if (savedReports.length > SAVED_REPORT_LIMIT){
      savedReports = savedReports.slice(0, SAVED_REPORT_LIMIT);
    }
    persistSavedReports();
    renderSavedReports();
    status(reason === 'copy' ? 'Saved after Copy.' : 'Comment saved.');
  }
  function autoSaveCurrentReport(reason = 'auto'){
    if (!builderReportOutput) return;
    const lastSig = builderReportOutput.dataset.lastSig || '';
    const data = getCurrentReportData();
    if (!data || !data.text) return;
    const sig = computeReportSignature(data.text, data);
    if (sig === lastSig) return; // no changes since last generate
    const exists = savedReports.some(r => r.signature === sig);
    if (exists) return;
    saveCurrentReport(reason);
  }
  function restoreSavedReport(entry){
    if (!entry) return;
    if (builderStudentNameInput) builderStudentNameInput.value = entry.studentName || '';
    if (builderCorePerformanceSelect) builderCorePerformanceSelect.value = entry.coreLevel || '';
    if (builderTermSelector) builderTermSelector.value = entry.termLabel || '';
    // restore comment selections
    if (!builderBankRendered && builderCommentBankEl){
      buildBuilderCommentBank();
      builderBankRendered = true;
    }
    builderCommentCheckboxes = Array.from(document.querySelectorAll('#builderCommentBank input[type="checkbox"]'));
    const idSet = new Set(entry.commentIds || []);
    builderCommentCheckboxes.forEach(cb => {
      cb.checked = idSet.has(cb.dataset.id);
    });
    updateBuilderSelectedTags();
    // restore assignments
    buildBuilderAssignmentsList();
    const assignSet = new Set(entry.assignments || []);
    const assignmentCbs = Array.from(document.querySelectorAll('#builderAssignmentsList input[type="checkbox"]'));
    assignmentCbs.forEach(cb => {
      cb.checked = assignSet.has(cb.value);
    });
    // restore text
    if (builderReportOutput){
      setBuilderReportOutputText(entry.text || '', 'instant');
      builderReportOutput.dataset.lastSig = entry.signature || '';
    }
    builderGenerateReport();
  }
  function renderSavedReports(){
    if (!savedReportsListEl) return;
    savedReportsListEl.innerHTML = '';
    if (!savedReports.length){
      const empty = document.createElement('div');
      empty.className = 'muted';
      empty.style.fontSize = '12px';
      empty.textContent = 'No saved comments yet.';
      savedReportsListEl.appendChild(empty);
      return;
    }
    savedReports.forEach(item => {
      const wrap = document.createElement('div');
      wrap.className = 'saved-report-item';
      const header = document.createElement('header');
      const title = document.createElement('strong');
      title.textContent = item.studentName || 'Unnamed';
      const meta = document.createElement('div');
      meta.className = 'saved-report-meta';
      const term = document.createElement('span');
      term.textContent = item.termLabel ? item.termLabel : 'No term';
      const level = document.createElement('span');
      level.textContent = item.coreLevel || 'No level';
      const time = document.createElement('span');
      time.textContent = new Date(item.ts).toLocaleString();
      meta.appendChild(term);
      meta.appendChild(level);
      meta.appendChild(time);
      const actions = document.createElement('div');
      actions.className = 'saved-report-actions';
      const copyBtn = document.createElement('button');
      copyBtn.className = 'btn';
      copyBtn.textContent = 'Copy';
      copyBtn.dataset.action = 'copy';
      copyBtn.dataset.id = item.id;
      const restoreBtn = document.createElement('button');
      restoreBtn.className = 'btn';
      restoreBtn.textContent = 'Restore';
      restoreBtn.dataset.action = 'restore';
      restoreBtn.dataset.id = item.id;
      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'btn warn';
      deleteBtn.textContent = 'Delete';
      deleteBtn.dataset.action = 'delete';
      deleteBtn.dataset.id = item.id;
      actions.appendChild(copyBtn);
      actions.appendChild(restoreBtn);
      actions.appendChild(deleteBtn);
      header.appendChild(title);
      header.appendChild(actions);
      wrap.appendChild(header);
      wrap.appendChild(meta);
      const snippet = document.createElement('div');
      snippet.className = 'muted';
      snippet.style.fontSize = '12px';
      snippet.style.marginTop = '6px';
      snippet.textContent = item.text.slice(0, 180) + (item.text.length > 180 ? '…' : '');
      wrap.appendChild(snippet);
      if (Array.isArray(item.commentLabels) && item.commentLabels.length){
        const tags = document.createElement('div');
        tags.style.display = 'flex';
        tags.style.flexWrap = 'wrap';
        tags.style.gap = '4px';
        tags.style.marginTop = '6px';
        item.commentLabels.slice(0, 6).forEach(lbl => {
          const chip = document.createElement('span');
          chip.className = 'badge';
          chip.textContent = lbl.slice(0, 40);
          tags.appendChild(chip);
        });
        wrap.appendChild(tags);
      }
      savedReportsListEl.appendChild(wrap);
    });
  }
  function clearSelectedComments(){
    if (!builderCommentCheckboxes.length && builderCommentBankEl){
      buildBuilderCommentBank();
      builderCommentCheckboxes = Array.from(builderCommentBankEl.querySelectorAll('input[type="checkbox"]'));
    }
    builderCommentCheckboxes.forEach(cb => cb.checked = false);
    updateBuilderSelectedTags();
    builderGenerateReport();
  }
  function resetToEmptyState(){
    rows = [];
    originalRows = [];
    allColumns = [];
    columnOrder = [];
    visibleColumns = [];
    filteredIdx = [];
    visibleRowSet = new Set();
    sortState = {key:null, dir:null};
    undoStack = [];
    redoStack = [];
    firstNameKey = null;
    lastNameKey = null;
    studentNameColumn = null;
    studentNameWarning = '';
    currentFileName = null;
    metaEl.textContent = 'No file loaded';
    countsEl.textContent = '';
    colCountEl.textContent = '0';
    rowCountEl.textContent = '0';
    colsDiv.innerHTML = '';
    rowsDiv.innerHTML = '';
    thead.innerHTML = '';
    tbody.innerHTML = '';
    setDataActionButtonsDisabled(true);
    editHeadersBtn.disabled = true;
    transposeBtn.disabled = true;
    highlightActiveFile();
    renderWarning('');
    resetPerformanceFilterState();
  }

  function applySavedHeaderMapping(){
    if (!currentFileName) return;
    const s = loadSettings();
    const mAll = s.headerMaps || {};
    const mapping = mAll[currentFileName] || {};
    if (Object.keys(mapping).length){
      renameColumns(mapping, false);
    }
  }
  function remapRows(data, mapping){
    if (!Array.isArray(data)) return [];
    return data.map(obj => {
      const newObj = {};
      Object.keys(obj || {}).forEach(k => {
        const nk = mapping[k] || k;
        newObj[nk] = obj[k];
      });
      return newObj;
    });
  }
  function applyMappingToContext(ctx, mapping){
    if (!ctx || !Object.keys(mapping).length) return;
    ctx.rows = remapRows(ctx.rows, mapping);
    ctx.originalRows = remapRows(ctx.originalRows, mapping);
    ctx.allColumns = ctx.allColumns.map(c => mapping[c] || c);
    ctx.visibleColumns = ctx.visibleColumns.map(c => mapping[c] || c);
    ctx.columnOrder = ctx.columnOrder.map(c => mapping[c] || c);
    if (ctx.firstNameKey) ctx.firstNameKey = mapping[ctx.firstNameKey] || ctx.firstNameKey;
    if (ctx.lastNameKey) ctx.lastNameKey = mapping[ctx.lastNameKey] || ctx.lastNameKey;
    if (ctx.studentNameColumn) ctx.studentNameColumn = mapping[ctx.studentNameColumn] || ctx.studentNameColumn;
    if (ctx.commentConfig && Array.isArray(ctx.commentConfig.extraColumns)){
      ctx.commentConfig.extraColumns = ctx.commentConfig.extraColumns.map(c => mapping[c] || c);
    }
    stampEndIndicator(ctx.rows);
    stampEndIndicator(ctx.originalRows);
    ctx.allColumns = ensureListHasEnd(ctx.allColumns);
    ctx.columnOrder = ensureListHasEnd(ctx.columnOrder);
    const set = new Set(ctx.visibleColumns);
    set.add(END_OF_LINE_COL);
    ctx.visibleColumns = ctx.columnOrder.filter(c => set.has(c));
  }
  function persistHeaderMapping(fileName, mapping){
    if (!fileName) return;
    const s = loadSettings();
    const all = s.headerMaps || {};
    all[fileName] = {...(all[fileName] || {}), ...mapping};
    saveSettings({ headerMaps: all });
  }

  function invertMapping(map){
    const inv = {};
    for (const [k,v] of Object.entries(map)) inv[v] = k;
    return inv;
  }

  // ==== Utils ====
  function isLikelyMarkColumn(name){
    const lower = (name || '').toLowerCase();
    if (!lower) return false;
    return MARK_COLUMN_HINTS.some(token => lower.includes(token));
  }
  function datasetHasPercentSymbol(){
    const markCols = allColumns.filter(isLikelyMarkColumn);
    const colsToScan = markCols.length ? markCols : allColumns;
    if (!colsToScan.length) return false;
    const maxRows = Math.min(rows.length, 300);
    for (let i = 0; i < maxRows; i++){
      const row = rows[i];
      if (!row) continue;
      for (const col of colsToScan){
        const val = row[col];
        if (typeof val === 'string' && val.includes('%')) return true;
      }
    }
    return false;
  }
  function columnHasPercentValues(col){
    if (!col) return false;
    const maxRows = Math.min(rows.length, 500);
    for (let i = 0; i < maxRows; i++){
      const val = rows[i]?.[col];
      if (typeof val === 'string' && val.includes('%')) return true;
    }
    return false;
  }
  function deriveMarkMeta(value, columnName){
    const rawStr = String(value ?? '').trim();
    if (!rawStr) return null;
    const hint = isLikelyMarkColumn(columnName) || rawStr.includes('%');
    if (!hint) return null;
    const match = rawStr.replace(/,/g,'').match(/-?\d+(\.\d+)?/);
    if (!match) return null;
    let num = parseFloat(match[0]);
    if (Number.isNaN(num)) return null;
    if (num < 0) num = 0;
    if (num > 100) num = 100;
    const className = num > 79 ? 'mark-high' : (num >= 70 ? 'mark-mid' : 'mark-low');
    const formattedNumber = Number(num.toFixed(2)).toString();
    return { className, raw: formattedNumber, value: num, hasPercent: rawStr.includes('%') };
  }
  function applyMarkStyling(td, value, columnName){
    td.classList.remove('mark-high','mark-mid','mark-low');
    const meta = deriveMarkMeta(value ?? td.textContent, columnName);
    if (meta && markColorsEnabled){
      td.classList.add(meta.className);
    }
    return meta;
  }
  function normalizeMarkInput(value, columnName){
    const raw = value == null ? '' : String(value).trim();
    const meta = deriveMarkMeta(raw, columnName);
    if (!meta) return raw;
    return meta.hasPercent ? `${meta.raw} %` : meta.raw;
  }
  function formatMarkText(meta, rawValue){
    if (!meta) return rawValue ?? '';
    return meta.hasPercent ? `${meta.raw} %` : meta.raw;
  }
  function applyMarkColorSetting(enabled){
    markColorsEnabled = !!enabled;
    if (markColorsToggle) markColorsToggle.checked = markColorsEnabled;
    saveSettings({ markColors: markColorsEnabled });
    render();
  }
  function deepClone(x){ return JSON.parse(JSON.stringify(x)); }
  function escapeHtml(s){ return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
  function escapeRegExp(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function status(msg){ statusEl.textContent = msg; }
  function renderWarning(msg){
    if (msg){
      filesWarningTextEl.textContent = msg;
      filesWarningEl.classList.add('show');
    } else {
      filesWarningEl.classList.remove('show');
      filesWarningTextEl.textContent = '';
    }
  }
  function normalizeHeader(name){ return String(name || '').trim().toLowerCase(); }
  function addStudentNames(data){
    if (!studentNameColumn || !firstNameKey || !lastNameKey || !Array.isArray(data)) return;
    data.forEach(row => {
      if (!row) return;
      row[studentNameColumn] = buildStudentName(row);
    });
  }
  function getFirstNameFromRow(row, rowIndex){
    if (!row) return '';
    const first = firstNameKey ? String(row[firstNameKey] ?? '').trim() : '';
    if (first) return first;
    const composite = row[studentNameColumn] || buildStudentName(row) || '';
    const token = String(composite).trim().split(/\s+/)[0];
    if (token) return token;
    return rowIndex != null ? getRowLabel(rowIndex) : '';
  }
  function buildStudentName(row){
    if (!row) return '';
    const first = String(row[firstNameKey] ?? '').trim();
    const last = String(row[lastNameKey] ?? '').trim();
    if (first && last) return `${first} ${last}`;
    return first || last || '';
  }
  function refreshStudentNameForRow(rowIndex){
    if (!studentNameColumn || rowIndex == null) return;
    const row = rows[rowIndex];
    if (!row) return;
    row[studentNameColumn] = buildStudentName(row);
    updateStudentNameCells(rowIndex);
    handleColumnDataMutation(studentNameColumn);
  }
  function updateStudentNameCells(rowIndex){
    if (!studentNameColumn) return;
    const value = rows[rowIndex]?.[studentNameColumn] ?? '';
    const rowEl = tbody.querySelector(`tr[data-row-index="${rowIndex}"]`);
    if (rowEl){
      const cell = rowEl.querySelector(`td[data-col="${studentNameColumn}"]`);
      if (cell) cell.textContent = value;
    }
    const transposedRow = tbody.querySelector(`tr[data-column-key="${studentNameColumn}"]`);
    if (transposedRow){
      const idxPos = filteredIdx.indexOf(rowIndex);
      if (idxPos !== -1){
        const cell = transposedRow.querySelector(`td[data-cpos="${idxPos}"]`);
        if (cell) cell.textContent = value;
      }
    }
  }
  function isNameColumn(col){
    return !!col && (col === firstNameKey || col === lastNameKey);
  }
  function arraysEqual(a,b){
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++){
      if (a[i] !== b[i]) return false;
    }
    return true;
  }
  function stampEndIndicator(data){
    if (!Array.isArray(data)) return;
    data.forEach(row => {
      if (!row || typeof row !== 'object') return;
      row[END_OF_LINE_COL] = END_OF_LINE_VALUE;
    });
  }
  function ensureListHasEnd(list){
    const base = Array.isArray(list) ? list.slice() : [];
    const filtered = base.filter(c => c !== END_OF_LINE_COL);
    filtered.push(END_OF_LINE_COL);
    return filtered;
  }
  function enforceEndOfLineColumn(){
    stampEndIndicator(rows);
    stampEndIndicator(originalRows);
    allColumns = ensureListHasEnd(allColumns);
    columnOrder = ensureListHasEnd(columnOrder);
    if (!visibleColumns.includes(END_OF_LINE_COL)) visibleColumns.push(END_OF_LINE_COL);
    syncVisibleOrder();
    if (!visibleColumns.includes(END_OF_LINE_COL)){
      visibleColumns.push(END_OF_LINE_COL);
      syncVisibleOrder();
    }
  }

  function getLessonNumber(col){
    if (!col) return null;
    const text = String(col).toUpperCase();
    const match = text.match(/L\s*(\d+)/);
    if (!match) return null;
    const num = Number(match[1]);
    return Number.isNaN(num) ? null : num;
  }
  function selectLessonRange(min, max){
    if (!Array.isArray(columnOrder) || !columnOrder.length) return;
    const before = [...visibleColumns];
    const next = [];
    const addUnique = (col) => {
      if (!col) return;
      if (!columnOrder.includes(col)) return;
      if (next.includes(col)) return;
      next.push(col);
    };
    const studentCol = studentNameColumn
      || columnOrder.find(col => normalizeHeader(col) === normalizeHeader(STUDENT_NAME_LABEL));
    addUnique(studentCol);
    const orgCol = columnOrder.find(col => normalizeHeader(col) === 'orgid');
    addUnique(orgCol);
    columnOrder.forEach(col => {
      const num = getLessonNumber(col);
      if (num == null) return;
      if (num >= min && num <= max) addUnique(col);
    });
    addUnique(END_OF_LINE_COL);
    if (next.length <= 2) return; // only reserved columns, no lessons found
    visibleColumns = next;
    pushUndo({type:'setVisibleColumns', before, after:[...visibleColumns]});
    saveSettings({ visibleColumns });
    rebuildColsUI();
    render();
  }

  // ==== Hints modal ====
  const hintsBtn = document.getElementById('hintsBtn');
  const hintsModal = document.getElementById('hintsModal');
  const hintsCloseBtn = document.getElementById('hintsCloseBtn');
  const hintsGotItBtn = document.getElementById('hintsGotItBtn');

  function openHints(){
    if (typeof window.showAllTourHints === 'function'){
      window.showAllTourHints();
      return;
    }
    if (hintsModal) hintsModal.style.display = 'flex';
  }
  function closeHints(){
    if (hintsModal) hintsModal.style.display = 'none';
  }
  if (hintsBtn){
    hintsBtn.addEventListener('click', openHints);
  }
  if (hintsCloseBtn){
    hintsCloseBtn.addEventListener('click', closeHints);
  }
  if (hintsGotItBtn){
    hintsGotItBtn.addEventListener('click', closeHints);
  }
  if (hintsModal){
    hintsModal.addEventListener('click', (e) => {
      if (e.target === hintsModal) closeHints();
    });
  }
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && hintsModal && hintsModal.style.display === 'flex'){
      closeHints();
    }
  });
