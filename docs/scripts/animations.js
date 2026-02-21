(function initAnimations(){
  const WELCOME_KEY = 'teacher_tools_welcome_shown';
  const TOUR_KEY = 'teacher_tools_tour_shown';
  const TOUR_PADDING = 12;
  const TOUR_FALLBACK_TOP = 72;
  const TOUR_PANEL_MIN_WIDTH = 260;
  const TOUR_PANEL_MAX_WIDTH = 340;

  const dropzone = window.dropzone || document.getElementById('dropzone');

  const state = {
    steps: [],
    index: -1,
    isOpen: false,
    overlayEl: null,
    panelEl: null,
    onKeydown: null
  };

  function clamp(value, min, max){
    return Math.max(min, Math.min(max, value));
  }

  function isTargetVisible(el){
    if (!el || !el.getBoundingClientRect) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width < 2 || rect.height < 2) return false;
    return rect.bottom > 0 && rect.right > 0 && rect.top < window.innerHeight && rect.left < window.innerWidth;
  }

  function getTourSteps(){
    return [
      {
        el: document.getElementById('filesDrawer'),
        title: 'Files panel (left side)',
        desc: 'This panel holds all loaded class files. Click a file to switch the active class across Data, Print View, and Report Comments.'
      },
      {
        el: document.getElementById('filesDrawerToggle'),
        title: 'Collapse or expand files panel',
        desc: 'Use this button to collapse the left panel for more workspace, or expand it to manage files.'
      },
      {
        el: dropzone,
        title: 'Upload your data',
        desc: 'Drop CSV/XLSX files or click to pick files quickly.'
      },
      {
        el: document.getElementById('loadDirBtn'),
        title: 'Load a whole folder',
        desc: 'Pick a folder with CSV/XLSX files to load them all at once.'
      },
      {
        el: document.getElementById('zoomOutBtn'),
        title: 'Zoom controls',
        desc: 'Use minus/plus to zoom the workspace out or in for comfortable viewing.'
      },
      {
        el: document.getElementById('darkModeBtn'),
        title: 'Dark mode toggle',
        desc: 'Switch between light and dark themes at any time.'
      },
      {
        el: document.getElementById('hintsBtn'),
        title: 'Hints button',
        desc: 'Open quick hints whenever you want a refresher.'
      },
      {
        el: document.getElementById('tabCommentsBtn'),
        title: 'Generate report comments',
        desc: 'Open Report Comments to auto-generate feedback and copy it for students.'
      },
      {
        el: document.getElementById('tabPrintBtn'),
        title: 'Print rosters and sheets',
        desc: 'Use Print View to generate attendance, marking, or drill sheets for selected classes.'
      },
      {
        el: document.getElementById('transposeBtn'),
        title: 'Transpose the table',
        desc: 'Swap rows and columns. Exports respect the current layout.'
      },
      {
        el: document.getElementById('wrapHeadersToggle'),
        title: 'Wrap long headers',
        desc: 'Toggle wrapping to see long header names without resizing columns.'
      },
      {
        el: document.getElementById('columnDataFilterBtn'),
        title: 'Show data columns only',
        desc: 'Filter the column list to only headers that contain data.'
      },
      {
        el: document.getElementById('exportCsvBtn'),
        title: 'Export CSV',
        desc: 'Export visible columns/rows to CSV. XLSX is available too.'
      },
      {
        el: document.getElementById('exportXlsxBtn'),
        title: 'Export XLSX',
        desc: 'Download the visible table as an XLSX workbook.'
      },
      {
        el: document.getElementById('resetBtn'),
        title: 'Reset view',
        desc: 'Clear search, filters, and visibility settings to start fresh.'
      },
      {
        el: document.getElementById('editHeadersBtn'),
        title: 'Bulk edit headers',
        desc: 'Rename/trim headers for consistency. All changes can be undone.'
      },
      {
        el: document.getElementById('search'),
        title: 'Filter rows fast',
        desc: 'Type to filter visible rows instantly across selected columns.'
      }
    ];
  }

  function clearHighlights(){
    state.steps.forEach(step => step?.el?.classList?.remove('tour-highlight', 'tour-highlight-all', 'tour-focus-target'));
    state.overlayEl?.classList?.remove('tour-focus');
  }

  function teardownTour(markShown){
    clearHighlights();
    if (state.onKeydown){
      document.removeEventListener('keydown', state.onKeydown);
      state.onKeydown = null;
    }
    state.overlayEl?.remove();
    state.panelEl = null;
    state.overlayEl = null;
    state.steps = [];
    state.index = -1;
    state.isOpen = false;
    if (markShown){
      sessionStorage.setItem(TOUR_KEY, '1');
    }
  }

  function findNextVisibleIndex(startIdx, direction){
    for (let i = startIdx; i >= 0 && i < state.steps.length; i += direction){
      if (isTargetVisible(state.steps[i]?.el)) return i;
    }
    return -1;
  }

  function positionPanel(targetEl){
    const panel = state.panelEl;
    if (!panel) return;

    panel.style.display = 'block';
    panel.style.visibility = 'hidden';
    panel.style.opacity = '1';
    panel.style.transform = 'none';

    const panelRect = panel.getBoundingClientRect();
    const panelWidth = clamp(panelRect.width || TOUR_PANEL_MIN_WIDTH, TOUR_PANEL_MIN_WIDTH, TOUR_PANEL_MAX_WIDTH);
    const panelHeight = Math.max(panelRect.height || 140, 140);

    const maxLeft = Math.max(TOUR_PADDING, window.innerWidth - panelWidth - TOUR_PADDING);
    const maxTop = Math.max(TOUR_PADDING, window.innerHeight - panelHeight - TOUR_PADDING);

    let top = TOUR_FALLBACK_TOP;
    let left = clamp((window.innerWidth - panelWidth) / 2, TOUR_PADDING, maxLeft);

    if (isTargetVisible(targetEl)){
      const rect = targetEl.getBoundingClientRect();
      const below = rect.bottom + TOUR_PADDING;
      const above = rect.top - panelHeight - TOUR_PADDING;
      if (below <= maxTop){
        top = below;
      }else if (above >= TOUR_PADDING){
        top = above;
      }else{
        top = clamp(below, TOUR_PADDING, maxTop);
      }
      left = clamp(rect.left, TOUR_PADDING, maxLeft);
    }

    panel.style.top = `${clamp(top, TOUR_PADDING, maxTop)}px`;
    panel.style.left = `${clamp(left, TOUR_PADDING, maxLeft)}px`;
    panel.style.visibility = 'visible';
  }

  function renderStep(index){
    if (!state.isOpen || !state.panelEl) return;

    const step = state.steps[index];
    if (!step){
      teardownTour(true);
      return;
    }

    state.index = index;

    const badge = state.panelEl.querySelector('[data-tour-badge]');
    const titleEl = state.panelEl.querySelector('[data-tour-title]');
    const descEl = state.panelEl.querySelector('[data-tour-desc]');
    const backBtn = state.panelEl.querySelector('[data-tour-back]');
    const nextBtn = state.panelEl.querySelector('[data-tour-next]');

    if (badge) badge.textContent = String(index + 1);
    if (titleEl) titleEl.textContent = step.title || 'Tour';
    if (descEl) descEl.textContent = step.desc || '';
    if (backBtn) backBtn.disabled = findNextVisibleIndex(index - 1, -1) === -1;
    if (nextBtn) nextBtn.textContent = findNextVisibleIndex(index + 1, 1) === -1 ? 'Done' : 'Next';

    clearHighlights();
    const isFirstStep = index === 0;
    const firstStepAnchor = document.getElementById('filesDrawerToggle') || document.getElementById('loadDirBtn');
    if (!isFirstStep && isTargetVisible(step.el)){
      step.el.classList.add('tour-highlight', 'tour-highlight-all', 'tour-focus-target');
      state.overlayEl?.classList?.add('tour-focus');
      positionPanel(step.el);
    }else if (isFirstStep && isTargetVisible(firstStepAnchor)){
      // Keep step 1 clean: no large drawer highlight, only bubble near top-left controls.
      positionPanel(firstStepAnchor);
    }else{
      positionPanel(null);
    }
  }

  function createTourPanel(){
    const panel = document.createElement('div');
    panel.className = 'tour-panel';
    panel.style.position = 'fixed';
    panel.style.display = 'block';
    panel.style.opacity = '1';
    panel.style.pointerEvents = 'auto';
    panel.style.zIndex = '2102';
    panel.style.maxWidth = `${TOUR_PANEL_MAX_WIDTH}px`;
    panel.style.minWidth = `${TOUR_PANEL_MIN_WIDTH}px`;
    panel.style.background = '#fff';
    panel.style.border = '1px solid #e5e5e5';
    panel.style.borderRadius = '14px';
    panel.style.boxShadow = '0 14px 48px rgba(0,0,0,.16)';
    panel.style.padding = '14px 16px';
    panel.style.transform = 'none';

    const badge = document.createElement('div');
    badge.className = 'tour-badge';
    badge.dataset.tourBadge = '1';

    const title = document.createElement('h4');
    title.dataset.tourTitle = '1';

    const desc = document.createElement('p');
    desc.dataset.tourDesc = '1';

    const actions = document.createElement('div');
    actions.className = 'tour-actions';

    const backBtn = document.createElement('button');
    backBtn.className = 'btn';
    backBtn.textContent = 'Back';
    backBtn.dataset.tourBack = '1';

    const skipBtn = document.createElement('button');
    skipBtn.className = 'btn';
    skipBtn.textContent = 'Skip';
    skipBtn.dataset.tourSkip = '1';

    const nextBtn = document.createElement('button');
    nextBtn.className = 'btn primary';
    nextBtn.textContent = 'Next';
    nextBtn.dataset.tourNext = '1';

    actions.append(backBtn, skipBtn, nextBtn);
    panel.append(badge, title, desc, actions);

    backBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const prev = findNextVisibleIndex(state.index - 1, -1);
      if (prev !== -1) renderStep(prev);
    });

    nextBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      const next = findNextVisibleIndex(state.index + 1, 1);
      if (next === -1){
        teardownTour(true);
      }else{
        renderStep(next);
      }
    });

    skipBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      teardownTour(true);
    });

    return panel;
  }

  function startTour(force){
    if (!force && sessionStorage.getItem(TOUR_KEY)) return;

    teardownTour(false);

    state.steps = getTourSteps();
    if (!state.steps.length) return;

    const overlay = document.createElement('div');
    overlay.className = 'tour-overlay';
    overlay.style.pointerEvents = 'auto';

    const panel = createTourPanel();
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    state.overlayEl = overlay;
    state.panelEl = panel;
    state.isOpen = true;

    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) teardownTour(true);
    });

    state.onKeydown = (e) => {
      if (e.key === 'Escape' && state.isOpen){
        teardownTour(true);
      }
    };
    document.addEventListener('keydown', state.onKeydown);

    const firstVisible = findNextVisibleIndex(0, 1);
    if (firstVisible === -1){
      renderStep(0);
      return;
    }
    renderStep(firstVisible);
  }

  function createWelcomeOverlay(){
    const overlay = document.createElement('div');
    overlay.className = 'welcome-overlay';
    overlay.innerHTML = `
      <div class="welcome-card">
        <div class="welcome-title">Welcome to Teacher TOOLs</div>
        <div class="welcome-sub">Upload your CSV/XLSX files to get started. You can take a quick tour or skip it.</div>
        <div class="welcome-actions">
          <button class="btn welcome-btn" type="button" id="welcomeTourBtn">Show Tour</button>
          <button class="btn welcome-btn secondary" type="button" id="welcomeSkipBtn">Skip Tour</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    return overlay;
  }

  function closeWelcome(overlay){
    if (!overlay) return;
    overlay.classList.remove('show');
    setTimeout(() => overlay.remove(), 250);
  }

  function showWelcome(){
    if (sessionStorage.getItem(WELCOME_KEY)) return;

    const overlay = createWelcomeOverlay();
    requestAnimationFrame(() => overlay.classList.add('show'));
    if (dropzone) dropzone.classList.add('pulse');

    const tourBtn = document.getElementById('welcomeTourBtn');
    const skipBtn = document.getElementById('welcomeSkipBtn');

    const finishWelcome = () => {
      if (dropzone) dropzone.classList.remove('pulse');
      sessionStorage.setItem(WELCOME_KEY, '1');
      closeWelcome(overlay);
    };

    tourBtn?.addEventListener('click', () => {
      finishWelcome();
      requestAnimationFrame(() => startTour(true));
    });

    skipBtn?.addEventListener('click', () => {
      finishWelcome();
    });

    overlay.addEventListener('click', (e) => {
      if (e.target === overlay){
        finishWelcome();
      }
    });
  }

  // Hints button uses the same deterministic tour engine.
  window.showAllTourHints = function showAllTourHints(){
    startTour(true);
  };

  window.resetTourState = function resetTourState(){
    sessionStorage.removeItem(WELCOME_KEY);
    sessionStorage.removeItem(TOUR_KEY);
    teardownTour(false);
  };

  if (document.readyState === 'complete' || document.readyState === 'interactive'){
    showWelcome();
  }else{
    document.addEventListener('DOMContentLoaded', showWelcome);
  }
})();
