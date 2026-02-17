(function initAnimations(){
  const WELCOME_KEY = 'teacher_tools_welcome_shown';
  const TOUR_KEY = 'teacher_tools_tour_shown';
  const dropzone = window.dropzone || document.getElementById('dropzone');

  function getTourSteps(){
    return [
      {
        el: dropzone,
        title: 'Upload your data',
        desc: 'Drop CSV/XLSX files or click to pick. Multiple files appear in the FILES list.'
      },
      {
        el: document.getElementById('loadDirBtn'),
        title: 'Load a whole folder',
        desc: 'Pick a folder with CSV/XLSX files to load them all at once.'
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
        title: 'Export what you see',
        desc: 'Export visible columns/rows to CSV. XLSX is available too.'
      },
      {
        el: document.getElementById('exportXlsxBtn'),
        title: 'Export to XLSX',
        desc: 'Download the visible table as an XLSX workbook.'
      },
      {
        el: document.getElementById('resetBtn'),
        title: 'Reset filters quickly',
        desc: 'Clear search, filters, and visibility to start fresh.'
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
    ].filter(step => step.el);
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

  function startTour(){
    if (sessionStorage.getItem(TOUR_KEY)) return;
    const steps = getTourSteps();
    if (!steps.length) return;

    let idx = 0;
    const overlay = document.createElement('div');
    overlay.className = 'tour-overlay';
    const panel = document.createElement('div');
    panel.className = 'tour-panel';
    const badge = document.createElement('div');
    badge.className = 'tour-badge';
    panel.appendChild(badge);
    const titleEl = document.createElement('h4');
    const descEl = document.createElement('p');
    panel.appendChild(titleEl);
    panel.appendChild(descEl);
    const actions = document.createElement('div');
    actions.className = 'tour-actions';
    const skipBtn = document.createElement('button');
    skipBtn.className = 'btn';
    skipBtn.textContent = 'Skip';
    const backBtn = document.createElement('button');
    backBtn.className = 'btn';
    backBtn.textContent = 'Back';
    const nextBtn = document.createElement('button');
    nextBtn.className = 'btn primary';
    nextBtn.textContent = 'Next';
    actions.append(backBtn, skipBtn, nextBtn);
    panel.appendChild(actions);
    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    function cleanup(){
      overlay.remove();
    }
    function endTour(){
      cleanup();
      sessionStorage.setItem(TOUR_KEY, '1');
    }
    function go(stepIdx){
      if (stepIdx < 0 || stepIdx >= steps.length){
        endTour();
        return;
      }
      idx = stepIdx;
      const step = steps[idx];
      const rect = step.el.getBoundingClientRect();
      const panelWidth = 320;
      const padding = 12;
      let top = rect.bottom + padding;
      let left = rect.left;
      if (top + 180 > window.innerHeight) top = rect.top - 210;
      if (left + panelWidth > window.innerWidth) left = window.innerWidth - panelWidth - padding;
      if (left < padding) left = padding;
      if (top < padding) top = rect.bottom + padding;
      panel.style.top = `${top + window.scrollY}px`;
      panel.style.left = `${left + window.scrollX}px`;
      badge.textContent = String(idx + 1);
      titleEl.textContent = step.title;
      descEl.textContent = step.desc;
      backBtn.disabled = idx === 0;
      nextBtn.textContent = idx === steps.length - 1 ? 'Done' : 'Next';
      panel.classList.remove('show');
      requestAnimationFrame(() => panel.classList.add('show'));
    }

    skipBtn.addEventListener('click', endTour);
    backBtn.addEventListener('click', () => go(idx - 1));
    nextBtn.addEventListener('click', () => {
      if (idx === steps.length - 1) {
        endTour();
      } else {
        go(idx + 1);
      }
    });
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) endTour();
    });
    go(0);
  }

  function showAllTourHints(){
    const steps = getTourSteps();
    if (!steps.length) return;
    const overlay = document.createElement('div');
    overlay.className = 'tour-overlay';
    overlay.style.pointerEvents = 'none';
    const closeBtn = document.createElement('button');
    closeBtn.className = 'btn primary';
    closeBtn.textContent = 'Close hints';
    closeBtn.style.position = 'fixed';
    closeBtn.style.top = '16px';
    closeBtn.style.right = '16px';
    closeBtn.style.pointerEvents = 'auto';
    overlay.appendChild(closeBtn);

    const cleanup = () => {
      steps.forEach(s => s.el.classList.remove('tour-highlight', 'tour-highlight-all', 'tour-highlight-field', 'tour-focus-target'));
      overlay.remove();
    };

    steps.forEach((step, i) => {
      step.el.classList.add('tour-highlight', 'tour-highlight-all');
      if (['editHeadersBtn','exportCsvBtn','exportXlsxBtn','search'].includes(step.el.id)){
        step.el.classList.add('tour-highlight-field');
      }

      const panel = document.createElement('div');
      panel.className = 'tour-panel';
      panel.style.pointerEvents = 'auto';
      const badge = document.createElement('div');
      badge.className = 'tour-badge';
      badge.textContent = String(i + 1);
      const titleEl = document.createElement('h4');
      titleEl.textContent = step.title;
      const descEl = document.createElement('p');
      descEl.textContent = step.desc;
      panel.append(badge, titleEl, descEl);
      overlay.appendChild(panel);

      const rect = step.el.getBoundingClientRect();
      const panelWidth = 320;
      const padding = 28;
      let top = rect.bottom + padding;
      let left = rect.left;
      if (top + 220 > window.innerHeight) top = rect.top - 250;
      if (left + panelWidth > window.innerWidth) left = window.innerWidth - panelWidth - padding;
      if (left < padding) left = padding;
      if (top < padding) top = rect.bottom + padding;
      panel.style.top = `${top + window.scrollY}px`;
      panel.style.left = `${left + window.scrollX}px`;

      const showPanel = () => {
        panel.classList.add('show');
        overlay.classList.add('tour-focus');
        steps.forEach(s => s.el.classList.remove('tour-focus-target'));
        step.el.classList.add('tour-focus-target');
      };
      const hidePanel = () => {
        panel.classList.remove('show');
        overlay.classList.remove('tour-focus');
        step.el.classList.remove('tour-focus-target');
      };

      step.el.addEventListener('mouseenter', showPanel);
      step.el.addEventListener('mouseleave', hidePanel);
      panel.addEventListener('mouseenter', showPanel);
      panel.addEventListener('mouseleave', hidePanel);
    });

    closeBtn.addEventListener('click', cleanup);
    overlay.addEventListener('click', cleanup);
    document.body.appendChild(overlay);
  }

  // Expose hint helper for the Hints button
  window.showAllTourHints = showAllTourHints;

  function showWelcome(){
    if (sessionStorage.getItem(WELCOME_KEY)) return;
    const overlay = createWelcomeOverlay();
    requestAnimationFrame(() => overlay.classList.add('show'));
    if (dropzone) dropzone.classList.add('pulse');
    const tourBtn = document.getElementById('welcomeTourBtn');
    const skipBtn = document.getElementById('welcomeSkipBtn');
    const dismiss = (startTourAfter = false) => {
      overlay.classList.remove('show');
      setTimeout(() => overlay.remove(), 400);
      if (dropzone) dropzone.classList.remove('pulse');
      sessionStorage.setItem(WELCOME_KEY, '1');
      if (startTourAfter) {
        setTimeout(startTour, 400);
      }
    };
    tourBtn?.addEventListener('click', () => dismiss(true));
    skipBtn?.addEventListener('click', () => dismiss(false));
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) dismiss(false);
    });
  }

  if (document.readyState === 'complete' || document.readyState === 'interactive'){
    showWelcome();
  }else{
    document.addEventListener('DOMContentLoaded', showWelcome);
  }
})();
