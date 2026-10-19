/**
 * ULN to DOCX Studio — Frontend Client Application Logic
 * Flat Minimal Light Architecture (Zero-Borders, Zero-Emojis, Reactive Localhost API)
 */

document.addEventListener('DOMContentLoaded', () => {
  // ── State Management ───────────────────────────────────────────
  const state = {
    images: new Map(), // key: lowercase filename, value: { filename, thumbnail, temp_path, size }
    compilePollingInterval: null,
    searchMatches: [],
    currentMatchIndex: -1,
    // File Manager State
    fmFolder: null,
    fmCategory: 'all',
    fmFiles: [],
    fmSelectedFile: null,
    outputFolder: '',
    documentsFolder: '',
  };

  // ── Heartbeat & Auto-Shutdown on Tab Close ─────────────────────
  // Send initial heartbeat immediately on load, then every 3 seconds
  fetch('/api/heartbeat', { method: 'POST' }).catch(() => {});
  setInterval(() => {
    fetch('/api/heartbeat', { method: 'POST' }).catch(() => {});
  }, 3000);

  // Send shutdown beacon ONLY when tab or window is explicitly unloaded
  window.addEventListener('beforeunload', () => {
    navigator.sendBeacon('/api/shutdown');
  });

  // ── DOM References ─────────────────────────────────────────────
  const el = {
    versionBadge: document.getElementById('versionBadge'),
    btnPromptModal: document.getElementById('btnPromptModal'),
    btnExitApp: document.getElementById('btnExitApp'),

    // View Navigation Tabs
    tabEditor: document.getElementById('tabEditor'),
    tabFiles: document.getElementById('tabFiles'),
    viewEditor: document.getElementById('viewEditor'),
    viewFiles: document.getElementById('viewFiles'),

    // Settings
    fontFamily: document.getElementById('fontFamily'),
    fontSize: document.getElementById('fontSize'),
    marginTop: document.getElementById('marginTop'),
    marginBottom: document.getElementById('marginBottom'),
    marginLeft: document.getElementById('marginLeft'),
    marginRight: document.getElementById('marginRight'),
    qPrefix: document.getElementById('qPrefix'),
    qDelim: document.getElementById('qDelim'),
    qColor: document.getElementById('qColor'),
    optColor: document.getElementById('optColor'),
    enablePageNumbers: document.getElementById('enablePageNumbers'),
    openWordAfterCompile: document.getElementById('openWordAfterCompile'),
    backgroundMode: document.getElementById('backgroundMode'),

    // Images
    imageCountPill: document.getElementById('imageCountPill'),
    imageDropzone: document.getElementById('imageDropzone'),
    fileInputImages: document.getElementById('fileInputImages'),
    queueList: document.getElementById('queueList'),
    btnSortAZ: document.getElementById('btnSortAZ'),
    btnMatchPics: document.getElementById('btnMatchPics'),
    btnClearImages: document.getElementById('btnClearImages'),

    // Toolbar & Editor
    btnSaveUlnDirect: document.getElementById('btnSaveUlnDirect'),
    btnEmbedRaw: document.getElementById('btnEmbedRaw'),
    btnOpenOutputFolder: document.getElementById('btnOpenOutputFolder'),
    btnImportTxt: document.getElementById('btnImportTxt'),
    btnLoadSample: document.getElementById('btnLoadSample'),
    btnClearText: document.getElementById('btnClearText'),
    btnCompile: document.getElementById('btnCompile'),
    ulnEditor: document.getElementById('ulnEditor'),
    lineNumbers: document.getElementById('lineNumbers'),

    // Status bar
    statusMessage: document.getElementById('statusMessage'),
    progressWrap: document.getElementById('progressWrap'),
    progressFill: document.getElementById('progressFill'),
    btnCancelCompile: document.getElementById('btnCancelCompile'),
    statLines: document.getElementById('statLines'),
    statWords: document.getElementById('statWords'),
    statChars: document.getElementById('statChars'),

    // Search bar (Keyboard Ctrl+F)
    searchBar: document.getElementById('searchBar'),
    searchInput: document.getElementById('searchInput'),
    matchCount: document.getElementById('matchCount'),
    btnPrevMatch: document.getElementById('btnPrevMatch'),
    btnNextMatch: document.getElementById('btnNextMatch'),
    matchCase: document.getElementById('matchCase'),
    btnToggleReplace: document.getElementById('btnToggleReplace'),
    btnCloseSearch: document.getElementById('btnCloseSearch'),
    replaceRow: document.getElementById('replaceRow'),
    replaceInput: document.getElementById('replaceInput'),
    btnReplaceOne: document.getElementById('btnReplaceOne'),
    btnReplaceAll: document.getElementById('btnReplaceAll'),

    // File Manager View Elements
    fmFolderSelect: document.getElementById('fmFolderSelect'),
    customFolderWrap: document.getElementById('customFolderWrap'),
    fmCustomFolderPath: document.getElementById('fmCustomFolderPath'),
    btnApplyCustomFolder: document.getElementById('btnApplyCustomFolder'),
    btnFmOpenFolder: document.getElementById('btnFmOpenFolder'),
    btnFmRefresh: document.getElementById('btnFmRefresh'),
    btnFmSaveCurrent: document.getElementById('btnFmSaveCurrent'),
    fmFileCountPill: document.getElementById('fmFileCountPill'),
    fmTableBody: document.getElementById('fmTableBody'),
    fmPreviewBadge: document.getElementById('fmPreviewBadge'),
    fmPreviewName: document.getElementById('fmPreviewName'),
    fmPreviewPath: document.getElementById('fmPreviewPath'),
    fmPreviewSize: document.getElementById('fmPreviewSize'),
    fmPreviewDate: document.getElementById('fmPreviewDate'),
    fmPreviewPre: document.getElementById('fmPreviewPre'),
    btnFmLoadToEditor: document.getElementById('btnFmLoadToEditor'),
    btnFmOpenInWord: document.getElementById('btnFmOpenInWord'),
    btnFmDeleteSelected: document.getElementById('btnFmDeleteSelected'),

    // Modals
    modalEmbedRaw: document.getElementById('modalEmbedRaw'),
    embedDocxPathInput: document.getElementById('embedDocxPathInput'),
    btnConfirmEmbedRaw: document.getElementById('btnConfirmEmbedRaw'),

    modalSaveFile: document.getElementById('modalSaveFile'),
    saveFileNameInput: document.getElementById('saveFileNameInput'),
    btnConfirmSaveFile: document.getElementById('btnConfirmSaveFile'),

    modalPrompt: document.getElementById('modalPrompt'),
    promptTextarea: document.getElementById('promptTextarea'),
    btnResetPrompt: document.getElementById('btnResetPrompt'),
    btnCopyPrompt: document.getElementById('btnCopyPrompt'),
    btnSavePrompt: document.getElementById('btnSavePrompt'),

    modalPicMatch: document.getElementById('modalPicMatch'),
    picMatchSummary: document.getElementById('picMatchSummary'),
    matchCardsContainer: document.getElementById('matchCardsContainer'),

    toastContainer: document.getElementById('toastContainer'),
  };

  // ── Modal Handling (Reliable open & close) ─────────────────────
  function openModal(modalEl) {
    if (!modalEl) return;
    modalEl.removeAttribute('hidden');
    modalEl.style.display = 'flex';
  }

  function closeModal(modalEl) {
    if (!modalEl) return;
    modalEl.setAttribute('hidden', '');
    modalEl.style.display = 'none';
  }

  // Ensure all modals start closed
  [el.modalPrompt, el.modalPicMatch, el.modalSaveFile, el.modalEmbedRaw].forEach((m) => {
    if (m) closeModal(m);
  });

  // Close handlers on [data-close] buttons
  document.querySelectorAll('[data-close]').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const targetId = btn.getAttribute('data-close');
      const targetModal = document.getElementById(targetId);
      if (targetModal) closeModal(targetModal);
    });
  });

  // Close when clicking outside modal card (on backdrop)
  document.querySelectorAll('.modal-backdrop').forEach((backdrop) => {
    backdrop.addEventListener('click', (e) => {
      if (e.target === backdrop) closeModal(backdrop);
    });
  });

  // Close top modal on Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (!el.searchBar.hasAttribute('hidden')) {
        closeSearchBar();
        return;
      }
      [el.modalPrompt, el.modalPicMatch, el.modalSaveFile, el.modalEmbedRaw].forEach((m) => {
        if (m && m.style.display === 'flex') closeModal(m);
      });
    }
  });

  // ── Exit App Button ────────────────────────────────────────────
  el.btnExitApp.addEventListener('click', async () => {
    if (confirm('Bạn có chắc chắn muốn thoát ứng dụng và dừng máy chủ trên terminal?')) {
      showToast('Đang tắt máy chủ...', 'info');
      try {
        await fetch('/api/shutdown', { method: 'POST' });
      } catch (err) {}
      document.body.innerHTML = `
        <div style="height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;font-family:sans-serif;color:#64748b;background:#f8fafc;">
          <h2 style="color:#0f172a;margin-bottom:8px;font-weight:600;">Máy chủ đã tắt thành công</h2>
          <p style="font-size:14px;">Bạn có thể đóng tab trình duyệt này.</p>
        </div>
      `;
      setTimeout(() => window.close(), 1200);
    }
  });

  // ── View Switching (Editor vs File Manager) ─────────────────────
  el.tabEditor.addEventListener('click', () => switchView('viewEditor'));
  el.tabFiles.addEventListener('click', () => {
    switchView('viewFiles');
    loadFileManagement();
  });

  function switchView(viewId) {
    if (viewId === 'viewEditor') {
      el.tabEditor.classList.add('active');
      el.tabFiles.classList.remove('active');
      el.viewEditor.removeAttribute('hidden');
      el.viewEditor.style.display = 'flex';
      el.viewFiles.setAttribute('hidden', '');
      el.viewFiles.style.display = 'none';
    } else {
      el.tabFiles.classList.add('active');
      el.tabEditor.classList.remove('active');
      el.viewFiles.removeAttribute('hidden');
      el.viewFiles.style.display = 'flex';
      el.viewEditor.setAttribute('hidden', '');
      el.viewEditor.style.display = 'none';
    }
  }

  // ── Toast Notification Helper (Zero Emojis) ────────────────────
  function showToast(message, type = 'info', action = null) {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const labelMap = {
      success: 'OK',
      error: 'LỖI',
      info: 'TIN',
      warning: 'Ý',
    };

    const badge = document.createElement('div');
    badge.className = 'toast-badge';
    badge.textContent = labelMap[type] || 'TIN';

    const msg = document.createElement('div');
    msg.className = 'toast-msg';
    msg.textContent = message;

    toast.appendChild(badge);
    toast.appendChild(msg);

    if (action && action.label && action.onClick) {
      const btn = document.createElement('button');
      btn.className = 'toast-action';
      btn.textContent = action.label;
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        action.onClick();
        toast.remove();
      });
      toast.appendChild(btn);
    }

    el.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      toast.style.transition = 'opacity 0.2s ease, transform 0.2s ease';
      setTimeout(() => toast.remove(), 200);
    }, 4500);
  }

  // ── Initialize App & Environment ───────────────────────────────
  async function initApp() {
    try {
      const res = await fetch('/api/status');
      if (res.ok) {
        const data = await res.json();
        if (data.version && el.versionBadge) {
          el.versionBadge.textContent = `v${data.version}`;
        }
        state.outputFolder = data.output_folder || '';
        state.documentsFolder = data.default_folder || '';
        state.fmFolder = state.outputFolder;
      }
    } catch (err) {
      console.warn('API status error:', err);
    }

    if (!el.ulnEditor.value.trim()) {
      await loadSampleUln();
    } else {
      updateEditorStats();
      syncLineNumbers();
    }
  }

  // ── Sample ULN Loader ──────────────────────────────────────────
  async function loadSampleUln() {
    try {
      const res = await fetch('/api/sample');
      if (res.ok) {
        const data = await res.json();
        el.ulnEditor.value = data.sample_text || data.content || '';
        updateEditorStats();
        syncLineNumbers();
        showToast('Đã nạp văn bản mẫu ULN', 'info');
      }
    } catch (err) {
      console.error('Failed to load sample:', err);
    }
  }

  // ── Editor Line Numbers & Statistics ───────────────────────────
  function syncLineNumbers() {
    const lines = el.ulnEditor.value.split('\n');
    const lineCount = lines.length;
    let numbersHtml = '';
    for (let i = 1; i <= lineCount; i++) {
      numbersHtml += `${i}<br>`;
    }
    el.lineNumbers.innerHTML = numbersHtml;
  }

  function updateEditorStats() {
    const text = el.ulnEditor.value;
    const lines = text.split('\n').length;
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;

    el.statLines.textContent = `${lines.toLocaleString()} dòng`;
    el.statWords.textContent = `${words.toLocaleString()} từ`;
    el.statChars.textContent = `${chars.toLocaleString()} ký tự`;
  }

  el.ulnEditor.addEventListener('scroll', () => {
    el.lineNumbers.scrollTop = el.ulnEditor.scrollTop;
  });

  let statsDebounceTimer = null;
  el.ulnEditor.addEventListener('input', () => {
    syncLineNumbers();
    clearTimeout(statsDebounceTimer);
    statsDebounceTimer = setTimeout(updateEditorStats, 150);
  });

  el.ulnEditor.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      const start = el.ulnEditor.selectionStart;
      const end = el.ulnEditor.selectionEnd;
      el.ulnEditor.value = el.ulnEditor.value.substring(0, start) + '\t' + el.ulnEditor.value.substring(end);
      el.ulnEditor.selectionStart = el.ulnEditor.selectionEnd = start + 1;
      syncLineNumbers();
    }
  });

  // ── Heading Shortcuts ──────────────────────────────────────────
  document.querySelectorAll('.btn-h').forEach((btn) => {
    btn.addEventListener('click', () => {
      const level = btn.getAttribute('data-level');
      const tag = `[H${level}] `;
      insertAtLineStart(tag);
    });
  });

  function insertAtLineStart(prefix) {
    const textarea = el.ulnEditor;
    const pos = textarea.selectionStart;
    const text = textarea.value;
    const lineStart = text.lastIndexOf('\n', pos - 1) + 1;

    textarea.value = text.slice(0, lineStart) + prefix + text.slice(lineStart);
    textarea.selectionStart = textarea.selectionEnd = pos + prefix.length;
    textarea.focus();
    syncLineNumbers();
    updateEditorStats();
  }

  // ── Clear Text ─────────────────────────────────────────────────
  el.btnClearText.addEventListener('click', () => {
    if (!el.ulnEditor.value.trim()) return;
    if (confirm('Bạn có chắc chắn muốn xóa toàn bộ nội dung trong trình soạn thảo?')) {
      el.ulnEditor.value = '';
      syncLineNumbers();
      updateEditorStats();
      showToast('Đã xóa trình soạn thảo', 'info');
    }
  });

  el.btnLoadSample.addEventListener('click', loadSampleUln);

  // ── Save Current ULN Directly to File ──────────────────────────
  el.btnSaveUlnDirect.addEventListener('click', openSaveFileDialog);
  el.btnFmSaveCurrent.addEventListener('click', openSaveFileDialog);

  function openSaveFileDialog() {
    const text = el.ulnEditor.value.trim();
    if (!text) {
      showToast('Nội dung trình soạn thảo đang rỗng, không có gì để lưu!', 'warning');
      return;
    }
    const defaultName = `de_thi_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.uln`;
    el.saveFileNameInput.value = defaultName;
    openModal(el.modalSaveFile);
    setTimeout(() => el.saveFileNameInput.focus(), 100);
  }

  el.btnConfirmSaveFile.addEventListener('click', async () => {
    const filename = el.saveFileNameInput.value.trim();
    if (!filename) {
      showToast('Vui lòng nhập tên tệp!', 'warning');
      return;
    }
    const text = el.ulnEditor.value;
    try {
      const res = await fetch('/api/files/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename, content: text, folder: state.fmFolder }),
      });
      if (res.ok) {
        const data = await res.json();
        closeModal(el.modalSaveFile);
        showToast(`Đã lưu tệp: ${data.filename}`, 'success');
        if (el.viewFiles.style.display === 'flex') {
          loadFileManagement();
        }
      } else {
        const err = await res.json();
        showToast(`Lỗi lưu tệp: ${err.detail}`, 'error');
      }
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  });

  // ── Import .txt File ───────────────────────────────────────────
  el.btnImportTxt.addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.txt,.uln';
    input.onchange = (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        el.ulnEditor.value = evt.target.result;
        syncLineNumbers();
        updateEditorStats();
        showToast(`Đã nạp tệp: ${file.name}`, 'success');
      };
      reader.readAsText(file, 'UTF-8');
    };
    input.click();
  });

  // ── Image Queue & Upload ───────────────────────────────────────
  el.imageDropzone.addEventListener('click', () => el.fileInputImages.click());

  el.imageDropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.imageDropzone.classList.add('drag-active');
  });

  el.imageDropzone.addEventListener('dragleave', () => {
    el.imageDropzone.classList.remove('drag-active');
  });

  el.imageDropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    el.imageDropzone.classList.remove('drag-active');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleImageFiles(e.dataTransfer.files);
    }
  });

  el.fileInputImages.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleImageFiles(e.target.files);
    }
  });

  async function handleImageFiles(fileList) {
    const formData = new FormData();
    let count = 0;
    for (let i = 0; i < fileList.length; i++) {
      formData.append('files', fileList[i]);
      count++;
    }

    if (count === 0) return;

    showToast(`Đang tải lên và xử lý ${count} ảnh...`, 'info');

    try {
      const res = await fetch('/api/upload-images', {
        method: 'POST',
        body: formData,
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Lỗi khi tải ảnh');
      }

      const data = await res.json();
      if (data.uploaded) {
        data.uploaded.forEach((item) => {
          state.images.set(item.filename.toLowerCase(), item);
        });
        renderImageQueue();
        showToast(`Đã thêm ${data.uploaded.length} ảnh vào hàng đợi`, 'success');
      }
    } catch (err) {
      showToast(`Tải ảnh thất bại: ${err.message}`, 'error');
    } finally {
      el.fileInputImages.value = '';
    }
  }

  function renderImageQueue() {
    el.queueList.innerHTML = '';
    const imageList = Array.from(state.images.values());
    el.imageCountPill.textContent = `${imageList.length} ảnh`;

    if (imageList.length === 0) {
      el.queueList.innerHTML = '<div class="empty-state">Chưa có ảnh nào trong hàng đợi</div>';
      return;
    }

    imageList.forEach((img) => {
      const row = document.createElement('div');
      row.className = 'queue-item';

      const thumb = document.createElement('img');
      thumb.className = 'queue-thumb';
      thumb.src = img.thumbnail;
      thumb.alt = img.filename;

      const details = document.createElement('div');
      details.className = 'queue-details';

      const name = document.createElement('span');
      name.className = 'queue-name';
      name.textContent = img.filename;
      name.title = img.filename;

      const size = document.createElement('span');
      size.className = 'queue-size';
      size.textContent = `${(img.size / 1024).toFixed(1)} KB`;

      details.appendChild(name);
      details.appendChild(size);

      const delBtn = document.createElement('button');
      delBtn.className = 'btn-queue-del';
      delBtn.textContent = 'Xóa';
      delBtn.title = 'Xóa ảnh này';
      delBtn.addEventListener('click', () => {
        state.images.delete(img.filename.toLowerCase());
        renderImageQueue();
      });

      row.appendChild(thumb);
      row.appendChild(details);
      row.appendChild(delBtn);

      el.queueList.appendChild(row);
    });
  }

  el.btnSortAZ.addEventListener('click', () => {
    const sorted = Array.from(state.images.entries()).sort((a, b) => a[0].localeCompare(b[0]));
    state.images = new Map(sorted);
    renderImageQueue();
    showToast('Đã sắp xếp danh sách ảnh A-Z', 'info');
  });

  el.btnClearImages.addEventListener('click', () => {
    if (state.images.size === 0) return;
    if (confirm('Xóa toàn bộ ảnh trong hàng đợi?')) {
      state.images.clear();
      renderImageQueue();
      showToast('Đã xóa toàn bộ hàng đợi ảnh', 'info');
    }
  });

  // ── Exercise & Picture Matcher Cross-Check ───────────────────────
  el.btnMatchPics.addEventListener('click', () => {
    const text = el.ulnEditor.value;
    const lines = text.split('\n');
    const matches = [];

    let currentQuestion = '';
    const picRegex = /\[PIC:\s*([^\]]+)\]/i;

    lines.forEach((line) => {
      const trimmed = line.trim();
      if (/^(Câu|\d+|Question|Task|Exercise)\s*[\d\.\:]/i.test(trimmed)) {
        currentQuestion = trimmed.slice(0, 70);
      }
      const match = picRegex.exec(line);
      if (match) {
        matches.push({
          tag: match[0],
          filename: match[1].trim(),
          question: currentQuestion || 'Nội dung chung',
          line: trimmed,
        });
      }
    });

    el.matchCardsContainer.innerHTML = '';
    el.picMatchSummary.textContent = `Tìm thấy ${matches.length} thẻ [PIC] trong văn bản.`;

    if (matches.length === 0) {
      el.matchCardsContainer.innerHTML = '<div class="empty-state">Không tìm thấy thẻ [PIC] nào trong văn bản soạn thảo.</div>';
    } else {
      matches.forEach((item, idx) => {
        const foundImg = state.images.get(item.filename.toLowerCase());
        const card = document.createElement('div');
        card.className = 'match-card';

        const info = document.createElement('div');
        info.className = 'match-info';

        const headerRow = document.createElement('div');
        headerRow.className = 'match-header-row';

        const qnum = document.createElement('span');
        qnum.className = 'match-qnum';
        qnum.textContent = `#${idx + 1} ${item.question}`;

        const tag = document.createElement('span');
        tag.className = `match-tag ${foundImg ? '' : 'missing'}`;
        tag.textContent = foundImg ? `Có sẵn: ${item.filename}` : `Chưa có: ${item.filename}`;

        headerRow.appendChild(qnum);
        headerRow.appendChild(tag);

        const snippet = document.createElement('div');
        snippet.className = 'match-snippet';
        snippet.textContent = item.line;

        info.appendChild(headerRow);
        info.appendChild(snippet);

        const previewBox = document.createElement('div');
        previewBox.className = 'match-preview-box';

        if (foundImg) {
          const imgEl = document.createElement('img');
          imgEl.className = 'match-img';
          imgEl.src = foundImg.thumbnail;
          imgEl.alt = item.filename;
          previewBox.appendChild(imgEl);
        } else {
          const missingMsg = document.createElement('div');
          missingMsg.className = 'match-img-placeholder';
          missingMsg.textContent = 'Chưa có ảnh';
          previewBox.appendChild(missingMsg);
        }

        card.appendChild(info);
        card.appendChild(previewBox);
        el.matchCardsContainer.appendChild(card);
      });
    }

    openModal(el.modalPicMatch);
  });

  // ── Find & Replace Floating Bar (Ctrl+F Keyboard Shortcut) ───────
  el.btnCloseSearch.addEventListener('click', closeSearchBar);

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault();
      toggleSearchBar();
    }
  });

  function toggleSearchBar() {
    if (el.searchBar.hasAttribute('hidden')) {
      el.searchBar.removeAttribute('hidden');
      el.searchBar.style.display = 'flex';
      el.searchInput.focus();
      el.searchInput.select();
      performSearch();
    } else {
      closeSearchBar();
    }
  }

  function closeSearchBar() {
    el.searchBar.setAttribute('hidden', '');
    el.searchBar.style.display = 'none';
    el.ulnEditor.focus();
  }

  el.btnToggleReplace.addEventListener('click', () => {
    if (el.replaceRow.hasAttribute('hidden')) {
      el.replaceRow.removeAttribute('hidden');
      el.replaceRow.style.display = 'flex';
      el.replaceInput.focus();
    } else {
      el.replaceRow.setAttribute('hidden', '');
      el.replaceRow.style.display = 'none';
    }
  });

  el.searchInput.addEventListener('input', performSearch);
  el.matchCase.addEventListener('change', performSearch);

  function performSearch() {
    const query = el.searchInput.value;
    const text = el.ulnEditor.value;
    state.searchMatches = [];
    state.currentMatchIndex = -1;

    if (!query) {
      el.matchCount.textContent = '0 kết quả';
      return;
    }

    const isCase = el.matchCase.checked;
    const haystack = isCase ? text : text.toLowerCase();
    const needle = isCase ? query : query.toLowerCase();

    let startIndex = 0;
    while (startIndex < haystack.length) {
      const idx = haystack.indexOf(needle, startIndex);
      if (idx === -1) break;
      state.searchMatches.push({ start: idx, end: idx + needle.length });
      startIndex = idx + needle.length;
    }

    if (state.searchMatches.length > 0) {
      state.currentMatchIndex = 0;
      el.matchCount.textContent = `1 / ${state.searchMatches.length}`;
      highlightCurrentMatch();
    } else {
      el.matchCount.textContent = '0 kết quả';
    }
  }

  function highlightCurrentMatch() {
    if (state.currentMatchIndex < 0 || state.currentMatchIndex >= state.searchMatches.length) return;
    const match = state.searchMatches[state.currentMatchIndex];
    el.ulnEditor.focus();
    el.ulnEditor.setSelectionRange(match.start, match.end);
    el.matchCount.textContent = `${state.currentMatchIndex + 1} / ${state.searchMatches.length}`;
  }

  el.btnNextMatch.addEventListener('click', () => {
    if (state.searchMatches.length === 0) return;
    state.currentMatchIndex = (state.currentMatchIndex + 1) % state.searchMatches.length;
    highlightCurrentMatch();
  });

  el.btnPrevMatch.addEventListener('click', () => {
    if (state.searchMatches.length === 0) return;
    state.currentMatchIndex = (state.currentMatchIndex - 1 + state.searchMatches.length) % state.searchMatches.length;
    highlightCurrentMatch();
  });

  el.searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (e.shiftKey) {
        el.btnPrevMatch.click();
      } else {
        el.btnNextMatch.click();
      }
    }
  });

  el.btnReplaceOne.addEventListener('click', () => {
    if (state.currentMatchIndex < 0 || state.currentMatchIndex >= state.searchMatches.length) return;
    const match = state.searchMatches[state.currentMatchIndex];
    const replacement = el.replaceInput.value;
    const text = el.ulnEditor.value;

    el.ulnEditor.value = text.slice(0, match.start) + replacement + text.slice(match.end);
    performSearch();
    syncLineNumbers();
    updateEditorStats();
  });

  el.btnReplaceAll.addEventListener('click', () => {
    const query = el.searchInput.value;
    if (!query) return;
    const replacement = el.replaceInput.value;
    const isCase = el.matchCase.checked;

    let count = 0;
    if (isCase) {
      count = el.ulnEditor.value.split(query).length - 1;
      el.ulnEditor.value = el.ulnEditor.value.replaceAll(query, replacement);
    } else {
      const regex = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
      count = (el.ulnEditor.value.match(regex) || []).length;
      el.ulnEditor.value = el.ulnEditor.value.replace(regex, replacement);
    }

    performSearch();
    syncLineNumbers();
    updateEditorStats();
    showToast(`Đã thay thế ${count} vị trí`, 'info');
  });

  // ── AI Prompt Modal ────────────────────────────────────────────
  el.btnPromptModal.addEventListener('click', async () => {
    try {
      const res = await fetch('/api/prompt');
      if (res.ok) {
        const data = await res.json();
        el.promptTextarea.value = data.prompt_text || data.prompt || '';
      }
    } catch (err) {
      console.error('Failed to load prompt:', err);
    }
    openModal(el.modalPrompt);
  });

  el.btnCopyPrompt.addEventListener('click', () => {
    navigator.clipboard.writeText(el.promptTextarea.value).then(() => {
      showToast('Đã sao chép prompt vào Clipboard!', 'success');
    });
  });

  el.btnSavePrompt.addEventListener('click', async () => {
    try {
      const res = await fetch('/api/prompt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt_text: el.promptTextarea.value }),
      });
      if (res.ok) {
        showToast('Đã lưu quy tắc prompt thành công', 'success');
        closeModal(el.modalPrompt);
      }
    } catch (err) {
      showToast('Lỗi khi lưu prompt: ' + err.message, 'error');
    }
  });

  el.btnResetPrompt.addEventListener('click', async () => {
    if (confirm('Khôi phục prompt về phiên bản mặc định của hệ thống?')) {
      try {
        const res = await fetch('/api/prompt');
        if (res.ok) {
          const data = await res.json();
          el.promptTextarea.value = data.default_prompt || data.prompt || '';
          showToast('Đã khôi phục prompt mặc định', 'info');
        }
      } catch (err) {
        console.error(err);
      }
    }
  });

  // ── Embed Raw ULN to Existing DOCX ─────────────────────────────
  el.btnEmbedRaw.addEventListener('click', () => {
    const text = el.ulnEditor.value.trim();
    if (!text) {
      showToast('Trình soạn thảo rỗng, không có gì để nhúng!', 'warning');
      return;
    }
    el.embedDocxPathInput.value = '';
    openModal(el.modalEmbedRaw);
    setTimeout(() => el.embedDocxPathInput.focus(), 100);
  });

  async function handleConfirmEmbedRaw() {
    let rawPath = (el.embedDocxPathInput.value || '').trim();
    if (!rawPath) {
      showToast('Vui lòng nhập đường dẫn file DOCX!', 'warning');
      return;
    }
    rawPath = rawPath.replace(/^["']|["']$/g, '').trim();
    const text = el.ulnEditor.value.trim();

    showToast('Đang nhúng mã ULN vào tệp Word...', 'info');

    try {
      const res = await fetch('/api/embed-uln', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filepath: rawPath, uln_text: text }),
      });

      if (res.ok) {
        const data = await res.json();
        closeModal(el.modalEmbedRaw);
        showToast('Đã nhúng mã ULN thành công vào tệp Word!', 'success', {
          label: 'Mở Word',
          onClick: () => openCompiledDocx(data.filepath),
        });
      } else {
        const err = await res.json().catch(() => ({ detail: res.statusText }));
        showToast(`Nhúng thất bại: ${err.detail || res.statusText}`, 'error');
      }
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  }

  el.btnConfirmEmbedRaw.addEventListener('click', handleConfirmEmbedRaw);
  el.embedDocxPathInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      handleConfirmEmbedRaw();
    }
  });

  if (el.btnOpenOutputFolder) {
    el.btnOpenOutputFolder.addEventListener('click', async () => {
      try {
        const res = await fetch('/api/files/open-folder', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ folder: state.outputFolder || '' }),
        });
        if (res.ok) {
          showToast('Đã mở thư mục output trong Windows Explorer', 'success');
        } else {
          showToast('Không thể mở thư mục output', 'error');
        }
      } catch (err) {
        showToast('Không thể kết nối máy chủ: ' + err.message, 'error');
      }
    });
  }

  // ── Compilation Orchestration ──────────────────────────────────
  el.btnCompile.addEventListener('click', startCompilation);
  el.btnCancelCompile.addEventListener('click', cancelCompilation);

  async function startCompilation() {
    const text = el.ulnEditor.value.trim();
    if (!text) {
      showToast('Vui lòng nhập mã ULN trước khi xuất tài liệu!', 'warning');
      return;
    }

    const settings = {
      font_family: el.fontFamily.value,
      font_size: parseFloat(el.fontSize.value) || 12.0,
      margin_top: parseFloat(el.marginTop.value) || 2.0,
      margin_bottom: parseFloat(el.marginBottom.value) || 2.0,
      margin_left: parseFloat(el.marginLeft.value) || 3.0,
      margin_right: parseFloat(el.marginRight.value) || 1.5,
      question_prefix: el.qPrefix.value,
      question_delim: el.qDelim.value,
      question_color: el.qColor.value,
      option_color: el.optColor.value,
      enable_page_numbers: el.enablePageNumbers.checked,
    };

    const payload = {
      uln_text: text,
      settings: settings,
      open_word: el.openWordAfterCompile.checked,
      background_mode: el.backgroundMode.checked,
    };

    el.btnCompile.disabled = true;
    el.progressWrap.removeAttribute('hidden');
    el.progressWrap.style.display = 'flex';
    el.progressFill.style.width = '5%';
    el.statusMessage.textContent = 'Đang khởi tạo tài liệu...';

    try {
      const res = await fetch('/api/compile', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Không thể bắt đầu biên dịch');
      }

      clearInterval(state.compilePollingInterval);
      state.compilePollingInterval = setInterval(pollCompileStatus, 300);
    } catch (err) {
      resetCompileUI();
      showToast(`Lỗi biên dịch: ${err.message}`, 'error');
    }
  }

  async function pollCompileStatus() {
    try {
      const res = await fetch('/api/compile/status');
      if (!res.ok) return;

      const data = await res.json();
      const pct = Math.round((data.progress || 0) * 100);
      el.progressFill.style.width = `${pct}%`;
      el.statusMessage.textContent = data.message || 'Đang xử lý...';

      if (!data.is_compiling) {
        clearInterval(state.compilePollingInterval);

        if (data.error) {
          resetCompileUI();
          showToast(`Biên dịch thất bại: ${data.error}`, 'error');
        } else if (data.output_path) {
          el.progressFill.style.width = '100%';
          el.statusMessage.textContent = 'Xuất tài liệu Word hoàn tất!';
          showToast(`Đã xuất thành công: ${data.output_path}`, 'success', {
            label: 'Mở Word',
            onClick: () => openCompiledDocx(data.output_path),
          });
          setTimeout(resetCompileUI, 3000);
        } else {
          resetCompileUI();
        }
      }
    } catch (err) {
      console.warn('Poll error:', err);
    }
  }

  async function cancelCompilation() {
    try {
      await fetch('/api/compile/cancel', { method: 'POST' });
      showToast('Đang hủy quá trình biên dịch...', 'info');
    } catch (err) {
      console.error(err);
    }
  }

  function resetCompileUI() {
    clearInterval(state.compilePollingInterval);
    el.btnCompile.disabled = false;
    el.progressWrap.setAttribute('hidden', '');
    el.progressWrap.style.display = 'none';
    el.progressFill.style.width = '0%';
    el.statusMessage.textContent = 'Sẵn sàng';
  }

  async function openCompiledDocx(filepath) {
    try {
      await fetch('/api/open-docx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filepath }),
      });
    } catch (err) {
      showToast('Không thể mở tệp trong Word: ' + err.message, 'error');
    }
  }

  // ── Dedicated File Management Logic ────────────────────────────
  el.fmFolderSelect.addEventListener('change', () => {
    const val = el.fmFolderSelect.value;
    if (val === 'output') {
      state.fmFolder = state.outputFolder;
      el.customFolderWrap.setAttribute('hidden', '');
      loadFileManagement();
    } else if (val === 'documents') {
      state.fmFolder = state.documentsFolder;
      el.customFolderWrap.setAttribute('hidden', '');
      loadFileManagement();
    } else {
      el.customFolderWrap.removeAttribute('hidden');
      if (state.fmFolder && state.fmFolder !== state.outputFolder && state.fmFolder !== state.documentsFolder) {
        el.fmCustomFolderPath.value = state.fmFolder;
      }
      el.fmCustomFolderPath.focus();
    }
  });

  function applyCustomFolder() {
    let custom = (el.fmCustomFolderPath.value || '').trim();
    custom = custom.replace(/^["']|["']$/g, '').trim();
    if (!custom) {
      showToast('Vui lòng nhập đường dẫn thư mục!', 'warning');
      return;
    }
    state.fmFolder = custom;
    loadFileManagement();
  }

  el.btnApplyCustomFolder.addEventListener('click', applyCustomFolder);
  el.fmCustomFolderPath.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      applyCustomFolder();
    }
  });

  document.querySelectorAll('.fm-filter-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.fm-filter-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.fmCategory = btn.getAttribute('data-category') || 'all';
      loadFileManagement();
    });
  });

  el.btnFmRefresh.addEventListener('click', loadFileManagement);

  el.btnFmOpenFolder.addEventListener('click', async () => {
    let targetFolder = state.fmFolder;
    if (el.fmFolderSelect && el.fmFolderSelect.value === 'custom' && el.fmCustomFolderPath && el.fmCustomFolderPath.value.trim()) {
      targetFolder = el.fmCustomFolderPath.value.trim().replace(/^["']|["']$/g, '').trim();
    }
    if (!targetFolder) {
      targetFolder = state.outputFolder || '';
    }

    const payload = { folder: targetFolder };
    if (state.fmSelectedFile && state.fmSelectedFile.path) {
      payload.filepath = state.fmSelectedFile.path;
    }

    try {
      const res = await fetch('/api/files/open-folder', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      if (res.ok) {
        showToast('Đang mở thư mục trong Windows Explorer...', 'success');
      } else {
        const err = await res.json().catch(() => ({ detail: res.statusText }));
        showToast('Không thể mở thư mục: ' + (err.detail || res.statusText), 'error');
      }
    } catch (err) {
      showToast('Không thể kết nối máy chủ: ' + err.message, 'error');
    }
  });

  async function loadFileManagement() {
    el.fmTableBody.innerHTML = '<tr><td colspan="6" class="text-center">Đang tải danh sách tệp...</td></tr>';
    resetFmPreview();

    try {
      const folderParam = state.fmFolder || '';
      const url = `/api/files?folder=${encodeURIComponent(folderParam)}&category=${encodeURIComponent(state.fmCategory)}`;
      const res = await fetch(url);
      if (!res.ok) {
        const err = await res.json().catch(() => ({ detail: 'Không thể đọc thư mục' }));
        throw new Error(err.detail || 'Không thể đọc danh sách tệp');
      }

      const data = await res.json();
      state.fmFiles = data.files || [];
      state.fmFolder = data.folder || state.fmFolder;
      if (el.fmCustomFolderPath && state.fmFolder) {
        el.fmCustomFolderPath.value = state.fmFolder;
      }
      el.fmFileCountPill.textContent = `${state.fmFiles.length} tệp`;

      renderFileManagerTable();
    } catch (err) {
      el.fmTableBody.innerHTML = `<tr><td colspan="6" class="text-center" style="color:#ef4444">${err.message}</td></tr>`;
      showToast(err.message, 'error');
    }
  }

  function renderFileManagerTable() {
    el.fmTableBody.innerHTML = '';

    if (state.fmFiles.length === 0) {
      el.fmTableBody.innerHTML = '<tr><td colspan="6" class="text-center">Không có tệp nào trong thư mục này.</td></tr>';
      return;
    }

    state.fmFiles.forEach((file) => {
      const tr = document.createElement('tr');

      // Ext badge
      const tdExt = document.createElement('td');
      const extName = (file.ext || '').replace('.', '');
      const badgeClass = extName === 'docx' ? 'badge-docx' : (extName === 'uln' ? 'badge-uln' : 'badge-txt');
      tdExt.innerHTML = `<span class="badge-ext ${badgeClass}">${extName}</span>`;

      // Name
      const tdName = document.createElement('td');
      tdName.textContent = file.name;
      tdName.title = file.path;

      // Size
      const tdSize = document.createElement('td');
      tdSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;

      // Date
      const tdDate = document.createElement('td');
      tdDate.textContent = file.mtime_formatted;

      // Embedded tag
      const tdEmbed = document.createElement('td');
      if (file.is_docx) {
        tdEmbed.innerHTML = file.has_embedded_uln
          ? '<span class="tag-embedded">Có ULN</span>'
          : '<span class="tag-none">Không</span>';
      } else {
        tdEmbed.innerHTML = '<span class="tag-embedded">Mã nguồn</span>';
      }

      // Actions
      const tdActions = document.createElement('td');
      tdActions.className = 'fm-table-actions';

      if (file.is_docx) {
        const btnWord = document.createElement('button');
        btnWord.className = 'btn btn-xs btn-outline';
        btnWord.textContent = 'Mở Word';
        btnWord.addEventListener('click', (e) => {
          e.stopPropagation();
          openCompiledDocx(file.path);
        });
        tdActions.appendChild(btnWord);
      }

      if (file.is_text || (file.is_docx && file.has_embedded_uln)) {
        const btnLoad = document.createElement('button');
        btnLoad.className = 'btn btn-xs btn-outline';
        btnLoad.textContent = 'Nạp';
        btnLoad.title = 'Nạp vào Trình Soạn Thảo';
        btnLoad.addEventListener('click', async (e) => {
          e.stopPropagation();
          await loadFileToEditor(file);
        });
        tdActions.appendChild(btnLoad);
      }

      const btnDel = document.createElement('button');
      btnDel.className = 'btn btn-xs btn-danger-ghost';
      btnDel.textContent = 'Xóa';
      btnDel.addEventListener('click', async (e) => {
        e.stopPropagation();
        if (confirm(`Bạn có chắc chắn muốn xóa tệp: ${file.name}?`)) {
          await deleteManagedFile(file.path);
        }
      });
      tdActions.appendChild(btnDel);

      tr.appendChild(tdExt);
      tr.appendChild(tdName);
      tr.appendChild(tdSize);
      tr.appendChild(tdDate);
      tr.appendChild(tdEmbed);
      tr.appendChild(tdActions);

      tr.addEventListener('click', () => {
        document.querySelectorAll('#fmTableBody tr').forEach((r) => r.classList.remove('selected'));
        tr.classList.add('selected');
        selectManagedFile(file);
      });

      el.fmTableBody.appendChild(tr);
    });
  }

  async function selectManagedFile(file) {
    state.fmSelectedFile = file;
    el.fmPreviewBadge.textContent = (file.ext || '').toUpperCase();
    el.fmPreviewName.textContent = file.name;
    el.fmPreviewPath.textContent = file.path;
    el.fmPreviewSize.textContent = `${(file.size / 1024).toFixed(1)} KB (${file.size.toLocaleString()} bytes)`;
    el.fmPreviewDate.textContent = file.mtime_formatted;

    el.btnFmDeleteSelected.disabled = false;
    el.btnFmOpenInWord.disabled = !file.is_docx;

    el.fmPreviewPre.textContent = 'Đang đọc nội dung tệp...';

    try {
      const res = await fetch('/api/files/read', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filepath: file.path }),
      });

      if (res.ok) {
        const data = await res.json();
        el.fmPreviewPre.textContent = data.content || '(Tệp rỗng)';
        el.btnFmLoadToEditor.disabled = false;
      } else {
        const err = await res.json();
        el.fmPreviewPre.textContent = `[Không thể đọc xem trước]: ${err.detail}`;
        el.btnFmLoadToEditor.disabled = true;
      }
    } catch (err) {
      el.fmPreviewPre.textContent = `Lỗi: ${err.message}`;
      el.btnFmLoadToEditor.disabled = true;
    }
  }

  function resetFmPreview() {
    state.fmSelectedFile = null;
    el.fmPreviewBadge.textContent = 'Chưa chọn tệp';
    el.fmPreviewName.textContent = '—';
    el.fmPreviewPath.textContent = '—';
    el.fmPreviewSize.textContent = '—';
    el.fmPreviewDate.textContent = '—';
    el.fmPreviewPre.textContent = 'Chọn một tệp bên trái để xem trước nội dung...';
    el.btnFmLoadToEditor.disabled = true;
    el.btnFmOpenInWord.disabled = true;
    el.btnFmDeleteSelected.disabled = true;
  }

  el.btnFmLoadToEditor.addEventListener('click', async () => {
    if (state.fmSelectedFile) {
      await loadFileToEditor(state.fmSelectedFile);
    }
  });

  el.btnFmOpenInWord.addEventListener('click', () => {
    if (state.fmSelectedFile && state.fmSelectedFile.is_docx) {
      openCompiledDocx(state.fmSelectedFile.path);
    }
  });

  el.btnFmDeleteSelected.addEventListener('click', async () => {
    if (state.fmSelectedFile) {
      if (confirm(`Bạn có chắc chắn muốn xóa tệp: ${state.fmSelectedFile.name}?`)) {
        await deleteManagedFile(state.fmSelectedFile.path);
      }
    }
  });

  async function loadFileToEditor(file) {
    try {
      const res = await fetch('/api/files/read', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filepath: file.path }),
      });
      if (res.ok) {
        const data = await res.json();
        el.ulnEditor.value = data.content || '';
        syncLineNumbers();
        updateEditorStats();
        switchView('viewEditor');
        showToast(`Đã nạp nội dung từ: ${file.name}`, 'success');
      } else {
        const err = await res.json();
        showToast(`Không thể nạp tệp: ${err.detail}`, 'error');
      }
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  }

  async function deleteManagedFile(filepath) {
    try {
      const res = await fetch('/api/files/delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filepath }),
      });
      if (res.ok) {
        showToast('Đã xóa tệp thành công', 'success');
        loadFileManagement();
      } else {
        const err = await res.json();
        showToast(`Xóa thất bại: ${err.detail}`, 'error');
      }
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  }

  // ── Run App ────────────────────────────────────────────────────
  initApp();
});
