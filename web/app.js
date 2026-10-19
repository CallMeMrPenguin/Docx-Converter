/**
 * ULN to DOCX Studio — Frontend Client Application Logic
 * Flat Minimal Light Architecture (Zero-Borders, Reactive Localhost API)
 */

document.addEventListener('DOMContentLoaded', () => {
  // ── State Management ───────────────────────────────────────────
  const state = {
    images: new Map(), // key: lowercase filename, value: { filename, thumbnail, temp_path, size }
    compilePollingInterval: null,
    searchMatches: [],
    currentMatchIndex: -1,
    selectedDocx: null,
  };

  // ── DOM References ─────────────────────────────────────────────
  const el = {
    versionBadge: document.getElementById('versionBadge'),
    connectionStatus: document.getElementById('connectionStatus'),
    btnPromptModal: document.getElementById('btnPromptModal'),
    btnCheckUpdate: document.getElementById('btnCheckUpdate'),

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
    btnImportDocx: document.getElementById('btnImportDocx'),
    btnEmbedRaw: document.getElementById('btnEmbedRaw'),
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

    // Search bar
    searchBar: document.getElementById('searchBar'),
    btnToggleSearch: document.getElementById('btnToggleSearch'),
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

    // Modals
    modalPrompt: document.getElementById('modalPrompt'),
    promptTextarea: document.getElementById('promptTextarea'),
    btnResetPrompt: document.getElementById('btnResetPrompt'),
    btnCopyPrompt: document.getElementById('btnCopyPrompt'),
    btnSavePrompt: document.getElementById('btnSavePrompt'),

    modalDocx: document.getElementById('modalDocx'),
    docxFolderInput: document.getElementById('docxFolderInput'),
    btnScanFolder: document.getElementById('btnScanFolder'),
    chkOnlyUln: document.getElementById('chkOnlyUln'),
    docxTableBody: document.getElementById('docxTableBody'),
    docxPreviewContent: document.getElementById('docxPreviewContent'),
    docxScannerStatus: document.getElementById('docxScannerStatus'),
    btnLoadDocxToEditor: document.getElementById('btnLoadDocxToEditor'),

    modalPicMatch: document.getElementById('modalPicMatch'),
    picMatchSummary: document.getElementById('picMatchSummary'),
    matchCardsContainer: document.getElementById('matchCardsContainer'),

    toastContainer: document.getElementById('toastContainer'),
  };

  // ── Toast Notification Helper ──────────────────────────────────
  function showToast(message, type = 'info', action = null) {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    const iconMap = {
      success: '✓',
      error: '✕',
      info: 'i',
      warning: '!',
    };

    const icon = document.createElement('div');
    icon.className = 'toast-icon';
    icon.textContent = iconMap[type] || 'i';

    const msg = document.createElement('div');
    msg.className = 'toast-msg';
    msg.textContent = message;

    toast.appendChild(icon);
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
        if (data.default_folder && el.docxFolderInput) {
          el.docxFolderInput.value = data.default_folder;
        }
      }
    } catch (err) {
      console.warn('API status error:', err);
      if (el.connectionStatus) {
        el.connectionStatus.innerHTML = '<span class="status-dot" style="background:#ef4444;box-shadow:none;"></span> Mất kết nối';
        el.connectionStatus.style.color = '#ef4444';
        el.connectionStatus.style.backgroundColor = '#fef2f2';
      }
    }

    // Load sample ULN initially if editor is empty
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
        el.ulnEditor.value = data.sample_text || '';
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

  // Support Tab key in editor
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
      delBtn.textContent = '✕';
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
        tag.textContent = foundImg ? `✓ ${item.filename}` : `✕ Thiếu: ${item.filename}`;

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

  // ── Find & Replace Floating Bar ────────────────────────────────
  el.btnToggleSearch.addEventListener('click', toggleSearchBar);
  el.btnCloseSearch.addEventListener('click', closeSearchBar);

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault();
      toggleSearchBar();
    } else if (e.key === 'Escape' && !el.searchBar.hasAttribute('hidden')) {
      closeSearchBar();
    }
  });

  function toggleSearchBar() {
    if (el.searchBar.hasAttribute('hidden')) {
      el.searchBar.removeAttribute('hidden');
      el.searchInput.focus();
      el.searchInput.select();
      performSearch();
    } else {
      closeSearchBar();
    }
  }

  function closeSearchBar() {
    el.searchBar.setAttribute('hidden', '');
    el.ulnEditor.focus();
  }

  el.btnToggleReplace.addEventListener('click', () => {
    if (el.replaceRow.hasAttribute('hidden')) {
      el.replaceRow.removeAttribute('hidden');
      el.replaceInput.focus();
    } else {
      el.replaceRow.setAttribute('hidden', '');
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

  // ── Modal Handling ─────────────────────────────────────────────
  function openModal(modalEl) {
    modalEl.removeAttribute('hidden');
  }

  function closeModal(modalEl) {
    modalEl.setAttribute('hidden', '');
  }

  document.querySelectorAll('[data-close]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const targetId = btn.getAttribute('data-close');
      const targetModal = document.getElementById(targetId);
      if (targetModal) closeModal(targetModal);
    });
  });

  // Close modals on clicking backdrop
  document.querySelectorAll('.modal-backdrop').forEach((backdrop) => {
    backdrop.addEventListener('click', (e) => {
      if (e.target === backdrop) closeModal(backdrop);
    });
  });

  // ── AI Prompt Modal ────────────────────────────────────────────
  el.btnPromptModal.addEventListener('click', async () => {
    try {
      const res = await fetch('/api/prompt');
      if (res.ok) {
        const data = await res.json();
        el.promptTextarea.value = data.prompt_text || '';
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
          el.promptTextarea.value = data.default_prompt || '';
          showToast('Đã khôi phục prompt mặc định', 'info');
        }
      } catch (err) {
        console.error(err);
      }
    }
  });

  // ── DOCX Scanner & Importer Modal ──────────────────────────────
  el.btnImportDocx.addEventListener('click', () => {
    openModal(el.modalDocx);
    scanDocxDirectory();
  });

  el.btnScanFolder.addEventListener('click', scanDocxDirectory);
  el.chkOnlyUln.addEventListener('change', scanDocxDirectory);

  async function scanDocxDirectory() {
    const folder = el.docxFolderInput.value.trim();
    const onlyUln = el.chkOnlyUln.checked;
    el.docxTableBody.innerHTML = '<tr><td colspan="4" class="text-center">Đang quét thư mục...</td></tr>';
    el.btnLoadDocxToEditor.disabled = true;
    el.docxPreviewContent.textContent = 'Chọn một file để xem trước...';
    state.selectedDocx = null;

    try {
      const url = `/api/docx-list?folder=${encodeURIComponent(folder)}&only_uln=${onlyUln}`;
      const res = await fetch(url);
      if (!res.ok) throw new Error('Không thể đọc thư mục');
      const files = await res.json();

      el.docxTableBody.innerHTML = '';
      if (!files || files.length === 0) {
        el.docxTableBody.innerHTML = '<tr><td colspan="4" class="text-center">Không tìm thấy file DOCX nào.</td></tr>';
        el.docxScannerStatus.textContent = '0 file tìm thấy';
        return;
      }

      el.docxScannerStatus.textContent = `${files.length} file tìm thấy`;

      files.forEach((file) => {
        const tr = document.createElement('tr');

        const tdEmbed = document.createElement('td');
        tdEmbed.innerHTML = file.has_embedded_uln
          ? '<span class="tag-embedded">Có ULN</span>'
          : '<span class="tag-none">Không</span>';

        const tdName = document.createElement('td');
        tdName.textContent = file.filename;
        tdName.title = file.filepath;

        const tdSize = document.createElement('td');
        tdSize.textContent = `${(file.size / 1024).toFixed(0)} KB`;

        const tdTime = document.createElement('td');
        tdTime.textContent = new Date(file.modified * 1000).toLocaleString('vi-VN', {
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
        });

        tr.appendChild(tdEmbed);
        tr.appendChild(tdName);
        tr.appendChild(tdSize);
        tr.appendChild(tdTime);

        tr.addEventListener('click', () => {
          document.querySelectorAll('#docxTableBody tr').forEach((r) => r.classList.remove('selected'));
          tr.classList.add('selected');
          state.selectedDocx = file;

          if (file.has_embedded_uln && file.embedded_preview) {
            el.docxPreviewContent.textContent = file.embedded_preview;
            el.btnLoadDocxToEditor.disabled = false;
          } else {
            el.docxPreviewContent.textContent = '(File này không chứa mã ULN nhúng trong Custom XML Part)';
            el.btnLoadDocxToEditor.disabled = true;
          }
        });

        el.docxTableBody.appendChild(tr);
      });
    } catch (err) {
      el.docxTableBody.innerHTML = `<tr><td colspan="4" class="text-center" style="color:#ef4444">${err.message}</td></tr>`;
    }
  }

  el.btnLoadDocxToEditor.addEventListener('click', () => {
    if (!state.selectedDocx || !state.selectedDocx.embedded_preview) return;
    el.ulnEditor.value = state.selectedDocx.embedded_preview;
    syncLineNumbers();
    updateEditorStats();
    closeModal(el.modalDocx);
    showToast(`Đã nạp mã ULN từ: ${state.selectedDocx.filename}`, 'success');
  });

  // ── Embed Raw ULN to Existing DOCX ─────────────────────────────
  el.btnEmbedRaw.addEventListener('click', async () => {
    const text = el.ulnEditor.value.trim();
    if (!text) {
      showToast('Trình soạn thảo rỗng, không có gì để nhúng!', 'warning');
      return;
    }
    const filepath = prompt('Nhập đường dẫn đầy đủ của file DOCX cần nhúng mã ULN:');
    if (!filepath) return;

    try {
      const res = await fetch('/api/embed-uln', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filepath, uln_text: text }),
      });
      if (res.ok) {
        showToast('Đã nhúng mã ULN thành công vào tệp Word!', 'success');
      } else {
        const err = await res.json();
        showToast(`Nhúng thất bại: ${err.detail}`, 'error');
      }
    } catch (err) {
      showToast(`Lỗi: ${err.message}`, 'error');
    }
  });

  // ── Check Updates ──────────────────────────────────────────────
  el.btnCheckUpdate.addEventListener('click', async () => {
    try {
      showToast('Đang kiểm tra bản cập nhật...', 'info');
      const res = await fetch('/api/check-update');
      if (res.ok) {
        const data = await res.json();
        if (data.update_available) {
          showToast(`Có bản cập nhật mới: v${data.latest_version}`, 'info', {
            label: 'Xem trên GitHub',
            onClick: () => window.open(data.release_url, '_blank'),
          });
        } else {
          showToast(`Bạn đang sử dụng phiên bản mới nhất (v${data.current_version})`, 'success');
        }
      }
    } catch (err) {
      showToast('Không thể kiểm tra cập nhật', 'error');
    }
  });

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

    // UI state into compiling
    el.btnCompile.disabled = true;
    el.progressWrap.removeAttribute('hidden');
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

      // Start polling
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

  // ── Run App ────────────────────────────────────────────────────
  initApp();
});
