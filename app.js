/**
 * Excel Schema Merger - Application Core Logic
 * 순수 Vanilla JavaScript + SheetJS + JSZip
 */

(function () {
  'use strict';

  // --- State Management ---
  const state = {
    uploadedFiles: [], // Array of file objects with parsed data
    groups: [],        // Array of grouped schema objects
    options: {
      addSourceFile: true,
      flexibleOrder: true,
      trimEmptyRows: true
    },
    activePreviewGroup: null
  };

  // --- DOM Elements ---
  const DOM = {
    dropZone: document.getElementById('dropZone'),
    fileInput: document.getElementById('fileInput'),
    btnSampleDemo: document.getElementById('btnSampleDemo'),
    btnResetAll: document.getElementById('btnResetAll'),
    optAddSourceFile: document.getElementById('optAddSourceFile'),
    optFlexibleOrder: document.getElementById('optFlexibleOrder'),
    optTrimEmptyRows: document.getElementById('optTrimEmptyRows'),
    statsSection: document.getElementById('statsSection'),
    statTotalFiles: document.getElementById('statTotalFiles'),
    statTotalGroups: document.getElementById('statTotalGroups'),
    statTotalRows: document.getElementById('statTotalRows'),
    statSingleGroups: document.getElementById('statSingleGroups'),
    btnDownloadAllZip: document.getElementById('btnDownloadAllZip'),
    emptyState: document.getElementById('emptyState'),
    groupCardsList: document.getElementById('groupCardsList'),
    previewModal: document.getElementById('previewModal'),
    modalGroupName: document.getElementById('modalGroupName'),
    modalIncludedFilesCount: document.getElementById('modalIncludedFilesCount'),
    modalTotalRowsCount: document.getElementById('modalTotalRowsCount'),
    previewThead: document.getElementById('previewThead'),
    previewTbody: document.getElementById('previewTbody'),
    btnModalClose: document.getElementById('btnModalClose'),
    btnModalDismiss: document.getElementById('btnModalDismiss'),
    btnModalDownload: document.getElementById('btnModalDownload'),
    toastContainer: document.getElementById('toastContainer')
  };

  // --- Initialize App ---
  function init() {
    setupEventListeners();
    refreshLucideIcons();
  }

  // --- Event Listeners ---
  function setupEventListeners() {
    // Drop zone & File Input
    DOM.dropZone.addEventListener('click', () => DOM.fileInput.click());
    DOM.fileInput.addEventListener('change', handleFileInput);

    DOM.dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      DOM.dropZone.classList.add('dragover');
    });
    DOM.dropZone.addEventListener('dragleave', (e) => {
      e.preventDefault();
      DOM.dropZone.classList.remove('dragover');
    });
    DOM.dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      DOM.dropZone.classList.remove('dragover');
      if (e.dataTransfer && e.dataTransfer.files.length > 0) {
        handleFiles(e.dataTransfer.files);
      }
    });

    // Options toggles
    DOM.optAddSourceFile.addEventListener('change', (e) => {
      state.options.addSourceFile = e.target.checked;
      renderGroups();
    });
    DOM.optFlexibleOrder.addEventListener('change', (e) => {
      state.options.flexibleOrder = e.target.checked;
      rebuildGroups();
    });
    DOM.optTrimEmptyRows.addEventListener('change', (e) => {
      state.options.trimEmptyRows = e.target.checked;
      rebuildGroups();
    });

    // Actions
    DOM.btnResetAll.addEventListener('click', resetAll);
    DOM.btnSampleDemo.addEventListener('click', loadDemoData);
    DOM.btnDownloadAllZip.addEventListener('click', downloadAllGroupsZip);

    // Modal controls
    DOM.btnModalClose.addEventListener('click', closeModal);
    DOM.btnModalDismiss.addEventListener('click', closeModal);
    DOM.previewModal.addEventListener('click', (e) => {
      if (e.target === DOM.previewModal) closeModal();
    });
    DOM.btnModalDownload.addEventListener('click', () => {
      if (state.activePreviewGroup) {
        downloadGroupExcel(state.activePreviewGroup.id);
      }
    });
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && DOM.previewModal.style.display !== 'none') {
        closeModal();
      }
    });
  }

  function refreshLucideIcons() {
    if (window.lucide) {
      window.lucide.createIcons();
    }
  }

  // --- Toast Notification ---
  function showToast(message, type = 'info', duration = 3200) {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    let iconName = 'info';
    if (type === 'success') iconName = 'check-circle';
    if (type === 'error') iconName = 'alert-triangle';

    toast.innerHTML = `
      <i data-lucide="${iconName}" class="mini-icon"></i>
      <span>${escapeHtml(message)}</span>
    `;

    DOM.toastContainer.appendChild(toast);
    refreshLucideIcons();

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateX(30px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, duration);
  }

  // --- File Handling & Parsing ---
  function handleFileInput(e) {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(e.target.files);
    }
    DOM.fileInput.value = ''; // Reset input so same files can be re-added
  }

  async function handleFiles(fileList) {
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const filesToProcess = [];

    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      const ext = '.' + file.name.split('.').pop().toLowerCase();
      if (validExtensions.includes(ext)) {
        filesToProcess.push(file);
      } else {
        showToast(`'${file.name}' 은(는) 지원되지 않는 파일 형식입니다.`, 'error');
      }
    }

    if (filesToProcess.length === 0) return;

    showToast(`${filesToProcess.length}개 엑셀 파일 분석을 시작합니다...`, 'info', 2000);

    let parsedCount = 0;
    for (const file of filesToProcess) {
      try {
        const parsed = await parseExcelFile(file);
        if (parsed) {
          // Check if file with same name already exists; if so, rename or replace
          const existingIndex = state.uploadedFiles.findIndex(f => f.name === parsed.name);
          if (existingIndex >= 0) {
            state.uploadedFiles[existingIndex] = parsed;
          } else {
            state.uploadedFiles.push(parsed);
          }
          parsedCount++;
        }
      } catch (err) {
        console.error('File parse error:', err);
        showToast(`'${file.name}' 파싱 실패: ${err.message}`, 'error');
      }
    }

    if (parsedCount > 0) {
      showToast(`${parsedCount}개 파일 분석 및 규격 분류 완료!`, 'success');
      rebuildGroups();
    }
  }

  // SheetJS parsing helper
  function parseExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });

          if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            throw new Error('시트가 비어있는 엑셀 파일입니다.');
          }

          // Read the first sheet
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];

          // Convert sheet to 2D array
          const rawAoa = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: '',
            blankrows: false
          });

          if (!rawAoa || rawAoa.length === 0) {
            throw new Error('데이터가 없는 빈 시트입니다.');
          }

          // Header is row 0
          const rawHeaders = rawAoa[0] || [];
          const cleanHeaders = rawHeaders.map(h => String(h || '').trim()).filter(h => h.length > 0);

          if (cleanHeaders.length === 0) {
            throw new Error('유효한 헤더(컬럼명)를 찾을 수 없습니다.');
          }

          // Data rows (from row 1 onwards)
          let dataRows = rawAoa.slice(1);

          if (state.options.trimEmptyRows) {
            dataRows = dataRows.filter(row => {
              if (!row || !Array.isArray(row)) return false;
              return row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== '');
            });
          }

          resolve({
            id: 'file_' + Math.random().toString(36).substr(2, 9),
            name: file.name,
            size: file.size,
            sheetName: sheetName,
            headers: cleanHeaders,
            rawHeaders: rawHeaders.map(h => String(h || '').trim()),
            dataRows: dataRows,
            rowCount: dataRows.length
          });
        } catch (err) {
          reject(err);
        }
      };

      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }

  // --- Grouping Engine ---
  function computeSchemaSignature(headers, flexibleOrder) {
    const normalized = headers.map(h => h.toLowerCase());
    if (flexibleOrder) {
      return [...normalized].sort().join('###');
    }
    return normalized.join('###');
  }

  function rebuildGroups() {
    if (state.uploadedFiles.length === 0) {
      state.groups = [];
      renderUI();
      return;
    }

    const groupsMap = new Map();

    state.uploadedFiles.forEach(file => {
      const signature = computeSchemaSignature(file.headers, state.options.flexibleOrder);

      if (!groupsMap.has(signature)) {
        groupsMap.set(signature, {
          id: 'grp_' + Math.random().toString(36).substr(2, 9),
          signature: signature,
          canonicalHeaders: [...file.headers], // Canonical order is set by first file in group
          files: [],
          customName: ''
        });
      }

      groupsMap.get(signature).files.push(file);
    });

    let groupIndex = 1;
    state.groups = Array.from(groupsMap.values()).map(grp => {
      // Create auto-generated default filename
      const summaryHeader = grp.canonicalHeaders.slice(0, 3).join('_');
      const countStr = grp.files.length > 1 ? `${grp.files.length}개파일` : '단독파일';
      const defaultName = `병합_규격${groupIndex}_[${summaryHeader}]_${countStr}`;
      groupIndex++;

      return {
        ...grp,
        name: grp.customName || defaultName,
        totalRows: grp.files.reduce((sum, f) => sum + f.rowCount, 0)
      };
    });

    renderUI();
  }

  // --- UI Rendering ---
  function renderUI() {
    const hasFiles = state.uploadedFiles.length > 0;

    DOM.statsSection.style.display = hasFiles ? 'flex' : 'none';
    DOM.btnResetAll.style.display = hasFiles ? 'inline-flex' : 'none';
    DOM.emptyState.style.display = hasFiles ? 'none' : 'block';

    if (hasFiles) {
      // Calculate stats
      const totalFiles = state.uploadedFiles.length;
      const totalGroups = state.groups.length;
      const totalRows = state.groups.reduce((sum, g) => sum + g.totalRows, 0);
      const singleGroups = state.groups.filter(g => g.files.length === 1).length;

      DOM.statTotalFiles.textContent = totalFiles.toLocaleString();
      DOM.statTotalGroups.textContent = totalGroups.toLocaleString();
      DOM.statTotalRows.textContent = totalRows.toLocaleString();
      DOM.statSingleGroups.textContent = singleGroups.toLocaleString();

      renderGroups();
    } else {
      DOM.groupCardsList.innerHTML = '';
    }

    refreshLucideIcons();
  }

  function renderGroups() {
    DOM.groupCardsList.innerHTML = '';

    state.groups.forEach((group, idx) => {
      const isMulti = group.files.length > 1;
      const card = document.createElement('div');
      card.className = `group-card ${isMulti ? 'multi-file' : 'single-file'}`;
      card.dataset.groupId = group.id;

      // Group Header
      const headerEl = document.createElement('div');
      headerEl.className = 'group-header';

      const metaEl = document.createElement('div');
      metaEl.className = 'group-meta';

      const badge = document.createElement('span');
      badge.className = `group-badge ${isMulti ? 'badge-multi' : 'badge-single'}`;
      badge.textContent = isMulti ? `규격 그룹 #${idx + 1} (${group.files.length}개 파일 병합)` : `단독 규격 #${idx + 1}`;

      const nameInput = document.createElement('input');
      nameInput.type = 'text';
      nameInput.className = 'group-filename-input';
      nameInput.value = group.name;
      nameInput.title = '병합 다운로드 시 생성될 파일명 (클릭하여 수정 가능)';
      nameInput.addEventListener('change', (e) => {
        group.name = e.target.value.trim() || `규격그룹_${idx + 1}`;
      });

      const countsEl = document.createElement('div');
      countsEl.className = 'group-counts';
      countsEl.innerHTML = `
        <span><i data-lucide="columns" class="mini-icon"></i> ${group.canonicalHeaders.length}개 열</span>
        <span>&bull;</span>
        <span><i data-lucide="table" class="mini-icon"></i> 총 ${group.totalRows.toLocaleString()}개 데이터 행</span>
      `;

      metaEl.appendChild(badge);
      metaEl.appendChild(nameInput);
      metaEl.appendChild(countsEl);

      // Actions
      const actionsEl = document.createElement('div');
      actionsEl.className = 'group-actions';

      const btnPreview = document.createElement('button');
      btnPreview.className = 'btn btn-secondary btn-sm';
      btnPreview.innerHTML = `<i data-lucide="eye" class="mini-icon"></i> <span>미리보기</span>`;
      btnPreview.addEventListener('click', () => openPreviewModal(group));

      const btnDownload = document.createElement('button');
      btnDownload.className = 'btn btn-primary btn-sm';
      btnDownload.innerHTML = `<i data-lucide="download" class="mini-icon"></i> <span>병합 엑셀 다운로드</span>`;
      btnDownload.addEventListener('click', () => downloadGroupExcel(group.id));

      actionsEl.appendChild(btnPreview);
      actionsEl.appendChild(btnDownload);

      headerEl.appendChild(metaEl);
      headerEl.appendChild(actionsEl);

      // Group Body
      const bodyEl = document.createElement('div');
      bodyEl.className = 'group-body';

      // Columns Tag List
      const colSection = document.createElement('div');
      colSection.className = 'schema-cols-section';
      colSection.innerHTML = `
        <div class="schema-section-title">
          <i data-lucide="layout-grid" class="mini-icon"></i> 일치 규격 헤더 목록 (${group.canonicalHeaders.length}개)
        </div>
      `;
      const colTagsWrap = document.createElement('div');
      colTagsWrap.className = 'schema-columns-wrap';

      if (state.options.addSourceFile) {
        const srcTag = document.createElement('span');
        srcTag.className = 'col-tag';
        srcTag.style.borderColor = 'rgba(56, 189, 248, 0.4)';
        srcTag.style.color = '#38bdf8';
        srcTag.innerHTML = `<span class="col-index">[추가열]</span> _원본파일`;
        colTagsWrap.appendChild(srcTag);
      }

      group.canonicalHeaders.forEach((col, cIdx) => {
        const tag = document.createElement('span');
        tag.className = 'col-tag';
        tag.innerHTML = `<span class="col-index">${cIdx + 1}</span> ${escapeHtml(col)}`;
        colTagsWrap.appendChild(tag);
      });
      colSection.appendChild(colTagsWrap);

      // Files list in this group
      const filesSection = document.createElement('div');
      filesSection.className = 'schema-files-section';
      filesSection.innerHTML = `
        <div class="schema-section-title">
          <i data-lucide="files" class="mini-icon"></i> 포함된 엑셀 파일 목록 (${group.files.length}개)
        </div>
      `;
      const filesGrid = document.createElement('div');
      filesGrid.className = 'group-files-list';

      group.files.forEach(file => {
        const filePill = document.createElement('div');
        filePill.className = 'file-item-pill';
        filePill.innerHTML = `
          <div class="file-name-info" title="${escapeHtml(file.name)}">
            <i data-lucide="file-spreadsheet" class="mini-icon" style="color:#10b981; flex-shrink:0;"></i>
            <span class="file-name-text">${escapeHtml(file.name)}</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.4rem;">
            <span class="file-row-count">${file.rowCount.toLocaleString()}행</span>
            <button class="btn-file-remove" title="이 파일 제외하기" data-file-id="${file.id}">
              <i data-lucide="trash-2" class="mini-icon"></i>
            </button>
          </div>
        `;

        // Remove file button
        filePill.querySelector('.btn-file-remove').addEventListener('click', (e) => {
          e.stopPropagation();
          removeFile(file.id);
        });

        filesGrid.appendChild(filePill);
      });
      filesSection.appendChild(filesGrid);

      bodyEl.appendChild(colSection);
      bodyEl.appendChild(filesSection);

      card.appendChild(headerEl);
      card.appendChild(bodyEl);
      DOM.groupCardsList.appendChild(card);
    });

    refreshLucideIcons();
  }

  function removeFile(fileId) {
    const file = state.uploadedFiles.find(f => f.id === fileId);
    state.uploadedFiles = state.uploadedFiles.filter(f => f.id !== fileId);
    showToast(`'${file ? file.name : '파일'}'을(를) 목록에서 제외했습니다.`, 'info');
    rebuildGroups();
  }

  function resetAll() {
    if (state.uploadedFiles.length === 0) return;
    if (confirm('업로드된 모든 파일과 분석 결과를 초기화하시겠습니까?')) {
      state.uploadedFiles = [];
      state.groups = [];
      renderUI();
      showToast('초기화되었습니다.', 'info');
    }
  }

  // --- Merge Data Generator ---
  function buildMergedAoaForGroup(group) {
    const headers = [...group.canonicalHeaders];
    const finalHeaders = state.options.addSourceFile ? ['_원본파일', ...headers] : [...headers];
    const aoa = [finalHeaders];

    // For each file, map its data rows into canonical header columns
    group.files.forEach(file => {
      // Build a column map from file's header position to canonical header position
      const colMap = [];
      group.canonicalHeaders.forEach(canonicalCol => {
        const fileColIdx = file.headers.findIndex(
          h => h.trim().toLowerCase() === canonicalCol.trim().toLowerCase()
        );
        colMap.push(fileColIdx);
      });

      file.dataRows.forEach(row => {
        const mappedRow = colMap.map(srcIdx => {
          if (srcIdx === -1 || srcIdx >= row.length || row[srcIdx] === undefined) {
            return '';
          }
          return row[srcIdx];
        });

        if (state.options.addSourceFile) {
          aoa.push([file.name, ...mappedRow]);
        } else {
          aoa.push(mappedRow);
        }
      });
    });

    return aoa;
  }

  // --- Preview Modal ---
  function openPreviewModal(group) {
    state.activePreviewGroup = group;
    DOM.modalGroupName.textContent = group.name;
    DOM.modalIncludedFilesCount.textContent = `${group.files.length}개 파일 포함 (${group.files.map(f => f.name).join(', ')})`;
    DOM.modalTotalRowsCount.textContent = `총 ${group.totalRows.toLocaleString()}개 행`;

    const mergedAoa = buildMergedAoaForGroup(group);
    const headers = mergedAoa[0] || [];
    const sampleRows = mergedAoa.slice(1, 51); // Top 50 rows for preview

    // Render Table Header
    DOM.previewThead.innerHTML = `
      <tr>
        <th style="width: 50px;">#</th>
        ${headers.map(h => `<th class="${h === '_원본파일' ? 'source-cell' : ''}">${escapeHtml(String(h))}</th>`).join('')}
      </tr>
    `;

    // Render Table Body
    if (sampleRows.length === 0) {
      DOM.previewTbody.innerHTML = `
        <tr>
          <td colspan="${headers.length + 1}" style="text-align: center; color: var(--text-dim); padding: 2rem;">
            데이터 행이 없습니다.
          </td>
        </tr>
      `;
    } else {
      DOM.previewTbody.innerHTML = sampleRows.map((row, rIdx) => `
        <tr>
          <td style="color: var(--text-dim); font-family: monospace;">${rIdx + 1}</td>
          ${row.map((cell, cIdx) => {
            const isSource = state.options.addSourceFile && cIdx === 0;
            return `<td class="${isSource ? 'source-cell' : ''}">${escapeHtml(formatCellValue(cell))}</td>`;
          }).join('')}
        </tr>
      `).join('');
    }

    DOM.previewModal.style.display = 'flex';
    refreshLucideIcons();
  }

  function closeModal() {
    DOM.previewModal.style.display = 'none';
    state.activePreviewGroup = null;
  }

  function formatCellValue(val) {
    if (val === null || val === undefined) return '';
    if (val instanceof Date) {
      return val.toISOString().split('T')[0];
    }
    return String(val);
  }

  // --- Export Functions ---
  function downloadGroupExcel(groupId) {
    const group = state.groups.find(g => g.id === groupId);
    if (!group) return;

    try {
      showToast(`'${group.name}' 병합 파일 생성 중...`, 'info', 1500);

      const mergedAoa = buildMergedAoaForGroup(group);
      const ws = XLSX.utils.aoa_to_sheet(mergedAoa);
      const wb = XLSX.utils.book_new();

      // Clean sheet name
      const safeSheetName = '병합데이터';
      XLSX.utils.book_append_sheet(wb, ws, safeSheetName);

      // Sanitize output filename
      let fileName = (group.name || '병합결과').replace(/[/\\?%*:|"<>]/g, '_');
      if (!fileName.toLowerCase().endsWith('.xlsx')) {
        fileName += '.xlsx';
      }

      XLSX.writeFile(wb, fileName);
      showToast(`'${fileName}' 다운로드가 시작되었습니다!`, 'success');
    } catch (err) {
      console.error('Download error:', err);
      showToast(`다운로드 실패: ${err.message}`, 'error');
    }
  }

  async function downloadAllGroupsZip() {
    if (state.groups.length === 0) return;

    if (!window.JSZip) {
      showToast('압축 라이브러리(JSZip)를 불러오는 중입니다. 잠시 후 다시 시도해주세요.', 'error');
      return;
    }

    try {
      showToast('전체 그룹 압축 ZIP 생성 중...', 'info', 2500);
      const zip = new JSZip();

      state.groups.forEach((group, idx) => {
        const mergedAoa = buildMergedAoaForGroup(group);
        const ws = XLSX.utils.aoa_to_sheet(mergedAoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '병합데이터');

        const xlsxArrayBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

        let fileName = (group.name || `규격그룹_${idx + 1}`).replace(/[/\\?%*:|"<>]/g, '_');
        if (!fileName.toLowerCase().endsWith('.xlsx')) {
          fileName += '.xlsx';
        }

        zip.file(fileName, xlsxArrayBuffer);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const now = new Date();
      const dateStr = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0');
      const zipFilename = `Excel_Merged_Groups_${dateStr}.zip`;

      // Trigger download
      const link = document.createElement('a');
      link.href = URL.createObjectURL(zipBlob);
      link.download = zipFilename;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(link.href);

      showToast(`총 ${state.groups.length}개 그룹이 포함된 '${zipFilename}' 다운로드 완료!`, 'success');
    } catch (err) {
      console.error('ZIP generation error:', err);
      showToast(`ZIP 압축 실패: ${err.message}`, 'error');
    }
  }

  // --- Demo Sample Generator ---
  function loadDemoData() {
    showToast('체험용 샘플 엑셀 데이터 5개를 생성하여 불러옵니다...', 'info', 2000);

    // Schema A: Sales order schema (3 files)
    const salesHeader = ['주문번호', '주문일시', '고객명', '상품명', '수량', '결제금액', '주문상태'];
    const demoSales1 = [
      salesHeader,
      ['ORD-2024-001', '2024-01-05', '김철수', '무선 블루투스 이어폰', 1, 49000, '배송완료'],
      ['ORD-2024-002', '2024-01-09', '이영희', '기계식 게이밍 키보드', 2, 178000, '배송완료'],
      ['ORD-2024-003', '2024-01-14', '박민준', 'C타입 고속충전 케이블', 3, 27000, '배송완료']
    ];
    const demoSales2 = [
      salesHeader,
      ['ORD-2024-021', '2024-02-02', '최지우', '27인치 4K 모니터', 1, 389000, '배송완료'],
      ['ORD-2024-022', '2024-02-18', '정수빈', '인체공학 버티컬 마우스', 1, 55000, '결제완료']
    ];
    const demoSales3 = [
      salesHeader,
      ['ORD-2024-035', '2024-03-01', '강동현', '노트북 거치대', 2, 64000, '배송준비중'],
      ['ORD-2024-036', '2024-03-12', '윤아라', '고속 무선충전 패드', 1, 32000, '배송완료'],
      ['ORD-2024-037', '2024-03-20', '한승우', '노이즈캔슬링 헤드폰', 1, 289000, '배송완료']
    ];

    // Schema B: Personnel schema (2 files)
    const hrHeader = ['사번', '성명', '소속부서', '직급', '입사일자', '사내이메일'];
    const demoHr1 = [
      hrHeader,
      ['DEV-101', '김개발', '플랫폼개발팀', '수석엔지니어', '2021-03-02', 'kim.dev@company.com'],
      ['DEV-102', '이프론트', '웹프론트팀', '선임연구원', '2022-07-15', 'lee.fe@company.com'],
      ['DEV-103', '박백엔드', '서버코어팀', '연구원', '2023-11-01', 'park.be@company.com']
    ];
    const demoHr2 = [
      hrHeader,
      ['DES-201', '최디자인', 'UX/UI디자인실', '팀장', '2020-05-18', 'choi.ux@company.com'],
      ['DES-202', '정그래픽', '브랜드디자인팀', '선임디자이너', '2022-09-01', 'jung.brand@company.com']
    ];

    const sampleFiles = [
      { name: '2024년_01월_주문매출내역.xlsx', aoa: demoSales1 },
      { name: '2024년_02월_주문매출내역.xlsx', aoa: demoSales2 },
      { name: '2024년_03월_주문매출내역.xlsx', aoa: demoSales3 },
      { name: '2024년_임직원명부_개발본부.xlsx', aoa: demoHr1 },
      { name: '2024년_임직원명부_디자인실.xlsx', aoa: demoHr2 }
    ];

    state.uploadedFiles = sampleFiles.map(s => {
      const cleanHeaders = s.aoa[0].map(h => String(h).trim());
      const dataRows = s.aoa.slice(1);
      return {
        id: 'file_' + Math.random().toString(36).substr(2, 9),
        name: s.name,
        size: 15420,
        sheetName: 'Sheet1',
        headers: cleanHeaders,
        rawHeaders: cleanHeaders,
        dataRows: dataRows,
        rowCount: dataRows.length
      };
    });

    rebuildGroups();
    showToast('샘플 엑셀 파일 5개가 2개의 규격 그룹으로 자동 분류되었습니다!', 'success', 3500);
  }

  // --- Utility Functions ---
  function escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // Initialize on DOM load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
