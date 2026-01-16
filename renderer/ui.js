// Toast notification system
function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  document.body.appendChild(toast);

  setTimeout(() => toast.classList.add('show'), 10);

  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

// Tool navigation
const toolButtons = document.querySelectorAll('.tool-btn');
const toolPanels = document.querySelectorAll('.tool-panel');

toolButtons.forEach((btn) => {
  btn.addEventListener('click', () => {
    const toolName = btn.dataset.tool;

    // Update active states
    toolButtons.forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');

    toolPanels.forEach((panel) => {
      panel.classList.remove('active');
      if (panel.id === `${toolName}-tool`) {
        panel.classList.add('active');
      }
    });
  });
});

// Merge Tool Logic
const mergeDropzone = document.getElementById('mergeDropzone');
const selectMergeFilesBtn = document.getElementById('selectMergeFilesBtn');
const mergeBtn = document.getElementById('mergeBtn');
const mergeActionBar = document.getElementById('mergeActionBar');
const fileInput = document.getElementById('files');
const fileList = document.getElementById('fileList');

let files = [];
let filePreviewStates = {}; // Track preview state for each file

selectMergeFilesBtn.onclick = () => {
  fileInput.click();
};

// Drag and drop for merge
mergeDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  mergeDropzone.classList.add('drag-over');
});

mergeDropzone.addEventListener('dragleave', () => {
  mergeDropzone.classList.remove('drag-over');
});

mergeDropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  mergeDropzone.classList.remove('drag-over');

  const droppedFiles = [...e.dataTransfer.files]
    .filter(f => f.type === 'application/pdf')
    .map(f => ({ path: f.path, name: f.name }));

  if (droppedFiles.length > 0) {
    files = [...files, ...droppedFiles];
    renderFileList();
  } else {
    showToast('Please drop PDF files only', 'error');
  }
});

fileInput.onchange = () => {
  const newFiles = [...fileInput.files].map((f) => ({ path: f.path, name: f.name }));
  files = [...files, ...newFiles];
  renderFileList();
  fileInput.value = ''; // Reset input to allow re-adding same file
};

function togglePreview(index) {
  filePreviewStates[index] = !filePreviewStates[index];
  renderFileList();
}

async function renderFileList() {
  fileList.innerHTML = '';

  if (files.length === 0) {
    mergeActionBar.style.display = 'none';
    filePreviewStates = {};
    return;
  }

  mergeActionBar.style.display = 'block';
  fileList.className = 'file-list';
  
  for (let index = 0; index < files.length; index++) {
    const file = files[index];
    const showPreview = filePreviewStates[index] || false;
    
    const pdfContainer = document.createElement('div');
    pdfContainer.className = 'pdf-file-container';
    pdfContainer.dataset.fileIndex = index;
    
    // File header with controls
    const fileHeader = document.createElement('div');
    fileHeader.className = 'file-item';
    fileHeader.draggable = true;
    fileHeader.dataset.index = index;
    
    fileHeader.innerHTML = `
      <div class="file-info">
        <div class="file-number">${index + 1}</div>
        <div class="file-name" title="${file.name}">${file.name}</div>
      </div>
      <div class="file-actions">
        <button class="btn-preview" data-index="${index}">
          ${showPreview ? '📝 Hide' : '👁️ Preview'}
        </button>
        <button class="file-remove" data-index="${index}">✕ Remove</button>
      </div>
    `;
    
    fileHeader.addEventListener('dragstart', handleDragStart);
    fileHeader.addEventListener('dragover', handleDragOver);
    fileHeader.addEventListener('drop', handleDrop);
    fileHeader.addEventListener('dragend', handleDragEnd);
    
    pdfContainer.appendChild(fileHeader);
    
    // Thumbnail container (shown/hidden based on state)
    if (showPreview) {
      const thumbnailContainer = document.createElement('div');
      thumbnailContainer.className = 'thumbnail-container';
      thumbnailContainer.innerHTML = '<div class="loading-thumbnails">⏳ Loading preview...</div>';
      pdfContainer.appendChild(thumbnailContainer);
      
      // Load thumbnails asynchronously
      (async () => {
        try {
          const thumbnails = await window.pdfAPI.getPDFThumbnails(file.path);
          thumbnailContainer.innerHTML = '';
          
          thumbnails.forEach((thumb, pageIndex) => {
            const thumbItem = document.createElement('div');
            thumbItem.className = 'thumbnail-item';
            
            thumbItem.innerHTML = `
              <img src="${thumb}" alt="Page ${pageIndex + 1}" />
              <div class="thumb-label">Page ${pageIndex + 1}</div>
            `;
            
            thumbnailContainer.appendChild(thumbItem);
          });
        } catch (err) {
          thumbnailContainer.innerHTML = '<div class="error-thumbnails">❌ Failed to load preview</div>';
          console.error('Thumbnail error:', err);
        }
      })();
    }
    
    fileList.appendChild(pdfContainer);
  }

  // Attach event handlers
  document.querySelectorAll('.btn-preview').forEach((btn) => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const idx = parseInt(e.target.dataset.index);
      togglePreview(idx);
    };
  });

  // Attach remove handlers
  document.querySelectorAll('.file-remove').forEach((btn) => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const idx = parseInt(e.target.dataset.index);
      files.splice(idx, 1);
      delete filePreviewStates[idx];
      // Reindex preview states
      const newStates = {};
      Object.keys(filePreviewStates).forEach(key => {
        const oldIdx = parseInt(key);
        if (oldIdx > idx) {
          newStates[oldIdx - 1] = filePreviewStates[oldIdx];
        } else if (oldIdx < idx) {
          newStates[oldIdx] = filePreviewStates[oldIdx];
        }
      });
      filePreviewStates = newStates;
      renderFileList();
    };
  });
}

let draggedIndex = null;

function handleDragStart(e) {
  draggedIndex = parseInt(e.target.dataset.index);
  e.target.classList.add('dragging');
}

function handleDragOver(e) {
  e.preventDefault();
  const target = e.target.closest('.file-item');
  if (target && !target.classList.contains('dragging')) {
    document.querySelectorAll('.file-item').forEach((item) => {
      item.classList.remove('drag-over');
    });
    target.classList.add('drag-over');
  }
}

function handleDrop(e) {
  e.preventDefault();
  const target = e.target.closest('.file-item');
  if (target && draggedIndex !== null) {
    const dropIndex = parseInt(target.dataset.index);

    if (draggedIndex !== dropIndex) {
      const [removed] = files.splice(draggedIndex, 1);
      files.splice(dropIndex, 0, removed);
      renderFileList();
    }
  }
}

function handleDragEnd(e) {
  e.target.classList.remove('dragging');
  document.querySelectorAll('.file-item').forEach((item) => {
    item.classList.remove('drag-over');
  });
  draggedIndex = null;
}

mergeBtn.onclick = async () => {
  if (files.length === 0) {
    alert('Select PDFs to merge');
    return;
  }

  try {
    mergeBtn.disabled = true;
    mergeBtn.textContent = '⏳ Merging...';

    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      mergeBtn.disabled = false;
      mergeBtn.innerHTML = '<span>📥</span> Merge PDFs';
      return;
    }

    const filePaths = files.map((f) => f.path);
    await window.pdfAPI.merge(filePaths, savePath);

    alert('✅ PDFs merged successfully!');

    // Reset
    files = [];
    fileInput.value = '';
    renderFileList();
    mergeBtn.disabled = false;
    mergeBtn.innerHTML = '<span>📥</span> Merge PDFs';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to merge PDFs: ' + err.message);
    mergeBtn.disabled = false;
    mergeBtn.innerHTML = '<span>📥</span> Merge PDFs';
  }
};

// Split Tool Logic
const selectPdfBtn = document.getElementById('selectPdfBtn');
const splitFileInput = document.getElementById('splitFileInput');
const splitDropzone = document.getElementById('splitDropzone');
const pdfInfo = document.getElementById('pdfInfo');
const splitOptions = document.getElementById('splitOptions');
const addRangeBtn = document.getElementById('addRangeBtn');
const rangeList = document.getElementById('rangeList');
const splitBtn = document.getElementById('splitBtn');
const pagesPerFileInput = document.getElementById('pagesPerFile');
const pagesDisplay = document.getElementById('pagesDisplay');
const eachPageInfo = document.getElementById('eachPageInfo');
const fixedPrefixInput = document.getElementById('fixedPrefix');
const fixedSuffixInput = document.getElementById('fixedSuffix');
const eachPrefixInput = document.getElementById('eachPrefix');
const eachSuffixInput = document.getElementById('eachSuffix');

let selectedPdf = null;
let totalPdfPages = 0;
let ranges = [];
let currentSplitMode = 'range';
let splitThumbnails = [];
let selectedSplitPages = new Set();
let splitPreviewVisible = false;

// Drag and drop handlers
splitDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  splitDropzone.classList.add('drag-over');
});

splitDropzone.addEventListener('dragleave', () => {
  splitDropzone.classList.remove('drag-over');
});

splitDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  splitDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadPDF(file.path);
  } else {
    alert('❌ Please drop a PDF file');
  }
});

selectPdfBtn.onclick = () => {
  splitFileInput.click();
};

splitFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadPDF(file.path);
  }
};

async function loadPDF(filePath) {
  selectedPdf = filePath;
  selectedSplitPages.clear();

  try {
    totalPdfPages = await window.pdfAPI.getPDFPageCount(filePath);

    pdfInfo.innerHTML = `
      <div class="info-icon">📄</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${totalPdfPages} pages total</p>
        <button id="toggleSplitPreviewBtn" class="preview-btn">📄 Show Preview</button>
      </div>
      <button class="change-btn" onclick="document.getElementById('splitFileInput').click()">Change</button>
    `;
    pdfInfo.classList.add('active');
    splitOptions.style.display = 'block';
    splitPreviewVisible = false;
    splitThumbnails = []; // Reset thumbnails
    
    // Reattach event listener to the new button
    const previewBtn = document.getElementById('toggleSplitPreviewBtn');
    if (previewBtn) {
      previewBtn.onclick = toggleSplitPreview;
    }

    // Initialize based on current mode
    if (currentSplitMode === 'range') {
      ranges = [{ start: 1, end: totalPdfPages, name: 'part1' }];
      renderRanges();
    } else if (currentSplitMode === 'each') {
      updateEachPageInfo();
    }

    splitBtn.disabled = false;
  } catch (err) {
    console.error(err);
    alert('❌ Failed to read PDF: ' + err.message);
  }
}

async function toggleSplitPreview() {
  const previewContainer = document.getElementById('splitPreviewContainer');
  const previewBtn = document.getElementById('toggleSplitPreviewBtn');
  
  if (!splitPreviewVisible) {
    // Load thumbnails if not already loaded
    if (splitThumbnails.length === 0 && selectedPdf) {
      if (previewBtn) {
        previewBtn.disabled = true;
        previewBtn.textContent = '⏳ Loading...';
      }
      try {
        splitThumbnails = await window.pdfAPI.getPDFThumbnails(selectedPdf);
      } catch (err) {
        alert('❌ Failed to load preview: ' + err.message);
        if (previewBtn) {
          previewBtn.disabled = false;
          previewBtn.textContent = '📄 Show Preview';
        }
        return;
      }
      if (previewBtn) {
        previewBtn.disabled = false;
      }
    }
    // Show preview
    renderSplitPreview();
    splitPreviewVisible = true;
    if (previewBtn) previewBtn.textContent = '📄 Hide Preview';
  } else {
    // Hide preview
    if (previewContainer) {
      previewContainer.style.display = 'none';
    }
    splitPreviewVisible = false;
    if (previewBtn) previewBtn.textContent = '📄 Show Preview';
  }
}

function renderSplitPreview() {
  let previewContainer = document.getElementById('splitPreviewContainer');
  if (!previewContainer) {
    previewContainer = document.createElement('div');
    previewContainer.id = 'splitPreviewContainer';
    previewContainer.className = 'split-preview-container';
    const splitOptions = document.getElementById('splitOptions');
    if (splitOptions && splitOptions.parentElement) {
      // Insert before splitOptions within the split-container
      splitOptions.parentElement.insertBefore(previewContainer, splitOptions);
    }
  }
  
  previewContainer.style.display = 'block';
  previewContainer.innerHTML = '<h3 style="margin: 20px 0 10px; font-size: 16px;">📄 PDF Pages Preview</h3>';
  
  const thumbnailGrid = document.createElement('div');
  thumbnailGrid.className = 'split-thumbnail-grid';
  
  splitThumbnails.forEach((thumb, index) => {
    const thumbItem = document.createElement('div');
    thumbItem.className = 'split-thumbnail-item';
    thumbItem.dataset.pageIndex = index;
    
    thumbItem.innerHTML = `
      <img src="${thumb}" alt="Page ${index + 1}" />
      <div class="split-thumb-label">Page ${index + 1}</div>
      <div class="split-thumb-check">✓</div>
    `;
    
    thumbItem.onclick = () => {
      if (currentSplitMode === 'visual') {
        thumbItem.classList.toggle('selected');
        if (thumbItem.classList.contains('selected')) {
          selectedSplitPages.add(index + 1);
        } else {
          selectedSplitPages.delete(index + 1);
        }
      }
    };
    
    thumbnailGrid.appendChild(thumbItem);
  });
  
  previewContainer.appendChild(thumbnailGrid);
}

// Mode switching
const modeButtons = document.querySelectorAll('.mode-btn');
const modeContents = document.querySelectorAll('.mode-content');

modeButtons.forEach((btn) => {
  btn.addEventListener('click', () => {
    const mode = btn.dataset.mode;
    currentSplitMode = mode;

    modeButtons.forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');

    modeContents.forEach((content) => {
      content.classList.remove('active');
      if (content.id === `${mode}Mode`) {
        content.classList.add('active');
      }
    });

    if (mode === 'each' && totalPdfPages > 0) {
      updateEachPageInfo();
    }
  });
});

function renderRanges() {
  rangeList.innerHTML = '';

  ranges.forEach((range, index) => {
    const item = document.createElement('div');
    item.className = 'range-item';

    item.innerHTML = `
      <label>From:</label>
      <input type="number" class="start-input" value="${range.start}" min="1" max="${totalPdfPages}" data-index="${index}">
      <label>To:</label>
      <input type="number" class="end-input" value="${range.end}" min="1" max="${totalPdfPages}" data-index="${index}">
      <label>Name:</label>
      <input type="text" class="name-input" value="${range.name}" placeholder="filename" data-index="${index}">
      <button class="range-remove" data-index="${index}">✕ Remove</button>
    `;

    rangeList.appendChild(item);
  });

  // Attach event listeners
  document.querySelectorAll('.start-input').forEach((input) => {
    input.onchange = (e) => {
      const idx = parseInt(e.target.dataset.index);
      ranges[idx].start = parseInt(e.target.value) || 1;
    };
  });

  document.querySelectorAll('.end-input').forEach((input) => {
    input.onchange = (e) => {
      const idx = parseInt(e.target.dataset.index);
      ranges[idx].end = parseInt(e.target.value) || totalPdfPages;
    };
  });

  document.querySelectorAll('.name-input').forEach((input) => {
    input.onchange = (e) => {
      const idx = parseInt(e.target.dataset.index);
      ranges[idx].name = e.target.value || `part${idx + 1}`;
    };
  });

  document.querySelectorAll('.range-remove').forEach((btn) => {
    btn.onclick = (e) => {
      const idx = parseInt(e.target.dataset.index);
      ranges.splice(idx, 1);
      renderRanges();
    };
  });
}

addRangeBtn.onclick = () => {
  const newRange = {
    start: 1,
    end: totalPdfPages,
    name: `part${ranges.length + 1}`,
  };
  ranges.push(newRange);
  renderRanges();
};

pagesPerFileInput.oninput = () => {
  const val = pagesPerFileInput.value;
  pagesDisplay.textContent = val;
};

function updateEachPageInfo() {
  eachPageInfo.textContent = `Will create ${totalPdfPages} individual PDF files`;
}

splitBtn.onclick = async () => {
  if (!selectedPdf) {
    alert('Please select a PDF file first');
    return;
  }

  try {
    splitBtn.disabled = true;
    splitBtn.textContent = '⏳ Splitting...';

    const outputDir = await window.pdfAPI.selectFolder();
    if (!outputDir) {
      splitBtn.disabled = false;
      splitBtn.innerHTML = '<span>✂️</span> Split PDF';
      return;
    }

    let splitRanges = [];

    if (currentSplitMode === 'range') {
      if (ranges.length === 0) {
        alert('Add at least one range to split');
        splitBtn.disabled = false;
        splitBtn.innerHTML = '<span>✂️</span> Split PDF';
        return;
      }
      splitRanges = ranges;
    } else if (currentSplitMode === 'fixed') {
      const pagesPerFile = parseInt(pagesPerFileInput.value) || 5;
      const prefix = fixedPrefixInput.value.trim() || 'part';
      const suffix = fixedSuffixInput.value.trim();
      let partNum = 1;
      for (let i = 1; i <= totalPdfPages; i += pagesPerFile) {
        splitRanges.push({
          start: i,
          end: Math.min(i + pagesPerFile - 1, totalPdfPages),
          name: `${prefix}${partNum}${suffix}`,
        });
        partNum++;
      }
    } else if (currentSplitMode === 'each') {
      const prefix = eachPrefixInput.value.trim() || 'page';
      const suffix = eachSuffixInput.value.trim();
      for (let i = 1; i <= totalPdfPages; i++) {
        splitRanges.push({
          start: i,
          end: i,
          name: `${prefix}${i}${suffix}`,
        });
      }
    }

    const results = await window.pdfAPI.split(selectedPdf, splitRanges, outputDir);

    alert(`✅ PDF split successfully!\nCreated ${results.length} file(s)`);

    // Reset
    selectedPdf = null;
    totalPdfPages = 0;
    ranges = [];
    splitFileInput.value = '';
    pdfInfo.classList.remove('active');
    splitOptions.style.display = 'none';
    splitBtn.innerHTML = '<span>✂️</span> Split PDF';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to split PDF: ' + err.message);
    splitBtn.disabled = false;
    splitBtn.innerHTML = '<span>✂️</span> Split PDF';
  }
};

// Organize Tool Logic
const selectOrganizePdfBtn = document.getElementById('selectOrganizePdfBtn');
const organizeFileInput = document.getElementById('organizeFileInput');
const organizeDropzone = document.getElementById('organizeDropzone');
const organizePdfInfo = document.getElementById('organizePdfInfo');
const organizeOptions = document.getElementById('organizeOptions');
const pageGrid = document.getElementById('pageGrid');
const selectedCount = document.getElementById('selectedCount');
const rotateLeftBtn = document.getElementById('rotateLeftBtn');
const rotateRightBtn = document.getElementById('rotateRightBtn');
const deleteSelectedBtn = document.getElementById('deleteSelectedBtn');
const organizeBtn = document.getElementById('organizeBtn');
const scrollLeftBtn = document.getElementById('scrollLeftBtn');
const scrollRightBtn = document.getElementById('scrollRightBtn');

let organizePages = [];
let selectedOrganizePdf = null;

// Page grid scroll navigation
if (scrollLeftBtn && scrollRightBtn) {
  scrollLeftBtn.onclick = () => {
    pageGrid.scrollBy({ left: -250, behavior: 'smooth' });
  };
  
  scrollRightBtn.onclick = () => {
    pageGrid.scrollBy({ left: 250, behavior: 'smooth' });
  };
  
  // Update button states based on scroll position
  pageGrid.addEventListener('scroll', () => {
    scrollLeftBtn.disabled = pageGrid.scrollLeft <= 0;
    scrollRightBtn.disabled = pageGrid.scrollLeft >= (pageGrid.scrollWidth - pageGrid.clientWidth - 1);
  });
}

// Drag and drop handlers
organizeDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  organizeDropzone.classList.add('drag-over');
});

organizeDropzone.addEventListener('dragleave', () => {
  organizeDropzone.classList.remove('drag-over');
});

organizeDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  organizeDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadOrganizePDF(file.path);
  } else {
    alert('❌ Please drop a PDF file');
  }
});

selectOrganizePdfBtn.onclick = () => {
  organizeFileInput.click();
};

organizeFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadOrganizePDF(file.path);
  }
};

async function loadOrganizePDF(filePath) {
  selectedOrganizePdf = filePath;

  try {
    organizePdfInfo.innerHTML = `
      <div class="info-icon">📑</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>Loading pages...</p>
      </div>
    `;
    organizePdfInfo.classList.add('active');

    const thumbnails = await window.pdfAPI.getPDFThumbnails(filePath);
    
    // Extract filename from path
    const fileName = filePath.split(/[\\\/]/).pop();

    organizePages = thumbnails.map((thumb, index) => ({
      pageIndex: index,
      originalPageNumber: index + 1, // Store original page number
      thumbnail: thumb,
      rotation: 0,
      deleted: false,
      selected: false,
      newPosition: index,
    }));

    organizePdfInfo.innerHTML = `
      <div class="info-icon">📑</div>
      <div class="info-content">
        <strong>${fileName}</strong>
        <p>${organizePages.length} pages loaded</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('organizeFileInput').click()">Change</button>
    `;

    organizeOptions.style.display = 'block';
    renderPages();
  } catch (err) {
    console.error(err);
    alert('❌ Failed to load PDF: ' + err.message);
  }
}

function renderPages() {
  pageGrid.innerHTML = '';

  const sortedPages = [...organizePages].sort((a, b) => a.newPosition - b.newPosition);

  sortedPages.forEach((page, displayIndex) => {
    const card = document.createElement('div');
    card.className = 'page-card';
    card.draggable = !page.deleted;
    card.dataset.pageIndex = page.pageIndex;

    if (page.selected) card.classList.add('selected');
    if (page.deleted) card.classList.add('deleted');

    const rotationStyle = page.rotation !== 0 ? `transform: rotate(${page.rotation}deg);` : '';
    const rotationBadge = page.rotation !== 0 ? `<span class="rotation-badge">${page.rotation}°</span>` : '';

    // Show both: current position and original page number
    const positionLabel = displayIndex + 1 !== page.originalPageNumber
      ? `Pos ${displayIndex + 1} (Pg ${page.originalPageNumber})`
      : `Page ${page.originalPageNumber}`;

    card.innerHTML = `
      <input type="checkbox" class="page-checkbox" ${page.deleted ? 'disabled' : ''} ${page.selected ? 'checked' : ''}>
      <div class="page-number-badge">${page.originalPageNumber}</div>
      <img src="${page.thumbnail}" class="page-thumbnail" style="${rotationStyle}" alt="Page ${page.originalPageNumber}" draggable="false">
      <div class="page-info">
        <span class="page-number">${positionLabel}</span>
        ${rotationBadge}
      </div>
    `;

    // Checkbox handler
    const checkbox = card.querySelector('.page-checkbox');
    checkbox.onclick = (e) => {
      e.stopPropagation();
      page.selected = checkbox.checked;
      updateSelectedCount();
      card.classList.toggle('selected', page.selected);
    };

    // Prevent checkbox from interfering with drag
    checkbox.addEventListener('mousedown', (e) => {
      e.stopPropagation();
    });

    // Drag handlers - attach to the card
    card.addEventListener('dragstart', handleOrganizeDragStart);
    card.addEventListener('dragover', handleOrganizeDragOver);
    card.addEventListener('drop', handleOrganizeDrop);
    card.addEventListener('dragend', handleOrganizeDragEnd);
    card.addEventListener('dragenter', handleOrganizeDragEnter);
    card.addEventListener('dragleave', handleOrganizeDragLeave);

    pageGrid.appendChild(card);
  });

  updateSelectedCount();
}

let draggedPageIndex = null;

function handleOrganizeDragStart(e) {
  draggedPageIndex = parseInt(e.target.dataset.pageIndex);
  e.target.classList.add('dragging');
}

function handleOrganizeDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
}

function handleOrganizeDragEnter(e) {
  e.preventDefault();
  const target = e.currentTarget;
  if (target && !target.classList.contains('dragging')) {
    target.style.opacity = '0.5';
  }
}

function handleOrganizeDragLeave(e) {
  const target = e.currentTarget;
  if (target && !target.classList.contains('dragging')) {
    target.style.opacity = '1';
  }
}

function handleOrganizeDrop(e) {
  e.preventDefault();
  e.stopPropagation();

  const target = e.currentTarget;
  target.style.opacity = '1';

  if (target && draggedPageIndex !== null) {
    const targetIndex = parseInt(target.dataset.pageIndex);

    if (draggedPageIndex !== targetIndex) {
      const draggedPage = organizePages.find((p) => p.pageIndex === draggedPageIndex);
      const targetPage = organizePages.find((p) => p.pageIndex === targetIndex);

      if (draggedPage && targetPage) {
        // Get the current positions
        const draggedPosition = draggedPage.newPosition;
        const targetPosition = targetPage.newPosition;

        // Reorder all pages between dragged and target
        if (draggedPosition < targetPosition) {
          // Moving forward
          organizePages.forEach(page => {
            if (page.newPosition > draggedPosition && page.newPosition <= targetPosition) {
              page.newPosition--;
            }
          });
          draggedPage.newPosition = targetPosition;
        } else {
          // Moving backward
          organizePages.forEach(page => {
            if (page.newPosition >= targetPosition && page.newPosition < draggedPosition) {
              page.newPosition++;
            }
          });
          draggedPage.newPosition = targetPosition;
        }

        renderPages();
      }
    }
  }
}

function handleOrganizeDragEnd(e) {
  e.target.classList.remove('dragging');
  e.target.style.opacity = '1';

  // Reset all opacities
  document.querySelectorAll('.page-card').forEach(card => {
    card.style.opacity = '1';
  });

  draggedPageIndex = null;
}

function updateSelectedCount() {
  const count = organizePages.filter((p) => p.selected && !p.deleted).length;
  selectedCount.textContent = count;
}

rotateLeftBtn.onclick = () => {
  organizePages.forEach((page) => {
    if (page.selected && !page.deleted) {
      page.rotation = (page.rotation - 90 + 360) % 360;
    }
  });
  renderPages();
};

rotateRightBtn.onclick = () => {
  organizePages.forEach((page) => {
    if (page.selected && !page.deleted) {
      page.rotation = (page.rotation + 90) % 360;
    }
  });
  renderPages();
};

deleteSelectedBtn.onclick = () => {
  const selectedPages = organizePages.filter((p) => p.selected && !p.deleted);
  if (selectedPages.length === 0) {
    alert('Please select pages to delete');
    return;
  }

  if (confirm(`Delete ${selectedPages.length} page(s)?`)) {
    selectedPages.forEach((page) => {
      page.deleted = true;
      page.selected = false;
    });
    renderPages();
  }
};

organizeBtn.onclick = async () => {
  if (!selectedOrganizePdf) {
    alert('No PDF loaded');
    return;
  }

  const activePages = organizePages.filter((p) => !p.deleted);
  if (activePages.length === 0) {
    alert('Cannot save PDF with all pages deleted');
    return;
  }

  try {
    organizeBtn.disabled = true;
    organizeBtn.textContent = '⏳ Saving...';

    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      organizeBtn.disabled = false;
      organizeBtn.innerHTML = '<span>💾</span> Save Organized PDF';
      return;
    }

    const operations = organizePages.map((page) => ({
      pageIndex: page.pageIndex,
      rotation: page.rotation,
      delete: page.deleted,
      newPosition: page.newPosition,
    }));

    await window.pdfAPI.organize(selectedOrganizePdf, operations, savePath);

    alert('✅ PDF organized successfully!');

    // Reset
    selectedOrganizePdf = null;
    organizePages = [];
    organizeFileInput.value = '';
    organizePdfInfo.classList.remove('active');
    organizeOptions.style.display = 'none';
    organizeBtn.innerHTML = '<span>💾</span> Save Organized PDF';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to organize PDF: ' + err.message);
    organizeBtn.disabled = false;
    organizeBtn.innerHTML = '<span>💾</span> Save Organized PDF';
  }
};

// Watermark Tool Logic
const selectWatermarkPdfBtn = document.getElementById('selectWatermarkPdfBtn');
const watermarkFileInput = document.getElementById('watermarkFileInput');
const watermarkDropzone = document.getElementById('watermarkDropzone');
const watermarkPdfInfo = document.getElementById('watermarkPdfInfo');
const watermarkOptions = document.getElementById('watermarkOptions');
const watermarkBtn = document.getElementById('watermarkBtn');
const watermarkText = document.getElementById('watermarkText');
const watermarkFontSize = document.getElementById('watermarkFontSize');
const watermarkOpacity = document.getElementById('watermarkOpacity');
const watermarkRotation = document.getElementById('watermarkRotation');
const watermarkColor = document.getElementById('watermarkColor');
const watermarkScale = document.getElementById('watermarkScale');
const opacityValue = document.getElementById('opacityValue');
const scaleValue = document.getElementById('scaleValue');
const previewText = document.getElementById('previewText');
const previewImage = document.getElementById('previewImage');
const watermarkPreview = document.getElementById('watermarkPreview');
const selectImageBtn = document.getElementById('selectImageBtn');
const watermarkImageInput = document.getElementById('watermarkImageInput');
const selectedImageName = document.getElementById('selectedImageName');
const textOptions = document.getElementById('textOptions');
const imageOptions = document.getElementById('imageOptions');
const customPosition = document.getElementById('customPosition');
const customX = document.getElementById('customX');
const customY = document.getElementById('customY');

let selectedWatermarkPdf = null;
let selectedPosition = 'center';
let watermarkType = 'text';
let selectedImagePath = null;
let pdfPageDimensions = { width: 595, height: 842 }; // Default A4 dimensions

// Drag and drop handlers
watermarkDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  watermarkDropzone.classList.add('drag-over');
});

watermarkDropzone.addEventListener('dragleave', () => {
  watermarkDropzone.classList.remove('drag-over');
});

watermarkDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  watermarkDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadWatermarkPDF(file.path);
  } else {
    alert('❌ Please drop a PDF file');
  }
});

selectWatermarkPdfBtn.onclick = () => {
  watermarkFileInput.click();
};

watermarkFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadWatermarkPDF(file.path);
  }
};

async function loadWatermarkPDF(filePath) {
  selectedWatermarkPdf = filePath;

  try {
    const pageCount = await window.pdfAPI.getPDFPageCount(filePath);
    
    // Try to get dimensions, fallback to default A4 if it fails
    let dimensionText = '';
    if (window.pdfAPI.getPDFPageDimensions) {
      try {
        const dimensions = await window.pdfAPI.getPDFPageDimensions(filePath);
        pdfPageDimensions = dimensions;
        dimensionText = ` • ${Math.round(dimensions.width)}×${Math.round(dimensions.height)}pt`;
        
        // Update preview container aspect ratio to match PDF
        const aspectRatio = (dimensions.height / dimensions.width) * 100;
        watermarkPreview.style.setProperty('--aspect-ratio', aspectRatio.toFixed(2));
      } catch (dimError) {
        console.warn('Could not get PDF dimensions, using default A4:', dimError);
        pdfPageDimensions = { width: 595, height: 842 };
        watermarkPreview.style.setProperty('--aspect-ratio', '141.4');
      }
    } else {
      // API not available yet, use defaults
      pdfPageDimensions = { width: 595, height: 842 };
      watermarkPreview.style.setProperty('--aspect-ratio', '141.4');
    }

    watermarkPdfInfo.innerHTML = `
      <div class="info-icon">💧</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${pageCount} pages${dimensionText}</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('watermarkFileInput').click()">Change</button>
    `;
    watermarkPdfInfo.classList.add('active');
    watermarkOptions.style.display = 'block';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to load PDF: ' + err.message);
  }
}

// Type selector
document.querySelectorAll('.type-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.type-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    watermarkType = btn.dataset.type;

    if (watermarkType === 'text') {
      textOptions.style.display = 'block';
      imageOptions.style.display = 'none';
    } else {
      textOptions.style.display = 'none';
      imageOptions.style.display = 'block';
    }

    updateWatermarkPreview();
  });
});

// Image selection
selectImageBtn.onclick = () => {
  watermarkImageInput.click();
};

watermarkImageInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    selectedImagePath = file.path;
    selectedImageName.textContent = file.name;

    // Load image for preview
    const reader = new FileReader();
    reader.onload = (event) => {
      previewImage.src = event.target.result;
      updateWatermarkPreview();
    };
    reader.readAsDataURL(file);
  }
};

// Position buttons
document.querySelectorAll('.position-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.position-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    selectedPosition = btn.dataset.position;

    if (selectedPosition === 'custom') {
      customPosition.style.display = 'block';
    } else {
      customPosition.style.display = 'none';
    }

    updateWatermarkPreview();
  });
});

// Preview updates
const watermarkFont = document.getElementById('watermarkFont');
watermarkFont.oninput = updateWatermarkPreview;
watermarkText.oninput = updateWatermarkPreview;
watermarkFontSize.oninput = updateWatermarkPreview;
watermarkOpacity.oninput = () => {
  opacityValue.textContent = watermarkOpacity.value + '%';
  updateWatermarkPreview();
};
watermarkScale.oninput = () => {
  let value = parseInt(watermarkScale.value) || 30;

  // Limit to range 1 to 500
  if (value > 500) {
    watermarkScale.value = 500;
  } else if (value < 1) {
    watermarkScale.value = 1;
  }

  updateWatermarkPreview();
};
watermarkRotation.oninput = () => {
  let value = parseInt(watermarkRotation.value) || 0;

  // Limit to range -360 to 360
  if (value > 360) {
    watermarkRotation.value = 360;
  } else if (value < -360) {
    watermarkRotation.value = -360;
  }

  updateWatermarkPreview();
};
watermarkColor.oninput = updateWatermarkPreview;
customX.oninput = updateWatermarkPreview;
customY.oninput = updateWatermarkPreview;

function updateWatermarkPreview() {
  const opacity = watermarkOpacity.value / 100;
  const rotation = watermarkRotation.value;
  const margin = 10; // margin percentage from edge

  if (watermarkType === 'text') {
    const text = watermarkText.value || 'WATERMARK';
    const fontSize = watermarkFontSize.value;
    const font = watermarkFont.value;
    const color = watermarkColor.value;

    previewText.style.display = 'block';
    previewImage.style.display = 'none';

    previewText.textContent = text;
    previewText.style.fontSize = `${Math.min(fontSize / 1.45, 100)}px`;
    previewText.style.opacity = opacity;
    previewText.style.color = color;

    // Set font family based on selection
    if (font.includes('Times')) {
      previewText.style.fontFamily = 'Times New Roman, serif';
    } else if (font.includes('Courier')) {
      previewText.style.fontFamily = 'Courier New, monospace';
    } else if (font.includes('Symbol') || font.includes('ZapfDingbats')) {
      previewText.style.fontFamily = 'Symbol, serif';
    } else {
      previewText.style.fontFamily = 'Arial, Helvetica, sans-serif';
    }

    // Set font weight and style
    previewText.style.fontWeight = font.includes('Bold') ? 'bold' : 'normal';
    previewText.style.fontStyle = (font.includes('Italic') || font.includes('Oblique')) ? 'italic' : 'normal';

    // Position calculation - rotation always applied
    let positionTransform = '';
    if (selectedPosition === 'custom') {
      const x = customX.value || 50;
      const y = customY.value || 50;
      previewText.style.left = `${x}%`;
      previewText.style.top = `${y}%`;
      positionTransform = 'translate(-50%, -50%)';
    } else {
      // Preset positions
      switch (selectedPosition) {
        case 'center':
        case 'diagonal':
          previewText.style.left = '50%';
          previewText.style.top = '50%';
          positionTransform = 'translate(-50%, -50%)';
          break;
        case 'top':
          previewText.style.left = '50%';
          previewText.style.top = `${margin}%`;
          positionTransform = 'translate(-50%, 0)';
          break;
        case 'bottom':
          previewText.style.left = '50%';
          previewText.style.top = `${100 - margin}%`;
          positionTransform = 'translate(-50%, -100%)';
          break;
        case 'topLeft':
          previewText.style.left = `${margin}%`;
          previewText.style.top = `${margin}%`;
          positionTransform = 'translate(0, 0)';
          break;
        case 'topRight':
          previewText.style.left = `${100 - margin}%`;
          previewText.style.top = `${margin}%`;
          positionTransform = 'translate(-100%, 0)';
          break;
        case 'bottomLeft':
          previewText.style.left = `${margin}%`;
          previewText.style.top = `${100 - margin}%`;
          positionTransform = 'translate(0, -100%)';
          break;
        case 'bottomRight':
          previewText.style.left = `${100 - margin}%`;
          previewText.style.top = `${100 - margin}%`;
          positionTransform = 'translate(-100%, -100%)';
          break;
        default:
          previewText.style.left = '50%';
          previewText.style.top = '50%';
          positionTransform = 'translate(-50%, -50%)';
      }
    }
    previewText.style.transform = `${positionTransform} rotate(${rotation}deg)`;
  } else {
    previewText.style.display = 'none';
    previewImage.style.display = 'block';

    if (selectedImagePath) {
      // Scale down for preview to match PDF output (preview is smaller than actual PDF)
      const scale = (watermarkScale.value / 100) * 6.4;
      previewImage.style.opacity = opacity;

      // Position calculation for image - rotation and scale always applied
      let positionTransform = '';
      if (selectedPosition === 'custom') {
        const x = customX.value || 50;
        const y = customY.value || 50;
        previewImage.style.left = `${x}%`;
        previewImage.style.top = `${y}%`;
        positionTransform = 'translate(-50%, -50%)';
      } else {
        // Preset positions
        switch (selectedPosition) {
          case 'center':
          case 'diagonal':
            previewImage.style.left = '50%';
            previewImage.style.top = '50%';
            positionTransform = 'translate(-50%, -50%)';
            break;
          case 'top':
            previewImage.style.left = '50%';
            previewImage.style.top = `${margin}%`;
            positionTransform = 'translate(-50%, 0)';
            break;
          case 'bottom':
            previewImage.style.left = '50%';
            previewImage.style.top = `${100 - margin}%`;
            positionTransform = 'translate(-50%, -100%)';
            break;
          case 'topLeft':
            previewImage.style.left = `${margin}%`;
            previewImage.style.top = `${margin}%`;
            positionTransform = 'translate(0, 0)';
            break;
          case 'topRight':
            previewImage.style.left = `${100 - margin}%`;
            previewImage.style.top = `${margin}%`;
            positionTransform = 'translate(-100%, 0)';
            break;
          case 'bottomLeft':
            previewImage.style.left = `${margin}%`;
            previewImage.style.top = `${100 - margin}%`;
            positionTransform = 'translate(0, -100%)';
            break;
          case 'bottomRight':
            previewImage.style.left = `${100 - margin}%`;
            previewImage.style.top = `${100 - margin}%`;
            positionTransform = 'translate(-100%, -100%)';
            break;
          default:
            previewImage.style.left = '50%';
            previewImage.style.top = '50%';
            positionTransform = 'translate(-50%, -50%)';
        }
      }
      previewImage.style.transform = `${positionTransform} rotate(${rotation}deg) scale(${scale})`;
    }
  }
}

watermarkBtn.onclick = async () => {
  if (!selectedWatermarkPdf) {
    alert('Please select a PDF file first');
    return;
  }

  if (watermarkType === 'text' && !watermarkText.value.trim()) {
    alert('Please enter watermark text');
    return;
  }

  if (watermarkType === 'image' && !selectedImagePath) {
    alert('Please select a watermark image');
    return;
  }

  try {
    watermarkBtn.disabled = true;
    watermarkBtn.textContent = '⏳ Adding Watermark...';

    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      watermarkBtn.disabled = false;
      watermarkBtn.innerHTML = '<span>💧</span> Add Watermark';
      return;
    }

    // Normalize rotation to 0-360 range for backend
    let rotation = parseInt(watermarkRotation.value) || 0;
    if (rotation < 0) {
      rotation = 360 + (rotation % 360);
    }

    const options = {
      type: watermarkType,
      opacity: parseFloat(watermarkOpacity.value) / 100,
      rotation: rotation,
      position: selectedPosition,
    };

    if (watermarkType === 'text') {
      // Convert hex color to RGB
      const hex = watermarkColor.value;
      const r = parseInt(hex.slice(1, 3), 16) / 255;
      const g = parseInt(hex.slice(3, 5), 16) / 255;
      const b = parseInt(hex.slice(5, 7), 16) / 255;

      options.text = watermarkText.value;
      options.fontSize = parseInt(watermarkFontSize.value);
      options.fontName = watermarkFont.value;
      options.color = { r, g, b };
    } else {
      options.imagePath = selectedImagePath;
      options.scale = parseFloat(watermarkScale.value) / 100;
    }

    if (selectedPosition === 'custom') {
      options.x = parseInt(customX.value);
      options.y = parseInt(customY.value);
    }

    await window.pdfAPI.watermark(selectedWatermarkPdf, options, savePath);

    alert('✅ Watermark added successfully!');

    // Reset
    selectedWatermarkPdf = null;
    selectedImagePath = null;
    watermarkFileInput.value = '';
    watermarkImageInput.value = '';
    selectedImageName.textContent = 'No image selected';
    watermarkPdfInfo.classList.remove('active');
    watermarkOptions.style.display = 'none';
    watermarkBtn.disabled = false;
    watermarkBtn.innerHTML = '<span>💧</span> Add Watermark';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to add watermark: ' + err.message);
    watermarkBtn.disabled = false;
    watermarkBtn.innerHTML = '<span>💧</span> Add Watermark';
  }
};

// Initialize preview
updateWatermarkPreview();

// Page Numbers Tool Logic
const selectPageNumbersPdfBtn = document.getElementById('selectPageNumbersPdfBtn');
const pageNumbersFileInput = document.getElementById('pageNumbersFileInput');
const pageNumbersDropzone = document.getElementById('pageNumbersDropzone');
const pageNumbersPdfInfo = document.getElementById('pageNumbersPdfInfo');
const pageNumbersOptions = document.getElementById('pageNumbersOptions');
const pageNumbersBtn = document.getElementById('pageNumbersBtn');
const pageNumberFont = document.getElementById('pageNumberFont');
const pageNumberFontSize = document.getElementById('pageNumberFontSize');
const pageNumberColor = document.getElementById('pageNumberColor');
const pageNumberStyle = document.getElementById('pageNumberStyle');
const pageNumberFormat = document.getElementById('pageNumberFormat');
const pageNumberStart = document.getElementById('pageNumberStart');
const pageNumberMargin = document.getElementById('pageNumberMargin');
const pageNumberSkip = document.getElementById('pageNumberSkip');
const pageNumberExclude = document.getElementById('pageNumberExclude');
const pageNumberPlacementIndicator = document.getElementById('pageNumberPlacementIndicator');
const pageNumberFontPreview = document.getElementById('pageNumberFontPreview');

let selectedPageNumbersPdf = null;
let selectedPageNumberPosition = 'bottom';
let pageNumbersPdfDimensions = { width: 595, height: 842 }; // Default A4

// Update previews
function updatePageNumberPreviews() {
  const fontSize = pageNumberFontSize.value;
  const color = pageNumberColor.value;
  const style = pageNumberStyle.value;
  const format = pageNumberFormat.value;
  const font = pageNumberFont.value;
  const margin = 8; // percentage margin for preview

  // Convert number style for preview
  function getStyledNumber(num, style) {
    switch (style) {
      case 'upperRoman': return 'I';
      case 'lowerRoman': return 'i';
      case 'upperAlpha': return 'A';
      case 'lowerAlpha': return 'a';
      default: return '1';
    }
  }

  const styledNum = getStyledNumber(1, style);

  // Update font preview text based on format
  let previewText;
  switch (format) {
    case 'pageOfTotal':
      previewText = `Page ${styledNum} of 10`;
      break;
    case 'pageX':
      previewText = `Page ${styledNum}`;
      break;
    case 'brackets':
      previewText = `[${styledNum}]`;
      break;
    case 'dashes':
      previewText = `- ${styledNum} -`;
      break;
    default:
      previewText = styledNum;
  }

  // Update font preview
  pageNumberFontPreview.textContent = previewText;
  pageNumberFontPreview.style.fontSize = `${Math.min(fontSize * 1.5, 32)}px`;
  pageNumberFontPreview.style.color = color;

  // Set font family based on selection
  if (font.includes('Times')) {
    pageNumberFontPreview.style.fontFamily = 'Times New Roman, serif';
  } else if (font.includes('Courier')) {
    pageNumberFontPreview.style.fontFamily = 'Courier New, monospace';
  } else if (font.includes('Symbol') || font.includes('ZapfDingbats')) {
    pageNumberFontPreview.style.fontFamily = 'Symbol, serif';
  } else {
    pageNumberFontPreview.style.fontFamily = 'Arial, Helvetica, sans-serif';
  }

  // Set font weight and style
  pageNumberFontPreview.style.fontWeight = font.includes('Bold') ? 'bold' : 'normal';
  pageNumberFontPreview.style.fontStyle = (font.includes('Italic') || font.includes('Oblique')) ? 'italic' : 'normal';

  // Update placement indicator
  pageNumberPlacementIndicator.textContent = styledNum;
  pageNumberPlacementIndicator.style.fontSize = `${Math.min(fontSize * 0.8, 14)}px`;
  pageNumberPlacementIndicator.style.color = color;

  // Position based on selected position
  pageNumberPlacementIndicator.style.left = '';
  pageNumberPlacementIndicator.style.right = '';
  pageNumberPlacementIndicator.style.top = '';
  pageNumberPlacementIndicator.style.bottom = '';
  pageNumberPlacementIndicator.style.transform = '';

  switch (selectedPageNumberPosition) {
    case 'top':
      pageNumberPlacementIndicator.style.top = `${margin}%`;
      pageNumberPlacementIndicator.style.left = '50%';
      pageNumberPlacementIndicator.style.transform = 'translateX(-50%)';
      break;
    case 'topLeft':
      pageNumberPlacementIndicator.style.top = `${margin}%`;
      pageNumberPlacementIndicator.style.left = `${margin}%`;
      break;
    case 'topRight':
      pageNumberPlacementIndicator.style.top = `${margin}%`;
      pageNumberPlacementIndicator.style.right = `${margin}%`;
      break;
    case 'bottomLeft':
      pageNumberPlacementIndicator.style.bottom = `${margin}%`;
      pageNumberPlacementIndicator.style.left = `${margin}%`;
      break;
    case 'bottomRight':
      pageNumberPlacementIndicator.style.bottom = `${margin}%`;
      pageNumberPlacementIndicator.style.right = `${margin}%`;
      break;
    case 'bottom':
    default:
      pageNumberPlacementIndicator.style.bottom = `${margin}%`;
      pageNumberPlacementIndicator.style.left = '50%';
      pageNumberPlacementIndicator.style.transform = 'translateX(-50%)';
  }
}

// Drag and drop handlers
pageNumbersDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  pageNumbersDropzone.classList.add('drag-over');
});

pageNumbersDropzone.addEventListener('dragleave', () => {
  pageNumbersDropzone.classList.remove('drag-over');
});

pageNumbersDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  pageNumbersDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadPageNumbersPDF(file.path);
  } else {
    alert('❌ Please drop a PDF file');
  }
});

selectPageNumbersPdfBtn.onclick = () => {
  pageNumbersFileInput.click();
};

pageNumbersFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadPageNumbersPDF(file.path);
  }
};

async function loadPageNumbersPDF(filePath) {
  selectedPageNumbersPdf = filePath;

  try {
    const pageCount = await window.pdfAPI.getPDFPageCount(filePath);
    
    // Fetch actual PDF page dimensions
    pageNumbersPdfDimensions = await window.pdfAPI.getPDFPageDimensions(filePath);
    
    // Update preview aspect ratio based on actual page dimensions
    const previewContainer = document.querySelector('.pageNumber-placement-preview');
    if (previewContainer) {
      const aspectRatio = pageNumbersPdfDimensions.height / pageNumbersPdfDimensions.width;
      previewContainer.style.setProperty('--aspect-ratio', aspectRatio);
    }

    pageNumbersPdfInfo.innerHTML = `
      <div class="info-icon">🔢</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${pageCount} pages • Ready for page numbers</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('pageNumbersFileInput').click()">Change</button>
    `;
    pageNumbersPdfInfo.classList.add('active');
    pageNumbersOptions.style.display = 'block';
    updateFormatOptions();
    updatePageNumberPreviews();
  } catch (err) {
    console.error(err);
    alert('❌ Failed to load PDF: ' + err.message);
  }
}

// Position buttons for page numbers
document.querySelectorAll('#pageNumbersOptions .position-btn').forEach((btn) => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#pageNumbersOptions .position-btn').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');
    selectedPageNumberPosition = btn.dataset.position;
    updatePageNumberPreviews();
  });
});

// Update preview on input changes
// Update format options based on style
function updateFormatOptions() {
  const style = pageNumberStyle.value;
  const isNonDecimal = style !== 'decimal';

  // Get the pageOfTotal option
  const pageOfTotalOption = pageNumberFormat.querySelector('option[value="pageOfTotal"]');

  if (isNonDecimal) {
    // Hide "Page N of T" for Roman and alphabetical styles
    if (pageOfTotalOption) {
      pageOfTotalOption.style.display = 'none';
      // If currently selected, switch to default
      if (pageNumberFormat.value === 'pageOfTotal') {
        pageNumberFormat.value = 'number';
      }
    }
  } else {
    // Show "Page N of T" for decimal style
    if (pageOfTotalOption) {
      pageOfTotalOption.style.display = '';
    }
  }
}

pageNumberFont.oninput = updatePageNumberPreviews;
pageNumberFontSize.oninput = updatePageNumberPreviews;
pageNumberColor.oninput = updatePageNumberPreviews;
pageNumberStyle.oninput = () => {
  updateFormatOptions();
  updatePageNumberPreviews();
};
pageNumberFormat.oninput = updatePageNumberPreviews;

pageNumbersBtn.onclick = async () => {
  if (!selectedPageNumbersPdf) {
    alert('Please select a PDF file first');
    return;
  }

  try {
    pageNumbersBtn.disabled = true;
    pageNumbersBtn.textContent = '⏳ Adding Page Numbers...';

    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      pageNumbersBtn.disabled = false;
      pageNumbersBtn.innerHTML = '<span>🔢</span> Add Page Numbers';
      return;
    }

    // Convert hex color to RGB
    const hex = pageNumberColor.value;
    const r = parseInt(hex.slice(1, 3), 16) / 255;
    const g = parseInt(hex.slice(3, 5), 16) / 255;
    const b = parseInt(hex.slice(5, 7), 16) / 255;

    // Parse skip pages (hide number but count)
    const skipPagesInput = pageNumberSkip.value.trim();
    const skipPages = skipPagesInput
      ? skipPagesInput.split(',').map(p => parseInt(p.trim())).filter(p => !isNaN(p) && p > 0)
      : [];

    // Parse exclude pages (don't show number and don't count)
    const excludePagesInput = pageNumberExclude.value.trim();
    const excludePages = excludePagesInput
      ? excludePagesInput.split(',').map(p => parseInt(p.trim())).filter(p => !isNaN(p) && p > 0)
      : [];

    const options = {
      position: selectedPageNumberPosition,
      fontSize: parseInt(pageNumberFontSize.value),
      color: { r, g, b },
      style: pageNumberStyle.value,
      format: pageNumberFormat.value,
      fontStyle: pageNumberFont.value,
      startFrom: parseInt(pageNumberStart.value),
      margin: parseInt(pageNumberMargin.value),
      skipPages: skipPages,
      excludePages: excludePages,
    };

    await window.pdfAPI.pageNumbers(selectedPageNumbersPdf, options, savePath);

    alert('✅ Page numbers added successfully!');

    // Reset
    selectedPageNumbersPdf = null;
    pageNumbersFileInput.value = '';
    pageNumbersPdfInfo.classList.remove('active');
    pageNumbersOptions.style.display = 'none';
    pageNumbersBtn.disabled = false;
    pageNumbersBtn.innerHTML = '<span>🔢</span> Add Page Numbers';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to add page numbers: ' + err.message);
    pageNumbersBtn.disabled = false;
    pageNumbersBtn.innerHTML = '<span>🔢</span> Add Page Numbers';
  }
};

// PDF to Images Tool Logic
const pdfToImageDropzone = document.getElementById('pdfToImageDropzone');
const selectPdfToImageBtn = document.getElementById('selectPdfToImageBtn');
const pdfToImageFileInput = document.getElementById('pdfToImageFileInput');
const pdfToImagePdfInfo = document.getElementById('pdfToImagePdfInfo');
const pdfToImageOptions = document.getElementById('pdfToImageOptions');
const imageFormat = document.getElementById('imageFormat');
const imageQuality = document.getElementById('imageQuality');
const qualityGroup = document.getElementById('qualityGroup');
const imageDpi = document.getElementById('imageDpi');
const imagePageRange = document.getElementById('imagePageRange');
const selectOutputFolderBtn = document.getElementById('selectOutputFolderBtn');
const selectedFolderName = document.getElementById('selectedFolderName');
const convertToImagesBtn = document.getElementById('convertToImagesBtn');
const imageColorMode = document.getElementById('imageColorMode');
const imageNaming = document.getElementById('imageNaming');
const imageCustomPrefix = document.getElementById('imageCustomPrefix');
const customPrefixGroup = document.getElementById('customPrefixGroup');

let selectedPdfToImage = null;
let selectedOutputFolder = null;

// Show/hide quality option based on format
imageFormat.onchange = () => {
  const format = imageFormat.value;
  if (format === 'jpeg') {
    qualityGroup.style.display = 'block';
  } else {
    qualityGroup.style.display = 'none';
  }
};

// Show/hide custom prefix input
imageNaming.onchange = () => {
  if (imageNaming.value === 'custom') {
    customPrefixGroup.style.display = 'block';
  } else {
    customPrefixGroup.style.display = 'none';
  }
};

// Drag and drop handlers
pdfToImageDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  pdfToImageDropzone.classList.add('drag-over');
});

pdfToImageDropzone.addEventListener('dragleave', () => {
  pdfToImageDropzone.classList.remove('drag-over');
});

pdfToImageDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  pdfToImageDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadPdfToImage(file.path);
  } else {
    alert('❌ Please drop a PDF file');
  }
});

selectPdfToImageBtn.onclick = () => {
  pdfToImageFileInput.click();
};

pdfToImageFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadPdfToImage(file.path);
  }
};

async function loadPdfToImage(filePath) {
  selectedPdfToImage = filePath;

  try {
    const pageCount = await window.pdfAPI.getPDFPageCount(filePath);

    pdfToImagePdfInfo.innerHTML = `
      <div class="info-icon">📄</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${pageCount} pages • Ready to convert</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('pdfToImageFileInput').click()">Change</button>
    `;
    pdfToImagePdfInfo.classList.add('active');
    pdfToImageOptions.style.display = 'block';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to load PDF: ' + err.message);
  }
}

selectOutputFolderBtn.onclick = async () => {
  const folder = await window.pdfAPI.selectFolder();
  if (folder) {
    selectedOutputFolder = folder;
    selectedFolderName.textContent = folder;
    selectedFolderName.style.color = '#000';
  }
};

convertToImagesBtn.onclick = async () => {
  if (!selectedPdfToImage) {
    alert('Please select a PDF file first');
    return;
  }

  if (!selectedOutputFolder) {
    alert('Please select an output folder');
    return;
  }

  try {
    convertToImagesBtn.disabled = true;
    convertToImagesBtn.textContent = '⏳ Converting...';

    const options = {
      format: imageFormat.value,
      quality: parseInt(imageQuality.value),
      dpi: parseInt(imageDpi.value),
      pageRange: imagePageRange.value.trim() || 'all',
      colorMode: imageColorMode.value,
      naming: imageNaming.value,
      customPrefix: imageCustomPrefix.value.trim() || 'page',
      outputFolder: selectedOutputFolder
    };

    const files = await window.pdfAPI.pdfToImages(selectedPdfToImage, options);

    alert(`✅ Successfully converted to ${files.length} image(s)!`);

    // Reset
    selectedPdfToImage = null;
    selectedOutputFolder = null;
    pdfToImageFileInput.value = '';
    pdfToImagePdfInfo.classList.remove('active');
    pdfToImageOptions.style.display = 'none';
    selectedFolderName.textContent = 'No folder selected';
    selectedFolderName.style.color = '#999';
    convertToImagesBtn.disabled = false;
    convertToImagesBtn.innerHTML = '<span>🖼️</span> Convert to Images';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to convert PDF: ' + err.message);
    convertToImagesBtn.disabled = false;
    convertToImagesBtn.innerHTML = '<span>🖼️</span> Convert to Images';
  }
};

// Images to PDF Tool Logic
const imageToPdfDropzone = document.getElementById('imageToPdfDropzone');
const selectImagesBtn = document.getElementById('selectImagesBtn');
const imageToPdfFileInput = document.getElementById('imageToPdfFileInput');
const imageList = document.getElementById('imageList');
const imageToPdfOptions = document.getElementById('imageToPdfOptions');
const pdfPageSize = document.getElementById('pdfPageSize');
const pdfOrientation = document.getElementById('pdfOrientation');
const pdfMargin = document.getElementById('pdfMargin');
const pdfImageQuality = document.getElementById('pdfImageQuality');
const createPdfBtn = document.getElementById('createPdfBtn');

let selectedImages = [];

// Drag and drop handlers
imageToPdfDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  imageToPdfDropzone.classList.add('drag-over');
});

imageToPdfDropzone.addEventListener('dragleave', () => {
  imageToPdfDropzone.classList.remove('drag-over');
});

imageToPdfDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  imageToPdfDropzone.classList.remove('drag-over');

  const files = [...e.dataTransfer.files].filter(file =>
    file.type === 'image/png' || file.type === 'image/jpeg'
  );

  if (files.length > 0) {
    const paths = files.map(f => f.path);
    addImagesToList(paths);
  } else {
    alert('❌ Please drop PNG or JPEG images');
  }
});

selectImagesBtn.onclick = () => {
  imageToPdfFileInput.click();
};

imageToPdfFileInput.onchange = (e) => {
  const files = [...e.target.files];
  if (files.length > 0) {
    const paths = files.map(f => f.path);
    addImagesToList(paths);
  }
};

function addImagesToList(paths) {
  selectedImages = [...selectedImages, ...paths];
  renderImageList();
  imageList.style.display = 'block';
  imageToPdfOptions.style.display = 'block';
}

function renderImageList() {
  imageList.innerHTML = '';

  if (selectedImages.length === 0) {
    imageList.style.display = 'none';
    imageToPdfOptions.style.display = 'none';
    return;
  }

  selectedImages.forEach((imagePath, index) => {
    const item = document.createElement('div');
    item.className = 'image-item';
    item.draggable = true;
    item.dataset.index = index;

    const fileName = imagePath.split('\\\\').pop();

    item.innerHTML = `
      <span class="drag-handle">⋮⋮</span>
      <span class="image-name">🖼️ ${fileName}</span>
      <div class="image-actions">
        <button class="image-up-btn" title="Move Up">↑</button>
        <button class="image-down-btn" title="Move Down">↓</button>
        <button class="image-remove-btn" title="Remove">×</button>
      </div>
    `;

    // Remove button
    item.querySelector('.image-remove-btn').onclick = () => {
      selectedImages.splice(index, 1);
      renderImageList();
    };

    // Move up button
    item.querySelector('.image-up-btn').onclick = () => {
      if (index > 0) {
        [selectedImages[index], selectedImages[index - 1]] = [selectedImages[index - 1], selectedImages[index]];
        renderImageList();
      }
    };

    // Move down button
    item.querySelector('.image-down-btn').onclick = () => {
      if (index < selectedImages.length - 1) {
        [selectedImages[index], selectedImages[index + 1]] = [selectedImages[index + 1], selectedImages[index]];
        renderImageList();
      }
    };

    // Drag and drop reordering
    item.addEventListener('dragstart', (e) => {
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', index);
      item.classList.add('dragging');
    });

    item.addEventListener('dragend', () => {
      item.classList.remove('dragging');
    });

    item.addEventListener('dragover', (e) => {
      e.preventDefault();
      const draggingItem = imageList.querySelector('.dragging');
      if (draggingItem !== item) {
        const rect = item.getBoundingClientRect();
        const midpoint = rect.top + rect.height / 2;
        if (e.clientY < midpoint) {
          imageList.insertBefore(draggingItem, item);
        } else {
          imageList.insertBefore(draggingItem, item.nextSibling);
        }
      }
    });

    item.addEventListener('drop', (e) => {
      e.preventDefault();
      const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
      const toIndex = parseInt(item.dataset.index);

      if (fromIndex !== toIndex) {
        const [movedItem] = selectedImages.splice(fromIndex, 1);
        selectedImages.splice(toIndex, 0, movedItem);
        renderImageList();
      }
    });

    imageList.appendChild(item);
  });
}

createPdfBtn.onclick = async () => {
  if (selectedImages.length === 0) {
    alert('Please add at least one image');
    return;
  }

  try {
    createPdfBtn.disabled = true;
    createPdfBtn.textContent = '⏳ Creating PDF...';

    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      createPdfBtn.disabled = false;
      createPdfBtn.innerHTML = '<span>📄</span> Create PDF';
      return;
    }

    const options = {
      pageSize: pdfPageSize.value,
      orientation: pdfOrientation.value,
      margin: parseInt(pdfMargin.value),
      quality: pdfImageQuality.value
    };

    await window.pdfAPI.imagesToPdf(selectedImages, options, savePath);

    alert(`✅ PDF created successfully with ${selectedImages.length} image(s)!`);

    // Reset
    selectedImages = [];
    imageToPdfFileInput.value = '';
    renderImageList();
    createPdfBtn.disabled = false;
    createPdfBtn.innerHTML = '<span>📄</span> Create PDF';
  } catch (err) {
    console.error(err);
    alert('❌ Failed to create PDF: ' + err.message);
    createPdfBtn.disabled = false;
    createPdfBtn.innerHTML = '<span>📄</span> Create PDF';
  }
};

// Protect PDF Tool Logic
const protectDropzone = document.getElementById('protectDropzone');
const selectProtectPdfBtn = document.getElementById('selectProtectPdfBtn');
const protectFileInput = document.getElementById('protectFileInput');
const protectPdfInfo = document.getElementById('protectPdfInfo');
const protectOptions = document.getElementById('protectOptions');
const protectUserPassword = document.getElementById('protectUserPassword');
const protectConfirmPassword = document.getElementById('protectConfirmPassword');
const protectOwnerPassword = document.getElementById('protectOwnerPassword');
const protectPdfBtn = document.getElementById('protectPdfBtn');

let selectedProtectPdf = null;

// Drag and drop handlers
protectDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  protectDropzone.classList.add('drag-over');
});

protectDropzone.addEventListener('dragleave', () => {
  protectDropzone.classList.remove('drag-over');
});

protectDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  protectDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadProtectPdf(file.path);
  } else {
    showToast('Please drop a PDF file', 'error');
  }
});

selectProtectPdfBtn.onclick = () => {
  protectFileInput.click();
};

protectFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadProtectPdf(file.path);
  }
};

async function loadProtectPdf(filePath) {
  selectedProtectPdf = filePath;

  try {
    const pageCount = await window.pdfAPI.getPDFPageCount(filePath);

    protectPdfInfo.innerHTML = `
      <div class="info-icon">🔒</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${pageCount} pages • Ready to protect</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('protectFileInput').click()">Change</button>
    `;
    protectPdfInfo.classList.add('active');
    protectOptions.style.display = 'block';
  } catch (err) {
    console.error(err);
    showToast('Failed to load PDF: ' + err.message, 'error');
  }
}

protectPdfBtn.onclick = async () => {
  // Validation checks
  if (!selectedProtectPdf) {
    showToast('Please select a PDF file first', 'error');
    return;
  }

  const userPwd = protectUserPassword.value.trim();
  const confirmPwd = protectConfirmPassword.value.trim();

  if (!userPwd) {
    showToast('Please enter a password', 'error');
    protectUserPassword.focus();
    return;
  }

  if (userPwd !== confirmPwd) {
    showToast('Passwords do not match', 'error');
    protectConfirmPassword.focus();
    return;
  }

  // Only disable button AFTER all validations pass
  protectPdfBtn.disabled = true;
  protectPdfBtn.textContent = '⏳ Protecting PDF...';

  try {
    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      protectPdfBtn.disabled = false;
      protectPdfBtn.innerHTML = '<span>🔒</span> Protect PDF';
      return;
    }

    const options = {
      userPassword: userPwd,
      ownerPassword: protectOwnerPassword.value.trim() || userPwd
    };

    await window.pdfAPI.protectPdf(selectedProtectPdf, options, savePath);

    showToast('PDF protected successfully!', 'success');

    // Reset
    selectedProtectPdf = null;
    protectFileInput.value = '';
    protectPdfInfo.classList.remove('active');
    protectOptions.style.display = 'none';
    protectUserPassword.value = '';
    protectConfirmPassword.value = '';
    protectOwnerPassword.value = '';
    protectPdfBtn.disabled = false;
    protectPdfBtn.innerHTML = '<span>🔒</span> Protect PDF';
  } catch (err) {
    console.error(err);
    showToast('Failed to protect PDF: ' + err.message, 'error');
    protectPdfBtn.disabled = false;
    protectPdfBtn.innerHTML = '<span>🔒</span> Protect PDF';
  }
};

// Unlock PDF Tool Logic
const unlockDropzone = document.getElementById('unlockDropzone');
const selectUnlockPdfBtn = document.getElementById('selectUnlockPdfBtn');
const unlockFileInput = document.getElementById('unlockFileInput');
const unlockPdfInfo = document.getElementById('unlockPdfInfo');
const unlockOptions = document.getElementById('unlockOptions');
const unlockPassword = document.getElementById('unlockPassword');
const unlockPdfBtn = document.getElementById('unlockPdfBtn');

let selectedUnlockPdf = null;

// Drag and drop handlers
unlockDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  unlockDropzone.classList.add('drag-over');
});

unlockDropzone.addEventListener('dragleave', () => {
  unlockDropzone.classList.remove('drag-over');
});

unlockDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  unlockDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadUnlockPdf(file.path);
  } else {
    showToast('Please drop a PDF file', 'error');
  }
});

selectUnlockPdfBtn.onclick = () => {
  unlockFileInput.click();
};

unlockFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadUnlockPdf(file.path);
  }
};

async function loadUnlockPdf(filePath) {
  selectedUnlockPdf = filePath;

  unlockPdfInfo.innerHTML = `
    <div class="info-icon">🔓</div>
    <div class="info-content">
      <strong>${filePath}</strong>
      <p>Protected PDF • Ready to unlock</p>
    </div>
    <button class="change-btn" onclick="document.getElementById('unlockFileInput').click()">Change</button>
  `;
  unlockPdfInfo.classList.add('active');
  unlockOptions.style.display = 'block';
}

unlockPdfBtn.onclick = async () => {
  // Validation checks
  if (!selectedUnlockPdf) {
    showToast('Please select a PDF file first', 'error');
    return;
  }

  const password = unlockPassword.value.trim();

  if (!password) {
    showToast('Please enter the PDF password', 'error');
    unlockPassword.focus();
    return;
  }

  // Only disable button AFTER all validations pass
  unlockPdfBtn.disabled = true;
  unlockPdfBtn.textContent = '⏳ Unlocking PDF...';

  try {
    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      unlockPdfBtn.disabled = false;
      unlockPdfBtn.innerHTML = '<span>🔓</span> Unlock PDF';
      return;
    }

    await window.pdfAPI.unlockPdf(selectedUnlockPdf, password, savePath);

    showToast('PDF unlocked successfully!', 'success');

    // Reset
    selectedUnlockPdf = null;
    unlockFileInput.value = '';
    unlockPdfInfo.classList.remove('active');
    unlockOptions.style.display = 'none';
    unlockPassword.value = '';
    unlockPdfBtn.disabled = false;
    unlockPdfBtn.innerHTML = '<span>🔓</span> Unlock PDF';
  } catch (err) {
    console.error(err);
    showToast('Failed to unlock PDF: ' + err.message, 'error');
    unlockPdfBtn.disabled = false;
    unlockPdfBtn.innerHTML = '<span>🔓</span> Unlock PDF';
  }
};

// Compress PDF Tool Logic
const compressDropzone = document.getElementById('compressDropzone');
const selectCompressPdfBtn = document.getElementById('selectCompressPdfBtn');
const compressFileInput = document.getElementById('compressFileInput');
const compressPdfInfo = document.getElementById('compressPdfInfo');
const compressOptions = document.getElementById('compressOptions');
const compressActionBar = document.getElementById('compressActionBar');
const compressionLevel = document.getElementById('compressionLevel');
const compressImageQuality = document.getElementById('compressImageQuality');
const compressImageQualityValue = document.getElementById('compressImageQualityValue');
const removeMetadata = document.getElementById('removeMetadata');
const compressImages = document.getElementById('compressImages');
const compressPdfBtn = document.getElementById('compressPdfBtn');

let selectedCompressPdf = null;

// Update quality display
compressImageQuality.oninput = () => {
  compressImageQualityValue.textContent = `${compressImageQuality.value}%`;
};

// Drag and drop handlers
compressDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  compressDropzone.classList.add('drag-over');
});

compressDropzone.addEventListener('dragleave', () => {
  compressDropzone.classList.remove('drag-over');
});

compressDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  compressDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadCompressPdf(file.path);
  } else {
    showToast('Please drop a PDF file', 'error');
  }
});

selectCompressPdfBtn.onclick = () => {
  compressFileInput.click();
};

compressFileInput.onchange = async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadCompressPdf(file.path);
  }
};

async function loadCompressPdf(filePath) {
  selectedCompressPdf = filePath;

  try {
    const pageCount = await window.pdfAPI.getPDFPageCount(filePath);
    const fileSize = await window.pdfAPI.getFileSize(filePath);
    const fileSizeMB = (fileSize / (1024 * 1024)).toFixed(2);

    compressPdfInfo.innerHTML = `
      <div class="info-icon">🗜️</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${pageCount} pages • ${fileSizeMB} MB • Ready to compress</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('compressFileInput').click()">Change</button>
    `;
    compressPdfInfo.style.display = 'flex';
    compressOptions.style.display = 'block';
    compressActionBar.style.display = 'block';
  } catch (err) {
    console.error(err);
    showToast('Failed to load PDF: ' + err.message, 'error');
  }
}

compressPdfBtn.onclick = async () => {
  if (!selectedCompressPdf) {
    showToast('Please select a PDF file first', 'error');
    return;
  }

  compressPdfBtn.disabled = true;
  compressPdfBtn.textContent = '⏳ Compressing PDF...';

  try {
    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      compressPdfBtn.disabled = false;
      compressPdfBtn.innerHTML = '<span>🗜️</span> Compress PDF';
      return;
    }

    const options = {
      compressionLevel: compressionLevel.value,
      imageQuality: parseInt(compressImageQuality.value),
      removeMetadata: removeMetadata.checked,
      compressImages: compressImages.checked
    };

    const result = await window.pdfAPI.compressPdf(selectedCompressPdf, options, savePath);

    const savedMB = (result.saved / (1024 * 1024)).toFixed(2);
    showToast(`PDF compressed successfully! Saved ${savedMB} MB (${result.compressionRatio}% reduction)`, 'success');

    // Reset
    selectedCompressPdf = null;
    compressFileInput.value = '';
    compressPdfInfo.style.display = 'none';
    compressOptions.style.display = 'none';
    compressActionBar.style.display = 'none';
    compressPdfBtn.disabled = false;
    compressPdfBtn.innerHTML = '<span>🗜️</span> Compress PDF';
  } catch (err) {
    console.error(err);
    showToast('Failed to compress PDF: ' + err.message, 'error');
    compressPdfBtn.disabled = false;
    compressPdfBtn.innerHTML = '<span>🗜️</span> Compress PDF';
  }
};

// ============================================================================
// Edit Metadata Tool
// ============================================================================
const metadataDropzone = document.getElementById('metadataDropzone');
const selectMetadataPdfBtn = document.getElementById('selectMetadataPdfBtn');
const metadataFileInput = document.getElementById('metadataFileInput');
const metadataPdfInfo = document.getElementById('metadataPdfInfo');
const metadataForm = document.getElementById('metadataForm');
const metadataActionBar = document.getElementById('metadataActionBar');
const saveMetadataBtn = document.getElementById('saveMetadataBtn');

const metaTitle = document.getElementById('metaTitle');
const metaAuthor = document.getElementById('metaAuthor');
const metaSubject = document.getElementById('metaSubject');
const metaKeywords = document.getElementById('metaKeywords');
const metaCreator = document.getElementById('metaCreator');
const metaProducer = document.getElementById('metaProducer');
const metaCreationDate = document.getElementById('metaCreationDate');
const metaModDate = document.getElementById('metaModDate');

let selectedMetadataPdf = null;

// Drag and drop handlers
metadataDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  metadataDropzone.classList.add('drag-over');
});

metadataDropzone.addEventListener('dragleave', () => {
  metadataDropzone.classList.remove('drag-over');
});

metadataDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  metadataDropzone.classList.remove('drag-over');

  const file = e.dataTransfer.files[0];
  if (file && file.type === 'application/pdf') {
    await loadMetadataPdf(file.path);
  } else {
    showToast('Please drop a PDF file', 'error');
  }
});

// Browse button
selectMetadataPdfBtn.addEventListener('click', () => {
  metadataFileInput.click();
});

metadataFileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadMetadataPdf(file.path);
  }
});

// Load PDF and read metadata
async function loadMetadataPdf(filePath) {
  try {
    selectedMetadataPdf = filePath;

    // Get file info
    const pageCount = await window.pdfAPI.getPDFPageCount(filePath);
    const fileSize = await window.pdfAPI.getFileSize(filePath);

    // Display PDF info
    metadataPdfInfo.innerHTML = `
      <div class="info-icon">📝</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${pageCount} pages • ${fileSize} • Ready to edit metadata</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('metadataFileInput').click()">Change</button>
    `;
    metadataPdfInfo.style.display = 'flex';

    // Read metadata
    const metadata = await window.pdfAPI.readMetadata(filePath);

    // Populate form fields
    metaTitle.value = metadata.title || '';
    metaAuthor.value = metadata.author || '';
    metaSubject.value = metadata.subject || '';
    metaKeywords.value = metadata.keywords || '';
    metaCreator.value = metadata.creator || '';
    metaProducer.value = metadata.producer || '';

    // Display dates
    if (metadata.creationDate) {
      metaCreationDate.textContent = `Created: ${new Date(metadata.creationDate).toLocaleString()}`;
    } else {
      metaCreationDate.textContent = '';
    }

    if (metadata.modificationDate) {
      metaModDate.textContent = `Modified: ${new Date(metadata.modificationDate).toLocaleString()}`;
    } else {
      metaModDate.textContent = '';
    }

    // Show form and save button
    metadataForm.style.display = 'block';
    metadataActionBar.style.display = 'flex';
  } catch (error) {
    showToast(`Error loading PDF: ${error.message}`, 'error');
  }
}

// Save metadata
saveMetadataBtn.addEventListener('click', async () => {
  if (!selectedMetadataPdf) return;

  try {
    saveMetadataBtn.disabled = true;
    saveMetadataBtn.innerHTML = '<span>⏳</span> Saving...';

    // Get output path
    const savePath = await window.pdfAPI.showSaveDialog();
    if (!savePath) {
      saveMetadataBtn.disabled = false;
      saveMetadataBtn.innerHTML = '<span>💾</span> Save Metadata';
      return;
    }

    // Prepare metadata
    const metadata = {
      title: metaTitle.value,
      author: metaAuthor.value,
      subject: metaSubject.value,
      keywords: metaKeywords.value,
      creator: metaCreator.value,
      producer: metaProducer.value
    };

    // Update metadata
    await window.pdfAPI.updateMetadata(selectedMetadataPdf, metadata, savePath);

    showToast('Metadata saved successfully!', 'success');

    // Reset
    selectedMetadataPdf = null;
    metadataPdfInfo.innerHTML = '';
    metadataPdfInfo.style.display = 'none';
    metadataForm.style.display = 'none';
    metadataActionBar.style.display = 'none';
    metadataFileInput.value = '';

    // Clear form
    metaTitle.value = '';
    metaAuthor.value = '';
    metaSubject.value = '';
    metaKeywords.value = '';
    metaCreator.value = '';
    metaProducer.value = '';
    metaCreationDate.textContent = '';
    metaModDate.textContent = '';
  } catch (error) {
    showToast(`Error saving metadata: ${error.message}`, 'error');
  } finally {
    saveMetadataBtn.disabled = false;
    saveMetadataBtn.innerHTML = '<span>💾</span> Save Metadata';
  }
});

// ============================================================================
// Convert Tool
// ============================================================================
const convertDropzone = document.getElementById('convertDropzone');
const selectConvertFileBtn = document.getElementById('selectConvertFileBtn');
const convertFileInput = document.getElementById('convertFileInput');
const convertFileInfo = document.getElementById('convertFileInfo');
const conversionType = document.getElementById('conversionType');
const convertActionBar = document.getElementById('convertActionBar');
const convertBtn = document.getElementById('convertBtn');

let selectedConvertFile = null;

// File extension mapping
const FILE_EXTENSIONS = {
  'pdf-to-word': { input: '.pdf', output: '.docx' },
  'pdf-to-excel': { input: '.pdf', output: '.xlsx' },
  'pdf-to-ppt': { input: '.pdf', output: '.pptx' },
  'pdf-to-html': { input: '.pdf', output: '.html' },
  'word-to-pdf': { input: '.docx', output: '.pdf' },
  'excel-to-pdf': { input: '.xlsx', output: '.pdf' },
  'ppt-to-pdf': { input: '.pptx', output: '.pdf' },
  'html-to-pdf': { input: '.html', output: '.pdf' }
};

// Update file input accept based on conversion type
conversionType.addEventListener('change', () => {
  const type = conversionType.value;
  if (type && FILE_EXTENSIONS[type]) {
    convertFileInput.setAttribute('accept', FILE_EXTENSIONS[type].input);
  } else {
    convertFileInput.removeAttribute('accept');
  }

  // Clear selected file when conversion type changes
  if (selectedConvertFile) {
    selectedConvertFile = null;
    convertFileInfo.innerHTML = '';
    convertFileInfo.style.display = 'none';
    convertActionBar.style.display = 'none';
  }
});

// Drag and drop handlers
convertDropzone.addEventListener('dragover', (e) => {
  e.preventDefault();
  if (conversionType.value) {
    convertDropzone.classList.add('drag-over');
  }
});

convertDropzone.addEventListener('dragleave', () => {
  convertDropzone.classList.remove('drag-over');
});

convertDropzone.addEventListener('drop', async (e) => {
  e.preventDefault();
  convertDropzone.classList.remove('drag-over');

  if (!conversionType.value) {
    showToast('Please select a conversion type first', 'error');
    return;
  }

  const file = e.dataTransfer.files[0];
  if (file) {
    await loadConvertFile(file.path);
  }
});

// Browse button
selectConvertFileBtn.addEventListener('click', () => {
  if (!conversionType.value) {
    showToast('Please select a conversion type first', 'error');
    return;
  }
  convertFileInput.click();
});

convertFileInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (file) {
    await loadConvertFile(file.path);
  }
});

// Load file for conversion
async function loadConvertFile(filePath) {
  try {
    const type = conversionType.value;
    const expectedExt = FILE_EXTENSIONS[type].input;
    const fileExt = filePath.substring(filePath.lastIndexOf('.')).toLowerCase();

    // Validate file extension
    if (fileExt !== expectedExt) {
      showToast(`Please select a ${expectedExt} file for this conversion`, 'error');
      return;
    }

    selectedConvertFile = filePath;
    const fileSize = await window.pdfAPI.getFileSize(filePath);
    const outputExt = FILE_EXTENSIONS[type].output;

    // Display file info
    convertFileInfo.innerHTML = `
      <div class="info-icon">🔄</div>
      <div class="info-content">
        <strong>${filePath}</strong>
        <p>${fileSize} • Ready to convert to ${outputExt}</p>
      </div>
      <button class="change-btn" onclick="document.getElementById('convertFileInput').click()">Change</button>
    `;
    convertFileInfo.style.display = 'flex';
    convertActionBar.style.display = 'flex';
  } catch (error) {
    showToast(`Error loading file: ${error.message}`, 'error');
  }
}

// Convert file
convertBtn.addEventListener('click', async () => {
  if (!selectedConvertFile || !conversionType.value) return;

  try {
    convertBtn.disabled = true;
    convertBtn.innerHTML = '<span>⏳</span> Converting...';

    // Get output extension and prepare save dialog
    const outputExt = FILE_EXTENSIONS[conversionType.value].output;
    const extName = outputExt.substring(1).toUpperCase();
    const inputFileName = selectedConvertFile.split(/[\\\/]/).pop().replace(/\.[^.]+$/, '');

    const saveOptions = {
      title: `Save Converted ${extName} File`,
      defaultPath: `${inputFileName}${outputExt}`,
      filters: [{ name: `${extName} Files`, extensions: [outputExt.substring(1)] }]
    };

    // Get save path
    const savePath = await window.pdfAPI.showSaveDialog(saveOptions);
    if (!savePath) {
      convertBtn.disabled = false;
      convertBtn.innerHTML = '<span>🔄</span> Convert File';
      return;
    }

    // Perform conversion
    const result = await window.pdfAPI.convert(conversionType.value, selectedConvertFile, savePath);

    showToast('File converted successfully!', 'success');

    // Reset
    selectedConvertFile = null;
    convertFileInfo.innerHTML = '';
    convertFileInfo.style.display = 'none';
    convertActionBar.style.display = 'none';
    convertFileInput.value = '';
  } catch (error) {
    showToast(`Conversion failed: ${error.message}`, 'error');
  } finally {
    convertBtn.disabled = false;
    convertBtn.innerHTML = '<span>🔄</span> Convert File';
  }
});
