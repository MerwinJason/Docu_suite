// --- Navigation ---
const navLinks = document.querySelectorAll('.nav-links li');
const pages = document.querySelectorAll('.page');

navLinks.forEach(link => {
    link.addEventListener('click', () => {
        const targetId = link.getAttribute('data-target');
        navigateTo(targetId);
    });
});

function navigateTo(targetId) {
    if (!targetId) return;

    // Update Sidebar
    navLinks.forEach(l => l.classList.remove('active'));
    const activeLink = document.querySelector(`.nav-links li[data-target="${targetId}"]`);
    if (activeLink) {
        activeLink.classList.add('active');
        activeLink.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    }

    // Update Pages
    pages.forEach(p => p.classList.remove('active'));
    const targetPage = document.getElementById(targetId);
    if (targetPage) targetPage.classList.add('active');
}

// --- Notification System ---
function showToast(message, type = 'success') {
    const container = document.getElementById('notification-area');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `
        <i class="fa-solid ${type === 'success' ? 'fa-check-circle' : 'fa-exclamation-circle'}"></i>
        <span>${message}</span>
    `;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
}

// --- Utility: Parse Page Set ---
function parsePageSet(expr, totalPages) {
    expr = (expr || '').trim().toLowerCase();
    if (!expr) return new Set();

    const pages = new Set();

    if (expr.startsWith('first') || expr.startsWith('last')) {
        const parts = expr.replace(/[:=]/g, ' ').split(/\s+/);
        if (parts.length >= 2 && !isNaN(parts[1])) {
            const n = parseInt(parts[1]);
            if (parts[0] === 'first') {
                for (let i = 0; i < Math.min(n, totalPages); i++) pages.add(i);
            } else {
                for (let i = Math.max(0, totalPages - n); i < totalPages; i++) pages.add(i);
            }
            return pages;
        }
    }

    if (expr.startsWith('-') && !isNaN(expr.substring(1))) {
        const n = parseInt(expr.substring(1));
        for (let i = Math.max(0, totalPages - n); i < totalPages; i++) pages.add(i);
        return pages;
    }

    const tokens = expr.split(',').map(t => t.trim()).filter(t => t);

    tokens.forEach(tok => {
        try {
            if (!isNaN(tok)) {
                const i = parseInt(tok) - 1;
                if (i >= 0 && i < totalPages) pages.add(i);
            } else if (tok.includes('-')) {
                const parts = tok.split('-', 2);
                let startStr = parts[0].trim();
                let endStr = parts[1].trim();

                let start = startStr === '' ? 0 : parseInt(startStr) - 1;
                let end = endStr === '' ? totalPages - 1 : parseInt(endStr) - 1;

                if (startStr === 'last') start = totalPages - 1;
                if (endStr === 'last') end = totalPages - 1;

                if (isNaN(start)) start = 0;
                if (isNaN(end)) end = totalPages - 1;

                if (start < 0) start = 0;
                if (end >= totalPages) end = totalPages - 1;

                if (start <= end) {
                    for (let i = start; i <= end; i++) pages.add(i);
                }
            }
        } catch (e) {
            console.warn(`Could not parse token: ${tok}`);
        }
    });

    return pages;
}

// --- PDF Merger Feature ---
const pdfFiles = [];
const pdfDropZone = document.getElementById('pdf-drop-zone');
const pdfInput = document.getElementById('pdf-input');
const pdfList = document.getElementById('pdf-file-list');
const mergePdfBtn = document.getElementById('merge-pdf-btn');

if (pdfDropZone) setupDragDrop(pdfDropZone, pdfInput, handlePdfFiles);

function handlePdfFiles(files) {
    for (const file of files) {
        if (file.type === 'application/pdf') {
            pdfFiles.push({ file: file, id: Math.random().toString(36).substr(2, 9) });
        } else {
            showToast(`Skipped ${file.name} (not a PDF)`, 'error');
        }
    }
    renderPdfList();
    updateMergeButton(pdfFiles, mergePdfBtn);
}

function renderPdfList() {
    pdfList.innerHTML = '';
    pdfFiles.forEach((item, index) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        div.innerHTML = `
            <div class="file-info">
                <i class="fa-solid fa-file-pdf pdf-color"></i>
                <span>${item.file.name}</span>
                <input type="text" id="range-${item.id}" placeholder="Omit: e.g. 1, 3-5">
            </div>
            <i class="fa-solid fa-trash remove-file" onclick="removePdf(${index})"></i>
        `;
        pdfList.appendChild(div);
    });
}

window.removePdf = (index) => {
    pdfFiles.splice(index, 1);
    renderPdfList();
    updateMergeButton(pdfFiles, mergePdfBtn);
};

if (mergePdfBtn) {
    mergePdfBtn.addEventListener('click', async () => {
        if (pdfFiles.length < 2) return;

        try {
            mergePdfBtn.disabled = true;
            mergePdfBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Merging...';

            const mergedPdf = await PDFLib.PDFDocument.create();
            const globalOmitExpr = document.getElementById('global-page-range').value;

            for (const item of pdfFiles) {
                const arrayBuffer = await item.file.arrayBuffer();
                const pdf = await PDFLib.PDFDocument.load(arrayBuffer);
                const totalPages = pdf.getPageCount();

                const perFileExpr = document.getElementById(`range-${item.id}`).value;
                const globalExcludes = parsePageSet(globalOmitExpr, totalPages);
                const localExcludes = parsePageSet(perFileExpr, totalPages);

                const allExcludes = new Set([...globalExcludes, ...localExcludes]);

                const indicesToKeep = [];
                for (let i = 0; i < totalPages; i++) {
                    if (!allExcludes.has(i)) indicesToKeep.push(i);
                }

                if (indicesToKeep.length > 0) {
                    const copiedPages = await mergedPdf.copyPages(pdf, indicesToKeep);
                    copiedPages.forEach(page => mergedPdf.addPage(page));
                }
            }

            const pdfBytes = await mergedPdf.save();
            downloadFile(pdfBytes, 'merged_document.pdf', 'application/pdf');

            showToast('PDFs merged successfully!');
            pdfFiles.length = 0;
            renderPdfList();
        } catch (error) {
            console.error(error);
            showToast('Error merging PDFs', 'error');
        } finally {
            updateMergeButton(pdfFiles, mergePdfBtn);
            mergePdfBtn.innerHTML = '<i class="fa-solid fa-bolt"></i> Merge Files';
        }
    });
}

// --- Excel Merger Feature ---
const excelFiles = [];
const excelDropZone = document.getElementById('excel-drop-zone');
const excelInput = document.getElementById('excel-input');
const excelList = document.getElementById('excel-file-list');
const mergeExcelBtn = document.getElementById('merge-excel-btn');

if (excelDropZone) setupDragDrop(excelDropZone, excelInput, handleExcelFiles);

function handleExcelFiles(files) {
    for (const file of files) {
        if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv')) {
            excelFiles.push(file);
        } else {
            showToast(`Skipped ${file.name} (invalid type)`, 'error');
        }
    }
    renderExcelList();
    updateMergeButton(excelFiles, mergeExcelBtn);
}

function renderExcelList() {
    excelList.innerHTML = '';
    excelFiles.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'file-item';
        item.innerHTML = `
            <div class="file-info">
                <i class="fa-solid fa-file-excel excel-color"></i>
                <span>${file.name}</span>
            </div>
            <i class="fa-solid fa-trash remove-file" onclick="removeExcel(${index})"></i>
        `;
        excelList.appendChild(item);
    });
}

window.removeExcel = (index) => {
    excelFiles.splice(index, 1);
    renderExcelList();
    updateMergeButton(excelFiles, mergeExcelBtn);
};

if (mergeExcelBtn) {
    mergeExcelBtn.addEventListener('click', async () => {
        if (excelFiles.length < 2) return;

        try {
            mergeExcelBtn.disabled = true;
            mergeExcelBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Merging...';

            const headerRowIndex = parseInt(document.getElementById('excel-header-row').value) || 0;
            const mergeMode = document.getElementById('excel-merge-mode').value;
            const dedupeMode = document.getElementById('excel-dedupe-mode').value;
            const includeFilename = document.getElementById('include-filename').checked;

            let allData = [];
            let allColumns = new Set();
            let commonColumns = null;

            for (const file of excelFiles) {
                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer);
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex });

                if (jsonData.length > 0) {
                    const keys = Object.keys(jsonData[0]);
                    keys.forEach(k => allColumns.add(k));

                    if (commonColumns === null) {
                        commonColumns = new Set(keys);
                    } else {
                        commonColumns = new Set(keys.filter(x => commonColumns.has(x)));
                    }

                    jsonData.forEach(row => {
                        if (includeFilename) row['source_filename'] = file.name;
                        allData.push(row);
                    });
                }
            }

            let finalColumns = (mergeMode === 'union') ? Array.from(allColumns) : Array.from(commonColumns);
            if (includeFilename) finalColumns.unshift('source_filename');

            if (mergeMode === 'intersection') {
                allData = allData.map(row => {
                    const newRow = {};
                    finalColumns.forEach(col => {
                        if (row[col] !== undefined) newRow[col] = row[col];
                    });
                    return newRow;
                });
            }

            if (dedupeMode === 'drop') {
                const seen = new Set();
                allData = allData.filter(row => {
                    const signature = JSON.stringify(row);
                    if (seen.has(signature)) return false;
                    seen.add(signature);
                    return true;
                });
            }

            const newWorksheet = XLSX.utils.json_to_sheet(allData, { header: finalColumns });
            const newWorkbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Merged Data");

            XLSX.writeFile(newWorkbook, 'merged_data.xlsx');
            showToast('Excels merged successfully!');

            excelFiles.length = 0;
            renderExcelList();
        } catch (error) {
            console.error(error);
            showToast('Error merging Excels', 'error');
        } finally {
            updateMergeButton(excelFiles, mergeExcelBtn);
            mergeExcelBtn.innerHTML = '<i class="fa-solid fa-bolt"></i> Merge Files';
        }
    });
}

// --- PDF Tools Feature ---
let currentToolsPdf = null;
let currentToolsPdfBytes = null;
const toolsDropZone = document.getElementById('tools-drop-zone');
const toolsInput = document.getElementById('tools-input');
const toolsFileName = document.getElementById('tools-file-name');
const toolsControls = document.getElementById('tools-controls');

if (toolsDropZone) {
    setupDragDrop(toolsDropZone, toolsInput, async (files) => {
        if (files.length > 0 && files[0].type === 'application/pdf') {
            const file = files[0];
            toolsFileName.textContent = file.name;
            currentToolsPdfBytes = await file.arrayBuffer();
            currentToolsPdf = await PDFLib.PDFDocument.load(currentToolsPdfBytes);
            toolsControls.classList.remove('disabled');
            showToast('PDF loaded. Enter page range.');
        } else {
            showToast('Please select a valid PDF file', 'error');
        }
    });
}

const keepPagesBtn = document.getElementById('keep-pages-btn');
if (keepPagesBtn) keepPagesBtn.addEventListener('click', () => processPages('keep'));

const removePagesBtn = document.getElementById('remove-pages-btn');
if (removePagesBtn) removePagesBtn.addEventListener('click', () => processPages('remove'));

const rotateCwBtn = document.getElementById('rotate-cw-btn');
if (rotateCwBtn) rotateCwBtn.addEventListener('click', () => rotatePages(90));

const rotateCcwBtn = document.getElementById('rotate-ccw-btn');
if (rotateCcwBtn) rotateCcwBtn.addEventListener('click', () => rotatePages(-90));

async function processPages(mode) {
    if (!currentToolsPdf) return;

    const rangeInput = document.getElementById('page-range').value;
    const totalPages = currentToolsPdf.getPageCount();
    const selectedPages = parsePageSet(rangeInput, totalPages);

    if (selectedPages.size === 0) {
        showToast('No pages selected', 'error');
        return;
    }

    try {
        const newPdf = await PDFLib.PDFDocument.create();
        const indicesToKeep = [];

        if (mode === 'keep') {
            selectedPages.forEach(i => indicesToKeep.push(i));
        } else {
            for (let i = 0; i < totalPages; i++) {
                if (!selectedPages.has(i)) indicesToKeep.push(i);
            }
        }

        if (indicesToKeep.length === 0) {
            showToast('Result would be empty PDF', 'error');
            return;
        }

        indicesToKeep.sort((a, b) => a - b);
        const copiedPages = await newPdf.copyPages(currentToolsPdf, indicesToKeep);
        copiedPages.forEach(page => newPdf.addPage(page));

        const pdfBytes = await newPdf.save();
        downloadFile(pdfBytes, mode === 'keep' ? 'extracted_pages.pdf' : 'remaining_pages.pdf', 'application/pdf');
        showToast('PDF processed successfully!');

    } catch (error) {
        console.error(error);
        showToast('Error processing PDF', 'error');
    }
}

async function rotatePages(degrees) {
    if (!currentToolsPdf) return;

    const rangeInput = document.getElementById('page-range').value;
    const totalPages = currentToolsPdf.getPageCount();
    const selectedPages = parsePageSet(rangeInput, totalPages);

    if (selectedPages.size === 0) {
        showToast('No pages selected for rotation', 'error');
        return;
    }

    try {
        const newPdf = await PDFLib.PDFDocument.load(currentToolsPdfBytes);
        const pages = newPdf.getPages();

        selectedPages.forEach(i => {
            const page = pages[i];
            const currentRotation = page.getRotation().angle;
            page.setRotation(PDFLib.degrees(currentRotation + degrees));
        });

        const pdfBytes = await newPdf.save();
        downloadFile(pdfBytes, 'rotated_document.pdf', 'application/pdf');
        showToast('PDF rotated successfully!');
    } catch (error) {
        console.error(error);
        showToast('Error rotating PDF', 'error');
    }
}

// --- PDF Deduplicator Feature ---
let dedupePdf = null;
let dedupePdfBytes = null;
const dedupeDropZone = document.getElementById('dedupe-drop-zone');
const dedupeInput = document.getElementById('dedupe-input');
const dedupeFileName = document.getElementById('dedupe-file-name');
const dedupeBtn = document.getElementById('dedupe-btn');

if (dedupeDropZone) {
    setupDragDrop(dedupeDropZone, dedupeInput, async (files) => {
        if (files.length > 0 && files[0].type === 'application/pdf') {
            const file = files[0];
            dedupeFileName.textContent = file.name;
            dedupePdfBytes = await file.arrayBuffer();
            dedupePdf = await PDFLib.PDFDocument.load(dedupePdfBytes);
            dedupeBtn.disabled = false;
            showToast('PDF loaded. Click to remove duplicates.');
        } else {
            showToast('Please select a valid PDF file', 'error');
        }
    });
}

if (dedupeBtn) {
    dedupeBtn.addEventListener('click', async () => {
        if (!dedupePdf) return;

        try {
            dedupeBtn.disabled = true;
            dedupeBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing...';

            // Simulation
            alert("Note: Client-side deduplication requires advanced text extraction. This is a simulation.");

            const pdfBytes = await dedupePdf.save();
            downloadFile(pdfBytes, 'deduplicated.pdf', 'application/pdf');
            showToast('Process complete (Simulation)');

        } catch (error) {
            console.error(error);
            showToast('Error processing PDF', 'error');
        } finally {
            dedupeBtn.disabled = false;
            dedupeBtn.innerHTML = '<i class="fa-solid fa-magic"></i> Remove Duplicates';
        }
    });
}


// --- Universal Converter Feature ---
const converterFiles = [];
const converterDropZone = document.getElementById('converter-drop-zone');
const converterInput = document.getElementById('converter-input');
const converterList = document.getElementById('converter-list');
const downloadAllBtn = document.getElementById('download-all-btn');

if (converterDropZone) {
    setupDragDrop(converterDropZone, converterInput, handleConverterFiles);
}

if (downloadAllBtn) {
    downloadAllBtn.addEventListener('click', downloadAll);
}

function handleConverterFiles(files) {
    for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext === 'pdf' || ext === 'docx' || ext === 'pptx') {
            converterFiles.push({
                file: file,
                id: Math.random().toString(36).substr(2, 9),
                status: 'pending', // pending, processing, success, error
                resultBlob: null,
                resultName: null,
                previewHtml: null
            });
        } else {
            showToast(`Skipped ${file.name} (unsupported format)`, 'error');
        }
    }
    renderConverterList();
}

function renderConverterList() {
    converterList.innerHTML = '';
    let hasSuccess = false;

    converterFiles.forEach((item) => {
        const ext = item.file.name.split('.').pop().toLowerCase();
        let iconClass = '';
        if (ext === 'pdf') iconClass = 'fa-file-pdf pdf-color';
        else if (ext === 'docx') iconClass = 'fa-file-word word-color';
        else if (ext === 'pptx') iconClass = 'fa-file-powerpoint tools-color';

        // Determine options based on file type
        let optionsHtml = '';
        if (ext === 'pdf') {
            optionsHtml = `
                <select id="format-${item.id}">
                    <option value="docx">to Word (DOCX)</option>
                    <option value="pptx">to PowerPoint (PPTX)</option>
                    <option value="txt">to Text (TXT)</option>
                </select>
            `;
        } else if (ext === 'docx') {
            optionsHtml = `
                <select id="format-${item.id}">
                    <option value="pdf">to PDF</option>
                    <option value="txt">to Text (TXT)</option>
                </select>
            `;
        } else if (ext === 'pptx') {
            optionsHtml = `
                <select id="format-${item.id}">
                    <option value="pdf">to PDF</option>
                    <option value="txt">to Text (TXT)</option>
                </select>
            `;
        }

        if (item.status === 'success') hasSuccess = true;

        const div = document.createElement('div');
        div.className = 'converter-item';
        div.innerHTML = `
            <div class="converter-header">
                <div class="converter-file-info">
                    <i class="fa-solid ${iconClass}"></i>
                    <span>${item.file.name}</span>
                    ${item.status === 'success' ? '<span class="status-badge success">Converted</span>' : ''}
                    ${item.status === 'processing' ? '<span class="status-badge processing">Processing...</span>' : ''}
                </div>
                <div class="converter-actions">
                    ${item.status === 'pending' ? optionsHtml : ''}
                    
                    ${item.status === 'pending' ? `
                        <button class="icon-btn convert" onclick="convertItem('${item.id}')" title="Convert">
                            <i class="fa-solid fa-arrow-right"></i>
                        </button>
                    ` : ''}

                    ${item.status === 'success' ? `
                        <button class="icon-btn download" onclick="downloadItem('${item.id}')" title="Download">
                            <i class="fa-solid fa-download"></i>
                        </button>
                    ` : ''}

                    <button class="icon-btn" onclick="togglePreview('${item.id}')" title="Preview">
                        <i class="fa-solid fa-eye"></i>
                    </button>
                    
                    <button class="icon-btn" onclick="removeItem('${item.id}')" title="Remove">
                        <i class="fa-solid fa-trash"></i>
                    </button>
                </div>
            </div>
            <div class="preview-area" id="preview-${item.id}">
                ${item.previewHtml || '<p>Loading preview...</p>'}
            </div>
        `;
        converterList.appendChild(div);

        // Load preview if not already loaded
        if (!item.previewHtml) {
            loadPreview(item);
        }
    });

    if (downloadAllBtn) {
        downloadAllBtn.style.display = hasSuccess ? 'block' : 'none';
    }
}

async function loadPreview(item) {
    const previewEl = document.getElementById(`preview-${item.id}`);
    try {
        const ext = item.file.name.split('.').pop().toLowerCase();
        if (ext === 'docx') {
            const arrayBuffer = await item.file.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
            item.previewHtml = result.value;
        } else if (ext === 'pdf') {
            const text = await extractTextFromPdf(item.file);
            item.previewHtml = `<p>${text.replace(/\n/g, '<br>').substring(0, 1000)}...</p>`;
        } else if (ext === 'pptx') {
            item.previewHtml = '<p>Preview not available for PPTX yet.</p>';
        }
        if (previewEl) previewEl.innerHTML = item.previewHtml;
    } catch (e) {
        console.error(e);
        item.previewHtml = '<p>Preview unavailable</p>';
        if (previewEl) previewEl.innerHTML = item.previewHtml;
    }
}

window.togglePreview = (id) => {
    const el = document.getElementById(`preview-${id}`);
    if (el) el.classList.toggle('active');
};

window.removeItem = (id) => {
    const idx = converterFiles.findIndex(x => x.id === id);
    if (idx > -1) {
        converterFiles.splice(idx, 1);
        renderConverterList();
    }
};

window.convertItem = async (id) => {
    const item = converterFiles.find(x => x.id === id);
    if (!item) return;

    const formatSelect = document.getElementById(`format-${id}`);
    const targetFormat = formatSelect.value;

    item.status = 'processing';
    renderConverterList();

    try {
        const ext = item.file.name.split('.').pop().toLowerCase();

        if (targetFormat === 'pdf' && ext === 'docx') {
            // DOCX to PDF
            const arrayBuffer = await item.file.arrayBuffer();
            const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });

            const container = document.createElement('div');
            container.innerHTML = result.value;
            container.style.background = 'white';
            container.style.color = 'black';
            container.style.padding = '20px';
            container.style.position = 'fixed';
            container.style.left = '-9999px';
            document.body.appendChild(container);

            const opt = {
                margin: 1,
                filename: item.file.name.replace('.docx', '.pdf'),
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 2 },
                jsPDF: { unit: 'in', format: 'letter', orientation: 'portrait' }
            };

            const pdfBlob = await html2pdf().set(opt).from(container).output('blob');
            document.body.removeChild(container);

            item.resultBlob = pdfBlob;
            item.resultName = item.file.name.replace('.docx', '.pdf');

        } else if (targetFormat === 'docx' && ext === 'pdf') {
            // PDF to DOCX (Fixed using docx library)
            const text = await extractTextFromPdf(item.file);
            const blob = await createDocxBlob(text);
            item.resultBlob = blob;
            item.resultName = item.file.name.replace('.pdf', '.docx');

        } else if (targetFormat === 'pptx' && ext === 'pdf') {
            // PDF to PPTX
            const text = await extractTextFromPdf(item.file);
            const blob = await createPptxBlob(text);
            item.resultBlob = blob;
            item.resultName = item.file.name.replace('.pdf', '.pptx');

        } else if (targetFormat === 'pdf' && ext === 'pptx') {
            // PPTX to PDF (Basic text extraction + PDF generation)
            const text = await extractTextFromPptx(item.file);
            const blob = createPdfFromText(text);
            item.resultBlob = blob;
            item.resultName = item.file.name.replace('.pptx', '.pdf');

        } else if (targetFormat === 'txt') {
            // Any to TXT
            let text = '';
            if (ext === 'pdf') {
                text = await extractTextFromPdf(item.file);
            } else if (ext === 'docx') {
                const arrayBuffer = await item.file.arrayBuffer();
                const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
                text = result.value;
            } else if (ext === 'pptx') {
                text = await extractTextFromPptx(item.file);
            }
            item.resultBlob = new Blob([text], { type: 'text/plain' });
            item.resultName = item.file.name.replace(/\.[^/.]+$/, "") + ".txt";
        }

        item.status = 'success';
        showToast('Conversion successful!');
    } catch (error) {
        console.error(error);
        item.status = 'error';
        showToast('Conversion failed', 'error');
    }
    renderConverterList();
};

window.downloadItem = (id) => {
    const item = converterFiles.find(x => x.id === id);
    if (item && item.resultBlob) {
        const url = URL.createObjectURL(item.resultBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = item.resultName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
};

function downloadAll() {
    const successItems = converterFiles.filter(item => item.status === 'success' && item.resultBlob);
    if (successItems.length === 0) return;

    if (successItems.length === 1) {
        downloadItem(successItems[0].id);
    } else {
        const zip = new JSZip();
        successItems.forEach(item => {
            zip.file(item.resultName, item.resultBlob);
        });
        zip.generateAsync({ type: "blob" }).then(function (content) {
            const url = URL.createObjectURL(content);
            const a = document.createElement('a');
            a.href = url;
            a.download = "converted_files.zip";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        });
    }
}

async function extractTextFromPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
    let fullText = "";

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str).join(' ');
        fullText += pageText + "\n\n";
    }
    return fullText;
}

async function createDocxBlob(text) {
    const { Document, Packer, Paragraph, TextRun } = docx;

    // Split text by newlines to create paragraphs
    const paragraphs = text.split('\n').map(line => {
        return new Paragraph({
            children: [new TextRun(line)],
        });
    });

    const doc = new Document({
        sections: [{
            properties: {},
            children: paragraphs,
        }],
    });

    return await Packer.toBlob(doc);
}

async function createPptxBlob(text) {
    const pptx = new PptxGenJS();

    // Simple logic: Create a new slide for every ~500 characters or split by paragraphs
    const lines = text.split('\n').filter(line => line.trim() !== '');
    const linesPerSlide = 10; // Adjust as needed

    for (let i = 0; i < lines.length; i += linesPerSlide) {
        const slide = pptx.addSlide();
        const slideText = lines.slice(i, i + linesPerSlide).join('\n');
        slide.addText(slideText, { x: 0.5, y: 0.5, w: '90%', h: '90%', fontSize: 14, color: '363636' });
    }

    return await pptx.write("blob");
}

async function extractTextFromPptx(file) {
    const zip = new JSZip();
    const arrayBuffer = await file.arrayBuffer();
    const loadedZip = await zip.loadAsync(arrayBuffer);
    let fullText = "";

    // Iterate over slide files
    const slideFiles = Object.keys(loadedZip.files).filter(filename => filename.startsWith("ppt/slides/slide") && filename.endsWith(".xml"));

    // Sort slides numerically
    slideFiles.sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml/)[1]);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)[1]);
        return numA - numB;
    });

    for (const filename of slideFiles) {
        const content = await loadedZip.file(filename).async("text");
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(content, "text/xml");
        const textElements = xmlDoc.getElementsByTagName("a:t");

        for (let i = 0; i < textElements.length; i++) {
            fullText += textElements[i].textContent + " ";
        }
        fullText += "\n\n";
    }

    return fullText;
}

function createPdfFromText(text) {
    const doc = new jsPDF();
    const splitText = doc.splitTextToSize(text, 180); // Width limit
    let y = 10;

    for (const line of splitText) {
        if (y > 280) {
            doc.addPage();
            y = 10;
        }
        doc.text(line, 10, y);
        y += 7;
    }
    return doc.output('blob');
}

// --- Helpers ---
function setupDragDrop(zone, input, handler) {
    zone.addEventListener('click', () => input.click());
    input.addEventListener('change', (e) => handler(e.target.files));

    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));

    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        handler(e.dataTransfer.files);
    });
}

function updateMergeButton(list, btn) {
    if (btn) btn.disabled = list.length < 2;
}

function downloadFile(data, filename, type) {
    const blob = new Blob([data], { type: type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
