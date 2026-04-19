// ===== Global Variables =====
let currentFile = null;
let markdownContent = '';

// ===== DOM Elements =====
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');
const previewSection = document.getElementById('previewSection');
const fileName = document.getElementById('fileName');
const removeBtn = document.getElementById('removeBtn');
const previewPane = document.getElementById('previewPane');
const sourcePane = document.getElementById('sourcePane');
const exportPdf = document.getElementById('exportPdf');
const exportWord = document.getElementById('exportWord');
const loadingOverlay = document.getElementById('loadingOverlay');
const tabs = document.querySelectorAll('.tab');

// ===== Initialize =====
function init() {
    setupEventListeners();
    setupMarkedOptions();
}

function setupMarkedOptions() {
    marked.setOptions({
        breaks: true,
        gfm: true,
        headerIds: false
    });
}

// ===== Event Listeners =====
function setupEventListeners() {
    // File selection - label triggers input automatically
    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop on upload area
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);

    // Remove file
    removeBtn.addEventListener('click', resetUpload);

    // Export buttons
    exportPdf.addEventListener('click', exportToPdf);
    exportWord.addEventListener('click', exportToWord);

    // Tabs
    tabs.forEach(tab => {
        tab.addEventListener('click', handleTabClick);
    });
}

// ===== Drag and Drop Handlers =====
function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

// ===== File Handling =====
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

function processFile(file) {
    // Validate file type
    const validTypes = ['.md', '.markdown', '.txt'];
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

    if (!validTypes.some(type => fileExtension === type)) {
        alert('请选择有效的 Markdown 文件 (.md, .markdown, .txt)');
        return;
    }

    currentFile = file;

    // Read file content
    const reader = new FileReader();
    reader.onload = (e) => {
        markdownContent = e.target.result;
        displayPreview();
    };
    reader.readAsText(file);
}

function displayPreview() {
    // Show preview section
    uploadSection.style.display = 'none';
    previewSection.style.display = 'block';

    // Set file name
    fileName.textContent = currentFile.name;

    // Render markdown
    const htmlContent = marked.parse(markdownContent);
    previewPane.innerHTML = htmlContent;
    sourcePane.textContent = markdownContent;

    // Reset to preview tab
    switchTab('preview');
}

function resetUpload() {
    uploadSection.style.display = 'block';
    previewSection.style.display = 'none';
    currentFile = null;
    markdownContent = '';
    fileInput.value = '';
    previewPane.innerHTML = '';
    sourcePane.textContent = '';
}

// ===== Tab Handling =====
function handleTabClick(e) {
    const tabType = e.target.dataset.tab;
    switchTab(tabType);
}

function switchTab(tabType) {
    tabs.forEach(tab => {
        tab.classList.toggle('active', tab.dataset.tab === tabType);
    });

    if (tabType === 'preview') {
        previewPane.style.display = 'block';
        sourcePane.style.display = 'none';
    } else {
        previewPane.style.display = 'none';
        sourcePane.style.display = 'block';
    }
}

// ===== Export to PDF =====
async function exportToPdf() {
    if (!markdownContent) {
        alert('请先上传 Markdown 文件');
        return;
    }

    showLoading();

    try {
        // Create a dedicated print container with all styles inline
        const printDiv = document.createElement('div');
        printDiv.id = 'pdf-print-container';

        // Parse markdown and create styled HTML
        const htmlContent = marked.parse(markdownContent);

        printDiv.innerHTML = `
            <style>
                #pdf-print-container {
                    font-family: 'Nunito', Arial, sans-serif;
                    padding: 30px 40px;
                    background: white;
                    color: #2D3748;
                    line-height: 1.7;
                }
                #pdf-print-container h1 { font-size: 24px; margin: 20px 0 12px; border-bottom: 2px solid #E5E7EB; padding-bottom: 6px; font-weight: 700; }
                #pdf-print-container h2 { font-size: 18px; margin: 16px 0 8px; font-weight: 700; }
                #pdf-print-container h3 { font-size: 15px; margin: 12px 0 6px; font-weight: 700; }
                #pdf-print-container p { margin: 10px 0; }
                #pdf-print-container ul, #pdf-print-container ol { margin: 10px 0; padding-left: 24px; }
                #pdf-print-container li { margin: 4px 0; }
                #pdf-print-container code { background: #F1F3F5; padding: 2px 5px; border-radius: 3px; font-family: Consolas, monospace; font-size: 0.9em; }
                #pdf-print-container pre { background: #2D3748; color: #E2E8F0; padding: 12px 16px; border-radius: 6px; overflow-x: auto; margin: 12px 0; font-family: Consolas, monospace; font-size: 13px; }
                #pdf-print-container pre code { background: transparent; padding: 0; color: inherit; }
                #pdf-print-container blockquote { border-left: 3px solid #A78BFA; margin: 12px 0; padding-left: 12px; color: #666; font-style: italic; }
                #pdf-print-container table { width: 100%; border-collapse: collapse; margin: 12px 0; }
                #pdf-print-container th, #pdf-print-container td { border: 1px solid #E5E7EB; padding: 8px 12px; text-align: left; }
                #pdf-print-container th { background: #F8F9FA; font-weight: 700; }
                #pdf-print-container a { color: #A78BFA; }
                #pdf-print-container img { max-width: 100%; height: auto; }
                #pdf-print-container hr { border: none; border-top: 1px solid #E5E7EB; margin: 16px 0; }
            </style>
            ${htmlContent}
        `;

        // Append to body in a visible way (html2canvas needs visibility)
        document.body.appendChild(printDiv);

        // Wait for styles to apply
        await new Promise(r => setTimeout(r, 300));

        const opt = {
            margin: 12,
            filename: getExportFilename('.pdf'),
            image: { type: 'jpeg', quality: 0.92 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                logging: false,
                allowTaint: true,
                scrollY: 0,
                scrollX: 0,
                windowWidth: printDiv.scrollWidth,
                windowHeight: printDiv.scrollHeight
            },
            jsPDF: {
                unit: 'mm',
                format: 'a4',
                orientation: 'portrait'
            },
            pagebreak: { mode: 'avoid-all' }
        };

        await html2pdf().set(opt).from(printDiv).save();

        document.body.removeChild(printDiv);
        hideLoading();
    } catch (error) {
        hideLoading();
        const printDiv = document.getElementById('pdf-print-container');
        if (printDiv) document.body.removeChild(printDiv);
        alert('PDF 导出失败: ' + error.message);
        console.error(error);
    }
}

// ===== Export to Word (DOCX) =====
async function exportToWord() {
    if (!markdownContent) {
        alert('请先上传 Markdown 文件');
        return;
    }

    showLoading();

    try {
        const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle } = docx;

        // Parse markdown and convert to docx paragraphs
        const children = parseMarkdownToDocx(markdownContent);

        const doc = new Document({
            numbering: {
                config: [{
                    reference: 'default-numbering',
                    levels: [{
                        level: 0,
                        format: 'decimal',
                        text: '%1.',
                        alignment: 'start'
                    }]
                }]
            },
            sections: [{
                properties: {},
                children: children
            }]
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, getExportFilename('.docx'));

        hideLoading();
    } catch (error) {
        hideLoading();
        alert('Word 导出失败: ' + error.message);
        console.error(error);
    }
}

// ===== Parse Markdown to DOCX =====
function parseMarkdownToDocx(content) {
    const { Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType } = docx;
    const elements = [];
    const lines = content.split('\n');

    let inCodeBlock = false;
    let codeContent = '';
    let inTable = false;
    let tableRows = [];
    let isTableHeader = false;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // Code block handling
        if (line.startsWith('```')) {
            if (inCodeBlock) {
                // End code block
                if (codeContent.trim()) {
                    elements.push(new Paragraph({
                        children: [new TextRun({
                            text: codeContent.trim(),
                            font: 'Consolas',
                            size: 20,
                            shading: { fill: 'F5F5F5' }
                        })],
                        shading: { fill: 'F5F5F5' }
                    }));
                }
                codeContent = '';
                inCodeBlock = false;
            } else {
                // Start code block
                inCodeBlock = true;
            }
            continue;
        }

        if (inCodeBlock) {
            codeContent += line + '\n';
            continue;
        }

        // Table handling
        if (line.match(/^\|/)) {
            if (!inTable) {
                inTable = true;
                isTableHeader = true;
                tableRows = [];
            }

            // Check if it's separator line
            if (line.match(/^\|[-:]+\|[-:]+\|/)) {
                isTableHeader = false;
                continue;
            }

            // Parse table row
            const cells = line.split('|')
                .filter(cell => cell.trim() !== '')
                .map(cell => cell.trim());

            if (cells.length > 0) {
                tableRows.push({ cells, isHeader: isTableHeader });
            }
            continue;
        } else if (inTable) {
            // End of table - create docx table
            if (tableRows.length > 0) {
                const docxTable = createDocxTable(tableRows);
                elements.push(docxTable);
                elements.push(new Paragraph({ children: [] })); // Add spacing
            }
            inTable = false;
            tableRows = [];
            isTableHeader = false;
        }

        // Empty line
        if (line.trim() === '') {
            elements.push(new Paragraph({ children: [] }));
            continue;
        }

        // Headers
        if (line.startsWith('#')) {
            const level = (line.match(/^#+/) || [''])[0].length;
            const text = line.replace(/^#+\s*/, '');
            let headingLevel;

            switch (level) {
                case 1: headingLevel = HeadingLevel.HEADING_1; break;
                case 2: headingLevel = HeadingLevel.HEADING_2; break;
                case 3: headingLevel = HeadingLevel.HEADING_3; break;
                default: headingLevel = HeadingLevel.HEADING_4;
            }

            elements.push(new Paragraph({
                heading: headingLevel,
                children: [new TextRun({
                    text: text,
                    bold: true
                })]
            }));
            continue;
        }

        // Horizontal rule
        if (line.match(/^-{3,}$|^\*{3,}$/)) {
            elements.push(new Paragraph({
                border: {
                    bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC' }
                },
                children: []
            }));
            continue;
        }

        // Blockquote
        if (line.startsWith('>')) {
            elements.push(new Paragraph({
                indent: { left: 720 },
                children: [new TextRun({
                    text: line.replace(/^>\s*/, ''),
                    italics: true,
                    color: '666666'
                })]
            }));
            continue;
        }

        // Unordered list
        if (line.match(/^[\*\-\+]\s/)) {
            elements.push(new Paragraph({
                bullet: { level: 0 },
                children: parseInlineFormatting(line.replace(/^[\*\-\+]\s/, ''))
            }));
            continue;
        }

        // Ordered list
        if (line.match(/^\d+\.\s/)) {
            elements.push(new Paragraph({
                numbering: { reference: 'default-numbering', level: 0 },
                children: parseInlineFormatting(line.replace(/^\d+\.\s/, ''))
            }));
            continue;
        }

        // Regular paragraph with inline formatting
        elements.push(new Paragraph({
            children: parseInlineFormatting(line)
        }));
    }

    // Handle remaining table at end of content
    if (inTable && tableRows.length > 0) {
        const docxTable = createDocxTable(tableRows);
        elements.push(docxTable);
    }

    return elements;
}

// ===== Create DOCX Table =====
function createDocxTable(tableRows) {
    const { Table, TableRow, TableCell, WidthType, BorderStyle, TextRun, Paragraph } = docx;

    const numCols = Math.max(...tableRows.map(r => r.cells.length));

    const rows = tableRows.map(row => {
        const cells = row.cells.map(cellText => {
            return new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({
                        text: cellText,
                        bold: row.isHeader
                    })]
                })],
                width: { size: 100 / numCols, type: WidthType.PERCENTAGE },
                shading: row.isHeader ? { fill: 'F0F0F0' } : undefined
            });
        });

        // Fill missing cells
        while (cells.length < numCols) {
            cells.push(new TableCell({
                children: [new Paragraph({ children: [] })],
                width: { size: 100 / numCols, type: WidthType.PERCENTAGE }
            }));
        }

        return new TableRow({ children: cells });
    });

    return new Table({
        rows: rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' }
        }
    });
}

// ===== Parse Inline Formatting =====
function parseInlineFormatting(text) {
    const { TextRun } = docx;
    const runs = [];
    let remaining = text;

    while (remaining.length > 0) {
        // Bold text **text** or __text__
        const boldMatch = remaining.match(/\*\*(.+?)\*\*|__(.+?)__/);
        if (boldMatch && boldMatch.index === 0) {
            runs.push(new TextRun({
                text: boldMatch[1] || boldMatch[2],
                bold: true
            }));
            remaining = remaining.substring(boldMatch[0].length);
            continue;
        }

        // Italic text *text* or _text_
        const italicMatch = remaining.match(/\*(.+?)\*|_(.+?)_/);
        if (italicMatch && italicMatch.index === 0 && !remaining.startsWith('**') && !remaining.startsWith('__')) {
            runs.push(new TextRun({
                text: italicMatch[1] || italicMatch[2],
                italics: true
            }));
            remaining = remaining.substring(italicMatch[0].length);
            continue;
        }

        // Inline code `code`
        const codeMatch = remaining.match(/`([^`]+)`/);
        if (codeMatch && codeMatch.index === 0) {
            runs.push(new TextRun({
                text: codeMatch[1],
                font: 'Consolas',
                size: 20,
                shading: { fill: 'F0F0F0' }
            }));
            remaining = remaining.substring(codeMatch[0].length);
            continue;
        }

        // Strikethrough ~~text~~
        const strikeMatch = remaining.match(/~~(.+?)~~/);
        if (strikeMatch && strikeMatch.index === 0) {
            runs.push(new TextRun({
                text: strikeMatch[1],
                strike: true
            }));
            remaining = remaining.substring(strikeMatch[0].length);
            continue;
        }

        // Regular text until next special character
        const nextSpecial = remaining.search(/\*\*|__|~~|`|\*(?!\*)|_(?!_)/);
        if (nextSpecial === -1) {
            runs.push(new TextRun({ text: remaining }));
            break;
        } else if (nextSpecial === 0) {
            // Handle cases where pattern didn't match at position 0
            runs.push(new TextRun({ text: remaining[0] }));
            remaining = remaining.substring(1);
        } else {
            runs.push(new TextRun({ text: remaining.substring(0, nextSpecial) }));
            remaining = remaining.substring(nextSpecial);
        }
    }

    return runs;
}

// ===== Utility Functions =====
function getExportFilename(extension) {
    if (currentFile) {
        const baseName = currentFile.name.replace(/\.[^/.]+$/, '');
        return baseName + extension;
    }
    return 'converted-document' + extension;
}

function showLoading() {
    loadingOverlay.style.display = 'flex';
}

function hideLoading() {
    loadingOverlay.style.display = 'none';
}

// ===== Start Application =====
init();