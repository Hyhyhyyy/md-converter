// ===== Global Variables =====
let currentFile = null;
let markdownContent = '';

// ===== DOM Elements =====
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const selectBtn = document.getElementById('selectBtn');
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
    // File selection
    selectBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop
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
        // Create PDF container
        const pdfContainer = document.createElement('div');
        pdfContainer.className = 'pdf-container';
        pdfContainer.innerHTML = marked.parse(markdownContent);

        // PDF options
        const opt = {
            margin: [10, 10, 10, 10],
            filename: getExportFilename('.pdf'),
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                letterRendering: true
            },
            jsPDF: {
                unit: 'mm',
                format: 'a4',
                orientation: 'portrait'
            }
        };

        await html2pdf().set(opt).from(pdfContainer).save();

        hideLoading();
    } catch (error) {
        hideLoading();
        alert('PDF 导出失败: ' + error.message);
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
        const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } = docx;

        // Parse markdown and convert to docx paragraphs
        const doc = new Document({
            sections: [{
                properties: {},
                children: parseMarkdownToDocx(markdownContent)
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
    const { Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;
    const paragraphs = [];
    const lines = content.split('\n');

    let inCodeBlock = false;
    let codeContent = '';

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // Code block handling
        if (line.startsWith('```')) {
            if (inCodeBlock) {
                // End code block
                paragraphs.push(new Paragraph({
                    children: [new TextRun({
                        text: codeContent,
                        font: 'Consolas',
                        size: 20
                    })]
                }));
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

        // Empty line
        if (line.trim() === '') {
            paragraphs.push(new Paragraph({ children: [] }));
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

            paragraphs.push(new Paragraph({
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
            paragraphs.push(new Paragraph({
                border: {
                    bottom: { style: BorderStyle.SINGLE, size: 1 }
                },
                children: []
            }));
            continue;
        }

        // Blockquote
        if (line.startsWith('>')) {
            paragraphs.push(new Paragraph({
                indent: { left: 720 },
                children: [new TextRun({
                    text: line.replace(/^>\s*/, '')
                })]
            }));
            continue;
        }

        // Unordered list
        if (line.match(/^[\*\-\+]\s/)) {
            paragraphs.push(new Paragraph({
                bullet: { level: 0 },
                children: [new TextRun({
                    text: line.replace(/^[\*\-\+]\s/, '')
                })]
            }));
            continue;
        }

        // Ordered list
        if (line.match(/^\d+\.\s/)) {
            paragraphs.push(new Paragraph({
                numbering: { reference: 'default-numbering', level: 0 },
                children: [new TextRun({
                    text: line.replace(/^\d+\.\s/, '')
                })]
            }));
            continue;
        }

        // Regular paragraph with inline formatting
        paragraphs.push(new Paragraph({
            children: parseInlineFormatting(line)
        }));
    }

    return paragraphs;
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
        if (italicMatch && italicMatch.index === 0 && !remaining.startsWith('**')) {
            runs.push(new TextRun({
                text: italicMatch[1] || italicMatch[2],
                italics: true
            }));
            remaining = remaining.substring(italicMatch[0].length);
            continue;
        }

        // Inline code `code`
        const codeMatch = remaining.match(/`(.+?)`/);
        if (codeMatch && codeMatch.index === 0) {
            runs.push(new TextRun({
                text: codeMatch[1],
                font: 'Consolas',
                size: 20
            }));
            remaining = remaining.substring(codeMatch[0].length);
            continue;
        }

        // Regular text until next special character
        const nextSpecial = remaining.search(/\*\*|__|\*|_|`/);
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