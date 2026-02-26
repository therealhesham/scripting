const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');
const puppeteer = require('puppeteer');

// ============================================
// CONFIGURATION
// ============================================
const KEYWORD_START = 'تقرير المستثمر';
const MIN_ROWS_PER_TABLE = 15;
const MAX_ROWS_PER_TABLE = 60;

// ============================================
// إعداد Multer لاستلام الملفات من الـ Request
// ============================================
const upload = multer({
    dest: 'uploads/',
    limits: {
        fileSize: 50 * 1024 * 1024, // حد أقصى للحجم 50 ميجا
        fieldSize: 50 * 1024 * 1024
    }
});

const app = express();
app.use(cors('*')); // السماح لكل النطاقات

app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.json({ limit: '50mb' }));

// مجلد أساسي دائم للاحتفاظ بالملفات المخرجة على الـ VPS
const FINAL_OUTPUT_DIR = process.env.OUTPUT_DIR || path.join(__dirname, 'vps_extracted_files');
if (!fs.existsSync(FINAL_OUTPUT_DIR)) {
    fs.mkdirSync(FINAL_OUTPUT_DIR, { recursive: true });
}

// عرض المجلد كـ Static Files لكي يمكن الوصول للملفات بروابط مباشرة
app.use('/files', express.static(FINAL_OUTPUT_DIR));

// Load header/footer globals
let GLOBAL_HEADER_BASE64 = '';
const headerPath = path.join(__dirname, 'header.png');
if (fs.existsSync(headerPath)) {
    GLOBAL_HEADER_BASE64 = fs.readFileSync(headerPath).toString('base64');
}

let GLOBAL_FOOTER_BASE64 = '';
const footerPath = path.join(__dirname, 'footer.jpg');
if (fs.existsSync(footerPath)) {
    GLOBAL_FOOTER_BASE64 = fs.readFileSync(footerPath).toString('base64');
}

// ============================================
// HELPERS FROM ELEGANT SCRIPT
// ============================================

function getSheetMerges(sheet) {
    const mergeMap = new Map(); // "r,c" -> {rowspan, colspan, master: {r,c}}
    if (sheet.model.merges) {
        sheet.model.merges.forEach(rangeStr => {
            const [start, end] = rangeStr.split(':');
            const s = decodeAddress(start);
            const e = decodeAddress(end);

            const rowspan = e.row - s.row + 1;
            const colspan = e.col - s.col + 1;

            mergeMap.set(`${s.row},${s.col}`, { rowspan, colspan, isMaster: true });

            for (let r = s.row; r <= e.row; r++) {
                for (let c = s.col; c <= e.col; c++) {
                    if (r === s.row && c === s.col) continue;
                    mergeMap.set(`${r},${c}`, { isMaster: false, master: { r: s.row, c: s.col } });
                }
            }
        });
    }
    return mergeMap;
}

function decodeAddress(addr) {
    const colMatch = addr.match(/[A-Z]+/)[0];
    const rowMatch = addr.match(/[0-9]+/)[0];
    let col = 0;
    for (let i = 0; i < colMatch.length; i++) {
        col = col * 26 + (colMatch.charCodeAt(i) - 64);
    }
    return { row: parseInt(rowMatch), col: col };
}

function findTablesInSheet(sheet, merges) {
    let hits = [];
    const maxScanRow = 100;

    sheet.eachRow((row, rowNumber) => {
        if (rowNumber > maxScanRow) return;
        row.eachCell((cell, colNumber) => {
            if (cell.value && cell.value.toString().includes(KEYWORD_START)) {
                const mergeInfo = merges.get(`${rowNumber},${colNumber}`);
                if (mergeInfo && !mergeInfo.isMaster) return;

                hits.push({ startRow: rowNumber, startCol: colNumber });
            }
        });
    });

    hits.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));

    const uniqueTables = [];
    for (const hit of hits) {
        const isDuplicate = uniqueTables.some(t =>
            Math.abs(t.startRow - hit.startRow) <= 2 &&
            Math.abs(t.startCol - hit.startCol) <= 3
        );
        if (!isDuplicate) uniqueTables.push(hit);
    }

    for (let i = 0; i < uniqueTables.length; i++) {
        const current = uniqueTables[i];
        const nextOnRow = uniqueTables.find(t => t.startRow === current.startRow && t.startCol > current.startCol);

        let width = 4;
        if (nextOnRow) {
            const dist = nextOnRow.startCol - current.startCol;
            width = dist > 1 ? dist : 4;
        } else {
            width = 4;
        }
        current.width = width;
    }

    return uniqueTables;
}

function getSafeCellValue(cell) {
    let val = cell.value;
    if (val === null || val === undefined) return '';

    let rawValue = val;
    if (typeof val === 'object' && !(val instanceof Date)) {
        if (val.richText) {
            rawValue = val.richText.map(t => t.text).join('');
        } else if (val.text) {
            rawValue = val.text;
        } else if (val.result !== undefined) {
            rawValue = val.result;
        } else if (val.error) {
            rawValue = '';
        } else {
            try {
                const json = JSON.stringify(val);
                if (json === '{}' || json.includes('error')) rawValue = '';
                else rawValue = '';
            } catch (e) { rawValue = ''; }
        }
    }

    if (typeof rawValue === 'number') {
        const num = rawValue;
        if (Math.abs(num) < 0.000001) {
            const fmt = cell.numFmt || '';
            if (fmt.includes('"-"') || fmt.includes(' - ') || fmt.includes('_-')) return '-';
            else return '0';
        }
        if (Number.isInteger(num)) return num.toString();
        else return num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
    }
    else if (rawValue instanceof Date) {
        const day = rawValue.getDate().toString().padStart(2, '0');
        const month = (rawValue.getMonth() + 1).toString().padStart(2, '0');
        const year = rawValue.getFullYear();
        return `${day}/${month}/${year}`;
    }

    return rawValue ? rawValue.toString() : '';
}

function getExcelThemeColor(theme, tint) {
    const themes = {
        0: [255, 255, 255], 1: [0, 0, 0], 2: [231, 230, 230], 3: [68, 84, 106],
        4: [68, 114, 196], 5: [237, 125, 49], 6: [165, 165, 165], 7: [255, 192, 0],
        8: [91, 155, 213], 9: [112, 173, 71]
    };
    let rgb = themes[theme] || [255, 255, 255];
    if (tint !== undefined && tint !== 0) {
        if (tint > 0) rgb = rgb.map(c => c + (255 - c) * tint);
        else rgb = rgb.map(c => c * (1 + tint));
    }
    return '#' + rgb.map(c => Math.round(c).toString(16).padStart(2, '0')).join('');
}

function getStyleFromCell(cell) {
    let style = 'border: 1px solid #444; padding: 4px; font-size: 10pt; font-family: Arial, sans-serif; text-align: center; vertical-align: middle;';

    if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
        if (cell.fill.fgColor.argb) {
            let color = cell.fill.fgColor.argb;
            if (color.length === 8) color = color.substring(2);
            style += `background-color: #${color}; `;
        } else if (cell.fill.fgColor.theme !== undefined) {
            let color = getExcelThemeColor(cell.fill.fgColor.theme, cell.fill.fgColor.tint);
            style += `background-color: ${color}; `;
        }
    }

    if (cell.font) {
        if (cell.font.bold) style += 'font-weight: bold; ';
        if (cell.font.color) {
            if (cell.font.color.argb) {
                let color = cell.font.color.argb;
                if (color.length === 8) color = color.substring(2);
                style += `color: #${color}; `;
            } else if (cell.font.color.theme !== undefined) {
                let color = getExcelThemeColor(cell.font.color.theme, cell.font.color.tint);
                style += `color: ${color}; `;
            }
        }
    }

    if (cell.alignment && cell.alignment.horizontal) {
        let align = cell.alignment.horizontal;
        if (align === 'centerContinuous') align = 'center';
        style += `text-align: ${align}; `;
    }

    return style;
}

function generateHTML(sheet, startRow, startCol, endRow, endCol, merges, headerBase64, footerBase64) {
    let rows = '';

    for (let r = startRow; r <= endRow; r++) {
        rows += '<tr>';
        for (let c = startCol; c <= endCol; c++) {
            const mergeInfo = merges.get(`${r},${c}`);
            if (mergeInfo && !mergeInfo.isMaster) continue;

            const cell = sheet.getCell(r, c);
            let attrs = '';
            if (mergeInfo && mergeInfo.isMaster) {
                if (mergeInfo.rowspan > 1) attrs += ` rowspan="${mergeInfo.rowspan}"`;
                if (mergeInfo.colspan > 1) attrs += ` colspan="${mergeInfo.colspan}"`;
            }

            let text = getSafeCellValue(cell);
            let style = getStyleFromCell(cell);
            rows += `<td${attrs} style="${style}">${text}</td>`;
        }
        rows += '</tr>';
    }

    const headerHTML = `
        <div style="width: 100%; margin-bottom: 20px; border-bottom: 0px solid #ccc; padding-bottom: 0px;">
            ${headerBase64 ? `<img src="data:image/png;base64,${headerBase64}" style="width: 100%; height: auto;" />` : '<div style="font-size: 24px; font-weight: bold; color: firebrick; text-align: center;">HEADER IMAGE NOT FOUND (header.png)</div>'}
        </div>
    `;

    const footerHTML = footerBase64 ? `
        <div class="page-footer">
            <img src="data:image/jpeg;base64,${footerBase64}" style="width: 100%; height: auto; display: block;" />
        </div>
    ` : '';

    return `
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <style>
            body { margin: 0; padding: 0; display: block; background: #fff; font-family: Arial, sans-serif; }
            table { border-collapse: collapse; width: 100%; direction: rtl; margin-top: 10px; }
            td { word-wrap: break-word; }
            .page-footer {
                position: fixed;
                bottom: 0;
                left: 0;
                right: 0;
                width: 100%;
                margin: 0;
                padding: 0;
                font-size: 0;
                line-height: 0;
            }
            @media print {
                body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                @page { size: A4 portrait; margin: 0; }
                .break-avoid { page-break-inside: avoid; }
            }
        </style>
    </head>
    <body>
        <div style="width: 100%; max-width: 190mm; margin: 0 auto; padding-bottom: 30mm;">
            ${headerHTML}
            <table>${rows}</table>
        </div>
        ${footerHTML}
    </body>
    </html>
    `;
}

function sanitizeFilename(name) {
    if (!name) return 'Unknown';
    return name.replace(/[^a-zA-Z0-9\u0600-\u06FF \-_]/g, '').trim().substring(0, 50);
}

async function processTable(sheet, tableLoc, page, merges) {
    const { startRow, startCol, width } = tableLoc;
    let endCol = startCol + width - 1;

    let endRow = startRow + MIN_ROWS_PER_TABLE;
    let emptyCount = 0;

    for (let r = endRow; r < startRow + MAX_ROWS_PER_TABLE; r++) {
        let hasData = false;
        for (let c = startCol; c <= endCol; c++) {
            const cell = sheet.getCell(r, c);
            if (cell.value) { hasData = true; break; }
        }

        if (!hasData) {
            emptyCount++;
            if (emptyCount >= 2) {
                endRow = r - 2;
                break;
            }
        } else {
            emptyCount = 0;
            endRow = r;
        }
    }

    for (let c = endCol; c > startCol; c--) {
        let hasDataInCol = false;
        for (let r = startRow; r <= endRow; r++) {
            const cell = sheet.getCell(r, c);
            if (cell.value) {
                hasDataInCol = true;
                break;
            }
        }
        if (!hasDataInCol) {
            endCol--;
        } else {
            break;
        }
    }

    let investorName = "Unknown";
    let plateNumber = "Unknown";

    for (let c = startCol; c <= endCol; c++) {
        const val = getSafeCellValue(sheet.getCell(startRow, c));
        if (val && val.includes(KEYWORD_START)) {
            const parts = val.split(':');
            if (parts[1]) investorName = parts[1].split('(')[0].trim();
            break;
        }
    }

    let foundPlate = false;
    for (let r = startRow; r < startRow + 15; r++) {
        for (let c = startCol; c <= endCol; c++) {
            const val = getSafeCellValue(sheet.getCell(r, c));
            if (val && (val.includes('لوحة') || val.includes('اللوحة'))) {
                let candidate = getSafeCellValue(sheet.getCell(r, c + 1));
                if (!candidate) candidate = getSafeCellValue(sheet.getCell(r + 1, c));
                if (!candidate) candidate = getSafeCellValue(sheet.getCell(r, c + 2));

                if (candidate) {
                    plateNumber = candidate.trim();
                    foundPlate = true;
                }
            }
        }
        if (foundPlate) break;
    }

    if (!foundPlate) {
        for (let r = startRow; r < startRow + 15; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const val = getSafeCellValue(sheet.getCell(r, c));
                if (val && val.includes('نوع السيارة')) {
                    let candidate = getSafeCellValue(sheet.getCell(r, c + 1));
                    if (candidate) {
                        plateNumber = candidate.trim();
                        foundPlate = true;
                    }
                }
            }
        }
    }

    investorName = sanitizeFilename(investorName);
    plateNumber = sanitizeFilename(plateNumber);
    if (investorName === "Unknown") investorName = sanitizeFilename(sheet.name);
    if (plateNumber === "Unknown") plateNumber = `Report_Col${startCol}`;

    const investorDir = path.join(FINAL_OUTPUT_DIR, investorName);
    if (!fs.existsSync(investorDir)) {
        fs.mkdirSync(investorDir, { recursive: true });
    }

    let baseFilename = `${investorName}_${plateNumber}`;
    let filename = `${baseFilename}.pdf`;
    let counter = 1;
    while (fs.existsSync(path.join(investorDir, filename))) {
        filename = `${baseFilename}_${counter}.pdf`;
        counter++;
    }

    const finalPath = path.join(investorDir, filename);

    const html = generateHTML(sheet, startRow, startCol, endRow, endCol, merges, GLOBAL_HEADER_BASE64, GLOBAL_FOOTER_BASE64);
    await page.setContent(html, { waitUntil: 'load' });

    await page.pdf({
        path: finalPath,
        format: 'A4',
        landscape: false,
        printBackground: true,
        displayHeaderFooter: false,
        margin: { top: '10mm', right: '0mm', bottom: '0mm', left: '10mm' }
    });

    return investorName;
}

async function mergeInvestorPDFs(investorName) {
    const investorDir = path.join(FINAL_OUTPUT_DIR, investorName);
    if (!fs.existsSync(investorDir)) return null;

    const files = fs.readdirSync(investorDir)
        .filter(f => f.endsWith('.pdf') && !f.endsWith(' 2025.pdf'));

    if (files.length === 0) return null;

    const countHyphens = (str) => (str.match(/-/g) || []).length;

    files.sort((a, b) => {
        const hyphensA = countHyphens(a);
        const hyphensB = countHyphens(b);
        if (hyphensA !== hyphensB) {
            return hyphensA - hyphensB;
        }
        return a.localeCompare(b, 'ar');
    });

    try {
        const mergedPdf = await PDFDocument.create();
        for (const file of files) {
            const pdfBytes = fs.readFileSync(path.join(investorDir, file));
            const pdf = await PDFDocument.load(pdfBytes);
            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
            copiedPages.forEach((page) => mergedPdf.addPage(page));
        }

        const mergedPdfBytes = await mergedPdf.save();
        const outputName = `${investorName} 2025.pdf`;
        const outputPath = path.join(investorDir, outputName);
        fs.writeFileSync(outputPath, mergedPdfBytes);

        // Delete individual unmerged pdf parts (optional but cleaner)
        for (const file of files) {
            try {
                fs.unlinkSync(path.join(investorDir, file));
            } catch (ignore) { }
        }

        return outputName;
    } catch (error) {
        console.error(`Error merging PDFs for ${investorName}:`, error);
        return null;
    }
}

// ============================================
// نقطة الـ API الأساسية لعملية الاستخراج
// ============================================
app.post('/extracting', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ status: 'error', message: 'لم يتم إرسال أي ملف. الرجاء إرسال الملف باسم الحقل "file".' });
        }

        const EXCEL_FILE = req.file.path;
        const uniqueId = Date.now() + '_' + Math.floor(Math.random() * 10000);

        console.log(`⏳ جاري قراءة ملف الإكسل المرسل لتحليل الجداول... (Request ID: ${uniqueId})`);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(EXCEL_FILE);

        const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox', '--disable-setuid-sandbox'] });
        const page = await browser.newPage();

        const processedInvestors = new Set();
        let totalTablesProcessed = 0;

        for (const sheet of workbook.worksheets) {
            const merges = getSheetMerges(sheet);
            const tables = findTablesInSheet(sheet, merges);

            if (tables.length > 0) {
                console.log(`Processing sheet: ${sheet.name} (Found ${tables.length} unique tables)`);
                for (const table of tables) {
                    const investorName = await processTable(sheet, table, page, merges);
                    if (investorName) {
                        processedInvestors.add(investorName);
                        totalTablesProcessed++;
                    }
                }
            }
        }

        await browser.close();

        if (totalTablesProcessed === 0) {
            if (fs.existsSync(EXCEL_FILE)) fs.unlinkSync(EXCEL_FILE);
            return res.status(404).json({ status: 'warning', message: 'لم يتم العثور على أي جداول للطباعة.' });
        }

        console.log(`\n🖨️ بدء دمج الملفات لـ PDF...`);
        const investorLinks = {};
        const baseUrl = req.protocol + '://' + req.get('host') + '/files';

        let successCount = 0;

        for (const investorName of processedInvestors) {
            const mergedFileName = await mergeInvestorPDFs(investorName);
            if (mergedFileName) {
                successCount++;
                const encodedFolder = encodeURIComponent(investorName);
                const encodedFile = encodeURIComponent(mergedFileName);
                investorLinks[investorName] = [`${baseUrl}/${encodedFolder}/${encodedFile}`];
            }
        }

        if (fs.existsSync(EXCEL_FILE)) fs.unlinkSync(EXCEL_FILE);

        console.log(`✅ انتهت المهمة. ورفع الملفات لـ VPS.`);

        res.status(200).json({
            status: 'success',
            message: `تم استخراج ودمج ${successCount} ملف PDF نهائي بنجاح من أصل ${totalTablesProcessed} جدول.`,
            total_jobs: totalTablesProcessed,
            success_count: successCount,
            investors_files: investorLinks
        });

    } catch (error) {
        console.error(error);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ status: 'error', message: 'خطأ داخلي في الخادم.', error: error.message });
    }
});

const PORT = process.env.PORT || 3172;
app.listen(PORT, () => {
    console.log(`🚀 السيرفر يعمل الآن على http://localhost:${PORT}`);
    console.log(`📡 يمكنك عمل طلب POST على http://localhost:${PORT}/extracting وإرفاق الملف كـ form-data تحت اسم (file)`);
});
