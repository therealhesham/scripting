const { PDFDocument } = require('pdf-lib');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

// ==========================================
// CONFIGURATION
// ==========================================
const INPUT_FILE = 'input.xlsx';
const OUTPUT_DIR = './output_pdfs';
const KEYWORD_START = 'تقرير المستثمر';
const MIN_ROWS_PER_TABLE = 15;
const MAX_ROWS_PER_TABLE = 60;

// ==========================================
// MAIN LOGIC
// ==========================================

async function main() {
    console.log(`Starting processing for file: ${INPUT_FILE}`);

    if (!fs.existsSync(INPUT_FILE)) {
        console.error(`Error: File "${INPUT_FILE}" not found.`);
        return;
    }

    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(INPUT_FILE);

    // Read Header Image (replacing logo.png logic)
    if (fs.existsSync('header.png')) {
        const header = fs.readFileSync('header.png');
        GLOBAL_HEADER_BASE64 = header.toString('base64');
    }

    if (fs.existsSync('footer.jpg')) {
        const footer = fs.readFileSync('footer.jpg');
        GLOBAL_FOOTER_BASE64 = footer.toString('base64');
    }

    const browser = await puppeteer.launch({ headless: 'new' });
    const page = await browser.newPage();

    console.log(`Workbook loaded. Sheets found: ${workbook.worksheets.length}`);

    // Track unique investors to know where to merge later
    const processedInvestors = new Set();

    for (const sheet of workbook.worksheets) {
        // Pre-calculate merges for the whole sheet
        const merges = getSheetMerges(sheet);

        const tables = findTablesInSheet(sheet, merges);

        if (tables.length > 0) {
            console.log(`Processing sheet: ${sheet.name} (Found ${tables.length} unique tables)`);
            for (const table of tables) {
                const investorName = await processTable(sheet, table, page, merges);
                if (investorName) processedInvestors.add(investorName);
            }
        }
    }

    await browser.close();

    // Merge Phase
    console.log('Starting PDF Merge phase...');
    for (const investorName of processedInvestors) {
        await mergeInvestorPDFs(investorName);
    }

    console.log('All done!');
}

async function mergeInvestorPDFs(investorName) {
    const investorDir = path.join(OUTPUT_DIR, investorName);
    if (!fs.existsSync(investorDir)) return;

    // Filter out the "merged" file itself if it already exists from a previous run to avoid recursion
    const files = fs.readdirSync(investorDir)
        .filter(f => f.endsWith('.pdf') && !f.endsWith(' 2025.pdf'));

    if (files.length === 0) return;

    // Convert hyphen counting logic:
    // Files with many hyphens are usually the "Summary" files.
    // We want Summary files to be LAST.
    // Files with few hyphens to be FIRST.
    const countHyphens = (str) => (str.match(/-/g) || []).length;

    files.sort((a, b) => {
        // 1. Sort by Hyphen Count (Ascending)
        const hyphensA = countHyphens(a);
        const hyphensB = countHyphens(b);
        if (hyphensA !== hyphensB) {
            return hyphensA - hyphensB;
        }
        // 2. Fallback to Alphabetical
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
        console.log(`Merged PDF created: ${outputName}`);
    } catch (error) {
        console.error(`Error merging PDFs for ${investorName}:`, error);
    }
}

function getSheetMerges(sheet) {
    const mergeMap = new Map(); // "r,c" -> {rowspan, colspan, master: {r,c}}
    // sheet.model.merges is array of "A1:B2" strings
    if (sheet.model.merges) {
        sheet.model.merges.forEach(rangeStr => {
            const [start, end] = rangeStr.split(':');
            const s = decodeAddress(start);
            const e = decodeAddress(end);

            const rowspan = e.row - s.row + 1;
            const colspan = e.col - s.col + 1;

            // Mark Master
            mergeMap.set(`${s.row},${s.col}`, { rowspan, colspan, isMaster: true });

            // Mark Slaves
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

                // If this cell is a slave merge, ignore it (we only want master)
                const mergeInfo = merges.get(`${rowNumber},${colNumber}`);
                if (mergeInfo && !mergeInfo.isMaster) {
                    return;
                }

                hits.push({
                    startRow: rowNumber,
                    startCol: colNumber
                });
            }
        });
    });

    hits.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));

    // Dedup just in case
    const uniqueTables = [];
    for (const hit of hits) {
        const isDuplicate = uniqueTables.some(t =>
            Math.abs(t.startRow - hit.startRow) <= 2 &&
            Math.abs(t.startCol - hit.startCol) <= 3
        );
        if (!isDuplicate) uniqueTables.push(hit);
    }

    // Calc Width
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

// Helper to safely extract text from any cell
function getSafeCellValue(cell) {
    let val = cell.value;

    if (val === null || val === undefined) {
        return '';
    }

    // 1. Unwrap ExcelJS object wrappers to get the raw value
    let rawValue = val;
    if (typeof val === 'object' && !(val instanceof Date)) {
        if (val.richText) {
            rawValue = val.richText.map(t => t.text).join('');
        } else if (val.text) {
            rawValue = val.text;
        } else if (val.result !== undefined) {
            rawValue = val.result; // Key for Formulas! This might be a number.
        } else if (val.error) {
            rawValue = ''; // Error
        } else {
            // Try stringify if unknown object
            try {
                const json = JSON.stringify(val);
                if (json === '{}' || json.includes('error')) rawValue = '';
                else rawValue = ''; // Hide
            } catch (e) { rawValue = ''; }
        }
    }

    // 2. Format based on Type of Raw Value
    if (typeof rawValue === 'number') {
        const num = rawValue;
        // Check for "0" being displayed as "-" (Accounting format)
        if (Math.abs(num) < 0.000001) {
            const fmt = cell.numFmt || '';
            if (fmt.includes('"-"') || fmt.includes(' - ') || fmt.includes('_-')) {
                return '-';
            } else {
                return '0';
            }
        }

        if (Number.isInteger(num)) {
            return num.toString();
        } else {
            // Force 2 decimal places max
            return num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
        }
    }
    else if (rawValue instanceof Date) {
        const day = rawValue.getDate().toString().padStart(2, '0');
        const month = (rawValue.getMonth() + 1).toString().padStart(2, '0');
        const year = rawValue.getFullYear();
        return `${day}/${month}/${year}`;
    }

    // Default String
    return rawValue ? rawValue.toString() : '';
}

async function processTable(sheet, tableLoc, page, merges) {
    const { startRow, startCol, width } = tableLoc;
    let endCol = startCol + width - 1;

    // Determine Height dynamically
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

    // TRIM EMPTY COLUMNS from the right
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
            endCol--; // Reduce width
        } else {
            break; // Stop trimming as soon as we hit data
        }
    }

    // Name Extraction
    let investorName = "Unknown";
    let plateNumber = "Unknown";

    // Investor Name
    for (let c = startCol; c <= endCol; c++) {
        const val = getSafeCellValue(sheet.getCell(startRow, c));
        if (val && val.includes(KEYWORD_START)) { // Now val is guaranteed string
            const parts = val.split(':');
            if (parts[1]) investorName = parts[1].split('(')[0].trim();
            break;
        }
    }

    // Plate Number & Multi-car Handling
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

    // Fallback: Use Car Type if Plate missing
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

    // Create Investor Directory
    const investorDir = path.join(OUTPUT_DIR, investorName);
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
    console.log(`    Generating PDF: ${filename} in ${investorName} (Cols ${startCol}-${endCol})`);

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

    return investorName; // Return investor name for tracking
}

// ------------------------------------------------------------------
// COLOR & STYLE HELPERS
// ------------------------------------------------------------------

function getExcelThemeColor(theme, tint) {
    // Approximate Office 2013-2022 Standard Theme Colors
    const themes = {
        0: [255, 255, 255], // Light 1 (White)
        1: [0, 0, 0],       // Dark 1 (Black)
        2: [231, 230, 230], // Light 2 (Lt Gray)
        3: [68, 84, 106],   // Dark 2 (Dk Blue-Gray)
        4: [68, 114, 196],  // Accent 1 (Blue)
        5: [237, 125, 49],  // Accent 2 (Orange)
        6: [165, 165, 165], // Accent 3 (Gray)
        7: [255, 192, 0],   // Accent 4 (Gold)
        8: [91, 155, 213],  // Accent 5 (Blue)
        9: [112, 173, 71]   // Accent 6 (Green)
    };

    let rgb = themes[theme] || [255, 255, 255];

    if (tint !== undefined && tint !== 0) {
        if (tint > 0) {
            // Lighter: Color + (White - Color) * Tint
            rgb = rgb.map(c => c + (255 - c) * tint);
        } else {
            // Darker: Color * (1 + Tint)
            rgb = rgb.map(c => c * (1 + tint));
        }
    }

    // Round and Convert to Hex
    return '#' + rgb.map(c => Math.round(c).toString(16).padStart(2, '0')).join('');
}

function getStyleFromCell(cell) {
    let style = 'border: 1px solid #444; padding: 4px; font-size: 10pt; font-family: Arial, sans-serif; text-align: center; vertical-align: middle;';

    // 1. BACKGROUND COLOR
    if (cell.fill && cell.fill.type === 'pattern') {
        if (cell.fill.fgColor) {
            if (cell.fill.fgColor.argb) {
                let color = cell.fill.fgColor.argb;
                if (color.length === 8) color = color.substring(2);
                style += `background-color: #${color}; `;
            } else if (cell.fill.fgColor.theme !== undefined) {
                let color = getExcelThemeColor(cell.fill.fgColor.theme, cell.fill.fgColor.tint);
                style += `background-color: ${color}; `;
            }
        }
    }

    // 2. FONT COLOR & BOLD
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

    // 3. ALIGNMENT
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

            // Check Merge Status
            const mergeInfo = merges.get(`${r},${c}`);

            if (mergeInfo && !mergeInfo.isMaster) {
                // This is a slave cell, SKIP IT
                continue;
            }

            const cell = sheet.getCell(r, c);

            let attrs = '';
            if (mergeInfo && mergeInfo.isMaster) {
                if (mergeInfo.rowspan > 1) attrs += ` rowspan="${mergeInfo.rowspan}"`;
                if (mergeInfo.colspan > 1) attrs += ` colspan="${mergeInfo.colspan}"`;
            }

            // Value Extraction
            let text = getSafeCellValue(cell);

            // Style Extraction with Theme/Tint Support
            let style = getStyleFromCell(cell);

            rows += `<td${attrs} style="${style}">${text}</td>`;
        }
        rows += '</tr>';
    }

    // UPDATED HEADER HTML: Uses Full Width Image
    const headerHTML = `
        <div style="width: 100%; margin-bottom: 20px; border-bottom: 0px solid #ccc; padding-bottom: 0px;">
            ${headerBase64 ? `<img src="data:image/png;base64,${headerBase64}" style="width: 100%; height: auto;" />` : '<div style="font-size: 24px; font-weight: bold; color: firebrick; text-align: center;">HEADER IMAGE NOT FOUND (header.png)</div>'}
        </div>
    `;

    // Footer HTML: Fixed position at the very bottom of each page
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

// GLOBAL for simplicity in this script
let GLOBAL_HEADER_BASE64 = '';
let GLOBAL_FOOTER_BASE64 = '';

main().catch(console.error);
