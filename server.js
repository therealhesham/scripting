const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors('*')); // السماح لكل النطاقات
app.use(express.json());

const EXCEL_FILE = path.join(__dirname, 'data.xlsx');
const OUTPUT_DIR = path.join(__dirname, 'مخرجات_السيارات');
const MAIN_TAB = 'قائمة المستثمرين';

// ============================================
// دالة مساعدة لاستخراج النص الصافي من الخلية
// ============================================
const getCellString = (cell) => {
    if (!cell || cell.value === null || cell.value === undefined) return '';
    if (typeof cell.value === 'object') {
        if (cell.value.result !== undefined) return String(cell.value.result);
        if (cell.value.richText) return cell.value.richText.map(r => r.text).join('');
        if (cell.value.text) return String(cell.value.text);
        return '';
    }
    return String(cell.value);
};

// ============================================
// دالة مساعدة لتنظيف وتطابق أسماء الـ Sheets
// ============================================
const normalizeKey = (s) => {
    return String(s || '').replace(/\s+/g, ' ').trim().normalize('NFKC')
        .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه')
        .replace(/ؤ/g, 'و').replace(/ئ/g, 'ي').replace(/[\u064B-\u065F]/g, '')
        .replace(/[^\p{L}\p{N}\s]+/gu, '');
};

const findInvestorSheet = (wb, investorName) => {
    const target = String(investorName || '').replace(/\s+/g, ' ').trim();
    const exact = wb.worksheets.find(w => String(w.name).replace(/\s+/g, ' ').trim() === target);
    if (exact) return exact;
    const targetKey = normalizeKey(investorName);
    const byKey = wb.worksheets.find(w => normalizeKey(w.name) === targetKey);
    if (byKey) return byKey;

    for (const w of wb.worksheets) {
        if (w.name === MAIN_TAB) continue;
        const sheetKey = normalizeKey(w.name);
        if (sheetKey.includes(targetKey) || targetKey.includes(sheetKey)) return w;
        const targetWords = targetKey.split(' ');
        const sheetWords = sheetKey.split(' ');
        let matches = 0;
        for (const tw of targetWords) {
            if (tw.length > 2 && sheetWords.includes(tw)) matches++;
        }
        if (matches >= 2 && Math.abs(targetWords.length - sheetWords.length) <= 2) return w;
    }
    return null;
};

// ============================================
// نقطة الـ API الأساسية لعملية الاستخراج
// ============================================
app.post('/extracting', async (req, res) => {
    try {
        if (!fs.existsSync(OUTPUT_DIR)) {
            fs.mkdirSync(OUTPUT_DIR, { recursive: true });
        }

        console.log('⏳ جاري قراءة ملف الإكسل لتحليل الجداول...');
        const workbook = new ExcelJS.Workbook();

        if (!fs.existsSync(EXCEL_FILE)) {
            return res.status(400).json({ status: 'error', message: 'ملف data.xlsx غير موجود في المجلد.' });
        }

        await workbook.xlsx.readFile(EXCEL_FILE);

        const mainSheet = workbook.getWorksheet(MAIN_TAB);
        if (!mainSheet) {
            return res.status(400).json({ status: 'error', message: `لم يتم العثور على شيت باسم: "${MAIN_TAB}"` });
        }

        let headerRowIndex = -1;
        let colInvestor = -1;
        let colCarsCount = -1;

        for (let r = 1; r <= Math.min(20, mainSheet.rowCount); r++) {
            const row = mainSheet.getRow(r);
            row.eachCell((cell, colNumber) => {
                const text = cell.value ? cell.value.toString().trim() : '';
                if (text.includes('اسم المستثمر')) colInvestor = colNumber;
                if (text.includes('عدد السيارات')) colCarsCount = colNumber;
            });
            if (colInvestor !== -1 && colCarsCount !== -1) {
                headerRowIndex = r;
                break;
            }
        }

        if (headerRowIndex === -1) {
            return res.status(400).json({ status: 'error', message: 'لم يتم العثور على أعمدة "اسم المستثمر" و "عدد السيارات" في القائمة.' });
        }

        const START_ROW = headerRowIndex + 1;
        const printJobs = [];

        for (let r = START_ROW; r <= mainSheet.rowCount; r++) {
            const row = mainSheet.getRow(r);
            const investorName = row.getCell(colInvestor).value?.toString()?.trim();
            let carsCount = row.getCell(colCarsCount).value;

            if (carsCount && typeof carsCount === 'object' && carsCount.result !== undefined) {
                carsCount = carsCount.result;
            }
            carsCount = parseInt(carsCount);

            if (!investorName || !carsCount || isNaN(carsCount)) continue;

            const investorSheet = findInvestorSheet(workbook, investorName);
            if (!investorSheet) {
                console.log(`⚠️ تحذير: لم يتم العثور على شيت باسم ( ${investorName} )`);
                continue;
            }
            const actualSheetName = investorSheet.name;

            console.log(`\n👨‍💼 تحليل جداول المستثمر: ${investorName} (${carsCount} سيارة)`);

            let startCol = 1;

            for (let carIndex = 1; carIndex <= carsCount; carIndex++) {
                let endCol = startCol;
                while (true) {
                    let hasData = false;
                    for (let rowIdx = 1; rowIdx <= investorSheet.rowCount; rowIdx++) {
                        let strVal = getCellString(investorSheet.getRow(rowIdx).getCell(endCol)).trim();
                        if (strVal !== '') {
                            hasData = true;
                            break;
                        }
                    }

                    if (!hasData && endCol > startCol) {
                        endCol--;
                        break;
                    }
                    endCol++;
                    if (endCol > 500) break;
                }

                let lastRow = 1;
                for (let rowIdx = 1; rowIdx <= investorSheet.rowCount; rowIdx++) {
                    let rowHasData = false;
                    for (let c = startCol; c <= endCol; c++) {
                        if (getCellString(investorSheet.getRow(rowIdx).getCell(c)).trim() !== '') {
                            rowHasData = true;
                            break;
                        }
                    }
                    if (rowHasData) lastRow = rowIdx;
                }

                const sanitizeName = (name) => name.replace(/[<>:"/\\|?*]+/g, '_').trim();
                const investorFolder = path.join(OUTPUT_DIR, sanitizeName(investorName));

                if (!fs.existsSync(investorFolder)) {
                    fs.mkdirSync(investorFolder, { recursive: true });
                }

                const pdfFileName = path.join(investorFolder, `${sanitizeName(investorName)} - سيارة ${carIndex}.pdf`);

                printJobs.push({
                    sheetName: actualSheetName,
                    investorName: investorName,
                    startCol,
                    endCol,
                    lastRow,
                    outputFile: pdfFileName
                });

                startCol = endCol + 1;
                while (startCol <= 500) {
                    let checkData = false;
                    for (let rowIdx = 1; rowIdx <= investorSheet.rowCount; rowIdx++) {
                        if (getCellString(investorSheet.getRow(rowIdx).getCell(startCol)).trim() !== '') {
                            checkData = true;
                            break;
                        }
                    }
                    if (checkData) break;
                    startCol++;
                }
            }
        }

        if (printJobs.length === 0) {
            return res.status(404).json({ status: 'warning', message: 'لم يتم العثور على أي جداول للطباعة.' });
        }

        console.log(`\n🚀 تم التعرف على ${printJobs.length} جدول. جاري معالجتها لتعمل على (Ubuntu / Linux) باستخدام LibreOffice...`);

        const TMP_DIR = path.join(__dirname, 'temp_excel');
        if (!fs.existsSync(TMP_DIR)) fs.mkdirSync(TMP_DIR);

        try {
            // 1. إنشاء ملف إكسل مؤقت لكل جدول لضمان قراءة LibreOffice له كجدول مستقل
            for (let i = 0; i < printJobs.length; i++) {
                const job = printJobs[i];
                process.stdout.write(`⚙️ تجهيز الملف المؤقت: ${carIndex = job.outputFile.split(' ').pop()} ... `);

                const tempWb = new ExcelJS.Workbook();
                await tempWb.xlsx.readFile(EXCEL_FILE);

                let targetSheetId = null;
                tempWb.eachSheet((sheet, id) => {
                    if (sheet.name === job.sheetName) targetSheetId = id;
                });

                // حذف جميع الشيتات باستثناء الشيت المطلوب لكي لا يتم طباعتها في الـ PDF
                const sheetIdsToRemove = [];
                tempWb.eachSheet((sheet, id) => {
                    if (id !== targetSheetId) sheetIdsToRemove.push(id);
                });
                sheetIdsToRemove.forEach(id => tempWb.removeWorksheet(id));

                // إخفاء الأعمدة والصفوف الأخرى
                const targetSheet = tempWb.getWorksheet(targetSheetId);
                if (targetSheet) {
                    const maxCols = targetSheet.columnCount;
                    for (let c = 1; c <= maxCols + 5; c++) {
                        if (c < job.startCol || c > job.endCol) targetSheet.getColumn(c).hidden = true;
                    }
                    const maxRows = targetSheet.rowCount;
                    for (let r = 1; r <= maxRows + 20; r++) {
                        if (r > job.lastRow) targetSheet.getRow(r).hidden = true;
                    }

                    targetSheet.pageSetup.fitToPage = true;
                    targetSheet.pageSetup.fitToWidth = 1;
                    targetSheet.pageSetup.fitToHeight = 1;
                    targetSheet.views = [{ rightToLeft: true }];
                }

                job.tempXlsxPath = path.join(TMP_DIR, `job_${i}.xlsx`);
                job.tempPdfPath = path.join(TMP_DIR, `job_${i}.pdf`);
                await tempWb.xlsx.writeFile(job.tempXlsxPath);
                console.log('تم!');
            }

            console.log(`\n🖨️ بدء طباعة جميع الملفات لـ PDF عبر LibreOffice...`);
            const isWin = process.platform === "win32";
            const libreCmd = isWin ? 'soffice' : 'libreoffice';

            for (let i = 0; i < printJobs.length; i++) {
                const job = printJobs[i];
                try {
                    execSync(`${libreCmd} --headless --convert-to pdf "${job.tempXlsxPath}" --outdir "${TMP_DIR}"`, { stdio: 'ignore' });
                } catch (e) {
                    console.log(`⚠️ فشل تحويل الملف ${job.tempXlsxPath}. تأكد من تثبيت LibreOffice.`);
                }
            }

            console.log(`\n� نقل ملفات الـ PDF إلى الفولدرات النهائية...`);
            let successCount = 0;
            for (let i = 0; i < printJobs.length; i++) {
                const job = printJobs[i];
                if (fs.existsSync(job.tempPdfPath)) {
                    fs.renameSync(job.tempPdfPath, job.outputFile);
                    successCount++;
                }
                // تنظيف الملف المُصدر
                if (fs.existsSync(job.tempXlsxPath)) fs.unlinkSync(job.tempXlsxPath);
            }

            res.status(200).json({
                status: 'success',
                message: `تم استخراج ${successCount} ملف PDF بنجاح من أصل ${printJobs.length}.`,
                total_jobs: printJobs.length,
                success_count: successCount,
                output_dir: OUTPUT_DIR
            });
        } catch (err) {
            console.error(err);
            res.status(500).json({ status: 'error', message: 'حدث خطأ أثناء تجهيز الملفات أو تشغيل LibreOffice لتصدير PDF.', error: err.message });
        }

    } catch (error) {
        console.error(error);
        res.status(500).json({ status: 'error', message: 'خطأ داخلي في الخادم.', error: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 السيرفر يعمل الآن على http://localhost:${PORT}`);
    console.log(`📡 يمكنك عمل طلب POST على http://localhost:${PORT}/extracting`);
});
