const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const EXCEL_FILE = path.join(__dirname, 'data.xlsx');
const OUTPUT_DIR = path.join(__dirname, 'مخرجات_السيارات');
const MAIN_TAB = 'قائمة المستثمرين';

async function main() {
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
  }

  console.log('⏳ جاري قراءة ملف الإكسل لتحليل الجداول...');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);

  const mainSheet = workbook.getWorksheet(MAIN_TAB);
  if (!mainSheet) {
    console.error(`❌ لم يتم العثور على شيت باسم: "${MAIN_TAB}"`);
    return;
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
    console.error('❌ لم يتم العثور على أعمدة "اسم المستثمر" و "عدد السيارات" في القائمة.');
    return;
  }

  const START_ROW = headerRowIndex + 1;

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

  const normalizeKey = (s) => {
    return String(s || '').replace(/\s+/g, ' ').trim().normalize('NFKC')
      .replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه')
      .replace(/ؤ/g, 'و').replace(/ئ/g, 'ي').replace(/[\u064B-\u065F]/g, '')
      .replace(/[^\p{L}\p{N}\s]+/gu, ''); // added \s to keep spaces
  };

  const findInvestorSheet = (wb, investorName) => {
    const target = String(investorName || '').replace(/\s+/g, ' ').trim();
    const exact = wb.worksheets.find(w => String(w.name).replace(/\s+/g, ' ').trim() === target);
    if (exact) return exact;

    const targetKey = normalizeKey(investorName);
    const byKey = wb.worksheets.find(w => normalizeKey(w.name) === targetKey);
    if (byKey) return byKey;

    // Advanced fallback: try substring match or word matching
    for (const w of wb.worksheets) {
      if (w.name === MAIN_TAB) continue;
      const sheetKey = normalizeKey(w.name);
      // عبيدالله مبارك العوفي -> عبيدالله العوفى
      // احمد عبيد الله العوفى -> احمد عبيدالله
      // ماطر ناير راشد العلوني الجهني -> ماطر ناير راشد العلواني الجهني
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

  // قائمة المهام التي سنسلمها للـ PowerShell
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

      // البحث عن آخر صف يحتوي على بيانات في هذا النطاق من الأعمدة لضبط الجدول
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

      console.log(`   ✔️ جدول ${carIndex} محدد من عمود ${startCol} إلى ${endCol}, وآخر صف ${lastRow}`);

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
    console.log('❌ لم يتم العثور على أي جداول للطباعة.');
    return;
  }

  console.log(`\n🚀 تم التعرف على ${printJobs.length} جدول. جاري الآن طباعتها صوره طبق الأصل من Excel...`);

  // بناء كود PowerShell
  const ps1ScriptPath = path.join(__dirname, 'export_jobs.ps1');

  let psCode = `
$ErrorActionPreference = "Stop"
try {
    Write-Host "Opening Excel Application in background..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open('${EXCEL_FILE.replace(/'/g, "''")}')
`;

  for (let i = 0; i < printJobs.length; i++) {
    const job = printJobs[i];
    psCode += `
    Write-Host "Exporting: ${path.basename(job.outputFile)}"
    $ws = $wb.Sheets.Item('${job.sheetName.replace(/'/g, "''")}')
    # من الصف 1 إلى آخر صف، ومن أول عمود للجدول لآخر عمود للجدول
    $range = $ws.Range($ws.Cells.Item(1, ${job.startCol}), $ws.Cells.Item(${job.lastRow}, ${job.endCol}))
    $ws.PageSetup.Zoom = $false
    $ws.PageSetup.FitToPagesWide = 1
    $ws.PageSetup.FitToPagesTall = 1
    $range.ExportAsFixedFormat(0, '${job.outputFile.replace(/'/g, "''")}')
`;
  }

  psCode += `
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "All done!"
} catch {
    Write-Host "An error occurred: $_"
    if ($wb) { $wb.Close($false) }
    if ($excel) { 
        $excel.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    exit 1
}
`;

  fs.writeFileSync(ps1ScriptPath, '\uFEFF' + psCode, 'utf8');

  try {
    execSync(`powershell -ExecutionPolicy Bypass -File "${ps1ScriptPath}"`, { stdio: 'inherit' });
    console.log('\n✅ اكتملت العملية بنجاح! جميع ملفات الـ PDF صورة طبق الأصل الآن.');
  } catch (err) {
    console.error('\n❌ حدث خطأ أثناء تشغيل PowerShell لتصدير الـ PDF:', err.message);
  } finally {
    // تنظيف ملف الـ PowerShell المؤقت
    if (fs.existsSync(ps1ScriptPath)) {
      fs.unlinkSync(ps1ScriptPath);
    }
  }
}

main().catch(console.error);
