const ExcelJS = require('exceljs');
async function run() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./data.xlsx');
    console.log('Worksheets:', wb.worksheets.map(w => w.name));

    const mainSheet = wb.worksheets[0];
    console.log('First sheet rows:', mainSheet.rowCount);
    for (let r = 1; r <= Math.min(5, mainSheet.rowCount); r++) {
        const row = mainSheet.getRow(r);
        const vals = [];
        row.eachCell((cell, colId) => {
            vals.push(`[${colId}]: ${cell.value}`);
        });
        console.log(`Row ${r}:`, vals.join(' | '));
    }
}
run().catch(console.error);
