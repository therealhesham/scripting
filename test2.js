const ExcelJS = require('exceljs');
async function run() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('./data.xlsx');

    const ws = wb.getWorksheet('مقبل ناجي عوض العلوني');
    if (!ws) return console.log('not found');

    const maxCol = ws.columnCount;
    console.log('Max col:', maxCol);

    for (let c = 1; c <= Math.min(maxCol, 40); c++) {
        const colName = ws.getColumn(c).letter;
        let hasValue = false;
        let vals = [];
        for (let r = 1; r <= Math.min(ws.rowCount, 10); r++) {
            const v = ws.getRow(r).getCell(c).value;
            if (v) hasValue = true;
            vals.push(v ? v.toString().substring(0, 10) : '-');
        }
        console.log(`Col ${c} [${colName}] -> HasValue: ${hasValue} | First 10: ${vals.join(', ')}`);
    }
}
run().catch(console.error);
