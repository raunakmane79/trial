// api/search.js
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const EXCEL_PATH = process.env.EXCEL_PATH || path.join(process.cwd(), 'data', 'data.xlsx');

function getSheet(wb, sheetArg) {
  let name = sheetArg;
  if (!name) name = wb.SheetNames[0];
  else if (!wb.SheetNames.includes(name)) {
    const idx = Number(name);
    if (!Number.isNaN(idx) && wb.SheetNames[idx]) name = wb.SheetNames[idx];
    else throw new Error(`Sheet "${sheetArg}" not found`);
  }
  return { sheetName: name, sheet: wb.Sheets[name] };
}

module.exports = (req, res) => {
  try {
    const q = (req.query.q || '').toLowerCase().trim();
    if (!q) return res.status(400).send('Missing search query q');

    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(500).json({ error: `Excel not found at ${EXCEL_PATH}` });
    }
    const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true });
    const { sheetName, sheet } = getSheet(wb, req.query.sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const filtered = rows.filter(row =>
      Object.values(row).some(v => String(v).toLowerCase().includes(q))
    );

    const outWb = XLSX.utils.book_new();
    const outWs = XLSX.utils.json_to_sheet(filtered);
    XLSX.utils.book_append_sheet(outWb, outWs, "Results");

    const buf = XLSX.write(outWb, { type: "buffer", bookType: "xlsx" });
    const safeSheet = String(sheetName).replace(/[^a-z0-9_-]+/gi, '-');
    const filename = `results_${safeSheet}_${q}.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buf);
  } catch (e) {
    console.error(e);
    res.status(500).send('Error searching Excel');
  }
};
