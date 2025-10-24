// api/data.js
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
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(500).json({ error: `Excel not found at ${EXCEL_PATH}` });
    }
    const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true });
    const { sheetName, sheet } = getSheet(wb, req.query.sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    res.json({ sheet: sheetName, count: rows.length, rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
};
