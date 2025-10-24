// api/sheets.js
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

const EXCEL_PATH = process.env.EXCEL_PATH || path.join(process.cwd(), 'data', 'data.xlsx');

module.exports = (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      return res.status(500).json({ error: `Excel not found at ${EXCEL_PATH}` });
    }
    const wb = XLSX.readFile(EXCEL_PATH, { cellDates: true });
    res.json({ sheets: wb.SheetNames });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
};
