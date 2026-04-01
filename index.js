const ExcelJS = require('exceljs');
const path = require('path');
const { diffWords } = require('diff');
const config = require('./config.json');

const filePath = path.join(__dirname, config.fileName);

function toPlainString(value) {
  if (value == null) return '';

  // Primitive
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  // Rich text
  if (Array.isArray(value.richText)) {
    return value.richText
      .map(part => toPlainString(part.text))
      .join(''); // <-- keep empty join, do NOT add spaces
  }

  // Wrapped text
  if (typeof value.text !== 'undefined') {
    return toPlainString(value.text);
  }

  // Hyperlink
  if (value.hyperlink) {
    return toPlainString(value.text || value.hyperlink);
  }

  // Formula result
  if (typeof value.result !== 'undefined') {
    return toPlainString(value.result);
  }

  // Date
  if (value instanceof Date) {
    return value.toISOString();
  }

  return String(value);
}

// Extract a map of Type -> columnValue from a worksheet given column header names
function extractMap(worksheet, typeCol, valueCol, filterCol = null) {
  const map = new Map();
  let typeIdx = -1;
  let valueIdx = -1;
  let filterIdx = -1;

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    const header = String(cell.value).trim();
    if (header === typeCol) typeIdx = colNumber;
    if (header === valueCol) valueIdx = colNumber;
    if (filterCol && header === filterCol) filterIdx = colNumber;
  });

  if (typeIdx === -1 || valueIdx === -1) {
    throw new Error(`Could not find columns "${typeCol}" / "${valueCol}" in sheet "${worksheet.name}"`);
  }

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    const type = row.getCell(typeIdx).value;
    const value = row.getCell(valueIdx).value;
    const filter = filterIdx !== -1 ? row.getCell(filterIdx).value : null;

    if(filterIdx !== -1 && filter !== config.columns.filterValue) {
      return; // skip rows that don't match filter
    }

    if (type != null) {
      map.set(String(type).trim(), {
        original: value,
        plain: toPlainString(value).trim()
      });
    }
  });

  return map;
}

// Build a human-readable diff summary: lists added and removed words
function buildDiffSummary(oldText, newText) {
  const changes = diffWords(oldText, newText);
  const added = [];
  const removed = [];

  for (const part of changes) {
    if (part.added) added.push(part.value.trim());
    else if (part.removed) removed.push(part.value.trim());
  }

  const parts = [];
  if (removed.length) parts.push(`Removed: ${removed.join(', ')}`);
  if (added.length) parts.push(`Added: ${added.join(', ')}`);
  return parts.join(' | ');
}

async function processExcelFile() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const oldSheet = workbook.getWorksheet(config.sheets.old);
  const newSheet = workbook.getWorksheet(config.sheets.new);
  const triggersSheet = workbook.getWorksheet(config.sheets.triggers);

  if (!oldSheet || !newSheet || !triggersSheet) {
    throw new Error(`Workbook must contain sheets named "${config.sheets.old}", "${config.sheets.new}", and "${config.sheets.triggers}"`);
  }

  const oldMap = extractMap(oldSheet, config.columns.typeCol, config.columns.contentCol, config.columns.filterCol);
  const newMap = extractMap(newSheet, config.columns.typeCol, config.columns.contentCol, config.columns.filterCol);
  const triggerMap = extractMap(triggersSheet, config.columns.typeCol, config.columns.triggerCol);

  // Types ordered by 1st tab (Old), then any extras from 2nd tab (New) in their original order
  const allTypes = [
    ...oldMap.keys(),
    ...[...newMap.keys()].filter((k) => !oldMap.has(k)),
  ];

  // Remove existing Result sheet if present, then add a fresh one
  const existingResult = workbook.getWorksheet(config.sheets.result);
  if (existingResult) workbook.removeWorksheet(existingResult.id);

  const resultSheet = workbook.addWorksheet(config.sheets.result);
  const rc = config.resultColumns;
  resultSheet.addRow([rc.type, rc.oldContent, rc.newContent, rc.trigger, rc.modified, rc.differences]);

  // Style header row
  const headerRow = resultSheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.commit();

  for (const type of allTypes) {
    const oldContent = oldMap.get(type) ?? {};
    const newContent = newMap.get(type) ?? {};
    const trigger = triggerMap.get(type)?.original ?? '';
    const modified = oldContent.plain !== newContent.plain;
    const differences = modified ? buildDiffSummary(oldContent.plain ?? '', newContent.plain ?? '') : '';

    resultSheet.addRow([type, oldContent.original, newContent.original, trigger, modified ? 'Yes' : 'No', differences]);
  }

  // Auto-fit column widths (approximate)
  resultSheet.columns.forEach((col) => {
    let maxLen = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const len = cell.value ? String(cell.value).length : 0;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 2, 80);
  });
  resultSheet.getColumn(2).alignment = { wrapText: true };
  resultSheet.getColumn(3).alignment = { wrapText: true };

  await workbook.xlsx.writeFile(filePath);
  console.log(`Done. Result tab written to ${filePath}`);
}

processExcelFile().catch((err) => {
  console.error('Error:', err.message);
  process.exit(1);
});
