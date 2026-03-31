const ExcelJS = require('exceljs');
const path = require('path');
const { diffWords } = require('diff');
const config = require('./config.json');

const filePath = path.join(__dirname, config.fileName);

// Extract a map of Type -> columnValue from a worksheet given column header names
function extractMap(worksheet, typeCol, valueCol) {
  const map = new Map();
  let typeIdx = -1;
  let valueIdx = -1;

  worksheet.getRow(1).eachCell((cell, colNumber) => {
    const header = String(cell.value).trim();
    if (header === typeCol) typeIdx = colNumber;
    if (header === valueCol) valueIdx = colNumber;
  });

  if (typeIdx === -1 || valueIdx === -1) {
    throw new Error(`Could not find columns "${typeCol}" / "${valueCol}" in sheet "${worksheet.name}"`);
  }

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // skip header
    const type = row.getCell(typeIdx).value;
    const value = row.getCell(valueIdx).value;
    if (type != null) {
      map.set(String(type).trim(), value != null ? String(value).trim() : '');
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

  const oldMap = extractMap(oldSheet, config.columns.typeCol, config.columns.contentCol);
  const newMap = extractMap(newSheet, config.columns.typeCol, config.columns.contentCol);
  const triggerMap = extractMap(triggersSheet, config.columns.typeCol, config.columns.triggerCol);

  // Union of all types from Old and New
  const allTypes = [...new Set([...oldMap.keys(), ...newMap.keys()])].sort();

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
    const oldContent = oldMap.get(type) ?? '';
    const newContent = newMap.get(type) ?? '';
    const trigger = triggerMap.get(type) ?? '';
    const modified = oldContent !== newContent;
    const differences = modified ? buildDiffSummary(oldContent, newContent) : '';

    resultSheet.addRow([type, oldContent, newContent, trigger, modified ? 'Yes' : 'No', differences]);
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

  await workbook.xlsx.writeFile(filePath);
  console.log(`Done. Result tab written to ${filePath}`);
}

processExcelFile().catch((err) => {
  console.error('Error:', err.message);
  process.exit(1);
});
