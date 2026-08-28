import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export interface TabularSheet {
  name: string;
  rows: unknown[][];
}

const excelJsValue = (value: ExcelJS.CellValue): unknown => {
  if (value == null) return '';
  if (value instanceof Date || typeof value !== 'object') return value;
  if ('result' in value) return value.result ?? '';
  if ('richText' in value) return value.richText.map(part => part.text).join('');
  if ('text' in value) return value.text;
  if ('error' in value) return value.error;
  return String(value);
};

/**
 * Read only the displayed table values needed by the importers.  Modern xlsx
 * files use ExcelJS because SheetJS 0.18 can crash while traversing vendor
 * OOXML metadata before it reaches the worksheet values.
 */
export async function readTabularWorkbook(file: File): Promise<TabularSheet[]> {
  const data = await file.arrayBuffer();
  if (/\.xlsx$/i.test(file.name)) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);
    return workbook.worksheets.map(worksheet => {
      const rows: unknown[][] = [];
      worksheet.eachRow({includeEmpty: true}, (row, rowNumber) => {
        const cells: unknown[] = [];
        row.eachCell({includeEmpty: true}, (cell, columnNumber) => {
          cells[columnNumber - 1] = excelJsValue(cell.value);
        });
        rows[rowNumber - 1] = cells;
      });
      return {name: worksheet.name, rows};
    });
  }

  const workbook = XLSX.read(data, {type: 'array', cellDates: false});
  return workbook.SheetNames.map(name => ({
    name,
    rows: XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[name], {header: 1, raw: true, defval: ''}),
  }));
}
