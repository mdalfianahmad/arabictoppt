
import * as XLSX from 'xlsx';
import { ExcelRow, ColumnMapping, ColumnStyle } from '../types';

export const parseExcelFile = async (file: File): Promise<{ data: ExcelRow[], columns: string[] }> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        if (jsonData.length === 0) throw new Error("File is empty.");

        const headers = jsonData[0].map(h => String(h || '').trim()).filter(h => h !== '');
        const rows = XLSX.utils.sheet_to_json(worksheet) as ExcelRow[];

        resolve({ data: rows, columns: headers });
      } catch (err) {
        reject(err instanceof Error ? err : new Error("Failed to parse file."));
      }
    };
    reader.readAsArrayBuffer(file);
  });
};

const createStyle = (col = '', size = 24, x = 5, y = 10, w = 90, h = 15): ColumnStyle => ({
  column: col, fontSize: size, x, y, w, h
});

export const getDefaultMapping = (cols: string[]): ColumnMapping => ({
  box1: createStyle(cols[0] || '', 30, 5, 10, 90, 15),
  box2: createStyle(cols[1] || '', 30, 5, 30, 90, 15),
  box3: createStyle(cols[2] || '', 26, 5, 55, 90, 15),
  box4: createStyle('', 20, 5, 75, 43, 10),
  box5: createStyle('', 20, 52, 75, 43, 10)
});
