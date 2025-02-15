import * as XLSX from 'xlsx';

/*
  Diese Hilfsfunktion kann genutzt werden, um eine Excel-Datei asynchron einzulesen und
  als JSON-Daten zurückzugeben. (Optional – derzeit wird die Logik direkt in ExcelUpload.js verwendet)
*/
export const parseExcelFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        resolve(jsonData);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
};
