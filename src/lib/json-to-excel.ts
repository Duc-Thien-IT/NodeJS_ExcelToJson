import { read, utils, writeFile, WorkBook, WorkSheet } from "xlsx";

interface CellAddress {
  r: number;
  c: number;
}

interface MergeCell {
  s: CellAddress;
  e: CellAddress;
}

interface HeaderInfo {
  headerRow1: any[];
  headerRow2: any[];
  merges: MergeCell[];
}

function createHeadersFromJson(jsonData: any[]): HeaderInfo {
  const headerRow1: any[] = [];
  const headerRow2: any[] = [];
  const merges: MergeCell[] = [];
  let currentCol = 0;

  // Hàm để xử lý các trường nested
  const processNestedField = (key: string, value: any, colIndex: number) => {
    if (typeof value === "object" && value !== null) {
      const subKeys = Object.keys(value);
      
      // Thêm title chính vào dòng 1
      headerRow1.push(key);
      // Thêm khoảng trống cho các cột con
      for (let i = 1; i < subKeys.length; i++) {
        headerRow1.push("");
      }

      // Thêm các title con vào dòng 2
      subKeys.forEach(subKey => {
        headerRow2.push(subKey);
      });

      // Tạo merge cell cho trường này
      if (subKeys.length > 1) {
        merges.push({
          s: { r: 0, c: colIndex },
          e: { r: 0, c: colIndex + subKeys.length - 1 }
        });
      }

      return subKeys.length;
    }
    
    // Xử lý trường thông thường
    headerRow1.push(key);
    headerRow2.push("");
    return 1;
  };

  // Duyệt qua tất cả các trường trong JSON
  Object.entries(jsonData[0]).forEach(([key, value]) => {
    const columnsAdded = processNestedField(key, value, currentCol);
    currentCol += columnsAdded;
  });

  return { headerRow1, headerRow2, merges };
}

function flattenRow(row: any): any[] {
  const flatRow: any[] = [];

  Object.entries(row).forEach(([key, value]) => {
    if (typeof value === "object" && value !== null && !Array.isArray(value)) {
      // Nếu là object lồng nhau, làm phẳng các giá trị của nó
      const nestedValues = Object.values(value);
      flatRow.push(...nestedValues);
    } else {
      // Nếu là giá trị thông thường
      flatRow.push(value);
    }
  });

  return flatRow;
}

function convertJsonToExcel(jsonData: any, outputFile: string) {
  const workbook = utils.book_new();

  Object.keys(jsonData).forEach((sheetName) => {
    const sheetData = jsonData[sheetName].data;
    
    // Get headers and merges
    const { headerRow1, headerRow2, merges } = createHeadersFromJson(sheetData);

    // Create worksheet data
    const rows = [
      headerRow1,
      headerRow2,
      ...sheetData.map((row: any) => flattenRow(row))
    ];

    // Create worksheet
    const ws: WorkSheet = utils.aoa_to_sheet(rows);

    // Apply merges
    ws["!merges"] = merges;

    // Set column widths
    const colWidths: { [key: string]: number } = {};
    headerRow1.forEach((header, idx) => {
      if (header) {
        colWidths[utils.encode_col(idx)] = Math.max(15, header.toString().length * 1.5);
      }
    });
    ws["!cols"] = Object.keys(colWidths).map(key => ({ wch: colWidths[key] }));

    // Add the worksheet to workbook
    utils.book_append_sheet(workbook, ws, sheetName);
  });

  // Write to file
  writeFile(workbook, outputFile);
}

export const JsonToExcel = convertJsonToExcel;