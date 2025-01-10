import { read, utils, writeFile, WorkBook, WorkSheet } from "xlsx";

interface CellAddress {
  r: number;
  c: number;
}

interface MergeCell {
  s: CellAddress;
  e: CellAddress;
}

function createHeadersFromJson(jsonData: any[]): { 
  headerRow1: any[]; 
  headerRow2: any[]; 
  merges: MergeCell[] 
} {
  const headerRow1: any[] = [];
  const headerRow2: any[] = [];
  const merges: MergeCell[] = [];
  let currentCol = 0;

  // Xử lý các trường thông thường
  Object.keys(jsonData[0]).forEach((key) => {
    if (key !== "Công trong tháng") {
      headerRow1.push(key);  // Dòng 1 chứa tất cả các title
      headerRow2.push("");   // Dòng 2 để trống cho các cột thông thường
      currentCol += 1;
    }
  });

  // Xử lý phần "Công trong tháng"
  const congTrongThang = jsonData[0]["Công trong tháng"];
  if (congTrongThang && typeof congTrongThang === "object") {
    const startCol = currentCol;
    const subKeys = Object.keys(congTrongThang);
    
    // Thêm title "Công trong tháng" vào dòng 1
    headerRow1.push("Công trong tháng");
    // Thêm khoảng trống cho các cột con
    for (let i = 1; i < subKeys.length; i++) {
      headerRow1.push("");
    }

    // Thêm các title con vào dòng 2, chỉ ở phần "Công trong tháng"
    subKeys.forEach(subKey => {
      headerRow2.push(subKey);
    });

    // Tạo merge cell cho "Công trong tháng"
    merges.push({
      s: { r: 0, c: startCol },
      e: { r: 0, c: startCol + subKeys.length - 1 }
    });
  }

  return { headerRow1, headerRow2, merges };
}

function flattenRow(row: any): any[] {
  const flatRow: any[] = [];
  
  // Xử lý các trường thông thường trước
  Object.keys(row).forEach((key) => {
    if (key !== "Công trong tháng") {
      flatRow.push(row[key]);
    }
  });
  
  // Sau đó xử lý "Công trong tháng"
  if (row["Công trong tháng"]) {
    const congValues = Object.values(row["Công trong tháng"]);
    flatRow.push(...congValues);
  }
  
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