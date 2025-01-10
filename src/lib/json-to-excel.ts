import { read, readFile, utils, writeFile, WorkBook, WorkSheet } from "xlsx";
import { SheetParser } from "./sheet-parser";
import type { ExcelToJSONConfig, SheetData } from "../types";
import * as xlsx from 'xlsx';

//Xử lý các title
function createHeadersFromJson(jsonData: any[]): { headerRow1: any[]; headerRow2: any[] } {
    const headerRow1: any[] = [];
    const headerRow2: any[] = [];
  
    Object.keys(jsonData[0]).forEach((key) => {
      if (typeof jsonData[0][key] === "object" && jsonData[0][key] !== null) {
        headerRow1.push(key);
        const subKeys = Object.keys(jsonData[0][key]);
        headerRow2.push(...subKeys);
        for (let i = 1; i < subKeys.length; i++) {
          headerRow1.push("");
        }
      } else {
        headerRow1.push(key);
        headerRow2.push("");
      }
    });
  
    return { headerRow1, headerRow2 };
}
  
function generateDynamicMerges(headerRow1: any[]): any[] {
    const merges: any[] = [];
    let startCol = 0;
  
    for (let i = 0; i < headerRow1.length; i++) {
      if (headerRow1[i] !== "") {
        const endCol = i;
        if (startCol < endCol) {
          merges.push({
            s: { r: 0, c: startCol },
            e: { r: 0, c: endCol },
          });
        }
        startCol = i + 1;
      }
    }
  
    return merges;
}

function convertJsonToExcel(jsonData: any, outputFile: string) {
    const workbook = utils.book_new();
  
    Object.keys(jsonData).forEach((sheetName) => {
      const sheetData = jsonData[sheetName].data;
  
      // Tạo tiêu đề từ JSON
      const { headerRow1, headerRow2 } = createHeadersFromJson(sheetData);
  
      // Thêm tiêu đề và dữ liệu vào sheet
      const rows = [
        headerRow1,
        headerRow2,
        ...sheetData.map((row: any) => {
          const flatRow: any[] = [];
          Object.keys(row).forEach((key) => {
            if (typeof row[key] === "object" && row[key] !== null) {
              const subValues = Object.values(row[key]);
              flatRow.push(...subValues);
            } else {
              flatRow.push(row[key]);
            }
          });
          return flatRow;
        }),
      ];
  
      const ws: WorkSheet = utils.aoa_to_sheet(rows);
  
      // Áp dụng merge cells tự động
      ws["!merges"] = generateDynamicMerges(headerRow1);
  
      utils.book_append_sheet(workbook, ws, sheetName);
    });
  
    writeFile(workbook, outputFile);
}
  
export const JsonToExcel = convertJsonToExcel;