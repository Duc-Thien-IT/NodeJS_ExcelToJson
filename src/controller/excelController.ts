import { Request, Response } from 'express';
import * as xlsx from 'xlsx';
import * as fs from 'fs';
import { ExcelToJSONConfig, SheetData } from 'types';
import { excelToJson } from '../lib';
import { SheetParser } from '../lib/sheet-parser';

const processSheet = (sheet: xlsx.WorkSheet, sheetName: string) => {
  const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  // Kiểm tra nếu dữ liệu rỗng
  if (jsonData.length === 0) {
    throw new Error(`Sheet "${sheetName}" is empty`);
  }

  const headers = jsonData[0] as string[];

  // Kiểm tra cột yêu cầu
  const requiredColumns = headers.filter(header => header !== undefined && header !== '');

  if (requiredColumns.length === 0) {
    throw new Error(`Required column is empty in sheet "${sheetName}"`);
  }

  return jsonData;
};

//Convert excel to json success
export const convertExcelToJson = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);
    const firstSheetName = workbook.SheetNames[1];
    const sheet = workbook.Sheets[firstSheetName];

    const config: ExcelToJSONConfig = {
      sheetStubs: true,
      columnToKey: {
        A: "STT",
        B: "MSNV",
        C: "Họ và Tên",
        D: "Ngày nhận việc",
        E: "SỐ CMND",
        F: "Chức vụ",
        G: "Lương tháng",
        H: "Tổng công",
        I: "Ngày nghỉ chờ việc",
        J: "Công Ca Ngày",
        K: "Công Ca Đêm",
        L: "Công Nghỉ Lễ Được Hưởng Lương"
      },
      data: {
        startRow: 10
      },
      header: {
        rows: 4
      },
      sheets: [firstSheetName],
      requiredColumn: ['A'],
      defVal: {
        STT: 0,
        MSNV: "Unknown",
        "Họ Và Tên": "N/A",
        "Công Ca Ngày": 0,
        "Công Ca Đêm": 0,
        "Công Nghỉ Lễ Được Hưởng Lương": 0
      },
      includeMergeCells: true
    };

    const sheetData = {name: firstSheetName, sheet: sheet};
    const parser = new SheetParser(sheetData, workbook, config);
    const jsonData = parser.parseSheet();

    // Process merge cells
    const merges = sheet['!merges'] || [];
    const mergeInfo: { [key: string]: string } = {};
    merges.forEach((merge) => {
      const start = xlsx.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
      const end = xlsx.utils.encode_cell({ r: merge.e.r, c: merge.e.c });
      mergeInfo[start] = `${start}:${end}`;
    });

    const formattedData = jsonData.rows.map((item: any, rowIndex: number) => {
      const rowKey = xlsx.utils.encode_cell({ r: rowIndex + 1, c: 0 }).replace(/\d+$/, '');
      const mergeRange = mergeInfo[rowKey];
      const congTrongThang = {
        'Công Ca Ngày': item['Công trong tháng']?.['Công Ca Ngày'],
        'Công Ca Đêm': item['Công trong tháng']?.['Công Ca Đêm'],
        'Công Nghỉ Lễ Được Hưởng Lương': item['Công trong tháng']?.['Công Nghỉ Lễ Được Hưởng Lương']
      };

      return {
        ...item,
        'Công trong tháng': congTrongThang,
        ...(mergeRange && { 'mergeRange': mergeRange })
      };
    });

    fs.unlinkSync(req.file.path);
    res.status(200).json(formattedData);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};