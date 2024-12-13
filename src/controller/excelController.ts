import { Request, Response } from 'express';
import * as xlsx from 'xlsx';
import * as fs from 'fs';

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

export const handleFirstSheet = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const jsonData = processSheet(sheet, firstSheetName);

    fs.unlinkSync(req.file.path); 

    res.json(jsonData);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

export const handleAllSheets = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);
    const jsonData = workbook.SheetNames.map(sheetName => {
      try {
        const sheet = workbook.Sheets[sheetName];
        const data = processSheet(sheet, sheetName);
        return { sheetName, data };
      } catch (error: any) {
        console.error(`Error processing sheet "${sheetName}": ${error.message}`);
        return { sheetName, error: error.message };
      }
    });

    fs.unlinkSync(req.file.path); 

    res.json(jsonData);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};
