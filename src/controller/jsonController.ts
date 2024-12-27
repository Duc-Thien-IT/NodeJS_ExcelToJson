import * as xlsx from 'xlsx';
import { Request, Response } from 'express';
import * as fs from 'fs';
import { ExcelToJSONConfig } from 'types';
import { excelToJson, jsonToExcel } from 'lib';
import { SheetParser } from 'lib';

// Hàm chuyển từ JSON sang Excel
export const convertJsonToExcel = async (req: Request, res: Response): Promise<void> => {
  try {
    // Nhận dữ liệu JSON từ body request
    const jsonData = req.body;
    const outputFile = req.body.outputFile || 'output.xlsx';

    if (!jsonData) {
      res.status(400).send({ message: 'Cần có dữ liệu JSON để chuyển đổi.' });
      return;
    }

    // Gọi hàm jsonToExcel từ thư viện
    jsonToExcel(jsonData, outputFile);

    // Đọc file đã tạo và trả về cho client
    const file = fs.createReadStream(outputFile);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${outputFile}"`);
    file.pipe(res);

  } catch (error) {
    console.error(error);
    res.status(500).send({ message: 'Đã xảy ra lỗi khi chuyển đổi từ JSON sang Excel.' });
  }
};
