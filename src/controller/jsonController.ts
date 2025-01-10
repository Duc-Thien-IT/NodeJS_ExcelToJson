import * as fs from 'fs';
import { Request, Response } from 'express';
import { JsonToExcel } from '../lib/json-to-excel';

/**
 * API chuyển từ JSON sang Excel
 */
export const convertJsonToExcel = (req: Request, res: Response) => {
  try {
    const jsonData = req.body;
    const outputFile = req.body.outputFile || "output.xlsx";

    // Gọi hàm JsonToExcel thay vì convertJsonToExcel
    JsonToExcel(jsonData, outputFile);

    res.download(outputFile);
  } catch (err: any) {
    res.status(500).send({ error: err.message });
  }
};