import * as xlsx from 'xlsx';
import { Request, Response } from 'express';
import * as fs from 'fs';
import { ExcelToJSONConfig } from 'types';
import { excelToJson } from 'lib';
import { SheetParser } from 'lib';

export const convertJsonToExcel = async (req: Request, res: Response): Promise<void> => {
  
};