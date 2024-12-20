import { Request, Response } from 'express';
import * as fs from 'fs';
import json2xls from 'json-as-xlsx';
import { ExcelToJSONConfig } from '../types';
import { excelToJson, jsonToExcel } from '../lib';

export const handleJsonToExcel = async (req: Request, res: Response): Promise<void> => {
    try {
        if (!req.body.jsonData || !req.body.config) {
            res.status(400).json({ error: 'Missing jsonData or config' });
            return;
        }

        const jsonData = req.body.jsonData;
        const config: ExcelToJSONConfig = req.body.config;

        const outputFilePath = './output.xlsx';
        jsonToExcel(jsonData, config, outputFilePath);

        res.download(outputFilePath, 'output.xlsx', (err) => {
            if (err) {
                res.status(500).json({ error: err.message });
            }
            fs.unlinkSync(outputFilePath);
        });
    } catch (error: any) {
        res.status(500).json({ error: error.message });
    }
};