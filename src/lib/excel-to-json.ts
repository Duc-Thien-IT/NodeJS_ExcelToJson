import { read, readFile } from "xlsx";
import { SheetParser } from ".";
import type { ExcelToJSONConfig, SheetData } from "../types";
import type { WorkBook } from "xlsx";
import * as xlsx from 'xlsx';

function convertExcelToJson(config: ExcelToJSONConfig | string, sourceFile: string | Buffer): any {
	const _config: ExcelToJSONConfig = typeof config === "string" ? JSON.parse(config) : config;
	_config.sourceFile = _config.sourceFile || (typeof sourceFile === "string" ? sourceFile : undefined);
	_config.source = _config.source || (Buffer.isBuffer(sourceFile) ? sourceFile : undefined);

	if (!(_config.sourceFile || _config.source)) {
		throw new Error(":: 'sourceFile' or 'source' required for _config :: ");
	}

	const workbook: WorkBook = _config.source
		? read(_config.source, { sheetStubs: true, cellDates: true })
		: readFile(_config.sourceFile as string, { sheetStubs: true, cellDates: true });

	const sheetsToGet: (string | SheetData)[] = Array.isArray(_config.sheets)
		? _config.sheets
		: Object.keys(workbook.Sheets).slice(0, _config.sheets?.["numberOfSheetsToGet"]);

	let parsedData: { [key: string]: any } = {};

	if (Array.isArray(sheetsToGet) && sheetsToGet.length > 1) {
		sheetsToGet.forEach((sheet) => {
			const sheetConfig = typeof sheet === "string" ? { name: sheet } : sheet;
			const sheetParser = new SheetParser(sheetConfig, workbook, _config);
			parsedData[sheetConfig.name] = sheetParser.parseSheet();
		});
	} else {
		const sheetConfig = typeof sheetsToGet[0] === "string" ? { name: sheetsToGet[0] } : sheetsToGet[0];
		const sheetParser = new SheetParser(sheetConfig, workbook, _config);
		parsedData = sheetParser.parseSheet();
	}

	return parsedData;
}

export const jsonToExcel = (jsonData: any, config: ExcelToJSONConfig, outputPath: string): void => {
    const workbook = xlsx.utils.book_new();

    jsonData.forEach((sheetData: any, index: number) => {
        const sheetName = typeof config.sheets[index] === 'string' ? config.sheets[index] : (config.sheets[index] as { name: string }).name;
        const sheet = xlsx.utils.json_to_sheet(sheetData, {
            header: Object.values(config.columnToKey),
            skipHeader: config.header?.rows ?? 1,
        });

        xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
    });

    xlsx.writeFile(workbook, outputPath);
};

export const excelToJson = convertExcelToJson;
