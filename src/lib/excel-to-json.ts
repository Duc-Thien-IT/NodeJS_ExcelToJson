import { read, readFile, utils, writeFile, WorkBook, WorkSheet } from "xlsx";
import { SheetParser } from ".";
import type { ExcelToJSONConfig, SheetData } from "../types";
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
			//parsedData[sheetConfig.name] = sheetParser.parseSheet();
			parsedData[sheetConfig.name] = {
				data: sheetParser.parseSheet(),
				merges: workbook.Sheets[sheetConfig.name]['!merges']
			};
		});
	} else {
		const sheetConfig = typeof sheetsToGet[0] === "string" ? { name: sheetsToGet[0] } : sheetsToGet[0];
		const sheetParser = new SheetParser(sheetConfig, workbook, _config);
		//parsedData = sheetParser.parseSheet();
		parsedData = {
			data: sheetParser.parseSheet(),
			merges: workbook.Sheets[sheetConfig.name]['!merges']
		};
	}

	return parsedData;
}

export const excelToJson = convertExcelToJson;

//Chuyển từ json sang excel
function convertJsonToExcel(jsonData: any, outputFile: string) {
	const workbook = utils.book_new();
	Object.keys(jsonData).forEach(sheetName => {
		const sheetData = jsonData[sheetName].data;
		const merges = jsonData[sheetName].merges;
		const ws: WorkSheet = utils.json_to_sheet(sheetData);

		ws['!merges'] = merges;

		utils.book_append_sheet(workbook, ws, sheetName);
	});
	writeFile(workbook, outputFile);
}

// Export hàm convertJsonToExcel
export const jsonToExcel = convertJsonToExcel;