"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.excelToJson = void 0;
const xlsx_1 = require("xlsx");
const _1 = require(".");
function convertExcelToJson(config, sourceFile) {
    var _a;
    const _config = typeof config === "string" ? JSON.parse(config) : config;
    _config.sourceFile = _config.sourceFile || (typeof sourceFile === "string" ? sourceFile : undefined);
    _config.source = _config.source || (Buffer.isBuffer(sourceFile) ? sourceFile : undefined);
    if (!(_config.sourceFile || _config.source)) {
        throw new Error(":: 'sourceFile' or 'source' required for _config :: ");
    }
    const workbook = _config.source
        ? (0, xlsx_1.read)(_config.source, { sheetStubs: true, cellDates: true })
        : (0, xlsx_1.readFile)(_config.sourceFile, { sheetStubs: true, cellDates: true });
    const sheetsToGet = Array.isArray(_config.sheets)
        ? _config.sheets
        : Object.keys(workbook.Sheets).slice(0, (_a = _config.sheets) === null || _a === void 0 ? void 0 : _a["numberOfSheetsToGet"]);
    let parsedData = {};
    if (Array.isArray(sheetsToGet) && sheetsToGet.length > 1) {
        sheetsToGet.forEach((sheet) => {
            const sheetConfig = typeof sheet === "string" ? { name: sheet } : sheet;
            const sheetParser = new _1.SheetParser(sheetConfig, workbook, _config);
            parsedData[sheetConfig.name] = sheetParser.parseSheet();
        });
    }
    else {
        const sheetConfig = typeof sheetsToGet[0] === "string" ? { name: sheetsToGet[0] } : sheetsToGet[0];
        const sheetParser = new _1.SheetParser(sheetConfig, workbook, _config);
        parsedData = sheetParser.parseSheet();
    }
    return parsedData;
}
exports.excelToJson = convertExcelToJson;
