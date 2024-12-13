"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.SheetParser = void 0;
const helper_1 = require("@services/helper");
const cli_color_1 = __importDefault(require("cli-color"));
class SheetParser {
    constructor(sheetData, workbook, config) {
        this.sheetData = sheetData;
        this.workbook = workbook;
        this.config = config;
    }
    parseSheet() {
        var _a, _b;
        (0, helper_1.serverLog)(cli_color_1.default.blueBright(`Parsing sheet: ${this.sheetData.name}`));
        const sheet = this.workbook.Sheets[this.sheetData.name];
        const { columnToKey = {}, cellToKey, range, header, data, requiredColumn, defVal = {} } = this.config;
        const headerRowToKeys = header === null || header === void 0 ? void 0 : header.rowToKeys;
        const dataStartRow = (_a = data === null || data === void 0 ? void 0 : data.startRow) !== null && _a !== void 0 ? _a : 1;
        const requiredColumns = requiredColumn !== null && requiredColumn !== void 0 ? requiredColumn : Object.keys(columnToKey)[0];
        defVal["*"] = (_b = defVal["*"]) !== null && _b !== void 0 ? _b : null;
        const strictRangeColumns = this.getStrictRangeColumns(range);
        const strictRangeRows = this.getStrictRangeRows(range);
        let rows = [];
        let extraData = {};
        let reading = true;
        let row = dataStartRow - 1;
        while (reading) {
            row++;
            if (this.isOutOfRange(row, strictRangeRows)) {
                reading = false;
                break;
            }
            for (let column in columnToKey) {
                const cell = `${column}${row}`;
                if (this.isColumnOutOfRange(column, strictRangeColumns))
                    continue;
                if (this.isRequiredColumnEmpty(sheet, column, row, cell, requiredColumns)) {
                    reading = false;
                    break;
                }
                if (cell === "!ref" || !this.isColumnKeyValid(columnToKey, column))
                    continue;
                const rowData = (rows[row - dataStartRow] = rows[row - dataStartRow] || {});
                const columnData = this.getColumnData(columnToKey, column, headerRowToKeys);
                const cellData = this.getCellData(sheet, cell, columnData, defVal);
                rowData[columnData] = cellData;
                if (this.config.appendData)
                    Object.assign(rowData, this.config.appendData);
            }
        }
        if (cellToKey) {
            extraData = this.getExtraData(sheet, cellToKey);
        }
        return { rows, extraData };
    }
    getStrictRangeColumns(range) {
        if (!range)
            return { from: null, to: null };
        return {
            from: this.getCellColumn(this.getRangeBegin(range)),
            to: this.getCellColumn(this.getRangeEnd(range)),
        };
    }
    getStrictRangeRows(range) {
        if (!range)
            return { from: null, to: null };
        return {
            from: this.getCellRow({ cell: this.getRangeBegin(range) }),
            to: this.getCellRow({ cell: this.getRangeEnd(range) }),
        };
    }
    isOutOfRange(row, strictRangeRows) {
        return (strictRangeRows.from !== null &&
            strictRangeRows.to !== null &&
            (row < strictRangeRows.from || row > strictRangeRows.to));
    }
    isColumnOutOfRange(column, strictRangeColumns) {
        return strictRangeColumns && (column < strictRangeColumns.from || column > strictRangeColumns.to);
    }
    isRequiredColumnEmpty(sheet, column, row, cell, requiredColumns) {
        const requiredColumnsArray = Array.isArray(requiredColumns) ? requiredColumns : [requiredColumns];
        if (requiredColumnsArray.length > 0 && requiredColumnsArray.includes(column) && !sheet[cell]) {
            // console.log(`🚀 Required column ${column} is empty at row ${row} and cell ${cell}: Parser will be stopped`);
            (0, helper_1.serverLog)(cli_color_1.default.redBright(`Required column ${column} is empty at row ${row} and cell ${cell}: Parser will be stopped`));
            return true;
        }
        return false;
    }
    isColumnKeyValid(columnToKey, column) {
        return columnToKey && (columnToKey[column] || columnToKey["*"]);
    }
    getColumnData(columnToKey, column, headerRowToKeys) {
        var _a, _b;
        return ((_b = (_a = columnToKey === null || columnToKey === void 0 ? void 0 : columnToKey[column]) !== null && _a !== void 0 ? _a : columnToKey === null || columnToKey === void 0 ? void 0 : columnToKey["*"]) !== null && _b !== void 0 ? _b : (headerRowToKeys ? `{{${column}${headerRowToKeys}}}` : column));
    }
    getCellData(sheet, cell, columnData, defVal) {
        var _a, _b, _c;
        if (((_a = sheet[cell]) === null || _a === void 0 ? void 0 : _a.v) === undefined && !(this.config.sheetStubs && ((_b = sheet[cell]) === null || _b === void 0 ? void 0 : _b.t) === "z")) {
            return (_c = defVal[columnData]) !== null && _c !== void 0 ? _c : defVal["*"];
        }
        return this.getSheetCellValue(sheet[cell]);
    }
    getExtraData(sheet, cellToKey) {
        let extraData = {};
        for (let cell in cellToKey) {
            const key = cellToKey[cell];
            if (key === "")
                continue;
            extraData[key] = this.getSheetCellValue(sheet[cell]);
        }
        return extraData;
    }
    getCellRow({ cell }) {
        return Number(cell.replace(/[A-z]/gi, ""));
    }
    getCellColumn(cell) {
        return cell.replace(/[0-9]/g, "").toUpperCase();
    }
    getRangeBegin(cell) {
        const match = cell.match(/^[^:]*/);
        return match ? match[0] : "";
    }
    getRangeEnd(cell) {
        const match = cell.match(/[^:]*$/);
        return match ? match[0] : "";
    }
    getSheetCellValue(sheetCell) {
        if (!sheetCell)
            return undefined;
        if (sheetCell.t === "z" && this.config.sheetStubs)
            return null;
        return sheetCell.t === "n" || sheetCell.t === "d"
            ? sheetCell.v
            : (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
    }
}
exports.SheetParser = SheetParser;
