import type { WorkBook } from "xlsx";
import type { ExcelToJSONConfig, SheetData } from "../types";
import { serverLog } from "../services/helpers";
import clc from "cli-color";

export class SheetParser {
	private sheetData: SheetData;
	private workbook: WorkBook;
	private config: ExcelToJSONConfig;

	constructor(sheetData: SheetData, workbook: WorkBook, config: ExcelToJSONConfig) {
		this.sheetData = sheetData;
		this.workbook = workbook;
		this.config = config;
	}

	public parseSheet(): any {
		serverLog(clc.blueBright(`Parsing sheet: ${this.sheetData.name}`));
		const sheet = this.workbook.Sheets[this.sheetData.name];
		const { columnToKey = {}, cellToKey, range, header, data, requiredColumn, defVal = {} } = this.config;
		const headerRowToKeys = header?.rowToKeys;
		const dataStartRow = data?.startRow ?? 1;
		const requiredColumns = requiredColumn ?? Object.keys(columnToKey)[0];
		defVal["*"] = defVal["*"] ?? null;

		const strictRangeColumns = this.getStrictRangeColumns(range);
		const strictRangeRows = this.getStrictRangeRows(range);

		let rows: any[] = [];
		let extraData: any = {};
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
				if (this.isColumnOutOfRange(column, strictRangeColumns)) continue;
				if (this.isRequiredColumnEmpty(sheet, column, row, cell, requiredColumns)) {
					reading = false;
					break;
				}
				if (cell === "!ref" || !this.isColumnKeyValid(columnToKey, column)) continue;

				const rowData = (rows[row - dataStartRow] = rows[row - dataStartRow] || {});
				const columnData = this.getColumnData(columnToKey, column, headerRowToKeys);
				const cellData = this.getCellData(sheet, cell, columnData, defVal);

				rowData[columnData] = cellData;
				if (this.config.appendData) Object.assign(rowData, this.config.appendData);
			}
		}

		if (cellToKey) {
			extraData = this.getExtraData(sheet, cellToKey);
		}

		return { rows, extraData };
	}

	private getStrictRangeColumns(range: string | undefined) {
		if (!range) return { from: null, to: null };
		return {
			from: this.getCellColumn(this.getRangeBegin(range)),
			to: this.getCellColumn(this.getRangeEnd(range)),
		};
	}

	private getStrictRangeRows(range: string | undefined) {
		if (!range) return { from: null, to: null };
		return {
			from: this.getCellRow({ cell: this.getRangeBegin(range) }),
			to: this.getCellRow({ cell: this.getRangeEnd(range) }),
		};
	}

	private isOutOfRange(row: number, strictRangeRows: { from: any; to: any }) {
		return (
			strictRangeRows.from !== null &&
			strictRangeRows.to !== null &&
			(row < strictRangeRows.from || row > strictRangeRows.to)
		);
	}

	private isColumnOutOfRange(column: string, strictRangeColumns: { from: any; to: any }) {
		return strictRangeColumns && (column < strictRangeColumns.from || column > strictRangeColumns.to);
	}

	private isRequiredColumnEmpty(
		sheet: any,
		column: string,
		row: number,
		cell: string,
		requiredColumns: string | string[],
	) {
		const requiredColumnsArray = Array.isArray(requiredColumns) ? requiredColumns : [requiredColumns];
		if (requiredColumnsArray.length > 0 && requiredColumnsArray.includes(column) && !sheet[cell]) {
			// console.log(`🚀 Required column ${column} is empty at row ${row} and cell ${cell}: Parser will be stopped`);
			serverLog(
				clc.redBright(`Required column ${column} is empty at row ${row} and cell ${cell}: Parser will be stopped`),
			);
			return true;
		}
		return false;
	}

	private isColumnKeyValid(columnToKey: any, column: string) {
		return columnToKey && (columnToKey[column] || columnToKey["*"]);
	}

	private getColumnData(columnToKey: any, column: string, headerRowToKeys: any) {
		return (
			columnToKey?.[column] ?? columnToKey?.["*"] ?? (headerRowToKeys ? `{{${column}${headerRowToKeys}}}` : column)
		);
	}

	private getCellData(sheet: any, cell: string, columnData: string, defVal: any) {
		if (sheet[cell]?.v === undefined && !(this.config.sheetStubs && sheet[cell]?.t === "z")) {
			return defVal[columnData] ?? defVal["*"];
		}
		return this.getSheetCellValue(sheet[cell]);
	}

	private getExtraData(sheet: any, cellToKey: any) {
		let extraData: any = {};
		for (let cell in cellToKey) {
			const key = cellToKey[cell];
			if (key === "") continue;
			extraData[key] = this.getSheetCellValue(sheet[cell]);
		}
		return extraData;
	}

	private getCellRow({ cell }: { cell: string }) {
		return Number(cell.replace(/[A-z]/gi, ""));
	}

	private getCellColumn(cell: string) {
		return cell.replace(/[0-9]/g, "").toUpperCase();
	}

	private getRangeBegin(cell: string) {
		const match = cell.match(/^[^:]*/);
		return match ? match[0] : "";
	}

	private getRangeEnd(cell: string) {
		const match = cell.match(/[^:]*$/);
		return match ? match[0] : "";
	}

	private getSheetCellValue(sheetCell: { t: string; v: any; w: string }) {
		if (!sheetCell) return undefined;
		if (sheetCell.t === "z" && this.config.sheetStubs) return null;
		return sheetCell.t === "n" || sheetCell.t === "d"
			? sheetCell.v
			: (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
	}
}
