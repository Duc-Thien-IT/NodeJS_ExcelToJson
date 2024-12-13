export interface SheetData {
	name: string;
}

/**
 * Configuration interface for converting Excel data to JSON.
 */
export interface ExcelToJSONConfig {
	/**
	 * Whether to include empty cells in the output.
	 */
	sheetStubs?: boolean;

	/**
	 * Mapping of column letters to key names.
	 */
	columnToKey: Record<string, string>;

	/**
	 * Mapping of cell addresses to key names.
	 */
	cellToKey?: Record<string, string>;

	/**
	 * Range of cells to include in the output.
	 */
	range?: string;

	/**
	 * Configuration for the header row(s).
	 */
	header?: {
		/**
		 * Number of header rows.
		 */
		rows?: number;

		/**
		 * Mapping of row numbers to key names.
		 */
		rowToKeys?: string;
	};

	/**
	 * Configuration for the data rows.
	 */
	data: {
		/**
		 * Starting row number for data extraction.
		 */
		startRow: number;

		/**
		 * Ending row number for data extraction.
		 */
		endRow?: number;
	};

	/**
	 * Path to the source Excel file.
	 */
	sourceFile?: string;

	/**
	 * Source data, if not using a file.
	 */
	source?: any;

	/**
	 * List of sheets to process.
	 */
	sheets: (
		| string
		| {
				/**
				 * Name of the sheet.
				 */
				name: string;

				/**
				 * Number of sheets to get.
				 */
				numberOfSheetsToGet?: number;
		  }
	)[];

	/**
	 * List of required columns.
	 */
	requiredColumn?: string[];

	/**
	 * Default values for missing data.
	 */
	defVal?: Record<string, any>;

	/**
	 * Data to append to the output.
	 */
	appendData?: Record<string, any>;
}
