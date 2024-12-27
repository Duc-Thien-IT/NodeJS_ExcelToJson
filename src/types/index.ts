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
	sheetStubs?: boolean; // Biến này xác định xem có bao gồm các ô trống (empty cells) trong đầu ra JSON hay không. Nếu giá trị là true, các ô trống sẽ được bao gồm.

	/**
	 * Mapping of column letters to key names.
	 */
	columnToKey: Record<string, string>; //Là một đối tượng ánh xạ (mapping) từ các ký tự cột trong Excel (A, B, C, ...) sang các tên khóa (key names) trong đầu ra JSON. Ví dụ: { "A": "name", "B": "age" }.

	/**
	 * Mapping of cell addresses to key names.
	 */
	cellToKey?: Record<string, string>; //Tương tự như columnToKey, nhưng ánh xạ từ địa chỉ ô cụ thể (như A1, B2, ...) sang tên khóa. Đây là tùy chọn và không bắt buộc phải có.

	/**
	 * Range of cells to include in the output.
	 */
	range?: string; //Xác định phạm vi các ô cần trích xuất dữ liệu. Ví dụ: "A1:C10" để trích xuất dữ liệu từ ô A1 đến ô C10

	/**
	 * Configuration for the header row(s).
	 */
	header?: { //Cấu hình cho hàng tiêu đề (header rows).
		/**
		 * Number of header rows.
		 */
		rows?: number; //Số lượng hàng tiêu đề. Ví dụ, nếu bạn có một hàng tiêu đề, giá trị có thể là 1.

		/**
		 * Mapping of row numbers to key names.
		 */
		rowToKeys?: string; //Ánh xạ từ số hàng tiêu đề sang tên khóa (key names).
	};

	/**
	 * Configuration for the data rows.
	 */
	data: { //Cấu hình cho các hàng dữ liệu.
		/**
		 * Starting row number for data extraction.
		 */
		startRow: number; //Số hàng bắt đầu trích xuất dữ liệu.

		/**
		 * Ending row number for data extraction.
		 */
		endRow?: number; //Số hàng kết thúc trích xuất dữ liệu (tuỳ chọn).
	};

	/**
	 * Path to the source Excel file.
	 */
	sourceFile?: string; // Đường dẫn tới file Excel nguồn. Đây là đường dẫn tới file Excel mà bạn muốn chuyển đổi.

	/**
	 * Source data, if not using a file.
	 */
	source?: any; //Dữ liệu nguồn, nếu bạn không sử dụng file. Đây có thể là buffer hoặc một nguồn dữ liệu khác.

	/**
	 * List of sheets to process.
	 */
	sheets: ( //Danh sách các sheet cần xử lý. Có thể là tên sheet dưới dạng chuỗi hoặc đối tượng chứa tên sheet và số lượng sheet cần lấy.
		| string
		| {
				/**
				 * Name of the sheet.
				 */
				name: string; //Tên của sheet.

				/**
				 * Number of sheets to get.
				 */
				numberOfSheetsToGet?: number; //Số lượng sheet cần lấy.
		  }
	)[];

	/**
	 * List of required columns.
	 */
	requiredColumn?: string[]; //Danh sách các cột bắt buộc. Nếu các cột này không có dữ liệu, có thể phát sinh lỗi hoặc cảnh báo.



	/**
	 * Default values for missing data.
	 */
	defVal?: Record<string, any>; //Giá trị mặc định cho các dữ liệu bị thiếu. Đây là một đối tượng ánh xạ từ tên khóa sang giá trị mặc định.

	/**
	 * Data to append to the output.
	 */
	appendData?: Record<string, any>; //Dữ liệu cần thêm vào đầu ra JSON. Đây là một đối tượng chứa các dữ liệu bổ sung sẽ được thêm vào đầu ra JSON.
	includeMergeCells?: boolean;  // Tùy chọn để merge cells
}
