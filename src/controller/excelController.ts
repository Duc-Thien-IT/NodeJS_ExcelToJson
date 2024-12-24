import { Request, Response } from 'express';
import * as xlsx from 'xlsx';
import * as fs from 'fs';
import columns from 'cli-color/columns';
import { ExcelToJSONConfig } from 'types';
import { excelToJson } from '../lib';
//import * as printer from 'printer';

const processSheet = (sheet: xlsx.WorkSheet, sheetName: string) => {
  const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  // Kiểm tra nếu dữ liệu rỗng
  if (jsonData.length === 0) {
    throw new Error(`Sheet "${sheetName}" is empty`);
  }

  const headers = jsonData[0] as string[];

  // Kiểm tra cột yêu cầu
  const requiredColumns = headers.filter(header => header !== undefined && header !== '');

  if (requiredColumns.length === 0) {
    throw new Error(`Required column is empty in sheet "${sheetName}"`);
  }

  return jsonData;
};

//Test
export const convertExcelToJson = async (req: Request, res: Response):Promise<void> => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }

  const workbook = xlsx.readFile(req.file.path);
    const firstSheetName = workbook.SheetNames[1];
  
  const config: ExcelToJSONConfig = {
    sheetStubs: true,
    columnToKey: {
      A: "STT",
      B: "Fullname",
      C: "Test",
      D: "Ngày nhận việc",
      E: "SỐ CMND",
      F: "Chức vụ",
      G: "Lương tháng",
      H: "Tổng công",
      I: "Ngày nghỉ chờ việc",
      J: "Công Ca Ngày",
      K: "Công Ca Đêm",
      L: "Công Nghỉ Lễ Được Hưởng Lương"
    },
    data: {
      startRow: 10
    },
    header: {
      rows: 4
    },
    sheets: [firstSheetName],
    requiredColumn: ['A'],
    defVal: {
        STT: 0,
        MSNV: "Unknown",
        "Họ Và Tên": "N/A",
        "Công Ca Ngày": 0,
        "Công Ca Đêm": 0,
        "Công Nghỉ Lễ Được Hưởng Lương": 0
    },
    appendData: {
      ExtraInfo: "Additional data" // Dữ liệu cần thêm vào đầu ra JSON
    }
  }

  const excelFile: Buffer = req.file.buffer;

  try {
    const jsonData = excelToJson(config, excelFile);
    res.status(200).json({ data: jsonData });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

//Xử lý chuyển dữ liệu file sheet đầu tiên sang dạng Json
export const handleFirstSheet = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);
    const firstSheetName = workbook.SheetNames[1];
    // const sheet = workbook.Sheets[firstSheetName];
    // const jsonData = processSheet(sheet, firstSheetName);

    const config: ExcelToJSONConfig = {
      sheetStubs: true,
      columnToKey: {
        A: "STT",
        B: "Fullname",
        C: "Test",
        D: "Ngày nhận việc",
        E: "SỐ CMND",
        F: "Chức vụ",
        G: "Lương tháng",
        H: "Tổng công",
        I: "Ngày nghỉ chờ việc",
        J: "Công Ca Ngày",
        K: "Công Ca Đêm",
        L: "Công Nghỉ Lễ Được Hưởng Lương"
      },
      data: {
        startRow: 10
      },
      header: {
        rows: 4
      },
      sheets: [firstSheetName],
      requiredColumn: ['A'],
      defVal: {
          STT: 0,
          MSNV: "Unknown",
          "Họ Và Tên": "N/A",
          "Công Ca Ngày": 0,
          "Công Ca Đêm": 0,
          "Công Nghỉ Lễ Được Hưởng Lương": 0
      },
      appendData: {
        ExtraInfo: "Additional data" // Dữ liệu cần thêm vào đầu ra JSON
      }
    }

    const jsonData = await excelToJson(config, req.file.path);
    const formattedData = jsonData.rows.map((item: any) => {
      const congTrongThang = {
        'Công Ca Ngày': item['Công trong tháng (ngày)']?.['Công Ca Ngày'],
        'Công Ca Đêm': item['Công trong tháng (ngày)']?.['Công Ca Đêm'],
        'Công Nghỉ Lễ Được Hưởng Lương': item['Công trong tháng (ngày)']?.['Công Nghỉ Lễ Được Hưởng Lương']
      };
      return {
        ...item,
        'Công trong tháng': congTrongThang
      };
    });

    res.status(200).json(formattedData);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

//Xử lý chuyển đổi tất cả file sheet
export const handleAllSheets = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);
    const jsonData = workbook.SheetNames.map(sheetName => {
      try {
        const sheet = workbook.Sheets[sheetName];
        const data = processSheet(sheet, sheetName);
        return { sheetName, data };
      } catch (error: any) {
        console.error(`Error processing sheet "${sheetName}": ${error.message}`);
        return { sheetName, error: error.message };
      }
    });

    fs.unlinkSync(req.file.path); 

    res.json(jsonData);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

//chuyển file sheet sang json theo tên file sheet
export const handleSheetByName = async (req: Request, res: Response): Promise<void> => {
  try {
    const sheetName = req.params.sheetName;

    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);

    if (!workbook.SheetNames.includes(sheetName)) {
      res.status(404).json({ error: `Sheet "${sheetName}" not found` });
      return;
    }

    const sheet = workbook.Sheets[sheetName];
    const jsonData = processSheet(sheet, sheetName);

    fs.unlinkSync(req.file.path);

    res.json(jsonData);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

//Đọc file sheet theo dữ liệu cột cụ thể
export const handleFirstSheetWithDataOnly = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.file) {
      res.status(400).json({ error: 'No file uploaded' });
      return;
    }

    const workbook = xlsx.readFile(req.file.path);
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    // Lấy tất cả các ô trong sheet
    const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1'); 
    const dataWithValues: Record<string, any> = {};

    // Duyệt từng ô
    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = xlsx.utils.encode_cell({ r: row, c: col });
        const cell = sheet[cellAddress];

        if (cell && cell.v !== undefined && cell.v !== null) {
          dataWithValues[cellAddress] = cell.v; 
        }
      }
    }

    fs.unlinkSync(req.file.path);
    res.json(dataWithValues);
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

//Json sang excel
export const handleJsonToExcel = async (req: Request, res: Response): Promise<void> => {
  try {
    if (!req.body) {
      res.status(400).json({ error: 'No JSON data provided' });
      return;
    }

    const jsonData = req.body;
    const sheetData: any[][] = [];

    // Chuyển đổi JSON data thành mảng 2 chiều
    Object.entries(jsonData).forEach(([cell, value]) => {
      const { r: row, c: col } = xlsx.utils.decode_cell(cell);

      // Đảm bảo mảng 2 chiều có kích thước phù hợp
      while (sheetData.length <= row) {
        sheetData.push([]);
      }
      while (sheetData[row].length <= col) {
        sheetData[row].push(undefined);
      }

      sheetData[row][col] = value;
    });

    const worksheet = xlsx.utils.aoa_to_sheet(sheetData);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const fileName = 'output.xlsx';
    const filePath = `./${fileName}`;

    xlsx.writeFile(workbook, filePath);

    res.download(filePath, fileName, (err) => {
      if (err) {
        console.error("File download failed:", err);
        res.status(500).json({ error: 'Error downloading the file' });
      } else {
        fs.unlinkSync(filePath); 
      }
    });
  } catch (error: any) {
    res.status(500).json({ error: error.message });
  }
};

// export const handleJsonToExcelAndPrint = async (req: Request, res: Response): Promise<void> => {
//   try{
//     if(!req.body.jsonData || !req.body.printerName){
//       res.status(400).json({ error: "No Json data or printer name provided" });
//       return;
//     }

//     const jsonData = req.body.jsonData;
//     const printerName = req.body.printerName;
//     const sheetData: any[][] = [];

//     Object.entries(jsonData).forEach(([cell, value]) => {
//       const { r: row, c: col } = xlsx.utils.decode_cell(cell);

//       while (sheetData.length <= row) {
//         sheetData.push([]);
//       }
//       while (sheetData[row].length <= col) {
//         sheetData[row].push(undefined);
//       }

//       sheetData[row][col] = value;
//     });

//     const worksheet = xlsx.utils.aoa_to_sheet(sheetData);
//     const workbook = xlsx.utils.book_new();
//     xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

//     const fileName = 'output.xlsx';
//     const filePath = `./${fileName}`;

//     xlsx.writeFile(workbook, filePath);

//     printer.printFile({
//       filename: filePath,
//       printer: printerName, 
//       success: (jobID) => {
//         console.log(`Sent to printer with ID: ${jobID}`);
//         fs.unlinkSync(filePath); 
//         res.status(200).json({ message: 'File sent to printer' });
//       },
//       error: (err) => {
//         console.error(`Failed to print: ${err}`);
//         res.status(500).json({ error: 'Error printing the file' });
//       }
//     });
//   }
//   catch(error: any){
//     res.status(500).json({ error: error.message });
//   }
// }