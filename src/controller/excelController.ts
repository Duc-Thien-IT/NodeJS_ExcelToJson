import { Request, Response, NextFunction } from "express";
import { SheetsProvider } from "../services/sheetsProvider";

const sheetsProvider = new SheetsProvider();

export const handleFirstSheet = async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await sheetsProvider.excelToJson("templateId", req.file?.path || "");
    res.json({
      message: "File processed successfully",
      data: result,
    });
  } catch (error) {
    next(error);
  }
};

export const handleAllSheets = async (req: Request, res: Response, next: NextFunction) => {
  try {
    const result = await sheetsProvider.excelToJson("templateId", req.file?.path || "");
    res.json({
      message: "File processed successfully",
      data: result,
    });
  } catch (error) {
    next(error);
  }
};
