import { excelToJson } from "../lib/excel-to-json";
import { SheetTemplateProvider } from "../providers/dynamic-flow/sheet-template";
import { serverLog } from "./helpers";
import clc from "cli-color";

export class SheetsProvider {
  private readonly sheetTemplateProvider: SheetTemplateProvider = new SheetTemplateProvider();
  
  constructor() {}

  async excelToJson(templateId: string, sourceFile: string | Buffer): Promise<any> {
    try {
      const config = await this.sheetTemplateProvider.getOne({
        where: { id: templateId },
      });

      if (!config) {
        throw new Error("Template not found");
      }

      const result = await excelToJson(JSON.stringify(config.import), sourceFile);
      return result;
    } catch (error: any) {
      serverLog(`${clc.redBright("Error:")} ${error.message}`);
      throw new Error(error.message);
    }
  }
}
