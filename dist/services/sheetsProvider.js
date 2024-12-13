"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.SheetsProvider = void 0;
const excel_to_json_1 = require("../lib/excel-to-json");
const sheet_template_1 = require("@providers/dynamic-flow/sheet-template");
const helper_1 = require("./helper");
const cli_color_1 = __importDefault(require("cli-color"));
class SheetsProvider {
    constructor() {
        this.sheetTemplateProvider = new sheet_template_1.SheetTemplateProvider();
    }
    excelToJson(templateId, sourceFile) {
        return __awaiter(this, void 0, void 0, function* () {
            try {
                const config = yield this.sheetTemplateProvider.getOne({
                    where: { id: templateId },
                });
                if (!config) {
                    throw new Error("Template not found");
                }
                const result = yield (0, excel_to_json_1.excelToJson)(JSON.stringify(config.import), sourceFile);
                return result;
            }
            catch (error) {
                (0, helper_1.serverLog)(`${cli_color_1.default.redBright("Error:")} ${error.message}`);
                throw new Error(error.message);
            }
        });
    }
}
exports.SheetsProvider = SheetsProvider;
