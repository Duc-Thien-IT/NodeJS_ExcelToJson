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
Object.defineProperty(exports, "__esModule", { value: true });
exports.handleAllSheets = exports.handleFirstSheet = void 0;
const sheetsProvider_1 = require("../services/sheetsProvider");
const sheetsProvider = new sheetsProvider_1.SheetsProvider();
const handleFirstSheet = (req, res, next) => __awaiter(void 0, void 0, void 0, function* () {
    var _a;
    try {
        const result = yield sheetsProvider.excelToJson("templateId", ((_a = req.file) === null || _a === void 0 ? void 0 : _a.path) || "");
        res.json({
            message: "File processed successfully",
            data: result,
        });
    }
    catch (error) {
        next(error);
    }
});
exports.handleFirstSheet = handleFirstSheet;
const handleAllSheets = (req, res, next) => __awaiter(void 0, void 0, void 0, function* () {
    var _a;
    try {
        const result = yield sheetsProvider.excelToJson("templateId", ((_a = req.file) === null || _a === void 0 ? void 0 : _a.path) || "");
        res.json({
            message: "File processed successfully",
            data: result,
        });
    }
    catch (error) {
        next(error);
    }
});
exports.handleAllSheets = handleAllSheets;
