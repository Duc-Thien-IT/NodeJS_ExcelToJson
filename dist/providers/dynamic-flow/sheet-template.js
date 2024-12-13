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
exports.SheetTemplateProvider = void 0;
class SheetTemplateProvider {
    constructor() {
        this.templates = [
            {
                id: "templateId",
                import: {
                    sheets: ["Sheet1"],
                    columnToKey: { A: "name", B: "age", C: "email" },
                    data: { startRow: 2 },
                },
            },
        ];
    }
    getOne(query) {
        return __awaiter(this, void 0, void 0, function* () {
            const { id } = query.where;
            const template = this.templates.find((template) => template.id === id);
            return template || null;
        });
    }
}
exports.SheetTemplateProvider = SheetTemplateProvider;
