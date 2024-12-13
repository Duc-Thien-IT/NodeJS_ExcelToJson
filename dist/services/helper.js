"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.serverLog = void 0;
const cli_color_1 = __importDefault(require("cli-color"));
const serverLog = (message) => {
    console.log(cli_color_1.default.greenBright(`[SERVER LOG]: ${message}`));
};
exports.serverLog = serverLog;
