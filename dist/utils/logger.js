"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.logger = void 0;
const cli_color_1 = __importDefault(require("cli-color"));
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const logPath = path_1.default.join(__dirname, '../../logs/app.log');
const logToFile = (message) => {
    const logMessage = `${new Date().toISOString()} - ${message}\n`;
    fs_1.default.appendFileSync(logPath, logMessage);
};
exports.logger = {
    info: (message) => {
        const formattedMessage = cli_color_1.default.blueBright(message);
        console.log(formattedMessage);
        logToFile(`INFO: ${message}`);
    },
    error: (message) => {
        const formattedMessage = cli_color_1.default.redBright(message);
        console.error(formattedMessage);
        logToFile(`ERROR: ${message}`);
    },
    warn: (message) => {
        const formattedMessage = cli_color_1.default.yellowBright(message);
        console.warn(formattedMessage);
        logToFile(`WARN: ${message}`);
    },
};
