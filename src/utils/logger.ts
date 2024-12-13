import clc from 'cli-color';
import fs from 'fs';
import path from 'path';

const logPath = path.join(__dirname, '../../logs/app.log');

const logToFile = (message: string) => {
  const logMessage = `${new Date().toISOString()} - ${message}\n`;
  fs.appendFileSync(logPath, logMessage);
};

export const logger = {
  info: (message: string) => {
    const formattedMessage = clc.blueBright(message);
    console.log(formattedMessage);
    logToFile(`INFO: ${message}`);
  },
  error: (message: string) => {
    const formattedMessage = clc.redBright(message);
    console.error(formattedMessage);
    logToFile(`ERROR: ${message}`);
  },
  warn: (message: string) => {
    const formattedMessage = clc.yellowBright(message);
    console.warn(formattedMessage);
    logToFile(`WARN: ${message}`);
  },
};
