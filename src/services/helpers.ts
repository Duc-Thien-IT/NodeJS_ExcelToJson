import clc from "cli-color";

export const serverLog = (message: string) => {
  console.log(clc.greenBright(`[SERVER LOG]: ${message}`));
};
