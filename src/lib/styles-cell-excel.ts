import { WorkBook, utils, WorkSheet } from 'xlsx';

interface CellStyle {
    font?: {
      bold?: boolean;
      color?: string;
    };
    fill?: {
      type: string;
      pattern?: string;
      fgColor?: string;
      bgColor?: string;
    };
    border?: {
      top?: { style: string; color: string };
      bottom?: { style: string; color: string };
      left?: { style: string; color: string };
      right?: { style: string; color: string };
    };
    alignment?: {
      vertical?: string;
      horizontal?: string;
      wrapText?: boolean;
    };
}
  
function applyStyles(ws: WorkSheet, range: string, style: CellStyle) {
    const rangeRef = utils.decode_range(range);
    
    for (let row = rangeRef.s.r; row <= rangeRef.e.r; row++) {
      for (let col = rangeRef.s.c; col <= rangeRef.e.c; col++) {
        const cellRef = utils.encode_cell({ r: row, c: col });
        if (!ws[cellRef]) ws[cellRef] = { v: '', t: 's' };
        
        ws[cellRef].s = {
          ...ws[cellRef].s,
          ...style,
        };
      }
    }
}
  
function styleWorksheet(ws: WorkSheet, headerRange: string) {
    // Style cho header
    const headerStyle: CellStyle = {
      font: { bold: true, color: '000000' },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: '99FFFF'
      },
      border: {
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'thin', color: '000000' },
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' }
      },
      alignment: {
        vertical: 'center',
        horizontal: 'center',
        wrapText: true
      }
    };
  
    applyStyles(ws, headerRange, headerStyle);
    
    // Style cho data cells
    const dataStyle: CellStyle = {
      border: {
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'thin', color: '000000' },
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' }
      },
      alignment: {
        vertical: 'center',
        horizontal: 'center',
        wrapText: true
      }
    };
    
    // Tính range cho data cells
    const range = utils.decode_range(ws['!ref'] || 'A1');
    const dataRange = utils.encode_range({
      s: { r: 2, c: 0 },  // Bắt đầu từ dòng 3 (sau header)
      e: { r: range.e.r, c: range.e.c }
    });
    
    applyStyles(ws, dataRange, dataStyle);
}

export const StyleWorksheet = styleWorksheet;