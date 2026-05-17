import * as XLSX from 'xlsx';

export interface ExcelSheet {
  name: string;
  data: string[][];
}

export interface ExcelParseResult {
  sheets: ExcelSheet[];
  rawText: string;
}

export function parseExcel(file: File): Promise<ExcelParseResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        const sheets: ExcelSheet[] = [];
        let rawText = "";
        
        workbook.SheetNames.forEach((sheetName) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as string[][];
          
          sheets.push({
            name: sheetName,
            data: jsonData
          });
          
          rawText += `【Sheet: ${sheetName}】\n`;
          jsonData.forEach((row) => {
            rawText += row.join('\t') + '\n';
          });
          rawText += '\n';
        });
        
        resolve({ sheets, rawText });
      } catch (error) {
        reject(new Error('Excel 文件解析失败: ' + (error as Error).message));
      }
    };
    
    reader.onerror = () => reject(new Error('文件读取失败'));
    reader.readAsArrayBuffer(file);
  });
}

export function isExcelFile(filename: string): boolean {
  const ext = filename.toLowerCase().split('.').pop();
  return ['xlsx', 'xls', 'csv'].includes(ext || '');
}

export function formatExcelContent(result: ExcelParseResult): string {
  let content = '## 📊 Excel 文件内容\n\n';
  content += `共 ${result.sheets.length} 个工作表\n\n`;
  
  result.sheets.forEach((sheet, sheetIndex) => {
    content += `### ${sheetIndex + 1}. ${sheet.name}\n\n`;
    if (sheet.data.length > 0) {
      content += '| ' + sheet.data[0]?.join(' | ') + ' |\n';
      content += '| ' + sheet.data[0]?.map(() => '---').join(' | ') + ' |\n';
      sheet.data.slice(1).forEach((row) => {
        content += '| ' + row.join(' | ') + ' |\n';
      });
    }
    content += '\n';
  });
  
  return content;
}
