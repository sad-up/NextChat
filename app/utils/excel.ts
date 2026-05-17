import * as XLSX from 'xlsx';

export interface ExcelSheet {
  name: string;
  data: string[][];
}

export interface ExcelParseResult {
  sheets: ExcelSheet[];
  rawText: string;
}

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
const MAX_ROWS_PER_SHEET = 30; // 每个 sheet 只处理 30 行（极致简化）
const MAX_COLS_PER_ROW = 10; // 每行只处理 10 列
const MAX_CELL_LENGTH = 100; // 每个单元格最多 100 字符

async function delay(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export function parseExcel(file: File): Promise<ExcelParseResult> {
  return new Promise(async (resolve, reject) => {
    if (file.size > MAX_FILE_SIZE) {
      reject(new Error(`文件过大 (超过 ${MAX_FILE_SIZE / 1024 / 1024}MB)，请选择较小的文件`));
      return;
    }

    await delay(50); // 给浏览器更多时间响应

    const reader = new FileReader();
    
    reader.onload = async (e) => {
      try {
        await delay(50);
        
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: false,
          cellNF: false,
          cellHTML: false,
          cellText: true,
          WTF: false,
          dense: true, // 更紧凑的格式
          sheetStubs: false,
        });
        
        await delay(50);
        
        const sheets: ExcelSheet[] = [];
        let rawText = "";
        
        // 只处理前2个Sheet，再多的话可能太卡
        const sheetNames = workbook.SheetNames.slice(0, 2);
        
        for (let i = 0; i < sheetNames.length; i++) {
          const sheetName = sheetNames[i];
          await delay(30);
          
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '',
            raw: false,
          }) as string[][];
          
          const limitedData = jsonData.slice(0, MAX_ROWS_PER_SHEET).map(row => 
            row.slice(0, MAX_COLS_PER_ROW).map(cell => {
              const str = String(cell ?? '');
              return str.length > MAX_CELL_LENGTH ? str.substring(0, MAX_CELL_LENGTH) + '...' : str;
            })
          );
          
          sheets.push({
            name: sheetName,
            data: limitedData
          });
          
          rawText += `【Sheet: ${sheetName}】\n`;
          limitedData.forEach((row) => {
            rawText += row.join(' | ') + '\n';
          });
          rawText += '\n';
        }
        
        await delay(30);
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
  let content = '## 📊 Excel 文件内容 (已简化处理)\n\n';
  content += `共 ${result.sheets.length} 个工作表\n\n`;
  
  result.sheets.forEach((sheet, sheetIndex) => {
    content += `### ${sheetIndex + 1}. ${sheet.name}\n\n`;
    if (sheet.data.length > 0) {
      content += sheet.data.map(row => row.join(' | ')).join('\n');
    }
    content += '\n\n';
  });
  
  return content;
}
