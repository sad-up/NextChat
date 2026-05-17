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
const MAX_ROWS_PER_SHEET = 100; // 每个 sheet 最多处理 100 行（进一步减少防止卡死）
const MAX_COLS_PER_ROW = 20; // 每行最多处理 20 列

async function delay(ms: number) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export function parseExcel(file: File): Promise<ExcelParseResult> {
  return new Promise(async (resolve, reject) => {
    if (file.size > MAX_FILE_SIZE) {
      reject(new Error(`文件过大 (超过 ${MAX_FILE_SIZE / 1024 / 1024}MB)，请选择较小的文件`));
      return;
    }

    await delay(10);

    const reader = new FileReader();
    
    reader.onload = async (e) => {
      try {
        await delay(10);
        
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: false,
          cellNF: false,
          cellHTML: false,
          WTF: false,
        });
        
        await delay(10);
        
        const sheets: ExcelSheet[] = [];
        let rawText = "";
        
        for (let i = 0; i < workbook.SheetNames.length; i++) {
          const sheetName = workbook.SheetNames[i];
          await delay(10);
          
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '',
          }) as string[][];
          
          const limitedData = jsonData.slice(0, MAX_ROWS_PER_SHEET).map(row => 
            row.slice(0, MAX_COLS_PER_ROW).map(cell => {
              const str = String(cell ?? '');
              return str.length > 200 ? str.substring(0, 200) + '...' : str;
            })
          );
          
          sheets.push({
            name: sheetName,
            data: limitedData
          });
          
          rawText += `【Sheet: ${sheetName}】\n`;
          limitedData.forEach((row) => {
            rawText += row.join('\t') + '\n';
          });
          rawText += '\n';
        }
        
        await delay(10);
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
