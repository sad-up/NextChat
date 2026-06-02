import * as XLSX from "xlsx";

export interface ExcelSheet {
  name: string;
  data: string[][];
}

export interface ExcelParseResult {
  sheets: ExcelSheet[];
  rawText: string;
}

const MAX_FILE_SIZE = 10 * 1024 * 1024;
const MAX_ROWS_PER_SHEET = 30;
const MAX_COLS_PER_ROW = 10;
const MAX_CELL_LENGTH = 500;

async function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function parseExcel(file: File): Promise<ExcelParseResult> {
  return new Promise(async (resolve, reject) => {
    if (file.size > MAX_FILE_SIZE) {
      reject(
        new Error(
          `File is too large. Maximum size is ${MAX_FILE_SIZE / 1024 / 1024}MB.`,
        ),
      );
      return;
    }

    await delay(50);

    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        await delay(50);

        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, {
          type: "array",
          cellDates: false,
          cellNF: false,
          cellHTML: false,
          cellText: true,
          WTF: false,
          dense: true,
          sheetStubs: false,
        });

        await delay(50);

        const sheets: ExcelSheet[] = [];
        let rawText = "";
        const sheetNames = workbook.SheetNames.slice(0, 2);

        for (let i = 0; i < sheetNames.length; i++) {
          const sheetName = sheetNames[i];
          await delay(30);

          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: "",
            raw: false,
          }) as string[][];

          const limitedData = jsonData.slice(0, MAX_ROWS_PER_SHEET).map((row) =>
            row.slice(0, MAX_COLS_PER_ROW).map((cell) => {
              const str = String(cell ?? "");
              return str.length > MAX_CELL_LENGTH
                ? str.substring(0, MAX_CELL_LENGTH) + "..."
                : str;
            }),
          );

          sheets.push({
            name: sheetName,
            data: limitedData,
          });

          rawText += `[Sheet: ${sheetName}]\n`;
          limitedData.forEach((row) => {
            rawText += row.join(" | ") + "\n";
          });
          rawText += "\n";
        }

        await delay(30);
        resolve({ sheets, rawText });
      } catch (error) {
        reject(new Error("Excel parse failed: " + (error as Error).message));
      }
    };

    reader.onerror = () => reject(new Error("File read failed"));
    reader.readAsArrayBuffer(file);
  });
}

export function isExcelFile(filename: string): boolean {
  const ext = filename.toLowerCase().split(".").pop();
  return ["xlsx", "xls", "csv"].includes(ext || "");
}

export function formatExcelContent(result: ExcelParseResult): string {
  let content = "## Excel file content\n\n";
  content += `Sheets: ${result.sheets.length}\n\n`;

  result.sheets.forEach((sheet, sheetIndex) => {
    content += `### ${sheetIndex + 1}. ${sheet.name}\n\n`;
    if (sheet.data.length > 0) {
      content += sheet.data.map((row) => row.join(" | ")).join("\n");
    }
    content += "\n\n";
  });

  return content;
}
