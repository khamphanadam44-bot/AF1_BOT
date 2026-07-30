/**
 * ReconcileExcelReader
 * ------------------------------------------------------------------
 * แทนที่ reconcile-record-reader.ts และ report-data-reader.ts (2 ไฟล์เดิมที่มีเนื้อหา
 * ซ้ำกันทุกตัวอักษร) ด้วย class เดียว รับผิดชอบ "อ่าน Excel -> ReconcileRecord[]"
 * ทั้งฝั่ง AF1 Report และ Test Data (ต่างกันแค่ header row number)
 * ------------------------------------------------------------------
 */
import ExcelJS from "exceljs";
import { getCellText, normalizeValue } from "../../validators/shared/excel-cell.util";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import { ReconcileRecord } from "./record";

export interface ReconcileSheetData {
  headers: string[];
  records: ReconcileRecord[];
}

export class ReconcileExcelReader {
  /** อ่าน worksheet ที่เปิดอยู่แล้ว (ไม่เปิดไฟล์ซ้ำ) ตั้งแต่แถว headerRowNumber */
  parseWorksheet(
    worksheet: ExcelJS.Worksheet,
    headerRowNumber: number,
  ): ReconcileSheetData {
    const headerRow = worksheet.getRow(headerRowNumber);
    const maxColumn = Math.max(
      worksheet.columnCount,
      headerRow.cellCount,
      headerRow.actualCellCount,
    );

    const headers: string[] = [];
    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      headers[columnNumber - 1] = getCellText(headerRow.getCell(columnNumber));
    }

    const records: ReconcileRecord[] = [];
    const firstDataRowNumber = headerRowNumber + 1;

    for (
      let rowNumber = firstDataRowNumber;
      rowNumber <= worksheet.rowCount;
      rowNumber += 1
    ) {
      const row = worksheet.getRow(rowNumber);
      const values: Record<string, string> = {};
      let hasAnyValue = false;

      headers.forEach((header, index) => {
        if (!header) {
          return;
        }
        const cellText = normalizeValue(getCellText(row.getCell(index + 1)));
        values[header] = cellText; // canonicalHeader ทำเองใน ReconcileRecord.get()
        if (cellText !== "") {
          hasAnyValue = true;
        }
      });

      if (!hasAnyValue) {
        continue;
      }

      records.push(new ReconcileRecord(rowNumber, this.toCanonicalMap(values)));
    }

    return { headers, records };
  }

  /** อ่านไฟล์ Excel จาก path (เปิดไฟล์เอง) แล้วแปลงเป็น ReconcileSheetData — ใช้กับ Test Data */
  async readFile(
    filePath: string,
    headerRowNumber: number,
  ): Promise<ReconcileSheetData> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error(`Worksheet not found: ${filePath}`);
    }

    return this.parseWorksheet(worksheet, headerRowNumber);
  }

  /** normalize key ของ map ด้วย canonicalHeader ก่อนส่งเข้า ReconcileRecord */
  private toCanonicalMap(
    values: Record<string, string>,
  ): Record<string, string> {
    const canonical: Record<string, string> = {};
    Object.entries(values).forEach(([header, value]) => {
      canonical[canonicalHeader(header)] = value;
    });
    return canonical;
  }
}
