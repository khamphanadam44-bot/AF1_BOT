/**
 * ============================================================================
 * ไฟล์: row-mapper.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้เป็นตัวช่วยแปลง Excel 1 แถวให้กลายเป็น Object ที่อ่านง่าย
 *
 * ก่อนแปลง
 * --------
 * ข้อมูลใน Excel ถูกอ้างอิงด้วยหมายเลข Column เช่น Cell 1, Cell 2, Cell 3
 *
 * หลังแปลง
 * --------
 * สามารถเรียกข้อมูลด้วยชื่อ Header เช่น
 * rowData["Txn Date"]
 * rowData["Currency Id"]
 *
 * ตัวอย่าง
 * --------
 * Header: Txn Date | Currency | Amount
 * Data:   25/11/2025 | HKD | 100
 *
 * ผลลัพธ์:
 * {
 *   "Txn Date": "25/11/2025",
 *   "Currency": "HKD",
 *   "Amount": "100"
 * }
 *
 * ไฟล์นี้ใช้ร่วมกันทั้งตอนอ่าน Test Data และตอนอ่าน AF1 Report
 * ============================================================================
 */


/**
 * ส่วน import ด้านล่าง คือการนำเครื่องมือหรือโครงสร้างข้อมูล
 * จากไฟล์อื่นมาใช้ในไฟล์นี้ เปรียบเหมือนการหยิบอุปกรณ์ที่เตรียมไว้แล้ว
 * มาใช้งาน โดยไม่ต้องเขียนทุกอย่างซ้ำใหม่
 */

import ExcelJS from "exceljs";
import { getCellText } from "../../validators/shared/excel-cell.util"; 

/**
 * แปลง Excel Row เป็น Object
 *
 * ตัวอย่าง
 *
 * Header
 *
 * Txn Date
 * Currency
 * Amount
 *
 * Row
 *
 * 2025-11-25
 * HKD
 * 100
 *
 * Result
 *
 * {
 *    "Txn Date":"2025-11-25",
 *    "Currency":"HKD",
 *    "Amount":"100"
 * }
 */
export const mapRowToObject = (
  row: ExcelJS.Row,
  headers: string[],
): Record<string, unknown> => {

  const result: Record<string, unknown> = {};

  headers.forEach((header, index) => {

    if (!header) {
      return;
    }

    result[header] =
      getCellText(
        row.getCell(index + 1),
      );

  });

  return result;

};