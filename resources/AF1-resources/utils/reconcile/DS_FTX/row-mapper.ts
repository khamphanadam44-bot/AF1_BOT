/**
 * row-mapper.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * แปลงข้อมูล Excel จำนวน 1 แถว
 * ให้เป็น JavaScript Object
 *
 * ชื่อ Header จะถูกใช้เป็น Property Name หรือ Key
 * และค่าภายใน Cell จะถูกใช้เป็น Property Value
 *
 * ตัวอย่างข้อมูลใน Excel:
 *
 * | Test No. | From Currency (CCY) | To Currency (CCY) |
 * |----------|---------------------|-------------------|
 * | FTX-001  | EUR                 | JPY               |
 *
 * Object ที่ได้:
 *
 * {
 *   "Test No.": "FTX-001",
 *   "From Currency (CCY)": "EUR",
 *   "To Currency (CCY)": "JPY"
 * }
 *
 * Object ที่สร้างจากไฟล์นี้จะถูกนำไปใช้เป็น
 * - TestDataRow สำหรับ Test Data
 * - ReportRow สำหรับ Report DS_FTX
 *
 * หมายเหตุสำคัญ
 * - ค่าจาก Cell จะถูกแปลงเป็นข้อความ
 * - ช่องว่างด้านหน้าและด้านหลังของค่าจะถูกตัดออก
 * - Column ที่ไม่มีชื่อ Header จะไม่ถูกนำมาเก็บ
 * - ชื่อ Header จะถูกใช้ตามค่าที่ได้รับมาโดยตรง
 * - ไฟล์นี้ไม่รองรับ Header Alias
 * - ถ้ามีชื่อ Header ซ้ำ ข้อมูล Column หลังจะเขียนทับข้อมูลเดิม
 *
 * คำศัพท์
 * - Mapper   = ตัวแปลงข้อมูลจากรูปแบบหนึ่งไปเป็นอีกรูปแบบ
 * - Row      = แถวข้อมูล
 * - Header   = ชื่อหัวตาราง
 * - Key      = ชื่อ Property ภายใน Object
 * - Value    = ค่าที่อยู่ภายใน Property
 * - Object   = ชุดข้อมูลที่เก็บในรูปแบบ Key และ Value
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  getCellText,
} from "../../validators/shared/excel-cell.util";

/**
 * แปลง Excel Row จำนวน 1 แถวให้เป็น Object
 *
 * การทำงาน
 * 1. สร้าง Object ว่างสำหรับเก็บผลลัพธ์
 * 2. วนอ่าน Header ทีละตำแหน่ง
 * 3. ข้าม Column ที่ไม่มีชื่อ Header
 * 4. แปลง Array Index เป็น Excel Column Number
 * 5. อ่านค่าจาก Cell ใต้ Header
 * 6. แปลงค่า Cell เป็นข้อความ
 * 7. ตัดช่องว่างด้านหน้าและด้านหลัง
 * 8. บันทึกข้อมูลลงใน Object
 *
 * ตัวอย่าง:
 *
 * headers:
 *
 * [
 *   "Test No.",
 *   "Transaction ID/ Reconcile ID",
 *   "From Currency (CCY)"
 * ]
 *
 * ข้อมูล Excel Row:
 *
 * [
 *   "FTX-001",
 *   "TX001",
 *   "USD"
 * ]
 *
 * ผลลัพธ์:
 *
 * {
 *   "Test No.": "FTX-001",
 *   "Transaction ID/ Reconcile ID": "TX001",
 *   "From Currency (CCY)": "USD"
 * }
 *
 * @param row
 * แถวข้อมูลที่อ่านจาก ExcelJS
 *
 * @param headers
 * รายชื่อ Header เรียงตามตำแหน่ง Column
 *
 * ตัวอย่าง:
 * - headers[0] ตรงกับ Excel Column A
 * - headers[1] ตรงกับ Excel Column B
 * - headers[2] ตรงกับ Excel Column C
 *
 * @returns
 * Object ที่ใช้ Header เป็น Key และค่าจาก Cell เป็น Value
 *
 * แม้ Return Type จะประกาศ Value เป็น unknown
 * แต่ Code ปัจจุบันใช้ getCellText()
 * จึงเก็บค่าของ Cell เป็น string
 */
export const mapRowToObject = (
  row: ExcelJS.Row,
  headers: string[],
): Record<string, unknown> => {
  /**
   * สร้าง Object ว่างสำหรับเก็บผลลัพธ์
   *
   * ข้อมูลจากแต่ละ Column
   * จะถูกเพิ่มเข้าไปใน Object นี้
   */
  const result: Record<
    string,
    unknown
  > = {};

  /**
   * วนอ่าน Header ทีละรายการ
   *
   * header
   * = ชื่อ Header ปัจจุบัน
   *
   * index
   * = ตำแหน่ง Header ภายใน Array
   */
  headers.forEach(
    (
      header,
      index,
    ) => {
      /**
       * ถ้า Header เป็นข้อความว่าง
       * ให้ข้าม Column ปัจจุบัน
       *
       * ข้อมูลภายใน Cell ของ Column นี้
       * จะไม่ถูกเพิ่มเข้าไปใน Object
       *
       * return ตรงนี้ข้ามเฉพาะ Header ปัจจุบัน
       * ไม่ได้หยุดฟังก์ชัน mapRowToObject()
       */
      if (
        header === ""
      ) {
        return;
      }

      /**
       * แปลง Array Index เป็น ExcelJS Column Number
       *
       * Array เริ่มจาก Index 0:
       * - Index 0 = รายการแรก
       * - Index 1 = รายการที่สอง
       *
       * ExcelJS เริ่ม Column จาก 1:
       * - Column A = 1
       * - Column B = 2
       *
       * จึงต้องใช้ index + 1
       */
      const columnNumber =
        index + 1;

      /**
       * อ่านค่าจาก Cell ที่อยู่ใต้ Header
       *
       * getCellText()
       * = แปลงค่า ExcelJS Cell เป็นข้อความ
       *
       * trim()
       * = ตัดช่องว่างด้านหน้าและด้านหลัง
       *
       * ตัวอย่าง:
       * "  TX001  " จะกลายเป็น "TX001"
       */
      result[header] =
        getCellText(
          row.getCell(
            columnNumber,
          ),
        ).trim();
    },
  );

  /**
   * คืน Object ที่สร้างเสร็จแล้ว
   */
  return result;
};