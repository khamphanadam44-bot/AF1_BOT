/**
 * report-row-builder.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * อ่านข้อมูลจาก Worksheet ของ Report DS_FTX
 * แล้วแปลงแต่ละแถวให้เป็น ActualRow
 *
 * ActualRow จำนวน 1 รายการประกอบด้วย
 * 1. หมายเลขแถวจริงใน Report
 * 2. Matching Key จาก Header "Ref. TX No."
 * 3. ข้อมูลทั้งหมดของ Report แถวนั้น
 *
 * ActualRow เป็นข้อมูลฝั่งที่พบจริง
 * และจะถูกนำไปจับคู่กับ ExpectedRow จาก Test Data
 *
 * Flow การทำงาน
 *
 * Report DS_FTX Worksheet
 * → อ่าน Header
 * → อ่านข้อมูลทีละแถว
 * → แปลงแถวเป็น Object
 * → ข้ามแถวว่าง
 * → อ่าน Matching Key
 * → สร้าง ActualRow[]
 *
 * หมายเหตุสำคัญ
 * - Header เริ่มต้นอยู่ที่แถว 1
 * - ผู้เรียกสามารถส่งหมายเลข Header Row อื่นได้
 * - แถวว่างระหว่างข้อมูลจะถูกข้าม
 * - แถวที่มีข้อมูลแต่ Matching Key ว่างยังถูกเก็บใน ActualRow[]
 * - ActualRow จะเรียงตามลำดับแถวใน Report
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  ActualRow,
  ReportRow,
} from "./compare-types";

import {
  buildReportMatchingKey,
} from "./build-matching-key";

import {
  mapRowToObject,
} from "./row-mapper";

import {
  getCellText,
} from "../../validators/shared/excel-cell.util";

/**
 * หมายเลขแถว Header เริ่มต้นของ Report DS_FTX
 *
 * ปัจจุบัน Report DS_FTX มี Header อยู่ที่แถว 1
 *
 * ค่านี้เป็นเพียงค่าเริ่มต้น
 * ผู้เรียกสามารถส่งหมายเลขแถวอื่นให้ฟังก์ชันได้
 */
export const DEFAULT_REPORT_HEADER_ROW_NUMBER =
  1;

/**
 * แปลงค่าทั่วไปให้อยู่ในรูปแบบข้อความ
 *
 * การทำงาน
 * 1. ถ้าค่าเป็น null หรือ undefined ให้คืนข้อความว่าง
 * 2. แปลงค่าเป็น string
 * 3. ตัดช่องว่างด้านหน้าและด้านหลัง
 *
 * ตัวอย่าง:
 * null       → ""
 * undefined  → ""
 * 123        → "123"
 * " TX001 "  → "TX001"
 *
 * @param value ค่าที่ต้องการแปลงเป็นข้อความ
 * @returns ข้อความที่ตัดช่องว่างแล้ว
 */
const normalizeText = (
  value: unknown,
): string => {
  // ถ้าไม่มีค่า ให้คืนข้อความว่าง
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  // แปลงเป็น string และตัดช่องว่างหน้า–หลัง
  return String(
    value,
  ).trim();
};

/**
 * อ่านรายชื่อ Header จาก Worksheet ของ Report DS_FTX
 *
 * ตำแหน่งใน Header Array จะตรงกับตำแหน่ง Column
 *
 * ตัวอย่าง:
 *
 * [
 *   "Ref. TX No.",          // Excel Column A
 *   "Buy Currency Id",      // Excel Column B
 *   "Sell Currency Id",     // Excel Column C
 *   "Transaction Date"      // Excel Column D
 * ]
 *
 * Header ที่อ่านได้จะถูก
 * - แปลงเป็นข้อความด้วย getCellText()
 * - ตัดช่องว่างด้านหน้าและด้านหลัง
 *
 * แต่จะไม่ถูก
 * - เปลี่ยนตัวพิมพ์เล็กหรือใหญ่
 * - ลบช่องว่างภายใน
 * - เปลี่ยนผ่าน Alias
 *
 * ฟังก์ชันนี้ถูก Export เพราะนอกจากใช้สร้าง ActualRow แล้ว
 * รายการ Header ยังถูกนำไปใช้สร้าง Column ในไฟล์ผลลัพธ์
 *
 * @param worksheet Worksheet ของ Report DS_FTX
 * @param headerRowNumber หมายเลขแถว Header
 *
 * @returns รายการ Header เรียงตามตำแหน่ง Column
 */
export const getReportHeaders = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber =
    DEFAULT_REPORT_HEADER_ROW_NUMBER,
): string[] => {
  /**
   * ดึงแถวที่กำหนดให้เป็น Header
   */
  const headerRow =
    worksheet.getRow(
      headerRowNumber,
    );

  /**
   * หาจำนวน Column ที่ต้องอ่าน
   *
   * ใช้ค่าที่มากที่สุดระหว่าง
   *
   * worksheet.actualColumnCount
   * = จำนวน Column ที่ ExcelJS มองว่ามีข้อมูลใน Worksheet
   *
   * headerRow.cellCount
   * = จำนวน Cell ที่แถว Header ครอบคลุม
   *
   * หมายเหตุ:
   * Logic นี้เหมาะกับตารางที่ Column เรียงต่อเนื่องกัน
   */
  const lastColumnNumber =
    Math.max(
      worksheet.actualColumnCount,
      headerRow.cellCount,
    );

  /**
   * Array สำหรับเก็บชื่อ Header
   */
  const headers: string[] =
    [];

  /**
   * วนอ่าน Header ตั้งแต่ Column 1
   * จนถึง Column สุดท้ายที่คำนวณได้
   */
  for (
    let columnNumber = 1;
    columnNumber <=
    lastColumnNumber;
    columnNumber += 1
  ) {
    /**
     * อ่านชื่อ Header และตัดช่องว่างหน้า–หลัง
     *
     * ถ้า Header Cell ว่าง
     * ค่าที่ได้จะเป็นข้อความว่าง ""
     */
    const header =
      getCellText(
        headerRow.getCell(
          columnNumber,
        ),
      ).trim();

    /**
     * เพิ่ม Header เข้า Array
     *
     * Header ว่างจะยังถูกเพิ่มเป็น ""
     * เพื่อรักษาตำแหน่ง Column ให้ตรงกับ Excel
     */
    headers.push(
      header,
    );
  }

  return headers;
};

/**
 * ตรวจสอบว่า ReportRow มีข้อมูลจริงอย่างน้อย 1 Field หรือไม่
 *
 * ใช้สำหรับข้าม
 * - แถวว่างระหว่างข้อมูล
 * - แถวว่างท้าย Worksheet
 *
 * การทำงาน
 * 1. ดึง Value ทั้งหมดจาก Object
 * 2. Normalize Value ทีละรายการ
 * 3. ถ้ามีค่าใดไม่ว่าง ให้ถือว่าแถวมีข้อมูล
 *
 * หมายเหตุ
 * mapRowToObject() ไม่เก็บข้อมูลของ Column ที่ไม่มี Header
 *
 * ดังนั้น ถ้าแถวมีข้อมูลเฉพาะใน Column ที่ Header ว่าง
 * ฟังก์ชันนี้จะถือว่าแถวนั้นไม่มีข้อมูล
 *
 * @param data ข้อมูล Report จำนวน 1 แถว
 *
 * @returns
 * true  = แถวมีข้อมูลอย่างน้อย 1 Field
 * false = แถวไม่มีข้อมูล
 */
const hasRowData = (
  data: ReportRow,
): boolean => {
  return Object.values(
    data,
  ).some(
    (
      value,
    ) => {
      return normalizeText(
        value,
      ) !== "";
    },
  );
};

/**
 * สร้าง ActualRow ทั้งหมดจาก Worksheet ของ Report DS_FTX
 *
 * ขั้นตอน
 * 1. ตรวจสอบหมายเลขแถว Header
 * 2. อ่าน Header
 * 3. เริ่มอ่านข้อมูลจากแถวถัดจาก Header
 * 4. แปลงข้อมูลแต่ละแถวเป็น Object
 * 5. ข้ามแถวที่ไม่มีข้อมูล
 * 6. อ่าน Matching Key จาก "Ref. TX No."
 * 7. สร้าง ActualRow
 *
 * @param worksheet
 * Worksheet ของ Report DS_FTX
 *
 * @param headerRowNumber
 * หมายเลขแถว Header
 *
 * ค่าเริ่มต้นคือแถว 1
 *
 * @returns
 * ActualRow[] เรียงตามหมายเลขแถวใน Report
 */
export const buildActualRows = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber =
    DEFAULT_REPORT_HEADER_ROW_NUMBER,
): ActualRow[] => {
  /**
   * ป้องกันหมายเลขแถวที่น้อยกว่า 1
   *
   * ExcelJS เริ่มนับแถวจาก 1
   * จึงไม่สามารถใช้แถว 0 หรือเลขติดลบได้
   *
   * หมายเหตุ:
   * ตรวจเฉพาะค่าที่น้อยกว่า 1
   * ไม่ได้ตรวจว่า Header Row เกินจำนวนแถวหรือไม่
   */
  if (
    headerRowNumber < 1
  ) {
    throw new Error(
      `Invalid Report header row number: ${headerRowNumber}`,
    );
  }

  /**
   * อ่าน Header ทั้งหมดจาก Report
   */
  const headers =
    getReportHeaders(
      worksheet,
      headerRowNumber,
    );

  /**
   * Array สำหรับเก็บ ActualRow ทั้งหมด
   */
  const actualRows:
  ActualRow[] = [];

  /**
   * เริ่มอ่านข้อมูลจากแถวถัดจาก Header
   *
   * ตัวอย่าง:
   * Header อยู่แถว 1
   * ข้อมูลจะเริ่มอ่านจากแถว 2
   */
  for (
    let rowNumber =
      headerRowNumber + 1;

    /**
     * วนอ่านจนถึงหมายเลขแถวสุดท้ายของ Worksheet
     *
     * ใช้ worksheet.rowCount
     * เพื่อไม่ให้ข้อมูลหลังแถวว่างคั่นกลางหายไป
     */
    rowNumber <=
    worksheet.rowCount;
    rowNumber += 1
  ) {
    /**
     * ดึง Excel Row ตามหมายเลขแถวปัจจุบัน
     */
    const excelRow =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * แปลง Excel Row ให้เป็น Object
     * โดยใช้ชื่อ Header เป็น Key
     *
     * ตัวอย่าง:
     *
     * {
     *   "Ref. TX No.": "TX001",
     *   "Buy Currency Id": "EUR",
     *   "Sell Currency Id": "JPY",
     *   "Transaction Date": "2025-11-25"
     * }
     *
     * as ReportRow เป็น Type Assertion
     * เพื่อบอก TypeScript ว่าต้องการใช้ Object
     * ในรูปแบบ ReportRow
     *
     * ไม่ได้ทำการแปลงข้อมูลเพิ่มเติมตอน Run
     */
    const data =
      mapRowToObject(
        excelRow,
        headers,
      ) as ReportRow;

    /**
     * ถ้าแถวไม่มีข้อมูล ให้ข้าม
     *
     * continue หมายถึงจบรอบปัจจุบัน
     * และไปอ่านแถวถัดไป
     */
    if (
      !hasRowData(
        data,
      )
    ) {
      continue;
    }

    /**
     * อ่าน Matching Key จาก Header "Ref. TX No."
     *
     * buildReportMatchingKey()
     * จะตัดช่องว่างด้านหน้าและด้านหลัง
     *
     * ถ้าไม่พบ Header หรือไม่มีค่า
     * matchingKey จะเป็นข้อความว่าง ""
     */
    const matchingKey =
      buildReportMatchingKey(
        data,
      );

    /**
     * รวมข้อมูลของแถวปัจจุบันเป็น ActualRow
     *
     * แม้ Matching Key จะว่าง
     * แถวนี้ก็ยังถูกเก็บไว้ใน actualRows
     *
     * การใช้งานภายหลัง:
     * - actual-row-index.ts จะไม่เพิ่มแถวที่ Key ว่างลง Index
     * - compare-validator.ts จึงไม่สามารถใช้แถวนี้จับคู่ได้
     * - compare-result-writer.ts ยังสามารถนำแถวนี้
     *   ไปแสดงเป็น Report Row ที่ไม่พบ Test Data ที่ Map กัน
     */
    actualRows.push({
      // หมายเลขแถวจริงใน Report
      rowNumber,

      // Matching Key จาก Ref. TX No.
      matchingKey,

      // ข้อมูลทั้งหมดของ Report แถวนี้
      data,
    });
  }

  /**
   * คืน ActualRow ทั้งหมด
   * โดยเรียงตามลำดับแถวใน Report
   */
  return actualRows;
};