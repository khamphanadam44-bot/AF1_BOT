/**
 * test-data-header-validator.ts
 * ------------------------------------------------------------
 * ตรวจสอบ Header ของไฟล์ Test Data
 *
 * หน้าที่หลัก:
 * 1. เลือก Worksheet แรกของ Test Data
 * 2. อ่าน Header จากแถวที่กำหนด
 * 3. ตรวจจำนวน Fee Group จาก Header จริง
 * 4. สร้าง Common Alias และ Fee Header Alias
 * 5. เปรียบเทียบ Actual Header กับ Expected Header
 * 6. คืนรายชื่อ Header ที่ค้นหาไม่พบ
 *
 * หมายเหตุ:
 * - ไม่ได้สร้าง Sheet "Header Validation"
 * - ไม่ได้ Highlight Header
 * - ไม่ได้บันทึก Workbook
 * - ไม่ Throw Error เมื่อ Header ขาด
 *
 * คำศัพท์:
 * - Actual Header
 *   ชื่อ Header ที่พบจริงในไฟล์ Test Data
 *
 * - Expected Header
 *   ชื่อ Header ที่ระบบต้องการให้มี
 *
 * - Missing Header
 *   Expected Header ที่ค้นหาไม่พบ
 *
 * - Alias
 *   ชื่ออื่นที่อนุญาตให้ใช้แทนชื่อ Header หลัก
 * ------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  createHeaderAliases,
} from "../shared/header-matcher";

import {
  getHeadersFromRow,
} from "../shared/excel-cell.util";

import {
  findHeaderMatchResults,
  getMissingHeaders,
} from "../shared/header-validation-sheet";

import {
  detectFeeTypeCount,
} from "../../../config/testdata-helper";

/**
 * ตรวจสอบ Header ของไฟล์ Test Data
 *
 * ขั้นตอน:
 * 1. เลือก Worksheet แรก
 * 2. อ่าน Actual Header จากแถวที่กำหนด
 * 3. ตรวจจำนวน Fee Group จาก Actual Header
 * 4. สร้าง Header Alias
 * 5. เปรียบเทียบ Expected Header กับ Actual Header
 * 6. คืนรายชื่อ Header ที่ไม่พบ
 *
 * @param workbook
 * Workbook ของ Test Data ที่เปิดด้วย ExcelJS แล้ว
 *
 * @param expectedHeaders
 * รายการ Header ที่ระบบต้องการตรวจ
 *
 * @param headerRowNumber
 * หมายเลขแถว Header ของ Test Data
 *
 * @returns
 * Array ของชื่อ Expected Header ที่ค้นหาไม่พบ
 *
 * ตัวอย่าง:
 *
 * [] หมายถึง พบ Header ครบทั้งหมด
 *
 * ["Txn Date"] หมายถึง ไม่พบ Header "Txn Date"
 */
export const validateTestDataHeader = (
  workbook: ExcelJS.Workbook,
  expectedHeaders: string[],
  headerRowNumber: number,
): string[] => {
  console.log(
    "\n===== TEST DATA HEADER VALIDATION =====",
  );

  /**
   * เลือก Worksheet ลำดับแรกของ Workbook
   *
   * getWorksheet(1) หมายถึง Worksheet ลำดับที่ 1
   * ไม่ได้หมายถึง Worksheet ที่มีชื่อว่า "1"
   */
  const worksheet =
    workbook.getWorksheet(1);

  if (!worksheet) {
    throw new Error(
      "Worksheet not found",
    );
  }

  /**
   * อ่าน Header จริงจากแถวที่กำหนด
   *
   * ตัวอย่าง:
   * headerRowNumber = 5
   * หมายถึงอ่าน Header จากแถวที่ 5
   */
  const actualHeaders =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * ตรวจหมายเลข Fee Group สูงสุดจาก Header จริง
   *
   * ตัวอย่าง:
   * หากพบ Fee Type 1 ถึง Fee Type 5
   * จะได้ feeTypeCount เท่ากับ 5
   *
   * หากไม่พบ Fee Header
   * จะได้ feeTypeCount เท่ากับ 0
   */
  const feeTypeCount =
    detectFeeTypeCount(
      actualHeaders,
    );

  /**
   * สร้าง Alias สำหรับใช้จับคู่ Header
   *
   * ประกอบด้วย:
   * 1. Common Header Alias
   *    เช่น Txn Date กับ Transaction Date
   *
   * 2. Fee Header Alias
   *    เช่น Fee Amount Type 2 กับ Fee Amount 2
   *
   * จำนวน Fee Header Alias จะสร้างตามจำนวน
   * Fee Group ที่ตรวจพบใน Test Data จริง
   */
  const aliases =
    createHeaderAliases(
      feeTypeCount,
    );

  /**
   * เปรียบเทียบ Expected Header ทุกตัว
   * กับ Actual Header ที่พบใน Test Data
   *
   * รองรับการจับคู่ผ่าน Alias
   */
  const results =
    findHeaderMatchResults(
      actualHeaders,
      expectedHeaders,
      aliases,
    );

  /**
   * คืนเฉพาะรายชื่อ Header ที่ค้นหาไม่พบ
   *
   * ฟังก์ชันนี้ไม่ Throw Error เมื่อ Header ขาด
   * ผู้เรียกจะเป็นผู้ตัดสินใจว่าจะดำเนินการอย่างไรต่อ
   */
  return getMissingHeaders(
    results,
  );
};