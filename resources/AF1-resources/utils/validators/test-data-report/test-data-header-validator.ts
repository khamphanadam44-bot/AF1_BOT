/**
 * test-data-header-validator.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้ตรวจสอบ Header ของไฟล์ Test Data
 * ซึ่งเป็นคนละไฟล์กับ Downloaded Report
 *
 * การทำงานหลัก
 * 1. เลือก Worksheet แรกของ Test Data
 * 2. อ่าน Header จากแถวที่กำหนด
 * 3. สร้าง Alias กลางและ Fee Header Alias
 * 4. เปรียบเทียบ Actual Header กับ Expected Header
 * 5. คืนรายชื่อ Header ที่ไม่พบ
 *
 * หมายเหตุ
 * - ไฟล์นี้ไม่ได้สร้าง Sheet "Header Validation"
 * - ไฟล์นี้ไม่ได้ Highlight Header
 * - ไฟล์นี้ไม่ได้บันทึก Workbook ลงไฟล์
 * - ไฟล์นี้ไม่ Throw Error เมื่อ Header ขาด
 *
 * คำศัพท์
 * - Test Data       = ไฟล์ข้อมูลที่ใช้เป็นเงื่อนไขในการทดสอบ
 * - Actual Header   = Header ที่พบจริงในไฟล์ Test Data
 * - Expected Header = Header ที่ระบบคาดหวังจาก Requirement
 * - Missing Header  = Header ที่ระบบค้นหาไม่พบ
 * - Alias           = ชื่ออื่นที่อนุญาตให้ใช้แทนชื่อหลัก
 * ------------------------------------------------------------------
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
  getFeeTypeCount,
} from "../../../config/testdata-helper";

/**
 * ตรวจสอบ Header ของไฟล์ Test Data
 *
 * ขั้นตอนการทำงาน
 * 1. เลือก Worksheet แรกจาก Workbook
 * 2. อ่าน Header จากแถว headerRowNumber
 * 3. สร้าง Alias ที่ใช้ในการค้นหา Header
 * 4. คำนวณผลว่า Expected Header แต่ละรายการพบหรือไม่
 * 5. คืนเฉพาะรายชื่อ Expected Header ที่ไม่พบ
 *
 * ตัวอย่าง:
 *
 * expectedHeaders:
 *
 * [
 *   "Transaction ID/ Reconcile ID",
 *   "Txn Date",
 *   "Currency Id",
 * ]
 *
 * Actual Header ในไฟล์:
 *
 * [
 *   "Transaction ID / Reconcile ID",
 *   "Transaction Date",
 * ]
 *
 * ผลลัพธ์:
 *
 * ["Currency Id"]
 *
 * เพราะ
 * - Transaction ID / Reconcile ID พบผ่าน Alias
 * - Transaction Date พบผ่าน Alias ของ Txn Date
 * - Currency Id ไม่พบ
 *
 * @param workbook
 * Workbook ของไฟล์ Test Data ที่เปิดด้วย ExcelJS แล้ว
 *
 * @param expectedHeaders
 * รายการ Header ที่ระบบคาดหวังว่าจะต้องพบใน Test Data
 *
 * @param headerRowNumber
 * หมายเลขแถวที่เป็น Header ของ Test Data
 *
 * ตัวอย่าง:
 * ถ้า Header อยู่แถวที่ 5 ให้ส่งค่า 5
 *
 * ปัจจุบัน DS_PTX และ DS_FTX ใช้ Header Row แถวที่ 5
 * แต่ฟังก์ชันนี้ไม่ได้ตรวจ Report Code ด้วยตัวเอง
 * ผู้เรียกต้องส่งหมายเลขแถวที่ถูกต้องเข้ามา
 *
 * @returns
 * Array ของชื่อ Expected Header ที่ค้นหาไม่พบ
 *
 * - Array ว่าง [] = พบ Header ครบ
 * - Array มีข้อมูล = มี Header ขาด
 */
export const validateTestDataHeader = (
  workbook: ExcelJS.Workbook,
  expectedHeaders: string[],
  headerRowNumber: number,
): string[] => {
  // แสดงหัวข้อการตรวจสอบใน Console
  console.log(
    "\n===== TEST DATA HEADER VALIDATION =====",
  );

  /**
   * เลือก Worksheet ลำดับแรกของ Workbook
   *
   * getWorksheet(1)
   * หมายถึง Worksheet ลำดับที่ 1
   * ไม่ได้หมายถึง Worksheet ที่มีชื่อว่า "1"
   */
  const worksheet =
    workbook.getWorksheet(1);

  /**
   * ถ้า Workbook ไม่มี Worksheet แรก
   * ให้หยุดการทำงานและแจ้ง Error
   */
  if (!worksheet) {
    throw new Error(
      "Worksheet not found",
    );
  }

  /**
   * อ่าน Header ทั้งหมดจากแถวที่กำหนด
   *
   * ตัวอย่าง:
   * headerRowNumber = 5
   * ระบบจะอ่าน Header จากแถวที่ 5
   *
   * ผลลัพธ์จะเป็น Array เช่น:
   *
   * [
   *   "Transaction ID/ Reconcile ID",
   *   "Txn Date",
   *   "Currency Id",
   * ]
   */
  const actualHeaders =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * สร้าง Alias สำหรับใช้จับคู่ Header
   *
   * Alias ที่ได้ประกอบด้วย
   * 1. Common Header Alias
   *    เช่น Txn Date และ Transaction Date
   *
   * 2. Fee Header Alias
   *    เช่น Fee Amount Type 2 และ Fee Amount 2
   *
   * จำนวน Fee Header จะอ่านจาก getFeeTypeCount()
   *
   * หมายเหตุ:
   * Code ปัจจุบันไม่ได้ส่ง Custom Alias เฉพาะ Report
   * เข้าไปใน createHeaderAliases()
   */
  const aliases =
    createHeaderAliases(
      getFeeTypeCount(),
    );

  /**
   * เปรียบเทียบ Expected Header ทุกตัว
   * กับ Actual Header ที่พบใน Test Data
   *
   * รองรับการ Match ผ่าน Alias
   *
   * ผลลัพธ์แต่ละรายการจะมีข้อมูล เช่น
   * - expectedHeader
   * - matchedHeader
   * - isFound
   * - similarHeader
   * - remark
   */
  const results =
    findHeaderMatchResults(
      actualHeaders,
      expectedHeaders,
      aliases,
    );

  /**
   * กรองผลลัพธ์ให้เหลือเฉพาะชื่อ Header ที่ไม่พบ
   *
   * ตัวอย่างผลลัพธ์:
   *
   * []
   * = พบ Header ครบทั้งหมด
   *
   * ["Currency Id", "Payment Method"]
   * = ไม่พบ Header 2 รายการ
   *
   * ฟังก์ชันนี้ไม่ Throw Error เมื่อพบ Missing Header
   * ผู้เรียกต้องนำ Array ที่ได้ไปตรวจสอบต่อ
   */
  return getMissingHeaders(
    results,
  );
};