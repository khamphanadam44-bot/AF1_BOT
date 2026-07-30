/**
 * field-helpers.util.ts
 * ------------------------------------------------------------------
 * รวมฟังก์ชันช่วยเหลือที่ใช้ร่วมกันระหว่าง
 * - Normal Field Validator = ตรวจสอบ Field ข้อมูลทั่วไป
 * - Fee Group Validator    = ตรวจสอบข้อมูลกลุ่มค่าธรรมเนียม
 *
 * หน้าที่หลัก
 * 1. ค้นหาตำแหน่ง Header โดยรองรับ Alias
 * 2. ดึง Cell จากชื่อ Header
 * 3. ใส่สีและข้อความตามผลการตรวจสอบ
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  createHeaderAliases,
  findMatchingHeaderIndex,
} from "../shared/header-matcher";

import {
  applyFill,
  CHECK_MESSAGE,
  COLORS,
  REQUIRED_MESSAGE,
} from "../shared/excel-style.util";

import {
  getFeeTypeCount,
} from "../../../config/testdata-helper";

/**
 * สร้าง Alias ของ Header สำหรับ Test Data
 *
 * Alias ประกอบด้วย
 * - Alias กลาง เช่น Txn Date และ Transaction Date
 * - Alias ของ Fee Header ตามจำนวน Fee Type
 *
 * ค่านี้จะถูกสร้างครั้งเดียวตอนเริ่มโหลดไฟล์นี้
 */
const TEST_DATA_HEADER_ALIASES =
  createHeaderAliases(
    getFeeTypeCount(),
  );

/**
 * ค้นหา Array Index ของ Column จากชื่อ Header
 *
 * รองรับการค้นหาผ่าน Alias เช่น
 * Expected Header: "Txn Date"
 * Actual Header:   "Transaction Date"
 *
 * @returns
 * - 0 ขึ้นไป เมื่อพบ Header
 * - -1 เมื่อไม่พบ Header
 */
export const findHeaderColumnIndex = (
  headers: string[],
  expectedHeader: string,
): number => {
  return findMatchingHeaderIndex(
    headers,
    expectedHeader,
    TEST_DATA_HEADER_ALIASES,
  );
};

/**
 * ดึง Cell ข้อมูลจากชื่อ Expected Header
 *
 * ขั้นตอน
 * 1. ค้นหา Array Index ของ Header
 * 2. ถ้าไม่พบ Header ให้คืน undefined
 * 3. ถ้าพบ ให้นำ Index + 1 ไปดึง Excel Cell
 *
 * ต้องบวก 1 เพราะ
 * - Array เริ่มจาก Index 0
 * - ExcelJS เริ่ม Column จากหมายเลข 1
 *
 * @returns
 * - ExcelJS.Cell เมื่อพบ Header
 * - undefined เมื่อไม่พบ Header
 */
export const getCellByHeader = (
  row: ExcelJS.Row,
  headers: string[],
  expectedHeader: string,
): ExcelJS.Cell | undefined => {
  // ค้นหาตำแหน่ง Header ภายใน Array
  const headerIndex =
    findHeaderColumnIndex(
      headers,
      expectedHeader,
    );

  /**
   * ถ้า headerIndex เป็น -1 แสดงว่าไม่พบ Header
   *
   * ถ้าพบ Header ให้นำ Index + 1
   * ไปดึง Cell จากแถวที่กำลังตรวจสอบ
   */
  return headerIndex === -1
    ? undefined
    : row.getCell(
        headerIndex + 1,
      );
};

/**
 * กำหนดให้ Cell เป็น Required Field ที่ว่าง
 *
 * ผลลัพธ์
 * - พื้นหลังสีแดง
 * - ใส่ข้อความ "โปรดกรอกข้อมูล"
 *
 * หมายเหตุ
 * คำสั่ง cell.value จะเขียนข้อความทับค่าเดิมใน Cell
 */
export const markRequiredCell = (
  cell: ExcelJS.Cell,
): void => {
  // ใส่สีแดง
  applyFill(
    cell,
    COLORS.RED,
  );

  // ใส่ข้อความ "โปรดกรอกข้อมูล"
  cell.value =
    REQUIRED_MESSAGE;
};

/**
 * กำหนดให้ Cell ต้องตรวจสอบเพิ่มเติม
 *
 * ใช้เฉพาะ Fee Group Case 3:
 * มี Fee Amount แต่ไม่มี Fee Type และ Fee Charge Account
 *
 * ผลลัพธ์
 * - พื้นหลังสีเหลือง
 * - ใส่ข้อความ "โปรดตรวจสอบข้อมูล"
 *
 * หมายเหตุ
 * คำสั่ง cell.value จะเขียนข้อความทับค่าเดิมใน Cell
 */
export const markCheckCell = (
  cell: ExcelJS.Cell,
): void => {
  // ใส่สีเหลือง
  applyFill(
    cell,
    COLORS.YELLOW,
  );

  // ใส่ข้อความ "โปรดตรวจสอบข้อมูล"
  cell.value =
    CHECK_MESSAGE;
};

/**
 * กำหนดให้ Cell มีข้อมูลหรือผ่านการตรวจสอบ
 *
 * ผลลัพธ์
 * - พื้นหลังสีเขียวอ่อน
 * - ไม่แก้ไขค่าข้อมูลเดิมใน Cell
 */
export const markSuccessCell = (
  cell: ExcelJS.Cell,
): void => {
  // ใส่สีเขียวอ่อนและคงค่าเดิมไว้
  applyFill(
    cell,
    COLORS.FIELD_GREEN,
  );
};