/**
 * field-helpers.util.ts
 * ------------------------------------------------------------
 * รวมฟังก์ชันช่วยเหลือที่ใช้ร่วมกันระหว่าง:
 *
 * 1. Normal Field Validator
 *    ตรวจข้อมูลทั่วไปที่เป็น Required Field
 *
 * 2. Fee Group Validator
 *    ตรวจข้อมูลในกลุ่มค่าธรรมเนียม
 *
 * หน้าที่หลัก:
 * 1. ค้นหาตำแหน่ง Header โดยรองรับ Alias
 * 2. สร้าง Header Alias ตามจำนวน Fee Group ที่พบจริง
 * 3. ดึง Cell จากชื่อ Header
 * 4. ใส่สีและข้อความตามผลการตรวจสอบ
 * ------------------------------------------------------------
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
  detectFeeTypeCount,
} from "../../../config/testdata-helper";

/**
 * ค้นหา Array Index ของ Column จากชื่อ Header
 *
 * รองรับการค้นหาผ่าน Alias
 *
 * ตัวอย่าง:
 *
 * Expected Header:
 * "Txn Date"
 *
 * Actual Header:
 * "Transaction Date"
 *
 * ฟังก์ชันจะมองว่า Header ทั้งสองชื่อนี้ตรงกัน
 *
 * สำหรับ Fee Group:
 * จำนวน Alias จะถูกสร้างตามหมายเลข Fee Group สูงสุด
 * ที่ตรวจพบจาก Header ใน Test Data จริง
 *
 * @returns
 * 0 ขึ้นไป = พบ Header
 * -1        = ไม่พบ Header
 */
export const findHeaderColumnIndex = (
  headers: string[],
  expectedHeader: string,
): number => {
  /**
   * ตรวจจำนวน Fee Group จาก Header จริง
   *
   * ตัวอย่าง:
   * หาก Header มี Fee Type 1 ถึง Fee Type 5
   * จะได้ feeTypeCount เท่ากับ 5
   */
  const feeTypeCount =
    detectFeeTypeCount(
      headers,
    );

  /**
   * สร้าง Alias ตามจำนวน Fee Group ที่ตรวจพบ
   *
   * ไม่สร้าง Alias แบบคงที่ตอนเปิดโปรแกรม
   * เพราะแต่ละไฟล์อาจมีจำนวน Fee Group ไม่เท่ากัน
   */
  const aliases =
    createHeaderAliases(
      feeTypeCount,
    );

  return findMatchingHeaderIndex(
    headers,
    expectedHeader,
    aliases,
  );
};

/**
 * ดึง Cell ข้อมูลจากชื่อ Expected Header
 *
 * ขั้นตอน:
 * 1. ค้นหา Array Index ของ Header
 * 2. ถ้าไม่พบ Header ให้คืน undefined
 * 3. ถ้าพบ Header ให้นำ Index + 1 ไปดึง Excel Cell
 *
 * ต้องบวก 1 เพราะ:
 * - Array เริ่มนับ Index จาก 0
 * - ExcelJS เริ่มนับ Column จาก 1
 *
 * @returns
 * ExcelJS.Cell = พบ Header
 * undefined    = ไม่พบ Header
 */
export const getCellByHeader = (
  row: ExcelJS.Row,
  headers: string[],
  expectedHeader: string,
): ExcelJS.Cell | undefined => {
  const headerIndex =
    findHeaderColumnIndex(
      headers,
      expectedHeader,
    );

  /**
   * หาก headerIndex เป็น -1
   * แสดงว่าไม่พบ Header ในไฟล์
   */
  if (headerIndex === -1) {
    return undefined;
  }

  /**
   * บวก 1 เพื่อแปลง Array Index
   * ให้เป็นหมายเลข Column ของ ExcelJS
   */
  return row.getCell(
    headerIndex + 1,
  );
};

/**
 * กำหนดให้ Cell เป็น Required Field ที่ว่าง
 *
 * ผลลัพธ์:
 * - พื้นหลังสีแดง
 * - ใส่ข้อความ "โปรดกรอกข้อมูล"
 *
 * ใช้เมื่อ:
 * Field เป็น Required แต่ไม่มีข้อมูล
 *
 * หมายเหตุ:
 * cell.value จะเขียนข้อความลงใน Cell
 */
export const markRequiredCell = (
  cell: ExcelJS.Cell,
): void => {
  applyFill(
    cell,
    COLORS.RED,
  );

  cell.value =
    REQUIRED_MESSAGE;
};

/**
 * กำหนดให้ Cell ต้องตรวจสอบเพิ่มเติม
 *
 * ผลลัพธ์:
 * - พื้นหลังสีเหลือง
 * - ใส่ข้อความ "โปรดตรวจสอบข้อมูล"
 *
 * ตัวอย่างกรณีใช้งาน:
 * มี Fee Amount แต่ข้อมูลสำคัญบางช่อง
 * ภายใน Fee Group ไม่ครบ
 *
 * หมายเหตุ:
 * cell.value จะเขียนข้อความลงใน Cell
 */
export const markCheckCell = (
  cell: ExcelJS.Cell,
): void => {
  applyFill(
    cell,
    COLORS.YELLOW,
  );

  cell.value =
    CHECK_MESSAGE;
};

/**
 * กำหนดให้ Cell มีข้อมูลหรือผ่านการตรวจสอบ
 *
 * ผลลัพธ์:
 * - พื้นหลังสีเขียวอ่อน
 * - ไม่แก้ไขค่าข้อมูลเดิมใน Cell
 */
export const markSuccessCell = (
  cell: ExcelJS.Cell,
): void => {
  applyFill(
    cell,
    COLORS.FIELD_GREEN,
  );
};