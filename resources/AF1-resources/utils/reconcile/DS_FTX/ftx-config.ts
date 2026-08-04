/**
 * ftx-config.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * กำหนด Mapping Rule สำหรับเปรียบเทียบ
 * ข้อมูล Test Data กับ Report DS_FTX
 *
 * ไฟล์นี้ระบุว่า
 * 1. ต้องอ่านค่าจาก Header ใดใน Test Data
 * 2. ต้องนำไปเทียบกับ Header ใดใน Report
 * 3. ต้องใช้วิธีเปรียบเทียบแบบ TEXT หรือ DATE
 *
 * ตัวอย่าง Mapping:
 *
 * Test Data:
 * From Currency (CCY)
 *
 * เทียบกับ Report:
 * Buy Currency Id
 *
 * ไฟล์นี้ยังไม่ได้เปรียบเทียบค่าจริง
 * การเปรียบเทียบจะเกิดขึ้นใน ftx-rules.ts *
 * Flow การทำงาน
 *
 * DS_FTX_COMPARE_RULES
 * → ftx-rules.ts วนอ่าน Rule
 * → resolveExpectedValue() อ่านค่าจาก Test Data
 * → อ่านค่าจาก Report ด้วย reportField
 * → Normalize ตาม compareType
 * → สร้าง CompareResult
 *
 * หมายเหตุสำคัญ
 * - ชื่อ Header ใน Config ต้องตรงกับ Object Key
 * - ไม่รองรับ Alias ในขั้นตอนอ่านค่า
 * - ลำดับ Rule ใน Array จะเป็นลำดับผลลัพธ์ที่ถูกสร้าง
 * - ถ้าเพิ่ม Rule ใหม่ ftx-rules.ts จะนำไปตรวจอัตโนมัติ
 * ------------------------------------------------------------------
 */

import {
  ExpectedRow,
} from "./ftx-types";

/**
 * ประเภทของวิธีเปรียบเทียบข้อมูล
 *
 * ปัจจุบันรองรับ 2 รูปแบบ
 *
 * TEXT
 * = เปรียบเทียบข้อมูลแบบข้อความ
 *
 * ftx-rules.ts จะ
 * - เปลี่ยน null หรือ undefined เป็นข้อความว่าง
 * - ตัดช่องว่างด้านหน้าและด้านหลัง
 * - รวมช่องว่างหลายช่องให้เหลือช่องเดียว
 * - เปลี่ยนข้อความเป็นตัวพิมพ์ใหญ่
 *
 * DATE
 * = เปรียบเทียบข้อมูลแบบวันที่
 *
 * ftx-rules.ts จะพยายามแปลงวันที่
 * ให้อยู่ในรูปแบบ yyyy-mm-dd ก่อนเปรียบเทียบ
 */
export type CompareType =
  | "TEXT"
  | "DATE";

/**
 * รูปแบบของกฎการเปรียบเทียบจำนวน 1 รายการ
 *
 * ตัวอย่าง:
 *
 * {
 *   reportField: "Buy Currency Id",
 *   testDataField: "From Currency (CCY)",
 *   compareType: "TEXT"
 * }
 *
 * หมายความว่า
 * ให้นำค่าจาก Test Data Header "From Currency (CCY)"
 * ไปเทียบกับ Report Header "Buy Currency Id"
 * ด้วยวิธีเปรียบเทียบแบบข้อความ
 */
export interface CompareRule {
  /**
   * ชื่อ Header ฝั่ง Report DS_FTX
   *
   * ftx-rules.ts จะใช้ค่านี้อ่านข้อมูลจาก:
   *
   * actualRow.data[reportField]
   *
   * และใช้เป็นชื่อ Field ใน CompareResult
   */
  reportField: string;

  /**
   * ชื่อ Header ฝั่ง Test Data
   *
   * resolveExpectedValue() จะใช้ค่านี้อ่านข้อมูลจาก:
   *
   * expectedRow.data[testDataField]
   */
  testDataField: string;

  /**
   * วิธีที่ใช้ Normalize และเปรียบเทียบค่า
   *
   * ค่าที่อนุญาต:
   * - TEXT
   * - DATE
   */
  compareType: CompareType;
}

/**
 * รายการ Mapping Rule ของ Report DS_FTX
 *
 * ปัจจุบันตรวจสอบ Core Field ทั้งหมด 3 รายการ
 *
 * Rule 1
 * Report:    Buy Currency Id
 * Test Data: From Currency (CCY)
 * Type:      TEXT
 *
 * Rule 2
 * Report:    Sell Currency Id
 * Test Data: Settled Currency (CCY)
 * Type:      TEXT
 *
 * Rule 3
 * Report:    Transaction Date
 * Test Data: Txn Date
 * Type:      DATE
 *
 * หมายเหตุสำคัญ
 * Rule ของ Sell Currency Id ใช้ Settled Currency (CCY)
 * ไม่ได้ใช้ To Currency (CCY)
 */
export const DS_FTX_COMPARE_RULES:
CompareRule[] = [
  /**
   * Rule 1:
   * เปรียบเทียบสกุลเงินต้นทาง
   *
   * From Currency (CCY)
   * → Buy Currency Id
   */
  {
    reportField:
      "Buy Currency Id",

    testDataField:
      "From Currency (CCY)",

    compareType:
      "TEXT",
  },

  /**
   * Rule 2:
   * เปรียบเทียบสกุลเงินที่ Settlement
   *
   * Settled Currency (CCY)
   * → Sell Currency Id
   */
  {
    reportField:
      "Sell Currency Id",

    testDataField:
      "Settled Currency (CCY)",

    compareType:
      "TEXT",
  },

  /**
   * Rule 3:
   * เปรียบเทียบวันที่ทำรายการ
   *
   * Txn Date
   * → Transaction Date
   */
  {
    reportField:
      "Transaction Date",

    testDataField:
      "Txn Date",

    compareType:
      "DATE",
  },
];

/**
 * อ่านค่าที่ระบบคาดหวังจาก ExpectedRow
 * ตามชื่อ Header ที่กำหนดใน CompareRule
 *
 * ตัวอย่าง:
 *
 * rule.testDataField:
 * "From Currency (CCY)"
 *
 * ระบบจะอ่าน:
 *
 * expectedRow.data["From Currency (CCY)"]
 *
 * ถ้าพบ Property:
 * จะคืนค่าที่อยู่ภายใน Property นั้น
 *
 * ถ้าไม่พบ Property:
 * JavaScript จะคืนค่า undefined
 *
 * หมายเหตุ
 * - ใช้ชื่อ Header แบบตรงตัว
 * - ไม่ค้นหาผ่าน Alias
 * - ไม่ Normalize ค่า
 * - ไม่ตรวจว่าค่าว่างหรือไม่
 * - ไม่ Throw Error เมื่อไม่พบ Header
 *
 * การ Normalize และตัดสิน PASS/FAIL
 * เป็นหน้าที่ของ ftx-rules.ts *
 * @param expectedRow
 * ExpectedRow ที่สร้างจาก Test Data
 *
 * @param rule
 * กฎการเปรียบเทียบที่กำลังตรวจสอบ
 *
 * @returns
 * - ค่าจาก Test Data เมื่อพบ Header
 * - undefined เมื่อไม่พบ Header
 */
export const resolveExpectedValue = (
  expectedRow: ExpectedRow,
  rule: CompareRule,
): unknown => {
  return expectedRow.data[
    rule.testDataField
  ];
};