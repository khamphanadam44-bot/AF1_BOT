/**
 * build-matching-key.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้อ่านและปรับรูปแบบ Matching Key
 * สำหรับกระบวนการ Reconcile ของ DS_FTX
 *
 * Matching Key ใช้สำหรับจับคู่ข้อมูลระหว่าง
 *
 * Test Data:
 * Header "Transaction ID/ Reconcile ID"
 *
 * Report DS_FTX:
 * Header "Ref. TX No."
 *
 * ถึงแม้ชื่อ Header ของทั้งสองไฟล์จะแตกต่างกัน
 * แต่ค่าภายใน Cell ต้องเป็น Transaction เดียวกัน
 *
 * ตัวอย่าง:
 *
 * Test Data:
 * Transaction ID/ Reconcile ID = "TX001"
 *
 * Report:
 * Ref. TX No. = "TX001"
 *
 * ระบบจึงสามารถจับคู่ข้อมูลสองแถวนี้เข้าด้วยกันได้
 *
 * หมายเหตุสำคัญ
 * - ไฟล์นี้ไม่ได้สร้าง Matching Key จากหลาย Field
 * - ไฟล์นี้อ่านค่าจาก Header ที่กำหนดไว้โดยตรง
 * - การค้นหาชื่อ Header ไม่รองรับ Alias
 * - Matching Key ไม่ถูกเปลี่ยนเป็นตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 * - ลบเฉพาะช่องว่างด้านหน้าและด้านหลัง
 * ------------------------------------------------------------------
 */

import {
  ReportRow,
  TestDataRow,
} from "./compare-types";

/**
 * ชื่อ Header ที่ใช้เป็น Matching Key ใน Test Data
 *
 * ต้องตรงกับชื่อ Property ที่อยู่ใน TestDataRow
 *
 * ตัวอย่าง:
 *
 * {
 *   "Transaction ID/ Reconcile ID": "TX001"
 * }
 *
 * หมายเหตุ
 * การเข้าถึงข้อมูลด้วย Object Key เป็นแบบตรงตัว
 *
 * ตัวอย่าง Header ต่อไปนี้ถือว่าเป็นคนละ Key:
 * - Transaction ID/ Reconcile ID
 * - Transaction ID / Reconcile ID
 *
 * แม้ Header Validator อาจรองรับชื่อที่สองผ่าน Alias
 * แต่ฟังก์ชันในไฟล์นี้ยังอ่านด้วยชื่อแรกเท่านั้น
 */
export const TEST_DATA_MATCHING_KEY_HEADER =
  "Transaction ID/ Reconcile ID";

/**
 * ชื่อ Header ที่ใช้เป็น Matching Key ใน Report DS_FTX
 *
 * ต้องตรงกับชื่อ Property ที่อยู่ใน ReportRow
 *
 * ตัวอย่าง:
 *
 * {
 *   "Ref. TX No.": "TX001"
 * }
 */
export const REPORT_MATCHING_KEY_HEADER =
  "Ref. TX No.";

/**
 * ปรับค่า Matching Key ให้อยู่ในรูปแบบข้อความ
 * ที่พร้อมนำไปจับคู่
 *
 * ขั้นตอนการทำงาน
 * 1. ตรวจสอบว่าค่าเป็น null หรือ undefined หรือไม่
 * 2. ถ้าไม่มีค่า ให้คืนข้อความว่าง
 * 3. แปลงค่าเป็น string
 * 4. ลบช่องว่างด้านหน้าและด้านหลัง
 *
 * ตัวอย่าง:
 *
 * " TX001 "
 * จะกลายเป็น
 * "TX001"
 *
 * หมายเหตุสำคัญ
 *
 * ฟังก์ชันนี้ไม่ได้
 * - ลบช่องว่างภายในข้อความ
 * - เปลี่ยนตัวพิมพ์เล็กเป็นตัวพิมพ์ใหญ่
 * - เปลี่ยนตัวพิมพ์ใหญ่เป็นตัวพิมพ์เล็ก
 * - ลบเครื่องหมายพิเศษ
 * - สร้าง Matching Key จากหลาย Field
 *
 * ดังนั้นค่าต่อไปนี้ถือว่าแตกต่างกัน:
 *
 * "TX001"
 * "tx001"
 * "TX 001"
 *
 * @param value
 * ค่า Matching Key ที่อ่านได้จาก Test Data หรือ Report
 *
 * @returns
 * - Matching Key ที่แปลงเป็น string และตัดช่องว่างแล้ว
 * - ข้อความว่าง "" เมื่อค่าเป็น null หรือ undefined
 */
export const normalizeMatchingKey = (
  value: unknown,
): string => {
  /**
   * ถ้าไม่มีค่า Matching Key
   * ให้คืนข้อความว่าง
   */
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  /**
   * แปลงค่าเป็น string
   * และลบช่องว่างด้านหน้าและด้านหลัง
   *
   * ตัวอย่าง:
   * 12345      → "12345"
   * " TX001 "  → "TX001"
   */
  return String(
    value,
  ).trim();
};

/**
 * อ่าน Matching Key จากข้อมูล Test Data จำนวน 1 แถว
 *
 * อ่านค่าจาก Header:
 * "Transaction ID/ Reconcile ID"
 *
 * ขั้นตอน
 * 1. รับ TestDataRow
 * 2. อ่านค่าจากชื่อ Header ที่กำหนด
 * 3. ส่งค่าไปปรับรูปแบบด้วย normalizeMatchingKey()
 * 4. คืน Matching Key ที่พร้อมนำไปใช้งาน
 *
 * ตัวอย่าง Input:
 *
 * {
 *   "Test No.": "BOTDMS_001",
 *   "Transaction ID/ Reconcile ID": " TX001 "
 * }
 *
 * ผลลัพธ์:
 *
 * "TX001"
 *
 * หมายเหตุ
 * ถ้า Test Data ใช้ชื่อ Header ต่างจากค่าคงที่
 * เช่น "Transaction ID / Reconcile ID"
 * Code จะอ่านค่าไม่พบและคืนข้อความว่าง
 *
 * @param row ข้อมูล Test Data จำนวน 1 แถว
 * @returns Matching Key จาก Test Data
 */
export const buildTestDataMatchingKey = (
  row: TestDataRow,
): string => {
  /**
   * อ่านค่าด้วยชื่อ Header แบบตรงตัว
   *
   * ถ้าไม่พบ Property นี้
   * ค่าที่ได้จะเป็น undefined
   */
  return normalizeMatchingKey(
    row[
      TEST_DATA_MATCHING_KEY_HEADER
    ],
  );
};

/**
 * อ่าน Matching Key จากข้อมูล Report DS_FTX จำนวน 1 แถว
 *
 * อ่านค่าจาก Header:
 * "Ref. TX No."
 *
 * ขั้นตอน
 * 1. รับ ReportRow
 * 2. อ่านค่าจากชื่อ Headerที่กำหนด
 * 3. ส่งค่าไปปรับรูปแบบด้วย normalizeMatchingKey()
 * 4. คืน Matching Key ที่พร้อมนำไปใช้งาน
 *
 * ตัวอย่าง Input:
 *
 * {
 *   "Ref. TX No.": " TX001 ",
 *   "Buy Currency Id": "USD"
 * }
 *
 * ผลลัพธ์:
 *
 * "TX001"
 *
 * หมายเหตุ
 * ถ้า Report ใช้ชื่อ Header ต่างจาก "Ref. TX No."
 * Code จะอ่านค่าไม่พบและคืนข้อความว่าง
 *
 * @param row ข้อมูล Report DS_FTX จำนวน 1 แถว
 * @returns Matching Key จาก Report
 */
export const buildReportMatchingKey = (
  row: ReportRow,
): string => {
  /**
   * อ่านค่าด้วยชื่อ Header แบบตรงตัว
   *
   * ถ้าไม่พบ Property นี้
   * ค่าที่ได้จะเป็น undefined
   */
  return normalizeMatchingKey(
    row[
      REPORT_MATCHING_KEY_HEADER
    ],
  );
};