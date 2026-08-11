/**
 * ftx-types.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * กำหนด Type และ Interface กลางที่ใช้ใน Script 3
 * สำหรับกระบวนการ Reconcile ของ Report DS_FTX
 *
 * ไฟล์นี้ช่วยกำหนดว่า Object แต่ละประเภท
 * ต้องมี Property อะไรและใช้ชนิดข้อมูลแบบใด
 *
 * ข้อมูลหลักที่กำหนดในไฟล์นี้
 * 1. CompareStatus = สถานะผลการเปรียบเทียบ
 * 2. TestDataRow   = ข้อมูลดิบ 1 แถวจาก Test Data
 * 3. ReportRow     = ข้อมูลดิบ 1 แถวจาก Report
 * 4. ExpectedRow   = ข้อมูลที่ระบบคาดหวังจาก Test Data
 * 5. ActualRow     = ข้อมูลจริงที่อ่านจาก Report
 * 6. CompareResult = ผลการตรวจสอบข้อมูล 1 Field
 *
 * Flow ของข้อมูล
 *
 * Test Data
 * → แปลงเป็น ExpectedRow
 *
 * Report DS_FTX
 * → แปลงเป็น ActualRow
 *
 * ExpectedRow + ActualRow
 * → เปรียบเทียบเป็น CompareResult
 *
 * หมายเหตุ
 * - ไฟล์นี้ไม่ได้อ่านไฟล์ Excel
 * - ไฟล์นี้ไม่ได้เปรียบเทียบข้อมูล
 * - ไฟล์นี้ไม่ได้สร้างไฟล์ผลลัพธ์
 * - Type และ Interface ใช้ตรวจสอบโครงสร้างตอนเขียน Code
 *   และจะไม่เหลืออยู่หลัง TypeScript ถูกแปลงเป็น JavaScript
 *
 * คำศัพท์
 * - Type       = การกำหนดชนิดหรือค่าที่อนุญาต
 * - Interface  = แบบแปลนโครงสร้างของ Object
 * - Property   = ข้อมูลแต่ละช่องภายใน Object
 * - Expected   = ค่าที่ระบบคาดหวัง
 * - Actual     = ค่าที่พบจริง
 * - Reconcile  = การจับคู่และตรวจสอบข้อมูลสองฝั่ง
 * ------------------------------------------------------------------
 */

/**
 * สถานะผลการเปรียบเทียบข้อมูลของ DS_FTX
 *
 * ปัจจุบันมีเพียง 2 สถานะ
 *
 * PASS
 * = ผลการตรวจสอบผ่าน
 *
 * FAIL
 * = ผลการตรวจสอบไม่ผ่าน
 *
 * รายการที่เข้า Exclusion Rule จะใช้สถานะ PASS
 * พร้อม Remark อธิบายสาเหตุ
 *
 * จึงไม่มีสถานะ SKIP ใน Logic ปัจจุบัน
 */
export type CompareStatus =
  | "PASS"
  | "FAIL";

/**
 * รูปแบบข้อมูลดิบ 1 แถวจากไฟล์ Test Data
 *
 * Property Name หรือ Key
 * = ชื่อ Header ของ Test Data
 *
 * Property Value
 * = ค่าที่อ่านได้จาก Cell ใต้ Header นั้น
 *
 * ตัวอย่าง:
 *
 * {
 *   "Test No.": "BOTDMS_001",
 *   "Transaction ID/ Reconcile ID": "TX001",
 *   "From Currency (CCY)": "THB"
 * }
 *
 * ใช้ unknown เพราะค่าจาก Excel อาจเป็นได้หลายชนิด เช่น
 * - string
 * - number
 * - boolean
 * - Date
 * - null
 * - Formula Object
 * - Rich Text Object
 *
 * หมายเหตุ
 * Interface นี้ไม่ได้บังคับชื่อ Header ตายตัว
 * จึงรองรับ Property ที่เป็น string ทุกชื่อ
 */
export interface TestDataRow {
  [header: string]: unknown;
}

/**
 * รูปแบบข้อมูลดิบ 1 แถวจาก Report DS_FTX
 *
 * Property Name หรือ Key
 * = ชื่อ Header ของ Report
 *
 * Property Value
 * = ค่าที่อ่านได้จาก Cell ใต้ Header นั้น
 *
 * ตัวอย่าง:
 *
 * {
 *   "Ref. TX No.": "TX001",
 *   "Buy Currency Id": "THB",
 *   "Transaction Date": "2026-07-30"
 * }
 *
 * ใช้ unknown เพราะค่าจาก Excel
 * สามารถมีได้หลายชนิดข้อมูล
 *
 * หมายเหตุ
 * Interface นี้ไม่ได้บังคับชื่อ Header ตายตัว
 * จึงรองรับ Property ที่เป็น string ทุกชื่อ
 */
export interface ReportRow {
  [header: string]: unknown;
}

/**
 * รูปแบบข้อมูลที่ระบบคาดหวังจาก Test Data จำนวน 1 แถว
 *
 * ExpectedRow ถูกสร้างโดย ftx-row-builder.ts
 *
 * หน้าที่ของ ExpectedRow คือรวบรวม
 * - หมายเลขแถวจริง
 * - Test No.
 * - Matching Key
 * - ข้อมูลทั้งหมดของ Test Data แถวนั้น
 *
 * เพื่อนำไปจับคู่และเปรียบเทียบกับ ActualRow
 */
export interface ExpectedRow {
  /**
   * หมายเลขแถวจริงในไฟล์ Test Data
   *
   * ตัวอย่าง:
   * ถ้าข้อมูลอยู่แถวที่ 6 ใน Excel
   * rowNumber จะมีค่าเป็น 6
   */
  rowNumber: number;

  /**
   * ค่า Test No. จาก Test Data
   *
   * ใช้เพื่อ
   * - ระบุ Test Case
   * - แสดงในไฟล์ผลลัพธ์
   * - ช่วยให้ผู้ใช้งานย้อนกลับไปตรวจ Test Data ได้ง่าย
   *
   * ไม่ได้นำไปเปรียบเทียบกับ Report
   * เพราะ Report DS_FTX ไม่มี Header Test No.
   */
  testScriptNo: string;

  /**
   * กุญแจสำหรับจับคู่ข้อมูลระหว่าง Test Data และ Report
   *
   * ฝั่ง Test Data:
   * Transaction ID/ Reconcile ID
   *
   * ฝั่ง Report:
   * Ref. TX No.
   *
   * ระบบจะใช้ค่านี้ค้นหาว่า
   * ExpectedRow ตรงกับ ActualRow แถวใด
   */
  matchingKey: string;

  /**
   * ข้อมูลดิบทั้งหมดของ Test Data แถวนั้น
   *
   * ชื่อ Header จะเป็น Key
   * และค่าจาก Cell จะเป็น Value
   *
   * ข้อมูลนี้จะถูกนำไปใช้สร้าง Expected Value
   * สำหรับเปรียบเทียบแต่ละ Field
   */
  data: TestDataRow;
}

/**
 * รูปแบบข้อมูลจริงที่อ่านจาก Report DS_FTX จำนวน 1 แถว
 *
 * ActualRow ถูกสร้างโดย ftx-row-builder.ts
 *
 * หน้าที่ของ ActualRow คือรวบรวม
 * - หมายเลขแถวจริงใน Report
 * - Matching Key
 * - ข้อมูลทั้งหมดของ Report แถวนั้น
 *
 * เพื่อนำไปจับคู่และเปรียบเทียบกับ ExpectedRow
 */
export interface ActualRow {
  /**
   * หมายเลขแถวจริงในไฟล์ Report
   *
   * ตัวอย่าง:
   * ถ้าข้อมูลอยู่แถวที่ 10 ใน Excel
   * rowNumber จะมีค่าเป็น 10
   */
  rowNumber: number;

  /**
   * Matching Key ที่อ่านจาก Header "Ref. TX No."
   * ของ Report DS_FTX
   *
   * ใช้สำหรับค้นหา ExpectedRow
   * ที่มี Matching Key เดียวกัน
   */
  matchingKey: string;

  /**
   * ข้อมูลดิบทั้งหมดของ Report แถวนั้น
   *
   * ชื่อ Header จะเป็น Key
   * และค่าจาก Cell จะเป็น Value
   */
  data: ReportRow;
}

/**
 * รูปแบบผลการตรวจสอบข้อมูลจำนวน 1 รายการ
 *
 * โดยทั่วไป CompareResult จำนวน 1 รายการ
 * จะแทนผลการตรวจสอบ 1 Field
 *
 * ตัวอย่าง:
 *
 * {
 *   matchingKey: "TX001",
 *   testScriptNo: "BOTDMS_001",
 *   testDataRowNumber: 6,
 *   reportRowNumber: 10,
 *   field: "Buy Currency Id",
 *   expected: "THB",
 *   actual: "THB",
 *   status: "PASS",
 *   remark: "Matched"
 * }
 *
 * นอกจากผลการเปรียบเทียบ Field แล้ว
 * CompareResult ยังใช้กับกรณีพิเศษ เช่น
 * - Matching Key ว่าง
 * - ไม่พบ Matching Key
 * - Matching Key ซ้ำ
 * - เข้า Exclusion Rule
 */
export interface CompareResult {
  /**
   * Matching Key ของรายการที่กำลังตรวจสอบ
   *
   * โดยปกติจะมาจาก ExpectedRow
   *
   * ถ้า Matching Key ใน Test Data ว่าง
   * Property นี้จะเป็นข้อความว่าง ""
   */
  matchingKey: string;

  /**
   * ค่า Test No. จาก Test Data
   *
   * ใช้แสดงในไฟล์ผลลัพธ์
   * เพื่อช่วยระบุว่าเป็นผลของ Test Case ใด
   */
  testScriptNo: string;

  /**
   * หมายเลขแถวจริงในไฟล์ Test Data
   *
   * ใช้ช่วยให้ผู้ใช้งานกลับไปตรวจสอบ
   * ข้อมูลต้นทางได้ง่าย
   */
  testDataRowNumber: number;

  /**
   * หมายเลขแถวจริงใน Report DS_FTX
   *
   * ถ้าพบและจับคู่ ActualRow ได้
   * จะเป็นหมายเลขแถวจริงใน Report
   *
   * ถ้าไม่มีแถว Report ที่ถูกนำมาเปรียบเทียบ
   * จะใช้ค่า 0 เช่น
   * - Matching Key ใน Test Data ว่าง
   * - ไม่พบ Matching Key ใน Report
   * - รายการเข้า Exclusion Rule
   *
   * กรณี Matching Key ซ้ำ
   * Code ปัจจุบันจะใช้หมายเลขแถวแรกที่พบ
   */
  reportRowNumber: number;

  /**
   * ชื่อ Field หรือหัวข้อที่กำลังตรวจสอบ
   *
   * ตัวอย่าง Field ปกติ:
   * - Buy Currency Id
   * - Sell Currency Id
   * - Transaction Date
   *
   * ตัวอย่างกรณีพิเศษ:
   * - Matching Key
   * - Exclusion Rule
   */
  field: string;

  /**
   * ค่าที่ระบบคาดหวังจาก Test Data
   *
   * กรณีเปรียบเทียบ Field ปกติ
   * ค่านี้จะถูกสร้างจาก ExpectedRow
   *
   * กรณีพิเศษอาจเป็นข้อความอธิบาย เช่น
   * - "Excluded from DS_FTX"
   * - "1 Report row"
   */
  expected: unknown;

  /**
   * ค่าที่พบจริงจาก Report DS_FTX
   *
   * กรณีไม่พบข้อมูลอาจเป็นข้อความว่าง ""
   *
   * กรณีพิเศษอาจเป็นข้อความอธิบาย เช่น
   * - "Not compared"
   * - "2 Report rows"
   */
  actual: unknown;

  /**
   * สถานะผลการตรวจสอบ
   *
   * PASS = ผ่าน
   * FAIL = ไม่ผ่าน
   */
  status: CompareStatus;

  /**
   * รายละเอียดหรือสาเหตุของผลการตรวจสอบ
   *
   * ตัวอย่าง:
   * - Matched
   * - Value mismatch
   * - Matching Key is empty in Test Data
   * - Matching Key Not Found in DS_FTX
   * - Duplicate Matching Key in DS_FTX
   * - รายละเอียดของ Exclusion Rule
   */
  remark: string;
}