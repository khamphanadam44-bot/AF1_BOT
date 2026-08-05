/**
 * ============================================================================
 * ไฟล์: ptx-types.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้เป็นเหมือน “แบบฟอร์มกลาง” ที่กำหนดว่าข้อมูลแต่ละชนิดใน Script 3
 * ต้องมีหน้าตาอย่างไร เพื่อให้ทุกไฟล์สื่อสารกันด้วยโครงสร้างเดียวกัน
 *
 * ตัวอย่างข้อมูลที่กำหนดไว้
 * -------------------------
 * - ExpectedRow: ข้อมูลที่คาดหวัง ซึ่งสร้างมาจาก Test Data
 * - ActualRow: ข้อมูลจริงที่อ่านมาจาก AF1 Report
 * - CompareResult: ผลตรวจราย Field เช่น PASS หรือ FAIL
 * - GroupedCompareResult: ผลหลาย Field ที่ถูกรวมตาม Matching Key
 *
 * ประโยชน์ของไฟล์นี้
 * ------------------
 * - ลดการส่งข้อมูลผิดรูปแบบระหว่างไฟล์
 * - ช่วยให้ TypeScript ตรวจจับข้อผิดพลาดก่อนรัน
 * - ทำให้ผู้พัฒนารู้ว่าข้อมูลแต่ละชุดต้องมี Property ใดบ้าง
 *
 * ไฟล์นี้ไม่มี Business Logic และไม่ได้เปรียบเทียบข้อมูลด้วยตัวเอง
 * ============================================================================
 */

/**
 * Status ของผล Compare
 */
export type CompareStatus =
  | "PASS"
  | "FAIL"
  | "SKIP";

/**
 * Row ของ Test Data
 */
export interface TestDataRow {
  [header: string]: unknown;
}

/**
 * Row ของ AF1 Report
 */
export interface ReportRow {
  [header: string]: unknown;
}

/**
 * ข้อมูลที่คาดหวังจาก Test Data (Expected Data)
 * สร้างจาก Test Data
 */
export interface ExpectedRow {

  /**
   * Row Number ของ Test Data
   */
  rowNumber: number;

  /**
   * Test Script No.
   */
  testScriptNo: string;

  /**
   * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
   *
   * กรณีไม่มี Fee จะเป็น Internal Key:
   * NO_FEE_ROW_<rowNumber>
   */
  matchingKey: string;

  /**
   * Running Number ของ Fee
   *
   * 0 = ไม่มี Fee
   * 1, 2, 3... = ลำดับ Fee
   */
  runningNumber: number;

  /**
   * Fee Type
   */
  feeType: string;

  /**
   * Fee Amount
   */
  feeAmount: number;

  /**
   * ระบุว่า Test Case นี้มี Fee หรือไม่
   */
  hasFee: boolean;

  /**
   * ข้อมูลทั้งแถวของ Test Data
   */
  data: TestDataRow;

}

/**
 * ข้อมูลจริงจาก Report (Actual Data)
 * สร้างจาก AF1 Report
 */
export interface ActualRow {

  /**
   * Row Number ของ Report
   */
  rowNumber: number;

  /**
   * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
   */
  matchingKey: string;

  /**
   * ข้อมูลทั้งแถวของ Report
   */
  data: ReportRow;

}

/**
 * ผลการ Compare ราย Field
 */
export interface CompareResult {

  /**
   * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
   */
  matchingKey: string;

  /**
   * Row ของ Test Data
   */
  testDataRowNumber: number;

  /**
   * Row ของ Report
   *
   * ถ้าไม่พบ Matching Key หรือ SKIP จะเป็น 0
   */
  reportRowNumber: number;

  /**
   * ชื่อ Field
   */
  field: string;

  /**
   * Expected Value
   */
  expected: unknown;

  /**
   * Actual Value
   */
  actual: unknown;

  /**
   * PASS / FAIL / SKIP
   */
  status: CompareStatus;

  /**
   * สาเหตุ
   */
  remark: string;

}

/**
 * ============================================================================
 * ผลเปรียบเทียบที่รวมเป็นกลุ่ม
 * ----------------------------------------------------------------------------
 * 1 Test Data Row + 1 Matching Key = 1 Group
 * Matching Key เดียวกันจึงมีได้หลาย Group
 * ใช้สำหรับ Export Excel
 * ============================================================================
 */
export interface GroupedCompareResult {

  /**
   * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
   */
  matchingKey: string;

  /**
   * Row ของ Test Data
   */
  testDataRowNumber: number;

  /**
   * Row ของ Report
   */
  reportRowNumber: number;

  /**
   * ผล Compare ของแต่ละ Field
   */
  fields: Record<string, CompareResult>;

}
