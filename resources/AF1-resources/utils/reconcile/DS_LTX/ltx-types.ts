/**
 * Result Models ของ Reconcile Service
 * ------------------------------------------------------------------

 * ------------------------------------------------------------------
 */

/**
 * สถานะผลการ Reconcile ตาม Requirement:
 * - PASS   : ข้อมูลทั้ง Test Data และ AF1 Report ตรงกัน (เขียว)
 * - FAIL   : ไม่ตรงตาม Requirement การ reconcile แบบชัดเจน/ไม่กำกวม (แดง)
 * - REVIEW : มีข้อมูลฝั่งใดฝั่งหนึ่งว่าง หรือมีข้อมูลทั้ง 2 ฝั่งแต่ไม่ตรงกัน
 *            ต้องตรวจสอบใหม่อีกครั้ง (เหลือง)
 */
export type ReconcileStatus = "PASS" | "FAIL" | "REVIEW";

/** ผลเทียบ 1 field rule ของ 1 คู่ (Test Case <-> Report row เดียว) */
export interface FieldCheckResult {
  fieldHeader: string;
  /** @deprecated ใช้ status แทน — เก็บไว้เพื่อ backward-compatibility เท่านั้น (= status === "PASS") */
  isMatch: boolean;
  status: ReconcileStatus;
  remark: string;
}

/** ผลของ 1 แถว AF1 Report ที่ถูกจับคู่กับ Test Case แล้ว */
export interface RowAnnotation {
  testCaseNo: string;
  /** @deprecated ใช้ status แทน — เก็บไว้เพื่อ backward-compatibility เท่านั้น (= status === "PASS") */
  isPassed: boolean;
  status: ReconcileStatus;
  remark: string;
  failedFieldHeaders: string[];
}

/** Mapping ของ 1 Report field ไปยังตำแหน่ง column ใน Test Data */
export interface ReportFieldTestDataMapping {
  reportField: string;
  testDataColumnPosition: number;
  testDataFieldName: string;
}

/**
 * Test Case (หรือ suffix DR/FE ของ Test Case) ที่หาแถวคู่กันใน AF1 Report ไม่เจอเลย
 * (เดิมแค่ console.warn ทิ้ง ไม่ปรากฏใน Sheet ผลลัพธ์ — เป็นบั๊กที่แก้ในรอบนี้ เพราะทำให้
 * QA ไม่เห็นเลยว่า Test Case นี้ไม่ได้ถูก reconcile จริง ต้องเปิด terminal log ดูเท่านั้น)
 */
export interface UnmatchedCase {
  testCaseNo: string;
  suffix: string;
  remark: string;
}
