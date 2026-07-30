/**
 * IExpectedCaseBuilder
 * ------------------------------------------------------------------
 * Strategy interface: แต่ละ Report (DS_LTX, DS_PTX, ...) มี business logic การแปลง
 * Test Data เป็น "Expected Case" ไม่เหมือนกัน — ReconcileService เรียกผ่าน interface
 * นี้เท่านั้น ไม่ผูกกับ implementation ของ report ไหนเป็นพิเศษ
 *
 * Bug fix / Redesign (Code Review):
 * เดิม ExpectedCase คือ "1 Test Case = 1 amount รวม (SUM Fee ของทุกแถวที่เกี่ยวข้อง)"
 * แล้วเอายอดรวมนั้นไปหาแถว Report ที่ยอดเงินตรงกัน (amount-based fuzzy match)
 *
 * ตรวจสอบกับไฟล์จริงแล้วพบว่าวิธีนี้ผิดหลักการ: แต่ละแถว Test Data (ทั้งแถวหลักและ
 * แถว "-Return") มี Reference Transaction Number ของตัวเอง (มาจาก "Transaction ID/
 * Reconcile ID") ที่ตรงกับ Report ตรง ๆ อยู่แล้ว AF1 Report ไม่ได้มีแถวเดียวที่ยอดรวม
 * เท่ากับผลรวม Fee ทั้งหมด แต่แยกเป็นคนละแถวตาม Reference Number ของแต่ละแถว Test Data
 *
 * เปลี่ยนเป็น "1 แถว Test Data (ไม่ว่าจะเป็นแถวหลักหรือแถว -Return) = 1 ExpectedCase"
 * ที่รู้ Reference Number ที่คาดหวังของแถว DR/FE ของตัวเองตรง ๆ ไม่ต้องเดา/หายอดเงิน
 * ------------------------------------------------------------------
 */
import { ReconcileReportConfig } from "./ltx-config";
import { ReconcileRecord } from "../shared/record";

export interface ExpectedCase {
  /** Test No. ที่จะแสดงผลใน Sheet (ตัด suffix "-Return" ออกแล้วถ้ามี เพื่อให้จัดกลุ่มกับแถวหลัก) */
  displayTestCaseNo: string;

  /** แถว Test Data ต้นทางที่ใช้เทียบ field values (Transaction Date, Currency, ฯลฯ) ของ ExpectedCase นี้ */
  primaryRecord: ReconcileRecord;

  /** Reference Transaction Number ที่คาดหวังของแถว DR — undefined = แถว Test Data นี้ไม่ได้คาดหวังแถว DR */
  expectedDrReference: string | undefined;

  /** Reference Transaction Number ที่คาดหวังของแถว FE — undefined = แถว Test Data นี้ไม่ได้คาดหวังแถว FE */
  expectedFeReference: string | undefined;

  /**
   * ยอดเงินที่คาดหวังของแถว FE = SUM ของ Fee Amount Type 1-5 ของแถว Test Data นี้เอง
   * (ไม่ใช่ plain field เดียวเหมือน DR จึงต้องคำนวณแยกไว้ล่วงหน้า ไม่ผ่าน fieldRule ทั่วไป)
   *
   * Bug fix (Code Review): เดิมตอนเปลี่ยนมาใช้ ID-based matching ไม่มีการตรวจ Transaction
   * Amount ของแถว FE เลยแม้แต่จุดเดียว (หายไปพร้อมกับ amount-based matching เดิมที่เคย
   * ตรวจ "โดยอ้อม" ผ่านการหายอดที่ตรงกัน) ทั้งที่ Requirement ต้องการให้ยอด Transaction
   * Amount ของแถว FE ตรงกับผลรวม Fee ทุกรายการ (บางเคสมี Fee มากกว่า 1 รายการ)
   */
  expectedFeAmount: string;
}

export interface IExpectedCaseBuilder {
  build(
    headers: string[],
    testDataRecords: ReconcileRecord[],
    config: ReconcileReportConfig,
  ): ExpectedCase[];
}
