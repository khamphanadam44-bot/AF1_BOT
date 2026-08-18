/**
 * ======================================================
 * ไฟล์: summary-types.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * กำหนดรูปแบบข้อมูลที่ใช้ร่วมกันในขั้นตอนสร้าง Automation Summary
 * ประกอบด้วยสถานะผลตรวจ แถวจาก Compare Result และข้อมูลสรุปของการรันแต่ละครั้ง
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
export type SummaryStatus = "PASS" | "FAIL" | "SKIP";

export interface CompareResultRow {
  testScriptNo: string;
  matchingKey: string;
  status: SummaryStatus;
  remark: string;
  reportValues: Record<string, unknown>;
}

export interface AutomationSummaryInfo {
  reportFileName: string;

  /**
   * เวลาเริ่ม Script 1 ของ Report นี้
   * ใช้แสดง Execution Date และ Actual Execution Time Start
   */
  automationStartedAt: Date;

  /** เวลาเริ่ม Script 4 ของ Report นี้ */
  script4StartedAt: Date;

  /**
   * เวลารวมของ Script 1-3 หน่วย Millisecond
   * Script 4 จะบวกเวลาของตัวเองก่อนเขียน Duration Time
   */
  completedStageDurationMilliseconds: number;

  runId: string;
  verifiedBy: string;
  totalChecked: number;
  passed: number;
  failed: number;
}
