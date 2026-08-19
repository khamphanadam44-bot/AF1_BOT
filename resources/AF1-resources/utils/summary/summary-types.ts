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
  executionDate: string;
  executionTime: string;
  runId: string;
  verifiedBy: string;
  totalChecked: number;
  passed: number;
  failed: number;
}
