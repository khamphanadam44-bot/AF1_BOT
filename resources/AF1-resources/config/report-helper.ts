/**
 * ======================================================
 * ไฟล์: report-helper.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวช่วยอ่านค่า REPORT_CONFIG โดยไฟล์อื่นไม่ต้องเข้าถึงโครงสร้าง Config โดยตรง
 * ใช้ดึง Header ที่จำเป็น เลขแถว Header และชื่อ Header สำรองของ Report ที่เลือก
 *
 * ======================================================
 */
import { REPORT_CONFIG } from "./report-config";


export type ReportName =
  keyof typeof REPORT_CONFIG;

/**
 * ดึง Header ทั้งหมดของ Report
 */
export function getRequiredReportHeaders(

  reportName: string,

): string[] {

  const groups =
    REPORT_CONFIG[
      reportName as ReportName
    ].requiredHeaders;

  return Object.values(groups)
    .flat()
    .map(String);

}


/**
 * ดึง Header Row Number
 */
export function getHeaderRowNumber(

  reportName: string,

): number {

  return REPORT_CONFIG[
    reportName as ReportName
  ].headerRowNumber;

}
/**
 * ดึง Alias
 */
export function getAliases(

  reportName: string,

): Record<string, string[]> {

  const aliases =
    REPORT_CONFIG[
      reportName as ReportName
    ].aliases;

  return Object.fromEntries(

    Object.entries(aliases).map(

      ([key, value]) => [

        key,

        [...value],

      ],

    ),

  );

}