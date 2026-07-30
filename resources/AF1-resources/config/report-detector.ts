/**
 * ======================================================
 * ไฟล์: report-detector.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * อ่านชื่อไฟล์ Report แล้วค้นหาว่าไฟล์นั้นเป็น Report ประเภทใด เช่น DS_PTX หรือ DS_FTX
 * ถ้าชื่อไฟล์ไม่มีรหัส Report ที่ระบบรู้จัก ฟังก์ชันจะหยุดและแจ้ง Error เพื่อป้องกันการใช้ Config ผิดชุด
 *
 * ======================================================
 */
import { REPORT_CONFIG } from "./report-config";

export type ReportCode = keyof typeof REPORT_CONFIG;

/**
 * อ่านชื่อไฟล์ แล้วหา Report Code
 *
 * ตัวอย่าง
 * EXPORT_DS_PTX_20251125.xlsx
 * -> DS_PTX
 */
export const detectReportCodeFromFileName = (
  fileName: string,
): ReportCode => {

  const upperFileName = fileName.toUpperCase();

  const reportCodes = Object.keys(REPORT_CONFIG) as ReportCode[];

  const matched = reportCodes.find((code) =>
    upperFileName.includes(code),
  );

  if (!matched) {

    throw new Error(
      `ไม่สามารถระบุ Report Type จากชื่อไฟล์ "${fileName}" ได้`,
    );

  }

  return matched;

};