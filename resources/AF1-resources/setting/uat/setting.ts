/**
 * setting.ts
 * ค่าตั้งต้นสำหรับระบบ UAT
 * ------------------------------------------------------------
 */

import {
  TEST_DATA_INPUT_PATH,
} from "../../config/paths.config";

/**
 * อ่าน Environment Variable แบบบังคับ
 *
 * หากไม่พบค่าหรือเป็นค่าว่าง
 * ระบบจะหยุดและแจ้ง Error
 */
const requireEnv = (
  key: string,
): string => {
  const value =
    process.env[key];

  if (
    !value ||
    value.trim() === ""
  ) {
    throw new Error(
      `Missing required environment variable: ${key}\n` +
        `กรุณากำหนด ${key} ในไฟล์ .env ก่อนรัน`,
    );
  }

  return value;
};

/**
 * ข้อมูลสำหรับเข้าใช้งานระบบ UAT
 */
export const webSetting = {
  url:
    process.env.AF1_UAT_URL ??
    "http://192.168.35.99:7081/portal/index",

  get username(): string {
    return requireEnv(
      "AF1_UAT_USERNAME",
    );
  },

  get password(): string {
    return requireEnv(
      "AF1_UAT_PASSWORD",
    );
  },
};

/**
 * รายชื่อ Report ทั้งหมดที่ระบบรองรับ
 *
 * ใช้สำหรับตรวจสอบชื่อ Report
 * ที่รับมาจาก Terminal
 */
export const dmsReportNames = [
  "DS_PTX",
  "DS_FTX",
  "DS_FTU",
  "DF_FXU",
  "DS_LTX",
  "DF_OLB",
  "DF_FXM",
] as const;

/**
 * Type ของชื่อ Report
 *
 * TypeScript จะอนุญาตเฉพาะชื่อ
 * ที่อยู่ใน dmsReportNames เท่านั้น
 */
export type DmsReportName =
  (typeof dmsReportNames)[number];


/**
 * Report เริ่มต้น
 *
 * ใช้เมื่อไม่ได้เลือก Report
 * ผ่าน Terminal
 */
/**
 * ช่วงวันที่สำหรับ Export Report
 */
export const datereport = {
  dateset: "25/11/2025",
  dateto: "27/11/2025",
};

/**
 * Path ของไฟล์ Test Data กลาง
 *
 * ทุก Report ใช้ Test Data ไฟล์เดียวกัน
 */
export const testDataPath =
  TEST_DATA_INPUT_PATH;