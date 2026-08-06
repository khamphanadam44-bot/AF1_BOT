/**
 * setting.ts
 * ค่าตั้งต้นสำหรับระบบ UAT
 * ------------------------------------------------------------
 */


import {
  getTestDataInputDir,
} from "../../config/paths.config";

import {
  getSingleExcelFile,
} from "../../utils/file-system.util";

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
 * ค้นหา Test Data ตามชื่อ Report ที่กำลัง Run
 *
 * ขั้นตอน:
 * 1. อ่าน AF1_SHAREPATH จากไฟล์ .env
 * 2. สร้าง Path โฟลเดอร์ตามชื่อ Report
 * 3. ค้นหาไฟล์ Excel เพียง 1 ไฟล์ในโฟลเดอร์
 *
 * ตัวอย่าง:
 * report=DS_PTX
 *
 * ระบบจะค้นหาภายใน:
 * AF1_SHAREPATH/af1_test_data/DS_PTX
 *
 * ชื่อไฟล์ Excel เป็นชื่ออะไรก็ได้
 */
export const getTestDataPath = (
  reportCode: string,
): string => {
  /**
   * อ่าน Share Path จากไฟล์ .env
   */
  const sharePath =
    requireEnv(
      "AF1_SHAREPATH",
    );

  /**
   * สร้าง Path ของโฟลเดอร์ Test Data
   * โดยใช้ชื่อ Report ที่ได้รับเข้ามา
   */
  const testDataDirectory =
    getTestDataInputDir(
      sharePath,
      reportCode,
    );

  /**
   * ค้นหา Excel เพียง 1 ไฟล์
   * ภายในโฟลเดอร์ของ Report
   */
  return getSingleExcelFile(
    testDataDirectory,
  );
};