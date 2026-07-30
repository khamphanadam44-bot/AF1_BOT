/**
 * testdata-helper.ts
 * ------------------------------------------------------------
 * Helper สำหรับอ่าน Test Data Config ตาม Report
 *
 * แยกรายการ Header ออกเป็น:
 *
 * 1. Header ทั้งหมด
 *    ใช้ตรวจว่า Header มีอยู่ในไฟล์ครบหรือไม่
 *
 * 2. Normal Required Header
 *    ใช้ตรวจข้อมูลในแต่ละแถว
 *    โดยไม่รวม Fee Group
 * ------------------------------------------------------------
 */

import {
  TESTDATA_CONFIG,
} from "./testdata-config";

import type {
  TestDataReportCode,
} from "./testdata-config";

/**
 * ตรวจสอบและคืน Config ของ Report
 *
 * ตัวอย่าง:
 * DS_PTX → TESTDATA_CONFIG.DS_PTX
 * DS_FTX → TESTDATA_CONFIG.DS_FTX
 */
const getTestDataReportConfig = (
  reportCode: TestDataReportCode,
) => {
  const reportConfig =
    TESTDATA_CONFIG[reportCode];

  if (!reportConfig) {
    throw new Error(
      `Test Data Config not found for report: ${reportCode}`,
    );
  }

  return reportConfig;
};

/**
 * ดึง Header ทั้งหมดของ Report
 *
 * ใช้สำหรับตรวจว่า Header ในไฟล์ Test Data
 * มีอยู่ครบตาม Config หรือไม่
 *
 * รวมกลุ่ม:
 * - matchingKey
 * - core
 * - customer
 * - conditional
 * - reference
 * - feeGroup
 */
export function getRequiredTestDataHeaders(
  reportCode: TestDataReportCode,
): string[] {
  const reportConfig =
    getTestDataReportConfig(
      reportCode,
    );

  const groups =
    reportConfig.requiredHeaders;

  return [
    ...groups.matchingKey,
    ...groups.core,
    ...groups.customer,
    ...groups.conditional,
    ...groups.reference,
    ...groups.feeGroup,
  ].map(String);
}

/**
 * ดึงเฉพาะ Header ที่ต้องตรวจแบบ Normal Field
 *
 * ไม่รวม feeGroup เพราะ Fee Group
 * ต้องตรวจด้วย validateFeeGroupFields()
 *
 * หากรวม feeGroup ในฟังก์ชันนี้:
 * - Fee จะถูกตรวจซ้ำ
 * - Fee ชุดที่ไม่ได้ใช้อาจถูกมองว่าเป็นข้อมูลไม่ครบ
 */
export function getNormalRequiredTestDataHeaders(
  reportCode: TestDataReportCode,
): string[] {
  const reportConfig =
    getTestDataReportConfig(
      reportCode,
    );

  const groups =
    reportConfig.requiredHeaders;

  return [
    ...groups.matchingKey,
    ...groups.core,
    ...groups.customer,
    ...groups.conditional,
    ...groups.reference,
  ].map(String);
}


/**
 * ดึงหมายเลขแถว Header ของ Test Data
 *
 * ตอนนี้:
 * DS_PTX → แถวที่ 5
 * DS_FTX → แถวที่ 5
 */
export function getTestDataHeaderRowNumber(
  reportCode: TestDataReportCode,
): number {
  const reportConfig =
    getTestDataReportConfig(
      reportCode,
    );

  return reportConfig.headerRowNumber;
}

/**
 * นับจำนวน Fee Group จากรายการ Header ใน Config
 *
 * นับเฉพาะ Header รูปแบบ "Fee Type N"
 * เพราะ Fee Group หนึ่งชุดต้องมี Fee Type หนึ่งช่องเสมอ
 */
const countFeeTypes = (
  feeHeaders: readonly string[],
): number => {
  return feeHeaders.filter(
    (header) =>
      /^Fee Type \d+$/i.test(
        String(header).trim(),
      ),
  ).length;
};

/**
 * คืนจำนวน Fee Group ที่ต้องใช้
 *
 * เมื่อส่ง reportCode:
 * - DS_LTX → 5
 * - DS_PTX → 2
 * - Report ที่ไม่มี Fee Group → 0
 *
 * เมื่อไม่ส่ง reportCode:
 * คืนจำนวนสูงสุดของทุก Report เพื่อใช้สร้าง Header Alias กลาง
 */
export function getFeeTypeCount(
  reportCode?: TestDataReportCode,
): number {
  if (reportCode) {
    return countFeeTypes(
      TESTDATA_CONFIG[
        reportCode
      ].requiredHeaders.feeGroup,
    );
  }

  return Math.max(
    0,
    ...Object.values(
      TESTDATA_CONFIG,
    ).map(
      (config) =>
        countFeeTypes(
          config.requiredHeaders.feeGroup,
        ),
    ),
  );
}
