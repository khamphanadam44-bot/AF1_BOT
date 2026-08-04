/**
 * testdata-helper.ts
 * ------------------------------------------------------------
 * Helper สำหรับอ่าน Test Data Config ตาม Report
 *
 * หน้าที่หลัก:
 * 1. ดึงรายการ Header ที่ต้องตรวจของแต่ละ Report
 * 2. ดึงรายการ Normal Required Header
 * 3. ดึงหมายเลขแถว Header
 * 4. ตรวจจับจำนวน Fee Group จาก Header ในไฟล์จริง
 *
 * DS_PTX และ DS_LTX:
 * - ตรวจจับจำนวน Fee Group จาก Header ใน Test Data
 *
 * DS_FTX และ DS_FTU:
 * - ไม่มีการตรวจ Fee Group
 * - คืนจำนวน Fee Group เป็น 0
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
 * DS_LTX → TESTDATA_CONFIG.DS_LTX
 * DS_FTX → TESTDATA_CONFIG.DS_FTX
 * DS_FTU → TESTDATA_CONFIG.DS_FTU
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
 * ดึง Header ทั้งหมดที่ต้องตรวจของ Report
 *
 * รวม Header จากกลุ่ม:
 * - matchingKey
 * - core
 * - customer
 * - conditional
 * - reference
 * - feeGroup
 *
 * Fee Group:
 * - DS_PTX และ DS_LTX สร้าง Header ตามจำนวนที่ตรวจพบ
 * - DS_FTX และ DS_FTU ไม่มี Fee Group
 */
export function getRequiredTestDataHeaders(
  reportCode: TestDataReportCode,
  actualHeaders: readonly string[],
): string[] {
  const reportConfig =
    getTestDataReportConfig(
      reportCode,
    );

  const groups =
    reportConfig.requiredHeaders;

  const feeTypeCount =
    getFeeTypeCount(
      reportCode,
      actualHeaders,
    );

  /**
   * DS_PTX และ DS_LTX:
   * feeGroup เป็นฟังก์ชันสำหรับสร้าง Header
   *
   * DS_FTX และ DS_FTU:
   * feeGroup เป็น Array ว่าง
   */
  const feeHeaders =
    typeof groups.feeGroup === "function"
      ? groups.feeGroup(
          feeTypeCount,
        )
      : [...groups.feeGroup];

  return [
    ...groups.matchingKey,
    ...groups.core,
    ...groups.customer,
    ...groups.conditional,
    ...groups.reference,
    ...feeHeaders,
  ].map(String);
}

/**
 * ดึงเฉพาะ Header ที่ต้องตรวจแบบ Normal Field
 *
 * ไม่รวม Fee Group เพราะ Fee Group จะถูกตรวจแยก
 * ด้วย validateFeeGroupFields()
 *
 * หากนำ Fee Group มารวมในฟังก์ชันนี้:
 * - Fee อาจถูกตรวจซ้ำ
 * - Fee Group ที่ไม่ได้ใช้อาจถูกมองว่าข้อมูลไม่ครบ
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
 * ปัจจุบันทุก Report ใช้ Header ที่แถว 5
 * แต่ยังแยกไว้ใน Config เพื่อรองรับการเปลี่ยนแปลงในอนาคต
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
 * ปรับรูปแบบ Header ก่อนนำไปตรวจ Fee Group
 *
 * รองรับกรณี:
 * - มีช่องว่างซ้ำกัน
 * - มีอักขระช่องว่างพิเศษ
 * - มีอักขระที่มองไม่เห็น
 * - มีช่องว่างด้านหน้าและด้านหลัง
 */
const normalizeFeeHeader = (
  header: unknown,
): string => {
  return String(header ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim();
};

/**
 * รูปแบบ Header ที่สามารถใช้ตรวจหมายเลข Fee Group
 *
 * ตัวอย่าง Header ที่รองรับ:
 * - Fee Type 1
 * - Fee Charge Type 1
 * - Fee Charge Account No. Type 1
 * - Fee Charge Account Number Type 1
 * - Fee Amount Type 1
 * - Fee Amount 2
 * - Charge Account Type 1
 * - Bank Code Type 1
 * - Branch Code Type 1
 * - Fee Currency Type 1
 * - Fee Currency 2
 * - Fee Timing Type 1
 */
const FEE_GROUP_HEADER_PATTERNS: readonly RegExp[] = [
  /^fee type (\d+)$/i,
  /^fee charge type (\d+)$/i,
  /^fee charge account (?:no\.?|number) type (\d+)$/i,
  /^fee amount(?: type)? (\d+)$/i,
  /^charge account type (\d+)$/i,
  /^bank code type (\d+)$/i,
  /^branch code type (\d+)$/i,
  /^fee currency(?: type)? (\d+)$/i,
  /^fee timing type (\d+)$/i,
];

/**
 * ตรวจหมายเลข Fee Group สูงสุดจาก Header ที่พบจริง
 *
 * ตัวอย่าง:
 *
 * หากพบ:
 * - Fee Type 1
 * - Fee Type 2
 * - Fee Type 3
 *
 * จะคืนค่า:
 * 3
 *
 * หากไม่พบ Header ของ Fee Group:
 * จะคืนค่า:
 * 0
 *
 * หมายเหตุ:
 * ใช้หมายเลขสูงสุดแทนการนับจำนวน Header
 * เพื่อให้ระบบยังตรวจพบกรณีลำดับ Fee Group ขาดหาย
 *
 * ตัวอย่าง:
 * พบ Fee Type 1 และ Fee Type 3 แต่ไม่มี Fee Type 2
 * ระบบจะคืนค่า 3 เพื่อให้ขั้นตอนตรวจ Header
 * สามารถรายงานว่า Fee Group 2 หายได้
 */
export function detectFeeTypeCount(
  actualHeaders: readonly string[],
): number {
  let highestFeeIndex = 0;

  for (const header of actualHeaders) {
    const normalizedHeader =
      normalizeFeeHeader(header);

    for (
      const pattern
      of FEE_GROUP_HEADER_PATTERNS
    ) {
      const matched =
        normalizedHeader.match(pattern);

      if (!matched) {
        continue;
      }

      const feeIndex =
        Number(matched[1]);

      if (
        Number.isInteger(feeIndex) &&
        feeIndex > highestFeeIndex
      ) {
        highestFeeIndex =
          feeIndex;
      }

      /**
       * Header หนึ่งช่องตรงกับ Pattern เดียวก็เพียงพอ
       * จึงไม่ต้องตรวจ Pattern ที่เหลือ
       */
      break;
    }
  }

  return highestFeeIndex;
}

/**
 * คืนจำนวน Fee Group ที่ต้องตรวจของแต่ละ Report
 *
 * DS_PTX และ DS_LTX:
 * - ตรวจจำนวนจาก Header ใน Test Data จริง
 *
 * DS_FTX และ DS_FTU:
 * - ไม่มี Fee Group
 * - คืนค่า 0 เสมอ
 */
export function getFeeTypeCount(
  reportCode: TestDataReportCode,
  actualHeaders: readonly string[],
): number {
  const reportConfig =
    getTestDataReportConfig(
      reportCode,
    );

  /**
   * ถ้า feeGroup ไม่ใช่ฟังก์ชัน
   * หมายความว่า Report นี้ไม่ได้ตรวจ Fee Group
   *
   * ปัจจุบันคือ:
   * - DS_FTX
   * - DS_FTU
   */
  if (
    typeof reportConfig.requiredHeaders.feeGroup !==
    "function"
  ) {
    return 0;
  }

  /**
   * DS_PTX และ DS_LTX
   * ตรวจจำนวน Fee Group จาก Header ที่พบในไฟล์จริง
   */
  return detectFeeTypeCount(
    actualHeaders,
  );
}