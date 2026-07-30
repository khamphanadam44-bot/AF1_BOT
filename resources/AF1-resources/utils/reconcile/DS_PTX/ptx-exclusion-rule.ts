/**
 * ======================================================
 * ไฟล์: ptx-exclusion-rule.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * ตรวจกรณีที่รายการไม่ควรแสดงใน DS_PTX ตามเงื่อนไข Resident, THB และบัญชีปลายทางประเภท FCD
 * ฟังก์ชันจะคืนค่า true ต่อเมื่อข้อมูลเข้าเงื่อนไขที่กำหนดครบทุกข้อเท่านั้น
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import {
  TestDataRow,
} from "./compare-types";

export const RESIDENT_THB_TO_FCD_REMARK =
  "กรณี ลูกค้า Resident โอน (From) THB บาท ออกจากบัญชีไปยัง FCD จะต้องไม่แสดงรายการใน PTX";

const normalizeRuleValue = (
  value: unknown,
): string => {

  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();

};

export const isResidentThbToFcdExclusionCase = (
  rowData: TestDataRow,
): boolean => {

  const residentStatus =
    normalizeRuleValue(
      rowData[
        "From Customer (Resident/Non Resident)"
      ],
    );

  const fromCurrency =
    normalizeRuleValue(
      rowData[
        "From Currency (CCY)"
      ],
    );

  const toAccountType =
    normalizeRuleValue(
      rowData[
        "To Account Type (Beneficiary)"
      ],
    );

  const toCurrency =
    normalizeRuleValue(
      rowData[
        "To Currency (CCY)"
      ],
    );

  /**
   * รองรับทั้ง:
   * Resident
   * Resident (0600)
   */
  const isResident =
    residentStatus === "RESIDENT" ||
    residentStatus.startsWith(
      "RESIDENT (",
    );

  const isFromTHB =
    fromCurrency === "THB";

  /**
   * กรณีระบุ Account Type เป็น FCD โดยตรง
   */
  const hasExplicitFcdAccountType =
    toAccountType === "FCD" ||
    toAccountType.startsWith(
      "FCD ",
    ) ||
    toAccountType.startsWith(
      "FCD(",
    );

  /**
   * หาก To Account Type ไม่มีข้อมูล
   * ใช้ To Currency จาก Test Data เป็นตัวสำรอง
   *
   * Currency ที่ไม่ใช่ THB ถือเป็นปลายทางสกุลต่างประเทศ
   */
  const hasForeignToCurrency =
    toCurrency !== "" &&
    toCurrency !== "THB";

  const isToFCD =
    hasExplicitFcdAccountType ||
    (
      toAccountType === "" &&
      hasForeignToCurrency
    );

  return (
    isResident &&
    isFromTHB &&
    isToFCD
  );

};