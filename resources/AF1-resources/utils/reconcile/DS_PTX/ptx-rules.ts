/**
 * ============================================================================
 * ไฟล์: ptx-rules.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้เป็นส่วนที่ “เปรียบเทียบค่าจริง” ระหว่าง Test Data กับ AF1 Report
 *
 * เมื่อระบบจับคู่แถวได้แล้ว ไฟล์นี้จะตรวจทีละ Field ว่าค่าที่คาดหวัง
 * และค่าที่อยู่ใน Report ตรงกันหรือไม่ แล้วสร้างผลเป็น PASS หรือ FAIL
 *
 * กลุ่มข้อมูลที่ตรวจ
 * -----------------
 * 1. Core Fields
 *    เป็นข้อมูลหลักที่ต้องตรวจ เช่น วันที่ สกุลเงิน ประเภท Transaction
 *    และจำนวนเงิน
 *
 * 2. Conditional / Customer Fields
 *    เป็นข้อมูลที่ตรวจตามเงื่อนไข เช่น Cust Code, Cust Name
 *    และข้อมูลของ Involved Party
 *
 * กติกาสำคัญ
 * -----------
 * - ข้อความจะตัดช่องว่างและไม่สนใจตัวพิมพ์เล็ก/ใหญ่
 * - ตัวเลขต้องแปลงเป็น Number ได้ก่อนจึงจะเปรียบเทียบ
 * - จำนวนเงินยอมให้คลาดเคลื่อนได้ตาม tolerance ที่กำหนดใน Config
 * - วันที่ DS_PTX มี Logic พิเศษ โดยตรวจทั้ง Data Set Date และวันที่
 *   ที่ฝังอยู่ใน Reference Transaction Number
 * - Conditional Field ที่ Actual ว่าง จะถือว่าไม่ต้องตรวจและให้ PASS
 * - ถ้า Actual มีค่า แต่ Expected ไม่มีค่า จะให้ FAIL
 *
 * ผลลัพธ์จากไฟล์นี้ยังไม่ใช่ Excel แต่เป็นข้อมูลกลางที่ระบุว่า
 * Field ใด PASS, FAIL และมีสาเหตุอะไร เพื่อนำไปเขียนลง Excel ภายหลัง
 * ============================================================================
 */


/**
 * ส่วน import ด้านล่าง คือการนำเครื่องมือหรือโครงสร้างข้อมูล
 * จากไฟล์อื่นมาใช้ในไฟล์นี้ เปรียบเหมือนการหยิบอุปกรณ์ที่เตรียมไว้แล้ว
 * มาใช้งาน โดยไม่ต้องเขียนทุกอย่างซ้ำใหม่
 */

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
  TestDataRow,
} from "./ptx-types";

import {
  getMappingCoreHeaders,
  getMappingCustomerHeaders,
} from "../../../config/mapping-helper";


import {
  CompareType,
  getCompareRule,
  resolveExpectedValue,
} from "./ptx-config";

/**
 * Remark สำหรับรายการที่ไม่ควรแสดงใน DS_PTX
 */
export const RESIDENT_THB_TO_FCD_REMARK =
  "กรณี ลูกค้า Resident โอน (From) THB บาท " +
  "ออกจากบัญชีไปยัง FCD จะต้องไม่แสดงรายการใน PTX";

/**
 * ปรับข้อความให้อยู่ในรูปแบบเดียวกัน
 * ก่อนนำไปตรวจ Business Rule
 */
const normalizeRuleValue = (
  value: unknown,
): string => {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
};

/**
 * ตรวจ Exclusion Rule ของ DS_PTX
 *
 * เงื่อนไข:
 * 1. ลูกค้าเป็น Resident
 * 2. From Currency เป็น THB
 * 3. บัญชีปลายทางเป็น FCD
 *
 * หาก To Account Type ว่าง
 * จะใช้ To Currency เป็นข้อมูลสำรอง
 */
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

  const isResident =
    residentStatus === "RESIDENT" ||
    residentStatus.startsWith(
      "RESIDENT (",
    );

  const isFromThb =
    fromCurrency === "THB";

  const hasExplicitFcdAccountType =
    toAccountType === "FCD" ||
    toAccountType.startsWith(
      "FCD ",
    ) ||
    toAccountType.startsWith(
      "FCD(",
    );

  const hasForeignToCurrency =
    toCurrency !== "" &&
    toCurrency !== "THB";

  const isToFcd =
    hasExplicitFcdAccountType ||
    (
      toAccountType === "" &&
      hasForeignToCurrency
    );

  return (
    isResident &&
    isFromThb &&
    isToFcd
  );
};

/**
 * ============================================================================
 * แปลงค่าให้อยู่ในรูปตัวเลข
 * ============================================================================
 */
const toNumber = (
  value: unknown,
): number | undefined => {

  if (
    value === undefined ||
    value === null ||
    value === ""
  ) {
    return undefined;
  }

  const result = Number(value);

  return Number.isNaN(result)
    ? undefined
    : result;

};

/**
 * ============================================================================
 * เปรียบเทียบข้อความ
 * ============================================================================
 */
const compareText = (
  expected: unknown,
  actual: unknown,
): boolean => {

  return (
    String(expected ?? "")
      .trim()
      .toUpperCase()
    ===
    String(actual ?? "")
      .trim()
      .toUpperCase()
  );

};

/**
 * ============================================================================
 * เปรียบเทียบวันที่
 * ============================================================================
 *
 * ตอนนี้ใช้ Trim ก่อน
 * อนาคตจะรองรับ yyyy-MM-dd / dd/MM/yyyy
 */
const compareDate = (
  expected: unknown,
  actual: unknown,
): boolean => {

  return compareText(
    expected,
    actual,
  );

};

/**
 * ============================================================================
 * เปรียบเทียบตัวเลข
 * ============================================================================
 */
const compareNumber = (
  expected: unknown,
  actual: unknown,
): boolean => {

  const expectedNumber =
    toNumber(expected);

  const actualNumber =
    toNumber(actual);

  if (
    expectedNumber === undefined ||
    actualNumber === undefined
  ) {
    return false;
  }

  return (
    expectedNumber === actualNumber
  );

};

/**
 * ============================================================================
 * เปรียบเทียบจำนวนเงิน
 * ============================================================================
 */
const compareAmount = (
  expected: unknown,
  actual: unknown,
  tolerance = 0,
): boolean => {

  const expectedNumber =
    toNumber(expected);

  const actualNumber =
    toNumber(actual);

  if (
    expectedNumber === undefined ||
    actualNumber === undefined
  ) {
    return false;
  }

  return (
    Math.abs(
      expectedNumber -
      actualNumber,
    ) <= tolerance
  );

};

/**
 * ============================================================================
 * เลือกวิธีเปรียบเทียบค่าตามประเภทข้อมูล
 * ============================================================================
 */
const compareValue = (

  expected: unknown,

  actual: unknown,

  compareType: CompareType = "TEXT",

  tolerance = 0,

): boolean => {

  /**
   * Empty
   */
  if (

    (expected === "" ||
      expected === undefined ||
      expected === null)

    &&

    (actual === "" ||
      actual === undefined ||
      actual === null)

  ) {

    return true;

  }

  switch (compareType) {

    case "TEXT":

      return compareText(
        expected,
        actual,
      );

    case "DATE":

      return compareDate(
        expected,
        actual,
      );

    case "NUMBER":

      return compareNumber(
        expected,
        actual,
      );

    case "AMOUNT":

      return compareAmount(
        expected,
        actual,
        tolerance,
      );

    default:

      return compareText(
        expected,
        actual,
      );

  }

};

/**
 * ============================================================================
 * ฟังก์ชันช่วยเหลือที่ใช้ร่วมกัน
 * ============================================================================
 */

/**
 * ตรวจสอบค่าว่าง
 */
const isBlank = (
  value: unknown,
): boolean => {

  return (
    value === undefined ||
    value === null ||
    String(value).trim() === ""
  );

};

/**
 * สร้าง CompareResult
 */
const createResult = (
  expected: ExpectedRow,
  actual: ActualRow,
  field: string,
  expectedValue: unknown,
  actualValue: unknown,
  status: "PASS" | "FAIL",
  remark: string,
): CompareResult => {

  return {

    matchingKey:
      expected.matchingKey,

    testDataRowNumber:
      expected.rowNumber,

    reportRowNumber:
      actual.rowNumber,

    field,

    expected:
      expectedValue,

    actual:
      actualValue,

    status,

    remark,

  };

};

/**
 * เติมเลข 0 ด้านหน้า
 *
 * ตัวอย่าง:
 * 5 → 05
 */
const padTwoDigits = (
  value: number,
): string => {

  return String(
    value,
  ).padStart(
    2,
    "0",
  );

};

/**
 * ตรวจสอบว่าวันที่มีอยู่จริงหรือไม่
 */
const isValidDateParts = (
  year: number,
  month: number,
  day: number,
): boolean => {

  const date =
    new Date(
      year,
      month - 1,
      day,
    );

  return (
    date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day
  );

};

/**
 * สร้างวันที่มาตรฐาน
 *
 * รูปแบบ:
 * YYYY-MM-DD
 */
const createDateKey = (
  year: number,
  month: number,
  day: number,
): string | undefined => {

  if (
    !isValidDateParts(
      year,
      month,
      day,
    )
  ) {

    return undefined;

  }

  return [
    String(year),
    padTwoDigits(month),
    padTwoDigits(day),
  ].join(
    "-",
  );

};

/**
 * Normalize วันที่ให้อยู่ในรูปแบบ:
 *
 * YYYY-MM-DD
 *
 * รองรับ:
 * - ExcelJS Date Object
 * - DD/MM/YYYY
 * - DD-MM-YYYY
 * - YYYY-MM-DD
 * - YYYY/MM/DD
 * - Excel Serial Number
 */
const normalizeDate = (
  value: unknown,
): string | undefined => {

  if (
    value === undefined ||
    value === null
  ) {

    return undefined;

  }

  /**
   * กรณี ExcelJS คืน Date Object
   */
  if (
    value instanceof Date
  ) {

    if (
      Number.isNaN(
        value.getTime(),
      )
    ) {

      return undefined;

    }

    return createDateKey(

      value.getFullYear(),

      value.getMonth() + 1,

      value.getDate(),

    );

  }

  /**
   * กรณี Excel Serial Number
   */
  if (
    typeof value === "number"
  ) {

    const milliseconds =
      Math.round(
        (
          value - 25569
        ) *
        86400 *
        1000,
      );

    const date =
      new Date(
        milliseconds,
      );

    if (
      Number.isNaN(
        date.getTime(),
      )
    ) {

      return undefined;

    }

    return createDateKey(

      date.getUTCFullYear(),

      date.getUTCMonth() + 1,

      date.getUTCDate(),

    );

  }

  const text =
    String(
      value,
    ).trim();

  if (
    text === ""
  ) {

    return undefined;

  }

  /**
   * DD/MM/YYYY
   * DD-MM-YYYY
   */
  const dayFirstMatch =
    text.match(
      /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/,
    );

  if (
    dayFirstMatch
  ) {

    const day =
      Number(
        dayFirstMatch[1],
      );

    const month =
      Number(
        dayFirstMatch[2],
      );

    const year =
      Number(
        dayFirstMatch[3],
      );

    return createDateKey(
      year,
      month,
      day,
    );

  }

  /**
   * YYYY-MM-DD
   * YYYY/MM/DD
   */
  const yearFirstMatch =
    text.match(
      /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/,
    );

  if (
    yearFirstMatch
  ) {

    const year =
      Number(
        yearFirstMatch[1],
      );

    const month =
      Number(
        yearFirstMatch[2],
      );

    const day =
      Number(
        yearFirstMatch[3],
      );

    return createDateKey(
      year,
      month,
      day,
    );

  }

  return undefined;

};

/**
 * ดึงวันที่จาก Reference Transaction Number
 *
 * ตัวอย่าง:
 *
 * KBO3042511260000000081670_01_GPMH_KBO_RM
 *
 * ตำแหน่งที่ 7-12:
 * 251126
 *
 * ผลลัพธ์:
 * 2025-11-26
 */
const extractDateFromReferenceTransactionNumber = (
  referenceTransactionNumber: unknown,
): string | undefined => {

  if (
    isBlank(
      referenceTransactionNumber,
    )
  ) {

    return undefined;

  }

  const reference =
    String(
      referenceTransactionNumber,
    ).trim();

  if (
    reference.length < 12
  ) {

    return undefined;

  }

  /**
   * ตำแหน่งที่ 7-12 แบบนับจาก 1
   *
   * JavaScript:
   * slice(6, 12)
   */
  const rawDate =
    reference.slice(
      6,
      12,
    );

  if (
    !/^\d{6}$/.test(
      rawDate,
    )
  ) {

    return undefined;

  }

  const year =
    2000 +
    Number(
      rawDate.slice(
        0,
        2,
      ),
    );

  const month =
    Number(
      rawDate.slice(
        2,
        4,
      ),
    );

  const day =
    Number(
      rawDate.slice(
        4,
        6,
      ),
    );

  return createDateKey(
    year,
    month,
    day,
  );

};

/**
 * Compare วันที่ของ DS_PTX
 *
 * จุดพิจารณาที่ 1:
 * Txn Date = Data Set Date
 *
 * จุดพิจารณาที่ 2:
 * หากไม่เท่ากัน ให้ตรวจ
 * Txn Date = วันที่จาก Reference Transaction Number
 */
const compareDsPtxTransactionDate = (
  expectedTxnDate: unknown,
  actual: ActualRow,
): {
  pass: boolean;
  remark: string;
} => {

  const normalizedTxnDate =
    normalizeDate(
      expectedTxnDate,
    );

  const normalizedDataSetDate =
    normalizeDate(
      actual.data[
        "Data Set Date"
      ],
    );

  const referenceTransactionNumber =
    actual.data[
      "Reference Transaction Number"
    ] ??
    actual.matchingKey;

  const referenceTxnDate =
    extractDateFromReferenceTransactionNumber(
      referenceTransactionNumber,
    );

  if (
    !normalizedTxnDate
  ) {

    return {

      pass:
        false,

      remark:
        "Invalid Txn Date",

    };

  }

  /**
   * จุดพิจารณาที่ 1
   */
  if (
    normalizedDataSetDate &&
    normalizedTxnDate ===
      normalizedDataSetDate
  ) {

    return {

      pass:
        true,

      remark:
        "",

    };

  }

  /**
   * จุดพิจารณาที่ 2
   */
  if (
    referenceTxnDate &&
    normalizedTxnDate ===
      referenceTxnDate
  ) {

    return {

      pass:
        true,

      remark:
        "",

    };

  }

  return {

    pass:
      false,

    remark:
      "Txn Date does not match Data Set Date or Reference Transaction Date",

  };

};

/**
 * ============================================================================
 * เปรียบเทียบข้อมูลหลัก (Core Fields)
 * ============================================================================
 */
export const compareCoreFields = (
  reportName: string,
  expected: ExpectedRow,
  actual: ActualRow,
): CompareResult[] => {

  const results: CompareResult[] = [];

  const headers =
    getMappingCoreHeaders(
      reportName,
    );

  for (
    const header of headers
  ) {

    const expectedValue =
      resolveExpectedValue(
        expected,
        header,
        reportName,
      );

    const actualValue =
      actual.data[
        header
      ];

    const rule =
      getCompareRule(
        reportName,
        header,
      );

    if (
      !rule
    ) {

      results.push(

        createResult(
          expected,
          actual,
          header,
          expectedValue,
          actualValue,
          "FAIL",
          "Compare Rule Not Found",
        ),

      );

      continue;

    }

    /**
     * DS_PTX Date Logic
     */
    if (
      reportName === "DS_PTX" &&
      header ===
        "Receive Payment Transaction Date"
    ) {

      const dateResult =
        compareDsPtxTransactionDate(
          expectedValue,
          actual,
        );

      results.push(

        createResult(
          expected,
          actual,
          header,
          expectedValue,
          actualValue,
          dateResult.pass
            ? "PASS"
            : "FAIL",
          dateResult.remark,
        ),

      );

      continue;

    }

    const pass =
      compareValue(

        expectedValue,

        actualValue,

        rule.compareType,

        rule.tolerance,

      );

    results.push(

      createResult(
        expected,
        actual,
        header,
        expectedValue,
        actualValue,
        pass
          ? "PASS"
          : "FAIL",
        pass
          ? ""
          : "Value Mismatch",
      ),

    );

  }

  return results;

};

/**
 * ============================================================================
 * เปรียบเทียบข้อมูลตามเงื่อนไขและข้อมูลลูกค้า
 * ============================================================================
 */
export const compareCustomerFields = (
  reportName: string,
  expected: ExpectedRow,
  actual: ActualRow,
): CompareResult[] => {

  const results: CompareResult[] = [];

  const headers =
    getMappingCustomerHeaders(
      reportName,
    );

  for (
    const header of headers
  ) {

    const expectedValue =
      resolveExpectedValue(
        expected,
        header,
        reportName,
      );

    const actualValue =
      actual.data[
        header
      ];

    const rule =
      getCompareRule(
        reportName,
        header,
      );

    if (
      !rule
    ) {

      results.push(

        createResult(
          expected,
          actual,
          header,
          expectedValue,
          actualValue,
          "FAIL",
          "Compare Rule Not Found",
        ),

      );

      continue;

    }

    /**
     * Actual Report ไม่มีค่า
     *
     * Conditional Field:
     * ถ้า Report ไม่ส่งค่า ไม่ใช้ตัดสิน FAIL
     */
    if (
      isBlank(
        actualValue,
      )
    ) {

      results.push(

        createResult(
          expected,
          actual,
          header,
          expectedValue,
          actualValue,
          "PASS",
          "Skipped (Actual Value Is Blank)",
        ),

      );

      continue;

    }

    /**
     * Actual มีค่า แต่ Expected ไม่มีค่า
     */
    if (
      isBlank(
        expectedValue,
      )
    ) {

      results.push(

        createResult(
          expected,
          actual,
          header,
          expectedValue,
          actualValue,
          "FAIL",
          "Expected Value Is Blank",
        ),

      );

      continue;

    }

    const pass =
      compareValue(

        expectedValue,

        actualValue,

        rule.compareType,

        rule.tolerance,

      );

    results.push(

      createResult(
        expected,
        actual,
        header,
        expectedValue,
        actualValue,
        pass
          ? "PASS"
          : "FAIL",
        pass
          ? ""
          : "Value Mismatch",
      ),

    );

  }

  return results;

};
