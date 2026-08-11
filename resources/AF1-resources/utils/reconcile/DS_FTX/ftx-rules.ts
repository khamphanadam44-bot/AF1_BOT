/**
 * ftx-rules.ts
 * ------------------------------------------------------------------
 * Business Rules กลางของ DS_FTX
 *
 * หน้าที่:
 * 1. Normalize ข้อความและวันที่ก่อนเปรียบเทียบ
 * 2. เปรียบเทียบ Core Fields ระหว่าง Test Data กับ Report
 * 3. ตรวจสอบ Exclusion Rules ของ DS_FTX
 * ------------------------------------------------------------------
 */

import {
  DS_FTX_COMPARE_RULES,
  resolveExpectedValue,
} from "./ftx-config";

import type {
  ActualRow,
  CompareResult,
  ExpectedRow,
  TestDataRow,
} from "./ftx-types";

/** Header ที่ใช้ตรวจ From Currency */
const FROM_CURRENCY_HEADER =
  "From Currency (CCY)";

/** Header ที่ใช้ตรวจ Settled Currency */
const SETTLED_CURRENCY_HEADER =
  "Settled Currency (CCY)";

/** Header ที่ใช้ตรวจ Settled Amount */
const SETTLED_AMOUNT_HEADER =
  "Settled Amount (CCY)";

/** Header ที่ใช้ตรวจประเภทลูกค้า */
const CUSTOMER_TYPE_HEADER =
  "From Customer Type Description";

/** สกุลเงินฐาน */
const BASE_CURRENCY = "THB";

/** สกุลเงิน USD */
const USD_CURRENCY = "USD";

/** วงเงินสูงสุดของ Resident Exclusion Rule */
const RESIDENT_AMOUNT_LIMIT_USD = 50000;

/** ผลการตรวจสอบ Exclusion Rule */
export interface FtxExclusionResult {
  isExcluded: boolean;
  remark: string;
}

/**
 * Normalize ข้อความสำหรับ Core Field
 *
 * การทำงาน:
 * - เปลี่ยน null หรือ undefined เป็นข้อความว่าง
 * - ตัดช่องว่างด้านหน้าและด้านหลัง
 * - รวมช่องว่างหลายช่องให้เหลือหนึ่งช่อง
 * - เปลี่ยนเป็นตัวพิมพ์ใหญ่
 */
const normalizeText = (
  value: unknown,
): string => {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
};

/**
 * Normalize วันที่ให้อยู่ในรูปแบบ yyyy-mm-dd
 */
const normalizeDate = (
  value: unknown,
): string => {
  const text =
    String(
      value ?? "",
    ).trim();

  /**
   * รองรับ:
   * - dd/mm/yyyy
   * - dd-mm-yyyy
   */
  const dayFirstMatch =
    text.match(
      /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/,
    );

  if (dayFirstMatch) {
    const [
      ,
      day,
      month,
      year,
    ] = dayFirstMatch;

    return (
      `${year}-` +
      `${month.padStart(2, "0")}-` +
      `${day.padStart(2, "0")}`
    );
  }

  /**
   * รองรับ:
   * - yyyy/mm/dd
   * - yyyy-mm-dd
   * - yyyy-mm-dd ตามด้วยเวลา
   */
  const isoMatch =
    text.match(
      /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/,
    );

  if (isoMatch) {
    const [
      ,
      year,
      month,
      day,
    ] = isoMatch;

    return (
      `${year}-` +
      `${month.padStart(2, "0")}-` +
      `${day.padStart(2, "0")}`
    );
  }

  /**
   * ถ้าไม่ตรง Pattern
   * ให้ลองแปลงด้วย JavaScript Date
   */
  const parsedDate =
    new Date(text);

  if (
    !Number.isNaN(
      parsedDate.getTime(),
    )
  ) {
    const year =
      parsedDate.getFullYear();

    const month =
      String(
        parsedDate.getMonth() + 1,
      ).padStart(
        2,
        "0",
      );

    const day =
      String(
        parsedDate.getDate(),
      ).padStart(
        2,
        "0",
      );

    return `${year}-${month}-${day}`;
  }

  /**
   * ถ้าแปลงเป็นวันที่ไม่ได้
   * ให้เปรียบเทียบเป็นข้อความแทน
   */
  return normalizeText(value);
};

/**
 * เปรียบเทียบ Core Fields ตาม DS_FTX_COMPARE_RULES
 *
 * ปัจจุบันตรวจสอบ:
 * - From Currency (CCY) กับ Buy Currency Id
 * - Settled Currency (CCY) กับ Sell Currency Id
 * - Txn Date กับ Transaction Date
 */
export const compareCoreFields = (
  expectedRow: ExpectedRow,
  actualRow: ActualRow,
): CompareResult[] => {
  return DS_FTX_COMPARE_RULES.map(
    (rule): CompareResult => {
      /**
       * อ่านค่าที่คาดหวังจาก Test Data
       */
      const expectedValue =
        resolveExpectedValue(
          expectedRow,
          rule,
        );

      /**
       * อ่านค่าจริงจาก Report DS_FTX
       */
      const actualValue =
        actualRow.data[
          rule.reportField
        ];

      /**
       * เลือกวิธี Normalize ตามประเภท Rule
       */
      const isMatched =
        rule.compareType === "DATE"
          ? normalizeDate(
              expectedValue,
            ) ===
            normalizeDate(
              actualValue,
            )
          : normalizeText(
              expectedValue,
            ) ===
            normalizeText(
              actualValue,
            );

      return {
        matchingKey:
          expectedRow.matchingKey,

        testScriptNo:
          expectedRow.testScriptNo,

        testDataRowNumber:
          expectedRow.rowNumber,

        reportRowNumber:
          actualRow.rowNumber,

        field:
          rule.reportField,

        expected:
          expectedValue,

        actual:
          actualValue,

        status:
          isMatched
            ? "PASS"
            : "FAIL",

        remark:
          isMatched
            ? ""
            : `${rule.reportField}: Value Mismatch | [Test Data]: ${String(expectedValue ?? "")} | [DS_FTX]: ${String(actualValue ?? "")}`,
      };
    },
  );
};

/**
 * Normalize ข้อความสำหรับ Exclusion Rule
 *
 * Logic เดิมของ Exclusion จะไม่รวมช่องว่างภายใน
 * จึงแยกออกจาก normalizeText() ของ Core Field
 */
const normalizeExclusionText = (
  value: unknown,
): string => {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value)
    .trim()
    .toUpperCase();
};

/**
 * แปลง Amount เป็น Number
 *
 * รองรับ comma คั่นหลักพัน เช่น:
 * "50,000" กลายเป็น 50000
 */
const normalizeAmount = (
  value: unknown,
): number | undefined => {
  if (
    value === null ||
    value === undefined
  ) {
    return undefined;
  }

  const text =
    String(value)
      .replace(/,/g, "")
      .trim();

  if (text === "") {
    return undefined;
  }

  const amount =
    Number(text);

  if (
    Number.isNaN(amount)
  ) {
    return undefined;
  }

  return amount;
};

/**
 * ตรวจว่าประเภทลูกค้าเป็น Resident หรือไม่
 *
 * ค่าที่รองรับ:
 * - R
 * - RESIDENT
 */
const isResidentCustomer = (
  value: unknown,
): boolean => {
  const customerType =
    normalizeExclusionText(value);

  return (
    customerType === "R" ||
    customerType === "RESIDENT"
  );
};

/**
 * ตรวจ Exclusion Rule ของ DS_FTX
 *
 * Rule 1:
 * - From Currency เป็น THB
 *
 * Rule 2 ต้องตรงทุกเงื่อนไข:
 * - From Currency มีค่า
 * - Settled Currency มีค่า
 * - From Currency ตรงกับ Settled Currency
 * - ลูกค้าเป็น Resident
 * - Settled Currency เป็น USD
 * - Settled Amount อยู่ระหว่าง 0 ถึง 50,000 USD
 */
export const evaluateFtxExclusion = (
  testDataRow: TestDataRow,
): FtxExclusionResult => {
  const fromCurrency =
    normalizeExclusionText(
      testDataRow[
        FROM_CURRENCY_HEADER
      ],
    );

  const settlementCurrency =
    normalizeExclusionText(
      testDataRow[
        SETTLED_CURRENCY_HEADER
      ],
    );

  const customerType =
    testDataRow[
      CUSTOMER_TYPE_HEADER
    ];

  const settledAmount =
    normalizeAmount(
      testDataRow[
        SETTLED_AMOUNT_HEADER
      ],
    );

  /**
   * Exclusion Rule 1:
   * From Currency เป็น THB
   */
  if (
    fromCurrency ===
    BASE_CURRENCY
  ) {
    return {
      isExcluded: true,

      remark:
        "Excluded: From Currency (CCY) เป็น THB จึงไม่อยู่ในขอบเขต DS_FTX",
    };
  }

  /**
   * Exclusion Rule 2:
   * Resident ทำรายการ USD ไม่เกิน 50,000
   */
  if (
    fromCurrency !== "" &&
    settlementCurrency !== "" &&
    fromCurrency ===
      settlementCurrency &&
    isResidentCustomer(
      customerType,
    ) &&
    settlementCurrency ===
      USD_CURRENCY &&
    settledAmount !==
      undefined &&
    settledAmount >= 0 &&
    settledAmount <=
      RESIDENT_AMOUNT_LIMIT_USD
  ) {
    return {
      isExcluded: true,

      remark:
        "Excluded: From Currency ตรงกับ Settled Currency, ลูกค้าเป็น Resident และ Settled Amount ไม่เกิน 50,000 USD",
    };
  }

  /**
   * ไม่เข้า Exclusion Rule
   */
  return {
    isExcluded: false,
    remark: "",
  };
};