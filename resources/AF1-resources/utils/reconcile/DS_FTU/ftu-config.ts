/**
 * ftu-config.ts
 * ------------------------------------------------------------------
 * Config และ Constant ของ DS_FTU สำหรับ Script 3 (Reconcile)
 *
 * หน้าที่:
 * - ระบุ Report Code และตำแหน่ง Header
 * - กำหนดชื่อ Header ของ AF1 Report และ Test Data
 * - กำหนดรหัส Leg Type สำหรับรายการ Buy/Sell
 * - กำหนดค่าที่ใช้ตรวจเงื่อนไข DS_FTU
 * - กำหนดข้อความ Remark สำหรับผล Reconcile
 *
 * ไฟล์นี้เก็บเฉพาะ Config และ Constant
 * ไม่มี Logic การตัดสิน Pass/Fail
 * ------------------------------------------------------------------
 */

import { getReportRuntimeConfig } from "../../../config/report-runtime.config";

export const FTU_REPORT_CODE = "DS_FTU";
export const FTU_REPORT_HEADER_ROW = 1;

export const FTU_TEST_DATA_HEADER_ROW =
  getReportRuntimeConfig(FTU_REPORT_CODE).testDataHeaderRowNumber;

export const FTU_THB_CURRENCY_CODE = "THB";
export const FTU_USD_CURRENCY_CODE = "USD";
export const FTU_USD_THRESHOLD = 50_000;
export const FTU_AMOUNT_TOLERANCE = 0.01;
export const FTU_COUNTRY_CODE_LENGTH = 2;
export const FTU_AMOUNT_THRESHOLD_LABEL = "FTU Amount Threshold";

export const FTU_DIRECTIONS = {
  buyForeignCurrency: "BUY_FCY",
  sellForeignCurrency: "SELL_FCY",
  unknown: "UNKNOWN_DIRECTION",
  noThbLeg: "NO_THB_LEG",
} as const;

export type FtuDirection =
  (typeof FTU_DIRECTIONS)[keyof typeof FTU_DIRECTIONS];

export const FTU_LEG_TYPES = {
  buyForeignCurrency: "182001",
  sellForeignCurrency: "182002",
} as const;

export const FTU_REPORT_FIELDS = {
  arrangementNumber: "Arr Number",
  dataSetDate: "Data Set Date",
  inflowPurpose: "Inflow Transaction Purpose",
  outflowPurpose: "Outflow Transaction Purpose",
  currencyId: "Currency Id",
  legType: "Leg Type",
  beneficiaryCountry: "Country Id of Beneficiary Involved Party",
  foreignCurrencyAmount: "Foreign Currency Amount",
} as const;

export const FTU_TEST_DATA_FIELDS = {
  testNo: "Test No.",
  transactionId: "Transaction ID/ Reconcile ID",
  transactionDate: "Txn Date",
  fromCurrency: "From Currency (CCY)",
  toCurrency: "To Currency (CCY)",
  purposeCode: "From BOT Purpose code",
  settledAmount: "Settled Amount (CCY)",
  settledCurrency: "Settled Currency (CCY)",
} as const;

export const FTU_REQUIRED_TEST_DATA_HEADERS = [
  FTU_TEST_DATA_FIELDS.testNo,
  FTU_TEST_DATA_FIELDS.transactionId,
  FTU_TEST_DATA_FIELDS.transactionDate,
  FTU_TEST_DATA_FIELDS.fromCurrency,
  FTU_TEST_DATA_FIELDS.toCurrency,
  FTU_TEST_DATA_FIELDS.purposeCode,
  FTU_TEST_DATA_FIELDS.settledCurrency,
  FTU_TEST_DATA_FIELDS.settledAmount,
] as const;

export const FTU_REMARKS = {
  buyForeignCurrency: "Buy Foreign Currency — ต้องพบใน DS_FTU",
  sellForeignCurrency: "Sell Foreign Currency — ต้องพบใน DS_FTU",
  /** แสดงเมื่อมีอย่างน้อยหนึ่ง Field ที่ต้อง Review */
  pleaseReview: "Please review",
  noThbLegExpectedAbsence: "ธุรกรรมไม่มีขา THB — ไม่ควรพบใน DS_FTU",
  noThbLegUnexpectedPresence: "ธุรกรรมไม่มีขา THB แต่พบใน DS_FTU โดยไม่ควรพบ",
  fallbackDirectionUnavailable:
    "Fallback Matching ทำไม่ได้ เพราะ From/To Currency ไม่ครบ",
  expectedAbsenceVerificationUnavailable:
    "ไม่สามารถยืนยัน Expected Absence ด้วย Fallback ได้ครบถ้วน",
  expectedAbsenceReportNotFound: "ไม่พบรายการใน Report ตามที่คาดหวัง",
  noThbLegReportNotFound: "ไม่พบรายการใน DS_FTU ตามที่คาดหวัง",
} as const;

/**
 * Normalize ที่เป็นกติกาเฉพาะของ DS_FTU
 * ใช้ร่วมกันระหว่าง Matcher และ Analyzer
 */
export const normalizeFtuText = (value: unknown): string =>
  String(value ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();