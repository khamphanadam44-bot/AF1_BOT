/**
 * ftu-config.ts
 * ------------------------------------------------------------------
 * Config และ Constant ของ DS_FTU สำหรับ Script 3 (Reconcile)
 *
 * เก็บเฉพาะข้อมูล ไม่มี Logic ใดๆ — Service อ่านข้อมูลตาม Header และ
 * Page เขียนข้อมูลตาม Target Column ที่กำหนดในไฟล์นี้
 * ------------------------------------------------------------------
 */

import { getReportRuntimeConfig } from "../../../config/report-runtime.config";

export const FTU_REPORT_CODE = "DS_FTU";
export const FTU_REPORT_HEADER_ROW = 1;

/**
 * Bug fix (Code Review): เดิม Hardcode = 5 ซ้ำกับ report-runtime.config.ts และ
 * ltx-config.ts — เปลี่ยนมา Derive จาก report-runtime.config.ts (Single Source
 * of Truth เดียวกับที่ ltx-config.ts ใช้) แทน
 */
export const FTU_TEST_DATA_HEADER_ROW =
  getReportRuntimeConfig(FTU_REPORT_CODE).testDataHeaderRowNumber;

export const FTU_BASE_CURRENCY = "THB";
export const FTU_USD_THRESHOLD = 50_000;
export const FTU_AMOUNT_TOLERANCE = 0.01;

export const FTU_LEG_TYPES = {
  buyForeignCurrency: "182001",
  sellForeignCurrency: "182002",
} as const;

export const FTU_COUNTRY_ID_BY_CURRENCY: Readonly<Record<string, string>> = {
  USD: "US",
  EUR: "DE",
  JPY: "JP",
  GBP: "GB",
  SGD: "SG",
  AUD : "AU",
  // mapping ตาม Requirement 3.7.3.2
};

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

export const FTU_TEST_FIELDS = {
  testNo: "Test No.",
  transactionId: "Transaction ID/ Reconcile ID",
  transactionDate: "Txn Date",
  fromCurrency: "From Currency (CCY)",
  toCurrency: "To Currency (CCY)",
  purposeCode: "From BOT Purpose code",
  settledAmount: "Settled Amount (CCY)",
  settledCurrency: "Settled Currency (CCY)",
} as const;

export const FTU_REMARKS = {
  buyForeignCurrency: "Buy Foreign Currency — ต้องพบใน DS_FTU",
  sellForeignCurrency: "Sell Foreign Currency — ต้องพบใน DS_FTU",

  noThbLegExpectedAbsence:
    "ธุรกรรมไม่มีขา THB — ไม่ควรพบใน DS_FTU",
  noThbLegUnexpectedPresence:
    "ธุรกรรมไม่มีขา THB แต่พบใน DS_FTU โดยไม่ควรพบ",

  thresholdExpectedAbsence: (usdAmount: number) =>
    `มูลค่า ${usdAmount.toLocaleString("en-US", { maximumFractionDigits: 2 })} USD ` +
    `>= 50,000 USD — ไม่ควรพบใน DS_FTU`,
  thresholdUnexpectedPresence: (usdAmount: number) =>
    `มูลค่า ${usdAmount.toLocaleString("en-US", { maximumFractionDigits: 2 })} USD ` +
    `>= 50,000 USD แต่พบใน DS_FTU โดยไม่ควรพบ`,

  ftxExceptionExpectedAbsence: (reason: string) =>
    `เข้าเงื่อนไข FTX Exception (${reason}) — ไม่ควรพบใน DS_FTU`,
  ftxExceptionUnexpectedPresence: (reason: string) =>
    `เข้าเงื่อนไข FTX Exception (${reason}) แต่พบใน DS_FTU โดยไม่ควรพบ`,
};
