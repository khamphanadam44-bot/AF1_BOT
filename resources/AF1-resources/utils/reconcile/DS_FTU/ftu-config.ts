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
export const FTU_USD_THRESHOLD = 50_000;

export const FTU_LEG_TYPES = {
  buyForeignCurrency: "182001",
  sellForeignCurrency: "182002",
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
};

export const FTU_TEST_DATA_FIELDS = {
  testNo: "Test No.",
  transactionId: "Transaction ID/ Reconcile ID",
  transactionDate: "Txn Date",
  fromCurrency: "From Currency (CCY)",
  toCurrency: "To Currency (CCY)",
  purposeCode: "From BOT Purpose code",
  settledAmount: "Settled Amount (CCY)",
  settledCurrency: "Settled Currency (CCY)",
};

export const FTU_REMARKS = {
  buyForeignCurrency: "Buy Foreign Currency — ต้องพบใน DS_FTU",
  sellForeignCurrency: "Sell Foreign Currency — ต้องพบใน DS_FTU",
   /** แสดงเมื่อมีอย่างน้อยหนึ่ง Field ที่ต้อง Review */
  pleaseReview: "Please review",
  noThbLegExpectedAbsence: "ธุรกรรมไม่มีขา THB — ไม่ควรพบใน DS_FTU",
  noThbLegUnexpectedPresence: "ธุรกรรมไม่มีขา THB แต่พบใน DS_FTU โดยไม่ควรพบ",

  thresholdExpectedAbsence: (usdAmount: number) =>
    `มูลค่า ${usdAmount.toLocaleString("en-US", {
      maximumFractionDigits: 2,
    })} USD >= 50,000 USD — ไม่ควรพบใน DS_FTU`,

  thresholdUnexpectedPresence: (usdAmount: number) =>
    `มูลค่า ${usdAmount.toLocaleString("en-US", {
      maximumFractionDigits: 2,
    })} USD >= 50,000 USD แต่พบใน DS_FTU โดยไม่ควรพบ`,
};
/**
 * TODO: FTX Exception
 *
 * Requirement ยังมีเงื่อนไขที่ต้องตรวจร่วมกับ DS_FTX
 * แต่ Field และ Logic ยังไม่ครบถ้วน
 * จึงยังไม่เปิดใช้งาน Rule นี้
  
  ftxExceptionExpectedAbsence: (reason: string) =>
    `เข้าเงื่อนไข FTX Exception (${reason}) — ไม่ควรพบใน DS_FTU`,
  ftxExceptionUnexpectedPresence: (reason: string) =>
    `เข้าเงื่อนไข FTX Exception (${reason}) แต่พบใน DS_FTU โดยไม่ควรพบ`,
  */
