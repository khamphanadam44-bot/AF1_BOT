/**
 * FXM-config.ts
 * ------------------------------------------------------------------
 * Config และ Constant ของ DF_FXM สำหรับ Script 3
 *
 * หน้าที่:
 * 1. กำหนด Report Code
 * 2. กำหนดตำแหน่ง Header ของ Report และ Test Data
 * 3. กำหนดชื่อ Header ฝั่ง DF_FXM Report
 * 4. กำหนดชื่อ Header ฝั่ง Test Data
 * 5. กำหนด Threshold ของ DF_FXM
 * 6. กำหนด Arrangement Type
 * 7. กำหนด Leg Type และ Leg Type Name
 *
 * ไฟล์นี้เก็บเฉพาะ Config และ Constant
 * ไม่มี Logic การจับคู่หรือการตัดสิน PASS/FAIL
 * ------------------------------------------------------------------
 */

import {
  getReportRuntimeConfig,
} from "../../../config/report-runtime.config";

/**
 * Report Code ที่ใช้ร่วมกันภายใน DF_FXM Reconciler
 */
export const FXM_REPORT_CODE =
  "DF_FXM";

/**
 * Header ของ DF_FXM Report อยู่ที่แถวที่ 1
 */
export const FXM_REPORT_HEADER_ROW =
  1;

/**
 * Header ของ Test Data อยู่ที่แถวที่ 5
 *
 * อ่านค่าจาก Report Runtime Config
 * เพื่อไม่กำหนดเลขแถวซ้ำหลายไฟล์
 */
export const FXM_TEST_DATA_HEADER_ROW =
  getReportRuntimeConfig(
    FXM_REPORT_CODE,
  ).testDataHeaderRowNumber;

/**
 * สกุลเงินบาท
 *
 * ใช้ประกอบการตัดสินทิศทางของ FX Transaction
 */
export const FXM_THB_CURRENCY_CODE =
  "THB";

/**
* Threshold ของ DF_FXM
 *
 * ตั้งแต่ 1,000,000 USD ขึ้นไป:
 * - ต้องพิจารณารายงานใน DF_FXM
 *
 * ต่ำกว่า 1,000,000 USD:
 * - ต้องไม่พบใน DF_FXM
 * - ต้องพิจารณารายงานใน DF_FXU
 */
export const FXM_USD_THRESHOLD =
  1_000_000;

/**
 * วันที่เริ่มมีผลของ Requirement DF_FXM
 *
 * รูปแบบ:
 * yyyy-MM-dd
 */
export const FXM_EFFECTIVE_DATE =
  "2025-06-02";

/**
 * Arrangement Type ที่ Requirement กำหนด
 */
export const FXM_ARRANGEMENT_TYPE =
  "018101";

/**
 * Leg Type ของ DF_FXM
 *
 * 182001:
 * ตัวแทนรับอนุญาตซื้อเงินตราต่างประเทศ
 * แลกกับสกุลเงินบาท
 *
 * 182002:
 * ตัวแทนรับอนุญาตขายเงินตราต่างประเทศ
 * แลกกับสกุลเงินบาท
 */
export const FXM_LEG_TYPES = {
  buyForeignCurrency:
    "182001",

  sellForeignCurrency:
    "182002",
} as const;

/**
 * Leg Type Name ที่ต้องสัมพันธ์กับ Leg Type
 */
export const FXM_LEG_TYPE_NAMES = {
  [FXM_LEG_TYPES.buyForeignCurrency]:
    "ตัวแทนรับอนุญาตซื้อเงินตราต่างประเทศแลกกับสกุลเงินบาท",

  [FXM_LEG_TYPES.sellForeignCurrency]:
    "ตัวแทนรับอนุญาตขายเงินตราต่างประเทศแลกกับสกุลเงินบาท",
} as const;

/**
 * Header ฝั่ง DF_FXM Report
 *
 * ชื่อทั้งหมดอ้างอิงจากไฟล์ Export จริง:
 * Worksheet "DF_FXM Transaction"
 */
export const FXM_REPORT_FIELDS = {
  departmentCode:
    "DEPT CODE",

  dataSetDate:
    "Data Set Date",

  arrangementNumber:
    "FI Arrangement Number",

  customerCode:
    "Cust Code",

  cmfCode:
    "CMF CODE",

  customerName:
    "Cust Name",

  arrangementType:
    "Arrangement Type",

  arrangementTypeName:
    "Arrangement Type Name",

  legType:
    "Leg Type",

  legTypeName:
    "Leg Type Name",

  currencyId:
    "Currency Code",

  currencyIdName:
    "Currency Code Name",

  originalAmount:
    "Original Amount",

  usdEquivalentAmount:
    "USD Equivalent Amount",
} as const;

/**
 * Header ฝั่ง Test Data
 *
 * รายการนี้ตรงกับ Requirement ใหม่
 * ใน TESTDATA_CONFIG.DF_FXM
 */
export const FXM_TEST_DATA_FIELDS = {
  transactionId:
    "Transaction ID/ Reconcile ID",

  fromCurrency:
    "From Currency (CCY)",

  toCurrency:
    "To Currency (CCY)",

  settledCurrency:
    "Settled Currency (CCY)",

  settledAmount:
    "Settled Amount (CCY)",

  transactionDate:
    "Txn Date",

  fromCustomerTypeCode:
    "From Customer Type Code",

  testNo:
    "Test No.",

  fromCustomerTypeDescription:
    "From Customer Type Description",
} as const;