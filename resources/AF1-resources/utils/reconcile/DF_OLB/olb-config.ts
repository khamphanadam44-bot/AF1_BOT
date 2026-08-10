/**
 * olb-config.ts
 * ------------------------------------------------------------------
 * Config ของ DF_OLB สำหรับ Script 3 (Reconcile)
 *
 * แยกชื่อ Header และค่าคงที่ออกจาก Logic เพื่อให้แก้ Mapping ได้ง่าย
 * และไม่กระทบ Reconcile ของ Report อื่น
 * ------------------------------------------------------------------
 */

import { getReportRuntimeConfig } from "../../../config/report-runtime.config";

export const OLB_REPORT_CODE = "DF_OLB";
export const OLB_REPORT_HEADER_ROW = 1;

export const OLB_TEST_DATA_HEADER_ROW =
  getReportRuntimeConfig(OLB_REPORT_CODE).testDataHeaderRowNumber;

/** Amount ต่างกันไม่เกิน 0.01 ถือว่าตรงกัน */
export const OLB_AMOUNT_TOLERANCE = 0.01;

export const OLB_REPORT_FIELDS = {
  arrangementNumber: "FI Arrangement Number",
  arrangementContractDate: "Arrangement Contract Date",
  thbOutstandingAmount: "THB Outstanding Amount",
  custCode: "Cust Code",
  custName: "Cust Name",
} as const;

export const OLB_TEST_DATA_FIELDS = {
  testNo: "Test No.",
  transactionId: "Transaction ID/ Reconcile ID",
  transactionDate: "Txn Date",
  thbEquivalentTransferAmount: "From THB Equivalent Transfer Amount",
  cifNo: "From CIF No. (Client/Sender)",
  cifName: "From CIF Name (Client/Sender)",
} as const;

export const OLB_REMARKS = {
  pleaseReview: "Please review",
  matched: "พบรายการ DF_OLB ที่ตรงกับ Test Data",
} as const;