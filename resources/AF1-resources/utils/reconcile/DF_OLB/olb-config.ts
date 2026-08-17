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

/** จำกัดรายการใน Remark เพื่อให้ผลลัพธ์อ่านง่ายเมื่อพบ Candidate จำนวนมาก */
export const OLB_MAX_CANDIDATES_IN_REMARK = 10;

/** Fallback ต้องมี Field ตรงอย่างน้อย 2 Field หาก Amount ไม่ได้เป็นหลักฐาน */
export const OLB_MIN_STRONG_EVIDENCE_FIELDS = 2;

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

/**
 * Mapping กลางที่ Matcher และ Analyzer ใช้ร่วมกัน
 * ช่วยลดการเขียนชื่อ Test Data/Report Field ซ้ำหลายจุด
 */
export const OLB_FIELD_MAPPINGS = {
  primary: {
    testDataField: OLB_TEST_DATA_FIELDS.transactionId,
    reportField: OLB_REPORT_FIELDS.arrangementNumber,
  },
  date: {
    testDataField: OLB_TEST_DATA_FIELDS.transactionDate,
    reportField: OLB_REPORT_FIELDS.arrangementContractDate,
  },
  amount: {
    testDataField:
      OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
    reportField: OLB_REPORT_FIELDS.thbOutstandingAmount,
  },
  cifNo: {
    testDataField: OLB_TEST_DATA_FIELDS.cifNo,
    reportField: OLB_REPORT_FIELDS.custCode,
  },
  cifName: {
    testDataField: OLB_TEST_DATA_FIELDS.cifName,
    reportField: OLB_REPORT_FIELDS.custName,
  },
} as const;

export type OlbFieldMapping =
  (typeof OLB_FIELD_MAPPINGS)[keyof typeof OLB_FIELD_MAPPINGS];

export const OLB_REQUIRED_TEST_DATA_HEADERS = [
  OLB_TEST_DATA_FIELDS.testNo,
  ...Object.values(OLB_FIELD_MAPPINGS).map(
    (mapping) => mapping.testDataField,
  ),
] as const;

export const OLB_REMARKS = {
  pleaseReview: "Please review",
  matched: "พบรายการ DF_OLB ที่ตรงกับ Test Data",
} as const;