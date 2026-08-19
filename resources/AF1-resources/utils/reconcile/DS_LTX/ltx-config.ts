/**
 * Config สำหรับ Reconcile DS_LTX ใน Script 3
 *
 * จำนวน Fee Group ไม่กำหนดตายตัวในไฟล์นี้ เพราะ Builder
 * จะตรวจจาก Header ของ Test Data จริงในแต่ละรอบการทำงาน
 */

import type { ReportCode } from "../../../config/report-config";
import { getReportRuntimeConfig } from "../../../config/report-runtime.config";
import {
  RESIDENT_THB_TO_FCD_MUST_NOT_EXIST,
  type ReportPresenceRule,
} from "../shared/presence-rule";

export type CompareMode =
  | "exact"
  | "amountTolerance"
  | "fixedValue"
  | "dateWithIdFallback";

/** กติกาสำหรับเปรียบเทียบหนึ่ง Field ระหว่าง AF1 Report กับ Test Data */
export interface ReconcileFieldRule {
  reportField: string;
  testDataField: string | null;
  compareMode: CompareMode;
  fixedValue?: string;
  tolerance?: number;
  fallbackTestDataField?: string;
  fallbackReportField?: string;
  isRequiredForAllCases: boolean;
  skipWhenTestDataFieldEmpty?: string;
  onlyWhenReportFieldHasValue?: string;
  applicableSuffixes?: string[];
}

interface ReconcileGroupKeyFields {
  reportAccountField: string;
  testDataAccountField: string;
  reportCurrencyField: string;
  testDataCurrencyField: string;
}

export interface ReconcileReportConfig {
  headerRowNumber: number;
  testDataHeaderRowNumber: number;
  groupKeyFields: ReconcileGroupKeyFields;
  referenceNumberReportField: string;
  drSuffixLabel: string;
  feSuffixLabel: string;
  testDataIdField: string;
  testDataTestNoField: string;
  transactionAmountReportField: string;
  drAmountTestDataField: string;
  drAmountFallbackTestDataField: string;
  reportPresenceRules: ReportPresenceRule[];
  fieldRules: ReconcileFieldRule[];
  outputOnlyFields: string[];
}

/** ยอมให้ Amount ต่างกันได้ไม่เกิน 1 สตางค์ */
export const DEFAULT_AMOUNT_TOLERANCE = 0.01;

/** ค่าคงที่ของ DS_LTX เพื่อไม่ให้ Magic Value กระจายใน Config */
const LTX_REPORT_CODE: ReportCode = "DS_LTX";
const LTX_REPORT_HEADER_ROW = 1;
const LTX_DR_SUFFIX = "DR";
const LTX_FE_SUFFIX = "FE";
const RELATIONSHIP_OTHER_CODE = "172064";

const LTX_REPORT_FIELDS = {
  referenceNumber: "Reference Transaction Number",
  arrangementNumber: "FI Arrangement Number",
  transactionAmount: "Transaction Amount",
  transactionDate: "Transaction Date",
  currencyId: "Currency Id",
  outflowPurpose: "Outflow Transaction Purpose",
  inflowPurpose: "Inflow Transaction Purpose",
  beneficiaryName: "Beneficiary or Sender Name",
  beneficiaryCountry: "Country Id of Beneficiary or Sender",
  beneficiaryRelationship: "Relationship with Beneficiary or Sender",
  custCode: "Cust Code",
  custName: "Cust Name",
} as const;

const LTX_TEST_DATA_FIELDS = {
  transactionId: "Transaction ID/ Reconcile ID",
  testNo: "Test No.",
  transactionDate: "Txn Date",
  account: "From Account ( A/C Client/Sender)",
  currency: "From Currency (CCY)",
  transferAmount: "From Transfer Amount",
  debitAmount: "From Debit Amount",
  purposeCode: "From BOT Purpose code",
  beneficiaryName: "To CIF Name (Beneficiary)",
  beneficiaryCountry: "To Region (Bene Country)",
  custCode: "From CIF No. (Client/Sender)",
  custName: "From CIF Name (Client/Sender)",
} as const;

/**
 * Config นี้เก็บไว้ภายในโมดูล เพื่อให้ไฟล์อื่นเข้าถึงผ่าน
 * getReconcileConfig() และได้รับการตรวจความครบถ้วนก่อนใช้งาน
 */
const RECONCILE_CONFIG: Partial<
  Record<ReportCode, ReconcileReportConfig>
> = {
  DS_LTX: {
    headerRowNumber: LTX_REPORT_HEADER_ROW,
    testDataHeaderRowNumber:
      getReportRuntimeConfig(LTX_REPORT_CODE).testDataHeaderRowNumber,
    groupKeyFields: {
      reportAccountField: LTX_REPORT_FIELDS.arrangementNumber,
      testDataAccountField: LTX_TEST_DATA_FIELDS.account,
      reportCurrencyField: LTX_REPORT_FIELDS.currencyId,
      testDataCurrencyField: LTX_TEST_DATA_FIELDS.currency,
    },
    referenceNumberReportField: LTX_REPORT_FIELDS.referenceNumber,
    drSuffixLabel: LTX_DR_SUFFIX,
    feSuffixLabel: LTX_FE_SUFFIX,
    testDataIdField: LTX_TEST_DATA_FIELDS.transactionId,
    testDataTestNoField: LTX_TEST_DATA_FIELDS.testNo,
    transactionAmountReportField: LTX_REPORT_FIELDS.transactionAmount,
    drAmountTestDataField: LTX_TEST_DATA_FIELDS.transferAmount,
    drAmountFallbackTestDataField: LTX_TEST_DATA_FIELDS.debitAmount,
    reportPresenceRules: [RESIDENT_THB_TO_FCD_MUST_NOT_EXIST],
    fieldRules: [
      /** วันที่ตรงกัน หรือใช้วันที่ตำแหน่ง 7–12 ของ Reference เป็น fallback */
      {
        reportField: LTX_REPORT_FIELDS.transactionDate,
        testDataField: LTX_TEST_DATA_FIELDS.transactionDate,
        compareMode: "dateWithIdFallback",
        fallbackReportField: LTX_REPORT_FIELDS.referenceNumber,
        isRequiredForAllCases: true,
      },
      /** สกุลเงินต้องตรงกันทุก Test Case */
      {
        reportField: LTX_REPORT_FIELDS.currencyId,
        testDataField: LTX_TEST_DATA_FIELDS.currency,
        compareMode: "exact",
        isRequiredForAllCases: true,
      },
      /** ทั้ง DR และ FE ต้องใช้ Outflow Purpose ตรงกับ BOT Purpose code */
      {
        reportField: LTX_REPORT_FIELDS.outflowPurpose,
        testDataField: LTX_TEST_DATA_FIELDS.purposeCode,
        compareMode: "exact",
        isRequiredForAllCases: true,
      },
      /** ข้อมูลผู้รับและข้อมูลลูกค้าตรวจเฉพาะแถวธุรกรรมหลัก DR */
      {
        reportField: LTX_REPORT_FIELDS.beneficiaryName,
        testDataField: LTX_TEST_DATA_FIELDS.beneficiaryName,
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: [LTX_DR_SUFFIX],
      },
      {
        reportField: LTX_REPORT_FIELDS.beneficiaryCountry,
        testDataField: LTX_TEST_DATA_FIELDS.beneficiaryCountry,
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: [LTX_DR_SUFFIX],
        skipWhenTestDataFieldEmpty: LTX_TEST_DATA_FIELDS.beneficiaryName,
      },
      {
        reportField: LTX_REPORT_FIELDS.beneficiaryRelationship,
        testDataField: null,
        compareMode: "fixedValue",
        fixedValue: RELATIONSHIP_OTHER_CODE,
        isRequiredForAllCases: false,
        applicableSuffixes: [LTX_DR_SUFFIX],
        skipWhenTestDataFieldEmpty: LTX_TEST_DATA_FIELDS.beneficiaryName,
      },
      {
        reportField: LTX_REPORT_FIELDS.custCode,
        testDataField: LTX_TEST_DATA_FIELDS.custCode,
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: [LTX_DR_SUFFIX],
      },
      {
        reportField: LTX_REPORT_FIELDS.custName,
        testDataField: LTX_TEST_DATA_FIELDS.custName,
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: [LTX_DR_SUFFIX],
        skipWhenTestDataFieldEmpty: LTX_TEST_DATA_FIELDS.custCode,
      },
      /** Inflow Purpose มีความหมายเฉพาะแถวที่ Report ระบุว่าเป็น Inflow */
      {
        reportField: LTX_REPORT_FIELDS.inflowPurpose,
        testDataField: LTX_TEST_DATA_FIELDS.purposeCode,
        compareMode: "exact",
        isRequiredForAllCases: false,
        onlyWhenReportFieldHasValue: LTX_REPORT_FIELDS.inflowPurpose,
      },
    ],
    /**
     * Field กลุ่มนี้ใช้ตรวจว่ามี Header และนำไปแสดงผลเท่านั้น
     * เพราะยังไม่มี Expected Value ที่ยืนยันสำหรับการ Compare
     */
    outputOnlyFields: [
      "Loan Deposit Transaction Type",
      "Payment Method",
      "From Transaction Type",
      "To Transaction Type",
      "Installment Number",
      "Approval Document Number",
      "CMF CODE",
      "Data Set Date",
      "Data Submission Period",
    ],
  },
};

export const getHandledReportFields = (
  config: ReconcileReportConfig,
): string[] => [
  config.referenceNumberReportField,
  config.groupKeyFields.reportAccountField,
  config.transactionAmountReportField,
  ...config.fieldRules.map((rule) => rule.reportField),
  ...config.outputOnlyFields,
];

export const getReconcileConfig = (
  reportCode: string,
): ReconcileReportConfig => {
  const config = RECONCILE_CONFIG[reportCode as ReportCode];

  if (!config || config.groupKeyFields.reportAccountField === "") {
    throw new Error(
      `Reconcile Config ของ Report "${reportCode}" ยังไม่ครบถ้วน`,
    );
  }

  return config;
};