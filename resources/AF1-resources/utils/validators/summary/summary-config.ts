/**
 * summary-config.ts
 * ------------------------------------------------------------------
 * Config ของ Script 4 แยกตาม Report Code
 *
 * Service อ่านข้อมูลตาม Header และ Page เขียนข้อมูลตาม Target Column
 * ที่กำหนดในไฟล์นี้ จึงไม่ต้องใส่ Business Mapping ของแต่ละ Report
 * กระจายอยู่ใน summary.service.ts หรือ summary.page.ts
 * ------------------------------------------------------------------
 *

import type { ReportCode } from "../../../config/report-config";
import { getReportRuntimeConfig } from "../../../config/report-runtime.config";
import {
  getCheckedReportHeaderDir,
  getCheckedTestDataDir,
  getReconcileOutputDir,
  getSummaryOutputDir,
} from "../../../config/paths.config";

export interface HeaderAliases {
  readonly aliases: readonly string[];
}

export interface SummaryFieldMapping {
  readonly key: string;
  readonly sourceHeader: HeaderAliases;
  readonly targetColumn: number;
  readonly valueType?: "text" | "number";
}

export interface SummaryTemplateLocator {
  readonly title: string;
  readonly reportCode: string;
  readonly reportFileName: string;
  readonly executionDate: string;
  readonly executionTime: string;
  readonly runId: string;
  readonly verifiedBy: string;
  readonly totalChecked: string;
  readonly passedMatch: string;
  readonly failedUnmatch: string;
  readonly testScriptDataTitle: string;
  readonly detailHeaderRow: number;
  readonly detailStartRow: number;
  readonly detailStartColumn: number;
  readonly fixedDetailEndColumn: number;
  readonly columns: {
    readonly testResult: number;
    readonly reason: number;
    readonly testScriptNo: number;
    readonly testCaseScenario: number;
    readonly firstFeeColumn: number;
  };
  readonly feeColumnsPerGroup: number;
  readonly templateFeeGroupCount: number;
}

export interface SummaryReportConfig {
  readonly reportCode: ReportCode;
  readonly reconcileDirectory: string;
  readonly reconcileFilePrefix: string;
  readonly reconcileSheetName: string;
  readonly reconcileHeaderRowNumber: number;
  readonly reconcileIdentityHeaders: {
    readonly testScriptNo: HeaderAliases;
    readonly testResult: HeaderAliases;
    readonly reason: HeaderAliases;
    readonly referenceTransactionNumber: HeaderAliases;
  };
  readonly reconcileFields: readonly SummaryFieldMapping[];
  readonly checkedReportDirectory: string;
  readonly checkedReportSourceSheetName: string;
  readonly checkedReportTargetSheetName: string;
  readonly checkedTestDataDirectory: string;
  readonly checkedTestDataSourceSheetName: string;
  readonly checkedTestDataTargetSheetName: string;
  readonly testDataHeaderRowNumber: number;
  readonly testDataIdentityHeaders: {
    readonly testScriptNo: HeaderAliases;
    readonly referenceTransactionNumber: HeaderAliases;
  };
  readonly testDataFields: readonly SummaryFieldMapping[];
  readonly includeFeeGroups: boolean;
  readonly feeHeaderPatterns: {
    readonly feeType: RegExp;
    readonly feeAccount: RegExp;
    readonly feeAmount: RegExp;
    readonly feeCurrency: RegExp;
  };
  readonly referenceSuffixes: readonly string[];
  readonly templateFilePath: string;
  readonly templateSheetName: string;
  readonly summaryOutputDirectory: string;
  readonly summaryFilePrefix: string;
  readonly verifiedBy: string;
  readonly locator: SummaryTemplateLocator;
}

const COMMON_RECONCILE_IDENTITY_HEADERS: SummaryReportConfig["reconcileIdentityHeaders"] =
  {
    testScriptNo: { aliases: ["Test Script No.", "Test Script No"] },
    testResult: { aliases: ["Test Result.", "Test Result"] },
    reason: { aliases: ["Remark", "Reason"] },
    referenceTransactionNumber: {
      aliases: ["Reference Transaction Number"],
    },
  };

const COMMON_TEST_DATA_IDENTITY_HEADERS: SummaryReportConfig["testDataIdentityHeaders"] =
  {
    testScriptNo: { aliases: ["Test No.", "Test No", "Test Script No."] },
    referenceTransactionNumber: {
      aliases: [
        "Transaction ID/ Reconcile ID",
        "Reference Transaction Number",
      ],
    },
  };

const FEE_HEADER_PATTERNS: SummaryReportConfig["feeHeaderPatterns"] = {
  feeType: /^fee type\s*(\d+)$/i,
  feeAccount: /^fee charge account no\.?\s*type\s*(\d+)$/i,
  feeAmount: /^fee amount(?:\s*type)?\s*(\d+)$/i,
  feeCurrency: /^fee currency(?:\s*type)?\s*(\d+)$/i,
};

const BASE_LOCATOR = {
  title: "B2",
  reportCode: "D13",
  reportFileName: "C5",
  executionDate: "C6",
  executionTime: "C7",
  runId: "C8",
  verifiedBy: "C9",
  totalChecked: "I7",
  passedMatch: "J7",
  failedUnmatch: "K7",
  testScriptDataTitle: "L13",
  detailHeaderRow: 14,
  detailStartRow: 15,
  detailStartColumn: 2,
} as const;

const LTX_LOCATOR: SummaryTemplateLocator = {
  ...BASE_LOCATOR,
  fixedDetailEndColumn: 28,
  columns: {
    testResult: 2,
    reason: 3,
    testScriptNo: 12,
    testCaseScenario: 13,
    firstFeeColumn: 21,
  },
  feeColumnsPerGroup: 4,
  templateFeeGroupCount: 2,
};

const FTU_LOCATOR: SummaryTemplateLocator = {
  ...BASE_LOCATOR,
  fixedDetailEndColumn: 20,
  columns: {
    testResult: 2,
    reason: 3,
    testScriptNo: 12,
    testCaseScenario: 13,
    firstFeeColumn: 21,
  },
  feeColumnsPerGroup: 4,
  templateFeeGroupCount: 0,
};

const LTX_RECONCILE_FIELDS: readonly SummaryFieldMapping[] = [
  {
    key: "referenceTransactionNumber",
    sourceHeader: { aliases: ["Reference Transaction Number"] },
    targetColumn: 4,
  },
  {
    key: "dataSetDate",
    sourceHeader: { aliases: ["Data Set Date"] },
    targetColumn: 5,
  },
  {
    key: "fiArrangementNumber",
    sourceHeader: { aliases: ["FI Arrangement Number"] },
    targetColumn: 6,
  },
  {
    key: "custCode",
    sourceHeader: { aliases: ["Cust Code", "CMF CODE", "CMF Code"] },
    targetColumn: 7,
  },
  {
    key: "paymentMethod",
    sourceHeader: { aliases: ["Payment Method"] },
    targetColumn: 8,
  },
  {
    key: "currencyId",
    sourceHeader: { aliases: ["Currency Id", "Currency ID"] },
    targetColumn: 9,
  },
  {
    key: "transactionAmount",
    sourceHeader: { aliases: ["Transaction Amount"] },
    targetColumn: 10,
    valueType: "number",
  },
];

const LTX_TEST_DATA_FIELDS: readonly SummaryFieldMapping[] = [
  {
    key: "testCaseScenario",
    sourceHeader: { aliases: ["Test Case / Scenario"] },
    targetColumn: 13,
  },
  {
    key: "txnDate",
    sourceHeader: { aliases: ["Txn Date"] },
    targetColumn: 14,
  },
  {
    key: "referenceTransactionNumber",
    sourceHeader: {
      aliases: [
        "Transaction ID/ Reconcile ID",
        "Reference Transaction Number",
      ],
    },
    targetColumn: 15,
  },
  {
    key: "fromCifNo",
    sourceHeader: { aliases: ["From CIF No. (Client/Sender)"] },
    targetColumn: 16,
  },
  {
    key: "fromAccount",
    sourceHeader: {
      aliases: [
        "From Account ( A/C Client/Sender)",
        "From Account (A/C Client/Sender)",
      ],
    },
    targetColumn: 17,
  },
  {
    key: "fromCurrency",
    sourceHeader: { aliases: ["From Currency (CCY)"] },
    targetColumn: 18,
  },
  {
    key: "fromDebitAmount",
    sourceHeader: { aliases: ["From Debit Amount", "From Debit Amount "] },
    targetColumn: 19,
    valueType: "number",
  },
  {
    key: "fromTransferAmount",
    sourceHeader: { aliases: ["From Transfer Amount"] },
    targetColumn: 20,
    valueType: "number",
  },
];

const FTU_RECONCILE_FIELDS: readonly SummaryFieldMapping[] = [
  {
    key: "arrNumber",
    sourceHeader: { aliases: ["Arr Number"] },
    targetColumn: 4,
  },
  {
    key: "dataSetDate",
    sourceHeader: { aliases: ["Data Set Date"] },
    targetColumn: 5,
  },
  {
    key: "inflowTransactionPurpose",
    sourceHeader: { aliases: ["Inflow Transaction Purpose"] },
    targetColumn: 6,
  },
  {
    key: "outflowTransactionPurpose",
    sourceHeader: { aliases: ["Outflow Transaction Purpose"] },
    targetColumn: 7,
  },
  {
    key: "currencyId",
    sourceHeader: { aliases: ["Currency Id", "Currency ID"] },
    targetColumn: 8,
  },
  {
    key: "legType",
    sourceHeader: { aliases: ["Leg Type"] },
    targetColumn: 9,
  },
  {
    key: "foreignCurrencyAmount",
    sourceHeader: { aliases: ["Foreign Currency Amount"] },
    targetColumn: 10,
    valueType: "number",
  },
];

const FTU_TEST_DATA_FIELDS: readonly SummaryFieldMapping[] = [
  {
    key: "testCaseScenario",
    sourceHeader: { aliases: ["Test Case / Scenario"] },
    targetColumn: 13,
  },
  {
    key: "txnDate",
    sourceHeader: { aliases: ["Txn Date"] },
    targetColumn: 14,
  },
  {
    key: "referenceTransactionNumber",
    sourceHeader: {
      aliases: ["Transaction ID/ Reconcile ID"],
    },
    targetColumn: 15,
  },
  {
    key: "fromCurrency",
    sourceHeader: { aliases: ["From Currency (CCY)"] },
    targetColumn: 16,
  },
  {
    key: "toCurrency",
    sourceHeader: { aliases: ["To Currency (CCY)"] },
    targetColumn: 17,
  },
  {
    key: "botPurposeCode",
    sourceHeader: {
      aliases: ["From BOT Purpose code", "BOT Purpose code"],
    },
    targetColumn: 18,
  },
  {
    key: "settledAmount",
    sourceHeader: {
      aliases: ["Settled Amount (CCY)", "Settled Amount"],
    },
    targetColumn: 19,
    valueType: "number",
  },
  {
    key: "settledCurrency",
    sourceHeader: {
      aliases: ["Settled Currency (CCY)", "Settled Currency"],
    },
    targetColumn: 20,
  },
];

const buildCommonConfig = (
  reportCode: ReportCode,
): Omit<
  SummaryReportConfig,
  | "reconcileIdentityHeaders"
  | "reconcileFields"
  | "checkedReportSourceSheetName"
  | "checkedReportTargetSheetName"
  | "testDataIdentityHeaders"
  | "testDataFields"
  | "includeFeeGroups"
  | "referenceSuffixes"
  | "locator"
> => {
  const runtimeConfig = getReportRuntimeConfig(reportCode);

  return {
    reportCode,
    reconcileDirectory: getReconcileOutputDir(reportCode),
    reconcileFilePrefix: `${reportCode}_Reconcile_`,
    reconcileSheetName: `${reportCode}_Reconcile`,
    reconcileHeaderRowNumber: 1,
    checkedReportDirectory: getCheckedReportHeaderDir(reportCode),
    checkedTestDataDirectory: getCheckedTestDataDir(reportCode),
    checkedTestDataSourceSheetName: runtimeConfig.testDataSheetName,
    checkedTestDataTargetSheetName: runtimeConfig.testDataSheetName,
    testDataHeaderRowNumber: runtimeConfig.testDataHeaderRowNumber,
    feeHeaderPatterns: FEE_HEADER_PATTERNS,
    templateFilePath: runtimeConfig.summaryTemplateFilePath,
    templateSheetName: runtimeConfig.summaryTemplateSheetName,
    summaryOutputDirectory: getSummaryOutputDir(reportCode),
    summaryFilePrefix: `${reportCode}_Automation_Summary_`,
    verifiedBy: runtimeConfig.verifiedBy,
  };
};

/**
 * Bug fix (Code Review): เดิม getSummaryConfig ใช้ if-chain (if reportCode === "DS_LTX"
 * ... if reportCode === "DS_FTU" ... throw) ซึ่งไม่ตรงกับ Pattern Registry ที่ Config
 * อื่นในระบบใช้ (RECONCILE_CONFIG, EXPECTED_CASE_BUILDERS, RECONCILE_RUNNERS) — เปลี่ยน
 * เป็น Registry Pattern เดียวกัน เพิ่ม Report ใหม่ในอนาคตแค่เติม Key ใน
 * SUMMARY_CONFIG_BUILDERS ไม่ต้องแก้ Logic ในฟังก์ชัน getSummaryConfig เลย
 
type SummaryConfigBuilder = (
  common: ReturnType<typeof buildCommonConfig>,
) => SummaryReportConfig;

const SUMMARY_CONFIG_BUILDERS: Partial<
  Record<ReportCode, SummaryConfigBuilder>
> = {
  DS_LTX: (common) => ({
    ...common,
    reconcileIdentityHeaders: COMMON_RECONCILE_IDENTITY_HEADERS,
    reconcileFields: LTX_RECONCILE_FIELDS,
    checkedReportSourceSheetName: "DS_LTX",
    checkedReportTargetSheetName: "DS_LTX",
    testDataIdentityHeaders: COMMON_TEST_DATA_IDENTITY_HEADERS,
    testDataFields: LTX_TEST_DATA_FIELDS,
    includeFeeGroups: true,
    referenceSuffixes: ["DR", "FE"],
    locator: LTX_LOCATOR,
  }),

  DS_FTU: (common) => ({
    ...common,
    reconcileIdentityHeaders: {
      ...COMMON_RECONCILE_IDENTITY_HEADERS,
      /**
       * FTU จับคู่ Transaction ID/Reconcile ID
       * ของ Test Data กับ Arr Number ใน Report
       
      referenceTransactionNumber: {
        aliases: ["Arr Number"],
      },
    },
    reconcileFields: FTU_RECONCILE_FIELDS,
    checkedReportSourceSheetName: "DS_FTU Transaction",
    checkedReportTargetSheetName: "DS_FTU",
    testDataIdentityHeaders: COMMON_TEST_DATA_IDENTITY_HEADERS,
    testDataFields: FTU_TEST_DATA_FIELDS,
    includeFeeGroups: false,
    referenceSuffixes: [],
    locator: FTU_LOCATOR,
  }),
};

/**
 * สร้าง Summary Config ของ Report ที่รองรับ Script 4 แล้ว
 
export const getSummaryConfig = (
  reportCode: ReportCode,
): SummaryReportConfig => {
  const builder = SUMMARY_CONFIG_BUILDERS[reportCode];

  if (!builder) {
    throw new Error(
      `Script 4 Summary Config is not configured for Report "${reportCode}".`,
    );
  }

  return builder(buildCommonConfig(reportCode));
};
*/