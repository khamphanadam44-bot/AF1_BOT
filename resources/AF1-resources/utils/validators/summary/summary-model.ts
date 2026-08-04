/**
 * summary-model.ts
 * ------------------------------------------------------------------
 * Data Model ของ Script 4
 *
 * ใช้ Fee Group แบบ Array เพื่อไม่ Fix จำนวน Fee Type ไว้ที่ 2
 * Report ที่มี Fee Type 1-N สามารถใช้ Model เดียวกันได้
 * ------------------------------------------------------------------
 

export type SummaryTestResult = "Pass" | "Fail";

export interface ReconcileSummaryCounts {
  totalChecked: number;
  passed: number;
  failed: number;
}

/**
 * สีของ Cell Test Result ที่อ่านมาจาก Reconcile Result
 * เก็บเฉพาะสีที่ต้องนำไปใช้ใน Summary ไม่ Copy Style อื่นจาก Source
 
export interface SummaryTestResultColorStyle {
  fillArgb?: string;
  fontColorArgb?: string;
}

export interface ReconcileDetailRecord {
  sourceRowNumber: number;
  testScriptNo: string;
  testResult: SummaryTestResult;
  testResultColorStyle: SummaryTestResultColorStyle;
  reason: string;
  referenceTransactionNumber: string;
  fieldValues: Readonly<Record<string, string>>;
}

export interface ReconcileSummarySource {
  reportCode: string;
  reconcileFilePath: string;
  reconcileFileName: string;
  reconcileSheetName: string;
  counts: ReconcileSummaryCounts;
  details: ReconcileDetailRecord[];
}

export interface SummaryFeeGroup {
  feeIndex: number;
  feeType: string;
  feeChargeAccountNo: string;
  feeAmount: string;
  feeCurrency: string;
}

export interface TestScriptDataRecord {
  sourceRowNumber: number;
  testScriptNo: string;
  referenceTransactionNumber: string;
  fieldValues: Readonly<Record<string, string>>;
  feeGroups: SummaryFeeGroup[];
}

export interface SummaryDetailRow {
  reconcile: ReconcileDetailRecord;
  testData: TestScriptDataRecord;
}

export interface SummaryRunMetadata {
  reportFileName: string;
  executionDate: string;
  executionTime: string;
  runId: string;
  verifiedBy: string;
  runTimestamp: string;
}

export interface SummaryGenerationResult {
  summaryFilePath: string;
  source: ReconcileSummarySource;
  metadata: SummaryRunMetadata;
  detailRowCount: number;
  displayedFeeGroupCount: number;
} */
