/**
 * Script 3 reconcile strategy for DS_FTU.
 *
 * Flow:
 * 1. Read the latest raw DS_FTU report and Test Data.
 * 2. Validate structural headers.
 * 3. Match Transaction ID/Reconcile ID with Arr Number.
 * 4. Evaluate FTU presence rules and compare mapped fields.
 * 5. Write the result with the same workbook format used by DS_LTX.
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { ReconcileRecord } from "../shared/record";
import { formatCompareRemark } from "../shared/remark";
import { ReconcileResultSheetWriter, ResultRow } from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import {
  FTU_AMOUNT_TOLERANCE,
  FTU_COUNTRY_ID_BY_CURRENCY,
  FTU_LEG_TYPES,
  FTU_REPORT_CODE,
  FTU_REPORT_FIELDS,
  FTU_REPORT_HEADER_ROW,
  FTU_TEST_DATA_HEADER_ROW,
  FTU_TEST_FIELDS,
} from "./ftu-config";
import {
  FtuExceptionEvaluator,
  FtuExpectedCase,
  FtuRuleEvaluator,
} from "./ftu-rules";
import { normalize, parseAmount } from "./ftu-parse.util";

/**
 * Bug fix (Code Review): เดิม normalize/parseAmount ประกาศซ้ำในไฟล์นี้เอง
 * (Copy-Paste เป๊ะจาก ftu-rules.ts) — ย้ายไปรวมไว้ที่ ftu-parse.util.ts
 * แล้ว Import มาใช้แทน ดู Import ด้านบน
 */

const parseDate = (value: unknown): Date | null => {
  const text = String(value ?? "").trim();
  const dayFirst = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  const yearFirst = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

  let year: number;
  let month: number;
  let day: number;

  if (dayFirst) {
    day = Number(dayFirst[1]);
    month = Number(dayFirst[2]);
    year = Number(dayFirst[3]);
  } else if (yearFirst) {
    year = Number(yearFirst[1]);
    month = Number(yearFirst[2]);
    day = Number(yearFirst[3]);
  } else {
    return null;
  }

  const date = new Date(Date.UTC(year, month - 1, day));
  const isValid =
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;

  return isValid ? date : null;
};

const nextWeekday = (sourceDate: Date): Date => {
  const result = new Date(sourceDate.getTime());

  do {
    result.setUTCDate(result.getUTCDate() + 1);
  } while (result.getUTCDay() === 0 || result.getUTCDay() === 6);

  return result;
};

const isSameDate = (left: Date, right: Date): boolean =>
  left.getTime() === right.getTime();

interface ResolvedCase {
  result: ResultRow;
  matchedRowNumbers: number[];
}

type AddComparisonRemark = (
  reportField: string,
  testDataField: string,
  expected: string,
  actual: string,
) => void;

/**
 * Placeholder ชั่วคราวสำหรับ R5 (FTX Exception) ตาม Requirement 3.7.3.1
 *
 * TODO (รอ QA ผู้รับผิดชอบ DS_FTX): แทนที่ด้วย FtxExceptionEvaluator ตัวจริง
 * เมื่อโค้ดฝั่ง DS_FTX พร้อม — Contract ดูได้ที่ FtuExceptionEvaluator
 * interface ใน ftu-rules.ts
 *
 * เงื่อนไข R5 ตาม Requirement:
 *   < 50,000 USD แต่เป็น Spot Today ตามข้อยกเว้น, Exercise counterparty
 *   ไม่ตรง Arrangement หรือ Not Sell/Not Deposit
 *   -> ไม่รายงานใน DS_FTU, ต้องไปรายงานใน DS_FTX แทน (ต้องมี Exercising
 *      Involved Party Name)
 *
 * พฤติกรรมตอนนี้: ไม่ตรวจ R5 เลย (เหมือนก่อนแก้ทุกประการ ไม่เปลี่ยนผลลัพธ์
 * ปัจจุบัน) แต่จะ Warn ครั้งเดียวต่อการรัน 1 รอบ เพื่อไม่ให้ผล PASS
 * ดูสมบูรณ์เกินจริงระหว่างที่ยังไม่มี Evaluator จริง
 */
const createPendingFtxExceptionEvaluator = (): FtuExceptionEvaluator => {
  let hasWarned = false;

  return {
    getExceptionReason: (): string | undefined => {
      if (!hasWarned) {
        hasWarned = true;
        console.warn(
          "⚠️  [DS_FTU] FtxExceptionEvaluator ยังไม่ได้ Implement — " +
            "Rule R5 (FTX Exception, Requirement 3.7.3.1) จะไม่ถูกตรวจสอบในรอบนี้ " +
            "ผล PASS ที่ได้อาจไม่ครอบคลุมกรณี FTX Exception จริง " +
            "ดู TODO ที่หัวไฟล์ ftu-reconcile.ts",
        );
      }

      return undefined;
    },
  };
};

export class FtuReconcileService {
  constructor(
    private readonly workbookPreparer = new ReconcileWorkbookPreparer(),
    private readonly excelReader = new ReconcileExcelReader(),
    private readonly sheetWriter = new ReconcileResultSheetWriter(),
    private readonly ruleEvaluator = new FtuRuleEvaluator({
      exceptionEvaluator: createPendingFtxExceptionEvaluator(),
    }),
  ) {}

  async reconcile(testDataFilePath: string): Promise<string> {
    console.log(`\n===== RECONCILE - ${FTU_REPORT_CODE} =====`);

    const prepared = await this.workbookPreparer.prepare(
      FTU_REPORT_CODE,
      FTU_REPORT_HEADER_ROW,
    );

    const reportData = this.excelReader.parseWorksheet(
      prepared.reportWorksheet,
      FTU_REPORT_HEADER_ROW,
    );
    const testData = await this.excelReader.readFile(
      testDataFilePath,
      FTU_TEST_DATA_HEADER_ROW,
    );

    this.validateHeaders(prepared.reportHeaders, testData.headers);

    const expectedCases = this.ruleEvaluator.buildExpectedCases(
      testData.records,
    );
    const reportRecordsById = this.indexReportRecords(reportData.records);
    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const caseResults: ResultRow[] = [];

    for (const expectedCase of expectedCases) {
      const matchedRecords =
        reportRecordsById.get(normalize(expectedCase.transactionId)) ?? [];
      const resolvedCase = expectedCase.validationError
        ? this.resolveInvalidCase(expectedCase, matchedRecords)
        : expectedCase.expectedPresence
          ? this.resolveExpectedPresence(expectedCase, matchedRecords)
          : this.resolveExpectedAbsence(expectedCase, matchedRecords);

      /**
       * ใช้ Business Result โดยตรง
       *
       * ไม่เพิ่ม Data Quality Warning ลงใน Remark
       * เพราะกรณี Test No. ว่างถูกจัดการด้วย
       * resolveIdentity() ใน shared/record.ts แล้ว
       */
      const result = resolvedCase.result;

      caseResults.push(result);

      if (resolvedCase.matchedRowNumbers.length === 0) {
        unmatchedRows.push(result);

        continue;
      }

      for (const rowNumber of resolvedCase.matchedRowNumbers) {
        annotationByRowNumber.set(rowNumber, result);
      }
    }

    this.sheetWriter.writeHeaderRow(
      prepared.resultSheet,
      prepared.reportHeaders,
    );

    const nextRowNumber = this.sheetWriter.writeRowsInRequestedOrder(
      prepared.resultSheet,
      prepared.reportWorksheet,
      prepared.reportHeaders,
      FTU_REPORT_HEADER_ROW + 1,
      prepared.reportWorksheet.rowCount,
      annotationByRowNumber,
      unmatchedRows,
    );

    this.sheetWriter.finalizeAutoFilter(
      prepared.resultSheet,
      prepared.reportHeaders,
      nextRowNumber - 1,
    );

    prepared.workbook.removeWorksheet(prepared.reportWorksheet.id);
    await prepared.workbook.xlsx.writeFile(prepared.reconcileFilePath);

    this.logSummary(prepared.reconcileFilePath, caseResults);
    return prepared.reconcileFilePath;
  }

  private validateHeaders(
    reportHeaders: string[],
    testDataHeaders: string[],
  ): void {
    const reportName = requireMappingReportName(FTU_REPORT_CODE);
    const requiredReportHeaders = getUniqueMappingHeaders(reportName);

    this.assertHeaders(reportHeaders, requiredReportHeaders, "Raw Report");
    this.assertHeaders(
      testDataHeaders,
      Object.values(FTU_TEST_FIELDS),
      "Test Data",
    );
  }

  private assertHeaders(
    actualHeaders: string[],
    requiredHeaders: readonly string[],
    sourceName: string,
  ): void {
    const actualHeaderSet = new Set(
      actualHeaders
        .filter((header) => header.trim() !== "")
        .map(canonicalHeader),
    );
    const missingHeaders = requiredHeaders.filter(
      (header) => !actualHeaderSet.has(canonicalHeader(header)),
    );

    if (missingHeaders.length > 0) {
      throw new Error(
        `[${FTU_REPORT_CODE}] ${sourceName} missing header(s): ` +
          missingHeaders.join(", "),
      );
    }
  }

  private indexReportRecords(
    records: ReconcileRecord[],
  ): Map<string, ReconcileRecord[]> {
    const recordsById = new Map<string, ReconcileRecord[]>();

    for (const record of records) {
      const arrangementNumber = normalize(
        record.get(FTU_REPORT_FIELDS.arrangementNumber),
      );

      if (arrangementNumber === "") {
        continue;
      }

      const matchingRecords = recordsById.get(arrangementNumber) ?? [];
      matchingRecords.push(record);
      recordsById.set(arrangementNumber, matchingRecords);
    }

    return recordsById;
  }

  private resolveExpectedAbsence(
    expectedCase: FtuExpectedCase,
    matchedRecords: ReconcileRecord[],
  ): ResolvedCase {
    if (matchedRecords.length === 0) {
      return {
        result: {
          testCaseNo: expectedCase.testCaseNo,
          status: "PASS",
          remark: expectedCase.successRemark,
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: [],
          isExpectedAbsence: true,
        },
        matchedRowNumbers: [],
      };
    }

    const unexpectedPresenceRemark =
      expectedCase.unexpectedPresenceRemark ?? expectedCase.successRemark;

    return {
      result: {
        testCaseNo: expectedCase.testCaseNo,
        status: "FAIL",
        remark:
          `${unexpectedPresenceRemark}\n` +
          formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_FIELDS.transactionId,
            expectedCase.transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            matchedRecords[0].get(FTU_REPORT_FIELDS.arrangementNumber),
          ),
        matchedRowNumber: matchedRecords[0].rowNumber,
        failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      },
      matchedRowNumbers: matchedRecords.map((record) => record.rowNumber),
    };
  }

  private resolveInvalidCase(
    expectedCase: FtuExpectedCase,
    matchedRecords: ReconcileRecord[],
  ): ResolvedCase {
    const fromCurrency = expectedCase.record.get(FTU_TEST_FIELDS.fromCurrency);

    const toCurrency = expectedCase.record.get(FTU_TEST_FIELDS.toCurrency);

    const isFromCurrencyMissing = normalize(fromCurrency) === "";

    const isToCurrencyMissing = normalize(toCurrency) === "";

    /**
     * Test Data ว่างเป็น Data Quality Remark
     * แต่ไม่ใช่เหตุให้ Business Result เป็น FAIL
     *
     * Transaction ID ยังคงถูกใช้จับคู่กับ Arr Number ตามเดิม
     */
    if (isFromCurrencyMissing || isToCurrencyMissing) {
      const matchedRecord = matchedRecords[0];

      const actualLegType =
        matchedRecord?.get(FTU_REPORT_FIELDS.legType) ?? "ไม่พบข้อมูล";

      const actualCountry =
        matchedRecord?.get(FTU_REPORT_FIELDS.beneficiaryCountry) ??
        "ไม่พบข้อมูล";

      const remarks: string[] = [];
      const reviewFieldHeaders: string[] = [];

      if (isFromCurrencyMissing) {
        remarks.push(
          formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_FIELDS.fromCurrency,
            fromCurrency,
            FTU_REPORT_FIELDS.legType,
            actualLegType,
          ),
        );

        reviewFieldHeaders.push(FTU_REPORT_FIELDS.legType);
      }

      if (isToCurrencyMissing) {
        remarks.push(
          formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_FIELDS.toCurrency,
            toCurrency,
            FTU_REPORT_FIELDS.beneficiaryCountry,
            actualCountry,
          ),
        );

        reviewFieldHeaders.push(FTU_REPORT_FIELDS.beneficiaryCountry);
      }

      return {
        result: {
          testCaseNo: expectedCase.testCaseNo,
          status: "PASS",
          remark: remarks.join("\n"),
          matchedRowNumber: matchedRecord?.rowNumber,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders,
          isExpectedAbsence: matchedRecords.length === 0,
        },
        matchedRowNumbers: matchedRecords.map((record) => record.rowNumber),
      };
    }

    /**
     * ค่าไม่ว่างแต่รูปแบบธุรกรรมไม่ถูกต้อง เช่น THB -> THB
     * ยังคงเป็น FAIL ตาม Business Rule เดิม
     */
    return {
      result: {
        testCaseNo: expectedCase.testCaseNo,
        status: "FAIL",
        remark:
          formatCompareRemark(
            FTU_REPORT_CODE,
            `${FTU_TEST_FIELDS.fromCurrency} / ` + FTU_TEST_FIELDS.toCurrency,
            `${fromCurrency} / ${toCurrency}`,
            "Expected Currency Direction",
            "THB -> FCY หรือ FCY -> THB",
          ) + `\n${expectedCase.validationError ?? "Data Validation Error"}`,
        matchedRowNumber: matchedRecords[0]?.rowNumber,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      },
      matchedRowNumbers: matchedRecords.map((record) => record.rowNumber),
    };
  }

  private resolveExpectedPresence(
    expectedCase: FtuExpectedCase,
    matchedRecords: ReconcileRecord[],
  ): ResolvedCase {
    if (matchedRecords.length === 0) {
      return {
        result: {
          testCaseNo: expectedCase.testCaseNo,
          status: "FAIL",
          remark: formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_FIELDS.transactionId,
            expectedCase.transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            "ไม่พบข้อมูล",
          ),
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
          reviewFieldHeaders: [],
          isExpectedAbsence: false,
        },
        matchedRowNumbers: [],
      };
    }

    const failedHeaders = new Set<string>();
    const reviewHeaders = new Set<string>();
    const remarks: string[] = [];

    const addFailure: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      failedHeaders.add(reportField);

      remarks.push(
        formatCompareRemark(
          FTU_REPORT_CODE,
          testDataField,
          expected,
          reportField,
          actual,
        ),
      );
    };

    /**
     * ใช้เมื่อ Test Data ว่าง
     *
     * เพิ่มเฉพาะ Remark และ Highlight สี Review
     * โดยไม่เพิ่ม Field เข้า failedHeaders
     */
    const addInformation: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      reviewHeaders.add(reportField);

      remarks.push(
        formatCompareRemark(
          FTU_REPORT_CODE,
          testDataField,
          expected,
          reportField,
          actual,
        ),
      );
    };

    for (const record of matchedRecords) {
      this.compareRecordFields(
        expectedCase,
        record,
        addFailure,
        addInformation,
      );
    }

    this.compareAggregateAmount(
      expectedCase,
      matchedRecords,
      addFailure,
      addInformation,
    );

    const aggregateText =
      matchedRecords.length > 1
        ? `Aggregate ${matchedRecords.length} raw rows by Arr Number.`
        : "";
    const status = failedHeaders.size === 0 ? "PASS" : "FAIL";

    const finalRemark = [
      status === "PASS" ? expectedCase.successRemark : "",
      ...remarks,
      aggregateText,
    ]
      .map((message) => message.trim())
      .filter((message) => message !== "")
      .join("\n");

    return {
      result: {
        testCaseNo: expectedCase.testCaseNo,
        status,
        remark: finalRemark,
        matchedRowNumber: matchedRecords[0].rowNumber,
        failedKeyFieldHeaders: [...failedHeaders],
        reviewFieldHeaders: [...reviewHeaders],
        isExpectedAbsence: false,
      },
      matchedRowNumbers: matchedRecords.map((record) => record.rowNumber),
    };
  }

  private compareRecordFields(
    expectedCase: FtuExpectedCase,
    record: ReconcileRecord,
    addFailure: (...parameters: Parameters<AddComparisonRemark>) => void,
    addInformation: (...parameters: Parameters<AddComparisonRemark>) => void,
  ): void {
    const expectedDate = expectedCase.record.get(
      FTU_TEST_FIELDS.transactionDate,
    );
    const actualDate = record.get(FTU_REPORT_FIELDS.dataSetDate);

    if (normalize(expectedDate) === "") {
      addInformation(
        FTU_REPORT_FIELDS.dataSetDate,
        FTU_TEST_FIELDS.transactionDate,
        expectedDate,
        actualDate,
      );
    } else if (!this.isAcceptedReportDate(expectedDate, actualDate)) {
      addFailure(
        FTU_REPORT_FIELDS.dataSetDate,
        FTU_TEST_FIELDS.transactionDate,
        `${expectedDate} หรือ weekday ถัดไป`,
        actualDate,
      );
    }

    const expectedLegType =
      expectedCase.direction === "BUY_FCY"
        ? FTU_LEG_TYPES.buyForeignCurrency
        : FTU_LEG_TYPES.sellForeignCurrency;
    const actualLegType = record.get(FTU_REPORT_FIELDS.legType);

    if (normalize(actualLegType) !== expectedLegType) {
      addFailure(
        FTU_REPORT_FIELDS.legType,
        `${FTU_TEST_FIELDS.fromCurrency}/${FTU_TEST_FIELDS.toCurrency}`,
        expectedLegType,
        actualLegType,
      );
    }

    this.comparePurpose(expectedCase, record, addFailure, addInformation);

    const toCurrency = normalize(
      expectedCase.record.get(FTU_TEST_FIELDS.toCurrency),
    );
    const expectedCountry = FTU_COUNTRY_ID_BY_CURRENCY[toCurrency];
    const actualCountry = record.get(FTU_REPORT_FIELDS.beneficiaryCountry);

    if (!expectedCountry || normalize(actualCountry) !== expectedCountry) {
      addFailure(
        FTU_REPORT_FIELDS.beneficiaryCountry,
        FTU_TEST_FIELDS.toCurrency,
        expectedCountry ?? `No country mapping for ${toCurrency}`,
        actualCountry,
      );
    }

    const expectedCurrency = expectedCase.record.get(
      FTU_TEST_FIELDS.settledCurrency,
    );
    const actualCurrency = record.get(FTU_REPORT_FIELDS.currencyId);

    if (normalize(expectedCurrency) === "") {
      addInformation(
        FTU_REPORT_FIELDS.currencyId,
        FTU_TEST_FIELDS.settledCurrency,
        expectedCurrency,
        actualCurrency,
      );
    } else if (normalize(actualCurrency) !== normalize(expectedCurrency)) {
      addFailure(
        FTU_REPORT_FIELDS.currencyId,
        FTU_TEST_FIELDS.settledCurrency,
        expectedCurrency,
        actualCurrency,
      );
    }
  }

  private comparePurpose(
    expectedCase: FtuExpectedCase,
    record: ReconcileRecord,
    addFailure: (...parameters: Parameters<AddComparisonRemark>) => void,
    addInformation: (...parameters: Parameters<AddComparisonRemark>) => void,
  ): void {
    const expectedPurpose = expectedCase.record.get(
      FTU_TEST_FIELDS.purposeCode,
    );
    const expectedPurposeField =
      expectedCase.direction === "BUY_FCY"
        ? FTU_REPORT_FIELDS.inflowPurpose
        : FTU_REPORT_FIELDS.outflowPurpose;
    const oppositePurposeField =
      expectedCase.direction === "BUY_FCY"
        ? FTU_REPORT_FIELDS.outflowPurpose
        : FTU_REPORT_FIELDS.inflowPurpose;
    const actualPurpose = record.get(expectedPurposeField);
    const oppositePurpose = record.get(oppositePurposeField);

    if (normalize(expectedPurpose) === "") {
      addInformation(
        expectedPurposeField,
        FTU_TEST_FIELDS.purposeCode,
        expectedPurpose,
        actualPurpose,
      );
    } else if (normalize(actualPurpose) !== normalize(expectedPurpose)) {
      addFailure(
        expectedPurposeField,
        FTU_TEST_FIELDS.purposeCode,
        expectedPurpose,
        actualPurpose,
      );
    }

    if (normalize(oppositePurpose) !== "") {
      addFailure(
        oppositePurposeField,
        FTU_TEST_FIELDS.purposeCode,
        "ต้องว่างตามทิศทางธุรกรรม",
        oppositePurpose,
      );
    }
  }

  private compareAggregateAmount(
    expectedCase: FtuExpectedCase,
    matchedRecords: ReconcileRecord[],
    addFailure: (...parameters: Parameters<AddComparisonRemark>) => void,
    addInformation: (...parameters: Parameters<AddComparisonRemark>) => void,
  ): void {
    const expectedText = expectedCase.record.get(FTU_TEST_FIELDS.settledAmount);
    const expectedAmount = parseAmount(expectedText);
    const actualAmounts = matchedRecords.map((record) =>
      parseAmount(record.get(FTU_REPORT_FIELDS.foreignCurrencyAmount)),
    );

    if (normalize(expectedText) === "") {
      addInformation(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_TEST_FIELDS.settledAmount,
        expectedText,
        actualAmounts
          .map((amount) => String(amount ?? "Invalid amount"))
          .join(" + "),
      );

      return;
    }

    if (
      expectedAmount === null ||
      actualAmounts.some((amount) => amount === null)
    ) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_TEST_FIELDS.settledAmount,
        expectedText,
        actualAmounts
          .map((amount) => String(amount ?? "Invalid amount"))
          .join(" + "),
      );
      return;
    }

    const actualTotal = actualAmounts.reduce<number>(
      (total, amount) => total + (amount ?? 0),
      0,
    );

    if (Math.abs(expectedAmount - actualTotal) > FTU_AMOUNT_TOLERANCE) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_TEST_FIELDS.settledAmount,
        expectedAmount.toString(),
        actualTotal.toString(),
      );
    }
  }

  private isAcceptedReportDate(expected: string, actual: string): boolean {
    const expectedDate = parseDate(expected);
    const actualDate = parseDate(actual);

    if (!expectedDate || !actualDate) {
      return false;
    }

    return (
      isSameDate(actualDate, expectedDate) ||
      isSameDate(actualDate, nextWeekday(expectedDate))
    );
  }

  private logSummary(outputPath: string, results: ResultRow[]): void {
    const passCount = results.filter(
      (result) => result.status === "PASS",
    ).length;
    const failCount = results.length - passCount;

    console.log(`Output File : ${outputPath}`);
    console.log(
      `Test Case : ${results.length} | Pass : ${passCount} | Fail : ${failCount}`,
    );
    console.log(
      "Reference reports in Remark are not used as reconcile keys or inputs.",
    );
  }
}

export const reconcileFtuReport = (testDataFilePath: string): Promise<string> =>
  new FtuReconcileService().reconcile(testDataFilePath);
