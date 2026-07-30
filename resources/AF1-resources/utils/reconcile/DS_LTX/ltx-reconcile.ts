/**
 * ReconcileService (Main Orchestrator)
 *
 * โครงสร้าง Class ที่เกี่ยวข้อง:
 *   getReconcileConfig          -> อ่านและตรวจ Config ของ LTX
 *   ReconcileWorkbookPreparer   -> Copy Report + สร้าง Sheet ผลลัพธ์
 *   ReconcileExcelReader        -> อ่าน Excel (Report/Test Data) เป็น ReconcileRecord[]
 *   IExpectedCaseBuilder        -> Test Data -> Expected Case (strategy ต่อ report)
 *   ReconcileMatcher            -> จับคู่ Test Case กับ Report row (Reference Number ตรง ๆ)
 *   FieldRuleValidatorSet       -> เทียบ field อื่นที่ไม่ใช่ 3 หัวข้อหลัก (ผล = ไฮไลท์เหลืองเท่านั้น)
 *   ReconcileResultSheetWriter  -> เขียนผลลง Sheet (copy ทุกแถวจาก AF1 Report + แปะ annotation)
 *
*  Test_Result Pass/Fail ตัดสินจาก "3 หัวข้อ" เท่านั้น:
 *   1. Reference Transaction Number (Report) vs Transaction ID/ Reconcile ID (Test Data)
 *   2. FI Arrangement Number (Report) vs From Account (A/C Client/Sender) (Test Data)
 *   3. Transaction Amount (Report) vs SUM ของ Fee Amount ทุกรายการ (แถว FE) หรือ
 *      From Transfer Amount/Debit Amount (แถว DR)
 * ครบทั้ง 3 ข้อ = Pass, ข้อใดข้อหนึ่งไม่ตรง = Fail (checkKeyConditions ด้านล่าง)
 *
 * ส่วน field อื่นทั้งหมด (Payment Method, Currency, Beneficiary Name, ฯลฯ) ยังคงเทียบตาม
 * Requirement เดิมผ่าน FieldRuleValidatorSet เหมือนเดิมทุกอย่าง แต่ผลลัพธ์ตอนนี้ไม่ว่าจะ
 * เป็น FAIL หรือ REVIEW (ตามตรรกะเดิมของ FieldComparer) จะถูกปฏิบัติเหมือนกันหมด คือ
 * "ไฮไลท์เหลือง + ใส่ข้อความใน Remark" เท่านั้น ไม่กระทบ Pass/Fail ของแถวอีกต่อไป
 * ------------------------------------------------------------------
 */
import { getReconcileConfig, ReconcileReportConfig } from "./ltx-config";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { ReconcileMatcher } from "../shared/record-matcher";
import { AmountComparator } from "./amount-compare";
import { FieldRuleValidatorSet } from "./field-validator";
import {
  ReconcileResultSheetWriter,
  ResultRow,
  RowStatus,
} from "../shared/result-writer";
import { ExpectedCase, IExpectedCaseBuilder } from "./expected-case";
import { LtxExpectedCaseBuilder } from "./expected-case-builder";
import { ReconcileRecord } from "../shared/record";
import {
  getPresenceRuleFields,
  ReportPresenceRuleEvaluator,
} from "../shared/presence-rule";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import {
  getMappingHeaderRowNumber,
  getMappingMatchingKeyHeaders,
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import type { ReportCode } from "../../../config/report-config";

/**
 * เลือก ExpectedCaseBuilder ที่ถูกต้องให้แต่ละ report — จุดเดียวที่ต้องแก้เมื่อเพิ่ม report ใหม่
 *
 * Bug fix (Code Review): เดิม Type เป็น Record<string, ...> กว้างเกินไป เปลี่ยนเป็น
 * Partial<Record<ReportCode, ...>> เพื่อให้ TypeScript ช่วยจับ Typo ตอนเพิ่ม Report ใหม่
 * ตั้งแต่ Compile Time (สอดคล้องกับ Pattern เดียวกับ RECONCILE_CONFIG ใน ltx-config.ts)
 */
const EXPECTED_CASE_BUILDERS: Partial<Record<ReportCode, IExpectedCaseBuilder>> = {
  DS_LTX: new LtxExpectedCaseBuilder(),
};

export class ReconcileService {
  constructor(
    private readonly workbookPreparer: ReconcileWorkbookPreparer = new ReconcileWorkbookPreparer(),
    private readonly excelReader: ReconcileExcelReader = new ReconcileExcelReader(),
    private readonly matcher: ReconcileMatcher = new ReconcileMatcher(),
    private readonly fieldValidatorSet: FieldRuleValidatorSet = new FieldRuleValidatorSet(),
    private readonly sheetWriter: ReconcileResultSheetWriter = new ReconcileResultSheetWriter(),
    private readonly amountComparator: AmountComparator = new AmountComparator(),
    private readonly presenceRuleEvaluator: ReportPresenceRuleEvaluator = new ReportPresenceRuleEvaluator(),
  ) {}

  private validateMappingConfiguration(
    reportCode: string,
    config: ReconcileReportConfig,
  ): {
    mappingHeaderRowNumber: number;
    matchingKeyHeader: string;
  } {
    const mappingReportName = requireMappingReportName(reportCode);
    const mappingHeaderRowNumber = getMappingHeaderRowNumber(mappingReportName);
    const [matchingKeyHeader] = getMappingMatchingKeyHeaders(mappingReportName);

    if (!matchingKeyHeader) {
      throw new Error(
        `[${reportCode}] Mapping Config ไม่มี Matching Key Header.`,
      );
    }

    if (mappingHeaderRowNumber !== config.headerRowNumber) {
      throw new Error(
        `[${reportCode}] Header Row Config mismatch.\n` +
          `mapping-config.ts = ${mappingHeaderRowNumber}\n` +
          `ltx-config.ts = ${config.headerRowNumber}`,
      );
    }

    if (
      canonicalHeader(matchingKeyHeader) !==
      canonicalHeader(config.referenceNumberReportField)
    ) {
      throw new Error(
        `[${reportCode}] Matching Key Config mismatch.\n` +
          `mapping-config.ts = "${matchingKeyHeader}"\n` +
          `ltx-config.ts = "${config.referenceNumberReportField}"`,
      );
    }

    const configuredHeaders = new Set(
      getUniqueMappingHeaders(mappingReportName).map((header) =>
        canonicalHeader(header),
      ),
    );

    const unknownRuleHeaders = config.fieldRules
      .map((rule) => rule.reportField)
      .filter(
        (reportField) => !configuredHeaders.has(canonicalHeader(reportField)),
      );

    if (unknownRuleHeaders.length > 0) {
      throw new Error(
        `[${reportCode}] Reconcile Config อ้างถึง Report Header ` +
          "ที่ไม่มีใน Mapping Config: " +
          [...new Set(unknownRuleHeaders)].join(", "),
      );
    }

    return {
      mappingHeaderRowNumber,
      matchingKeyHeader,
    };
  }

  private getExpectedCaseBuilder(reportCode: string): IExpectedCaseBuilder {
    const builder = EXPECTED_CASE_BUILDERS[reportCode as ReportCode];
    if (!builder) {
      throw new Error(
        `ยังไม่มี ExpectedCaseBuilder สำหรับ report "${reportCode}" (ดู EXPECTED_CASE_BUILDERS)`,
      );
    }
    return builder;
  }

  private formatRuleRemark(remark: string, reportCode: string): string {
    return remark.replaceAll("{REPORT_CODE}", reportCode);
  }

  private validatePresenceRuleHeaders(
    reportCode: string,
    actualHeaders: string[],
    config: ReconcileReportConfig,
  ): void {
    const requiredRuleFields = getPresenceRuleFields(
      config.reportPresenceRules,
    );

    if (requiredRuleFields.length === 0) {
      return;
    }

    const normalizedActualHeaders = new Set(
      actualHeaders
        .filter((header) => header.trim() !== "")
        .map((header) => canonicalHeader(header)),
    );

    const missingHeaders = requiredRuleFields.filter(
      (requiredHeader) =>
        !normalizedActualHeaders.has(canonicalHeader(requiredHeader)),
    );

    if (missingHeaders.length === 0) {
      return;
    }

    throw new Error(
      `[${reportCode}] Test Data ไม่มี Header ` +
        "ที่จำเป็นสำหรับ Reconcile Presence Rule: " +
        missingHeaders.join(", "),
    );
  }

  private checkKeyConditions(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    suffix: string,
    matchedRecord: ReconcileRecord,
  ): {
    status: RowStatus;
    failedKeyFieldHeaders: string[];
    failRemarks: string[];
  } {
    const failedKeyFieldHeaders: string[] = [];
    const failRemarks: string[] = [];

    const expectedId = expectedCase.primaryRecord
      .get(config.testDataIdField)
      .trim()
      .toUpperCase();
    const actualReference = matchedRecord
      .get(config.referenceNumberReportField)
      .trim()
      .toUpperCase();
    const actualId = actualReference.endsWith(suffix.toUpperCase())
      ? actualReference.slice(0, actualReference.length - suffix.length)
      : actualReference;

    if (expectedId === "" || actualId !== expectedId) {
      failedKeyFieldHeaders.push(config.referenceNumberReportField);
      failRemarks.push(
        `[TS] : ${config.testDataIdField} = "${expectedId}" | ` +
          `[${reportCode}] : ${config.referenceNumberReportField} = "${actualReference}"`,
      );
    }

    const expectedAccount = expectedCase.primaryRecord
      .get(config.groupKeyFields.testDataAccountField)
      .trim();
    const actualAccount = matchedRecord
      .get(config.groupKeyFields.reportAccountField)
      .trim();

    if (
      expectedAccount === "" ||
      actualAccount === "" ||
      expectedAccount.toLowerCase() !== actualAccount.toLowerCase()
    ) {
      failedKeyFieldHeaders.push(config.groupKeyFields.reportAccountField);
      failRemarks.push(
        `[TS] : ${config.groupKeyFields.testDataAccountField} = "${expectedAccount}" | ` +
          `[${reportCode}] : ${config.groupKeyFields.reportAccountField} = "${actualAccount}"`,
      );
    }

    const actualAmount = matchedRecord.get(config.transactionAmountReportField);
    let isAmountOk: boolean;

    if (suffix === config.feSuffixLabel) {
      isAmountOk = this.amountComparator.compare(
        expectedCase.expectedFeAmount,
        actualAmount,
      ).isMatch;
    } else {
      const matchesPrimary = this.amountComparator.compare(
        expectedCase.primaryRecord.get(config.drAmountTestDataField),
        actualAmount,
      ).isMatch;
      const matchesFallback = this.amountComparator.compare(
        expectedCase.primaryRecord.get(config.drAmountFallbackTestDataField),
        actualAmount,
      ).isMatch;
      isAmountOk = matchesPrimary || matchesFallback;
    }

    if (!isAmountOk) {
      failedKeyFieldHeaders.push(config.transactionAmountReportField);
      if (suffix === config.feSuffixLabel) {
        failRemarks.push(
          `[TS] : SUM(Fee Amount Type 1-5) = "${expectedCase.expectedFeAmount}" | ` +
            `[${reportCode}] : ${config.transactionAmountReportField} = "${actualAmount}"`,
        );
      } else {
        const expectedTransferAmount = expectedCase.primaryRecord
          .get(config.drAmountTestDataField)
          .trim();
        const expectedDebitAmount = expectedCase.primaryRecord
          .get(config.drAmountFallbackTestDataField)
          .trim();
        failRemarks.push(
          `[TS] : ${config.drAmountTestDataField} = "${expectedTransferAmount}" ` +
            `/ ${config.drAmountFallbackTestDataField} = "${expectedDebitAmount}" | ` +
            `[${reportCode}] : ${config.transactionAmountReportField} = "${actualAmount}"`,
        );
      }
    }

    return {
      status: failedKeyFieldHeaders.length === 0 ? "PASS" : "FAIL",
      failedKeyFieldHeaders,
      failRemarks,
    };
  }

  private buildReviewSection(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    matchedRecord: ReconcileRecord,
    suffix: string,
  ): { reviewFieldHeaders: string[]; remark: string } {
    const fieldResults = this.fieldValidatorSet.validateAll(
      reportCode,
      config.fieldRules,
      expectedCase.primaryRecord,
      matchedRecord,
      suffix,
    );

    const mismatches = fieldResults.filter(
      (result) => result.status !== "PASS",
    );

    if (mismatches.length === 0) {
      return { reviewFieldHeaders: [], remark: "" };
    }

    const remark = `${mismatches.map((result) => result.remark).join("\n")}\nPlease review`;

    return {
      reviewFieldHeaders: mismatches.map((result) => result.fieldHeader),
      remark,
    };
  }

  private resolveSlot(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    expectedReference: string,
    suffix: string,
    reportRecords: ReconcileRecord[],
  ): ResultRow {
    const presenceDecision = this.presenceRuleEvaluator.evaluate(
      expectedCase.primaryRecord,
      config.reportPresenceRules,
    );

    const matchedRecord = this.matcher.findByExactReference(
      reportRecords,
      config.referenceNumberReportField,
      expectedReference,
    );

    if (presenceDecision.expectation === "MUST_NOT_EXIST") {
      if (!matchedRecord) {
        return {
          testCaseNo: expectedCase.displayTestCaseNo,
          status: "PASS",
          remark: this.formatRuleRemark(
            presenceDecision.passRemark,
            reportCode,
          ),
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: [],
          isExpectedAbsence: true,
        };
      }

      return {
        testCaseNo: expectedCase.displayTestCaseNo,
        status: "FAIL",
        remark:
          this.formatRuleRemark(presenceDecision.failRemark, reportCode) +
          `\n[${reportCode}] : ${config.referenceNumberReportField} = ` +
          `"${expectedReference}"`,
        matchedRowNumber: matchedRecord.rowNumber,
        failedKeyFieldHeaders: [config.referenceNumberReportField],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    if (!matchedRecord) {
      return {
        testCaseNo: expectedCase.displayTestCaseNo,
        status: "FAIL",
        remark:
          `ไม่พบแถว ${suffix} ใน ${reportCode} ที่ ` +
          `${config.referenceNumberReportField} = "${expectedReference}"`,
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    const { status, failedKeyFieldHeaders, failRemarks } =
      this.checkKeyConditions(
        reportCode,
        config,
        expectedCase,
        suffix,
        matchedRecord,
      );
    const { reviewFieldHeaders, remark: reviewRemark } =
      this.buildReviewSection(
        reportCode,
        config,
        expectedCase,
        matchedRecord,
        suffix,
      );
    const remarkLines = [
      ...failRemarks,
      ...(reviewRemark.trim() !== "" ? [reviewRemark] : []),
    ];
    return {
      testCaseNo: expectedCase.displayTestCaseNo,
      status,
      remark: remarkLines.join("\n"),
      matchedRowNumber: matchedRecord.rowNumber,
      failedKeyFieldHeaders,
      reviewFieldHeaders,
      isExpectedAbsence: false,
    };
  }

  async reconcile(
    reportCode: string,
    testDataFilePath: string,
  ): Promise<string> {
    console.log(`\n===== RECONCILE - ${reportCode} =====`);

    const config = getReconcileConfig(reportCode);
    const { mappingHeaderRowNumber } = this.validateMappingConfiguration(
      reportCode,
      config,
    );
    const expectedCaseBuilder = this.getExpectedCaseBuilder(reportCode);

    const {
      workbook,
      reportWorksheet,
      resultSheet,
      reportHeaders,
      reconcileFilePath,
    } = await this.workbookPreparer.prepare(reportCode, mappingHeaderRowNumber);

    const { records: reportRecords } = this.excelReader.parseWorksheet(
      reportWorksheet,
      mappingHeaderRowNumber,
    );
    const testData = await this.excelReader.readFile(
      testDataFilePath,
      config.testDataHeaderRowNumber,
    );

    this.validatePresenceRuleHeaders(reportCode, testData.headers, config);

    const expectedCases = expectedCaseBuilder.build(
      testData.headers,
      testData.records,
      config,
    );

    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];

    const collect = (row: ResultRow): void => {
      if (row.matchedRowNumber !== undefined) {
        annotationByRowNumber.set(row.matchedRowNumber, row);
      } else {
        unmatchedRows.push(row);
      }
    };

    expectedCases.forEach((expectedCase) => {
      if (expectedCase.expectedDrReference) {
        collect(
          this.resolveSlot(
            reportCode,
            config,
            expectedCase,
            expectedCase.expectedDrReference,
            config.drSuffixLabel,
            reportRecords,
          ),
        );
      }

      if (expectedCase.expectedFeReference) {
        collect(
          this.resolveSlot(
            reportCode,
            config,
            expectedCase,
            expectedCase.expectedFeReference,
            config.feSuffixLabel,
            reportRecords,
          ),
        );
      }
    });

    const unexpectedMissingRows = unmatchedRows.filter(
      (row) => row.status === "FAIL",
    );

    if (unexpectedMissingRows.length > 0) {
      console.warn(
        `⚠️ พบ ${unexpectedMissingRows.length} แถวที่หาคู่กันใน ${reportCode} ` +
          `ไม่เจอและไม่เข้าเงื่อนไข Expected Absence`,
      );
    }

    this.sheetWriter.writeHeaderRow(resultSheet, reportHeaders);

    const lastRowNumber = this.sheetWriter.writeRowsInRequestedOrder(
      resultSheet,
      reportWorksheet,
      reportHeaders,
      mappingHeaderRowNumber + 1,
      reportWorksheet.rowCount,
      annotationByRowNumber,
      unmatchedRows,
    );

    this.sheetWriter.finalizeAutoFilter(
      resultSheet,
      reportHeaders,
      lastRowNumber - 1,
    );

    workbook.removeWorksheet(reportWorksheet.id);

    await workbook.xlsx.writeFile(reconcileFilePath);

    const allRows = [...annotationByRowNumber.values(), ...unmatchedRows];
    const totalPass = allRows.filter((row) => row.status === "PASS").length;
    const totalFail = allRows.filter((row) => row.status === "FAIL").length;
    const totalWithReview = allRows.filter(
      (row) => row.reviewFieldHeaders.length > 0,
    ).length;

    console.log(`Output File : ${reconcileFilePath}`);
    console.log(
      `Test Case ที่ประมวลผล : ${expectedCases.length} | แถวผลลัพธ์ทั้งหมด : ${allRows.length} | ` +
        `Pass : ${totalPass} | Fail : ${totalFail} | (ในจำนวนนี้มี field สีเหลืองต้อง review : ${totalWithReview} แถว)`,
    );

    return reconcileFilePath;
  }
}

/**
 * Backward-compatible function wrapper — ให้ tests/script3-reconcile.spec.ts เดิม
 * เรียกใช้ได้โดยไม่ต้องแก้ไฟล์ test เลย
 */
export const reconcileReport = (
  reportCode: string,
  testDataFilePath: string,
): Promise<string> => {
  return new ReconcileService().reconcile(reportCode, testDataFilePath);
};
