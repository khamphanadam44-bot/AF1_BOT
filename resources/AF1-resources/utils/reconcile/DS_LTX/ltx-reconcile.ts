/**
 * ควบคุม Flow Reconcile ของ DS_LTX
 *
 * 1. สร้าง Expected Case จาก Test Data
 * 2. สงวน Report Row ที่มี Exact Reference
 * 3. จับคู่ Exact ก่อน Fallback และไม่ใช้ Report Row ซ้ำ
 * 4. ตรวจ Reference, Account และ Amount เพื่อกำหนด PASS/FAIL
 *
 * Field อื่นใช้สำหรับ Review และไม่เปลี่ยนสถานะหลักของ Test Case
 */

import { getHandledReportFields, getReconcileConfig } from "./ltx-config";
import type { ReconcileReportConfig } from "./ltx-config";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { AmountComparator } from "./ltx-amount-compare";
import { FieldRuleValidatorSet } from "./ltx-field-validator";
import { LtxMatcher } from "./ltx-matcher";
import type { LtxMatchResult } from "./ltx-matcher";
import { ReconcileResultSheetWriter } from "../shared/result-writer";
import type { ResultRow, RowStatus } from "../shared/result-writer";
import {
  LtxExpectedCaseBuilder,
  type ExpectedCase,
} from "./ltx-expected-case-builder";
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

type CollectedResultRow = ResultRow & {
  reserveReportRow: boolean;
};

export class ReconcileService {
  constructor(
    private workbookPreparer: ReconcileWorkbookPreparer = new ReconcileWorkbookPreparer(),
    private excelReader: ReconcileExcelReader = new ReconcileExcelReader(),
    private matcher: LtxMatcher = new LtxMatcher(),
    private fieldValidatorSet: FieldRuleValidatorSet = new FieldRuleValidatorSet(),
    private sheetWriter: ReconcileResultSheetWriter = new ReconcileResultSheetWriter(),
    private amountComparator: AmountComparator = new AmountComparator(),
    private presenceRuleEvaluator: ReportPresenceRuleEvaluator = new ReportPresenceRuleEvaluator(),
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

    const unknownOutputOnlyHeaders = config.outputOnlyFields
      .filter(
        (reportField) => !configuredHeaders.has(canonicalHeader(reportField)),
      );

    if (unknownOutputOnlyHeaders.length > 0) {
      throw new Error(
        `[${reportCode}] Output Only Config อ้างถึง Report Header ` +
          "ที่ไม่มีใน Mapping Config: " +
          [...new Set(unknownOutputOnlyHeaders)].join(", "),
      );
    }

    const evaluatedHeaders = new Set(
      [
        config.referenceNumberReportField,
        config.groupKeyFields.reportAccountField,
        config.transactionAmountReportField,
        ...config.fieldRules.map((rule) => rule.reportField),
      ].map((header) => canonicalHeader(header)),
    );

    const conflictingOutputOnlyHeaders = config.outputOnlyFields
      .filter((reportField) =>
        evaluatedHeaders.has(canonicalHeader(reportField)),
      );

    if (conflictingOutputOnlyHeaders.length > 0) {
      throw new Error(
        `[${reportCode}] Report Header ถูกกำหนดเป็นทั้ง Field ที่ตรวจ ` +
          "และ Output Only: " +
          [...new Set(conflictingOutputOnlyHeaders)].join(", "),
      );
    }

    const handledHeaders = new Set(
      getHandledReportFields(config).map((header) => canonicalHeader(header)),
    );

    const uncoveredMappingHeaders = getUniqueMappingHeaders(
      mappingReportName,
    ).filter((header) => !handledHeaders.has(canonicalHeader(header)));

    if (uncoveredMappingHeaders.length > 0) {
      throw new Error(
        `[${reportCode}] Mapping Field ยังไม่ได้กำหนดวิธีตรวจ ` +
          "หรือระบุเป็น Output Only: " +
          uncoveredMappingHeaders.join(", "),
      );
    }

    return {
      mappingHeaderRowNumber,
      matchingKeyHeader,
    };
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
    reviewFieldHeaders: string[];
    reviewRemarks: string[];
  } {
    const failedKeyFieldHeaders: string[] = [];
    const failRemarks: string[] = [];
    const reviewFieldHeaders: string[] = [];
    const reviewRemarks: string[] = [];

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

    if (expectedId === "") {
      reviewFieldHeaders.push(config.referenceNumberReportField);
      reviewRemarks.push(
        `Test Data ไม่มี ${config.testDataIdField} ` +
          `จึงไม่สามารถตรวจ Exact Reference ได้ และใช้ผลจาก LTX Fallback Matching`,
      );
    } else if (actualId !== expectedId) {
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
      isAmountOk = this.amountComparator.matches(
        expectedCase.expectedFeAmount,
        actualAmount,
      );
    } else {
      const matchesPrimary = this.amountComparator.matches(
        expectedCase.primaryRecord.get(config.drAmountTestDataField),
        actualAmount,
      );
      const matchesFallback = this.amountComparator.matches(
        expectedCase.primaryRecord.get(config.drAmountFallbackTestDataField),
        actualAmount,
      );
      isAmountOk = matchesPrimary || matchesFallback;
    }

    if (!isAmountOk) {
      failedKeyFieldHeaders.push(config.transactionAmountReportField);
      if (suffix === config.feSuffixLabel) {
        failRemarks.push(
          `[TS] : SUM(Fee Amount) = "${expectedCase.expectedFeAmount}" | ` +
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
      reviewFieldHeaders,
      reviewRemarks,
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

  /**
   * รวม Review Remark จาก Matching และ Field Compare
   * แล้วต่อท้าย Please review เพียงครั้งเดียว
   */
  private combineReviewRemarks(...remarks: string[]): string {
    const messages = remarks
      .map((remark) => remark.replace(/\n?Please review\s*$/i, "").trim())
      .filter((remark) => remark !== "");

    return messages.length === 0 ? "" : `${messages.join("\n")}\nPlease review`;
  }

  /**
   * เมื่อไม่มี Report Row ที่แน่นอน ระบบจะไม่สร้าง Actual Value ปลอม
   * แต่แจ้ง Field ที่ยังไม่สามารถเปรียบเทียบได้ตาม Slot ปัจจุบัน
   */
  private buildUnavailableFieldRemark(
    config: ReconcileReportConfig,
    suffix: string,
  ): string {
    const unavailableFields = [
      ...new Set(
        config.fieldRules
          .filter(
            (rule) =>
              !rule.applicableSuffixes ||
              rule.applicableSuffixes.includes(suffix),
          )
          .map((rule) => rule.reportField),
      ),
    ];

    if (unavailableFields.length === 0) {
      return "";
    }

    return (
      "ไม่สามารถเปรียบเทียบ Field จาก AF1 Report ได้ " +
      "เนื่องจากไม่มี Report row ที่แน่นอน: " +
      unavailableFields.join(", ")
    );
  }

  private removeFailedHeadersFromReview(
    reviewFieldHeaders: string[],
    failedKeyFieldHeaders: string[],
  ): string[] {
    const failedHeaders = new Set(
      failedKeyFieldHeaders.map((header) => canonicalHeader(header)),
    );

    return [...new Set(reviewFieldHeaders)].filter(
      (header) => !failedHeaders.has(canonicalHeader(header)),
    );
  }

  private buildUnresolvableCaseResult(
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
  ): CollectedResultRow {
    const testNo = expectedCase.primaryRecord
      .get(config.testDataTestNoField)
      .trim();
    const transactionId = expectedCase.primaryRecord
      .get(config.testDataIdField)
      .trim();
    const remarks: string[] = [];

    if (testNo === "") {
      remarks.push(`Test Data ไม่มี ${config.testDataTestNoField}`);
    }

    if (transactionId === "") {
      remarks.push(`Test Data ไม่มี ${config.testDataIdField}`);
    }

    remarks.push(
      "ไม่สามารถสร้าง DR หรือ FE Slot สำหรับ Reconcile ได้ " +
        "เพราะไม่มี Expected Reference และไม่มี Amount ที่มากกว่า 0",
    );

    return {
      testCaseNo: expectedCase.displayTestCaseNo,
      status: "FAIL",
      remark: remarks.join("\n"),
      matchedRowNumber: undefined,
      failedKeyFieldHeaders: [],
      reviewFieldHeaders: [],
      isExpectedAbsence: false,
      reserveReportRow: false,
    };
  }

  private resolveSlot(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    expectedReference: string,
    suffix: string,
    reportRecords: ReconcileRecord[],
    /**
     * แถว AF1 ที่ Expected Case อื่นจับคู่ไปแล้ว (ทั้ง DR และ FE ของ
     * Test Case อื่น) — ส่งต่อให้ LtxMatcher กรอง Candidate ออก
     * เพื่อกันการจับคู่ซ้ำข้าม Test Case (ดู ltx-matcher.ts)
     */
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): CollectedResultRow {
    const presenceDecision = this.presenceRuleEvaluator.evaluate(
      expectedCase.primaryRecord,
      config.reportPresenceRules,
    );

    const expectedId = expectedCase.primaryRecord
      .get(config.testDataIdField)
      .trim();
    const normalizedExpectedReference = expectedReference.trim();

    const testNo = expectedCase.primaryRecord
      .get(config.testDataTestNoField)
      .trim();

    const missingTestNoRemark =
      testNo === "" ? `Test Data ไม่มี ${config.testDataTestNoField}` : "";

    /**
     * ใส่ Comment กรณี Test No. ว่างให้ Result ทุกเส้นทาง
     * โดยไม่กระทบผล Matching และ PASS/FAIL
     */
    const withMissingTestNoRemark = (...remarks: string[]): string =>
      [missingTestNoRemark, ...remarks]
        .filter((remark) => remark.trim() !== "")
        .join("\n");

    let matchResult: LtxMatchResult;

    if (presenceDecision.expectation === "MUST_NOT_EXIST") {
      matchResult = this.matcher.findPresence(
        reportRecords,
        usedReportRowNumbers,
        expectedCase,
        config,
        normalizedExpectedReference,
        suffix,
      );
    } else {
      matchResult = this.matcher.findMatch(
        reportCode,
        reportRecords,
        usedReportRowNumbers,
        reservedReportRowNumbers,
        expectedCase,
        config,
        normalizedExpectedReference,
        suffix,
      );
    }

    const matchedRecord = matchResult.matchedRecord;

    const isUnresolvedMustExistMatch =
      presenceDecision.expectation !== "MUST_NOT_EXIST" &&
      (matchResult.strategy === "AMBIGUOUS_EXACT" ||
        matchResult.strategy === "AMBIGUOUS_FALLBACK" ||
        matchResult.strategy === "EXACT_ALREADY_USED" ||
        matchResult.strategy === "FALLBACK_ALREADY_USED");

    if (isUnresolvedMustExistMatch) {
      return {
        testCaseNo: expectedCase.displayTestCaseNo,
        status: "FAIL",
        remark: withMissingTestNoRemark(
          matchResult.informationalRemark,
          matchResult.failureRemark,
          this.buildUnavailableFieldRemark(config, suffix),
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
        reserveReportRow: false,
      };
    }

    /**
     * Presence Rule กำหนดว่า Test Data รายการนี้
     * ต้องไม่มีอยู่ใน AF1 Report
     */
    if (presenceDecision.expectation === "MUST_NOT_EXIST") {
      /**
       * ไม่พบ Report Row ตามที่ Business Rule คาดไว้
       * จึงถือว่า PASS แบบ Expected Absence
       *
       * ถ้า Transaction ID ว่าง
       * ให้บันทึก Comment เพิ่ม แต่ไม่เปลี่ยนผล
       * ของ Expected Absence
       */
      if (!matchedRecord) {
        const reviewRemark = this.combineReviewRemarks(
          matchResult.reviewRemark,
        );

        return {
          testCaseNo: expectedCase.displayTestCaseNo,

          status: "PASS",

          remark: withMissingTestNoRemark(
            this.formatRuleRemark(presenceDecision.passRemark, reportCode),
            matchResult.informationalRemark,
            reviewRemark,
            expectedId === ""
              ? `Test Data ไม่มี ${config.testDataIdField}`
              : "",
          ),

          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: matchResult.reviewFieldHeaders,
          isExpectedAbsence: true,
          reserveReportRow: false,
        };
      }

      /**
       * พบ Report Row ทั้งที่ Business Rule กำหนด
       * ว่ารายการนี้ต้องไม่มี จึงถือว่า FAIL
       */
      const matchedRowWasAlreadyUsed = usedReportRowNumbers.has(
        matchedRecord.rowNumber,
      );
      const alreadyUsedRemark = matchedRowWasAlreadyUsed
        ? `พบรายการที่ Report row ${matchedRecord.rowNumber} ` +
          "ซึ่งถูกจับคู่กับ Test Case อื่นแล้ว แต่ยังถือว่ารายการมีอยู่ใน Report"
        : "";
      const missingExpectedIdRemark =
        expectedId === "" ? `Test Data ไม่มี ${config.testDataIdField}` : "";

      return {
        testCaseNo: expectedCase.displayTestCaseNo,

        status: "FAIL",

        remark: withMissingTestNoRemark(
          this.formatRuleRemark(presenceDecision.failRemark, reportCode),

          `[${reportCode}] : ${config.referenceNumberReportField} = ` +
            `"${matchedRecord.get(config.referenceNumberReportField).trim()}"`,
          matchResult.informationalRemark,
          alreadyUsedRemark,
          missingExpectedIdRemark,
        ),
        matchedRowNumber: matchedRowWasAlreadyUsed
          ? undefined
          : matchedRecord.rowNumber,

        failedKeyFieldHeaders: [config.referenceNumberReportField],

        reviewFieldHeaders: [],
        isExpectedAbsence: false,
        reserveReportRow: false,
      };
    }

    if (!matchedRecord) {
      const fallbackReason =
        expectedId === ""
          ? `Test Data ไม่มี ${config.testDataIdField}`
          : `ไม่พบ Exact Reference = "${normalizedExpectedReference}"`;

      return {
        testCaseNo: expectedCase.displayTestCaseNo,

        status: "FAIL",

        remark: withMissingTestNoRemark(
          fallbackReason,
          matchResult.informationalRemark,
          `ไม่พบแถว ${suffix} ใน ${reportCode} ทั้งจาก Exact Reference ` +
            "และ LTX Fallback Matching " +
            "(Account + Currency + DR/FE + Amount)",
          this.buildUnavailableFieldRemark(config, suffix),
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
        reserveReportRow: false,
      };
    }

    const {
      status,
      failedKeyFieldHeaders,
      failRemarks,
      reviewFieldHeaders: keyReviewFieldHeaders,
      reviewRemarks: keyReviewRemarks,
    } = this.checkKeyConditions(
      reportCode,
      config,
      expectedCase,
      suffix,
      matchedRecord,
    );
    const {
      reviewFieldHeaders: fieldReviewHeaders,
      remark: fieldReviewRemark,
    } = this.buildReviewSection(
      reportCode,
      config,
      expectedCase,
      matchedRecord,
      suffix,
    );

    const reviewFieldHeaders = this.removeFailedHeadersFromReview(
      [
        ...matchResult.reviewFieldHeaders,
        ...keyReviewFieldHeaders,
        ...fieldReviewHeaders,
      ],
      failedKeyFieldHeaders,
    );

    const reviewRemark = this.combineReviewRemarks(
      matchResult.reviewRemark,
      ...keyReviewRemarks,
      fieldReviewRemark,
    );

    const remark = withMissingTestNoRemark(
      matchResult.informationalRemark,
      ...failRemarks,
      reviewRemark,
    );

    /**
     * Fallback ที่มี Transaction ID แต่ไปจับ Reference ของรายการอื่น
     * ต้องไม่จอง Report row เพราะแถวนั้นอาจเป็น Exact Reference
     * ของ Expected Case ถัดไป ส่วน Exact Match และ No-ID Fallback
     * ยังจองแถวได้ตามปกติ
     */
    const referenceFailed = failedKeyFieldHeaders.some(
      (header) =>
        canonicalHeader(header) ===
        canonicalHeader(config.referenceNumberReportField),
    );

    return {
      testCaseNo: expectedCase.displayTestCaseNo,
      status,
      remark,
      matchedRowNumber: matchedRecord.rowNumber,
      failedKeyFieldHeaders,
      reviewFieldHeaders,
      isExpectedAbsence: false,
      reserveReportRow: !referenceFailed,
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
    const expectedCaseBuilder = new LtxExpectedCaseBuilder();

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
    const reservedReportRowNumbers = this.matcher.findReservedReportRows(
      reportRecords,
      expectedCases,
      config,
    );

    const annotationByRowNumber = new Map<number, CollectedResultRow>();
    const unmatchedRows: ResultRow[] = [];

    /** แถว AF1 ที่จับคู่แล้วจะไม่ถูกนำไปใช้กับ MUST_EXIST Slot อื่น */
    const usedReportRowNumbers = new Set<number>();

    const moveResultToUnmatched = (
      row: CollectedResultRow,
      reportRowNumber: number,
    ): CollectedResultRow => {
      const referenceRemark =
        `ผลลัพธ์นี้อ้างถึง Report row ${reportRowNumber} ` +
        "แต่แสดงเป็น Unmatched Result เพื่อไม่ให้เขียนทับผลลัพธ์ของแถวเดียวกัน";

      return {
        ...row,
        matchedRowNumber: undefined,
        remark: [row.remark, referenceRemark]
          .filter((remark) => remark.trim() !== "")
          .join("\n"),
      };
    };

    const collect = (row: CollectedResultRow): void => {
      if (row.matchedRowNumber === undefined) {
        unmatchedRows.push(row);
        return;
      }

      const reportRowNumber = row.matchedRowNumber;
      const previous = annotationByRowNumber.get(reportRowNumber);

      if (!previous) {
        annotationByRowNumber.set(reportRowNumber, row);

        if (row.reserveReportRow) {
          usedReportRowNumbers.add(reportRowNumber);
        }

        return;
      }

      const previousReservesReportRow =
        usedReportRowNumbers.has(reportRowNumber);

      if (previousReservesReportRow && row.reserveReportRow) {
        throw new Error(
          `[${reportCode}] AF1 Row ${reportRowNumber} ถูกจับคู่ซ้ำ: ` +
            `"${previous.testCaseNo}" กับ "${row.testCaseNo}". ` +
            "กรุณาตรวจสอบ LtxMatcher และ usedReportRowNumbers",
        );
      }

      if (previousReservesReportRow) {
        unmatchedRows.push(moveResultToUnmatched(row, reportRowNumber));
        return;
      }

      if (row.reserveReportRow) {
        unmatchedRows.push(
          moveResultToUnmatched(
            previous,
            reportRowNumber,
          ),
        );
        annotationByRowNumber.set(reportRowNumber, row);
        usedReportRowNumbers.add(reportRowNumber);
        return;
      }

      unmatchedRows.push(moveResultToUnmatched(row, reportRowNumber));
    };

    for (const expectedCase of expectedCases) {
      const expectedDrReference = expectedCase.expectedDrReference ?? "";

      const expectedFeReference = expectedCase.expectedFeReference ?? "";

      /**
       * ExpectedCaseBuilder เป็นแหล่งตัดสินเพียงจุดเดียวว่า Test Data
       * ต้องมี DR/FE Slot หรือไม่ โดย Amount มากกว่า 0 ถือว่ามี Slot
       * ส่วน Amount Tolerance ใช้เฉพาะตอนเปรียบเทียบ Expected กับ Actual
       */
      const shouldResolveDr = expectedCase.hasExpectedDr;
      const shouldResolveFe = expectedCase.hasExpectedFe;

      if (shouldResolveDr) {
        collect(
          this.resolveSlot(
            reportCode,
            config,
            expectedCase,
            expectedDrReference,
            config.drSuffixLabel,
            reportRecords,
            usedReportRowNumbers,
            reservedReportRowNumbers,
          ),
        );
      }

      if (shouldResolveFe) {
        collect(
          this.resolveSlot(
            reportCode,
            config,
            expectedCase,
            expectedFeReference,
            config.feSuffixLabel,
            reportRecords,
            usedReportRowNumbers,
            reservedReportRowNumbers,
          ),
        );
      }

      if (!shouldResolveDr && !shouldResolveFe) {
        collect(this.buildUnresolvableCaseResult(config, expectedCase));
      }
    }

    const unmatchedFailRows = unmatchedRows.filter(
      (row) => row.status === "FAIL",
    );

    if (unmatchedFailRows.length > 0) {
      console.warn(
        `⚠️ พบผลลัพธ์ FAIL ${unmatchedFailRows.length} แถวใน ${reportCode} ` +
          "ที่ไม่มี Report row สำหรับเขียน Annotation โดยตรง",
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
      (row) =>
        row.reviewFieldHeaders.length > 0 || /Please review/i.test(row.remark),
    ).length;

    console.log(`Output File : ${reconcileFilePath}`);
    console.log(
      `Test Case ที่ประมวลผล : ${expectedCases.length} | แถวผลลัพธ์ทั้งหมด : ${allRows.length} | ` +
        `Pass : ${totalPass} | Fail : ${totalFail} | (ในจำนวนนี้มี field สีเหลืองต้อง review : ${totalWithReview} แถว)`,
    );

    return reconcileFilePath;
  }
}

export const reconcileReport = (
  reportCode: string,
  testDataFilePath: string,
): Promise<string> => {
  return new ReconcileService().reconcile(reportCode, testDataFilePath);
};