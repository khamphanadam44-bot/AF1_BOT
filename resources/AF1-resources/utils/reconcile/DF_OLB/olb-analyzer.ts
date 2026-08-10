/**
 * olb-analyzer.ts
 * ------------------------------------------------------------------
 * ตรวจ Candidate ที่ Matcher เลือกมา แล้วสร้าง PASS/FAIL/Remark
 * ไม่ค้นหา Candidate และไม่อ่านหรือเขียน Excel Workbook
 * ------------------------------------------------------------------
 */

import { ReconcileRecord } from "../shared/record";
import { formatCompareRemark } from "../shared/remark";
import { ResultRow } from "../shared/result-writer";
import {
  extractDateFromArrangementNumber,
  formatDate,
  isDateMatchWithArrangementFallback,
  parseAmount,
  parseDate,
} from "../shared/reconcile-parse.util";
import {
  OLB_AMOUNT_TOLERANCE,
  OLB_FIELD_MAPPINGS,
  OLB_REMARKS,
  OLB_REPORT_CODE,
  OLB_REPORT_FIELDS,
  OLB_TEST_DATA_FIELDS,
  OlbFieldMapping,
} from "./olb-config";
import { OlbCandidateResolution } from "./olb-matcher";
import {
  normalizeOlbId,
  normalizeOlbText,
} from "./olb-normalize.util";

type AddComparisonRemark = (
  mapping: OlbFieldMapping,
  expected: string,
  actual: string,
) => void;

export class OlbAnalyzer {
  analyze(
    testDataRecord: ReconcileRecord,
    candidate: OlbCandidateResolution,
  ): ResultRow {
    const testCaseNo = testDataRecord
      .get(OLB_TEST_DATA_FIELDS.testNo)
      .trim();

    if (!candidate.matchedRecord) {
      return this.buildUnresolvedResult(
        testCaseNo,
        testDataRecord,
        candidate.remark,
      );
    }

    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      candidate.matchedRecord,
      candidate.remark,
    );
  }

  /** ไม่พบ Candidate ก็ยังสร้าง Remark ของทุก Matching/Core Field */
  private buildUnresolvedResult(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    candidateRemark: string,
  ): ResultRow {
    const remarks = [candidateRemark];
    const reviewHeaders = new Set<string>();

    for (const mapping of Object.values(OLB_FIELD_MAPPINGS)) {
      const expected = testDataRecord.get(mapping.testDataField);

      remarks.push(
        formatCompareRemark(
          OLB_REPORT_CODE,
          mapping.testDataField,
          expected,
          mapping.reportField,
          "",
        ),
      );

      if (expected.trim() === "") {
        reviewHeaders.add(mapping.reportField);
      }
    }

    return {
      testCaseNo,
      status: "FAIL",
      remark: remarks.filter((remark) => remark.trim() !== "").join("\n"),
      matchedRowNumber: undefined,
      failedKeyFieldHeaders: [OLB_REPORT_FIELDS.arrangementNumber],
      reviewFieldHeaders: [...reviewHeaders],
      isExpectedAbsence: false,
    };
  }

  /** ตรวจทุก Field แล้วจึงสรุป PASS/FAIL ตอนท้าย */
  private compareMatchedRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    candidateRemark: string,
  ): ResultRow {
    const failedHeaders = new Set<string>();
    const reviewHeaders = new Set<string>();
    const remarks: string[] = candidateRemark ? [candidateRemark] : [];

    const addFailure: AddComparisonRemark = (
      mapping,
      expected,
      actual,
    ): void => {
      failedHeaders.add(mapping.reportField);
      remarks.push(this.buildRemark(mapping, expected, actual));
    };

    const addReview: AddComparisonRemark = (
      mapping,
      expected,
      actual,
    ): void => {
      reviewHeaders.add(mapping.reportField);
      remarks.push(this.buildRemark(mapping, expected, actual));
    };

    this.comparePrimaryKey(
      testDataRecord,
      matchedRecord,
      addFailure,
      addReview,
    );
    this.compareDate(testDataRecord, matchedRecord, addFailure, addReview);
    this.compareAmount(testDataRecord, matchedRecord, addFailure, addReview);
    this.compareCifNo(testDataRecord, matchedRecord, addReview);
    this.compareCifName(testDataRecord, matchedRecord, addReview);

    const status = failedHeaders.size === 0 ? "PASS" : "FAIL";

    if (reviewHeaders.size > 0) {
      remarks.push(OLB_REMARKS.pleaseReview);
    }

    const finalRemark = [
      status === "PASS" ? OLB_REMARKS.matched : "",
      ...remarks,
    ]
      .filter((remark) => remark.trim() !== "")
      .join("\n");

    return {
      testCaseNo,
      status,
      remark: finalRemark,
      matchedRowNumber: matchedRecord.rowNumber,
      failedKeyFieldHeaders: [...failedHeaders],
      reviewFieldHeaders: [...reviewHeaders],
      isExpectedAbsence: false,
    };
  }

  private buildRemark(
    mapping: OlbFieldMapping,
    expected: string,
    actual: string,
  ): string {
    return formatCompareRemark(
      OLB_REPORT_CODE,
      mapping.testDataField,
      expected,
      mapping.reportField,
      actual,
    );
  }

  private comparePrimaryKey(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const mapping = OLB_FIELD_MAPPINGS.primary;
    const expected = testDataRecord.get(mapping.testDataField);
    const actual = matchedRecord.get(mapping.reportField);

    if (normalizeOlbText(expected) === "") {
      addReview(mapping, expected, actual);
      return;
    }

    if (normalizeOlbText(expected) !== normalizeOlbText(actual)) {
      addFailure(mapping, expected, actual);
    }
  }

  private compareDate(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const mapping = OLB_FIELD_MAPPINGS.date;
    const expectedText = testDataRecord.get(mapping.testDataField);
    const actualText = matchedRecord.get(mapping.reportField);
    const expectedDate = parseDate(expectedText);

    if (!expectedDate) {
      addReview(mapping, expectedText, actualText);
      return;
    }

    if (
      isDateMatchWithArrangementFallback(
        expectedDate,
        actualText,
        matchedRecord.get(OLB_REPORT_FIELDS.arrangementNumber),
      )
    ) {
      return;
    }

    const arrangementDate = extractDateFromArrangementNumber(
      matchedRecord.get(OLB_REPORT_FIELDS.arrangementNumber),
    );

    addFailure(
      mapping,
      expectedText,
      `${actualText} | FI Arrangement Number Date: ${formatDate(
        arrangementDate,
      )}`,
    );
  }

  private compareAmount(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const mapping = OLB_FIELD_MAPPINGS.amount;
    const expectedText = testDataRecord.get(mapping.testDataField);
    const actualText = matchedRecord.get(mapping.reportField);
    const expectedAmount = parseAmount(expectedText);
    const actualAmount = parseAmount(actualText);

    if (expectedAmount === null) {
      addReview(mapping, expectedText, actualText);
      return;
    }

    if (
      actualAmount === null ||
      Math.abs(expectedAmount - actualAmount) > OLB_AMOUNT_TOLERANCE
    ) {
      addFailure(mapping, expectedText, actualText);
    }
  }

  private compareCifNo(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const mapping = OLB_FIELD_MAPPINGS.cifNo;
    const expected = testDataRecord.get(mapping.testDataField);
    const actual = matchedRecord.get(mapping.reportField);

    if (
      normalizeOlbId(expected) === "" ||
      normalizeOlbId(actual) === "" ||
      normalizeOlbId(expected) !== normalizeOlbId(actual)
    ) {
      addReview(mapping, expected, actual);
    }
  }

  private compareCifName(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const mapping = OLB_FIELD_MAPPINGS.cifName;
    const expected = testDataRecord.get(mapping.testDataField);
    const actual = matchedRecord.get(mapping.reportField);

    if (
      normalizeOlbText(expected) === "" ||
      normalizeOlbText(actual) === "" ||
      normalizeOlbText(expected) !== normalizeOlbText(actual)
    ) {
      addReview(mapping, expected, actual);
    }
  }
}