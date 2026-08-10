/**
 * olb-reconcile.ts
 * ------------------------------------------------------------------
 * Reconcile Test Data กับ AF1 Report สำหรับ DF_OLB
 *
 * หลักการสำคัญ:
 * 1. ใช้ Transaction ID/ Reconcile ID หา FI Arrangement Number ก่อน
 * 2. ห้ามหยุดตรวจทันทีเมื่อ Primary Key ว่าง/ไม่พบ/พบซ้ำ
 * 3. ใช้ Date, Amount, CIF No. และ CIF Name ช่วยเลือก Candidate
 * 4. ค่าว่างไม่นับเป็น Match และไม่ทำให้ Fail อัตโนมัติ
 * 5. Candidate คะแนนเสมอหรือหลักฐานไม่พอ จะไม่เลือกแถวเอง
 * 6. เมื่อเลือก Candidate ได้ ต้องตรวจทุก Matching/Core Field จนครบ
 * 7. สรุป PASS/FAIL หลังตรวจครบทุก Field เท่านั้น
 * ------------------------------------------------------------------
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { ReconcileRecord } from "../shared/record";
import { formatCompareRemark } from "../shared/remark";
import {
  ReconcileResultSheetWriter,
  ResultRow,
} from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import {
  OLB_AMOUNT_TOLERANCE,
  OLB_REMARKS,
  OLB_REPORT_CODE,
  OLB_REPORT_FIELDS,
  OLB_REPORT_HEADER_ROW,
  OLB_TEST_DATA_FIELDS,
  OLB_TEST_DATA_HEADER_ROW,
} from "./olb-config";
import {
  extractDateFromArrangementNumber,
  formatDate,
  isSameDate,
  normalizeId,
  normalizeText,
  parseAmount,
  parseDate,
} from "./olb-parse.util";

type CandidateMatchField =
  | "Date"
  | "Amount"
  | "CIF No."
  | "CIF Name";

type AddComparisonRemark = (
  reportField: string,
  testDataField: string,
  expected: string,
  actual: string,
) => void;

interface ScoredCandidate {
  record: ReconcileRecord;
  matchedFields: CandidateMatchField[];
  score: number;
}

interface CandidateResolution {
  matchedRecord?: ReconcileRecord;
  remark: string;
}

interface ResolvedCase {
  result: ResultRow;
  matchedRowNumber?: number;
}

const REQUIRED_TEST_DATA_HEADERS = [
  OLB_TEST_DATA_FIELDS.testNo,
  OLB_TEST_DATA_FIELDS.transactionId,
  OLB_TEST_DATA_FIELDS.transactionDate,
  OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
  OLB_TEST_DATA_FIELDS.cifNo,
  OLB_TEST_DATA_FIELDS.cifName,
] as const;

export class OlbReconcileService {
  async reconcile(testDataFilePath: string): Promise<string> {
    const workbookPreparer = new ReconcileWorkbookPreparer();
    const excelReader = new ReconcileExcelReader();
    const sheetWriter = new ReconcileResultSheetWriter();

    console.log(`\n===== RECONCILE - ${OLB_REPORT_CODE} =====`);

    const prepared = await workbookPreparer.prepare(
      OLB_REPORT_CODE,
      OLB_REPORT_HEADER_ROW,
    );

    const reportData = excelReader.parseWorksheet(
      prepared.reportWorksheet,
      OLB_REPORT_HEADER_ROW,
    );

    const testData = await excelReader.readFile(
      testDataFilePath,
      OLB_TEST_DATA_HEADER_ROW,
    );

    this.validateHeaders(prepared.reportHeaders, testData.headers);

    const reportRecords = reportData.records.filter(
      (record) =>
        normalizeText(
          record.get(OLB_REPORT_FIELDS.arrangementNumber),
        ) !== "",
    );

    /**
     * กัน Fallback ไปใช้ AF1 Row ที่มี Test Data แถวอื่นระบุ Primary Key ไว้แล้ว
     * แม้ Test Data แถวนั้นจะยังประมวลผลไม่ถึงก็ตาม
     */
    const reservedReportRowNumbers = this.findReservedReportRows(
      testData.records,
      reportRecords,
    );

    const usedReportRowNumbers = new Set<number>();
    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const results: ResultRow[] = [];

    for (const testDataRecord of testData.records) {
      const resolution = this.resolveRecord(
        testDataRecord,
        reportRecords,
        usedReportRowNumbers,
        reservedReportRowNumbers,
      );

      results.push(resolution.result);

      if (resolution.matchedRowNumber === undefined) {
        unmatchedRows.push(resolution.result);
        continue;
      }

      annotationByRowNumber.set(
        resolution.matchedRowNumber,
        resolution.result,
      );

      usedReportRowNumbers.add(resolution.matchedRowNumber);
    }

    sheetWriter.writeHeaderRow(prepared.resultSheet, prepared.reportHeaders);

    const nextRowNumber = sheetWriter.writeRowsInRequestedOrder(
      prepared.resultSheet,
      prepared.reportWorksheet,
      prepared.reportHeaders,
      OLB_REPORT_HEADER_ROW + 1,
      prepared.reportWorksheet.rowCount,
      annotationByRowNumber,
      unmatchedRows,
    );

    sheetWriter.finalizeAutoFilter(
      prepared.resultSheet,
      prepared.reportHeaders,
      nextRowNumber - 1,
    );

    prepared.workbook.removeWorksheet(prepared.reportWorksheet.id);
    await prepared.workbook.xlsx.writeFile(prepared.reconcileFilePath);

    this.logSummary(prepared.reconcileFilePath, results);
    return prepared.reconcileFilePath;
  }

  /** ตรวจ Header ที่จำเป็นก่อนเริ่ม Reconcile */
  private validateHeaders(
    reportHeaders: string[],
    testDataHeaders: string[],
  ): void {
    const reportName = requireMappingReportName(OLB_REPORT_CODE);
    const requiredReportHeaders = getUniqueMappingHeaders(reportName);

    this.assertHeaders(reportHeaders, requiredReportHeaders, "Raw Report");
    this.assertHeaders(
      testDataHeaders,
      REQUIRED_TEST_DATA_HEADERS,
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
        `[${OLB_REPORT_CODE}] ${sourceName} missing header(s): ` +
          missingHeaders.join(", "),
      );
    }
  }

  /** หา AF1 Row ที่ถูก Test Data แถวใดแถวหนึ่งระบุ Primary Key ไว้แล้ว */
  private findReservedReportRows(
    testDataRecords: ReconcileRecord[],
    reportRecords: ReconcileRecord[],
  ): Set<number> {
    const transactionIds = new Set(
      testDataRecords
        .map((record) =>
          normalizeText(record.get(OLB_TEST_DATA_FIELDS.transactionId)),
        )
        .filter((value) => value !== ""),
    );

    return new Set(
      reportRecords
        .filter((record) =>
          transactionIds.has(
            normalizeText(record.get(OLB_REPORT_FIELDS.arrangementNumber)),
          ),
        )
        .map((record) => record.rowNumber),
    );
  }

  /** วันที่ต้องตรงกับ Arrangement Contract Date หรือวันที่ใน Arrangement Number */
  private isReportDateMatch(
    expectedDate: Date,
    reportRecord: ReconcileRecord,
  ): boolean {
    const contractDate = parseDate(
      reportRecord.get(OLB_REPORT_FIELDS.arrangementContractDate),
    );

    if (contractDate && isSameDate(expectedDate, contractDate)) {
      return true;
    }

    const arrangementDate = extractDateFromArrangementNumber(
      reportRecord.get(OLB_REPORT_FIELDS.arrangementNumber),
    );

    return Boolean(
      arrangementDate && isSameDate(expectedDate, arrangementDate),
    );
  }

  /** คำนวณ Field ที่ Candidate ตรง โดยค่าว่างทั้งสองฝั่งจะไม่นับเป็น Match */
  private scoreCandidate(
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): ScoredCandidate {
    const matchedFields: CandidateMatchField[] = [];

    const expectedDate = parseDate(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.transactionDate),
    );

    if (expectedDate && this.isReportDateMatch(expectedDate, reportRecord)) {
      matchedFields.push("Date");
    }

    const expectedAmount = parseAmount(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
      ),
    );
    const actualAmount = parseAmount(
      reportRecord.get(OLB_REPORT_FIELDS.thbOutstandingAmount),
    );

    if (
      expectedAmount !== null &&
      actualAmount !== null &&
      Math.abs(expectedAmount - actualAmount) <= OLB_AMOUNT_TOLERANCE
    ) {
      matchedFields.push("Amount");
    }

    const expectedCifNo = normalizeId(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.cifNo),
    );
    const actualCustCode = normalizeId(
      reportRecord.get(OLB_REPORT_FIELDS.custCode),
    );

    if (
      expectedCifNo !== "" &&
      actualCustCode !== "" &&
      expectedCifNo === actualCustCode
    ) {
      matchedFields.push("CIF No.");
    }

    const expectedCifName = normalizeText(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.cifName),
    );
    const actualCustName = normalizeText(
      reportRecord.get(OLB_REPORT_FIELDS.custName),
    );

    if (
      expectedCifName !== "" &&
      actualCustName !== "" &&
      expectedCifName === actualCustName
    ) {
      matchedFields.push("CIF Name");
    }

    return {
      record: reportRecord,
      matchedFields,
      score: matchedFields.length,
    };
  }

  private rankCandidates(
    testDataRecord: ReconcileRecord,
    candidates: ReconcileRecord[],
  ): ScoredCandidate[] {
    return candidates
      .map((record) => this.scoreCandidate(testDataRecord, record))
      .sort((left, right) => right.score - left.score);
  }

  private formatCandidateList(candidates: ScoredCandidate[]): string {
    return candidates
      .slice(0, 10)
      .map((candidate) => {
        const matchedFields =
          candidate.matchedFields.length > 0
            ? candidate.matchedFields.join(" + ")
            : "ไม่มี Field ตรง";

        return (
          `Row ${candidate.record.rowNumber}: ` +
          `${candidate.record.get(OLB_REPORT_FIELDS.arrangementNumber)} ` +
          `[${matchedFields}]`
        );
      })
      .join(", ");
  }

  /**
   * หา Candidate โดยไม่หยุดเมื่อ Primary Key ว่าง/ไม่พบ/พบซ้ำ
   *
   * Fallback จะเลือกอัตโนมัติเมื่อ:
   * - Candidate คะแนนสูงสุดมีเพียง 1 แถว และ
   * - Amount ตรง หรือมี Field สำรองตรงอย่างน้อย 2 Field
   */
  private findBestCandidate(
    testDataRecord: ReconcileRecord,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): CandidateResolution {
    const transactionId = normalizeText(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.transactionId),
    );

    const unusedReportRecords = reportRecords.filter(
      (record) => !usedReportRowNumbers.has(record.rowNumber),
    );

    const primaryCandidates =
      transactionId === ""
        ? []
        : unusedReportRecords.filter(
            (record) =>
              normalizeText(
                record.get(OLB_REPORT_FIELDS.arrangementNumber),
              ) === transactionId,
          );

    if (primaryCandidates.length === 1) {
      return {
        matchedRecord: primaryCandidates[0],
        remark: "",
      };
    }

    if (primaryCandidates.length > 1) {
      const ranked = this.rankCandidates(testDataRecord, primaryCandidates);
      const bestScore = ranked[0]?.score ?? 0;
      const bestCandidates = ranked.filter(
        (candidate) => candidate.score === bestScore,
      );

      if (bestScore > 0 && bestCandidates.length === 1) {
        const best = bestCandidates[0];

        return {
          matchedRecord: best.record,
          remark:
            `Primary Matching พบหลายแถว (${primaryCandidates.length} แถว) ` +
            `แต่เลือก AF1 Row ${best.record.rowNumber} ได้จาก ` +
            best.matchedFields.join(" + "),
        };
      }

      return {
        remark:
          `Primary Matching พบหลายแถว (${primaryCandidates.length} แถว) ` +
          "และคะแนน Field อื่นสูงสุดเท่ากัน จึงไม่เลือก AF1 Row อัตโนมัติ\n" +
          this.formatCandidateList(bestCandidates),
      };
    }

    /**
     * Fallback ห้ามใช้ Row ที่ Test Data แถวอื่นมี Primary Key ตรงอยู่แล้ว
     * เพื่อป้องกันแถวที่ Transaction ID ว่างไปแย่ง Record ของ Test Case อื่น
     */
    const fallbackCandidates = unusedReportRecords.filter(
      (record) => !reservedReportRowNumbers.has(record.rowNumber),
    );

    const ranked = this.rankCandidates(testDataRecord, fallbackCandidates);
    const bestScore = ranked[0]?.score ?? 0;
    const bestCandidates = ranked.filter(
      (candidate) => candidate.score === bestScore,
    );

    if (bestScore === 0) {
      return {
        remark:
          "Fallback Matching ไม่พบ AF1 Row ที่มี Field สำรองตรงกับ Test Data",
      };
    }

    if (bestCandidates.length > 1) {
      return {
        remark:
          `Fallback Matching พบ Candidate คะแนนสูงสุดเท่ากัน ` +
          `(${bestCandidates.length} แถว, คะแนน ${bestScore}) ` +
          "จึงไม่เลือก AF1 Row อัตโนมัติ\n" +
          this.formatCandidateList(bestCandidates),
      };
    }

    const best = bestCandidates[0];
    const hasStrongEvidence =
      best.matchedFields.includes("Amount") || best.score >= 2;

    if (!hasStrongEvidence) {
      return {
        remark:
          `Fallback Matching พบ Candidate สูงสุดเพียง Row ` +
          `${best.record.rowNumber} แต่หลักฐานไม่เพียงพอ ` +
          `(ตรงเฉพาะ ${best.matchedFields.join(" + ")}) ` +
          "จึงไม่เลือก AF1 Row อัตโนมัติ",
      };
    }

    return {
      matchedRecord: best.record,
      remark:
        `Fallback Matching เลือก AF1 Row ${best.record.rowNumber} จาก ` +
        `${best.matchedFields.join(" + ")}\n` +
        `${OLB_REPORT_FIELDS.arrangementNumber} = ` +
        `"${best.record.get(OLB_REPORT_FIELDS.arrangementNumber)}"`,
    };
  }

  private resolveRecord(
    testDataRecord: ReconcileRecord,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): ResolvedCase {
    const testCaseNo = testDataRecord
      .get(OLB_TEST_DATA_FIELDS.testNo)
      .trim();

    const candidate = this.findBestCandidate(
      testDataRecord,
      reportRecords,
      usedReportRowNumbers,
      reservedReportRowNumbers,
    );

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

  /**
   * ไม่พบ Candidate ก็ยังสร้าง Remark ของทุก Matching/Core Field
   * เพื่อให้ User เห็น Test Data ที่ Script ใช้ค้นหาครบถ้วน
   */
  private buildUnresolvedResult(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    candidateRemark: string,
  ): ResolvedCase {
    const remarks = [candidateRemark];
    const reviewHeaders = new Set<string>();

    const comparisonPairs = [
      {
        testDataField: OLB_TEST_DATA_FIELDS.transactionId,
        reportField: OLB_REPORT_FIELDS.arrangementNumber,
      },
      {
        testDataField: OLB_TEST_DATA_FIELDS.transactionDate,
        reportField: OLB_REPORT_FIELDS.arrangementContractDate,
      },
      {
        testDataField: OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
        reportField: OLB_REPORT_FIELDS.thbOutstandingAmount,
      },
      {
        testDataField: OLB_TEST_DATA_FIELDS.cifNo,
        reportField: OLB_REPORT_FIELDS.custCode,
      },
      {
        testDataField: OLB_TEST_DATA_FIELDS.cifName,
        reportField: OLB_REPORT_FIELDS.custName,
      },
    ] as const;

    for (const pair of comparisonPairs) {
      const expected = testDataRecord.get(pair.testDataField);

      remarks.push(
        formatCompareRemark(
          OLB_REPORT_CODE,
          pair.testDataField,
          expected,
          pair.reportField,
          "",
        ),
      );

      if (expected.trim() === "") {
        reviewHeaders.add(pair.reportField);
      }
    }

    return {
      result: {
        testCaseNo,
        status: "FAIL",
        remark: remarks.filter((remark) => remark.trim() !== "").join("\n"),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [OLB_REPORT_FIELDS.arrangementNumber],
        reviewFieldHeaders: [...reviewHeaders],
        isExpectedAbsence: false,
      },
    };
  }

  /** ตรวจ Candidate ที่เลือกได้ทุก Field แล้วค่อยสรุป PASS/FAIL ตอนท้าย */
  private compareMatchedRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    candidateRemark: string,
  ): ResolvedCase {
    const failedHeaders = new Set<string>();
    const reviewHeaders = new Set<string>();
    const remarks: string[] = candidateRemark ? [candidateRemark] : [];

    const addFailure: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      failedHeaders.add(reportField);
      remarks.push(
        formatCompareRemark(
          OLB_REPORT_CODE,
          testDataField,
          expected,
          reportField,
          actual,
        ),
      );
    };

    const addReview: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      reviewHeaders.add(reportField);
      remarks.push(
        formatCompareRemark(
          OLB_REPORT_CODE,
          testDataField,
          expected,
          reportField,
          actual,
        ),
      );
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
      result: {
        testCaseNo,
        status,
        remark: finalRemark,
        matchedRowNumber: matchedRecord.rowNumber,
        failedKeyFieldHeaders: [...failedHeaders],
        reviewFieldHeaders: [...reviewHeaders],
        isExpectedAbsence: false,
      },
      matchedRowNumber: matchedRecord.rowNumber,
    };
  }

  private comparePrimaryKey(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const expected = testDataRecord.get(OLB_TEST_DATA_FIELDS.transactionId);
    const actual = matchedRecord.get(OLB_REPORT_FIELDS.arrangementNumber);

    if (normalizeText(expected) === "") {
      addReview(
        OLB_REPORT_FIELDS.arrangementNumber,
        OLB_TEST_DATA_FIELDS.transactionId,
        expected,
        actual,
      );
      return;
    }

    if (normalizeText(expected) !== normalizeText(actual)) {
      addFailure(
        OLB_REPORT_FIELDS.arrangementNumber,
        OLB_TEST_DATA_FIELDS.transactionId,
        expected,
        actual,
      );
    }
  }

  private compareDate(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const expectedText = testDataRecord.get(
      OLB_TEST_DATA_FIELDS.transactionDate,
    );
    const actualText = matchedRecord.get(
      OLB_REPORT_FIELDS.arrangementContractDate,
    );
    const expectedDate = parseDate(expectedText);

    if (!expectedDate) {
      addReview(
        OLB_REPORT_FIELDS.arrangementContractDate,
        OLB_TEST_DATA_FIELDS.transactionDate,
        expectedText,
        actualText,
      );
      return;
    }

    if (this.isReportDateMatch(expectedDate, matchedRecord)) {
      return;
    }

    const arrangementDate = extractDateFromArrangementNumber(
      matchedRecord.get(OLB_REPORT_FIELDS.arrangementNumber),
    );

    addFailure(
      OLB_REPORT_FIELDS.arrangementContractDate,
      OLB_TEST_DATA_FIELDS.transactionDate,
      expectedText,
      `${actualText} | FI Arrangement Number Date: ` +
        formatDate(arrangementDate),
    );
  }

  private compareAmount(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const expectedText = testDataRecord.get(
      OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
    );
    const actualText = matchedRecord.get(
      OLB_REPORT_FIELDS.thbOutstandingAmount,
    );
    const expectedAmount = parseAmount(expectedText);
    const actualAmount = parseAmount(actualText);

    if (expectedAmount === null) {
      addReview(
        OLB_REPORT_FIELDS.thbOutstandingAmount,
        OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
        expectedText,
        actualText,
      );
      return;
    }

    if (
      actualAmount === null ||
      Math.abs(expectedAmount - actualAmount) > OLB_AMOUNT_TOLERANCE
    ) {
      addFailure(
        OLB_REPORT_FIELDS.thbOutstandingAmount,
        OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
        expectedText,
        actualText,
      );
    }
  }

  private compareCifNo(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const expected = testDataRecord.get(OLB_TEST_DATA_FIELDS.cifNo);
    const actual = matchedRecord.get(OLB_REPORT_FIELDS.custCode);

    if (
      normalizeId(expected) === "" ||
      normalizeId(actual) === "" ||
      normalizeId(expected) !== normalizeId(actual)
    ) {
      addReview(
        OLB_REPORT_FIELDS.custCode,
        OLB_TEST_DATA_FIELDS.cifNo,
        expected,
        actual,
      );
    }
  }

  private compareCifName(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const expected = testDataRecord.get(OLB_TEST_DATA_FIELDS.cifName);
    const actual = matchedRecord.get(OLB_REPORT_FIELDS.custName);

    if (
      normalizeText(expected) === "" ||
      normalizeText(actual) === "" ||
      normalizeText(expected) !== normalizeText(actual)
    ) {
      addReview(
        OLB_REPORT_FIELDS.custName,
        OLB_TEST_DATA_FIELDS.cifName,
        expected,
        actual,
      );
    }
  }

  private logSummary(outputPath: string, results: ResultRow[]): void {
    const passCount = results.filter(
      (result) => result.status === "PASS",
    ).length;
    const failCount = results.length - passCount;

    console.log(`Output File : ${outputPath}`);
    console.log(
      `Test Case : ${results.length} | ` +
        `Pass : ${passCount} | ` +
        `Fail : ${failCount}`,
    );
  }
}

export const reconcileOlbReport = (
  testDataFilePath: string,
): Promise<string> => new OlbReconcileService().reconcile(testDataFilePath);
