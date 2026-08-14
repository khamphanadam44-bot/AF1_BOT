/**
 * LtxMatcher
 * ------------------------------------------------------------------
 * จับคู่ Expected Case ของ DS_LTX กับ AF1 Report Row
 *
 * ลำดับ:
 * 1. Exact Reference (ค่าเต็มรวม suffix เช่น TX001DR หรือ TX001FE)
 * 2. Fallback: Account + Currency + DR/FE + Amount
 * 3. ถ้าพบหลายแถว ให้เลือกแถวที่ Key และ Supporting Field ตรงมากที่สุด
 * 4. ถ้าคะแนนสูงสุดเสมอกัน จะไม่เดาและคืนผล Ambiguous
 *
 * DR และ FE ที่มี Transaction ID ฐานเดียวกันไม่ถือว่าซ้ำ
 * เพราะค่า Exact Reference เต็มมี suffix ต่างกัน
 *
 * findMatch() ใช้ usedReportRowNumbers เพื่อกันแถว AF1 เดียวกันถูกจับคู่ซ้ำ
 * ส่วน findPresence() ตรวจจาก Report ทั้งหมด เพราะ Expected Absence ต้องตอบว่า
 * รายการมีอยู่จริงหรือไม่ โดยไม่ขึ้นกับว่าแถวนั้นถูก Test Case อื่นใช้แล้วหรือยัง
 *
 * ไฟล์นี้ไม่ตัดสิน PASS/FAIL ของ Test Case และไม่เขียน Excel
 * ------------------------------------------------------------------
 */

import { ReconcileRecord } from "../shared/record";

import { AmountComparator } from "./ltx-amount-compare";

import type { ReconcileReportConfig } from "./ltx-config";

import type { ExpectedCase } from "./ltx-expected-case-builder";

import { FieldRuleValidatorSet } from "./ltx-field-validator";

export type LtxMatchStrategy =
  | "EXACT"
  | "EXACT_BEST_MATCH"
  | "FALLBACK"
  | "FALLBACK_BEST_MATCH"
  | "NONE"
  | "EXACT_ALREADY_USED"
  | "FALLBACK_ALREADY_USED"
  | "AMBIGUOUS_EXACT"
  | "AMBIGUOUS_FALLBACK";

export interface LtxMatchResult {
  matchedRecord: ReconcileRecord | undefined;
  strategy: LtxMatchStrategy;
  informationalRemark: string;
  reviewRemark: string;
  reviewFieldHeaders: string[];
  failureRemark: string;
}

interface LtxCandidateScore {
  record: ReconcileRecord;
  requiredKeyMatchFields: string[];
  secondaryMatchFields: string[];
  supportingMatchFields: string[];
  supportingMismatchCount: number;
}

interface LtxBestCandidateResult {
  selectedScore: LtxCandidateScore | undefined;
  highestScores: LtxCandidateScore[];
}

export class LtxMatcher {
  constructor(
    private readonly amountComparator: AmountComparator = new AmountComparator(),
    private readonly fieldValidatorSet: FieldRuleValidatorSet = new FieldRuleValidatorSet(),
  ) {}

  private normalize(value: string): string {
    return value.trim().toUpperCase();
  }

  /**
   * ค่าว่างสองฝั่งไม่ถือว่าตรงกัน
   * เพื่อป้องกันการเลือก Candidate จากข้อมูลที่ไม่มีหลักฐาน
   */
  private isSameNonEmptyText(expected: string, actual: string): boolean {
    const normalizedExpected = this.normalize(expected);
    const normalizedActual = this.normalize(actual);

    return (
      normalizedExpected !== "" &&
      normalizedActual !== "" &&
      normalizedExpected === normalizedActual
    );
  }

  private isSameSlot(
    record: ReconcileRecord,
    config: ReconcileReportConfig,
    suffix: string,
  ): boolean {
    const reference = this.normalize(
      record.get(config.referenceNumberReportField),
    );
    const normalizedSuffix = this.normalize(suffix);

    return (
      reference !== "" &&
      normalizedSuffix !== "" &&
      reference.endsWith(normalizedSuffix)
    );
  }

  /**
   * DR ใช้ From Transfer Amount หรือ From Debit Amount
   * FE ใช้ SUM Fee Amount ของ Expected Case
   */
  private isAmountMatch(
    record: ReconcileRecord,
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): boolean {
    const actualAmount = record.get(config.transactionAmountReportField);

    if (this.normalize(suffix) === this.normalize(config.feSuffixLabel)) {
      return this.amountComparator.compare(
        expectedCase.expectedFeAmount,
        actualAmount,
      ).isMatch;
    }

    const primaryMatch = this.amountComparator.compare(
      expectedCase.primaryRecord.get(config.drAmountTestDataField),
      actualAmount,
    ).isMatch;

    const fallbackMatch = this.amountComparator.compare(
      expectedCase.primaryRecord.get(config.drAmountFallbackTestDataField),
      actualAmount,
    ).isMatch;

    return primaryMatch || fallbackMatch;
  }

  private findExactCandidates(
    reportRecords: ReconcileRecord[],
    config: ReconcileReportConfig,
    expectedReference: string,
  ): ReconcileRecord[] {
    const target = this.normalize(expectedReference);

    if (target === "") {
      return [];
    }

    return reportRecords.filter(
      (record) =>
        this.normalize(record.get(config.referenceNumberReportField)) ===
        target,
    );
  }

  private findFallbackCandidates(
    reportRecords: ReconcileRecord[],
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): ReconcileRecord[] {
    const expectedAccount = expectedCase.primaryRecord.get(
      config.groupKeyFields.testDataAccountField,
    );
    const expectedCurrency = expectedCase.primaryRecord.get(
      config.groupKeyFields.testDataCurrencyField,
    );

    /**
     * Account และ Currency เป็นข้อมูลขั้นต่ำของ Fallback
     * จึงไม่ใช้ Amount เพียงอย่างเดียวในการเดาแถว
     */
    if (
      this.normalize(expectedAccount) === "" ||
      this.normalize(expectedCurrency) === ""
    ) {
      return [];
    }

    return reportRecords.filter(
      (record) =>
        this.isSameNonEmptyText(
          expectedAccount,
          record.get(config.groupKeyFields.reportAccountField),
        ) &&
        this.isSameNonEmptyText(
          expectedCurrency,
          record.get(config.groupKeyFields.reportCurrencyField),
        ) &&
        this.isSameSlot(record, config, suffix) &&
        this.isAmountMatch(record, expectedCase, config, suffix),
    );
  }

  private getMissingFallbackInputFields(
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): string[] {
    const missingFields = [
      config.groupKeyFields.testDataAccountField,
      config.groupKeyFields.testDataCurrencyField,
    ].filter(
      (field) => this.normalize(expectedCase.primaryRecord.get(field)) === "",
    );

    const hasFallbackAmount = this.hasExpectedAmountForSlot(
      expectedCase,
      config,
      suffix,
    );

    if (!hasFallbackAmount) {
      missingFields.push(
        this.normalize(suffix) === this.normalize(config.feSuffixLabel)
          ? "SUM(Fee Amount)"
          : `${config.drAmountTestDataField} / ` +
              config.drAmountFallbackTestDataField,
      );
    }

    return missingFields;
  }

  /**
   * ใช้ผลการตัดสิน Slot จาก ExpectedCaseBuilder เป็นแหล่งเดียว
   * ไม่ใช้ Amount Tolerance เป็นเกณฑ์ว่ามีหรือไม่มี Transaction
   */
  private hasExpectedAmountForSlot(
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): boolean {
    return this.normalize(suffix) === this.normalize(config.feSuffixLabel)
      ? expectedCase.hasExpectedFe
      : expectedCase.hasExpectedDr;
  }

  /**
   * Key Score ใช้ลำดับความสำคัญสูงกว่า Supporting Field
   * จึงไม่มีกรณีที่ Field ทั่วไปจำนวนมากกลบ Account หรือ Amount ที่ผิด
   */
  private buildCandidateScore(
    reportCode: string,
    record: ReconcileRecord,
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    suffix: string,
  ): LtxCandidateScore {
    const requiredKeyMatchFields: string[] = [];
    const secondaryMatchFields: string[] = [];

    if (
      this.isSameNonEmptyText(
        expectedReference,
        record.get(config.referenceNumberReportField),
      )
    ) {
      requiredKeyMatchFields.push(config.referenceNumberReportField);
    }

    if (
      this.isSameNonEmptyText(
        expectedCase.primaryRecord.get(
          config.groupKeyFields.testDataAccountField,
        ),
        record.get(config.groupKeyFields.reportAccountField),
      )
    ) {
      requiredKeyMatchFields.push(config.groupKeyFields.reportAccountField);
    }

    if (
      this.isSameNonEmptyText(
        expectedCase.primaryRecord.get(
          config.groupKeyFields.testDataCurrencyField,
        ),
        record.get(config.groupKeyFields.reportCurrencyField),
      )
    ) {
      secondaryMatchFields.push(config.groupKeyFields.reportCurrencyField);
    }

    if (this.isAmountMatch(record, expectedCase, config, suffix)) {
      requiredKeyMatchFields.push(config.transactionAmountReportField);
    }

    const keyFieldHeaders = new Set(
      [
        config.referenceNumberReportField,
        config.groupKeyFields.reportAccountField,
        config.groupKeyFields.reportCurrencyField,
        config.transactionAmountReportField,
      ].map((header) => this.normalize(header)),
    );

    const supportingResults = this.fieldValidatorSet
      .validateAll(
        reportCode,
        config.fieldRules,
        expectedCase.primaryRecord,
        record,
        suffix,
      )
      .filter(
        (result) => !keyFieldHeaders.has(this.normalize(result.fieldHeader)),
      );

    return {
      record,
      requiredKeyMatchFields,
      secondaryMatchFields,
      supportingMatchFields: supportingResults
        .filter((result) => result.status === "PASS")
        .map((result) => result.fieldHeader),
      supportingMismatchCount: supportingResults.filter(
        (result) => result.status !== "PASS",
      ).length,
    };
  }

  private isSameScore(
    left: LtxCandidateScore,
    right: LtxCandidateScore,
  ): boolean {
    return (
      left.requiredKeyMatchFields.length ===
        right.requiredKeyMatchFields.length &&
      left.secondaryMatchFields.length === right.secondaryMatchFields.length &&
      left.supportingMatchFields.length ===
        right.supportingMatchFields.length &&
      left.supportingMismatchCount === right.supportingMismatchCount
    );
  }

  private compareCandidateScores(
    left: LtxCandidateScore,
    right: LtxCandidateScore,
  ): number {
    const requiredKeyDifference =
      right.requiredKeyMatchFields.length - left.requiredKeyMatchFields.length;

    if (requiredKeyDifference !== 0) {
      return requiredKeyDifference;
    }

    const secondaryDifference =
      right.secondaryMatchFields.length - left.secondaryMatchFields.length;

    if (secondaryDifference !== 0) {
      return secondaryDifference;
    }

    const supportingMatchDifference =
      right.supportingMatchFields.length - left.supportingMatchFields.length;

    if (supportingMatchDifference !== 0) {
      return supportingMatchDifference;
    }

    return left.supportingMismatchCount - right.supportingMismatchCount;
  }

  private selectBestCandidate(
    reportCode: string,
    candidates: ReconcileRecord[],
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    suffix: string,
  ): LtxBestCandidateResult {
    const scores = candidates.map((record) =>
      this.buildCandidateScore(
        reportCode,
        record,
        expectedCase,
        config,
        expectedReference,
        suffix,
      ),
    );

    const sortedScores = [...scores].sort((left, right) =>
      this.compareCandidateScores(left, right),
    );

    const bestScore = sortedScores[0];

    if (!bestScore) {
      return {
        selectedScore: undefined,
        highestScores: [],
      };
    }

    const highestScores = sortedScores.filter((score) =>
      this.isSameScore(score, bestScore),
    );

    return {
      selectedScore: highestScores.length === 1 ? bestScore : undefined,
      highestScores,
    };
  }

  private buildFallbackReason(
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
  ): string {
    const expectedId = expectedCase.primaryRecord
      .get(config.testDataIdField)
      .trim();

    return expectedId === ""
      ? `Test Data ไม่มี ${config.testDataIdField}`
      : `ไม่พบ Exact Reference = "${expectedReference.trim()}"`;
  }

  private buildFallbackInformationalRemark(
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    matchedRecord: ReconcileRecord,
  ): string {
    const matchedReference = matchedRecord
      .get(config.referenceNumberReportField)
      .trim();

    return (
      `${this.buildFallbackReason(expectedCase, config, expectedReference)}\n` +
      "Mapping ด้วย LTX Fallback: Account + Currency + DR/FE + Amount " +
      `พบคู่กับ ${config.referenceNumberReportField} = "${matchedReference}"`
    );
  }

  private buildSelectionReviewRemark(
    sourceLabel: string,
    candidates: ReconcileRecord[],
    selectedScore: LtxCandidateScore,
  ): string {
    const candidateRows = candidates
      .map((record) => record.rowNumber)
      .join(", ");

    return (
      `${sourceLabel} พบ Candidate ${candidates.length} แถว ` +
      `(Report row: ${candidateRows})\n` +
      `ระบบเลือก Report row ${selectedScore.record.rowNumber} ` +
      "เนื่องจากมีข้อมูลตรงกับ Test Data มากที่สุด " +
      `(Required Key ${selectedScore.requiredKeyMatchFields.length} รายการ, ` +
      `Secondary ${selectedScore.secondaryMatchFields.length} รายการ, ` +
      `Supporting Field ${selectedScore.supportingMatchFields.length} รายการ)`
    );
  }

  private buildAmbiguousRemark(
    sourceLabel: string,
    candidates: ReconcileRecord[],
    highestScores: LtxCandidateScore[],
  ): string {
    const candidateRows = candidates
      .map((record) => record.rowNumber)
      .join(", ");
    const tiedRows = highestScores
      .map((score) => score.record.rowNumber)
      .join(", ");

    return (
      `${sourceLabel} พบ Candidate ${candidates.length} แถว ` +
      `(Report row: ${candidateRows}) และแถวที่มีข้อมูลตรงมากที่สุด ` +
      `ยังเสมอกัน (Report row: ${tiedRows}) ` +
      "จึงไม่สามารถระบุคู่ที่แน่นอนได้"
    );
  }

  private emptyResult(strategy: LtxMatchStrategy): LtxMatchResult {
    return {
      matchedRecord: undefined,
      strategy,
      informationalRemark: "",
      reviewRemark: "",
      reviewFieldHeaders: [],
      failureRemark: "",
    };
  }

  private getAvailableRecords(
    records: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
  ): ReconcileRecord[] {
    return records.filter(
      (record) => !usedReportRowNumbers.has(record.rowNumber),
    );
  }

  private buildAlreadyUsedExactRemark(
    expectedReference: string,
    exactCandidates: ReconcileRecord[],
  ): string {
    const usedRows = exactCandidates
      .map((record) => record.rowNumber)
      .join(", ");

    return (
      `พบ Exact Reference = "${expectedReference.trim()}" ` +
      `ที่ Report row ${usedRows} แต่ทุกแถวถูกจับคู่กับ Test Case อื่นแล้ว ` +
      "ระบบจึงหยุดและไม่ใช้ Fallback ไปเลือก Transaction อื่น"
    );
  }

  private buildAlreadyUsedFallbackRemark(
    candidates: ReconcileRecord[],
  ): string {
    const usedRows = candidates
      .map((record) => record.rowNumber)
      .join(", ");

    return (
      "พบ LTX Fallback Candidate ที่ Report row " +
      `${usedRows} แต่ทุกแถวถูกจับคู่กับ Test Case อื่นแล้ว ` +
      "ระบบจึงไม่ใช้ Report row ซ้ำ"
    );
  }

  private buildUsedFallbackSelectionRemark(
    allCandidates: ReconcileRecord[],
    availableCandidates: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
  ): string {
    const usedRows = allCandidates
      .filter((record) => usedReportRowNumbers.has(record.rowNumber))
      .map((record) => record.rowNumber);

    if (usedRows.length === 0) {
      return "";
    }

    const availableRows = availableCandidates
      .map((record) => record.rowNumber)
      .join(", ");

    return (
      `LTX Fallback Matching พบ Candidate ที่ Report row ${usedRows.join(", ")} ` +
      "ถูกจับคู่กับ Test Case อื่นแล้ว " +
      `จึงพิจารณาเฉพาะ Report row ที่ยังใช้ได้: ${availableRows}`
    );
  }

  private selectPresenceRecord(
    candidates: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
  ): ReconcileRecord {
    const availableCandidate = candidates.find(
      (record) => !usedReportRowNumbers.has(record.rowNumber),
    );

    return availableCandidate ?? candidates[0];
  }

  private buildPresenceRemark(
    sourceLabel: string,
    candidates: ReconcileRecord[],
    selectedRecord: ReconcileRecord,
    usedReportRowNumbers: ReadonlySet<number>,
  ): string {
    if (candidates.length === 1) {
      return "";
    }

    const candidateRows = candidates
      .map((record) => record.rowNumber)
      .join(", ");
    const selectedRowWasUsed = usedReportRowNumbers.has(
      selectedRecord.rowNumber,
    );
    const selectionReason = selectedRowWasUsed
      ? "ทุก Candidate ถูก Test Case อื่นใช้แล้ว แต่ยังถือว่าพบรายการใน Report"
      : "เลือกแถวที่ยังไม่ถูก Test Case อื่นใช้มาแสดงผล";

    return (
      `${sourceLabel} สำหรับ Presence Check พบ ${candidates.length} แถว ` +
      `(Report row: ${candidateRows}); ${selectionReason}`
    );
  }

  /**
   * ตรวจว่ารายการมีอยู่ใน Report หรือไม่สำหรับ MUST_NOT_EXIST
   * โดยค้นจาก Report ทั้งหมด ไม่ตัดแถวที่ถูกใช้แล้วออก
   */
  findPresence(
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    suffix: string,
  ): LtxMatchResult {
    const exactCandidates = this.findExactCandidates(
      reportRecords,
      config,
      expectedReference,
    );

    if (exactCandidates.length > 0) {
      const matchedRecord = this.selectPresenceRecord(
        exactCandidates,
        usedReportRowNumbers,
      );

      return {
        ...this.emptyResult(
          exactCandidates.length === 1 ? "EXACT" : "EXACT_BEST_MATCH",
        ),
        matchedRecord,
        informationalRemark: this.buildPresenceRemark(
          "Exact Reference Matching",
          exactCandidates,
          matchedRecord,
          usedReportRowNumbers,
        ),
      };
    }

    const missingFallbackInputFields = this.getMissingFallbackInputFields(
      expectedCase,
      config,
      suffix,
    );

    if (missingFallbackInputFields.length > 0) {
      const reviewRemark =
        "LTX Presence Fallback ทำไม่ได้ เพราะ Test Data ไม่มีข้อมูล: " +
        missingFallbackInputFields.join(", ");

      return {
        ...this.emptyResult("NONE"),
        reviewRemark,
        reviewFieldHeaders: [config.referenceNumberReportField],
      };
    }

    const fallbackCandidates = this.findFallbackCandidates(
      reportRecords,
      expectedCase,
      config,
      suffix,
    );

    if (fallbackCandidates.length === 0) {
      return this.emptyResult("NONE");
    }

    const matchedRecord = this.selectPresenceRecord(
      fallbackCandidates,
      usedReportRowNumbers,
    );

    return {
      ...this.emptyResult(
        fallbackCandidates.length === 1 ? "FALLBACK" : "FALLBACK_BEST_MATCH",
      ),
      matchedRecord,
      informationalRemark: [
        this.buildFallbackInformationalRemark(
          expectedCase,
          config,
          expectedReference,
          matchedRecord,
        ),
        this.buildPresenceRemark(
          "LTX Fallback Matching",
          fallbackCandidates,
          matchedRecord,
          usedReportRowNumbers,
        ),
      ]
        .filter((remark) => remark !== "")
        .join("\n"),
    };
  }

  /**
   * จับคู่ MUST_EXIST โดยไม่อนุญาตให้ใช้ Report row ซ้ำ
   */
  findMatch(
    reportCode: string,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    suffix: string,
  ): LtxMatchResult {
    const allExactCandidates = this.findExactCandidates(
      reportRecords,
      config,
      expectedReference,
    );
    const availableExactCandidates = this.getAvailableRecords(
      allExactCandidates,
      usedReportRowNumbers,
    );

    if (
      allExactCandidates.length > 0 &&
      availableExactCandidates.length === 0
    ) {
      return {
        ...this.emptyResult("EXACT_ALREADY_USED"),
        failureRemark: this.buildAlreadyUsedExactRemark(
          expectedReference,
          allExactCandidates,
        ),
      };
    }

    if (availableExactCandidates.length === 1) {
      const matchedRecord = availableExactCandidates[0];
      const hasOtherUsedExactRows = allExactCandidates.length > 1;

      return {
        ...this.emptyResult(
          hasOtherUsedExactRows ? "EXACT_BEST_MATCH" : "EXACT",
        ),
        matchedRecord,
        reviewRemark: hasOtherUsedExactRows
          ? "Exact Reference พบมากกว่าหนึ่งแถว และระบบเลือกแถวที่ยังไม่ถูกใช้"
          : "",
        reviewFieldHeaders: hasOtherUsedExactRows
          ? [config.referenceNumberReportField]
          : [],
      };
    }

    if (availableExactCandidates.length > 1) {
      const selection = this.selectBestCandidate(
        reportCode,
        availableExactCandidates,
        expectedCase,
        config,
        expectedReference,
        suffix,
      );

      if (selection.selectedScore) {
        return {
          ...this.emptyResult("EXACT_BEST_MATCH"),
          matchedRecord: selection.selectedScore.record,
          reviewRemark: this.buildSelectionReviewRemark(
            "Exact Reference Matching",
            availableExactCandidates,
            selection.selectedScore,
          ),
          reviewFieldHeaders: [config.referenceNumberReportField],
        };
      }

      return {
        ...this.emptyResult("AMBIGUOUS_EXACT"),
        failureRemark: this.buildAmbiguousRemark(
          "Exact Reference Matching",
          availableExactCandidates,
          selection.highestScores,
        ),
      };
    }

    const missingFallbackInputFields = this.getMissingFallbackInputFields(
      expectedCase,
      config,
      suffix,
    );

    if (missingFallbackInputFields.length > 0) {
      return {
        ...this.emptyResult("NONE"),
        informationalRemark:
          "LTX Fallback Matching ทำไม่ได้ เพราะ Test Data ไม่มีข้อมูล: " +
          missingFallbackInputFields.join(", "),
      };
    }

    const allFallbackCandidates = this.findFallbackCandidates(
      reportRecords,
      expectedCase,
      config,
      suffix,
    );

    const fallbackCandidates = this.getAvailableRecords(
      allFallbackCandidates,
      usedReportRowNumbers,
    );

    if (
      allFallbackCandidates.length > 0 &&
      fallbackCandidates.length === 0
    ) {
      return {
        ...this.emptyResult("FALLBACK_ALREADY_USED"),
        informationalRemark: this.buildFallbackReason(
          expectedCase,
          config,
          expectedReference,
        ),
        failureRemark: this.buildAlreadyUsedFallbackRemark(
          allFallbackCandidates,
        ),
      };
    }

    if (fallbackCandidates.length === 0) {
      return this.emptyResult("NONE");
    }

    const usedCandidateReviewRemark =
      this.buildUsedFallbackSelectionRemark(
        allFallbackCandidates,
        fallbackCandidates,
        usedReportRowNumbers,
      );

    if (fallbackCandidates.length === 1) {
      const matchedRecord = fallbackCandidates[0];
      const hasUsedCandidates = usedCandidateReviewRemark !== "";

      return {
        ...this.emptyResult(
          hasUsedCandidates ? "FALLBACK_BEST_MATCH" : "FALLBACK",
        ),
        matchedRecord,
        informationalRemark: this.buildFallbackInformationalRemark(
          expectedCase,
          config,
          expectedReference,
          matchedRecord,
        ),
        reviewRemark: usedCandidateReviewRemark,
        reviewFieldHeaders: hasUsedCandidates
          ? [config.referenceNumberReportField]
          : [],
      };
    }

    const selection = this.selectBestCandidate(
      reportCode,
      fallbackCandidates,
      expectedCase,
      config,
      expectedReference,
      suffix,
    );

    if (selection.selectedScore) {
      return {
        ...this.emptyResult("FALLBACK_BEST_MATCH"),
        matchedRecord: selection.selectedScore.record,
        informationalRemark: this.buildFallbackInformationalRemark(
          expectedCase,
          config,
          expectedReference,
          selection.selectedScore.record,
        ),
        reviewRemark: [
          usedCandidateReviewRemark,
          this.buildSelectionReviewRemark(
            "LTX Fallback Matching",
            fallbackCandidates,
            selection.selectedScore,
          ),
        ]
          .filter((remark) => remark !== "")
          .join("\n"),
        reviewFieldHeaders: [config.referenceNumberReportField],
      };
    }

    return {
      ...this.emptyResult("AMBIGUOUS_FALLBACK"),
      informationalRemark: this.buildFallbackReason(
        expectedCase,
        config,
        expectedReference,
      ),
      failureRemark: [
        usedCandidateReviewRemark,
        this.buildAmbiguousRemark(
          "LTX Fallback Matching",
          fallbackCandidates,
          selection.highestScores,
        ),
      ]
        .filter((remark) => remark !== "")
        .join("\n"),
    };
  }
}