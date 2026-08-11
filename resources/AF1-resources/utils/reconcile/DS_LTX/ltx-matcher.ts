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
 * ไฟล์นี้ไม่ตัดสิน PASS/FAIL ของ Test Case และไม่เขียน Excel
 * ------------------------------------------------------------------
 */

import {
  ReconcileRecord,
} from "../shared/record";

import {
  AmountComparator,
} from "./ltx-amount-compare";

import type {
  ReconcileReportConfig,
} from "./ltx-config";

import type {
  ExpectedCase,
} from "./ltx-expected-case-builder";

import {
  FieldRuleValidatorSet,
} from "./ltx-field-validator";

export type LtxMatchStrategy =
  | "EXACT"
  | "EXACT_BEST_MATCH"
  | "FALLBACK"
  | "FALLBACK_BEST_MATCH"
  | "NONE"
  | "AMBIGUOUS_EXACT"
  | "AMBIGUOUS_FALLBACK";

export interface LtxMatchResult {
  matchedRecord: ReconcileRecord | undefined;
  strategy: LtxMatchStrategy;
  candidateCount: number;
  candidateRows: number[];
  informationalRemark: string;
  reviewRemark: string;
  reviewFieldHeaders: string[];
  failureRemark: string;
}

interface LtxCandidateScore {
  record: ReconcileRecord;
  keyMatchFields: string[];
  supportingMatchFields: string[];
  supportingMismatchCount: number;
}

interface LtxBestCandidateResult {
  selectedScore: LtxCandidateScore | undefined;
  highestScores: LtxCandidateScore[];
}

export class LtxMatcher {
  constructor(
    private readonly amountComparator:
      AmountComparator =
      new AmountComparator(),
    private readonly fieldValidatorSet:
      FieldRuleValidatorSet =
      new FieldRuleValidatorSet(),
  ) {}

  private normalize(
    value: string,
  ): string {
    return value
      .trim()
      .toUpperCase();
  }

  /**
   * ค่าว่างสองฝั่งไม่ถือว่าตรงกัน
   * เพื่อป้องกันการเลือก Candidate จากข้อมูลที่ไม่มีหลักฐาน
   */
  private isSameNonEmptyText(
    expected: string,
    actual: string,
  ): boolean {
    const normalizedExpected =
      this.normalize(expected);
    const normalizedActual =
      this.normalize(actual);

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
      record.get(
        config.referenceNumberReportField,
      ),
    );
    const normalizedSuffix =
      this.normalize(suffix);

    return (
      reference !== "" &&
      normalizedSuffix !== "" &&
      reference.endsWith(
        normalizedSuffix,
      )
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
    const actualAmount = record.get(
      config.transactionAmountReportField,
    );

    if (
      this.normalize(suffix) ===
      this.normalize(config.feSuffixLabel)
    ) {
      return this.amountComparator.compare(
        expectedCase.expectedFeAmount,
        actualAmount,
      ).isMatch;
    }

    const primaryMatch =
      this.amountComparator.compare(
        expectedCase.primaryRecord.get(
          config.drAmountTestDataField,
        ),
        actualAmount,
      ).isMatch;

    const fallbackMatch =
      this.amountComparator.compare(
        expectedCase.primaryRecord.get(
          config.drAmountFallbackTestDataField,
        ),
        actualAmount,
      ).isMatch;

    return primaryMatch || fallbackMatch;
  }

  private findExactCandidates(
    reportRecords: ReconcileRecord[],
    config: ReconcileReportConfig,
    expectedReference: string,
  ): ReconcileRecord[] {
    const target =
      this.normalize(expectedReference);

    if (
      target === ""
    ) {
      return [];
    }

    return reportRecords.filter(
      (record) =>
        this.normalize(
          record.get(
            config.referenceNumberReportField,
          ),
        ) === target,
    );
  }

  private findFallbackCandidates(
    reportRecords: ReconcileRecord[],
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): ReconcileRecord[] {
    const expectedAccount =
      expectedCase.primaryRecord.get(
        config.groupKeyFields
          .testDataAccountField,
      );
    const expectedCurrency =
      expectedCase.primaryRecord.get(
        config.groupKeyFields
          .testDataCurrencyField,
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
          record.get(
            config.groupKeyFields
              .reportAccountField,
          ),
        ) &&
        this.isSameNonEmptyText(
          expectedCurrency,
          record.get(
            config.groupKeyFields
              .reportCurrencyField,
          ),
        ) &&
        this.isSameSlot(
          record,
          config,
          suffix,
        ) &&
        this.isAmountMatch(
          record,
          expectedCase,
          config,
          suffix,
        ),
    );
  }

  private getMissingFallbackInputFields(
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
  ): string[] {
    return [
      config.groupKeyFields
        .testDataAccountField,
      config.groupKeyFields
        .testDataCurrencyField,
    ].filter(
      (field) =>
        this.normalize(
          expectedCase.primaryRecord.get(
            field,
          ),
        ) === "",
    );
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
    const keyMatchFields:
      string[] = [];

    if (
      this.isSameNonEmptyText(
        expectedReference,
        record.get(
          config.referenceNumberReportField,
        ),
      )
    ) {
      keyMatchFields.push(
        config.referenceNumberReportField,
      );
    }

    if (
      this.isSameNonEmptyText(
        expectedCase.primaryRecord.get(
          config.groupKeyFields
            .testDataAccountField,
        ),
        record.get(
          config.groupKeyFields
            .reportAccountField,
        ),
      )
    ) {
      keyMatchFields.push(
        config.groupKeyFields
          .reportAccountField,
      );
    }

    if (
      this.isSameNonEmptyText(
        expectedCase.primaryRecord.get(
          config.groupKeyFields
            .testDataCurrencyField,
        ),
        record.get(
          config.groupKeyFields
            .reportCurrencyField,
        ),
      )
    ) {
      keyMatchFields.push(
        config.groupKeyFields
          .reportCurrencyField,
      );
    }

    if (
      this.isAmountMatch(
        record,
        expectedCase,
        config,
        suffix,
      )
    ) {
      keyMatchFields.push(
        config.transactionAmountReportField,
      );
    }

    const keyFieldHeaders = new Set(
      [
        config.referenceNumberReportField,
        config.groupKeyFields
          .reportAccountField,
        config.groupKeyFields
          .reportCurrencyField,
        config.transactionAmountReportField,
      ].map(
        (header) =>
          this.normalize(header),
      ),
    );

    const supportingResults =
      this.fieldValidatorSet.validateAll(
        reportCode,
        config.fieldRules,
        expectedCase.primaryRecord,
        record,
        suffix,
      ).filter(
        (result) =>
          !keyFieldHeaders.has(
            this.normalize(
              result.fieldHeader,
            ),
          ),
      );

    return {
      record,
      keyMatchFields,
      supportingMatchFields:
        supportingResults
          .filter(
            (result) =>
              result.status === "PASS",
          )
          .map(
            (result) =>
              result.fieldHeader,
          ),
      supportingMismatchCount:
        supportingResults.filter(
          (result) =>
            result.status !== "PASS",
        ).length,
    };
  }

  private isSameScore(
    left: LtxCandidateScore,
    right: LtxCandidateScore,
  ): boolean {
    return (
      left.keyMatchFields.length ===
        right.keyMatchFields.length &&
      left.supportingMatchFields.length ===
        right.supportingMatchFields.length &&
      left.supportingMismatchCount ===
        right.supportingMismatchCount
    );
  }

  private selectBestCandidate(
    reportCode: string,
    candidates: ReconcileRecord[],
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    suffix: string,
  ): LtxBestCandidateResult {
    const scores = candidates.map(
      (record) =>
        this.buildCandidateScore(
          reportCode,
          record,
          expectedCase,
          config,
          expectedReference,
          suffix,
        ),
    );

    const sortedScores = [
      ...scores,
    ].sort(
      (left, right) =>
        right.keyMatchFields.length -
          left.keyMatchFields.length ||
        right.supportingMatchFields.length -
          left.supportingMatchFields.length ||
        left.supportingMismatchCount -
          right.supportingMismatchCount,
    );

    const bestScore =
      sortedScores[0];

    if (
      !bestScore
    ) {
      return {
        selectedScore: undefined,
        highestScores: [],
      };
    }

    const highestScores =
      sortedScores.filter(
        (score) =>
          this.isSameScore(
            score,
            bestScore,
          ),
      );

    return {
      selectedScore:
        highestScores.length === 1
          ? bestScore
          : undefined,
      highestScores,
    };
  }

  private buildFallbackReason(
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
  ): string {
    const expectedId =
      expectedCase.primaryRecord
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
    const matchedReference =
      matchedRecord.get(
        config.referenceNumberReportField,
      ).trim();

    return (
      `${this.buildFallbackReason(
        expectedCase,
        config,
        expectedReference,
      )}\n` +
      "Mapping ด้วย LTX Fallback: Account + Currency + DR/FE + Amount " +
      `พบคู่กับ ${config.referenceNumberReportField} = "${matchedReference}"`
    );
  }

  private buildSelectionReviewRemark(
    sourceLabel: string,
    candidates: ReconcileRecord[],
    selectedScore: LtxCandidateScore,
  ): string {
    const candidateRows =
      candidates
        .map(
          (record) =>
            record.rowNumber,
        )
        .join(", ");

    return (
      `${sourceLabel} พบ Candidate ${candidates.length} แถว ` +
      `(Report row: ${candidateRows})\n` +
      `ระบบเลือก Report row ${selectedScore.record.rowNumber} ` +
      "เนื่องจากมีข้อมูลตรงกับ Test Data มากที่สุด " +
      `(Key ${selectedScore.keyMatchFields.length} รายการ, ` +
      `Supporting Field ${selectedScore.supportingMatchFields.length} รายการ)`
    );
  }

  private buildAmbiguousRemark(
    sourceLabel: string,
    candidates: ReconcileRecord[],
    highestScores: LtxCandidateScore[],
  ): string {
    const candidateRows =
      candidates
        .map(
          (record) =>
            record.rowNumber,
        )
        .join(", ");
    const tiedRows =
      highestScores
        .map(
          (score) =>
            score.record.rowNumber,
        )
        .join(", ");

    return (
      `${sourceLabel} พบ Candidate ${candidates.length} แถว ` +
      `(Report row: ${candidateRows}) และแถวที่มีข้อมูลตรงมากที่สุด ` +
      `ยังเสมอกัน (Report row: ${tiedRows}) ` +
      "จึงไม่สามารถระบุคู่ที่แน่นอนได้"
    );
  }

  private emptyResult(
    strategy: LtxMatchStrategy,
  ): LtxMatchResult {
    return {
      matchedRecord: undefined,
      strategy,
      candidateCount: 0,
      candidateRows: [],
      informationalRemark: "",
      reviewRemark: "",
      reviewFieldHeaders: [],
      failureRemark: "",
    };
  }

  findMatch(
    reportCode: string,
    reportRecords: ReconcileRecord[],
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    expectedReference: string,
    suffix: string,
  ): LtxMatchResult {
    const exactCandidates =
      this.findExactCandidates(
        reportRecords,
        config,
        expectedReference,
      );

    if (
      exactCandidates.length === 1
    ) {
      return {
        ...this.emptyResult("EXACT"),
        matchedRecord:
          exactCandidates[0],
        candidateCount: 1,
        candidateRows: [
          exactCandidates[0].rowNumber,
        ],
      };
    }

    if (
      exactCandidates.length > 1
    ) {
      const selection =
        this.selectBestCandidate(
          reportCode,
          exactCandidates,
          expectedCase,
          config,
          expectedReference,
          suffix,
        );

      if (
        selection.selectedScore
      ) {
        return {
          ...this.emptyResult(
            "EXACT_BEST_MATCH",
          ),
          matchedRecord:
            selection.selectedScore.record,
          candidateCount:
            exactCandidates.length,
          candidateRows:
            exactCandidates.map(
              (record) =>
                record.rowNumber,
            ),
          reviewRemark:
            this.buildSelectionReviewRemark(
              "Exact Reference Matching",
              exactCandidates,
              selection.selectedScore,
            ),
          reviewFieldHeaders: [
            config.referenceNumberReportField,
          ],
        };
      }

      return {
        ...this.emptyResult(
          "AMBIGUOUS_EXACT",
        ),
        candidateCount:
          exactCandidates.length,
        candidateRows:
          exactCandidates.map(
            (record) =>
              record.rowNumber,
          ),
        failureRemark:
          this.buildAmbiguousRemark(
            "Exact Reference Matching",
            exactCandidates,
            selection.highestScores,
          ),
      };
    }

    const missingFallbackInputFields =
      this.getMissingFallbackInputFields(
        expectedCase,
        config,
      );

    if (
      missingFallbackInputFields.length > 0
    ) {
      return {
        ...this.emptyResult("NONE"),
        informationalRemark:
          "LTX Fallback Matching ทำไม่ได้ เพราะ Test Data ไม่มีข้อมูล: " +
          missingFallbackInputFields.join(", "),
      };
    }

    const fallbackCandidates =
      this.findFallbackCandidates(
        reportRecords,
        expectedCase,
        config,
        suffix,
      );

    if (
      fallbackCandidates.length === 0
    ) {
      return this.emptyResult("NONE");
    }

    if (
      fallbackCandidates.length === 1
    ) {
      const matchedRecord =
        fallbackCandidates[0];

      return {
        ...this.emptyResult("FALLBACK"),
        matchedRecord,
        candidateCount: 1,
        candidateRows: [
          matchedRecord.rowNumber,
        ],
        informationalRemark:
          this.buildFallbackInformationalRemark(
            expectedCase,
            config,
            expectedReference,
            matchedRecord,
          ),
      };
    }

    const selection =
      this.selectBestCandidate(
        reportCode,
        fallbackCandidates,
        expectedCase,
        config,
        expectedReference,
        suffix,
      );

    if (
      selection.selectedScore
    ) {
      return {
        ...this.emptyResult(
          "FALLBACK_BEST_MATCH",
        ),
        matchedRecord:
          selection.selectedScore.record,
        candidateCount:
          fallbackCandidates.length,
        candidateRows:
          fallbackCandidates.map(
            (record) =>
              record.rowNumber,
          ),
        informationalRemark:
          this.buildFallbackInformationalRemark(
            expectedCase,
            config,
            expectedReference,
            selection.selectedScore.record,
          ),
        reviewRemark:
          this.buildSelectionReviewRemark(
            "LTX Fallback Matching",
            fallbackCandidates,
            selection.selectedScore,
          ),
        reviewFieldHeaders: [
          config.referenceNumberReportField,
        ],
      };
    }

    return {
      ...this.emptyResult(
        "AMBIGUOUS_FALLBACK",
      ),
      candidateCount:
        fallbackCandidates.length,
      candidateRows:
        fallbackCandidates.map(
          (record) =>
            record.rowNumber,
        ),
      informationalRemark:
        this.buildFallbackReason(
          expectedCase,
          config,
          expectedReference,
        ),
      failureRemark:
        this.buildAmbiguousRemark(
          "LTX Fallback Matching",
          fallbackCandidates,
          selection.highestScores,
        ),
    };
  }
}