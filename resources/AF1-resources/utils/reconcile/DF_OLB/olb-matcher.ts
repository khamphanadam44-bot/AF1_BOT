/**
 * olb-matcher.ts
 * ------------------------------------------------------------------
 * รับผิดชอบเฉพาะการหา AF1 Row ที่เหมาะสมกับ Test Data 1 แถว
 * ไม่ตัดสิน PASS/FAIL และไม่เขียน Excel
 * ------------------------------------------------------------------
 */

import { ReconcileRecord } from "../shared/record";
import {
  isDateMatchWithArrangementFallback,
  parseAmount,
  parseDate,
} from "../shared/reconcile-parse.util";
import {
  OLB_AMOUNT_TOLERANCE,
  OLB_REPORT_FIELDS,
  OLB_TEST_DATA_FIELDS,
} from "./olb-config";
import {
  normalizeOlbId,
  normalizeOlbText,
} from "./olb-normalize.util";

type CandidateMatchField =
  | "Date"
  | "Amount"
  | "CIF No."
  | "CIF Name";

interface ScoredCandidate {
  record: ReconcileRecord;
  matchedFields: CandidateMatchField[];
  score: number;
}

export interface OlbCandidateResolution {
  matchedRecord?: ReconcileRecord;
  remark: string;
}

export class OlbMatcher {
  /**
   * หา AF1 Row ที่ Test Data แถวใดแถวหนึ่งระบุ Primary Key ไว้แล้ว
   * เพื่อป้องกัน Fallback ของ Test Case อื่นมาใช้แถวนั้นก่อน
   */
  findReservedReportRows(
    testDataRecords: ReconcileRecord[],
    reportRecords: ReconcileRecord[],
  ): Set<number> {
    const transactionIds = new Set(
      testDataRecords
        .map((record) =>
          normalizeOlbText(record.get(OLB_TEST_DATA_FIELDS.transactionId)),
        )
        .filter((value) => value !== ""),
    );

    return new Set(
      reportRecords
        .filter((record) =>
          transactionIds.has(
            normalizeOlbText(
              record.get(OLB_REPORT_FIELDS.arrangementNumber),
            ),
          ),
        )
        .map((record) => record.rowNumber),
    );
  }

  /** คำนวณ Field ที่ Candidate ตรง โดยค่าว่างจะไม่นับเป็น Match */
  private scoreCandidate(
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): ScoredCandidate {
    const matchedFields: CandidateMatchField[] = [];

    const expectedDate = parseDate(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.transactionDate),
    );

    if (
      expectedDate &&
      isDateMatchWithArrangementFallback(
        expectedDate,
        reportRecord.get(OLB_REPORT_FIELDS.arrangementContractDate),
        reportRecord.get(OLB_REPORT_FIELDS.arrangementNumber),
      )
    ) {
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

    const expectedCifNo = normalizeOlbId(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.cifNo),
    );
    const actualCustCode = normalizeOlbId(
      reportRecord.get(OLB_REPORT_FIELDS.custCode),
    );

    if (
      expectedCifNo !== "" &&
      actualCustCode !== "" &&
      expectedCifNo === actualCustCode
    ) {
      matchedFields.push("CIF No.");
    }

    const expectedCifName = normalizeOlbText(
      testDataRecord.get(OLB_TEST_DATA_FIELDS.cifName),
    );
    const actualCustName = normalizeOlbText(
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
   * Matching ตามลำดับ:
   * 1. Primary Key พบ 1 แถว -> เลือกทันที
   * 2. Primary Key พบหลายแถว -> เลือกคะแนน Field สนับสนุนสูงสุดแบบไม่เสมอ
   * 3. Primary Key ว่าง/ไม่พบ -> Fallback โดยใช้แถวที่ยังไม่ถูกใช้หรือจอง
   */
  findBestCandidate(
    testDataRecord: ReconcileRecord,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): OlbCandidateResolution {
    const transactionId = normalizeOlbText(
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
              normalizeOlbText(
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
      return this.resolveDuplicatePrimary(testDataRecord, primaryCandidates);
    }

    const fallbackCandidates = unusedReportRecords.filter(
      (record) => !reservedReportRowNumbers.has(record.rowNumber),
    );

    return this.resolveFallback(testDataRecord, fallbackCandidates);
  }

  private resolveDuplicatePrimary(
    testDataRecord: ReconcileRecord,
    primaryCandidates: ReconcileRecord[],
  ): OlbCandidateResolution {
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

  private resolveFallback(
    testDataRecord: ReconcileRecord,
    fallbackCandidates: ReconcileRecord[],
  ): OlbCandidateResolution {
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
          "Fallback Matching พบ Candidate คะแนนสูงสุดเท่ากัน " +
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
}