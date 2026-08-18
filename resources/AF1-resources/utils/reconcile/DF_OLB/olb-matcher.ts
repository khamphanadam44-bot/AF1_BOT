/**
 * olb-matcher.ts
 * ------------------------------------------------------------------
 * เลือก AF1 Row ที่เหมาะสมกับ Test Data โดยไม่ตัดสิน PASS/FAIL
 *
 * Matching Flow:
 * 1. หา Exact Match จาก Transaction ID = FI Arrangement Number
 * 2. ถ้า Exact Match ไม่พบ ให้เข้า Fallback Matching
 * 3. Fallback ตรวจ Candidate ทุก AF1 Row ที่ยังไม่ถูกใช้
 * 4. Amount ต้องตรงกันภายใน OLB_AMOUNT_TOLERANCE
 *    จึงมีสิทธิ์เป็น Candidate
 * 5. Candidate ที่ Amount ผ่าน จะจัดอันดับจาก
 *    Date + CIF No. + CIF Name
 * 6. เลือก AF1 Row เฉพาะเมื่อมีคะแนนสูงสุดเพียงแถวเดียว
 * 7. หากคะแนนสูงสุดเสมอกัน จะไม่เลือก Row อัตโนมัติ
 *
 * หมายเหตุ:
 * - ค่าว่างทั้งสองฝั่งไม่ถือเป็น Exact Match
 * - Amount เป็น Mandatory Gate ไม่ใช่คะแนนสำหรับ Ranking
 * - ไฟล์นี้มีหน้าที่เลือก Candidate เท่านั้น
 *   การตัดสิน PASS/FAIL อยู่ใน olb-analyzer.ts
 * ------------------------------------------------------------------
 */

import type { ReconcileRecord } from "../shared/record";

import {
  isDateMatchWithArrangementFallback,
  parseAmount,
  parseDate,
} from "../shared/reconcile-parse.util";

import {
  OLB_AMOUNT_TOLERANCE,
  OLB_MAX_CANDIDATES_IN_REMARK,
  OLB_REPORT_FIELDS,
  OLB_TEST_DATA_FIELDS,
} from "./olb-config";

import {
  normalizeOlbId,
  normalizeOlbText,
} from "./olb-normalize.util";

/**
 * Field ที่ใช้แสดงว่า Candidate ตรงกับ Test Data ในส่วนใดบ้าง
 *
 * Amount ถูกเก็บไว้ใน matchedFields เพื่อแสดงใน Remark
 * แต่จะไม่นำมาคิด Score เพราะเป็น Mandatory Gate อยู่แล้ว
 */
type CandidateMatchField =
  | "Date"
  | "Amount"
  | "CIF No."
  | "CIF Name";

interface ScoredCandidate {
  record: ReconcileRecord;

  /** Field ที่ Candidate ตรงกับ Test Data */
  matchedFields: CandidateMatchField[];

  /**
   * true เมื่อ Amount ของ Test Data และ AF1
   * ต่างกันไม่เกิน OLB_AMOUNT_TOLERANCE
   */
  amountMatched: boolean;

  /**
   * คะแนนสำหรับ Ranking
   *
   * คิดเฉพาะ:
   * - Date
   * - CIF No.
   * - CIF Name
   *
   * Amount ไม่รวมในคะแนน เพราะเป็น Mandatory Gate
   */
  score: number;
}

export interface OlbCandidateResolution {
  matchedRecord?: ReconcileRecord;
  remark: string;
}

export class OlbMatcher {
  /**
   * สงวน AF1 Row ที่มี FI Arrangement Number
   * ตรงกับ Transaction ID ของ Test Case ใด Test Case หนึ่ง
   *
   * เพื่อป้องกัน Fallback ของ Test Case อื่น
   * นำ Exact Row นี้ไปใช้ก่อน
   */
  findReservedReportRows(
    testDataRecords: ReconcileRecord[],
    reportRecords: ReconcileRecord[],
  ): Set<number> {
    const transactionIds = new Set(
      testDataRecords
        .map((record) =>
          normalizeOlbText(
            record.get(
              OLB_TEST_DATA_FIELDS.transactionId,
            ),
          ),
        )
        /**
         * Transaction ID ว่างไม่สามารถใช้ Exact Matching ได้
         *
         * ป้องกันกรณี:
         * "" === ""
         * แล้วถูกเข้าใจผิดว่าเป็น Exact Match
         */
        .filter((value) => value !== ""),
    );

    return new Set(
      reportRecords
        .filter((record) =>
          transactionIds.has(
            normalizeOlbText(
              record.get(
                OLB_REPORT_FIELDS.arrangementNumber,
              ),
            ),
          ),
        )
        .map((record) => record.rowNumber),
    );
  }

  /**
   * ตรวจ Candidate ทุก Field
   *
   * Amount:
   * - ต้องอ่านเป็นตัวเลขได้ทั้งสองฝั่ง
   * - ผลต่างต้องไม่เกิน OLB_AMOUNT_TOLERANCE
   *
   * Field อื่น:
   * - ค่าว่างจะไม่นับเป็น Match
   */
  private scoreCandidate(
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): ScoredCandidate {
    const matchedFields: CandidateMatchField[] = [];

    // =====================================================
    // Date
    // =====================================================

    const expectedDate = parseDate(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.transactionDate,
      ),
    );

    if (
      expectedDate &&
      isDateMatchWithArrangementFallback(
        expectedDate,
        reportRecord.get(
          OLB_REPORT_FIELDS.arrangementContractDate,
        ),
        reportRecord.get(
          OLB_REPORT_FIELDS.arrangementNumber,
        ),
      )
    ) {
      matchedFields.push("Date");
    }

    // =====================================================
    // Amount
    //
    // Mandatory Gate สำหรับ Fallback Candidate
    // =====================================================

    const expectedAmount = parseAmount(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
      ),
    );

    const actualAmount = parseAmount(
      reportRecord.get(
        OLB_REPORT_FIELDS.thbOutstandingAmount,
      ),
    );

    const amountMatched =
      expectedAmount !== null &&
      actualAmount !== null &&
      Math.abs(
        expectedAmount - actualAmount,
      ) <= OLB_AMOUNT_TOLERANCE;

    if (amountMatched) {
      matchedFields.push("Amount");
    }

    // =====================================================
    // CIF No. / Cust Code
    // =====================================================

    const expectedCifNo = normalizeOlbId(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.cifNo,
      ),
    );

    const actualCustCode = normalizeOlbId(
      reportRecord.get(
        OLB_REPORT_FIELDS.custCode,
      ),
    );

    if (
      expectedCifNo !== "" &&
      actualCustCode !== "" &&
      expectedCifNo === actualCustCode
    ) {
      matchedFields.push("CIF No.");
    }

    // =====================================================
    // CIF Name / Cust Name
    // =====================================================

    const expectedCifName = normalizeOlbText(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.cifName,
      ),
    );

    const actualCustName = normalizeOlbText(
      reportRecord.get(
        OLB_REPORT_FIELDS.custName,
      ),
    );

    if (
      expectedCifName !== "" &&
      actualCustName !== "" &&
      expectedCifName === actualCustName
    ) {
      matchedFields.push("CIF Name");
    }

    /**
     * Amount ไม่ถูกนำมาคิด Ranking Score
     *
     * เพราะ Candidate ทุก Row ที่เข้า Ranking
     * ต้องผ่าน Amount Gate อยู่แล้ว
     */
    const score = matchedFields.filter(
      (field) => field !== "Amount",
    ).length;

    return {
      record: reportRecord,
      matchedFields,
      amountMatched,
      score,
    };
  }

  /**
   * ตรวจ Candidate ทุก Row และเรียงตาม Score จากมากไปน้อย
   */
  private rankCandidates(
    testDataRecord: ReconcileRecord,
    candidates: ReconcileRecord[],
  ): ScoredCandidate[] {
    return candidates
      .map((record) =>
        this.scoreCandidate(
          testDataRecord,
          record,
        ),
      )
      .sort(
        (left, right) =>
          right.score - left.score,
      );
  }

  /**
   * ตรวจ Candidate ทุก Row ก่อน
   * แล้วเหลือเฉพาะ Row ที่ Amount ผ่าน Tolerance
   *
   * ทำให้ Amount เป็น Mandatory Gate
   * ไม่ใช่เพียงหนึ่งคะแนนในการ Matching
   */
  private rankAmountMatchedCandidates(
    testDataRecord: ReconcileRecord,
    candidates: ReconcileRecord[],
  ): ScoredCandidate[] {
    return this.rankCandidates(
      testDataRecord,
      candidates,
    ).filter(
      (candidate) =>
        candidate.amountMatched,
    );
  }

  /**
   * สร้างรายการ Candidate สำหรับ Remark
   *
   * จำกัดจำนวน Row ตาม Config
   * เพื่อไม่ให้ข้อความยาวเกินไป
   */
  private formatCandidateList(
    candidates: ScoredCandidate[],
  ): string {
    return candidates
      .slice(
        0,
        OLB_MAX_CANDIDATES_IN_REMARK,
      )
      .map((candidate) => {
        const matchedFields =
          candidate.matchedFields.length > 0
            ? candidate.matchedFields.join(" + ")
            : "ไม่มี Field ตรง";

        return (
          `Row ${candidate.record.rowNumber}: ` +
          `${candidate.record.get(
            OLB_REPORT_FIELDS.arrangementNumber,
          )} ` +
          `[${matchedFields}]`
        );
      })
      .join(", ");
  }

  /**
   * เลือก AF1 Row ที่เหมาะสมกับ Test Data
   *
   * Priority:
   * 1. Exact Transaction ID
   * 2. Duplicate Exact → Amount + Field อื่นช่วยตัดสิน
   * 3. Exact ไม่พบ → Fallback Matching
   */
  findBestCandidate(
    testDataRecord: ReconcileRecord,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): OlbCandidateResolution {
    const transactionId = normalizeOlbText(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.transactionId,
      ),
    );

    /**
     * Row ที่ถูก Test Case ก่อนหน้าใช้ไปแล้ว
     * จะไม่สามารถถูกนำมา Match ซ้ำ
     */
    const unusedReportRecords =
      reportRecords.filter(
        (record) =>
          !usedReportRowNumbers.has(
            record.rowNumber,
          ),
      );

    /**
     * Exact Matching
     *
     * Transaction ID ว่าง → ไม่หา Exact
     *
     * ทำให้:
     * "" กับ ""
     * ไม่ถูกถือว่าเป็น Exact Match
     */
    const primaryCandidates =
      transactionId === ""
        ? []
        : unusedReportRecords.filter(
            (record) =>
              normalizeOlbText(
                record.get(
                  OLB_REPORT_FIELDS.arrangementNumber,
                ),
              ) === transactionId,
          );

    // =====================================================
    // Exact Match พบเพียง 1 Row
    // =====================================================

    if (primaryCandidates.length === 1) {
      return {
        matchedRecord:
          primaryCandidates[0],
        remark: "",
      };
    }

    // =====================================================
    // Exact Match พบหลาย Row
    // =====================================================

    if (primaryCandidates.length > 1) {
      return this.resolveDuplicatePrimary(
        testDataRecord,
        primaryCandidates,
      );
    }

    // =====================================================
    // Exact Match ไม่พบ
    // → Fallback
    //
    // ห้ามนำ Reserved Exact Row ของ Case อื่นมาใช้
    // =====================================================

    const fallbackCandidates =
      unusedReportRecords.filter(
        (record) =>
          !reservedReportRowNumbers.has(
            record.rowNumber,
          ),
      );

    return this.resolveFallback(
      testDataRecord,
      fallbackCandidates,
    );
  }

  /**
   * Transaction ID ตรงกับ AF1 มากกว่า 1 Row
   *
   * ใช้ Amount เป็น Mandatory Gate
   * จากนั้นใช้ Date + CIF No. + CIF Name ช่วยเลือก
   *
   * จะเลือกได้เฉพาะเมื่อมี Best Candidate เพียง Row เดียว
   */
  private resolveDuplicatePrimary(
    testDataRecord: ReconcileRecord,
    primaryCandidates: ReconcileRecord[],
  ): OlbCandidateResolution {
    const ranked =
      this.rankAmountMatchedCandidates(
        testDataRecord,
        primaryCandidates,
      );

    /**
     * Exact Transaction ID ซ้ำ
     * แต่ไม่มี Row ไหน Amount ตรง
     */
    if (ranked.length === 0) {
      return {
        remark:
          `Primary Matching พบ Transaction ID ตรงหลายแถว ` +
          `(${primaryCandidates.length} แถว) ` +
          `แต่ไม่มี AF1 Row ที่ Amount ตรงภายใน Tolerance ` +
          `(${OLB_AMOUNT_TOLERANCE}) ` +
          "จึงไม่เลือก AF1 Row อัตโนมัติ",
      };
    }

    const bestScore =
      ranked[0].score;

    const bestCandidates =
      ranked.filter(
        (candidate) =>
          candidate.score === bestScore,
      );

    // =====================================================
    // มี Best Candidate เพียง Row เดียว
    // =====================================================

    if (bestCandidates.length === 1) {
      const best =
        bestCandidates[0];

      return {
        matchedRecord: best.record,

        remark:
          `Primary Matching พบ Transaction ID ตรงหลายแถว ` +
          `(${primaryCandidates.length} แถว) ` +
          `แต่เลือก AF1 Row ${best.record.rowNumber} ได้ ` +
          `เนื่องจาก Amount ตรงและมีข้อมูลตรงมากที่สุด ` +
          `(${best.matchedFields.join(" + ")})`,
      };
    }

    // =====================================================
    // คะแนนสูงสุดยังเสมอกัน
    // =====================================================

    return {
      remark:
        `Primary Matching พบ Transaction ID ตรงหลายแถว ` +
        `และ Candidate ที่ Amount ตรงมีคะแนน Field อื่นสูงสุดเท่ากัน ` +
        `(${bestCandidates.length} แถว, คะแนน ${bestScore}) ` +
        "จึงไม่เลือก AF1 Row อัตโนมัติ\n" +
        this.formatCandidateList(
          bestCandidates,
        ),
    };
  }

  /**
   * Fallback Matching
   *
   * Rule:
   * 1. Test Data Amount ต้องอ่านเป็นตัวเลขได้
   * 2. ตรวจ Candidate ทุก AF1 Row
   * 3. Amount ต้องผ่าน OLB_AMOUNT_TOLERANCE
   * 4. ตัด Row ที่ Amount ไม่ผ่านออก
   * 5. ใช้ Date + CIF No. + CIF Name คิด Ranking Score
   * 6. เลือกเฉพาะเมื่อ Best Candidate มีเพียง Row เดียว
   */
  private resolveFallback(
    testDataRecord: ReconcileRecord,
    fallbackCandidates: ReconcileRecord[],
  ): OlbCandidateResolution {
    // =====================================================
    // Test Data Amount ต้องใช้ Matching ได้
    // =====================================================

    const expectedAmount = parseAmount(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount,
      ),
    );

    if (expectedAmount === null) {
      return {
        remark:
          "Fallback Matching ไม่สามารถทำได้ เนื่องจาก " +
          `"${OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount}" ` +
          "ใน Test Data เป็นค่าว่างหรือไม่สามารถอ่านเป็นตัวเลขได้",
      };
    }

    // =====================================================
    // ตรวจทุก Candidate
    // แล้วเหลือเฉพาะ Row ที่ Amount ผ่าน Tolerance
    // =====================================================

    const ranked =
      this.rankAmountMatchedCandidates(
        testDataRecord,
        fallbackCandidates,
      );

    // =====================================================
    // ไม่มี AF1 Row ที่ Amount ตรง
    // =====================================================

    if (ranked.length === 0) {
      return {
        remark:
          "Fallback Matching ไม่พบ AF1 Row ที่ " +
          `${OLB_REPORT_FIELDS.thbOutstandingAmount} ` +
          "ตรงกับ " +
          `${OLB_TEST_DATA_FIELDS.thbEquivalentTransferAmount} ` +
          `(${expectedAmount}) ` +
          `ภายใน Tolerance ${OLB_AMOUNT_TOLERANCE}`,
      };
    }

    // =====================================================
    // หา Score สูงสุด
    //
    // Score ไม่รวม Amount
    // เพราะทุก Candidate ผ่าน Amount Gate แล้ว
    // =====================================================

    const bestScore =
      ranked[0].score;

    const bestCandidates =
      ranked.filter(
        (candidate) =>
          candidate.score === bestScore,
      );

    // =====================================================
    // คะแนนสูงสุดเสมอกันหลาย Row
    // → ไม่เลือก
    // =====================================================

    if (bestCandidates.length > 1) {
      return {
        remark:
          "Fallback Matching พบ AF1 Candidate ที่ Amount ตรง " +
          "และมีคะแนน Field อื่นสูงสุดเท่ากัน " +
          `(${bestCandidates.length} แถว, คะแนน ${bestScore}) ` +
          "จึงไม่เลือก AF1 Row อัตโนมัติ\n" +
          this.formatCandidateList(
            bestCandidates,
          ),
      };
    }

    // =====================================================
    // มี Best Candidate เพียง Row เดียว
    // =====================================================

    const best =
      bestCandidates[0];

    return {
      matchedRecord: best.record,

      remark:
        `Fallback Matching เลือก AF1 Row ${best.record.rowNumber} ` +
        `เนื่องจาก Amount ตรงภายใน Tolerance ` +
        `และมีข้อมูลตรงกับ Test Data มากที่สุด ` +
        `(${best.matchedFields.join(" + ")})\n` +
        `${OLB_REPORT_FIELDS.arrangementNumber} = ` +
        `"${best.record.get(
          OLB_REPORT_FIELDS.arrangementNumber,
        )}"`,
    };
  }
}