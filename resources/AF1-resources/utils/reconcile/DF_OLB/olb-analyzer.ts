/**
 * olb-analyzer.ts
 * ------------------------------------------------------------------
 * วิเคราะห์ Candidate ของ DF_OLB และสร้างผล PASS/FAIL/Review
 *
 * หน้าที่:
 * 1. รับ AF1 Row ที่เลือกมาจาก olb-matcher.ts
 * 2. ตรวจ Expected Absence ของ DF_OLB ก่อนตัดสินผลปกติ
 *    - ไม่มี THB ใน From/To Currency
 *    - From Customer เป็น Resident
 * 3. Expected Absence: ไม่พบ AF1 → PASS, พบ AF1 → FAIL
 * 4. ถ้า Expected Presence แต่ Matcher เลือก Row ไม่ได้ → FAIL
 * 5. Primary Key, Date และ Amount ใช้กำหนด PASS/FAIL
 * 6. CIF Number, CIF Name และ Arrangement Date Fallback เป็น Review-only
 * 7. Amount เปรียบเทียบโดยใช้ OLB_AMOUNT_TOLERANCE
 * ------------------------------------------------------------------
 */

import type { ReconcileRecord } from "../shared/record";
import { formatCompareRemark } from "../shared/remark";
import type { ResultRow } from "../shared/result-writer";
import {
  extractDateFromArrangementNumber,
  formatDate,
  isSameDate,
  parseAmount,
  parseDate,
} from "../shared/reconcile-parse.util";
import {
  OLB_AMOUNT_TOLERANCE,
  OLB_EXCLUDE_REASONS,
  OLB_FIELD_MAPPINGS,
  OLB_REMARKS,
  OLB_REPORT_CODE,
  OLB_REPORT_FIELDS,
  OLB_RESIDENT_VALUE,
  OLB_TEST_DATA_FIELDS,
  OLB_THB_CURRENCY_CODE,
} from "./olb-config";
import type { OlbFieldMapping } from "./olb-config";
import type { OlbCandidateResolution } from "./olb-matcher";
import {
  normalizeOlbId,
  normalizeOlbText,
} from "./olb-normalize.util";

/**
 * Callback สำหรับเพิ่ม Compare Remark
 *
 * ใช้ร่วมกันทั้ง:
 * - Failure
 * - Review
 */
type AddComparisonRemark = (
  mapping: OlbFieldMapping,
  expected: string,
  actual: string,
  message?: string,
) => void;

export class OlbAnalyzer {
  /**
   * วิเคราะห์ Test Data 1 Row
   *
   * ถ้า Test No. ว่าง จะใช้เลขแถวของ Test Data
   * เพื่อให้สามารถย้อนกลับไปตรวจข้อมูลต้นทางได้
   */
  analyze(
    testDataRecord: ReconcileRecord,
    candidate: OlbCandidateResolution,
    presenceContextRecord: ReconcileRecord = testDataRecord,
  ): ResultRow {
    const testCaseNo =
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.testNo,
      ).trim() ||
      `TEST DATA ROW ${testDataRecord.rowNumber}`;

    // =====================================================
    // Expected Absence / Exclude Rules
    // =====================================================

    const excludeReasons =
      this.findExcludeReasons(
        presenceContextRecord,
      );

    if (excludeReasons.length > 0) {
      return this.buildExpectedAbsenceResult(
        testCaseNo,
        candidate,
        excludeReasons,
      );
    }

    // =====================================================
    // Expected Presence
    // =====================================================

    /**
     * รายการนี้ควรอยู่ใน DF_OLB
     * แต่ Matcher ยังไม่สามารถระบุ AF1 Row ที่ถูกต้องได้
     *
     * เช่น:
     * - ไม่มี Amount ที่ตรง
     * - Candidate คะแนนสูงสุดเสมอกัน
     * - Amount ฝั่ง Test Data อ่านไม่ได้
     */
    if (!candidate.matchedRecord) {
      return this.buildUnresolvedResult(
        testCaseNo,
        candidate.remark,
      );
    }

    /**
     * Matcher เลือก AF1 Row ได้แล้ว
     * จึงเริ่ม Compare Field จริง
     */
    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      candidate.matchedRecord,
      candidate.remark,
    );
  }

  /**
   * ตรวจเงื่อนไข Expected Absence ของ DF_OLB
   *
   * Rule ปัจจุบัน:
   *  มีข้อมูลครบ และไม่มี THB ทั้งสองฝั่ง
   *  เป็น Resident
   * Test Case เดียวสามารถเข้า Exclude ได้มากกว่า 1 เหตุผล
   */
  private findExcludeReasons(
    testDataRecord: ReconcileRecord,
  ): string[] {
    const reasons: string[] = [];

    // =====================================================
    // Rule 1: ไม่มี THB
    // =====================================================

    const fromCurrency =
      normalizeOlbText(
        testDataRecord.get(
          OLB_TEST_DATA_FIELDS.fromCurrency,
        ),
      );

    const toCurrency =
      normalizeOlbText(
        testDataRecord.get(
          OLB_TEST_DATA_FIELDS.toCurrency,
        ),
      );

    /**
     * ต้องมี Currency ครบทั้งสองฝั่งก่อน
     * จึงจะสรุปว่า "ไม่มี THB"
     *
     * ถ้าฝั่งใดฝั่งหนึ่งว่าง จะไม่ตีความเองว่าเป็น Expected Absence
     */
    const hasCompleteCurrencyData =
      fromCurrency !== "" &&
      toCurrency !== "";

    const hasThb =
      fromCurrency === OLB_THB_CURRENCY_CODE ||
      toCurrency === OLB_THB_CURRENCY_CODE;

    if (
      hasCompleteCurrencyData &&
      !hasThb
    ) {
      reasons.push(
        `${OLB_EXCLUDE_REASONS.noThb} ` +
          `(From CCY = "${fromCurrency}", ` +
          `To CCY = "${toCurrency}")`,
      );
    }

    // =====================================================
    // Rule 2: Resident
    // =====================================================

    const residency =
      normalizeOlbText(
        testDataRecord.get(
          OLB_TEST_DATA_FIELDS.fromResidency,
        ),
      );

    if (
      residency === OLB_RESIDENT_VALUE
    ) {
      reasons.push(
        `${OLB_EXCLUDE_REASONS.resident} ` +
          `(From Customer = "${residency}")`,
      );
    }

    return reasons;
  }

  /**
   * สร้างผลสำหรับ Test Case ที่เข้า Exclude Rule
   *
   * Expected Absence:
   * - ไม่พบ AF1 Row → PASS
   * - พบ AF1 Row → FAIL เพราะเป็น Unexpected Presence
   */
  private buildExpectedAbsenceResult(
    testCaseNo: string,
    candidate: OlbCandidateResolution,
    excludeReasons: string[],
  ): ResultRow {
    const reasonRemark =
      excludeReasons
        .map(
          (reason, index) =>
            `${index + 1}. ${reason}`,
        )
        .join("\n");

    // =====================================================
    // ไม่พบใน DF_OLB ตามที่คาดหวัง
    // =====================================================

    if (!candidate.matchedRecord) {
      return {
        testCaseNo,
        status: "PASS",
        remark:
          `ไม่พบรายการใน ${OLB_REPORT_CODE} ตามที่คาดหวัง ` +
          `เนื่องจากเข้าเงื่อนไข Exclude:\n` +
          reasonRemark,
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: [],
        isExpectedAbsence: true,
      };
    }

    // =====================================================
    // พบใน DF_OLB ทั้งที่ควรถูก Exclude
    // =====================================================

    const remarks = [
      `พบรายการใน ${OLB_REPORT_CODE} ` +
        "ทั้งที่ Test Case เข้าเงื่อนไข Exclude",
      `เหตุผล Exclude:\n${reasonRemark}`,
      candidate.remark,
    ]
      .filter(
        (remark) =>
          remark.trim() !== "",
      )
      .join("\n");

    return {
      testCaseNo,
      status: "FAIL",
      remark: remarks,
      matchedRowNumber:
        candidate.matchedRecord.rowNumber,
      failedKeyFieldHeaders: [
        OLB_REPORT_FIELDS.arrangementNumber,
      ],
      reviewFieldHeaders: [],
      isExpectedAbsence: true,
    };
  }

  /**
   * สร้างผลกรณี Matcher ไม่สามารถระบุ AF1 Row ได้
   */
  private buildUnresolvedResult(
    testCaseNo: string,
    candidateRemark: string,
  ): ResultRow {
    const remark =
      candidateRemark.trim() ||
      "ไม่สามารถระบุ AF1 Row ที่ตรงกับ Test Data ได้";

    return {
      testCaseNo,
      status: "FAIL",
      remark,
      matchedRowNumber: undefined,

      /**
       * Matching ไม่สามารถระบุ Row ได้
       * จึงให้ FI Arrangement Number เป็น Key Field
       * ที่เกี่ยวข้องกับ Failure
       */
      failedKeyFieldHeaders: [
        OLB_REPORT_FIELDS.arrangementNumber,
      ],

      reviewFieldHeaders: [],
      isExpectedAbsence: false,
    };
  }

  /**
   * ตรวจทุก Field ของ AF1 Row ที่ Matcher เลือกมา
   *
   * หลังตรวจครบทุก Field แล้ว
   * จึงสรุปสถานะ PASS / FAIL ตอนท้าย
   */
  private compareMatchedRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    candidateRemark: string,
  ): ResultRow {
    const failedHeaders = new Set<string>();
    const reviewHeaders = new Set<string>();

    /**
     * ถ้า Matcher มี Remark เช่น Fallback Matching
     * ให้เก็บไว้เป็น Remark แรก
     */
    const remarks: string[] =
      candidateRemark.trim() !== ""
        ? [candidateRemark]
        : [];

    /**
     * เพิ่ม Field ที่ FAIL
     */
    const addFailure: AddComparisonRemark = (
      mapping,
      expected,
      actual,
      message,
    ): void => {
      failedHeaders.add(
        mapping.reportField,
      );

      if (message?.trim()) {
        remarks.push(message);
      }

      remarks.push(
        this.buildRemark(
          mapping,
          expected,
          actual,
        ),
      );
    };

    /**
     * เพิ่ม Field ที่ต้อง Review
     */
    const addReview: AddComparisonRemark = (
      mapping,
      expected,
      actual,
    ): void => {
      reviewHeaders.add(
        mapping.reportField,
      );

      remarks.push(
        this.buildRemark(
          mapping,
          expected,
          actual,
        ),
      );
    };

    // =====================================================
    // Primary Key
    // =====================================================

    this.comparePrimaryKey(
      testDataRecord,
      matchedRecord,
      addFailure,
      addReview,
    );

    // =====================================================
    // Date
    // =====================================================

    this.compareDate(
      testDataRecord,
      matchedRecord,
      addFailure,
      addReview,
    );

    // =====================================================
    // Amount
    // =====================================================

    this.compareAmount(
      testDataRecord,
      matchedRecord,
      addFailure,
    );

    // =====================================================
    // CIF No.
    // =====================================================

    this.compareCifNo(
      testDataRecord,
      matchedRecord,
      addReview,
    );

    // =====================================================
    // CIF Name
    // =====================================================

    this.compareCifName(
      testDataRecord,
      matchedRecord,
      addReview,
    );

    // =====================================================
    // สรุป Status หลังตรวจทุก Field
    // =====================================================

    const status =
      failedHeaders.size === 0
        ? "PASS"
        : "FAIL";

    /**
     * หากมี Review Field อย่างน้อย 1 Field
     * ให้เพิ่มข้อความ Please Review
     */
    if (reviewHeaders.size > 0) {
      remarks.push(
        OLB_REMARKS.pleaseReview,
      );
    }

    /**
     * PASS จะมี Remark "Matched"
     * แล้วตามด้วย Remark อื่นที่เกี่ยวข้อง
     */
    const finalRemark = [
      status === "PASS"
        ? OLB_REMARKS.matched
        : "",
      ...remarks,
    ]
      .filter(
        (remark) =>
          remark.trim() !== "",
      )
      .join("\n");

    return {
      testCaseNo,
      status,
      remark: finalRemark,
      matchedRowNumber: matchedRecord.rowNumber,
      failedKeyFieldHeaders: [
        ...failedHeaders,
      ],
      reviewFieldHeaders: [
        ...reviewHeaders,
      ],
      isExpectedAbsence: false,
    };
  }

  /**
   * สร้าง Compare Remark รูปแบบกลาง
   */
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

  /**
   * ตรวจ Primary Key
   *
   * Test Data:
   * Transaction ID / Reconcile ID
   *
   * AF1:
   * FI Arrangement Number
   *
   * Rule:
   * - Transaction ID ว่าง → Review
   * - มีค่าและตรงกัน → ผ่าน
   * - มีค่าแต่ไม่ตรงกัน → FAIL
   *
   * หมายเหตุ:
   * Fallback Matching มีหน้าที่ค้นหา AF1 Row ที่เหมาะสม
   * แต่ไม่ได้ทำให้ Primary Key ที่ไม่ตรงกลายเป็น PASS
   */
  private comparePrimaryKey(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const mapping =
      OLB_FIELD_MAPPINGS.primary;

    const expected =
      testDataRecord.get(
        mapping.testDataField,
      );

    const actual =
      matchedRecord.get(
        mapping.reportField,
      );

    const normalizedExpected =
      normalizeOlbText(expected);

    const normalizedActual =
      normalizeOlbText(actual);

    /**
     * Transaction ID ฝั่ง Test Data ว่าง
     *
     * Matcher อาจหา Row ได้จาก Fallback
     * แต่ Primary Key ยังต้อง Review
     */
    if (normalizedExpected === "") {
      addReview(
        mapping,
        expected,
        actual,
      );

      return;
    }

    /**
     * มี Transaction ID แต่ไม่ตรงกับ
     * FI Arrangement Number
     */
    if (
      normalizedExpected !==
      normalizedActual
    ) {
      addFailure(
        mapping,
        expected,
        actual,
      );
    }
  }

  /**
   * ตรวจ Date
   * 1. Txn Date = Arrangement Contract Date
   *    → ผ่าน
   * 2. Arrangement Contract Date ไม่ตรง
   *    แต่ Date ที่อ่านจาก FI Arrangement Number ตรง
   *    → Review-only
   * 3. Date ไม่ตรงทั้งสองทาง
   *    → FAIL
   */
  private compareDate(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const mapping =
      OLB_FIELD_MAPPINGS.date;

    const expectedText =
      testDataRecord.get(
        mapping.testDataField,
      );

    const actualText =
      matchedRecord.get(
        mapping.reportField,
      );

    const expectedDate =
      parseDate(expectedText);

    /**
     * Test Data Date ว่าง
     * หรือไม่สามารถ Parse เป็น Date ได้
     */
    if (!expectedDate) {
      addReview(
        mapping,
        expectedText,
        actualText,
      );

      return;
    }

    const actualDate =
      parseDate(actualText);

    /**
     * Arrangement Contract Date ตรงโดยตรง
     */
    if (
      actualDate &&
      isSameDate(
        expectedDate,
        actualDate,
      )
    ) {
      return;
    }

    /**
     * Date ไม่ตรงโดยตรง
     * ลองอ่าน Date จาก FI Arrangement Number
     */
    const arrangementDate =
      extractDateFromArrangementNumber(
        matchedRecord.get(
          OLB_REPORT_FIELDS.arrangementNumber,
        ),
      );

    /**
     * Date จาก Arrangement Number ตรง
     *
     * ถือเป็น Review-only
     */
    if (
      arrangementDate &&
      isSameDate(
        expectedDate,
        arrangementDate,
      )
    ) {
      addReview(
        mapping,
        expectedText,
        `${actualText} | ` +
          `FI Arrangement Number Date: ` +
          `${formatDate(arrangementDate)}`,
      );

      return;
    }

    /**
     * Date ไม่ตรงทั้ง:
     * - Arrangement Contract Date
     * - FI Arrangement Number Date
     */
    addFailure(
      mapping,
      expectedText,
      `${actualText} | ` +
        `FI Arrangement Number Date: ` +
        `${formatDate(arrangementDate)}`,
    );
  }

  /**
   * ตรวจ Amount
   *
   * Test Data:
   * From THB Equivalent Transfer Amount
   *
   * AF1:
   * THB Outstanding Amount
   *
   * Rule:
   * - Test Data Amount ว่าง/อ่านไม่ได้ → FAIL
   * - AF1 Amount ว่าง/อ่านไม่ได้ → FAIL
   * - ผลต่าง <= OLB_AMOUNT_TOLERANCE → ผ่าน
   * - ผลต่าง > OLB_AMOUNT_TOLERANCE → FAIL
   *
   * Rule นี้ต้องตรงกับ Amount Gate
   * ที่ใช้ใน olb-matcher.ts
   */
  private compareAmount(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
  ): void {
    const mapping =
      OLB_FIELD_MAPPINGS.amount;

    const expectedText =
      testDataRecord.get(
        mapping.testDataField,
      );

    const actualText =
      matchedRecord.get(
        mapping.reportField,
      );

    const expectedAmount =
      parseAmount(expectedText);

    const actualAmount =
      parseAmount(actualText);

    /**
     * Test Data Amount ว่างหรืออ่านไม่ได้
     *
     * Expected Absence ถูกตรวจและ return ไปก่อนหน้านี้แล้ว
     * ดังนั้นเมื่อมาถึง compareAmount() หมายถึงรายการนี้เป็น
     * Expected Presence และ Amount ต้องเป็นข้อมูลบังคับ
     *
     * กรณี Exact Match อาจเลือก AF1 Row ได้จาก Transaction ID
     * แม้ Test Data Amount จะว่าง จึงต้อง FAIL ที่ Analyzer
     */
    if (expectedAmount === null) {
      addFailure(
        mapping,
        expectedText,
        actualText,
         `FAIL - Test Data ไม่มี ${mapping.testDataField} ` +
            `กรุณากรอกข้อมูล Amount ก่อน Reconcile`,
      );

      return;
    }

    /**
     * AF1 Amount อ่านไม่ได้
     * หรือผลต่างเกิน Tolerance
     */
    if (
      actualAmount === null ||
      Math.abs(
        expectedAmount - actualAmount,
      ) > OLB_AMOUNT_TOLERANCE
    ) {
      addFailure(
        mapping,
        expectedText,
        actualText,
      );
    }
  }

  /**
   * ตรวจ CIF Number
   *
   * Test Data:
   * From CIF No. (Client/Sender)
   *
   * AF1:
   * Cust Code
   *
   * เป็นข้อมูลประกอบ:
   * - ค่าว่าง
   * - หรือค่าไม่ตรง
   *
   * → Review-only
   */
  private compareCifNo(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const mapping =
      OLB_FIELD_MAPPINGS.cifNo;

    const expected =
      testDataRecord.get(
        mapping.testDataField,
      );

    const actual =
      matchedRecord.get(
        mapping.reportField,
      );

    const normalizedExpected =
      normalizeOlbId(expected);

    const normalizedActual =
      normalizeOlbId(actual);

    if (
      normalizedExpected === "" ||
      normalizedActual === "" ||
      normalizedExpected !== normalizedActual
    ) {
      addReview(
        mapping,
        expected,
        actual,
      );
    }
  }

  /**
   * ตรวจ CIF Name
   *
   * Test Data:
   * From CIF Name (Client/Sender)
   *
   * AF1:
   * Cust Name
   *
   * เป็นข้อมูลประกอบ:
   * - ค่าว่าง
   * - หรือค่าไม่ตรง
   *
   * → Review-only
   */
  private compareCifName(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const mapping =
      OLB_FIELD_MAPPINGS.cifName;

    const expected =
      testDataRecord.get(
        mapping.testDataField,
      );

    const actual =
      matchedRecord.get(
        mapping.reportField,
      );

    const normalizedExpected =
      normalizeOlbText(expected);

    const normalizedActual =
      normalizeOlbText(actual);

    if (
      normalizedExpected === "" ||
      normalizedActual === "" ||
      normalizedExpected !== normalizedActual
    ) {
      addReview(
        mapping,
        expected,
        actual,
      );
    }
  }
}