/**
 * FieldComparer
 * ------------------------------------------------------------------
 * เปรียบเทียบ "1 field rule เดียว" ระหว่าง Test Data (expected) กับ AF1 Report (actual)
 * รองรับทั้ง 4 compareMode: fixedValue / amountTolerance / exact / dateWithIdFallback
 * ใช้ร่วมกันทั้ง CoreFieldValidator และ ConditionalFieldValidator
 * ------------------------------------------------------------------
 */
import type { ReconcileFieldRule } from "./ltx-config";
import { AmountComparator } from "./ltx-amount-compare";
import { ReconcileRecord } from "../shared/record";
import { formatCompareRemark, formatFixedValueRemark } from "../shared/remark";
import {
  isDateMatchWithArrangementFallback,
  parseDate,
} from "../shared/reconcile-parse.util";

/** สถานะผลเปรียบเทียบของ Field หนึ่งรายการ */
export type ReconcileStatus = "PASS" | "FAIL" | "REVIEW";

/** ผลเปรียบเทียบ Field หนึ่งรายการของคู่ Test Data และ AF1 Report */
export interface FieldCheckResult {
  fieldHeader: string;
  isMatch: boolean;
  status: ReconcileStatus;
  remark: string;
}

export class FieldComparer {
  constructor(
    private readonly amountComparator: AmountComparator = new AmountComparator(),
  ) {}

  private isNumericId(value: string): boolean {
    return /^\d+$/.test(value.trim());
  }

  /** normalize รหัสตัวเลข (เช่น Cust Code) ก่อนเทียบ — ตัด leading zero ทิ้ง */
  private normalizeIdForCompare(value: string): string {
    const trimmed = value.trim();

    if (!this.isNumericId(trimmed)) {
      return trimmed.toLowerCase();
    }

    const normalized = trimmed.replace(/^0+(?=\d)/, "");

    return normalized === "" ? "0" : normalized;
  }

  /** แปลงค่าที่ใช้แสดงใน Remark — ป้องกัน null/undefined และตัดช่องว่างหัวท้าย */
  /**
   * สร้างข้อความเมื่อข้อมูลระหว่าง Test Script และ AF1 Report ไม่ตรงกัน (สถานะ REVIEW)
   * Format ตาม Requirement:
   *   [TS] : <TestDataField> = "<value>" | [AF1-<ReportLabel>] : <ReportField> = "<value>"
   *

   */
  private buildReviewRemark(
    reportCode: string,
    rule: ReconcileFieldRule,
    expected: string,
    actual: string,
  ): string {
    const testDataLabel = rule.testDataField ?? "Test Data";

    return formatCompareRemark(
      reportCode,
      testDataLabel,
      expected,
      rule.reportField,
      actual,
    );
  }

  /** สร้างข้อความเมื่อ AF1 Report ผิด business rule แบบชัดเจน/ไม่กำกวม (สถานะ FAIL) */
  private buildFailRemark(
    reportCode: string,
    rule: ReconcileFieldRule,
    expected: string,
    actual: string,
  ): string {
    return formatFixedValueRemark(
      reportCode,
      rule.reportField,
      actual,
      expected,
    );
  }

  private toFieldCheckResult(
    fieldHeader: string,
    status: ReconcileStatus,
    remark: string,
  ): FieldCheckResult {
    return { fieldHeader, status, isMatch: status === "PASS", remark };
  }

  /**
   * เปรียบเทียบ compareMode = "fixedValue" (ค่าคงที่ตาม Requirement เช่น Payment Method ต้องเป็น 234004)
   *
   * กติกาตาม Requirement (ตัวอย่าง Payment Method):
   * - AF1 Report ไม่ตรงกับค่าคงที่ -> FAIL เสมอ (ไม่สนใจฝั่ง Test Data)
   * - AF1 Report ตรงกับค่าคงที่ แต่ Test Data (ฝั่งอ้างอิง) ว่าง -> REVIEW
   * - AF1 Report ตรงกับค่าคงที่ และ Test Data มีค่า (หรือไม่ได้กำหนด field ให้ cross-check) -> PASS
   */
  private checkFixedValue(
    reportCode: string,
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    actual: string,
  ): FieldCheckResult {
    const expected = rule.fixedValue ?? "";
    const isCorrect =
      actual.trim().toLowerCase() === expected.trim().toLowerCase();

    if (!isCorrect) {
      return this.toFieldCheckResult(
        rule.reportField,
        "FAIL",
        this.buildFailRemark(reportCode, rule, expected, actual),
      );
    }

    if (rule.testDataField) {
      const crossCheckValue = testDataRecord.get(rule.testDataField);
      if (crossCheckValue.trim() === "") {
        return this.toFieldCheckResult(
          rule.reportField,
          "REVIEW",
          this.buildReviewRemark(reportCode, rule, crossCheckValue, actual),
        );
      }
    }

    return this.toFieldCheckResult(rule.reportField, "PASS", "");
  }

  /**
   * เปรียบเทียบ compareMode = "dateWithIdFallback"
   * รอบที่ 1: เทียบ Transaction Date ตรง ๆ ก่อน
   * รอบที่ 2: ถ้าไม่ตรง ให้ดึงวันที่จากตำแหน่ง 7-12 ของ Reference Transaction Number มาเทียบแทน
   * ถ้าเทียบทั้ง 2 รอบแล้วยังไม่ตรง -> REVIEW (ไม่ใช่ FAIL ตาม Requirement:
   * "ห้ามให้ผลเป็น Please review ก่อนเช็คขั้นที่ 2" หมายความว่าหลังเช็คครบ 2 ขั้นแล้ว
   * ผลที่เป็นไปได้คือ Pass หรือ Please Review เท่านั้น ไม่มี Fail สำหรับ field นี้)
   */
  private checkDateWithIdFallback(
    reportCode: string,
    rule: ReconcileFieldRule,
    reportRecord: ReconcileRecord,
    expected: string,
    actual: string,
  ): FieldCheckResult {
    const expectedDate = parseDate(expected);
    const referenceNumber = reportRecord.get(
      "Reference Transaction Number",
    );

    if (
      expectedDate &&
      isDateMatchWithArrangementFallback(
        expectedDate,
        actual,
        referenceNumber,
      )
    ) {
      return this.toFieldCheckResult(rule.reportField, "PASS", "");
    }

    // เทียบทั้ง 2 รอบแล้วยังไม่ตรง -> ต้องตรวจสอบเพิ่มเติม (REVIEW)
    return this.toFieldCheckResult(
      rule.reportField,
      "REVIEW",
      this.buildReviewRemark(reportCode, rule, expected, actual),
    );
  }

  /**
   * เปรียบเทียบ compareMode = "amountTolerance" — ตัวเลข ยอมรับส่วนต่าง +-tolerance
   * ถ้าฝั่งใดฝั่งหนึ่งแปลงเป็นตัวเลขไม่ได้ (ว่าง/ไม่ใช่ตัวเลข) หรือส่วนต่างเกิน tolerance -> REVIEW
   *
   * ถ้าเทียบกับ testDataField หลักแล้วไม่ผ่าน และมี fallbackTestDataField กำหนดไว้
   * จะลองเทียบกับ fallback อีกรอบก่อนสรุปว่า REVIEW (ตาม Business Rule: "หาก Transaction
   * Amount ไม่ตรงกับ From Transfer Amount ให้ไปดูที่ From Debit Amount")
   */
  private checkAmountTolerance(
    reportCode: string,
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    expected: string,
    actual: string,
  ): FieldCheckResult {
    const result = this.amountComparator.compare(
      expected,
      actual,
      rule.tolerance,
    );

    if (result.isMatch) {
      return this.toFieldCheckResult(rule.reportField, "PASS", "");
    }

    if (rule.fallbackTestDataField) {
      const fallbackExpected = testDataRecord.get(rule.fallbackTestDataField);
      const fallbackResult = this.amountComparator.compare(
        fallbackExpected,
        actual,
        rule.tolerance,
      );

      if (fallbackResult.isMatch) {
        return this.toFieldCheckResult(rule.reportField, "PASS", "");
      }
    }

    return this.toFieldCheckResult(
      rule.reportField,
      "REVIEW",
      this.buildReviewRemark(reportCode, rule, expected, actual),
    );
  }

  /**
   * เปรียบเทียบ compareMode = "exact" — ต้องตรงกันเป๊ะหลัง normalize (ตัด leading zero ของรหัสตัวเลข)
   * ฝั่งใดฝั่งหนึ่งว่าง หรือมีค่าทั้ง 2 ฝั่งแต่ไม่ตรงกัน -> REVIEW (ตามตัวอย่าง Cust Name ใน Requirement)
   */
  private checkExact(
    reportCode: string,
    rule: ReconcileFieldRule,
    expected: string,
    actual: string,
  ): FieldCheckResult {
    const isEitherEmpty = expected.trim() === "" || actual.trim() === "";

    if (isEitherEmpty) {
      return this.toFieldCheckResult(
        rule.reportField,
        "REVIEW",
        this.buildReviewRemark(reportCode, rule, expected, actual),
      );
    }

    const isMatch =
      this.normalizeIdForCompare(expected) ===
      this.normalizeIdForCompare(actual);

    if (isMatch) {
      return this.toFieldCheckResult(rule.reportField, "PASS", "");
    }

    return this.toFieldCheckResult(
      rule.reportField,
      "REVIEW",
      this.buildReviewRemark(reportCode, rule, expected, actual),
    );
  }

  check(
    reportCode: string,
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): FieldCheckResult {
    const actual = reportRecord.get(rule.reportField);

    if (rule.compareMode === "fixedValue") {
      return this.checkFixedValue(reportCode, rule, testDataRecord, actual);
    }

    const expected = rule.testDataField
      ? testDataRecord.get(rule.testDataField)
      : "";

    if (rule.compareMode === "dateWithIdFallback") {
      return this.checkDateWithIdFallback(
        reportCode,
        rule,
        reportRecord,
        expected,
        actual,
      );
    }

    if (rule.compareMode === "amountTolerance") {
      return this.checkAmountTolerance(
        reportCode,
        rule,
        testDataRecord,
        expected,
        actual,
      );
    }

    // compareMode === "exact"
    return this.checkExact(reportCode, rule, expected, actual);
  }

  /** เช็คว่า Field rule นี้ใช้กับแถว suffix ปัจจุบันไหม (เช่น field จำกัดไว้เฉพาะแถว DR) */
  isApplicableToSuffix(rule: ReconcileFieldRule, suffix: string): boolean {
    return !rule.applicableSuffixes || rule.applicableSuffixes.includes(suffix);
  }
}