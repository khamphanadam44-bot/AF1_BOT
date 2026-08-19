/**
 * เปรียบเทียบหนึ่ง Field Rule ระหว่าง Test Data กับ DS_LTX Report
 * และคืนสถานะ PASS, FAIL หรือ REVIEW ให้ Field Validator นำไปสรุปผลต่อ
 */

import { formatCompareRemark, formatFixedValueRemark } from "../shared/remark";
import {
  isDateMatchWithArrangementFallback,
  parseDate,
} from "../shared/reconcile-parse.util";
import type { ReconcileRecord } from "../shared/record";
import { AmountComparator } from "./ltx-amount-compare";
import type { ReconcileFieldRule } from "./ltx-config";

type ReconcileStatus = "PASS" | "FAIL" | "REVIEW";

export interface FieldCheckResult {
  fieldHeader: string;
  status: ReconcileStatus;
  remark: string;
}

export class FieldComparer {
  constructor(
    private readonly amountComparator: AmountComparator = new AmountComparator(),
  ) {}

  /** รหัสตัวเลขตัด Leading Zero ก่อนเทียบ แต่ข้อความทั่วไปไม่ถูกเปลี่ยนรูปแบบ */
  private normalizeIdForCompare(value: string): string {
    const normalized = value.trim().toLowerCase();

    if (!/^\d+$/.test(normalized)) {
      return normalized;
    }

    return normalized.replace(/^0+(?=\d)/, "");
  }

  /** ใช้รูปแบบ Remark กลางเพื่อให้ทุก Compare Mode แสดง Expected/Actual เหมือนกัน */
  private buildReviewRemark(
    reportCode: string,
    rule: ReconcileFieldRule,
    expected: string,
    actual: string,
  ): string {
    return formatCompareRemark(
      reportCode,
      rule.testDataField ?? "Test Data",
      expected,
      rule.reportField,
      actual,
    );
  }

  private toResult(
    fieldHeader: string,
    status: ReconcileStatus,
    remark = "",
  ): FieldCheckResult {
    return {
      fieldHeader,
      status,
      remark,
    };
  }

  /**
   * Fixed value ที่ผิดเป็น FAIL เพราะขัด Business Rule โดยตรง
   * แต่ถ้าค่าฝั่ง Report ถูกและข้อมูล Cross-check ว่าง ให้ REVIEW แทน
   */
  private checkFixedValue(
    reportCode: string,
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    actual: string,
  ): FieldCheckResult {
    const expected = rule.fixedValue ?? "";

    if (actual.trim().toLowerCase() !== expected.trim().toLowerCase()) {
      return this.toResult(
        rule.reportField,
        "FAIL",
        formatFixedValueRemark(
          reportCode,
          rule.reportField,
          actual,
          expected,
        ),
      );
    }

    if (rule.testDataField) {
      const crossCheckValue = testDataRecord.get(rule.testDataField);

      if (crossCheckValue.trim() === "") {
        return this.toResult(
          rule.reportField,
          "REVIEW",
          this.buildReviewRemark(reportCode, rule, crossCheckValue, actual),
        );
      }
    }

    return this.toResult(rule.reportField, "PASS");
  }

  /**
   * ตรวจ Report Date โดยตรงก่อน แล้วจึงใช้วันที่ตำแหน่ง 7–12
   * ของ Reference เป็น fallback; ไม่ตรงทั้งสองทางจึงเป็น REVIEW
   */
  private checkDateWithIdFallback(
    reportCode: string,
    rule: ReconcileFieldRule,
    reportRecord: ReconcileRecord,
    expected: string,
    actual: string,
  ): FieldCheckResult {
    const expectedDate = parseDate(expected);
    const fallbackReference = rule.fallbackReportField
      ? reportRecord.get(rule.fallbackReportField)
      : "";

    if (
      expectedDate &&
      isDateMatchWithArrangementFallback(
        expectedDate,
        actual,
        fallbackReference,
      )
    ) {
      return this.toResult(rule.reportField, "PASS");
    }

    return this.toResult(
      rule.reportField,
      "REVIEW",
      this.buildReviewRemark(reportCode, rule, expected, actual),
    );
  }

  /**
   * Amount ผ่านเมื่อผลต่างอยู่ใน Tolerance หากยอดหลักไม่ตรง
   * ต้องลองยอด Fallback ที่ Config กำหนดก่อนสรุปเป็น REVIEW
   */
  private checkAmountTolerance(
    reportCode: string,
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    expected: string,
    actual: string,
  ): FieldCheckResult {
    if (this.amountComparator.matches(expected, actual, rule.tolerance)) {
      return this.toResult(rule.reportField, "PASS");
    }

    if (rule.fallbackTestDataField) {
      const fallbackExpected = testDataRecord.get(rule.fallbackTestDataField);

      if (
        this.amountComparator.matches(
          fallbackExpected,
          actual,
          rule.tolerance,
        )
      ) {
        return this.toResult(rule.reportField, "PASS");
      }
    }

    return this.toResult(
      rule.reportField,
      "REVIEW",
      this.buildReviewRemark(reportCode, rule, expected, actual),
    );
  }

  /**
   * Exact comparison ต้องมีค่าทั้งสองฝั่งและตรงกันหลัง Normalize
   * ค่าว่างหรือค่าไม่ตรงเป็น REVIEW เพื่อให้ผู้ตรวจสอบเห็น Expected/Actual
   */
  private checkExact(
    reportCode: string,
    rule: ReconcileFieldRule,
    expected: string,
    actual: string,
  ): FieldCheckResult {
    const isMatch =
      expected.trim() !== "" &&
      actual.trim() !== "" &&
      this.normalizeIdForCompare(expected) ===
        this.normalizeIdForCompare(actual);

    return isMatch
      ? this.toResult(rule.reportField, "PASS")
      : this.toResult(
          rule.reportField,
          "REVIEW",
          this.buildReviewRemark(reportCode, rule, expected, actual),
        );
  }

  /** เลือกวิธีตรวจจาก compareMode ของ Rule โดยไม่ผูกชื่อ Field ไว้ใน Logic */
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

    return this.checkExact(reportCode, rule, expected, actual);
  }

  /** Rule ที่จำกัดเป็น DR ต้องไม่ถูกนำไปตรวจซ้ำกับแถวค่าธรรมเนียม FE */
  isApplicableToSuffix(rule: ReconcileFieldRule, suffix: string): boolean {
    return !rule.applicableSuffixes || rule.applicableSuffixes.includes(suffix);
  }
}