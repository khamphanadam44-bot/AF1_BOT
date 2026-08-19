/**
 * ltx-field-validator
 */
import type { ReconcileRecord } from "../shared/record";
import type { ReconcileFieldRule } from "./ltx-config";
import {
  FieldComparer,
  type FieldCheckResult,
} from "./ltx-field-compare";

/**
 * ใช้ Template Method เพื่อรวมขั้นตอนเลือก Rule และเปรียบเทียบ Field ไว้จุดเดียว
 * Subclass กำหนดเฉพาะวิธีเลือก Core หรือ Conditional Rule
 */

abstract class FieldRuleValidator {
  constructor(protected comparer = new FieldComparer()) {}

  protected abstract selectRules(
    fieldRules: ReconcileFieldRule[],
    suffix: string,
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): ReconcileFieldRule[];

  validate(
    reportCode: string,
    fieldRules: ReconcileFieldRule[],
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
    suffix: string,
  ): FieldCheckResult[] {
    return this.selectRules(
      fieldRules,
      suffix,
      testDataRecord,
      reportRecord,
    ).map((rule) =>
      this.comparer.check(reportCode, rule, testDataRecord, reportRecord),
    );
  }
}

class CoreFieldValidator extends FieldRuleValidator {
  /** Core Rule ต้องตรวจทุก Case แต่ยังคงเคารพข้อจำกัดของ Suffix */
  protected selectRules(
    fieldRules: ReconcileFieldRule[],
    suffix: string,
  ): ReconcileFieldRule[] {
    return fieldRules.filter(
      (rule) =>
        rule.isRequiredForAllCases &&
        this.comparer.isApplicableToSuffix(rule, suffix),
    );
  }
}

class ConditionalFieldValidator extends FieldRuleValidator {
  /** Conditional Rule ตรวจเฉพาะเมื่อ Suffix และข้อมูลเข้าเงื่อนไข */
  protected selectRules(
    fieldRules: ReconcileFieldRule[],
    suffix: string,
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): ReconcileFieldRule[] {
    return fieldRules.filter(
      (rule) =>
        !rule.isRequiredForAllCases &&
        this.comparer.isApplicableToSuffix(rule, suffix) &&
        this.shouldCheck(rule, testDataRecord, reportRecord),
    );
  }

  /** ข้าม Rule เมื่อ Field ที่ใช้เปิดเงื่อนไขไม่มีค่าฝั่งที่กำหนด */
  private shouldCheck(
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): boolean {
    if (
      rule.onlyWhenReportFieldHasValue &&
      reportRecord.get(rule.onlyWhenReportFieldHasValue).trim() === ""
    ) {
      return false;
    }

    const conditionField =
      rule.skipWhenTestDataFieldEmpty ?? rule.testDataField;

    if (
      conditionField &&
      testDataRecord.get(conditionField).trim() === ""
    ) {
      return false;
    }

    return (
      !rule.testDataField ||
      testDataRecord.get(rule.testDataField).trim() !== ""
    );
  }
}

export class FieldRuleValidatorSet {
  constructor(
    private coreValidator = new CoreFieldValidator(),
    private conditionalValidator = new ConditionalFieldValidator(),
  ) {}

  /** คงลำดับ Core ก่อน Conditional เพื่อให้ผลลัพธ์อ่านได้สม่ำเสมอ */
  validateAll(
    reportCode: string,
    fieldRules: ReconcileFieldRule[],
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
    suffix: string,
  ): FieldCheckResult[] {
    return [
      ...this.coreValidator.validate(
        reportCode,
        fieldRules,
        testDataRecord,
        reportRecord,
        suffix,
      ),
      ...this.conditionalValidator.validate(
        reportCode,
        fieldRules,
        testDataRecord,
        reportRecord,
        suffix,
      ),
    ];
  }
}