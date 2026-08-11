/**
 * FieldRuleValidator 
 * ------------------------------------------------------------------
 * Template Method เดียว: FieldRuleValidator.validate()
 * ทำหน้าที่ filter + map เหมือนกันทั้งคู่ ส่วน subclass override แค่ selectRules()
 * ------------------------------------------------------------------
 */
import { ReconcileFieldRule } from "./ltx-config";
import { FieldComparer } from "./ltx-field-compare";
import { ReconcileRecord } from "../shared/record";
import { FieldCheckResult } from "./ltx-types";

export abstract class FieldRuleValidator {
  protected readonly comparer: FieldComparer;

  constructor(comparer: FieldComparer = new FieldComparer()) {
    this.comparer = comparer;
  }

  /** เลือก rule ที่ validator ตัวนี้รับผิดชอบ (override โดย subclass) */
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

/** Core Field (isRequiredForAllCases = true) — ต้องตรวจสอบทุก Test Case ไม่มีเงื่อนไขข้าม */
export class CoreFieldValidator extends FieldRuleValidator {
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

/** Conditional Field (isRequiredForAllCases = false) — ตรวจเฉพาะเมื่อเข้าเงื่อนไข */
export class ConditionalFieldValidator extends FieldRuleValidator {
  protected selectRules(
    fieldRules: ReconcileFieldRule[],
    suffix: string,
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): ReconcileFieldRule[] {
    return fieldRules
      .filter(
        (rule) =>
          !rule.isRequiredForAllCases &&
          this.comparer.isApplicableToSuffix(rule, suffix),
      )
      .filter((rule) => this.shouldCheck(rule, testDataRecord, reportRecord));
  }

  /**
   * เงื่อนไขข้าม (นอกเหนือจาก suffix ที่เช็คแยกไว้แล้ว):
   * 1) มี onlyWhenReportFieldHasValue กำหนดไว้ แต่ field นั้นฝั่ง AF1 Report ว่าง -> ข้าม
   *    (เช่น Inflow Transaction Purpose: ตรวจเฉพาะแถวที่ direction เป็น Inflow จริง ๆ
   *    สังเกตจาก field นี้ฝั่ง Report เองมีค่าหรือไม่ — ไม่ได้ดูฝั่ง Test Data)
   * 2) field เงื่อนไขฝั่ง Test Data (skipWhenTestDataFieldEmpty หรือ testDataField ของตัวเอง) ว่าง
   * 3) ค่า Test Data ของ field ตัวเองก็ว่างด้วย
   */
  private shouldCheck(
    rule: ReconcileFieldRule,
    testDataRecord: ReconcileRecord,
    reportRecord: ReconcileRecord,
  ): boolean {
    if (rule.onlyWhenReportFieldHasValue) {
      const reportConditionValue = reportRecord.get(
        rule.onlyWhenReportFieldHasValue,
      );
      if (reportConditionValue.trim() === "") {
        return false;
      }
    }

    const conditionField =
      rule.skipWhenTestDataFieldEmpty ?? rule.testDataField;
    const conditionValue = conditionField
      ? testDataRecord.get(conditionField)
      : "";

    if (conditionField && conditionValue.trim() === "") {
      return false;
    }

    if (rule.testDataField) {
      const ownValue = testDataRecord.get(rule.testDataField);
      if (ownValue.trim() === "") {
        return false;
      }
    }

    return true;
  }
}

/** รวมผล Core + Conditional ของ 1 คู่ — ใช้แทน checkFieldRules ใน reconcile-validator.ts เดิม */
export class FieldRuleValidatorSet {
  constructor(
    private readonly coreValidator: CoreFieldValidator = new CoreFieldValidator(),
    private readonly conditionalValidator: ConditionalFieldValidator = new ConditionalFieldValidator(),
  ) {}

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
