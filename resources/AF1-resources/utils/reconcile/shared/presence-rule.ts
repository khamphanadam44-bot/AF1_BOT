/**
 * presence-rule.ts
 * ------------------------------------------------------------------
 * Rule กลางสำหรับตัดสินว่า Test Data Record ต้องมีหรือห้ามมีใน AF1 Report
 * ไม่ผูกกับ Test Case No., วันที่, Transaction ID หรือชื่อ Report
 * ------------------------------------------------------------------
 */

export type ReportPresenceExpectation = "MUST_EXIST" | "MUST_NOT_EXIST";

export type TestDataConditionOperator = "EQUALS" | "NOT_EQUALS" | "NOT_EMPTY";

export interface TestDataCondition {
  readonly field: string;
  readonly operator: TestDataConditionOperator;
  readonly value?: string;
}

export interface ReportPresenceRule {
  readonly ruleCode: string;
  readonly expectation: ReportPresenceExpectation;
  readonly allConditions: readonly TestDataCondition[];
  readonly passRemark: string;
  readonly failRemark: string;
}

export interface ReportPresenceDecision {
  readonly expectation: ReportPresenceExpectation;
  readonly matchedRuleCode?: string;
  readonly passRemark: string;
  readonly failRemark: string;
}

/** Domain Object ที่ Evaluator ต้องการใช้เพียง Method get() */
export interface PresenceRuleRecord {
  get(headerName: string): string;
}

/**
 * ดึงรายชื่อ Test Data Header ที่ Presence Rules ต้องใช้
 *
 * ใช้สำหรับตรวจโครงสร้างไฟล์ใน Script 3 เท่านั้น
 */
export const getPresenceRuleFields = (
  rules: readonly ReportPresenceRule[],
): string[] => {
  const fields = rules.flatMap((rule) =>
    rule.allConditions.map((condition) => condition.field),
  );

  /**
   * ตัด Header ซ้ำ
   *
   * ตัวอย่าง To Currency มีทั้ง:
   * - NOT_EMPTY
   * - NOT_EQUALS
   *
   * แต่ควรตรวจชื่อ Header เพียงครั้งเดียว
   */
  return [...new Set(fields)];
};

export class ReportPresenceRuleEvaluator {
  private normalize(value: string): string {
    return String(value ?? "")
      .trim()
      .replace(/\s+/g, " ")
      .toUpperCase();
  }

  private matchesCondition(
    record: PresenceRuleRecord,
    condition: TestDataCondition,
  ): boolean {
    const actualValue = this.normalize(record.get(condition.field));
    const expectedValue = this.normalize(condition.value ?? "");

    switch (condition.operator) {
      case "EQUALS":
        return actualValue === expectedValue;

      case "NOT_EQUALS":
        return actualValue !== expectedValue;

      case "NOT_EMPTY":
        return actualValue !== "";

      default: {
        const unsupportedOperator: never = condition.operator;
        throw new Error(
          `Unsupported Presence Rule operator: ${unsupportedOperator}`,
        );
      }
    }
  }

  private matchesRule(
    record: PresenceRuleRecord,
    rule: ReportPresenceRule,
  ): boolean {
    return rule.allConditions.every((condition) =>
      this.matchesCondition(record, condition),
    );
  }

  evaluate(
    record: PresenceRuleRecord,
    rules: readonly ReportPresenceRule[] = [],
  ): ReportPresenceDecision {
    const matchedRule = rules.find((rule) => this.matchesRule(record, rule));

    if (!matchedRule) {
      return {
        expectation: "MUST_EXIST",
        passRemark: "",
        failRemark: "",
      };
    }

    return {
      expectation: matchedRule.expectation,
      matchedRuleCode: matchedRule.ruleCode,
      passRemark: matchedRule.passRemark,
      failRemark: matchedRule.failRemark,
    };
  }
}
/**
 * Resident โอน THB ไป FCD/สกุลเงินต่างประเทศ
 * รายการต้องไม่มีใน Report
 */
export const RESIDENT_THB_TO_FCD_MUST_NOT_EXIST: ReportPresenceRule = {
  ruleCode: "RESIDENT_THB_TO_FCD_MUST_NOT_EXIST",
  expectation: "MUST_NOT_EXIST",

  allConditions: [
    {
      field: "From Customer (Resident/Non Resident)",
      operator: "EQUALS",
      value: "Resident",
    },
    {
      field: "From Currency (CCY)",
      operator: "EQUALS",
      value: "THB",
    },
    {
      field: "To Currency (CCY)",
      operator: "NOT_EMPTY",
    },
    {
      field: "To Currency (CCY)",
      operator: "NOT_EQUALS",
      value: "THB",
    },
  ],

  passRemark:
    "กรณีลูกค้า Resident โอน (From) THB บาทออกจากบัญชีไปยัง FCD " +
    "หรือสกุลเงินต่างประเทศ รายการต้องไม่แสดงใน {REPORT_CODE}",

  failRemark:
    "พบรายการใน {REPORT_CODE} ทั้งที่ลูกค้า Resident โอน (From) THB บาท " +
    "ออกจากบัญชีไปยัง FCD หรือสกุลเงินต่างประเทศ " +
    "ซึ่งตาม Requirement รายการต้องไม่แสดง",
};