/**
 * ReconcileMatcher
 * ------------------------------------------------------------------
 * Report-agnostic: จัดกลุ่ม record ตาม group key + หา record ที่ match กันแบบไม่กำกวม
 * report อื่น (PTX, FTX, ...) reuse class นี้ได้ตรงๆ ผ่าน Config ของ Report
 *
 * หมายเหตุ: findUniqueAmountMatch / groupReportRecordsByKey / buildReportGroupKey /
 * buildTestDataGroupKey เก็บไว้เหมือนเดิมทั้งหมด (ไม่ได้ใช้ในเส้นทางหลักของ DS_LTX
 * อีกต่อไปหลังเปลี่ยนมาใช้ findByExactReference แต่ยังคงไว้เผื่อ report อื่นในอนาคต
 * ที่อาจไม่มี Transaction ID ให้ lookup ตรง ๆ แบบ DS_LTX)
 *
 * TODO: AmountComparator ยังอยู่ที่ DS_LTX/ltx-amount-compare.ts
 * แม้ Logic จะสามารถใช้ร่วมกับ Report อื่นได้ แต่ยังไม่ได้ย้ายมาไว้ใน shared/
 *
 * ปัจจุบันไฟล์นี้จึงยังมี Dependency ไปยัง DS_LTX อยู่ 1 จุด
 * หาก Report อื่นต้องใช้ AmountComparator ร่วมกัน
 * ควรพิจารณาย้าย Logic นี้มาไว้ใน shared/ ในอนาคต
 * ------------------------------------------------------------------
 */
import { normalizeValue } from "../../validators/shared/excel-cell.util";
import type { ReconcileGroupKeyFields } from "../DS_LTX/ltx-config";
import { AmountComparator } from "../DS_LTX/ltx-amount-compare";
import { ReconcileRecord } from "./record";

export interface AmountMatchResult {
  record: ReconcileRecord | undefined;
  /** true = เจอมากกว่า 1 แถวที่ยอดเงินตรงกัน (ไม่รู้จะเลือกอันไหน ต้องตรวจสอบ manual) */
  isAmbiguous: boolean;
  matchedBy: "primary" | "fallback" | "none";
}

export class ReconcileMatcher {
  constructor(
    private readonly amountComparator: AmountComparator = new AmountComparator(),
  ) {}

  private normalizeKeyPart(value: string): string {
    return normalizeValue(value).toLowerCase();
  }

  buildReportGroupKey(
    record: ReconcileRecord,
    groupKeyFields: ReconcileGroupKeyFields,
  ): string {
    const account = record.get(groupKeyFields.reportAccountField);
    const currency = record.get(groupKeyFields.reportCurrencyField);
    return `${this.normalizeKeyPart(account)}||${this.normalizeKeyPart(currency)}`;
  }

  buildTestDataGroupKey(
    record: ReconcileRecord,
    groupKeyFields: ReconcileGroupKeyFields,
  ): string {
    const account = record.get(groupKeyFields.testDataAccountField);
    const currency = record.get(groupKeyFields.testDataCurrencyField);
    return `${this.normalizeKeyPart(account)}||${this.normalizeKeyPart(currency)}`;
  }

  /** จัดกลุ่ม Report record ทั้งหมดตาม group key — ใช้ lookup ตอนหา candidate ของแต่ละ Test Case */
  groupReportRecordsByKey(
    reportRecords: ReconcileRecord[],
    groupKeyFields: ReconcileGroupKeyFields,
  ): Map<string, ReconcileRecord[]> {
    const groups = new Map<string, ReconcileRecord[]>();

    reportRecords.forEach((record) => {
      const key = this.buildReportGroupKey(record, groupKeyFields);
      if (key === "||") {
        return; // Account และ Currency ว่างทั้งคู่ (เช่นแถว "Balance This Period") -> ไม่เอาเข้ากลุ่ม
      }
      const existing = groups.get(key) ?? [];
      existing.push(record);
      groups.set(key, existing);
    });

    return groups;
  }

  /**
   * หา record เดียวจาก candidate ที่ยอดเงินตรงกับ expected แบบไม่กำกวม
   * รองรับ fallback field (เช่น Transfer Amount ไม่ตรง -> ลองเทียบกับ Debit Amount)
   */
  findUniqueAmountMatch(
    candidates: ReconcileRecord[],
    reportAmountField: string,
    expectedAmount: string,
    expectedAmountFallback?: string,
    tolerance?: number,
  ): AmountMatchResult {
    const matchedByPrimary = candidates.filter(
      (candidate) =>
        this.amountComparator.compare(
          expectedAmount,
          candidate.get(reportAmountField),
          tolerance,
        ).isMatch,
    );

    if (matchedByPrimary.length === 1) {
      return {
        record: matchedByPrimary[0],
        isAmbiguous: false,
        matchedBy: "primary",
      };
    }
    if (matchedByPrimary.length > 1) {
      return { record: undefined, isAmbiguous: true, matchedBy: "primary" };
    }
    if (!expectedAmountFallback) {
      return { record: undefined, isAmbiguous: false, matchedBy: "none" };
    }

    const matchedByFallback = candidates.filter(
      (candidate) =>
        this.amountComparator.compare(
          expectedAmountFallback,
          candidate.get(reportAmountField),
          tolerance,
        ).isMatch,
    );

    if (matchedByFallback.length === 1) {
      return {
        record: matchedByFallback[0],
        isAmbiguous: false,
        matchedBy: "fallback",
      };
    }
    return {
      record: undefined,
      isAmbiguous: matchedByFallback.length > 1,
      matchedBy: "none",
    };
  }

  /**
   * หา record ที่ Reference Transaction Number ตรงกับ expectedReference แบบเป๊ะ ๆ
   * (เทียบแบบ case-insensitive + ตัดช่องว่างหัวท้าย) — ใช้แทนการเทียบยอดเงินแบบเดิม
   */
  findByExactReference(
    records: ReconcileRecord[],
    referenceField: string,
    expectedReference: string,
  ): ReconcileRecord | undefined {
    const target = expectedReference.trim().toUpperCase();
    if (target === "") {
      return undefined;
    }
    return records.find(
      (record) => record.get(referenceField).trim().toUpperCase() === target,
    );
  }

  /**
   * หา record คู่กันโดยสลับ suffix ท้าย Reference Transaction Number ของ record ต้นทาง
   * เช่น DR "...082000DR" -> หา FE "...082000FE" ตรงๆ (แม่นยำกว่าเทียบยอดเงินอย่างเดียว)
   */
  findRecordBySuffixSwap(
    candidates: ReconcileRecord[],
    referenceField: string,
    sourceRecord: ReconcileRecord,
    sourceSuffix: string,
    targetSuffix: string,
  ): ReconcileRecord | undefined {
    const sourceReference = sourceRecord.get(referenceField).toUpperCase();

    if (!sourceReference.endsWith(sourceSuffix.toUpperCase())) {
      return undefined;
    }

    const baseReference = sourceReference.slice(
      0,
      sourceReference.length - sourceSuffix.length,
    );
    const targetReference = `${baseReference}${targetSuffix.toUpperCase()}`;

    return candidates.find(
      (candidate) =>
        candidate.get(referenceField).toUpperCase() === targetReference,
    );
  }
}
