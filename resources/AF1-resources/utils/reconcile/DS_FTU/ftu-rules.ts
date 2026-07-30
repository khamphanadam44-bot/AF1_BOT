/**
 * Pure business-rule evaluator for DS_FTU.
 * It does not read or write Excel, which keeps the rules easy to unit test.
 */

import { ReconcileRecord } from "../shared/record";
import {
  FTU_BASE_CURRENCY,
  FTU_REPORT_CODE,
  FTU_REMARKS,
  FTU_TEST_FIELDS,
  FTU_USD_THRESHOLD,
} from "./ftu-config";
import { normalize, parseAmount } from "./ftu-parse.util";

export type FtuDirection = "BUY_FCY" | "SELL_FCY" | "NOT_FTU" | "INVALID";

export interface FtuExpectedCase {
  testCaseNo: string;
  transactionId: string;
  record: ReconcileRecord;
  direction: FtuDirection;
  expectedPresence: boolean;
  successRemark: string;
  unexpectedPresenceRemark?: string;
  usdEquivalentAmount?: number;
  validationError?: string;
}

export interface FtuExceptionEvaluator {
  /**
   * Return the confirmed R5 reason when a Test Data record is an exception.
   * Return undefined when the transaction is not an R5 exception.
   */
  getExceptionReason(record: ReconcileRecord): string | undefined;
}

export interface FtuRuleEvaluatorOptions {
  /**
   * Rate rule: foreign-currency amount multiplied by the rate
   * must equal its USD equivalent.
   */
  usdRateByCurrency?: Readonly<Record<string, number>>;
  exceptionEvaluator?: FtuExceptionEvaluator;
}

/** ใช้ฟังก์ชันแปลงค่าร่วมจาก ftu-parse.util.ts */

export class FtuRuleEvaluator {
  private readonly usdRateByCurrency: Readonly<Record<string, number>>;
  private readonly exceptionEvaluator?: FtuExceptionEvaluator;

  constructor(options: FtuRuleEvaluatorOptions = {}) {
    this.usdRateByCurrency = options.usdRateByCurrency ?? {};
    this.exceptionEvaluator = options.exceptionEvaluator;
  }

  getDirection(record: ReconcileRecord): FtuDirection {
    const fromCurrency = normalize(record.get(FTU_TEST_FIELDS.fromCurrency));
    const toCurrency = normalize(record.get(FTU_TEST_FIELDS.toCurrency));

    if (fromCurrency === "" || toCurrency === "") {
      return "INVALID";
    }

    if (
      fromCurrency === FTU_BASE_CURRENCY &&
      toCurrency === FTU_BASE_CURRENCY
    ) {
      return "INVALID";
    }

    if (
      fromCurrency === FTU_BASE_CURRENCY &&
      toCurrency !== FTU_BASE_CURRENCY
    ) {
      return "SELL_FCY";
    }

    if (
      fromCurrency !== FTU_BASE_CURRENCY &&
      toCurrency === FTU_BASE_CURRENCY
    ) {
      return "BUY_FCY";
    }

    return "NOT_FTU";
  }

  buildExpectedCases(records: ReconcileRecord[]): FtuExpectedCase[] {
    const transactionIds = new Set<string>();

    return records
      .map((record): FtuExpectedCase | null => {
        /**
         * Test No. ใช้สำหรับแสดงผล แต่ไม่ใช่ Matching Key หลัก
         *
         * ถ้า Test No. ว่าง:
         * - Script 2 จะ Highlight ช่อง Test No. เป็นสีแดง
         * - Script 3 ยังประเมินกฎ FTU และ Reconcile ต่อ
         * - Column Test Script No. ยังคงเป็นช่องว่าง
         * - Transaction ID ยังคงใช้เป็น Matching Key ตามเดิม
         */
        const identity = record.resolveIdentity(
          FTU_TEST_FIELDS.testNo,
          FTU_TEST_FIELDS.transactionId,
        );

        const testCaseNo = identity.displayValue;
        const transactionId = identity.matchingReference;

        /**
         * ถ้า Transaction ID ว่าง:
         * - Script 2 จะตรวจและ Highlight ช่องที่ว่าง
         * - Script 3 จะข้ามรายการนี้
         * - Script 3 จะไม่หยุด และจะทำรายการถัดไปต่อ
         */
        if (transactionId === "") {
          console.warn(
            `⚠️ [${FTU_REPORT_CODE}] Skip Test ${testCaseNo}: ` +
              `"${FTU_TEST_FIELDS.transactionId}" is empty.`,
          );

          return null;
        }

        const normalizedTransactionId = normalize(transactionId);

        if (transactionIds.has(normalizedTransactionId)) {
          throw new Error(
            `[${FTU_REPORT_CODE}] Duplicate Transaction ID in Test Data: ` +
              `"${transactionId}".`,
          );
        }

        transactionIds.add(normalizedTransactionId);

        const direction = this.getDirection(record);

        if (direction === "INVALID") {
          const fromCurrency = normalize(
            record.get(FTU_TEST_FIELDS.fromCurrency),
          );

          const toCurrency = normalize(
            record.get(FTU_TEST_FIELDS.toCurrency),
          );

          const validationError =
            fromCurrency === "" || toCurrency === ""
              ? `Data Validation Error: "${FTU_TEST_FIELDS.fromCurrency}" ` +
                `and "${FTU_TEST_FIELDS.toCurrency}" are required.`
              : `Data Validation Error: ${FTU_BASE_CURRENCY} -> ` +
                `${FTU_BASE_CURRENCY} is not an FX transaction.`;

          return {
            testCaseNo,
            transactionId,
            record,
            direction,
            expectedPresence: false,
            successRemark: validationError,
            validationError,
          };
        }

        if (direction === "NOT_FTU") {
          return {
            testCaseNo,
            transactionId,
            record,
            direction,
            expectedPresence: false,
            successRemark: FTU_REMARKS.noThbLegExpectedAbsence,
            unexpectedPresenceRemark:
              FTU_REMARKS.noThbLegUnexpectedPresence,
          };
        }

        /**
         * Test Data ที่ว่างไม่ใช่เหตุให้ Reconcile เป็น FAIL
         *
         * ตรวจ Threshold เฉพาะเมื่อ Settled Currency และ Settled Amount
         * มีค่าครบเท่านั้น ส่วน Remark ของค่าที่ว่างจะถูกสร้างใน
         * ftu-reconcile.ts หลังจากจับคู่กับข้อมูล AF1 แล้ว
         */
        const settledCurrencyText = normalize(
          record.get(FTU_TEST_FIELDS.settledCurrency),
        );

        const settledAmountText = String(
          record.get(FTU_TEST_FIELDS.settledAmount) ?? "",
        ).trim();

        const settledCurrency =
          settledCurrencyText === ""
            ? undefined
            : this.requireSettledCurrency(record, testCaseNo);

        const settledAmount =
          settledAmountText === ""
            ? undefined
            : this.requireSettledAmount(record, testCaseNo);

        const usdEquivalentAmount =
          settledCurrency !== undefined && settledAmount !== undefined
            ? this.toUsdEquivalent(
                settledAmount,
                settledCurrency,
                testCaseNo,
              )
            : undefined;

        if (
          usdEquivalentAmount !== undefined &&
          usdEquivalentAmount >= FTU_USD_THRESHOLD
        ) {
          return {
            testCaseNo,
            transactionId,
            record,
            direction,
            expectedPresence: false,
            successRemark:
              FTU_REMARKS.thresholdExpectedAbsence(usdEquivalentAmount),
            unexpectedPresenceRemark:
              FTU_REMARKS.thresholdUnexpectedPresence(usdEquivalentAmount),
            usdEquivalentAmount,
          };
        }

        const exceptionReason =
          this.exceptionEvaluator?.getExceptionReason(record);

        if (exceptionReason) {
          return {
            testCaseNo,
            transactionId,
            record,
            direction,
            expectedPresence: false,
            successRemark:
              FTU_REMARKS.ftxExceptionExpectedAbsence(exceptionReason),
            unexpectedPresenceRemark:
              FTU_REMARKS.ftxExceptionUnexpectedPresence(exceptionReason),
            usdEquivalentAmount,
          };
        }

        return {
          testCaseNo,
          transactionId,
          record,
          direction,
          expectedPresence: true,
          successRemark:
            direction === "BUY_FCY"
              ? FTU_REMARKS.buyForeignCurrency
              : FTU_REMARKS.sellForeignCurrency,
          usdEquivalentAmount,
        };
      })
      .filter(
        (expectedCase): expectedCase is FtuExpectedCase =>
          expectedCase !== null,
      );
  }

  private requireSettledCurrency(
    record: ReconcileRecord,
    testCaseNo: string,
  ): string {
    const currency = normalize(
      record.get(FTU_TEST_FIELDS.settledCurrency),
    );

    if (currency === "" || currency === FTU_BASE_CURRENCY) {
      throw new Error(
        `[${FTU_REPORT_CODE}] Test ${testCaseNo}: ` +
          `"${FTU_TEST_FIELDS.settledCurrency}" must be a foreign currency, ` +
          `but found "${currency}".`,
      );
    }

    return currency;
  }

  private requireSettledAmount(
    record: ReconcileRecord,
    testCaseNo: string,
  ): number {
    const amount = parseAmount(
      record.get(FTU_TEST_FIELDS.settledAmount),
    );

    if (amount === null || amount <= 0) {
      throw new Error(
        `[${FTU_REPORT_CODE}] Test ${testCaseNo}: ` +
          `"${FTU_TEST_FIELDS.settledAmount}" must be greater than zero.`,
      );
    }

    return amount;
  }

  private toUsdEquivalent(
    amount: number,
    currency: string,
    testCaseNo: string,
  ): number {
    if (currency === "USD") {
      return amount;
    }

    const configuredRate = this.usdRateByCurrency[currency];
    const environmentRate = readCurrencyToUsdRate(currency);
    const rate = configuredRate ?? environmentRate;

    if (rate === undefined || !Number.isFinite(rate) || rate <= 0) {
      const environmentVariable = `FTU_${currency}_TO_USD_RATE`;

      throw new Error(
        `[${FTU_REPORT_CODE}] Test ${testCaseNo}: ${currency} to USD rate ` +
          `is required ` +
          `to evaluate the 50,000 USD equivalent threshold. ` +
          `Set ${environmentVariable} or inject usdRateByCurrency.`,
      );
    }

    return amount * rate;
  }
}

export const readCurrencyToUsdRate = (
  currency: string,
): number | undefined => {
  const normalizedCurrency = normalize(currency);
  const environmentVariable =
    `FTU_${normalizedCurrency}_TO_USD_RATE`;

  const value = process.env[environmentVariable]?.trim();

  if (!value) {
    return undefined;
  }

  const rate = Number(value);

  if (!Number.isFinite(rate) || rate <= 0) {
    throw new Error(
      `${environmentVariable} must be a positive number, ` +
        `but found "${value}".`,
    );
  }

  return rate;
};