/**
 * Pure business-rule evaluator for DS_FTU.
 * It does not read or write Excel, which keeps the rules easy to unit test.
 */

import { ReconcileRecord } from "../shared/record";
import {
  FTU_REPORT_CODE,
  FTU_REMARKS,
  FTU_TEST_DATA_FIELDS,
  FTU_THB_CURRENCY_CODE,
  FTU_USD_THRESHOLD,
} from "./ftu-config";
import { normalize, parseAmount } from "./ftu-parse.util";

/**
 * ใช้ชื่อจาก ftu-config.ts เวอร์ชันปัจจุบัน
 */
const FTU_BASE_CURRENCY = FTU_THB_CURRENCY_CODE;
const FTU_TEST_FIELDS = FTU_TEST_DATA_FIELDS;

export type FtuDirection =
  | "BUY_FCY"
  | "SELL_FCY"
  | "NOT_FTU"
  | "INVALID";

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
  getExceptionReason(
    record: ReconcileRecord,
  ): string | undefined;
}

export interface FtuRuleEvaluatorOptions {
  /**
   * Rate rule: foreign-currency amount multiplied by the rate
   * must equal its USD equivalent.
   */
  usdRateByCurrency?: Readonly<Record<string, number>>;
  exceptionEvaluator?: FtuExceptionEvaluator;
}

/**
 * Remark กรณีเข้าเงื่อนไข FTX Exception
 *
 * ประกาศไว้ในไฟล์นี้ เพราะ ftu-config.ts เวอร์ชันปัจจุบัน
 * ยังไม่ได้ Export Remark สองรายการนี้
 */
const getFtxExceptionExpectedAbsenceRemark = (
  reason: string,
): string =>
  `เข้าเงื่อนไข FTX Exception (${reason}) — ` +
  `ไม่ควรพบใน DS_FTU`;

const getFtxExceptionUnexpectedPresenceRemark = (
  reason: string,
): string =>
  `เข้าเงื่อนไข FTX Exception (${reason}) ` +
  `แต่พบใน DS_FTU โดยไม่ควรพบ`;

export class FtuRuleEvaluator {
  private readonly usdRateByCurrency: Readonly<
    Record<string, number>
  >;

  private readonly exceptionEvaluator?: FtuExceptionEvaluator;

  constructor(
    options: FtuRuleEvaluatorOptions = {},
  ) {
    this.usdRateByCurrency =
      options.usdRateByCurrency ?? {};

    this.exceptionEvaluator =
      options.exceptionEvaluator;
  }

  /**
   * ตรวจสอบทิศทางของธุรกรรม
   */
  getDirection(
    record: ReconcileRecord,
  ): FtuDirection {
    const fromCurrency = normalize(
      record.get(FTU_TEST_FIELDS.fromCurrency),
    );

    const toCurrency = normalize(
      record.get(FTU_TEST_FIELDS.toCurrency),
    );

    /**
     * From Currency หรือ To Currency ว่าง
     */
    if (
      fromCurrency === "" ||
      toCurrency === ""
    ) {
      return "INVALID";
    }

    /**
     * THB ไป THB ไม่ใช่ธุรกรรม FX
     */
    if (
      fromCurrency === FTU_BASE_CURRENCY &&
      toCurrency === FTU_BASE_CURRENCY
    ) {
      return "INVALID";
    }

    /**
     * THB ไปสกุลเงินต่างประเทศ
     * หมายถึงขายเงินตราต่างประเทศ
     */
    if (
      fromCurrency === FTU_BASE_CURRENCY &&
      toCurrency !== FTU_BASE_CURRENCY
    ) {
      return "SELL_FCY";
    }

    /**
     * สกุลเงินต่างประเทศมาเป็น THB
     * หมายถึงซื้อเงินตราต่างประเทศ
     */
    if (
      fromCurrency !== FTU_BASE_CURRENCY &&
      toCurrency === FTU_BASE_CURRENCY
    ) {
      return "BUY_FCY";
    }

    /**
     * ไม่มีขา THB
     */
    return "NOT_FTU";
  }

  /**
   * สร้าง Expected Case จาก Test Data
   */
  buildExpectedCases(
    records: ReconcileRecord[],
  ): FtuExpectedCase[] {
    const transactionIds =
      new Set<string>();

    return records
      .map(
        (
          record,
        ): FtuExpectedCase | null => {
          /**
           * Test No. ใช้สำหรับแสดงผล
           * แต่ไม่ใช่ Matching Key หลัก
           *
           * ถ้า Test No. ว่าง:
           * - Script 2 Highlight ช่อง Test No. เป็นสีแดง
           * - Script 3 ยัง Reconcile ต่อ
           * - Transaction ID ยังคงเป็น Matching Key
           */
          const identity =
            record.resolveIdentity(
              FTU_TEST_FIELDS.testNo,
              FTU_TEST_FIELDS.transactionId,
            );

          const testCaseNo =
            identity.displayValue;

          const transactionId =
            identity.matchingReference;

          /**
           * ถ้า Transaction ID ว่าง:
           * - แสดง Warning
           * - ข้ามรายการนี้
           * - ไม่หยุด Script 3
           * - ทำรายการถัดไปต่อ
           */
          if (transactionId === "") {
            console.warn(
              `⚠️ [${FTU_REPORT_CODE}] ` +
                `Skip Test ${testCaseNo}: ` +
                `"${FTU_TEST_FIELDS.transactionId}" ` +
                `is empty.`,
            );

            return null;
          }

          const normalizedTransactionId =
            normalize(transactionId);

          /**
           * Transaction ID ซ้ำยังถือว่าเป็น Error
           * เพราะไม่สามารถระบุรายการที่จะจับคู่ได้ถูกต้อง
           */
          if (
            transactionIds.has(
              normalizedTransactionId,
            )
          ) {
            throw new Error(
              `[${FTU_REPORT_CODE}] ` +
                `Duplicate Transaction ID in Test Data: ` +
                `"${transactionId}".`,
            );
          }

          transactionIds.add(
            normalizedTransactionId,
          );

          const direction =
            this.getDirection(record);

          /**
           * Currency ไม่ครบ หรือ THB ไป THB
           */
          if (direction === "INVALID") {
            const fromCurrency =
              normalize(
                record.get(
                  FTU_TEST_FIELDS.fromCurrency,
                ),
              );

            const toCurrency =
              normalize(
                record.get(
                  FTU_TEST_FIELDS.toCurrency,
                ),
              );

            const validationError =
              fromCurrency === "" ||
              toCurrency === ""
                ? `Data Validation Error: ` +
                  `"${FTU_TEST_FIELDS.fromCurrency}" ` +
                  `and ` +
                  `"${FTU_TEST_FIELDS.toCurrency}" ` +
                  `are required.`
                : `Data Validation Error: ` +
                  `${FTU_BASE_CURRENCY} -> ` +
                  `${FTU_BASE_CURRENCY} ` +
                  `is not an FX transaction.`;

            return {
              testCaseNo,
              transactionId,
              record,
              direction,
              expectedPresence: false,
              successRemark:
                validationError,
              validationError,
            };
          }

          /**
           * ไม่มีขา THB
           */
          if (direction === "NOT_FTU") {
            return {
              testCaseNo,
              transactionId,
              record,
              direction,
              expectedPresence: false,
              successRemark:
                FTU_REMARKS
                  .noThbLegExpectedAbsence,
              unexpectedPresenceRemark:
                FTU_REMARKS
                  .noThbLegUnexpectedPresence,
            };
          }

          /**
           * ถ้า Settled Currency หรือ Settled Amount ว่าง
           * จะไม่หยุด Script และไม่ตรวจ Threshold
           *
           * Script 2 จะเป็นผู้ Highlight ช่องที่ว่าง
           */
          const settledCurrencyText =
            normalize(
              record.get(
                FTU_TEST_FIELDS.settledCurrency,
              ),
            );

          const settledAmountText =
            String(
              record.get(
                FTU_TEST_FIELDS.settledAmount,
              ) ?? "",
            ).trim();

          const settledCurrency =
            settledCurrencyText === ""
              ? undefined
              : this.requireSettledCurrency(
                  record,
                  testCaseNo,
                );

          const settledAmount =
            settledAmountText === ""
              ? undefined
              : this.requireSettledAmount(
                  record,
                  testCaseNo,
                );

          const usdEquivalentAmount =
            settledCurrency !== undefined &&
            settledAmount !== undefined
              ? this.toUsdEquivalent(
                  settledAmount,
                  settledCurrency,
                  testCaseNo,
                )
              : undefined;

          /**
           * ยอดเทียบเท่า USD ตั้งแต่ 50,000 ขึ้นไป
           * ไม่ควรพบใน DS_FTU
           */
          if (
            usdEquivalentAmount !==
              undefined &&
            usdEquivalentAmount >=
              FTU_USD_THRESHOLD
          ) {
            return {
              testCaseNo,
              transactionId,
              record,
              direction,
              expectedPresence: false,
              successRemark:
                FTU_REMARKS
                  .thresholdExpectedAbsence(
                    usdEquivalentAmount,
                  ),
              unexpectedPresenceRemark:
                FTU_REMARKS
                  .thresholdUnexpectedPresence(
                    usdEquivalentAmount,
                  ),
              usdEquivalentAmount,
            };
          }

          /**
           * ตรวจเงื่อนไข FTX Exception
           */
          const exceptionReason =
            this.exceptionEvaluator
              ?.getExceptionReason(record);

          if (exceptionReason) {
            return {
              testCaseNo,
              transactionId,
              record,
              direction,
              expectedPresence: false,
              successRemark:
                getFtxExceptionExpectedAbsenceRemark(
                  exceptionReason,
                ),
              unexpectedPresenceRemark:
                getFtxExceptionUnexpectedPresenceRemark(
                  exceptionReason,
                ),
              usdEquivalentAmount,
            };
          }

          /**
           * รายการ FTU ปกติ
           */
          return {
            testCaseNo,
            transactionId,
            record,
            direction,
            expectedPresence: true,
            successRemark:
              direction === "BUY_FCY"
                ? FTU_REMARKS
                    .buyForeignCurrency
                : FTU_REMARKS
                    .sellForeignCurrency,
            usdEquivalentAmount,
          };
        },
      )
      .filter(
        (
          expectedCase,
        ): expectedCase is FtuExpectedCase =>
          expectedCase !== null,
      );
  }

  /**
   * ตรวจสอบ Settled Currency
   */
  private requireSettledCurrency(
    record: ReconcileRecord,
    testCaseNo: string,
  ): string {
    const currency = normalize(
      record.get(
        FTU_TEST_FIELDS.settledCurrency,
      ),
    );

    if (
      currency === "" ||
      currency === FTU_BASE_CURRENCY
    ) {
      throw new Error(
        `[${FTU_REPORT_CODE}] ` +
          `Test ${testCaseNo}: ` +
          `"${FTU_TEST_FIELDS.settledCurrency}" ` +
          `must be a foreign currency, ` +
          `but found "${currency}".`,
      );
    }

    return currency;
  }

  /**
   * ตรวจสอบ Settled Amount
   */
  private requireSettledAmount(
    record: ReconcileRecord,
    testCaseNo: string,
  ): number {
    const amount = parseAmount(
      record.get(
        FTU_TEST_FIELDS.settledAmount,
      ),
    );

    if (
      amount === null ||
      amount <= 0
    ) {
      throw new Error(
        `[${FTU_REPORT_CODE}] ` +
          `Test ${testCaseNo}: ` +
          `"${FTU_TEST_FIELDS.settledAmount}" ` +
          `must be greater than zero.`,
      );
    }

    return amount;
  }

  /**
   * แปลงยอดสกุลเงินต่างประเทศเป็น USD
   */
  private toUsdEquivalent(
    amount: number,
    currency: string,
    testCaseNo: string,
  ): number {
    if (currency === "USD") {
      return amount;
    }

    const configuredRate =
      this.usdRateByCurrency[currency];

    const environmentRate =
      readCurrencyToUsdRate(currency);

    const rate =
      configuredRate ??
      environmentRate;

    if (
      rate === undefined ||
      !Number.isFinite(rate) ||
      rate <= 0
    ) {
      const environmentVariable =
        `FTU_${currency}_TO_USD_RATE`;

      throw new Error(
        `[${FTU_REPORT_CODE}] ` +
          `Test ${testCaseNo}: ` +
          `${currency} to USD rate is required ` +
          `to evaluate the 50,000 USD equivalent threshold. ` +
          `Set ${environmentVariable} ` +
          `or inject usdRateByCurrency.`,
      );
    }

    return amount * rate;
  }
}

/**
 * อ่านอัตราแลกเปลี่ยนจาก Environment Variable
 *
 * ตัวอย่าง:
 * FTU_EUR_TO_USD_RATE=1.08
 * FTU_JPY_TO_USD_RATE=0.0067
 */
export const readCurrencyToUsdRate = (
  currency: string,
): number | undefined => {
  const normalizedCurrency =
    normalize(currency);

  const environmentVariable =
    `FTU_${normalizedCurrency}_TO_USD_RATE`;

  const value =
    process.env[
      environmentVariable
    ]?.trim();

  if (!value) {
    return undefined;
  }

  const rate = Number(value);

  if (
    !Number.isFinite(rate) ||
    rate <= 0
  ) {
    throw new Error(
      `${environmentVariable} ` +
        `must be a positive number, ` +
        `but found "${value}".`,
    );
  }

  return rate;
};