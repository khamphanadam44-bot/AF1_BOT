/**
 * fxu-rules.ts
 * ------------------------------------------------------------------
 * Business Rule ของ DF_FXU สำหรับ Script 3
 *
 * หน้าที่:
 * 1. Normalize ค่าที่ใช้ใน Business Rule
 * 2. แปลง Settled Amount เป็นตัวเลข
 * 3. ตรวจข้อมูลบังคับที่ใช้ใน Business Rule
 * 4. ตรวจและระบุทิศทางของ FX Transaction
 *
 * ไฟล์นี้ไม่มีการอ่านหรือเขียน Excel
 * จึงสามารถใช้ทดสอบ Business Rule แยกจาก Reconcile Flow ได้
 * ------------------------------------------------------------------
 */

import {
  ReconcileRecord,
} from "../shared/record";

import {
  FXU_REPORT_CODE,
  FXU_TEST_DATA_FIELDS,
  FXU_THB_CURRENCY_CODE,
} from "./fxu-config";

/**
 * ทิศทางของ FX Transaction
 *
 * BUY_FCY:
 * สกุลเงินต่างประเทศแลกเป็นเงินบาท
 *
 * SELL_FCY:
 * เงินบาทแลกเป็นสกุลเงินต่างประเทศ
 *
 * CROSS_CURRENCY:
 * สกุลเงินต่างประเทศแลกกับอีกสกุลเงินหนึ่ง
 * โดยไม่มีขา THB
 *
 * NOT_FX:
 * Source Currency และ Destination Currency เหมือนกัน
 *
 * UNKNOWN:
 * ข้อมูล Currency ไม่ครบ
 */
export type FxuDirection =
  | "BUY_FCY"
  | "SELL_FCY"
  | "CROSS_CURRENCY"
  | "NOT_FX"
  | "UNKNOWN";


/**
 * Normalize ค่าทั่วไป
 *
 * ขั้นตอน:
 * 1. แปลงเป็น String
 * 2. ตัดช่องว่างหน้าและหลัง
 * 3. แปลงเป็นตัวพิมพ์ใหญ่
 */
export const normalizeFxuValue = (
  value: unknown,
): string => {
  return String(
    value ?? "",
  )
    .trim()
    .toUpperCase();
};

/**
 * แปลง Amount เป็นตัวเลข
 *
 * รองรับ:
 * - 1000000
 * - 1,000,000
 * - 999999.99
 *
 * คืนค่า null เมื่อ:
 * - เป็นค่าว่าง
 * - ไม่สามารถแปลงเป็นตัวเลขได้
 */
export const parseFxuAmount = (
  value: unknown,
): number | null => {
  const normalizedValue =
    String(
      value ?? "",
    )
      .replace(
        /,/g,
        "",
      )
      .trim();

  if (
    normalizedValue ===
    ""
  ) {
    return null;
  }

  const amount =
    Number(
      normalizedValue,
    );

  return Number.isFinite(
    amount,
  )
    ? amount
    : null;
};


/**
 * ตัวประเมิน Business Rule ของ DF_FXU
 */
export class FxuRuleEvaluator {

  /**
   * ตรวจ Field ที่ Business Rule จำเป็นต้องใช้
   *
   * ไม่รวม Test No. เพราะใช้แสดงผลเท่านั้น
   * และไม่ใช้ตัดสิน Business Rule
   *
   * ไม่รวม Transaction ID เพราะกรณีหา Matching Key หลักไม่พบ
   * Reconciler อาจใช้ Supporting Field ช่วยค้นหารายการได้
   */
  /**
   * ตรวจ Field ที่จำเป็นสำหรับ Business Rule ของ DF_FXU
   *
   * Field ที่ใช้ตาม Requirement:
   * - Txn Date
   * - From Currency
   * - To Currency
   * - Settled Currency
   * - Settled Amount
   *
   * Field ที่ไม่ใช้ตัดสิน Business Rule:
   * - Test No. ใช้แสดงหมายเลข Test Case
   * - Transaction ID ใช้เป็น Matching Key
   * - From Customer Type Code
   * - From Customer Type Description
   *
   * Customer Type ยังคงถูกตรวจใน Script 2
   * แต่จะยังไม่ถูกนำมาตัดสิน PASS/FAIL ใน Script 3
   * เนื่องจาก Requirement ยังไม่ได้กำหนดเงื่อนไข
   * ที่เชื่อม Customer Type กับ DF_FXU
   */
  validateRequiredRuleFields(
    record: ReconcileRecord,
  ): string[] {
    const validationErrors:
      string[] = [];

    /**
     * ตรวจเฉพาะ Field ที่ Requirement ใช้สำหรับ:
     * - ตรวจ FX Conversion
     * - ตรวจ Settlement Currency
     * - ตรวจ Threshold
     * - Matching Support
     */
    const requiredFields = [
      FXU_TEST_DATA_FIELDS
        .transactionDate,

      FXU_TEST_DATA_FIELDS
        .fromCurrency,

      FXU_TEST_DATA_FIELDS
        .toCurrency,

      FXU_TEST_DATA_FIELDS
        .settledCurrency,

      FXU_TEST_DATA_FIELDS
        .settledAmount,
    ];

    /**
     * ตรวจว่า Field ที่จำเป็นมีข้อมูลหรือไม่
     */
    for (
      const field of
      requiredFields
    ) {
      const value =
        String(
          record.get(
            field,
          ) ?? "",
        ).trim();

      if (
        value !==
        ""
      ) {
        continue;
      }

      validationErrors.push(
        `[${FXU_REPORT_CODE}] ` +
        `Test Data Field "${field}" is empty.`,
      );
    }

    /**
     * ตรวจรูปแบบ Settled Amount
     *
     * กรณีมีข้อมูล:
     * - ต้องสามารถแปลงเป็นตัวเลขได้
     *
     * กรณีไม่มีข้อมูล:
     * - Error จะถูกเพิ่มจาก requiredFields ด้านบนแล้ว
     */
    const settledAmountText =
      String(
        record.get(
          FXU_TEST_DATA_FIELDS
            .settledAmount,
        ) ?? "",
      ).trim();

    if (
      settledAmountText !== "" &&
      parseFxuAmount(
        settledAmountText,
      ) === null
    ) {
      validationErrors.push(
        `[${FXU_REPORT_CODE}] ` +
        `"${FXU_TEST_DATA_FIELDS.settledAmount}" ` +
        `must be a valid number.`,
      );
    }

    return validationErrors;
  }

  /**
   * ตรวจทิศทางของ FX Transaction
   */
  getDirection(
    record: ReconcileRecord,
  ): FxuDirection {
    const fromCurrency =
      normalizeFxuValue(
        record.get(
          FXU_TEST_DATA_FIELDS
            .fromCurrency,
        ),
      );

    const toCurrency =
      normalizeFxuValue(
        record.get(
          FXU_TEST_DATA_FIELDS
            .toCurrency,
        ),
      );

    /**
     * ข้อมูล Currency ไม่ครบ
     */
    if (
      fromCurrency === "" ||
      toCurrency === ""
    ) {
      return "UNKNOWN";
    }

    /**
     * Source Currency เท่ากับ Destination Currency
     * จึงไม่ใช่ FX Conversion
     */
    if (
      fromCurrency ===
      toCurrency
    ) {
      return "NOT_FX";
    }

    /**
     * เงินบาทแลกเป็นเงินตราต่างประเทศ
     */
    if (
      fromCurrency ===
        FXU_THB_CURRENCY_CODE &&
      toCurrency !==
        FXU_THB_CURRENCY_CODE
    ) {
      return "SELL_FCY";
    }

    /**
     * เงินตราต่างประเทศแลกเป็นเงินบาท
     */
    if (
      fromCurrency !==
        FXU_THB_CURRENCY_CODE &&
      toCurrency ===
        FXU_THB_CURRENCY_CODE
    ) {
      return "BUY_FCY";
    }

    /**
     * Source และ Destination เป็นเงินตราต่างประเทศ
     * คนละสกุลกัน โดยไม่มีขา THB
     */
    return "CROSS_CURRENCY";
  }


}