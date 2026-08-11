/**
 * FXM-rules.ts
 * ------------------------------------------------------------------
 * Business Rule ของ DF_FXM สำหรับ Script 3
 *
 * หน้าที่:
 * 1. ตรวจข้อมูลบังคับที่ใช้ใน Business Rule
 * 2. ตรวจว่ารายการเป็น FX Conversion หรือไม่
 * 3. แยกทิศทางซื้อหรือขายเงินตราต่างประเทศ
 * 4. ตรวจ Threshold ต่ำกว่า 1,000,000 USD
 * 5. ตัดสินว่ารายการต้องมีหรือไม่ต้องมีใน DF_FXM
 * 6. สร้าง Expected Leg Type และ Leg Type Name
 *
 * ไฟล์นี้ไม่มีการอ่านหรือเขียน Excel
 * จึงสามารถใช้ทดสอบ Business Rule แยกจาก Reconcile Flow ได้
 * ------------------------------------------------------------------
 */

import {
  ReconcileRecord,
} from "../shared/record";

import {
  FXM_LEG_TYPE_NAMES,
  FXM_LEG_TYPES,
  FXM_REPORT_CODE,
  FXM_TEST_DATA_FIELDS,
  FXM_THB_CURRENCY_CODE,
} from "./fxm-config";

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
export type FXMDirection =
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
export const normalizeFXMValue = (
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
export const parseFXMAmount = (
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
 * ตัวประเมิน Business Rule ของ DF_FXM
 */
export class FXMRuleEvaluator {

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
   * ตรวจ Field ที่จำเป็นสำหรับ Business Rule ของ DF_FXM
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
   * ที่เชื่อม Customer Type กับ DF_FXM
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
      FXM_TEST_DATA_FIELDS
        .transactionDate,

      FXM_TEST_DATA_FIELDS
        .fromCurrency,

      FXM_TEST_DATA_FIELDS
        .toCurrency,

      FXM_TEST_DATA_FIELDS
        .settledCurrency,

      FXM_TEST_DATA_FIELDS
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
        `[${FXM_REPORT_CODE}] ` +
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
          FXM_TEST_DATA_FIELDS
            .settledAmount,
        ) ?? "",
      ).trim();

    if (
      settledAmountText !== "" &&
      parseFXMAmount(
        settledAmountText,
      ) === null
    ) {
      validationErrors.push(
        `[${FXM_REPORT_CODE}] ` +
        `"${FXM_TEST_DATA_FIELDS.settledAmount}" ` +
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
  ): FXMDirection {
    const fromCurrency =
      normalizeFXMValue(
        record.get(
          FXM_TEST_DATA_FIELDS
            .fromCurrency,
        ),
      );

    const toCurrency =
      normalizeFXMValue(
        record.get(
          FXM_TEST_DATA_FIELDS
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
        FXM_THB_CURRENCY_CODE &&
      toCurrency !==
        FXM_THB_CURRENCY_CODE
    ) {
      return "SELL_FCY";
    }

    /**
     * เงินตราต่างประเทศแลกเป็นเงินบาท
     */
    if (
      fromCurrency !==
        FXM_THB_CURRENCY_CODE &&
      toCurrency ===
        FXM_THB_CURRENCY_CODE
    ) {
      return "BUY_FCY";
    }

    /**
     * Source และ Destination เป็นเงินตราต่างประเทศ
     * คนละสกุลกัน โดยไม่มีขา THB
     */
    return "CROSS_CURRENCY";
  }

  /**
   * คืน Expected Leg Type จากทิศทางของรายการ
   */
  getExpectedLegType(
    direction: FXMDirection,
  ): string | undefined {
    if (
      direction ===
      "BUY_FCY"
    ) {
      return FXM_LEG_TYPES
        .buyForeignCurrency;
    }

    if (
      direction ===
      "SELL_FCY"
    ) {
      return FXM_LEG_TYPES
        .sellForeignCurrency;
    }

    /**
     * CROSS_CURRENCY ยังไม่สามารถระบุ Leg Type ได้
     * จนกว่าจะได้ Settlement/Intermediary Use Case
     */
    return undefined;
  }

  /**
   * คืน Expected Leg Type Name
   * ให้สัมพันธ์กับ Expected Leg Type
   */
  getExpectedLegTypeName(
    direction: FXMDirection,
  ): string | undefined {
    if (
      direction ===
      "BUY_FCY"
    ) {
      return FXM_LEG_TYPE_NAMES[
        FXM_LEG_TYPES
          .buyForeignCurrency
      ];
    }

    if (
      direction ===
      "SELL_FCY"
    ) {
      return FXM_LEG_TYPE_NAMES[
        FXM_LEG_TYPES
          .sellForeignCurrency
      ];
    }

    return undefined;
  }

}