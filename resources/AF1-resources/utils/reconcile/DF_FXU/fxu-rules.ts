/**
 * fxu-rules.ts
 * ------------------------------------------------------------------
 * Business Rule ของ DF_FXU สำหรับ Script 3
 *
 * หน้าที่:
 * 1. ตรวจข้อมูลบังคับที่ใช้ใน Business Rule
 * 2. ตรวจว่ารายการเป็น FX Conversion หรือไม่
 * 3. แยกทิศทางซื้อหรือขายเงินตราต่างประเทศ
 * 4. ตรวจ Threshold ต่ำกว่า 1,000,000 USD
 * 5. ตัดสินว่ารายการต้องมีหรือไม่ต้องมีใน DF_FXU
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
  FXU_LEG_TYPE_NAMES,
  FXU_LEG_TYPES,
  FXU_REPORT_CODE,
  FXU_TEST_DATA_FIELDS,
  FXU_THB_CURRENCY_CODE,
  FXU_USD_THRESHOLD,
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
 * ผลการตัดสินว่ารายการควรอยู่ใน DF_FXU หรือไม่
 *
 * MUST_EXIST:
 * ต้องพบรายการใน DF_FXU
 *
 * MUST_NOT_EXIST:
 * ต้องไม่พบรายการใน DF_FXU
 *
 * CANNOT_DECIDE:
 * ข้อมูลไม่เพียงพอสำหรับตัดสิน
 */
export type FxuPresenceExpectation =
  | "MUST_EXIST"
  | "MUST_NOT_EXIST"
  | "CANNOT_DECIDE";

/**
 * ผลการประเมิน Business Rule ของ Test Data หนึ่งแถว
 */
export interface FxuRuleDecision {
  direction: FxuDirection;

  expectation:
    FxuPresenceExpectation;

  usdEquivalentAmount?:
    number;

  expectedLegType?:
    string;

  expectedLegTypeName?:
    string;

  validationErrors:
    string[];

  passRemark:
    string;

  failRemark:
    string;

  requiresReview:
    boolean;
}

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
 * จัดรูปแบบ Amount สำหรับแสดงใน Remark
 */
const formatAmount = (
  amount: number,
): string => {
  return amount.toLocaleString(
    "en-US",
    {
      maximumFractionDigits:
        2,
    },
  );
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

  /**
   * คืน Expected Leg Type จากทิศทางของรายการ
   */
  getExpectedLegType(
    direction: FxuDirection,
  ): string | undefined {
    if (
      direction ===
      "BUY_FCY"
    ) {
      return FXU_LEG_TYPES
        .buyForeignCurrency;
    }

    if (
      direction ===
      "SELL_FCY"
    ) {
      return FXU_LEG_TYPES
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
    direction: FxuDirection,
  ): string | undefined {
    if (
      direction ===
      "BUY_FCY"
    ) {
      return FXU_LEG_TYPE_NAMES[
        FXU_LEG_TYPES
          .buyForeignCurrency
      ];
    }

    if (
      direction ===
      "SELL_FCY"
    ) {
      return FXU_LEG_TYPE_NAMES[
        FXU_LEG_TYPES
          .sellForeignCurrency
      ];
    }

    return undefined;
  }

  /**
   * ประเมิน Test Data หนึ่งแถว
   *
   * ลำดับการตัดสิน:
   * 1. ตรวจข้อมูลบังคับ
   * 2. ตรวจ FX Conversion
   * 3. ตรวจ Threshold
   * 4. สร้าง Expected Leg Type
   */
  evaluate(
    record: ReconcileRecord,
  ): FxuRuleDecision {
    const validationErrors =
      this.validateRequiredRuleFields(
        record,
      );

    const direction =
      this.getDirection(
        record,
      );

    const usdEquivalentAmount =
      parseFxuAmount(
        record.get(
          FXU_TEST_DATA_FIELDS
            .settledAmount,
        ),
      );

    const expectedLegType =
      this.getExpectedLegType(
        direction,
      );

    const expectedLegTypeName =
      this.getExpectedLegTypeName(
        direction,
      );

    /**
     * ข้อมูลที่ต้องใช้ตัดสินไม่ครบ
     */
    if (
      validationErrors.length >
      0
    ) {
      return {
        direction,
        expectation:
          "CANNOT_DECIDE",

        usdEquivalentAmount:
          usdEquivalentAmount ??
          undefined,

        expectedLegType,
        expectedLegTypeName,

        validationErrors,

        passRemark:
          "",

        failRemark:
          validationErrors.join(
            "\n",
          ),

        requiresReview:
          true,
      };
    }

    /**
     * Source Currency เท่ากับ Destination Currency
     *
     * รายการไม่ใช่ FX Conversion
     * จึงต้องไม่พบใน DF_FXU
     */
    if (
      direction ===
      "NOT_FX"
    ) {
      return {
        direction,
        expectation:
          "MUST_NOT_EXIST",

        usdEquivalentAmount:
          usdEquivalentAmount ??
          undefined,

        expectedLegType,
        expectedLegTypeName,

        validationErrors: [],

        passRemark:
          "Source Currency เท่ากับ Destination Currency " +
          "จึงไม่ใช่ FX Conversion และต้องไม่พบใน DF_FXU",

        failRemark:
          "พบรายการใน DF_FXU ทั้งที่ Source Currency " +
          "เท่ากับ Destination Currency",

        requiresReview:
          false,
      };
    }

    /**
     * ป้องกันกรณี Amount เป็น null
     *
     * ตามปกติกรณีนี้จะถูกตรวจพบจาก
     * validateRequiredRuleFields() ก่อนแล้ว
     */
    if (
      usdEquivalentAmount ===
      null
    ) {
      return {
        direction,
        expectation:
          "CANNOT_DECIDE",

        expectedLegType,
        expectedLegTypeName,

        validationErrors: [
          `[${FXU_REPORT_CODE}] ` +
          `"${FXU_TEST_DATA_FIELDS.settledAmount}" ` +
          `is invalid.`,
        ],

        passRemark:
          "",

        failRemark:
          `ไม่สามารถอ่านค่า ` +
          `${FXU_TEST_DATA_FIELDS.settledAmount}`,

        requiresReview:
          true,
      };
    }

    /**
     * ตั้งแต่ 1,000,000 USD ขึ้นไป
     * ต้องไม่พบใน DF_FXU
     * เพราะต้องพิจารณารายงานใน DF_FXM
     */
    if (
      usdEquivalentAmount >=
      FXU_USD_THRESHOLD
    ) {
      return {
        direction,
        expectation:
          "MUST_NOT_EXIST",

        usdEquivalentAmount,

        expectedLegType,
        expectedLegTypeName,

        validationErrors: [],

        passRemark:
          `USD Equivalent Amount = ` +
          `${formatAmount(usdEquivalentAmount)} ` +
          `ซึ่งไม่น้อยกว่า ` +
          `${formatAmount(FXU_USD_THRESHOLD)} USD ` +
          `จึงต้องไม่พบใน DF_FXU`,

        failRemark:
          `พบรายการใน DF_FXU ทั้งที่ ` +
          `USD Equivalent Amount = ` +
          `${formatAmount(usdEquivalentAmount)} ` +
          `ซึ่งต้องพิจารณารายงานใน DF_FXM`,

        requiresReview:
          false,
      };
    }

    /**
     * เป็น FX Conversion และยอดต่ำกว่า 1,000,000 USD
     * จึงต้องพบใน DF_FXU
     */
    return {
      direction,
      expectation:
        "MUST_EXIST",

      usdEquivalentAmount,

      expectedLegType,
      expectedLegTypeName,

      validationErrors: [],

      passRemark:
        `เป็น FX Conversion และ ` +
        `USD Equivalent Amount = ` +
        `${formatAmount(usdEquivalentAmount)} ` +
        `ซึ่งต่ำกว่า ` +
        `${formatAmount(FXU_USD_THRESHOLD)} USD`,

      failRemark:
        `ไม่พบรายการใน DF_FXU ทั้งที่เป็น FX Conversion ` +
        `และ USD Equivalent Amount ต่ำกว่า ` +
        `${formatAmount(FXU_USD_THRESHOLD)} USD`,

      /**
       * CROSS_CURRENCY ยังไม่สามารถระบุ Leg Type ได้
       * เพราะต้องใช้ Settlement/Intermediary Use Case เพิ่มเติม
       */
      requiresReview:
        direction ===
        "CROSS_CURRENCY",
    };
  }
}