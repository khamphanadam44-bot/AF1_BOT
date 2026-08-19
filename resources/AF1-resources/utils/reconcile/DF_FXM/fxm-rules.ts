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
  FXM_USD_THRESHOLD,
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
 * ผลการตัดสินว่ารายการควรอยู่ใน DF_FXM หรือไม่
 *
 * MUST_EXIST:
 * ต้องพบรายการใน DF_FXM
 *
 * MUST_NOT_EXIST:
 * ต้องไม่พบรายการใน DF_FXM
 *
 * CANNOT_DECIDE:
 * ข้อมูลไม่เพียงพอสำหรับตัดสิน
 */
export type FXMPresenceExpectation =
  | "MUST_EXIST"
  | "MUST_NOT_EXIST"
  | "CANNOT_DECIDE";

/**
 * ผลการประเมิน Business Rule ของ Test Data หนึ่งแถว
 */
export interface FXMRuleDecision {
  direction: FXMDirection;

  expectation:
    FXMPresenceExpectation;

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
  ): FXMRuleDecision {
    const validationErrors =
      this.validateRequiredRuleFields(
        record,
      );

    const direction =
      this.getDirection(
        record,
      );

    const usdEquivalentAmount =
      parseFXMAmount(
        record.get(
          FXM_TEST_DATA_FIELDS
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
     * จึงต้องไม่พบใน DF_FXM
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
          "จึงไม่ใช่ FX Conversion และต้องไม่พบใน DF_FXM",

        failRemark:
          "พบรายการใน DF_FXM ทั้งที่ Source Currency " +
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
          `[${FXM_REPORT_CODE}] ` +
          `"${FXM_TEST_DATA_FIELDS.settledAmount}" ` +
          `is invalid.`,
        ],

        passRemark:
          "",

        failRemark:
          `ไม่สามารถอ่านค่า ` +
          `${FXM_TEST_DATA_FIELDS.settledAmount}`,

        requiresReview:
          true,
      };
    }

        /**
     * ต่ำกว่า 1,000,000 USD
     *
     * ยอดไม่เข้าเงื่อนไขของ DF_FXM
     * จึงต้องไม่พบใน DF_FXM
     * และต้องพิจารณารายงานใน DF_FXU
     */
    if (
      usdEquivalentAmount <
      FXM_USD_THRESHOLD
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
          `ซึ่งต่ำกว่า ` +
          `${formatAmount(FXM_USD_THRESHOLD)} USD ` +
          `จึงต้องไม่พบใน DF_FXM`,

        failRemark:
          `พบรายการใน DF_FXM ทั้งที่ ` +
          `USD Equivalent Amount = ` +
          `${formatAmount(usdEquivalentAmount)} ` +
          `ซึ่งต้องพิจารณารายงานใน DF_FXU`,

        requiresReview:
          false,
      };
    }

    /**
     * Cross Currency และยอดตั้งแต่
     * 1,000,000 USD ขึ้นไป
     *
     * Requirement ระบุว่าต้องใช้ Settlement Currency
     * และ Payment Intermediary เช่น NIUM
     * ช่วยตัดสินว่าใครเป็นผู้ทำ FX Conversion
     *
     * ขณะนี้ยังไม่มี Use Case ที่ยืนยันครบถ้วน
     * จึงยังไม่สามารถตัดสินว่ารายการ
     * ต้องพบหรือไม่ต้องพบใน DF_FXM
     */
    if (
      direction ===
      "CROSS_CURRENCY"
    ) {
      return {
        direction,

        expectation:
          "CANNOT_DECIDE",

        usdEquivalentAmount,

        expectedLegType,
        expectedLegTypeName,

        validationErrors: [],

        passRemark:
          "",

        failRemark:
          "Cross Currency: ยังไม่สามารถตัดสินได้ เนื่องจาก " +
          "Requirement ของ Settlement Currency และ " +
          "Payment Intermediary เช่น NIUM ยังไม่ครบถ้วน",

        requiresReview:
          true,
      };
    }

    /**
     * เป็น FX Conversion ที่มีขา THB
     * และยอดตั้งแต่ 1,000,000 USD ขึ้นไป
     * จึงต้องพบใน DF_FXM
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
        `ซึ่งไม่น้อยกว่า ` +
        `${formatAmount(FXM_USD_THRESHOLD)} USD`,

      failRemark:
        `ไม่พบรายการใน DF_FXM ทั้งที่เป็น FX Conversion ` +
        `และ USD Equivalent Amount ตั้งแต่ ` +
        `${formatAmount(FXM_USD_THRESHOLD)} USD ขึ้นไป`,

      /**
       * Cross Currency ถูกแยกไปตัดสินก่อนหน้านี้แล้ว
       *
       * จุดนี้เหลือเฉพาะ FX Conversion ที่มีขา THB
       * และสามารถตัดสิน Leg Type ได้
       */
      requiresReview:
        false,
    };
  }
}