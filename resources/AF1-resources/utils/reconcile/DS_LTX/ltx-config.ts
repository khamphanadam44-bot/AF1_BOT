/**
 * ltx-config.ts
 * ------------------------------------------------------------
 * Config สำหรับการ Reconcile DS_LTX ใน Script 3
 *
 * หน้าที่หลัก:
 * 1. กำหนด Header ที่ใช้จับกลุ่มข้อมูล
 * 2. กำหนด Header ที่ใช้สร้าง DR และ FE
 * 3. กำหนด Business Rule ของแต่ละ Field
 * 4. กำหนดวิธีเปรียบเทียบข้อมูล
 * 5. กำหนด Expected Presence และ Expected Absence
 *
 * หมายเหตุ:
 * จำนวน Fee Group จะไม่กำหนดในไฟล์นี้
 * แต่จะตรวจจาก Header ใน Test Data จริง
 * ภายใน ltx-expected-case-builder.ts
 * ------------------------------------------------------------
 */

import type {
  ReportPresenceRule,
} from "../shared/presence-rule";

import {
  RESIDENT_THB_TO_FCD_MUST_NOT_EXIST,
} from "../shared/presence-rule";

import type {
  ReportCode,
} from "../../../config/report-config";

import {
  getReportRuntimeConfig,
} from "../../../config/report-runtime.config";

/**
 * วิธีเปรียบเทียบค่าของแต่ละ Field
 */
export type CompareMode =
  | "exact"
  | "amountTolerance"
  | "fixedValue"
  | "dateWithIdFallback";

/**
 * Rule สำหรับเปรียบเทียบ Field
 * ระหว่าง AF1 Report กับ Test Data
 */
export interface ReconcileFieldRule {
  /**
   * ชื่อ Header ฝั่ง AF1 Report
   */
  reportField: string;

  /**
   * ชื่อ Header ฝั่ง Test Data
   *
   * ใช้ null เมื่อ Field ไม่มีข้อมูลต้นทางจาก Test Data
   * เช่น Rule ที่ใช้ fixedValue
   */
  testDataField: string | null;

  /**
   * วิธีเปรียบเทียบข้อมูล
   */
  compareMode: CompareMode;

  /**
   * ค่าคงที่ที่คาดหวัง
   *
   * ใช้เมื่อ compareMode เป็น fixedValue
   */
  fixedValue?: string;

  /**
   * ส่วนต่างของตัวเลขที่ยอมรับได้
   *
   * ใช้เมื่อ compareMode เป็น amountTolerance
   *
   * หากไม่กำหนด จะใช้ DEFAULT_AMOUNT_TOLERANCE
   */
  tolerance?: number;

  /**
   * Header สำรองฝั่ง Test Data
   *
   * ใช้เมื่อเปรียบเทียบกับ testDataField หลักไม่ผ่าน
   *
   * ตัวอย่าง:
   * หาก Transaction Amount ไม่ตรงกับ
   * From Transfer Amount ให้ลองเปรียบเทียบกับ
   * From Debit Amount
   */
  fallbackTestDataField?: string;

  /**
   * true:
   * เป็น Core Field ที่ต้องตรวจทุก Test Case
   *
   * false:
   * เป็น Conditional Field
   */
  isRequiredForAllCases: boolean;

  /**
   * ใช้กับ Conditional Field
   *
   * หาก Field นี้ใน Test Data ว่าง
   * ระบบจะข้ามการเปรียบเทียบ Rule นี้
   *
   * หากไม่กำหนด จะใช้ testDataField
   * ของ Rule นี้เป็นตัวตรวจ
   */
  skipWhenTestDataFieldEmpty?: string;

  /**
   * ใช้กับ Conditional Field
   *
   * ตรวจค่าจากฝั่ง AF1 Report
   * แทนการตรวจจาก Test Data
   *
   * หาก Field นี้ใน AF1 Report ว่าง
   * ระบบจะข้าม Rule นี้
   *
   * ใช้กับ Inflow Transaction Purpose
   * ซึ่งจะตรวจเฉพาะแถวที่เป็น Inflow
   */
  onlyWhenReportFieldHasValue?: string;

  /**
   * หมายเหตุสำหรับอธิบาย Business Rule
   */
  remark?: string;

  /**
   * จำกัดว่า Rule นี้ใช้กับ Suffix ใด
   *
   * ตัวอย่าง:
   * ["DR"] หมายถึงตรวจเฉพาะแถว DR
   *
   * หากไม่กำหนด:
   * Rule นี้ใช้กับทุก Suffix
   *
   * Field กลุ่มข้อมูลลูกค้า เช่น:
   * - Beneficiary Name
   * - Cust Code
   * - Cust Name
   *
   * ต้องตรวจเฉพาะแถว DR เพราะแถว FE
   * ใช้เก็บข้อมูลธนาคารผู้รับค่าธรรมเนียม
   * ไม่ใช่ข้อมูลลูกค้าจาก Test Data
   */
  applicableSuffixes?: string[];
}

/**
 * Header ที่ใช้จับกลุ่ม Record
 * ของ Test Case เดียวกัน
 */
export interface ReconcileGroupKeyFields {
  /**
   * Header เลขที่บัญชีฝั่ง AF1 Report
   */
  reportAccountField: string;

  /**
   * Header เลขที่บัญชีฝั่ง Test Data
   */
  testDataAccountField: string;

  /**
   * Headerสกุลเงินฝั่ง AF1 Report
   */
  reportCurrencyField: string;

  /**
   * Header สกุลเงินฝั่ง Test Data
   */
  testDataCurrencyField: string;
}

/**
 * Field ที่ตั้งใจ Copy ไปแสดงในผลลัพธ์เท่านั้น
 * แต่ยังไม่มี Expected Value หรือ Business Rule สำหรับเปรียบเทียบ
 */
export interface ReconcileOutputOnlyField {
  reportField: string;
  reason: string;
}

/**
 * Config หลักสำหรับ Reconcile Report
 */
export interface ReconcileReportConfig {
  /**
   * หมายเลขแถว Header ของ AF1 Report
   */
  headerRowNumber: number;

  /**
   * หมายเลขแถว Header ของ Test Data
   */
  testDataHeaderRowNumber: number;

  /**
   * Field ที่ใช้จับกลุ่มข้อมูล
   */
  groupKeyFields: ReconcileGroupKeyFields;

  /**
   * Header Reference Transaction Number
   * ฝั่ง AF1 Report
   */
  referenceNumberReportField: string;

  /**
   * Suffix ของแถว Debit
   */
  drSuffixLabel: string;

  /**
   * Suffix ของแถว Fee
   */
  feSuffixLabel: string;

  /**
   * Header Transaction ID ฝั่ง Test Data
   *
   * ใช้เป็น Base สำหรับสร้าง Reference:
   * - Transaction ID + DR
   * - Transaction ID + FE
   */
  testDataIdField: string;

  /**
   * Header Test No. ฝั่ง Test Data
   *
   * ใช้แสดงหมายเลข Test Case และแจ้งเตือนเมื่อไม่ได้กรอกค่า
   */
  testDataTestNoField: string;

  /**
   * Header จำนวนเงินฝั่ง AF1 Report
   */
  transactionAmountReportField: string;

  /**
   * Header จำนวนเงินหลักฝั่ง Test Data
   * สำหรับแถว DR
   */
  drAmountTestDataField: string;

  /**
   * Header จำนวนเงินสำรองฝั่ง Test Data
   *
   * ใช้เมื่อเทียบกับ drAmountTestDataField
   * แล้วไม่ผ่าน
   */
  drAmountFallbackTestDataField: string;

  /**
   * Rule ตัดสินว่ารายการ:
   * - ต้องมีใน Report
   * - ห้ามมีใน Report
   */
  reportPresenceRules:
    readonly ReportPresenceRule[];

  /**
   * รายการ Rule สำหรับเปรียบเทียบ Field
   */
  fieldRules: ReconcileFieldRule[];

  /**
   * Field ที่ตั้งใจ Copy ไปแสดง แต่ยังไม่ใช้เปรียบเทียบ
   *
   * ต้องระบุเหตุผลทุก Field เพื่อป้องกัน Mapping Field
   * หลุดจากการตรวจสอบโดยไม่มีคำอธิบาย
   */
  outputOnlyFields:
    readonly ReconcileOutputOnlyField[];
}

/**
 * ส่วนต่างของตัวเลขที่ยอมรับได้
 *
 * ตัวอย่าง:
 * Expected = 100
 * Actual = 100.01
 * ถือว่ายังอยู่ภายใน Tolerance
 */
export const DEFAULT_AMOUNT_TOLERANCE =
  0.01;

/**
 * ไม่มีการกำหนดจำนวน Fee Group ใน Config นี้
 *
 * จำนวน Fee Group ของ DS_LTX จะตรวจจาก Header
 * ใน Test Data จริงภายใน ltx-expected-case-builder.ts
 *
 * จึงไม่มี RECONCILE_FEE_TYPE_COUNT
 * และไม่มีจำนวน Fee Group แบบ Hard code
 */

/**
 * ค่าคงที่ทางธุรกิจของ DS_LTX
 *
 * แยกเป็น Constant เพื่อไม่ให้มี Magic Number
 * กระจายอยู่ภายใน Business Rule
 */
const RELATIONSHIP_OTHER_CODE =
  "172064";

/**
 * อ่านหมายเลขแถว Header ของ Test Data
 * จาก report-runtime.config.ts
 *
 * ทำให้ Script 2, Script 3 และ Script 4
 * ใช้ Config จากแหล่งเดียวกัน
 */
const LTX_TEST_DATA_HEADER_ROW_NUMBER =
  getReportRuntimeConfig(
    "DS_LTX",
  ).testDataHeaderRowNumber;

/**
 * Config สำหรับ Reconcile Report
 *
 * ใช้ Partial<Record<ReportCode, ...>>
 * เพื่อให้ TypeScript ช่วยตรวจ Report Code
 * และช่วยตรวจการพิมพ์ชื่อ Report ผิด
 */
export const RECONCILE_CONFIG: Partial<
  Record<
    ReportCode,
    ReconcileReportConfig
  >
> = {
  // ====================================================
  // DS_LTX
  // ====================================================
  DS_LTX: {
    /**
     * Header ของ DS_LTX Report อยู่แถวที่ 1
     */
    headerRowNumber: 1,

    /**
     * Header ของ Test Data อ่านจาก Runtime Config
     */
    testDataHeaderRowNumber:
      LTX_TEST_DATA_HEADER_ROW_NUMBER,

    /**
     * Field ที่ใช้จับกลุ่ม Record
     * ของ Test Case เดียวกัน
     */
    groupKeyFields: {
      reportAccountField:
        "FI Arrangement Number",

      testDataAccountField:
        "From Account ( A/C Client/Sender)",

      reportCurrencyField:
        "Currency Id",

      testDataCurrencyField:
        "From Currency (CCY)",
    },

    /**
     * Reference Transaction Number
     *
     * DR:
     * Transaction ID + DR
     *
     * FE:
     * Transaction ID + FE
     */
    referenceNumberReportField:
      "Reference Transaction Number",

    drSuffixLabel:
      "DR",

    feSuffixLabel:
      "FE",

    testDataIdField:
      "Transaction ID/ Reconcile ID",

    testDataTestNoField:
      "Test No.",

    /**
     * Field จำนวนเงิน
     */
    transactionAmountReportField:
      "Transaction Amount",

    drAmountTestDataField:
      "From Transfer Amount",

    drAmountFallbackTestDataField:
      "From Debit Amount",

    /**
     * Rule Expected Presence/Absence
     *
     * Resident โอน THB ไปบัญชี FCD
     * ต้องไม่พบรายการใน DS_LTX Report
     */
    reportPresenceRules: [
      RESIDENT_THB_TO_FCD_MUST_NOT_EXIST,
    ],

    /**
     * Rule สำหรับเปรียบเทียบ Field
     */
    fieldRules: [
      // ==================================================
      // Core Fields
      // ต้องตรวจสอบทุก Test Case
      // ==================================================
      {
        reportField:
          "Transaction Date",

        testDataField:
          "Txn Date",

        compareMode:
          "dateWithIdFallback",

        isRequiredForAllCases:
          true,

        remark:
          "ต้องตรงกัน หากไม่ตรงให้ดึงจาก Reference Transaction Number ตำแหน่ง 7-12",
      },

      {
        reportField:
          "Currency Id",

        testDataField:
          "From Currency (CCY)",

        compareMode:
          "exact",

        isRequiredForAllCases:
          true,

        remark:
          "ต้องตรงกัน ห้ามเปลี่ยนสกุลเงิน",
      },

      {
        /**
         * ตรวจเฉพาะแถว DR
         *
         * แถว FE แสดงข้อมูลจาก Report เท่านั้น
         * เพราะยังไม่มี Rule เปรียบเทียบที่ยืนยันแล้ว
         */
        reportField:
          "Outflow Transaction Purpose",

        testDataField:
          "From BOT Purpose code",

        compareMode:
          "exact",

        isRequiredForAllCases:
          true,

        applicableSuffixes: [
          "DR",
        ],

        remark:
          "ตรวจเฉพาะแถว DR; แถว FE แสดงข้อมูลจาก Report เท่านั้น",
      },

      // ==================================================
      // Conditional Fields
      // ตรวจเฉพาะเมื่อเข้าเงื่อนไข
      // ==================================================
      {
        reportField:
          "Beneficiary or Sender Name",

        testDataField:
          "To CIF Name (Beneficiary)",

        compareMode:
          "exact",

        isRequiredForAllCases:
          false,

        applicableSuffixes: [
          "DR",
        ],

        remark:
          "ตรวจเฉพาะแถว DR เพราะแถว FE เป็นข้อมูลธนาคารผู้รับค่าธรรมเนียม",
      },

      {
        reportField:
          "Country Id of Beneficiary or Sender",

        testDataField:
          "To Region (Bene Country)",

        compareMode:
          "exact",

        isRequiredForAllCases:
          false,

        applicableSuffixes: [
          "DR",
        ],

        skipWhenTestDataFieldEmpty:
          "To CIF Name (Beneficiary)",

        remark:
          "ตรวจเมื่อ Beneficiary Name มีค่า และตรวจเฉพาะแถว DR",
      },

      {
        reportField:
          "Relationship with Beneficiary or Sender",

        testDataField:
          null,

        compareMode:
          "fixedValue",

        fixedValue:
          RELATIONSHIP_OTHER_CODE,

        isRequiredForAllCases:
          false,

        applicableSuffixes: [
          "DR",
        ],

        skipWhenTestDataFieldEmpty:
          "To CIF Name (Beneficiary)",

        remark:
          "Resident/Non Resident ใช้ค่า 172064 และตรวจเฉพาะแถว DR",
      },

      {
        reportField:
          "Cust Code",

        testDataField:
          "From CIF No. (Client/Sender)",

        compareMode:
          "exact",

        isRequiredForAllCases:
          false,

        applicableSuffixes: [
          "DR",
        ],

        remark:
          "ตรวจเมื่อมีค่าใน Test Data และตรวจเฉพาะแถว DR โดยตัด Leading Zero ก่อนเปรียบเทียบ",
      },

      {
        reportField:
          "Cust Name",

        testDataField:
          "From CIF Name (Client/Sender)",

        compareMode:
          "exact",

        isRequiredForAllCases:
          false,

        applicableSuffixes: [
          "DR",
        ],

        skipWhenTestDataFieldEmpty:
          "From CIF No. (Client/Sender)",

        remark:
          "ตรวจเมื่อ Cust Code มีค่า และตรวจเฉพาะแถว DR",
      },

      {
        reportField:
          "Inflow Transaction Purpose",

        testDataField:
          "From BOT Purpose code",

        compareMode:
          "exact",

        isRequiredForAllCases:
          false,

        onlyWhenReportFieldHasValue:
          "Inflow Transaction Purpose",

        remark:
          "ตรวจเฉพาะเมื่อแถวใน AF1 Report เป็น Inflow",
      },
    ],

    /**
     * Field กลุ่มนี้ถูกคัดลอกจาก AF1 Report ไปแสดงโดยตรง
     * และยังไม่ถูก Highlight จนกว่าจะมี Expected Value
     * หรือ Business Rule ที่ได้รับการยืนยัน
     */
    outputOnlyFields: [
      {
        reportField:
          "Loan Deposit Transaction Type",
        reason:
          "ตรวจเฉพาะการมี Header; ยังไม่มี Rule เปรียบเทียบค่าที่ได้รับการยืนยัน",
      },
      {
        reportField:
          "Payment Method",
        reason:
          "ยังไม่มี Expected Value ที่ยืนยันสำหรับ DS_LTX",
      },
      {
        reportField:
          "From Transaction Type",
        reason:
          "ยังไม่มี Expected Value ที่ยืนยันสำหรับ DS_LTX",
      },
      {
        reportField:
          "To Transaction Type",
        reason:
          "ยังไม่มี Expected Value ที่ยืนยันสำหรับ DS_LTX",
      },
      {
        reportField:
          "Installment Number",
        reason:
          "ไม่มีข้อมูลต้นทางใน Test Data สำหรับเปรียบเทียบ",
      },
      {
        reportField:
          "Approval Document Number",
        reason:
          "ไม่มีข้อมูลต้นทางใน Test Data สำหรับเปรียบเทียบ",
      },
      {
        reportField:
          "CMF CODE",
        reason:
          "ยังไม่มี Expected Value ที่ยืนยันสำหรับ DS_LTX",
      },
      {
        reportField:
          "Data Set Date",
        reason:
          "เป็น Reference Field สำหรับแสดงผลเท่านั้น",
      },
      {
        reportField:
          "Data Submission Period",
        reason:
          "เป็น Reference Field สำหรับแสดงผลเท่านั้น",
      },
    ],
  },
};

/**
 * รายชื่อ Report Field ที่ DS_LTX จัดประเภทไว้แล้วทั้งหมด
 *
 * ใช้ตรวจ Coverage ระหว่าง Mapping Config และ Reconcile Config
 */
export const getHandledReportFields = (
  config: ReconcileReportConfig,
): string[] => {
  return [
    config.referenceNumberReportField,
    config.groupKeyFields.reportAccountField,
    config.transactionAmountReportField,
    ...config.fieldRules.map(
      (rule) => rule.reportField,
    ),
    ...config.outputOnlyFields.map(
      (field) => field.reportField,
    ),
  ];
};

/**
 * คืน Config ของ Report ที่พร้อม Reconcile
 *
 * หากไม่พบ Config หรือ Config ไม่ครบ
 * ระบบจะ Throw Error และหยุดการ Reconcile
 */
export const getReconcileConfig = (
  reportCode: string,
): ReconcileReportConfig => {
  const config =
    RECONCILE_CONFIG[
      reportCode as ReportCode
    ];

  if (
    !config ||
    config.groupKeyFields
      .reportAccountField === ""
  ) {
    throw new Error(
      `Reconcile Config ของ Report "${reportCode}" ยังไม่ครบถ้วน`,
    );
  }

  return config;
};