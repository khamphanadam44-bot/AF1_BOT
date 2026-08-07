/**
 * ============================================================================
 * ไฟล์: ptx-config.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้เก็บกติกาและแผนที่ข้อมูลสำหรับการตรวจ DS_PTX
 *
 * ถ้า ptx-rules.ts เป็นคนลงมือเปรียบเทียบ
 * ไฟล์นี้ก็คือคู่มือที่บอกว่าต้องตรวจ Field ใดและใช้กติกาใด
 *
 * สิ่งที่กำหนดไว้ในไฟล์:
 * 1. Field Mapping
 * 2. Compare Rules
 * 3. Business Values
 * 4. Expected Value Resolver
 *
 * เมื่อต้องเพิ่มหรือเปลี่ยน Mapping ของ DS_PTX
 * ควรเริ่มตรวจสอบจากไฟล์นี้ก่อนแก้ Business Logic
 * ============================================================================
 */

/**
 * ส่วน import ด้านล่าง คือการนำเครื่องมือหรือโครงสร้างข้อมูล
 * จากไฟล์อื่นมาใช้ในไฟล์นี้ เปรียบเหมือนการหยิบอุปกรณ์ที่เตรียมไว้แล้ว
 * มาใช้งาน โดยไม่ต้องเขียนทุกอย่างซ้ำใหม่
 */

import {
  ExpectedRow,
} from "./ptx-types";

/**
 * ============================================================================
 * ประเภทของการเปรียบเทียบ
 * ============================================================================
 */

export type CompareType =
  | "TEXT"
  | "DATE"
  | "NUMBER"
  | "AMOUNT";

export interface CompareRule {

  /**
   * ชื่อ Field ใน AF1 Report
   */
  reportField: string;

  /**
   * วิธี Compare
   */
  compareType: CompareType;

  /**
   * ถ้า Expected เป็นค่าว่าง
   * ให้ข้ามการ Compare
   */
  skipIfExpectedBlank?: boolean;

  /**
   * ค่าความคลาดเคลื่อน
   */
  tolerance?: number;

}

/**
 * ============================================================================
 * รูปแบบข้อมูลของ Field Mapping
 * ============================================================================
 */

export interface CompareFieldMapping {

  /**
   * ชื่อ Field ใน AF1 Report
   */
  reportField: string;

  /**
   * คืนชื่อ Header ของ Test Data
   *
   * runningNumber:
   * 1 = Fee ลำดับที่ 1
   * 2 = Fee ลำดับที่ 2
   * 3 = Fee ลำดับที่ 3
   */
  getTestDataField: (
    runningNumber: number,
  ) => string;

}

/**
 * ============================================================================
 * Composite Fallback Matching Configuration
 * ============================================================================
 *
 * ใช้เมื่อ Test Data ไม่มีทั้ง:
 * - Test No.
 * - Transaction ID/ Reconcile ID
 *
 * Configuration นี้มีหน้าที่บอกว่า:
 * - ต้องอ่าน Expected Value จากที่ใด
 * - ต้องนำไปเทียบกับ Header ใดใน DS_PTX Report
 * - Field ใดเป็น Field บังคับ
 * - Field ใดใช้เพิ่มความแม่นยำเมื่อ Test Data มีข้อมูล
 *
 * ไฟล์นี้เก็บเฉพาะ Configuration
 * ยังไม่มีหน้าที่ค้นหา Candidate หรือตัดสินผล PASS/FAIL
 */

/**
 * แหล่งที่มาของ Expected Value
 *
 * TEST_DATA:
 * อ่านค่าจาก Header ของ Test Data โดยตรง
 *
 * CURRENT_FEE_AMOUNT:
 * อ่าน Fee Amount จาก Expected Row
 * ของ Fee Group ที่กำลังตรวจ
 */
export type PtxFallbackValueSource =
  | "TEST_DATA"
  | "CURRENT_FEE_AMOUNT";

/**
 * รูปแบบของ Field
 * ที่ใช้ใน Composite Fallback Matching
 */
export interface PtxFallbackField {

  /**
   * ชื่อที่ใช้เรียก Field ภายใน Logic
   * และใช้แสดงในข้อความอธิบายผล
   */
  fieldName: string;

  /**
   * แหล่งที่มาของ Expected Value
   */
  valueSource:
    PtxFallbackValueSource;

  /**
   * ชื่อ Header ใน Test Data
   *
   * ใช้เมื่อ valueSource เป็น TEST_DATA
   *
   * หาก valueSource เป็น CURRENT_FEE_AMOUNT
   * จะไม่ต้องกำหนด testDataField
   */
  testDataField?: string;

  /**
   * Header ใน DS_PTX Report
   * ที่สามารถใช้ตรวจ Field นี้ได้
   *
   * หากมีหลาย Header:
   * Candidate จะผ่านเมื่อค่าตรงกับ
   * Report Header อย่างน้อยหนึ่ง Header
   *
   * ตัวอย่าง From CIF:
   * - Cust Code
   * - CMF CODE
   */
  reportFields: string[];

  /**
   * true:
   * เป็น Field ขั้นต่ำที่ต้องมีและต้องตรง
   *
   * false:
   * ตรวจเฉพาะเมื่อ Test Data มีข้อมูล
   */
  required: boolean;

  /**
   * วิธี Normalize และ Compare ค่า
   */
  compareType:
    CompareType;

  /**
   * ค่าความคลาดเคลื่อนสำหรับจำนวนเงิน
   *
   * ตัวอย่าง:
   * 0.01 หมายถึงยอมให้จำนวนเงิน
   * ต่างกันได้ไม่เกิน 0.01
   */
  tolerance?: number;

}

/**
 * Field ที่ใช้ค้นหา DS_PTX Report Row
 * ด้วย Composite Fallback Matching
 *
 * Field ขั้นต่ำ:
 * 1. Txn Date
 * 2. From Currency (CCY)
 * 3. Fee Amount ของ Fee Group ที่กำลังตรวจ
 *
 * Field เพิ่มเติมเมื่อ Test Data มีข้อมูล:
 * 4. From CIF No. (Client/Sender)
 *    เทียบกับ Cust Code หรือ CMF CODE
 * 5. To CIF No. (Beneficiary)
 *    เทียบกับ Involved Party Id
 */
export const PTX_FALLBACK_FIELDS:
  PtxFallbackField[] = [

    /**
     * Field ขั้นต่ำที่ 1:
     * วันที่ทำรายการ
     */
    {
      fieldName:
        "Transaction Date",

      valueSource:
        "TEST_DATA",

      testDataField:
        "Txn Date",

      reportFields: [
        "Receive Payment Transaction Date",
      ],

      required:
        true,

      compareType:
        "DATE",
    },

    /**
     * Field ขั้นต่ำที่ 2:
     * สกุลเงินต้นทาง
     */
    {
      fieldName:
        "Currency",

      valueSource:
        "TEST_DATA",

      testDataField:
        "From Currency (CCY)",

      reportFields: [
        "Currency Id",
      ],

      required:
        true,

      compareType:
        "TEXT",
    },

    /**
     * Fieldขั้นต่ำที่ 3:
     * Fee Amount ของ Fee Group
     * ที่กำลังตรวจอยู่
     */
    {
      fieldName:
        "Fee Amount",

      valueSource:
        "CURRENT_FEE_AMOUNT",

      reportFields: [
        "Transaction Amount in Foreign Currency",
      ],

      required:
        true,

      compareType:
        "AMOUNT",

      tolerance:
        0.01,
    },

    /**
     * Field เพิ่มเติม:
     * ตรวจเฉพาะเมื่อ From CIF
     * ใน Test Data มีข้อมูล
     *
     * Candidate จะผ่านเมื่อค่าตรงกับ:
     * - Cust Code
     * หรือ
     * - CMF CODE
     */
    {
      fieldName:
        "From CIF",

      valueSource:
        "TEST_DATA",

      testDataField:
        "From CIF No. (Client/Sender)",

      reportFields: [
        "Cust Code",
        "CMF CODE",
      ],

      required:
        false,

      compareType:
        "TEXT",
    },

    /**
     * Field เพิ่มเติม:
     * ตรวจเฉพาะเมื่อ To CIF
     * ใน Test Data มีข้อมูล
     */
    {
      fieldName:
        "To CIF",

      valueSource:
        "TEST_DATA",

      testDataField:
        "To CIF No. (Beneficiary)",

      reportFields: [
        "Involved Party Id",
      ],

      required:
        false,

      compareType:
        "TEXT",
    },

  ];

/**
 * ============================================================================
 * DS_PTX Field Mapping
 * ============================================================================
 */

export const DS_PTX_FIELD_MAPPING: CompareFieldMapping[] = [

  /**
   * --------------------------------------------------------------------------
   * ข้อมูลหลักที่ต้องตรวจ (Core Fields)
   * --------------------------------------------------------------------------
   */

  {
    reportField:
      "Receive Payment Transaction Date",

    getTestDataField: () =>
      "Txn Date",
  },

  {
    reportField:
      "Currency Id",

    getTestDataField: () =>
      "From Currency (CCY)",
  },

  {
    reportField:
      "Transaction Amount in Foreign Currency",

    getTestDataField: (
      runningNumber,
    ) => {

      /**
       * ชื่อ Header ใน Test Data ปัจจุบัน:
       *
       * Fee 1 → Fee Amount Type 1
       * Fee 2 → Fee Amount 2
       * Fee 3 → Fee Amount 3
       */
      if (
        runningNumber === 1
      ) {

        return "Fee Amount Type 1";

      }

      return `Fee Amount ${runningNumber}`;

    },
  },

  /**
   * --------------------------------------------------------------------------
   * ข้อมูลที่ตรวจตามเงื่อนไข (Conditional Fields)
   * --------------------------------------------------------------------------
   */

  {
    reportField:
      "Cust Code",

    getTestDataField: () =>
      "From CIF No. (Client/Sender)",
  },

  {
    reportField:
      "CMF CODE",

    getTestDataField: () =>
      "From CIF No. (Client/Sender)",
  },

  {
    reportField:
      "Cust Name",

    getTestDataField: () =>
      "From CIF Name (Client/Sender)",
  },

  {
    reportField:
      "Involved Party Id",

    getTestDataField: () =>
      "To CIF No. (Beneficiary)",
  },

  {
    reportField:
      "Involved Party Name",

    getTestDataField: () =>
      "To CIF Name (Beneficiary)",
  },

  {
    reportField:
      "Country Id of Involved Party",

    getTestDataField: () =>
      "To Region (Bene Country)",
  },

];

/**
 * ============================================================================
 * DS_PTX Compare Rules
 * ============================================================================
 */

export const DS_PTX_COMPARE_RULES: CompareRule[] = [

  /**
   * --------------------------------------------------------------------------
   * ข้อมูลหลักที่ต้องตรวจ (Core Fields)
   * --------------------------------------------------------------------------
   */

  {
    reportField:
      "Receive Payment Transaction Date",

    compareType:
      "DATE",
  },

  {
    reportField:
      "Currency Id",

    compareType:
      "TEXT",
  },

  {
    reportField:
      "Payment Method",

    compareType:
      "TEXT",
  },

  {
    reportField:
      "Receive Payment Transaction Type",

    compareType:
      "TEXT",
  },

  {
    reportField:
      "Receive Payment Item Type",

    compareType:
      "TEXT",
  },

  {
    reportField:
      "Transaction Amount in Foreign Currency",

    compareType:
      "AMOUNT",

    tolerance:
      0.01,
  },

  /**
   * --------------------------------------------------------------------------
   * ข้อมูลที่ตรวจตามเงื่อนไข (Conditional Fields)
   * --------------------------------------------------------------------------
   */

  {
    reportField:
      "Cust Code",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

  {
    reportField:
      "CMF CODE",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

  {
    reportField:
      "Cust Name",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

  {
    reportField:
      "Involved Party Id",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

  {
    reportField:
      "Involved Party Name",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

  {
    reportField:
      "Country Id of Involved Party",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

  {
    reportField:
      "Receive Payment Item Description",

    compareType:
      "TEXT",

    skipIfExpectedBlank:
      true,
  },

];

/**
 * ============================================================================
 * ตัวช่วยหาค่าตามกติกาทางธุรกิจของ DS_PTX
 * ============================================================================
 */

/**
 * Payment Method
 */
export const resolvePaymentMethod = (
  _expected: ExpectedRow,
): string => {

  return "234004";

};

/**
 * Receive Payment Transaction Type
 */
export const resolveReceivePaymentTransactionType = (
  _expected: ExpectedRow,
): string => {

  return "270001";

};

/**
 * Receive Payment Item Type
 */
export const resolveReceivePaymentItemType = (
  _expected: ExpectedRow,
): string => {

  return "268002";

};

/**
 * Receive Payment Item Description
 */
export const resolveReceivePaymentItemDescription = (
  _expected: ExpectedRow,
): string => {

  return "";

};

/**
 * ============================================================================
 * Mapping Helper
 * ============================================================================
 */

export const getCompareFieldMappings = (
  reportName: string,
): CompareFieldMapping[] => {

  switch (
    reportName
  ) {

    case "DS_PTX":

      return DS_PTX_FIELD_MAPPING;

    default:

      throw new Error(
        `Compare Field Mapping not found: ${reportName}`,
      );

  }

};

export const getCompareFieldMapping = (
  reportName: string,
  reportField: string,
): CompareFieldMapping | undefined => {

  return getCompareFieldMappings(
    reportName,
  ).find(
    mapping =>
      mapping.reportField === reportField,
  );

};

/**
 * ============================================================================
 * Compare Rule Helper
 * ============================================================================
 */

export const getCompareRules = (
  reportName: string,
): CompareRule[] => {

  switch (
    reportName
  ) {

    case "DS_PTX":

      return DS_PTX_COMPARE_RULES;

    default:

      throw new Error(
        `Compare Rule not found : ${reportName}`,
      );

  }

};

export const getCompareRule = (
  reportName: string,
  reportField: string,
): CompareRule | undefined => {

  return getCompareRules(
    reportName,
  ).find(
    rule =>
      rule.reportField === reportField,
  );

};

/**
 * ============================================================================
 * Resolve Expected Value
 * ============================================================================
 *
 * Report Field
 *      ↓
 * Business Resolver หรือ Test Data Mapping
 *      ↓
 * Expected Value
 */
export const resolveExpectedValue = (
  expected: ExpectedRow,
  reportField: string,
  reportName = "DS_PTX",
): unknown => {

  /**
   * Business Fields
   */
  switch (
    reportField
  ) {

    case "Payment Method":

      return resolvePaymentMethod(
        expected,
      );

    case "Receive Payment Transaction Type":

      return resolveReceivePaymentTransactionType(
        expected,
      );

    case "Receive Payment Item Type":

      return resolveReceivePaymentItemType(
        expected,
      );

    case "Receive Payment Item Description":

      return resolveReceivePaymentItemDescription(
        expected,
      );

    default:

      break;

  }

  /**
   * Mapping Fields
   */
  const mapping =
    getCompareFieldMapping(
      reportName,
      reportField,
    );

  if (
    !mapping
  ) {

    return undefined;

  }

  const testDataField =
    mapping.getTestDataField(
      expected.runningNumber,
    );

  return expected.data[
    testDataField
  ];

};
