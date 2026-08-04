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

