/**
 * ============================================================================
 * ไฟล์: expected-row-builder.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้อ่าน Test Data แล้วแปลงแต่ละรายการให้เป็น ExpectedRow
 * ซึ่งเป็นข้อมูลที่ระบบ “คาดหวัง” ว่าควรพบใน AF1 Report
 *
 * กติกาหลัก
 * ----------
 * - Test Data 1 แถวสามารถมี Fee ได้หลายรายการ
 * - Fee 1 รายการจะถูกสร้างเป็น ExpectedRow 1 รายการ
 * - จึงเป็นไปได้ว่า Test Data 1 แถวจะกลายเป็นหลาย Expected Rows
 *
 * ตัวอย่าง
 * --------
 * Test Case หนึ่งมี Fee ลำดับ 1 และ Fee ลำดับ 2
 * ระบบจะสร้าง ExpectedRow จำนวน 2 รายการ และสร้าง Matching Key
 * ที่ลงท้ายด้วย 01 และ 02 ตามลำดับ
 *
 * กรณีไม่มี Fee
 * -------------
 * หาก Test Case ไม่มี Fee ทุกช่อง ระบบยังสร้าง ExpectedRow 1 รายการ
 * โดยใช้ Internal Key เช่น NO_FEE_ROW_10 เพื่อให้ผลลัพธ์แสดงเป็น SKIP
 * แทนที่จะทำให้ Test Case หายไปจากผลตรวจ
 *
 * แถวจะถูกข้ามเฉพาะเมื่อไม่มีทั้ง Test No. และ Transaction ID
 * ============================================================================
 */


/**
 * ส่วน import ด้านล่าง คือการนำเครื่องมือหรือโครงสร้างข้อมูล
 * จากไฟล์อื่นมาใช้ในไฟล์นี้ เปรียบเหมือนการหยิบอุปกรณ์ที่เตรียมไว้แล้ว
 * มาใช้งาน โดยไม่ต้องเขียนทุกอย่างซ้ำใหม่
 */

import ExcelJS from "exceljs";

import {
  getHeadersFromRow,
} from "../../validators/shared/excel-cell.util";

import {
  FEE_TYPE_COUNT,
  getFeeAmountHeader,
} from "../../../config/testdata-config";

import {
  mapRowToObject,
} from "./row-mapper";

import {
  buildMatchingKey,
} from "./build-matching-key";

import {
  ExpectedRow,
  TestDataRow,
} from "./compare-types";

/**
 * Header Row ของ Test Data
 */
const TEST_DATA_HEADER_ROW = 5;

/**
 * แปลงค่าเป็นข้อความ
 */
const toText = (
  value: unknown,
): string => {

  return String(
    value ?? "",
  ).trim();

};

/**
 * แปลงค่าเป็น Number
 */
const toNumber = (
  value: unknown,
): number => {

  if (
    value === undefined ||
    value === null ||
    toText(value) === ""
  ) {

    return 0;

  }

  const result =
    Number(
      value,
    );

  return Number.isNaN(
    result,
  )
    ? 0
    : result;

};

/**
 * ตรวจสอบว่าค่าเป็นค่าว่างหรือไม่
 */
const isBlank = (
  value: unknown,
): boolean => {

  return (
    value === undefined ||
    value === null ||
    toText(value) === ""
  );

};

/**
 * หา Test Script No.
 *
 * รองรับชื่อ Header หลายรูปแบบ
 */
const getTestScriptNo = (
  rowData: TestDataRow,
): string => {

  const possibleHeaders = [

    "Test No.",

    "Test Script No.",

    "Test No",

    "Test Script No",

  ];

  for (
    const header of possibleHeaders
  ) {

    const value =
      toText(
        rowData[header],
      );

    if (
      value !== ""
    ) {

      return value;

    }

  }

  return "";

};

/**
 * หา Transaction ID
 */
const getTransactionId = (
  rowData: TestDataRow,
): string => {

  return toText(
    rowData[
      "Transaction ID/ Reconcile ID"
    ],
  );

};

/**
 * ตรวจสอบว่าเป็นแถวว่างจริงหรือไม่
 *
 * กติกา:
 * - ไม่มี Test No.
 * - และไม่มี Transaction ID
 *
 * ถ้ามี Test No. แต่ไม่มี Fee
 * ต้องไม่ข้าม เพราะต้องสร้าง SKIP
 */
const isActualBlankRow = (
  rowData: TestDataRow,
): boolean => {

  const testScriptNo =
    getTestScriptNo(
      rowData,
    );

  const transactionId =
    getTransactionId(
      rowData,
    );

  return (
    testScriptNo === "" &&
    transactionId === ""
  );

};

/**
 * สร้าง ExpectedRow
 *
 * เดิมอยู่ใน expected-row.factory.ts
 */
const createExpectedRow = (
  rowNumber: number,
  sourceRow: TestDataRow,
  runningNumber: number,
  feeType: string,
  feeAmount: unknown,
  hasFee: boolean,
): ExpectedRow => {

  const transactionId =
    getTransactionId(
      sourceRow,
    );

  const testScriptNo =
    getTestScriptNo(
      sourceRow,
    );

  /**
   * กรณีมี Fee:
   * สร้าง Matching Key จริงเพื่อ Match กับ Report
   *
   * กรณีไม่มี Fee:
   * สร้าง Internal Key สำหรับ Group Output เท่านั้น
   */
  const matchingKey =
    hasFee
      ? buildMatchingKey(
          transactionId,
          runningNumber,
        )
      : `NO_FEE_ROW_${rowNumber}`;

  return {

    rowNumber,

    testScriptNo,

    matchingKey,

    runningNumber,

    feeType:
      toText(
        feeType,
      ),

    feeAmount:
      toNumber(
        feeAmount,
      ),

    hasFee,

    data:
      sourceRow,

  };

};

/**
 * สร้างแถวข้อมูลที่คาดหวังจาก Test Data
 */
export const buildExpectedRows = (
  worksheet: ExcelJS.Worksheet,
): ExpectedRow[] => {

  /**
   * อ่าน Header จากแถวที่ 5
   */
  const headers =
    getHeadersFromRow(
      worksheet,
      TEST_DATA_HEADER_ROW,
    );

  const expectedRows: ExpectedRow[] = [];

  /**
   * เริ่มอ่านข้อมูลหลัง Header
   */
  for (
    let rowNumber =
      TEST_DATA_HEADER_ROW + 1;

    rowNumber <= worksheet.rowCount;

    rowNumber += 1
  ) {

    const row =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * แปลง Excel Row เป็น Object
     */
    const rowData =
      mapRowToObject(
        row,
        headers,
      ) as TestDataRow;

    /**
     * ข้ามเฉพาะแถวว่างจริง
     *
     * ถ้ามี Test No. หรือ Transaction ID
     * ต้องประมวลผลต่อ แม้ไม่มี Fee
     */
    if (
      isActualBlankRow(
        rowData,
      )
    ) {

      continue;

    }

    /**
     * ใช้ตรวจว่า Test Case นี้
     * มี Fee อย่างน้อยหนึ่งรายการหรือไม่
     */
    let hasAnyFee = false;

    /**
     * ตรวจ Fee ทีละลำดับ
     */
    for (
      let feeIndex = 1;

      feeIndex <= FEE_TYPE_COUNT;

      feeIndex += 1
    ) {

      const feeType =
        toText(
          rowData[
            `Fee Type ${feeIndex}`
          ],
        );

      const feeAmountHeader =
  getFeeAmountHeader(
    feeIndex,
  );

const feeAmount =
  rowData[
    feeAmountHeader
  ];
        rowData[
          `Fee Amount Type ${feeIndex}`
        ];

      /**
       * Fee Type และ Fee Amount ว่างทั้งคู่
       *
       * แปลว่า Fee Slot นี้ไม่มีข้อมูล
       */
      if (
        feeType === "" &&
        isBlank(
          feeAmount,
        )
      ) {

        continue;

      }

      /**
       * พบ Fee อย่างน้อยหนึ่งรายการ
       */
      hasAnyFee = true;

      /**
       * 1 Fee = 1 Expected Row
       */
      expectedRows.push(

        createExpectedRow(

          rowNumber,

          rowData,

          feeIndex,

          feeType,

          feeAmount,

          true,

        ),

      );

    }

    /**
     * Test Case ไม่มี Fee ทุกลำดับ
     *
     * สร้าง Expected Row หนึ่งรายการ
     * เพื่อให้ Compare Engine สร้างผล SKIP
     */
    if (
      !hasAnyFee
    ) {

      expectedRows.push(

        createExpectedRow(

          rowNumber,

          rowData,

          0,

          "",

          0,

          false,

        ),

      );

    }

  }

  return expectedRows;

};
