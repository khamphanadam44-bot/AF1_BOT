/**
 * compare-validator.ts
 * ------------------------------------------------------------
 * ตัวควบคุม Logic การเปรียบเทียบข้อมูล DS_FTX
 *
 * หน้าที่หลัก:
 *
 * 1. ตรวจ Matching Key ของ Test Data
 * 2. ตรวจ Exclusion Rule ของ DS_FTX
 * 3. ค้นหา Matching Key ใน Report
 * 4. ตรวจ Matching Key ซ้ำใน Report
 * 5. เรียกเปรียบเทียบ Core Fields
 * 6. รวมผลทั้งหมดเป็น CompareResult[]
 * ------------------------------------------------------------
 */

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
} from "./compare-types";

import {
  buildActualRowIndex,
  findActualRowsByMatchingKey,
} from "./actual-row-index";

import {
  compareCoreFields,
} from "./compare-fields";

import {
  evaluateFtxExclusion,
} from "./ftx-exclusion-rule";

/**
 * แปลงค่าทั่วไปเป็นข้อความ
 *
 * ใช้สำหรับตรวจสอบว่า Matching Key
 * มีข้อมูลหรือไม่
 */
const normalizeText = (
  value: unknown,
): string => {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(
    value,
  ).trim();
};

/**
 * สร้างผล FAIL กรณี Matching Key
 * ใน Test Data เป็นค่าว่าง
 */
const createEmptyMatchingKeyResult = (
  expectedRow: ExpectedRow,
): CompareResult => {
  return {
    /**
     * Matching Key ไม่มีค่า
     */
    matchingKey:
      "",

    /**
     * Test No. จาก Test Data
     */
    testScriptNo:
      expectedRow.testScriptNo,

    /**
     * หมายเลขแถวจริงใน Test Data
     */
    testDataRowNumber:
      expectedRow.rowNumber,

    /**
     * ยังไม่สามารถค้นหาแถวใน Report ได้
     */
    reportRowNumber:
      0,

    /**
     * Field ที่เกิดปัญหา
     */
    field:
      "Matching Key",

    /**
     * สิ่งที่ระบบคาดหวัง
     */
    expected:
      "Transaction ID/ Reconcile ID must not be empty",

    /**
     * ค่าที่พบจริง
     */
    actual:
      "",

    /**
     * Matching Key ว่างจึงเป็น FAIL
     */
    status:
      "FAIL",

    /**
     * รายละเอียดข้อผิดพลาด
     */
    remark:
      "Matching Key is empty in Test Data",
  };
};

/**
 * สร้างผล PASS กรณีเข้า Exclusion Rule
 *
 * รายการที่เข้า Exclusion Rule
 * ไม่ต้องนำไปค้นหาใน Report DS_FTX
 *
 * แต่ยังต้องแสดงในผลลัพธ์เป็น PASS
 * พร้อม Remark อธิบายสาเหตุ
 */
const createExclusionResult = (
  expectedRow: ExpectedRow,
  remark: string,
): CompareResult => {
  return {
    matchingKey:
      expectedRow.matchingKey,

    testScriptNo:
      expectedRow.testScriptNo,

    testDataRowNumber:
      expectedRow.rowNumber,

    /**
     * ไม่ได้จับคู่กับแถวใน Report
     */
    reportRowNumber:
      0,

    /**
     * บอกว่าผลลัพธ์นี้มาจาก
     * Exclusion Rule
     */
    field:
      "Exclusion Rule",

    /**
     * รายการนี้ไม่ต้องแสดงใน DS_FTX
     */
    expected:
      "Excluded from DS_FTX",

    /**
     * ระบบไม่ได้นำไปเปรียบเทียบ
     * กับข้อมูลใน Report
     */
    actual:
      "Not compared",

    /**
     * ตาม Requirement:
     * Exclusion ต้องแสดงเป็น PASS
     */
    status:
      "PASS",

    /**
     * Remark ที่ได้จาก
     * ftx-exclusion-rule.ts
     */
    remark,
  };
};

/**
 * สร้างผล FAIL กรณีไม่พบ Matching Key
 * ใน Report DS_FTX
 */
const createMissingMatchingKeyResult = (
  expectedRow: ExpectedRow,
): CompareResult => {
  return {
    matchingKey:
      expectedRow.matchingKey,

    testScriptNo:
      expectedRow.testScriptNo,

    testDataRowNumber:
      expectedRow.rowNumber,

    /**
     * ไม่พบแถวใน Report
     */
    reportRowNumber:
      0,

    field:
      "Matching Key",

    /**
     * คาดหวังว่าจะพบ Matching Key นี้
     * ใน Ref. TX No.
     */
    expected:
      expectedRow.matchingKey,

    /**
     * แต่ค้นหาไม่พบ
     */
    actual:
      "",

    status:
      "FAIL",

    remark:
      `Matching Key Not Found in DS_FTX: ${expectedRow.matchingKey}`,
  };
};

/**
 * สร้างผล FAIL กรณี Matching Key
 * ซ้ำกันใน Report DS_FTX
 */
const createDuplicateMatchingKeyResult = (
  expectedRow: ExpectedRow,
  duplicatedRows: ActualRow[],
): CompareResult => {
  /**
   * ดึงหมายเลขแถวทั้งหมดที่พบ Matching Key ซ้ำ
   *
   * ตัวอย่าง:
   * [2, 8, 15]
   */
  const rowNumbers =
    duplicatedRows.map(
      (
        row,
      ) => {
        return row.rowNumber;
      },
    );

  return {
    matchingKey:
      expectedRow.matchingKey,

    testScriptNo:
      expectedRow.testScriptNo,

    testDataRowNumber:
      expectedRow.rowNumber,

    /**
     * ใช้หมายเลขแถวแรกที่พบ
     * เพื่อช่วยให้เปิด Excel ตรวจสอบง่ายขึ้น
     */
    reportRowNumber:
      duplicatedRows[0]
        ?.rowNumber ?? 0,

    field:
      "Matching Key",

    /**
     * ปกติหนึ่ง Matching Key
     * ต้องพบเพียงหนึ่งแถว
     */
    expected:
      "1 Report row",

    /**
     * แสดงจำนวนแถวที่พบจริง
     */
    actual:
      `${duplicatedRows.length} Report rows`,

    status:
      "FAIL",

    /**
     * แสดงหมายเลขแถวที่พบข้อมูลซ้ำทั้งหมด
     */
    remark:
      `Duplicate Matching Key in DS_FTX: ${expectedRow.matchingKey} was found at rows ${rowNumbers.join(", ")}`,
  };
};

/**
 * เปรียบเทียบ ExpectedRow จาก Test Data
 * กับ ActualRow จาก Report DS_FTX
 *
 * ลำดับการทำงาน:
 *
 * 1. สร้าง Index ของข้อมูล Report
 * 2. วนตรวจ Test Data ทีละแถว
 * 3. ตรวจ Matching Key ว่าง
 * 4. ตรวจ Exclusion Rule
 * 5. ค้นหา Matching Key ใน Report
 * 6. ตรวจกรณีไม่พบ
 * 7. ตรวจกรณีพบซ้ำ
 * 8. เปรียบเทียบ Core Fields
 *
 * @param expectedRows
 * ข้อมูลที่คาดหวังจาก Test Data
 *
 * @param actualRows
 * ข้อมูลจริงจาก Report DS_FTX
 *
 * @returns
 * ผลการเปรียบเทียบทั้งหมด
 */
export const compareFtxRows = (
  expectedRows: ExpectedRow[],
  actualRows: ActualRow[],
): CompareResult[] => {
  /**
   * สร้างตารางค้นหา Report
   * ด้วย Matching Key จาก Ref. TX No.
   */
  const actualRowIndex =
    buildActualRowIndex(
      actualRows,
    );

  /**
   * Array สำหรับเก็บผลลัพธ์ทั้งหมด
   */
  const results:
    CompareResult[] = [];

  /**
   * วนตรวจ Test Data ทีละแถว
   */
  for (
    const expectedRow of
    expectedRows
  ) {
    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 1:
     * ตรวจ Matching Key ของ Test Data
     * ----------------------------------------------------------
     *
     * Matching Key มาจาก:
     *
     * Transaction ID/ Reconcile ID
     */
    const matchingKey =
      normalizeText(
        expectedRow.matchingKey,
      );

    if (
      matchingKey === ""
    ) {
      results.push(
        createEmptyMatchingKeyResult(
          expectedRow,
        ),
      );

      /**
       * Matching Key ว่าง
       * จึงไม่สามารถทำขั้นตอนอื่นต่อได้
       */
      continue;
    }

    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 2:
     * ตรวจ Exclusion Rule ของ DS_FTX
     * ----------------------------------------------------------
     *
      * Rule ที่ตรวจ ได้แก่:
 *
 * 1. From Currency (CCY) = THB
 *
 * หรือ
 *
 * 2. ต้องผ่านทุกเงื่อนไขพร้อมกัน:
 *    - From Currency ตรงกับ Settled Currency
 *    - ลูกค้าเป็น Resident
 *    - Settled Currency เป็น USD
 *    - Settled Amount อยู่ระหว่าง 0 ถึง 50,000 USD
     */
    const exclusionResult =
      evaluateFtxExclusion(
        expectedRow.data,
      );

    if (
      exclusionResult.isExcluded
    ) {
      /**
       * ตาม Requirement:
       *
       * รายการ Exclusion
       * ต้องแสดงเป็น PASS พร้อม Remark
       *
       * และไม่ต้องนำไปค้นหาใน Report
       */
      results.push(
        createExclusionResult(
          expectedRow,
          exclusionResult.remark,
        ),
      );

      continue;
    }

    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 3:
     * ค้นหา Matching Key ใน Report
     * ----------------------------------------------------------
     *
     * Test Data:
     * Transaction ID/ Reconcile ID
     *
     * Report DS_FTX:
     * Ref. TX No.
     */
    const matchedActualRows =
      findActualRowsByMatchingKey(
        actualRowIndex,
        matchingKey,
      );

    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 4:
     * ไม่พบ Matching Key ใน Report
     * ----------------------------------------------------------
     */
    if (
      matchedActualRows.length ===
      0
    ) {
      results.push(
        createMissingMatchingKeyResult(
          expectedRow,
        ),
      );

      continue;
    }

    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 5:
     * พบ Matching Key ซ้ำใน Report
     * ----------------------------------------------------------
     *
     * Matching Key หนึ่งค่า
     * ควรจับคู่ได้เพียงหนึ่งแถว
     */
    if (
      matchedActualRows.length >
      1
    ) {
      results.push(
        createDuplicateMatchingKeyResult(
          expectedRow,
          matchedActualRows,
        ),
      );

      continue;
    }

    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 6:
     * ได้ Report ที่ตรงกันหนึ่งแถว
     * ----------------------------------------------------------
     */
    const actualRow =
      matchedActualRows[0];

    /**
     * TypeScript ยังมองว่า actualRow
     * อาจเป็น undefined ได้
     *
     * จึงตรวจป้องกันไว้อีกชั้น
     */
    if (
      actualRow === undefined
    ) {
      results.push(
        createMissingMatchingKeyResult(
          expectedRow,
        ),
      );

      continue;
    }

    /**
     * ----------------------------------------------------------
     * ขั้นตอนที่ 7:
     * เปรียบเทียบ Core Fields
     * ----------------------------------------------------------
     *
     * Field ที่ตรวจ:
     *
     * From Currency (CCY)
     *      ↕
     * Buy Currency Id
     *
     * Settled Currency (CCY)
     *      ↕
     * Sell Currency Id
     * Txn Date
     *      ↕
     * Transaction Date
     */
    const coreFieldResults =
      compareCoreFields(
        expectedRow,
        actualRow,
      );

    /**
     * นำผลของ Core Fields
     * มารวมกับผลลัพธ์ทั้งหมด
     */
    results.push(
      ...coreFieldResults,
    );
  }

  return results;
};