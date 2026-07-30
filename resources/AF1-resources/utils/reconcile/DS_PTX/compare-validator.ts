/**
 * ======================================================
 * ไฟล์: compare-validator.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวควบคุมการเปรียบเทียบ DS_PTX ระหว่าง Test Data กับ Report
 * ไฟล์นี้สร้างข้อมูลที่คาดหวัง จับคู่ Matching Key ตรวจ Exclusion และรวมผลระดับ Field เป็น PASS หรือ FAIL
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import ExcelJS from "exceljs";

import {
  buildExpectedRows,
} from "./expected-row-builder";

import {
  buildActualRows,
} from "./report-row-builder";

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
} from "./compare-types";

import {
  compareCoreFields,
  compareCustomerFields,
} from "./compare-fields";

import {
  writeCompareResult,
} from "./compare-result-writer";

import {
  getCompareResultOutputPath,
} from "./compare-file-helper";

import {
  isResidentThbToFcdExclusionCase,
  RESIDENT_THB_TO_FCD_REMARK,
} from "./ptx-exclusion-rule";

/**
 * Compare Expected Rows จาก Test Data
 * กับ Actual Rows จาก PTX Report
 */
const compareEngine = (
  reportName: string,
  expectedRows: ExpectedRow[],
  actualRows: ActualRow[],
): CompareResult[] => {

  const results: CompareResult[] = [];

  /**
   * เตรียม Map สำหรับค้นหา Report
   * ด้วย Matching Key
   */
  const actualRowMap =
    new Map<string, ActualRow>();

  for (
    const actualRow of actualRows
  ) {

    actualRowMap.set(
      actualRow.matchingKey,
      actualRow,
    );

  }

  /**
   * ตรวจ Test Data ทีละ Expected Row
   */
  for (
    const expectedRow of expectedRows
  ) {

    /**
     * ค้นหา Matching Key ใน PTX Report
     */
    const actualRow =
      actualRowMap.get(
        expectedRow.matchingKey,
      );

    /**
     * อ่านเงื่อนไขจาก Test Data
     *
     * From Customer = Resident
     * AND From Currency = THB
     * AND To Account Type = FCD
     */
    const isExclusionCase =
      isResidentThbToFcdExclusionCase(
        expectedRow.data,
      );

    /**
     * ==========================================================
     * กรณีพิเศษ:
     * Resident โอนจาก THB ไปยังบัญชี FCD
     *
     * เงื่อนไขนี้ต้องไม่พบรายการใน PTX
     * ==========================================================
     */
    if (
      isExclusionCase
    ) {

      /**
       * ไม่พบใน Report = PASS
       * พบใน Report = FAIL
       */
      const reportNotFound =
        !actualRow;

      results.push({

        matchingKey:
          expectedRow.matchingKey,

        testDataRowNumber:
          expectedRow.rowNumber,

        reportRowNumber:
          actualRow?.rowNumber ?? 0,

        field:
          "PTX Exclusion Rule",

        expected:
          "Not Found In PTX",

        actual:
          actualRow?.matchingKey ?? "",

        status:
          reportNotFound
            ? "PASS"
            : "FAIL",

        remark:
          reportNotFound
            ? RESIDENT_THB_TO_FCD_REMARK
            : `${RESIDENT_THB_TO_FCD_REMARK} แต่พบรายการใน PTX`,

      });

      /**
       * จบการตรวจ Expected Row นี้
       * ไม่ต้องตรวจ Field อื่นต่อ
       */
      continue;

    }

    /**
     * ==========================================================
     * กรณีไม่มี Fee
     *
     * ถ้าไม่ใช่ Exclusion Case และไม่มี Fee
     * Test Case นี้ไม่เกี่ยวข้องกับ DS_PTX
     * ==========================================================
     */
    if (
      !expectedRow.hasFee
    ) {

      results.push({

        matchingKey:
          expectedRow.matchingKey,

        testDataRowNumber:
          expectedRow.rowNumber,

        reportRowNumber:
          0,

        field:
          "No Fee",

        expected:
          "",

        actual:
          "",

        status:
          "SKIP",

        remark:
          "No Fee Data - DS_PTX Not Applicable",

      });

      continue;

    }

    /**
     * ==========================================================
     * กรณีปกติที่มี Fee
     * แต่ค้นหา Matching Key ใน Report ไม่พบ
     *
     * เนื่องจากไม่เข้าเงื่อนไขยกเว้น
     * ผลจึงต้องเป็น FAIL
     * ==========================================================
     */
    if (
      !actualRow
    ) {

      results.push({

        matchingKey:
          expectedRow.matchingKey,

        testDataRowNumber:
          expectedRow.rowNumber,

        reportRowNumber:
          0,

        field:
          "Matching Key",

        expected:
          expectedRow.matchingKey,

        actual:
          "",

        status:
          "FAIL",

        remark:
          "Matching Key Not Found",

      });

      continue;

    }

    /**
     * ==========================================================
     * พบข้อมูลใน Report และเป็นกรณีปกติ
     * ตรวจ Core Fields
     * ==========================================================
     */
    results.push(

      ...compareCoreFields(

        reportName,

        expectedRow,

        actualRow,

      ),

    );

    /**
     * ตรวจ Conditional Fields
     * และ Customer Fields
     */
    results.push(

      ...compareCustomerFields(

        reportName,

        expectedRow,

        actualRow,

      ),

    );

  }

  return results;

};

/**
 * ==============================================================
 * Function หลักสำหรับ Compare Report กับ Test Data
 * ==============================================================
 */
export const compareReportWithTestData = async (
  reportName: string,
  reportFilePath: string,
  testDataFilePath: string,
): Promise<CompareResult[]> => {

  /**
   * เปิดไฟล์ Test Data
   */
  const testWorkbook =
    new ExcelJS.Workbook();

  await testWorkbook.xlsx.readFile(
    testDataFilePath,
  );

  const testWorksheet =
    testWorkbook.worksheets[0];

  if (
    !testWorksheet
  ) {

    throw new Error(
      `Test Data worksheet not found: ${testDataFilePath}`,
    );

  }

  /**
   * เปิดไฟล์ PTX Report
   */
  const reportWorkbook =
    new ExcelJS.Workbook();

  await reportWorkbook.xlsx.readFile(
    reportFilePath,
  );

  /**
   * ค้นหา Worksheet จาก Report Name ก่อน
   *
   * หากไม่พบ จะใช้ Worksheet แรก
   */
  const reportWorksheet =
    reportWorkbook.getWorksheet(
      reportName,
    ) ??
    reportWorkbook.worksheets[0];

  if (
    !reportWorksheet
  ) {

    throw new Error(
      `Report worksheet not found: ${reportFilePath}`,
    );

  }

  /**
   * สร้าง Expected Rows จาก Test Data
   */
  const expectedRows =
    buildExpectedRows(
      testWorksheet,
    );

  /**
   * สร้าง Actual Rows จาก PTX Report
   */
  const actualRows =
    buildActualRows(
      reportWorksheet,
      reportName,
    );

  /**
   * เริ่ม Compare
   */
  const results =
    compareEngine(
      reportName,
      expectedRows,
      actualRows,
    );

  /**
   * นับจำนวน PASS
   */
  const passCount =
    results.filter(
      result =>
        result.status === "PASS",
    ).length;

  /**
   * นับจำนวน FAIL
   */
  const failCount =
    results.filter(
      result =>
        result.status === "FAIL",
    ).length;

  /**
   * นับจำนวน SKIP
   */
  const skipCount =
    results.filter(
      result =>
        result.status === "SKIP",
    ).length;

  /**
   * สร้าง Output Path
   */
  const outputFile =
    getCompareResultOutputPath(
      reportName,
    );

  /**
   * Export ผล Compare เป็น Excel
   */
  await writeCompareResult(

    results,

    expectedRows,

    actualRows,

    outputFile,

  );

  /**
   * แสดงผลสรุปทาง Console
   */
  console.log("");
  console.log("======================================");
  console.log(" Script 3 Compare Result");
  console.log("======================================");

  console.log(
    "Report Name   :",
    reportName,
  );

  console.log(
    "Expected Rows :",
    expectedRows.length,
  );

  console.log(
    "Actual Rows   :",
    actualRows.length,
  );

  console.log(
    "Total Results :",
    results.length,
  );

  console.log(
    "PASS          :",
    passCount,
  );

  console.log(
    "FAIL          :",
    failCount,
  );

  console.log(
    "SKIP          :",
    skipCount,
  );

  console.log(
    "Output File   :",
    outputFile,
  );

  console.log(
    "======================================",
  );

  return results;

};