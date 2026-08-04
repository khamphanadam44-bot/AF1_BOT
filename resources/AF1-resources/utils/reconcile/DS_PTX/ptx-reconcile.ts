/**
 * ======================================================
 * ไฟล์: ptx-reconcile.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวควบคุมหลักของ Script 3 สำหรับ DS_PTX
 *
 * ลำดับการทำงาน:
 * 1. หา Checked Report ล่าสุด
 * 2. สร้างตำแหน่งไฟล์ผล Reconcile
 * 3. เปิดไฟล์ Test Data และ AF1 Report
 * 4. เรียก ptx-row-builder.ts เพื่อสร้าง Expected Row และ Actual Row
 * 5. จับคู่ข้อมูลด้วย Matching Key
 * 6. ตรวจ Business Rule จาก ptx-rules.ts
 * 7. ส่งผลให้ ptx-result-writer.ts สร้างไฟล์ Excel
 *
 * ======================================================
 */

import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

import {
  getCheckedReportHeaderDir,
} from "../../../config/paths.config";

import {
  buildActualRows,
  buildExpectedRows,
} from "./ptx-row-builder";

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
} from "./ptx-types";

import {
  compareCoreFields,
  compareCustomerFields,
  isResidentThbToFcdExclusionCase,
  RESIDENT_THB_TO_FCD_REMARK,
} from "./ptx-rules";

import {
  writeCompareResult,
} from "./ptx-result-writer";


/**
 * Folder หลักสำหรับเก็บผล Reconcile
 */
const RECONCILE_REPORT_ROOT_FOLDER =
  path.resolve(
    process.cwd(),
    "Test_result",
    "Reconcile-report",
  );

/**
 * สร้าง Timestamp
 *
 * รูปแบบ:
 * yyyyMMdd_HHmmss
 */
const getTimestamp = (): string => {
  const now = new Date();

  const yyyy =
    now.getFullYear();

  const MM =
    String(
      now.getMonth() + 1,
    ).padStart(
      2,
      "0",
    );

  const dd =
    String(
      now.getDate(),
    ).padStart(
      2,
      "0",
    );

  const HH =
    String(
      now.getHours(),
    ).padStart(
      2,
      "0",
    );

  const mm =
    String(
      now.getMinutes(),
    ).padStart(
      2,
      "0",
    );

  const ss =
    String(
      now.getSeconds(),
    ).padStart(
      2,
      "0",
    );

  return (
    `${yyyy}${MM}${dd}_${HH}${mm}${ss}`
  );
};

/**
 * หา Checked Report ล่าสุด
 * ของ Report ที่ระบุ
 */
const getLatestReportPath = (
  reportCode: string,
): string => {
  const reportFolder =
    getCheckedReportHeaderDir(
      reportCode,
    );

  if (
    !fs.existsSync(
      reportFolder,
    )
  ) {
    throw new Error(
      `Report folder not found: ${reportFolder}`,
    );
  }

  const files =
    fs
      .readdirSync(
        reportFolder,
      )
      .filter(
        (file) =>
          file
            .toLowerCase()
            .endsWith(
              ".xlsx",
            ),
      );

  if (
    files.length === 0
  ) {
    throw new Error(
      `No report found in: ${reportFolder}`,
    );
  }

  /**
   * เรียงจากไฟล์ใหม่ที่สุด
   * ไปหาไฟล์เก่าที่สุด
   */
  files.sort(
    (
      firstFile,
      secondFile,
    ) => {
      const firstModifiedTime =
        fs
          .statSync(
            path.join(
              reportFolder,
              firstFile,
            ),
          )
          .mtime
          .getTime();

      const secondModifiedTime =
        fs
          .statSync(
            path.join(
              reportFolder,
              secondFile,
            ),
          )
          .mtime
          .getTime();

      return (
        secondModifiedTime -
        firstModifiedTime
      );
    },
  );

  return path.join(
    reportFolder,
    files[0],
  );
};

/**
 * สร้างตำแหน่งไฟล์ผล Reconcile
 * ของ DS_PTX
 */
const getCompareResultOutputPath = (
  reportName: string,
): string => {
  /**
   * ทำชื่อ Report ให้เป็นรูปแบบมาตรฐาน
   *
   * ds_ptx เป็น DS_PTX
   * DS-PTX เป็น DS_PTX
   */
  const normalizedReportName =
    String(
      reportName,
    )
      .trim()
      .toUpperCase()
      .replace(
        /-/g,
        "_",
      );

  const reconcileReportFolder =
    path.join(
      RECONCILE_REPORT_ROOT_FOLDER,
      normalizedReportName,
    );

  /**
   * สร้าง Folder หากยังไม่มี
   */
  fs.mkdirSync(
    reconcileReportFolder,
    {
      recursive: true,
    },
  );

  const fileName =
    `${normalizedReportName}_Compare_Result_` +
    `${getTimestamp()}.xlsx`;

  return path.join(
    reconcileReportFolder,
    fileName,
  );
};

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

/**
 * Entry Point หลักของ DS_PTX สำหรับ Script 3
 *
 * Script 3 ส่งเข้ามาเพียง:
 * 1. ชื่อ Report
 * 2. Path ของ Test Data
 *
 * ส่วนการหา Checked Report ล่าสุดและสร้าง Output Path
 * จะถูกจัดการภายใน ptx-reconcile.ts
 */
export const reconcilePtxReport = async (
  reportName: string,
  testDataFilePath: string,
): Promise<CompareResult[]> => {
  const reportFilePath =
    getLatestReportPath(
      reportName,
    );

  console.log(
    "REPORT PATH",
  );

  console.log(
    reportFilePath,
  );

  console.log(
    "================================",
  );

  console.log(
    "TEST DATA PATH",
  );

  console.log(
    testDataFilePath,
  );

  console.log(
    "================================",
  );

  return compareReportWithTestData(
    reportName,
    reportFilePath,
    testDataFilePath,
  );
};