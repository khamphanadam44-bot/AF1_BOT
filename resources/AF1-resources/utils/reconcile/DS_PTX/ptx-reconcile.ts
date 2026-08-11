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
  PTX_FALLBACK_FIELDS,
} from "./ptx-config";

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
} from "./ptx-types";

import {
  compareCoreFields,
  compareCustomerFields,
  compareValue,
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
 * ชื่อ Header ที่อาจใช้เก็บ Test No.
 *
 * ใช้รูปแบบเดียวกับ ptx-row-builder.ts
 * เพื่อให้การตัดสินว่าแถวใดใช้ Fallback
 * เป็นไปตามกติกาเดียวกัน
 */
const PTX_TEST_NUMBER_HEADERS = [
  "Test No.",
  "Test Script No.",
  "Test No",
  "Test Script No",
];

/**
 * ตรวจว่าค่าเป็นค่าว่างหรือไม่
 */
const isFallbackBlank = (
  value: unknown,
): boolean => {
  return (
    value === undefined ||
    value === null ||
    String(
      value,
    ).trim() === ""
  );
};

/**
 * ตรวจว่า Expected Row ต้องใช้
 * Composite Fallback Matching หรือไม่
 *
 * ใช้ Fallback เฉพาะเมื่อ:
 * 1. Expected Row มี Fee
 * 2. Test No. ทุกชื่อ Header เป็นค่าว่าง
 * 3. Transaction ID/ Reconcile ID เป็นค่าว่าง
 */
const shouldUseCompositeFallback = (
  expectedRow: ExpectedRow,
): boolean => {
  if (
    !expectedRow.hasFee
  ) {
    return false;
  }

  const hasTestNumber =
    PTX_TEST_NUMBER_HEADERS.some(
      (header) =>
        !isFallbackBlank(
          expectedRow.data[
          header
          ],
        ),
    );

  const hasTransactionId =
    !isFallbackBlank(
      expectedRow.data[
      "Transaction ID/ Reconcile ID"
      ],
    );

  return (
    !hasTestNumber &&
    !hasTransactionId
  );
};

/**
 * อ่าน Expected Value
 * ของ Composite Fallback Field
 *
 * TEST_DATA:
 * อ่านจาก Test Data Header
 *
 * CURRENT_FEE_AMOUNT:
 * อ่าน Fee Amount ของ Fee Group
 * ที่ Expected Row กำลังตรวจ
 */
const getFallbackExpectedValue = (
  expectedRow: ExpectedRow,
  fallbackField:
    (typeof PTX_FALLBACK_FIELDS)[number],
): unknown => {
  if (
    fallbackField.valueSource ===
    "CURRENT_FEE_AMOUNT"
  ) {
    return expectedRow.feeAmount;
  }

  if (
    !fallbackField.testDataField
  ) {
    return undefined;
  }

  return expectedRow.data[
    fallbackField.testDataField
  ];
};

/**
 * ตรวจว่า Report Row หนึ่งแถว
 * ตรงกับ Composite Fallback Fields
 * ของ Expected Row หรือไม่
 */
const isCompositeFallbackCandidate = (
  expectedRow: ExpectedRow,
  actualRow: ActualRow,
): boolean => {
  return PTX_FALLBACK_FIELDS.every(
    (fallbackField) => {
      const expectedValue =
        getFallbackExpectedValue(
          expectedRow,
          fallbackField,
        );

      /**
       * Field ขั้นต่ำที่ Expected ว่าง
       * จะไม่สามารถใช้ Report Row นี้
       * เป็น Candidate ได้
       */
      if (
        isFallbackBlank(
          expectedValue,
        )
      ) {
        return (
          !fallbackField.required
        );
      }

      /**
       * Report Field อาจมีมากกว่าหนึ่ง Header
       *
       * ตัวอย่าง From CIF:
       * - Cust Code
       * - CMF CODE
       *
       * Candidate ผ่านเมื่อค่าตรง
       * อย่างน้อยหนึ่ง Header
       */
      /**
       * อ่านค่าจาก Report Header
       * ที่กำหนดไว้สำหรับ Fallback Field นี้
       *
       * ตัวอย่าง From CIF:
       * - Cust Code
       * - CMF CODE
       */
      const actualValues =
        fallbackField
          .reportFields
          .map(
            (reportField) =>
              actualRow.data[
              reportField
              ],
          );

      const hasAnyActualValue =
        actualValues.some(
          (actualValue) =>
            !isFallbackBlank(
              actualValue,
            ),
        );

      /**
       * Field เพิ่มเติม เช่น From CIF และ To CIF
       *
       * หาก Test Data มีข้อมูล
       * แต่ Report Candidate ไม่มีข้อมูล
       * ในทุก Header ที่เกี่ยวข้อง
       * ให้ข้าม Field เพิ่มเติมนี้
       *
       * Candidate จะถูกตรวจต่อด้วย Field ขั้นต่ำ:
       * - Transaction Date
       * - Currency
       * - Fee Amount
       *
       * กฎนี้ไม่ใช้กับ Field ขั้นต่ำ
       * เพราะ Field ขั้นต่ำต้องมีและต้องตรงเสมอ
       */
      if (
        !fallbackField.required &&
        !hasAnyActualValue
      ) {
        return true;
      }

      /**
       * หาก Report Candidate มีข้อมูล
       * ต้องมีอย่างน้อยหนึ่ง Report Header
       * ที่ตรงกับ Expected Value
       *
       * ตัวอย่าง From CIF:
       * Cust Code หรือ CMF CODE
       * ต้องตรงอย่างน้อยหนึ่งช่อง
       */
      return actualValues.some(
        (actualValue) =>
          compareValue(
            expectedValue,
            actualValue,
            fallbackField.compareType,
            fallbackField.tolerance ?? 0,
          ),
      );
    },
  );
};

/**
 * สงวน Report Row ที่มี Matching Key
 * ตรงกับ Expected Row แบบ Exact Matching
 *
 * การสงวนทำก่อนเริ่มวน Expected Row
 * เพื่อป้องกันไม่ให้ Fallback นำ Report Row
 * ของ Exact Matching ที่อยู่ลำดับถัดไปไปใช้ก่อน
 */
const buildReservedExactReportRowNumbers = (
  expectedRows: ExpectedRow[],
  actualRows: ActualRow[],
): Set<number> => {
  const exactMatchingKeys =
    new Set<string>();

  for (
    const expectedRow of expectedRows
  ) {
    const transactionId =
      expectedRow.data[
      "Transaction ID/ Reconcile ID"
      ];

    if (
      expectedRow.hasFee &&
      !isFallbackBlank(
        transactionId,
      )
    ) {
      exactMatchingKeys.add(
        expectedRow.matchingKey,
      );
    }
  }

  const reservedRowNumbers =
    new Set<number>();

  for (
    const actualRow of actualRows
  ) {
    if (
      exactMatchingKeys.has(
        actualRow.matchingKey,
      )
    ) {
      reservedRowNumbers.add(
        actualRow.rowNumber,
      );
    }
  }

  return reservedRowNumbers;
};

/**
 * ค้นหา Report Candidate
 * สำหรับ Composite Fallback Matching
 *
 * ไม่อนุญาตให้ใช้:
 * - Report Row ที่สงวนให้ Exact Matching
 * - Report Row ที่เคยถูก Fallback จับคู่แล้ว
 */
const findCompositeFallbackCandidates = (
  expectedRow: ExpectedRow,
  actualRows: ActualRow[],
  reservedExactRowNumbers:
    Set<number>,
  usedFallbackRowNumbers:
    Set<number>,
): ActualRow[] => {
  return actualRows.filter(
    (actualRow) => {
      if (
        reservedExactRowNumbers.has(
          actualRow.rowNumber,
        )
      ) {
        return false;
      }

      if (
        usedFallbackRowNumbers.has(
          actualRow.rowNumber,
        )
      ) {
        return false;
      }

      return isCompositeFallbackCandidate(
        expectedRow,
        actualRow,
      );
    },
  );
};

/**
 * Compare Expected Rows จาก Test Data
 * กับ Actual Rows จาก PTX Report
 *
 * ลำดับการจับคู่:
 * 1. เตรียม Exact Matching Map
 * 2. สงวน Report Row สำหรับ Exact Matching
 * 3. ตรวจ Expected Row ทีละรายการ
 * 4. ใช้ Exact Matching เมื่อมี Transaction ID
 * 5. ใช้ Composite Fallback เมื่อไม่มีทั้ง
 *    Test No. และ Transaction ID
 * 6. ป้องกัน Report Row ถูก Fallback ใช้ซ้ำ
 * 7. ส่งคู่ข้อมูลที่พบไปตรวจ Field ต่อ
 */
const compareEngine = (
  reportName: string,
  expectedRows: ExpectedRow[],
  actualRows: ActualRow[],
): CompareResult[] => {
  const results:
    CompareResult[] = [];

  /**
   * เตรียม Map สำหรับ Exact Matching
   * ด้วย Reference Transaction Number
   */
  const actualRowMap =
    new Map<
      string,
      ActualRow
    >();

  for (
    const actualRow of actualRows
  ) {
    /**
     * Matching Key ว่างไม่สามารถใช้
     * Exact Matching ได้
     *
     * แต่ Actual Row ยังคงอยู่ใน actualRows
     * เพื่อให้ Fallback นำไปตรวจได้
     */
    if (
      isFallbackBlank(
        actualRow.matchingKey,
      )
    ) {
      continue;
    }

    actualRowMap.set(
      actualRow.matchingKey,
      actualRow,
    );
  }

  /**
   * สงวน Report Row ที่ต้องใช้
   * สำหรับ Exact Matching
   *
   * Fallback จะไม่นำ Row เหล่านี้ไปใช้
   */
  const reservedExactRowNumbers =
    buildReservedExactReportRowNumbers(
      expectedRows,
      actualRows,
    );

  /**
   * เก็บ Row Number ของ Report
   * ที่ถูก Composite Fallback ใช้แล้ว
   *
   * Report Row เดียวจึงไม่สามารถ
   * ถูก Fallback จับคู่ซ้ำได้
   */
  const usedFallbackRowNumbers =
    new Set<number>();

  /**
   * ตรวจ Test Data
   * ทีละ Expected Row
   */
  for (
    const expectedRow of expectedRows
  ) {
    const useCompositeFallback =
      shouldUseCompositeFallback(
        expectedRow,
      );

    /**
     * เริ่มต้นด้วย Exact Matching
     *
     * Expected Row ที่ใช้ Fallback
     * จะมี Internal Matching Key
     * จึงไม่พบใน actualRowMap
     */
    let actualRow =
      actualRowMap.get(
        expectedRow.matchingKey,
      );

    /**
     * จำนวน Candidate ที่พบจาก Fallback
     *
     * ใช้แยกผล:
     * 0 = Not Found
     * 1 = Matched
     * มากกว่า 1 = Ambiguous
     */
    let fallbackCandidateCount =
      0;

    if (
      useCompositeFallback
    ) {
      const fallbackCandidates =
        findCompositeFallbackCandidates(
          expectedRow,
          actualRows,
          reservedExactRowNumbers,
          usedFallbackRowNumbers,
        );

      fallbackCandidateCount =
        fallbackCandidates.length;

      /**
       * พบมากกว่า 1 Candidate
       *
       * ห้ามเลือก Candidate แถวแรก
       * ให้สร้าง FAIL: Ambiguous
       * แล้วตรวจ Expected Row ถัดไป
       */
      if (
        fallbackCandidateCount > 1
      ) {
        results.push({
          matchingKey:
            expectedRow.matchingKey,

          testDataRowNumber:
            expectedRow.rowNumber,

          reportRowNumber:
            0,

          field:
            "Composite Fallback Matching",

          expected:
            "1 matching Report Row",

          actual:
            `${fallbackCandidateCount} matching Report Rows`,

          status:
            "FAIL",

          remark:
            "Ambiguous Composite Fallback Match: " +
            `${fallbackCandidateCount} Report Rows Found`,
        });

        continue;
      }

      /**
       * พบ Candidate เพียง 1 แถว
       *
       * Mapping Report Row กับ Expected Row
       * แล้วส่งไป Compare Field ต่อ
       */
      if (
        fallbackCandidateCount === 1
      ) {
        actualRow =
          fallbackCandidates[0];

        /**
         * ป้องกัน Report Row เดียว
         * ถูก Fallback จับคู่ซ้ำ
         */
        usedFallbackRowNumbers.add(
          actualRow.rowNumber,
        );
      } else {
        /**
         * ไม่พบ Candidate
         *
         * ยังไม่สร้าง FAIL ตรงนี้
         * เพราะต้องตรวจ Exclusion Rule ก่อน
         */
        actualRow =
          undefined;
      }
    }

    /**
     * อ่านเงื่อนไข Exclusion จาก Test Data
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
            : (
              `${RESIDENT_THB_TO_FCD_REMARK} ` +
              "แต่พบรายการใน PTX"
            ),
      });

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
     * Composite Fallback ไม่พบ Candidate
     * ==========================================================
     *
     * ไม่ throw เพื่อให้ Script 3
     * ตรวจ Expected Row ถัดไปต่อได้
     */
    if (
      useCompositeFallback &&
      fallbackCandidateCount === 0
    ) {
      results.push({
        matchingKey:
          expectedRow.matchingKey,

        testDataRowNumber:
          expectedRow.rowNumber,

        reportRowNumber:
          0,

        field:
          "Composite Fallback Matching",

        expected:
          "Txn Date + Currency + Fee Amount" +
          " + Optional CIF Fields",

        actual:
          "",

        status:
          "FAIL",

        remark:
          "Composite Fallback Match Not Found",
      });

      continue;
    }

    /**
     * ==========================================================
     * Exact Matching ไม่พบ Matching Key
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
     * พบข้อมูลใน Report
     *
     * ใช้ได้ทั้ง:
     * - Exact Matching
     * - Composite Fallback Matching
     *
     * จากนั้นตรวจ Core Fields
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