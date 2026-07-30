/**
 * ds-ftx-reconciler.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวควบคุมหลักสำหรับกระบวนการ Reconcile Report DS_FTX
 *
 * ลำดับการทำงาน
 * 1. ตรวจสอบ Path ของ Report
 * 2. ตรวจสอบ Path ของ Test Data
 * 3. ตรวจสอบ Path ของไฟล์ผลลัพธ์
 * 4. ป้องกัน Output เขียนทับ Report หรือ Test Data
 * 5. เปิด Report และ Test Data พร้อมกัน
 * 6. เลือก Worksheet ที่ต้องใช้งาน
 * 7. สร้าง ExpectedRow[] จาก Test Data
 * 8. สร้าง ActualRow[] จาก Report
 * 9. เปรียบเทียบ Expected กับ Actual
 * 10. สร้างไฟล์ Excel ผลลัพธ์
 * 11. สร้างและคืนผลสรุปให้ Script 3
 *
 * หมายเหตุสำคัญ
 * - ไฟล์นี้รองรับเฉพาะ .xlsx
 * - ไม่ได้สร้าง Output Folder
 * - ไม่ Throw Error เมื่อผล Compare มี FAIL
 * - จำนวน PASS/FAIL เป็นจำนวน CompareResult
 *   ไม่ใช่จำนวน Test Data Row
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";

import {
  CompareResult,
} from "./compare-types";

import {
  buildExpectedRows,
} from "./expected-row-builder";

import {
  buildActualRows,
  getReportHeaders,
} from "./report-row-builder";

import {
  compareFtxRows,
} from "./compare-validator";

import {
  writeCompareResult,
} from "./compare-result-writer";

/**
 * ชื่อ Worksheet ที่ต้องพบใน Report DS_FTX
 *
 * การค้นหาไม่สนใจ
 * - ตัวพิมพ์เล็กหรือตัวพิมพ์ใหญ่
 * - ช่องว่างด้านหน้าและด้านหลัง
 *
 * ตัวอย่างชื่อที่ถือว่าตรงกัน:
 * - DS_FTX
 * - ds_ftx
 * - " DS_FTX "
 *
 * แต่ชื่อ "DS FTX" จะไม่ตรง
 * เพราะ Code ไม่ได้เปลี่ยนช่องว่างภายในเป็น _
 */
const DEFAULT_REPORT_SHEET_NAME =
  "DS_FTX";

/**
 * นามสกุลไฟล์ที่อนุญาต
 *
 * รองรับเฉพาะ .xlsx
 * แต่ไม่สนตัวพิมพ์เล็กหรือใหญ่
 */
const EXCEL_FILE_EXTENSION =
  ".xlsx";

/**
 * รูปแบบผลสรุปที่ส่งกลับไปให้ Script 3
 */
export interface FtxReconcileSummary {
  /**
   * Path ของ Report ที่นำมาตรวจสอบ
   */
  reportFilePath: string;

  /**
   * Path ของ Test Data ที่นำมาตรวจสอบ
   */
  testDataFilePath: string;

  /**
   * Path ของไฟล์ผลลัพธ์ที่สร้างเสร็จแล้ว
   */
  outputFilePath: string;

  /**
   * ชื่อ Worksheet ของ Report ที่ระบบเลือกใช้
   */
  reportWorksheetName: string;

  /**
   * ชื่อ Worksheet ของ Test Data ที่ระบบเลือกใช้
   */
  testDataWorksheetName: string;

  /**
   * จำนวน ExpectedRow ที่สร้างจาก Test Data
   *
   * หมายถึงจำนวนแถว Test Data ที่ไม่ใช่แถวว่าง
   * รวมถึงแถวที่ Matching Key ว่าง
   */
  expectedRowCount: number;

  /**
   * จำนวน ActualRow ที่สร้างจาก Report
   *
   * หมายถึงจำนวนแถว Report ที่ไม่ใช่แถวว่าง
   * รวมถึงแถวที่ Matching Key ว่าง
   */
  actualRowCount: number;

  /**
   * จำนวน CompareResult ทั้งหมด
   *
   * ไม่ใช่จำนวน Test Case
   *
   * ตัวอย่าง:
   * Test Data 1 แถวที่ตรวจ Core Field 3 Field
   * อาจสร้าง CompareResult 3 รายการ
   */
  totalResultCount: number;

  /**
   * จำนวน CompareResult ที่มีสถานะ PASS
   *
   * ไม่ใช่จำนวน Test Case ที่ผ่าน
   */
  passedResultCount: number;

  /**
   * จำนวน CompareResult ที่มีสถานะ FAIL
   *
   * ไม่ใช่จำนวน Test Case ที่ไม่ผ่าน
   */
  failedResultCount: number;

  /**
   * จำนวน CompareResult ที่มาจาก Exclusion Rule
   *
   * Exclusion จะมีสถานะ PASS พร้อม Remark
   */
  exclusionResultCount: number;

  /**
   * CompareResult ทั้งหมด
   *
   * Script 3 สามารถนำไปตรวจสอบต่อได้
   */
  results: CompareResult[];
}

/**
 * แปลงค่าทั่วไปให้เป็นข้อความ
 *
 * การทำงาน
 * - null หรือ undefined → ""
 * - ค่าอื่น → String และ trim()
 */
const toText = (
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
 * ปรับชื่อ Worksheet ให้อยู่ในรูปแบบกลาง
 *
 * การทำงาน
 * 1. แปลงเป็นข้อความ
 * 2. ตัดช่องว่างหน้าและหลัง
 * 3. เปลี่ยนเป็นตัวพิมพ์เล็ก
 *
 * ตัวอย่าง:
 * " DS_FTX " → "ds_ftx"
 * "ds_ftx"   → "ds_ftx"
 *
 * หมายเหตุ
 * ไม่ได้ลบหรือเปลี่ยนช่องว่างภายในชื่อ Sheet
 */
const normalizeWorksheetName = (
  worksheetName: unknown,
): string => {
  return toText(
    worksheetName,
  ).toLowerCase();
};

/**
 * ตรวจสอบว่า Path ไม่เป็นข้อความว่าง
 *
 * ถ้าว่างจะ Throw Error
 */
const validateNonEmptyPath = (
  filePath: string,
  fileDescription: string,
): void => {
  if (
    toText(
      filePath,
    ) === ""
  ) {
    throw new Error(
      `${fileDescription} path is empty`,
    );
  }
};

/**
 * ตรวจสอบนามสกุลไฟล์
 *
 * อนุญาตเฉพาะ .xlsx
 * โดยไม่สนตัวพิมพ์เล็กหรือใหญ่
 *
 * ตัวอย่างที่ผ่าน:
 * - Report.xlsx
 * - Report.XLSX
 *
 * ตัวอย่างที่ไม่ผ่าน:
 * - Report.xls
 * - Report.csv
 */
const validateExcelExtension = (
  filePath: string,
  fileDescription: string,
): void => {
  const extension =
    path
      .extname(
        filePath,
      )
      .toLowerCase();

  if (
    extension !==
    EXCEL_FILE_EXTENSION
  ) {
    throw new Error(
      `${fileDescription} must be an .xlsx file: ${filePath}`,
    );
  }
};

/**
 * ตรวจสอบว่า Input Path เป็นไฟล์จริง
 *
 * ขั้นตอน
 * 1. ตรวจว่า Path มีอยู่จริง
 * 2. ตรวจว่า Path ชี้ไปที่ File ไม่ใช่ Folder
 */
const validateInputFileExists = (
  filePath: string,
  fileDescription: string,
): void => {
  if (
    !fs.existsSync(
      filePath,
    )
  ) {
    throw new Error(
      `${fileDescription} not found: ${filePath}`,
    );
  }

  const fileStatus =
    fs.statSync(
      filePath,
    );

  if (
    !fileStatus.isFile()
  ) {
    throw new Error(
      `${fileDescription} is not a file: ${filePath}`,
    );
  }
};

/**
 * ตรวจสอบ Input Excel File
 *
 * ตรวจทั้งหมด 3 เรื่อง:
 * 1. Path ต้องไม่ว่าง
 * 2. นามสกุลต้องเป็น .xlsx
 * 3. ต้องมีไฟล์อยู่จริง
 *
 * ฟังก์ชันนี้ยังไม่ได้เปิดอ่านเนื้อหาภายใน Workbook
 */
const validateInputExcelFile = (
  filePath: string,
  fileDescription: string,
): void => {
  validateNonEmptyPath(
    filePath,
    fileDescription,
  );

  validateExcelExtension(
    filePath,
    fileDescription,
  );

  validateInputFileExists(
    filePath,
    fileDescription,
  );
};

/**
 * ตรวจสอบ Path ของ Output File
 *
 * ตรวจเฉพาะ:
 * - Path ต้องไม่ว่าง
 * - นามสกุลต้องเป็น .xlsx
 *
 * Output File ยังไม่จำเป็นต้องมีอยู่จริง
 * เพราะ compare-result-writer.ts จะเป็นผู้สร้างไฟล์
 *
 * หมายเหตุ
 * ฟังก์ชันนี้ไม่ได้สร้างหรือตรวจสอบ Output Folder
 */
const validateOutputExcelFile = (
  outputFilePath: string,
): void => {
  validateNonEmptyPath(
    outputFilePath,
    "DS_FTX output file",
  );

  validateExcelExtension(
    outputFilePath,
    "DS_FTX output file",
  );
};

/**
 * ตรวจสอบว่า Output File ไม่ใช่ Path เดียวกับ
 * - Report
 * - Test Data
 *
 * ใช้เพื่อป้องกันการเขียนทับไฟล์ต้นทาง
 *
 * หมายเหตุ
 * - เปรียบเทียบจากข้อความที่ผ่าน path.resolve()
 * - ไม่ได้ใช้ fs.realpathSync()
 * - ไม่ได้ตรวจ Symbolic Link
 * - ไม่ได้ตรวจว่า Output File อื่นมีอยู่แล้วหรือไม่
 */
const validateOutputDoesNotOverwriteInput = (
  reportFilePath: string,
  testDataFilePath: string,
  outputFilePath: string,
): void => {
  const resolvedReportPath =
    path.resolve(
      reportFilePath,
    );

  const resolvedTestDataPath =
    path.resolve(
      testDataFilePath,
    );

  const resolvedOutputPath =
    path.resolve(
      outputFilePath,
    );

  if (
    resolvedOutputPath ===
    resolvedReportPath
  ) {
    throw new Error(
      "DS_FTX output file cannot overwrite the Report file",
    );
  }

  if (
    resolvedOutputPath ===
    resolvedTestDataPath
  ) {
    throw new Error(
      "DS_FTX output file cannot overwrite the Test Data file",
    );
  }
};

/**
 * อ่านไฟล์ Excel และคืน Workbook
 *
 * ถ้า ExcelJS เปิดไฟล์ไม่สำเร็จ
 * จะสร้าง Error ใหม่พร้อม Path และสาเหตุเดิม
 *
 * หลังเปิดไฟล์แล้วจะตรวจว่า
 * Workbook มีอย่างน้อย 1 Worksheet
 */
const readWorkbook = async (
  filePath: string,
  fileDescription: string,
): Promise<ExcelJS.Workbook> => {
  const workbook =
    new ExcelJS.Workbook();

  try {
    await workbook.xlsx.readFile(
      filePath,
    );
  } catch (
    error
  ) {
    /**
     * ดึงข้อความจาก Error
     *
     * ถ้าไม่ใช่ Error Object
     * ให้แปลงค่าเป็น string
     */
    const errorMessage =
      error instanceof Error
        ? error.message
        : String(
            error,
          );

    throw new Error(
      `Cannot read ${fileDescription}: ${filePath}. Reason: ${errorMessage}`,
    );
  }

  /**
   * Workbook ต้องมีอย่างน้อย 1 Worksheet
   */
  if (
    workbook.worksheets.length ===
    0
  ) {
    throw new Error(
      `${fileDescription} does not contain any worksheet: ${filePath}`,
    );
  }

  return workbook;
};

/**
 * รวมชื่อ Worksheet ทั้งหมดเป็นข้อความ
 *
 * ใช้แสดงใน Error Message
 * เมื่อไม่พบ Worksheet ที่ต้องการ
 *
 * ตัวอย่าง:
 * Sheet1, DS_PTX, Summary
 */
const getWorksheetNameList = (
  workbook: ExcelJS.Workbook,
): string => {
  return workbook.worksheets
    .map(
      (
        worksheet,
      ) => {
        return worksheet.name;
      },
    )
    .join(
      ", ",
    );
};

/**
 * ค้นหา Worksheet ด้วยชื่อ
 *
 * การเปรียบเทียบ
 * - ไม่สนตัวพิมพ์เล็กหรือใหญ่
 * - ไม่สนช่องว่างด้านหน้าและด้านหลัง
 *
 * ถ้าพบหลาย Sheet ที่ Normalize แล้วชื่อเหมือนกัน
 * จะคืน Sheet แรกที่พบ
 */
const findWorksheetByName = (
  workbook: ExcelJS.Workbook,
  worksheetName: string,
): ExcelJS.Worksheet | undefined => {
  const normalizedExpectedName =
    normalizeWorksheetName(
      worksheetName,
    );

  return workbook.worksheets.find(
    (
      worksheet,
    ) => {
      return (
        normalizeWorksheetName(
          worksheet.name,
        ) ===
        normalizedExpectedName
      );
    },
  );
};

/**
 * เลือก Worksheet ของ Report DS_FTX
 *
 * Report ต้องมี Sheet ชื่อ "DS_FTX"
 *
 * ระบบไม่เลือก Sheet แรกให้อัตโนมัติ
 * เพื่อป้องกันการอ่านผิด Sheet โดยไม่แจ้ง Error
 *
 * ถ้าไม่พบ จะแสดง
 * - ชื่อ Sheet ที่ต้องการ
 * - Path ของ Report
 * - รายชื่อ Sheet ที่มีอยู่
 */
const getReportWorksheet = (
  workbook: ExcelJS.Workbook,
  reportFilePath: string,
): ExcelJS.Worksheet => {
  const worksheet =
    findWorksheetByName(
      workbook,
      DEFAULT_REPORT_SHEET_NAME,
    );

  if (
    worksheet === undefined
  ) {
    throw new Error(
      [
        `Report worksheet "${DEFAULT_REPORT_SHEET_NAME}" not found`,
        `Report file: ${reportFilePath}`,
        `Available worksheets: ${getWorksheetNameList(workbook)}`,
      ].join(
        " | ",
      ),
    );
  }

  return worksheet;
};

/**
 * เลือก Worksheet ของ Test Data
 *
 * Test Data ใช้ Worksheet ลำดับแรก
 * โดยไม่ได้ตรวจสอบชื่อ Sheet
 *
 * เนื่องจากไฟล์ Test Data เดียวกัน
 * ถูกใช้ร่วมกันหลาย Report
 */
const getTestDataWorksheet = (
  workbook: ExcelJS.Workbook,
  testDataFilePath: string,
): ExcelJS.Worksheet => {
  const worksheet =
    workbook.worksheets[0];

  if (
    worksheet === undefined
  ) {
    throw new Error(
      `Test Data worksheet not found: ${testDataFilePath}`,
    );
  }

  return worksheet;
};

/**
 * ตรวจสอบว่าสร้าง ExpectedRow จาก Test Data
 * ได้อย่างน้อย 1 รายการ
 *
 * ถ้าได้ 0 รายการจะ Throw Error
 *
 * Error Message ระบุ Header Row แถว 5
 * เพราะ reconcileDsFtx() เรียก buildExpectedRows()
 * โดยใช้ค่าเริ่มต้นของ expected-row-builder.ts
 */
const validateExpectedRows = (
  expectedRowCount: number,
  worksheetName: string,
): void => {
  if (
    expectedRowCount ===
    0
  ) {
    throw new Error(
      [
        "No DS_FTX Test Data row was created",
        `Worksheet: ${worksheetName}`,
        "Please check Test Data header row 5 and data rows below it",
      ].join(
        " | ",
      ),
    );
  }
};

/**
 * สร้างผลสรุปการ Reconcile
 *
 * PASS และ FAIL จะนับจาก CompareResult ทีละรายการ
 *
 * ตัวอย่าง:
 * Test Data 1 แถวตรวจ Core Field 3 Field
 * ถ้าผ่านทั้งหมด:
 *
 * passedResultCount = 3
 *
 * ไม่ใช่:
 *
 * passedResultCount = 1
 */
const buildReconcileSummary = (
  reportFilePath: string,
  testDataFilePath: string,
  outputFilePath: string,
  reportWorksheet: ExcelJS.Worksheet,
  testDataWorksheet: ExcelJS.Worksheet,
  expectedRowCount: number,
  actualRowCount: number,
  results: CompareResult[],
): FtxReconcileSummary => {
  /**
   * เลือก CompareResult ที่เป็น PASS
   */
  const passedResults =
    results.filter(
      (
        result,
      ) => {
        return (
          result.status ===
          "PASS"
        );
      },
    );

  /**
   * เลือก CompareResult ที่เป็น FAIL
   */
  const failedResults =
    results.filter(
      (
        result,
      ) => {
        return (
          result.status ===
          "FAIL"
        );
      },
    );

  /**
   * เลือก CompareResult ที่มาจาก Exclusion Rule
   *
   * Code ปัจจุบันใช้ชื่อ Field:
   * "Exclusion Rule"
   */
  const exclusionResults =
    results.filter(
      (
        result,
      ) => {
        return (
          result.field ===
          "Exclusion Rule"
        );
      },
    );

  return {
    reportFilePath,

    testDataFilePath,

    outputFilePath,

    reportWorksheetName:
      reportWorksheet.name,

    testDataWorksheetName:
      testDataWorksheet.name,

    expectedRowCount,

    actualRowCount,

    totalResultCount:
      results.length,

    passedResultCount:
      passedResults.length,

    failedResultCount:
      failedResults.length,

    exclusionResultCount:
      exclusionResults.length,

    results,
  };
};

/**
 * แสดงผลสรุปการ Reconcile ใน Terminal
 *
 * จำนวน PASS, FAIL และ Exclusion
 * เป็นจำนวน CompareResult
 */
const printReconcileSummary = (
  summary: FtxReconcileSummary,
): void => {
  console.log(
    "========================================",
  );

  console.log(
    "DS_FTX Reconcile Completed",
  );

  console.log(
    `Report Sheet      : ${summary.reportWorksheetName}`,
  );

  console.log(
    `Test Data Sheet   : ${summary.testDataWorksheetName}`,
  );

  console.log(
    `Expected Rows     : ${summary.expectedRowCount}`,
  );

  console.log(
    `Actual Rows       : ${summary.actualRowCount}`,
  );

  console.log(
    `Total Results     : ${summary.totalResultCount}`,
  );

  console.log(
    `PASS Results      : ${summary.passedResultCount}`,
  );

  console.log(
    `FAIL Results      : ${summary.failedResultCount}`,
  );

  console.log(
    `Exclusion Results : ${summary.exclusionResultCount}`,
  );

  console.log(
    `Output File       : ${summary.outputFilePath}`,
  );

  console.log(
    "========================================",
  );
};

/**
 * ฟังก์ชันหลักสำหรับ Reconcile DS_FTX
 *
 * @param reportFilePath
 * Path ของ Report DS_FTX
 *
 * @param testDataFilePath
 * Path ของ Test Data
 *
 * @param outputFilePath
 * Path ของไฟล์ผลลัพธ์
 *
 * Output Folder ต้องมีอยู่ก่อน
 *
 * @returns
 * ผลสรุปการ Reconcile พร้อม CompareResult[]
 *
 * หมายเหตุ
 * แม้ CompareResult จะมี FAIL
 * ฟังก์ชันนี้ยังสร้างไฟล์และคืน Summary ตามปกติ
 */
export const reconcileDsFtx = async (
  reportFilePath: string,
  testDataFilePath: string,
  outputFilePath: string,
): Promise<FtxReconcileSummary> => {
  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 1: ตรวจสอบ Input และ Output Path
   * ----------------------------------------------------------
   */
  validateInputExcelFile(
    reportFilePath,
    "DS_FTX Report file",
  );

  validateInputExcelFile(
    testDataFilePath,
    "DS_FTX Test Data file",
  );

  validateOutputExcelFile(
    outputFilePath,
  );

  validateOutputDoesNotOverwriteInput(
    reportFilePath,
    testDataFilePath,
    outputFilePath,
  );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 2: อ่าน Report และ Test Data พร้อมกัน
   * ----------------------------------------------------------
   *
   * Promise.all() เริ่มอ่าน Workbook ทั้งสองไฟล์พร้อมกัน
   * และรอจนกว่าทั้งสองไฟล์จะอ่านเสร็จ
   *
   * ถ้าไฟล์ใดไฟล์หนึ่งอ่านไม่สำเร็จ
   * Promise.all() จะ Throw Error
   */
  const [
    reportWorkbook,
    testDataWorkbook,
  ] = await Promise.all([
    readWorkbook(
      reportFilePath,
      "DS_FTX Report file",
    ),

    readWorkbook(
      testDataFilePath,
      "DS_FTX Test Data file",
    ),
  ]);

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 3: เลือก Worksheet
   * ----------------------------------------------------------
   *
   * Report:
   * ค้นหา Sheet ชื่อ DS_FTX
   *
   * Test Data:
   * ใช้ Sheet แรก
   */
  const reportWorksheet =
    getReportWorksheet(
      reportWorkbook,
      reportFilePath,
    );

  const testDataWorksheet =
    getTestDataWorksheet(
      testDataWorkbook,
      testDataFilePath,
    );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 4: สร้าง ExpectedRow[] จาก Test Data
   * ----------------------------------------------------------
   *
   * ไม่ส่ง headerRowNumber เข้าไป
   * จึงใช้ค่าเริ่มต้นแถว 5
   */
  const expectedRows =
    buildExpectedRows(
      testDataWorksheet,
    );

  /**
   * ถ้าไม่สามารถสร้าง ExpectedRow ได้เลย
   * ให้หยุดการทำงาน
   */
  validateExpectedRows(
    expectedRows.length,
    testDataWorksheet.name,
  );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 5: สร้าง ActualRow[] จาก Report
   * ----------------------------------------------------------
   *
   * ไม่ส่ง headerRowNumber เข้าไป
   * จึงใช้ค่าเริ่มต้นแถว 1
   *
   * หมายเหตุ:
   * Code ไม่ได้ Throw Error เมื่อ actualRows เป็น Array ว่าง
   */
  const actualRows =
    buildActualRows(
      reportWorksheet,
    );

  /**
   * อ่าน Header ของ Report จากแถว 1
   *
   * ใช้สำหรับสร้าง Column ในไฟล์ผลลัพธ์
   */
  const reportHeaders =
    getReportHeaders(
      reportWorksheet,
    );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 6: เปรียบเทียบ Expected กับ Actual
   * ----------------------------------------------------------
   *
   * compareFtxRows() ตรวจตามลำดับ:
   * 1. Matching Key ใน Test Data ว่าง
   * 2. Exclusion Rule
   * 3. ไม่พบ Matching Key ใน Report
   * 4. Matching Key ซ้ำใน Report
   * 5. Core Field
   */
  const compareResults =
    compareFtxRows(
      expectedRows,
      actualRows,
    );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 7: สร้างไฟล์ Excel ผลลัพธ์
   * ----------------------------------------------------------
   *
   * Writer จะสร้าง Workbook ใหม่
   * จึงไม่แก้ไข Report หรือ Test Data ต้นทาง
   *
   * Output Folder ต้องมีอยู่ก่อน
   */
  await writeCompareResult(
    compareResults,
    expectedRows,
    actualRows,
    reportHeaders,
    outputFilePath,
  );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 8: สร้างผลสรุป
   * ----------------------------------------------------------
   */
  const summary =
    buildReconcileSummary(
      reportFilePath,
      testDataFilePath,
      outputFilePath,
      reportWorksheet,
      testDataWorksheet,
      expectedRows.length,
      actualRows.length,
      compareResults,
    );

  /**
   * ----------------------------------------------------------
   * ขั้นตอนที่ 9: แสดงผลสรุปใน Terminal
   * ----------------------------------------------------------
   */
  printReconcileSummary(
    summary,
  );

  /**
   * คืน Summary ให้ Script 3
   *
   * ผู้เรียกสามารถตรวจ failedResultCount
   * เพื่อพิจารณาว่าจะให้ Test Fail หรือไม่
   */
  return summary;
};