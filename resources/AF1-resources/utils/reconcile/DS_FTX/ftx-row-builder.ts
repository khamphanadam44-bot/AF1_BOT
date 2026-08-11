/**
 * ftx-row-builder.ts
 * ------------------------------------------------------------------
 * เตรียม Expected Row และ Actual Row ของ DS_FTX ก่อนเปรียบเทียบ
 *
 * ลำดับการทำงาน:
 * 1. อ่าน Header จาก Worksheet
 * 2. แปลง Excel Row เป็น Object
 * 3. อ่านและ Normalize Matching Key
 * 4. สร้าง Expected Row จาก Test Data
 * 5. สร้าง Actual Row จาก AF1 Report
 *
 * กรณี Test No. ว่าง:
 * - ถ้ามี Transaction ID ให้ใช้ Transaction ID แสดงใน Test Script No.
 * - ถ้าไม่มีทั้ง Test No. และ Transaction ID ให้ใช้
 *   "Test Data Row N" เพื่อให้ Script 4 ค้นหา Test Data ต้นทางได้
 *
 * หมายเหตุ:
 * - Test Data Row N ใช้สำหรับอ้างอิงกลับไปยัง Test Data เท่านั้น
 * - ไม่ได้นำเลขแถวไปจับคู่กับ Report DS_FTX
 * - การจับคู่กับ Report ยังคงใช้ Transaction ID/ Reconcile ID
 *   เทียบกับ Ref. TX No.
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import type {
  ActualRow,
  ExpectedRow,
  ReportRow,
  TestDataRow,
} from "./ftx-types";

import {
  getCellText,
} from "../../validators/shared/excel-cell.util";

/** แถว Header เริ่มต้นของ Test Data */
export const DEFAULT_TEST_DATA_HEADER_ROW_NUMBER = 5;

/** แถว Header เริ่มต้นของ Report DS_FTX */
export const DEFAULT_REPORT_HEADER_ROW_NUMBER = 1;

/** Header ที่เก็บหมายเลข Test Case */
export const TEST_SCRIPT_NO_HEADER = "Test No.";

/** Header Matching Key ใน Test Data */
const TEST_DATA_MATCHING_KEY_HEADER =
  "Transaction ID/ Reconcile ID";

/** Header Matching Key ใน Report DS_FTX */
const REPORT_MATCHING_KEY_HEADER = "Ref. TX No.";

/** แปลงค่าเป็นข้อความและตัดช่องว่างหัวท้าย */
const normalizeText = (
  value: unknown,
): string => {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value).trim();
};

/**
 * Normalize Matching Key
 *
 * ไม่เปลี่ยนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 * และไม่ลบช่องว่างภายในข้อความ
 */
const normalizeMatchingKey = (
  value: unknown,
): string => {
  return normalizeText(value);
};

/** อ่าน Matching Key จาก Test Data */
const buildTestDataMatchingKey = (
  row: TestDataRow,
): string => {
  return normalizeMatchingKey(
    row[
      TEST_DATA_MATCHING_KEY_HEADER
    ],
  );
};

/** อ่าน Matching Key จาก Report DS_FTX */
const buildReportMatchingKey = (
  row: ReportRow,
): string => {
  return normalizeMatchingKey(
    row[
      REPORT_MATCHING_KEY_HEADER
    ],
  );
};

/**
 * สร้างค่าที่ใช้แสดงใน Column Test Script No.
 *
 * ลำดับความสำคัญ:
 * 1. Test No.
 * 2. Transaction ID/ Reconcile ID
 * 3. Test Data Row N
 *
 * ค่า Test Data Row N ใช้เพื่อให้ Script 4
 * สามารถย้อนกลับไปยังแถว Test Data ต้นทางได้
 */
const buildTestScriptNo = (
  testNo: unknown,
  transactionId: string,
  rowNumber: number,
): string => {
  const normalizedTestNo =
    normalizeText(testNo);

  if (normalizedTestNo !== "") {
    return normalizedTestNo;
  }

  if (transactionId !== "") {
    return transactionId;
  }

  return `Test Data Row ${rowNumber}`;
};

/**
 * แปลง Excel Row เป็น Object
 *
 * Column ที่ไม่มี Header จะไม่ถูกนำมาเก็บ
 */
const mapRowToObject = (
  row: ExcelJS.Row,
  headers: string[],
): Record<string, unknown> => {
  const result: Record<
    string,
    unknown
  > = {};

  headers.forEach(
    (header, index) => {
      if (header === "") {
        return;
      }

      result[header] =
        getCellText(
          row.getCell(
            index + 1,
          ),
        ).trim();
    },
  );

  return result;
};

/**
 * อ่านชื่อ Header โดยรักษาตำแหน่ง Column เดิม
 *
 * Header ที่ว่างจะถูกเก็บเป็นข้อความว่าง
 */
const readHeaders = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber: number,
): string[] => {
  const headerRow =
    worksheet.getRow(
      headerRowNumber,
    );

  const lastColumnNumber =
    Math.max(
      worksheet.actualColumnCount,
      headerRow.cellCount,
    );

  const headers: string[] = [];

  for (
    let columnNumber = 1;
    columnNumber <= lastColumnNumber;
    columnNumber += 1
  ) {
    headers.push(
      getCellText(
        headerRow.getCell(
          columnNumber,
        ),
      ).trim(),
    );
  }

  return headers;
};

/**
 * ตรวจว่า Object ของแถว
 * มีข้อมูลอย่างน้อยหนึ่ง Field หรือไม่
 */
const hasRowData = (
  data: Record<string, unknown>,
): boolean => {
  return Object.values(data).some(
    (value) =>
      normalizeText(value) !== "",
  );
};

/**
 * อ่าน Header ของ Report DS_FTX
 *
 * Export ฟังก์ชันนี้เพราะ Result Writer
 * ใช้รายการ Header เพื่อสร้าง Column ผลลัพธ์ด้วย
 */
export const getReportHeaders = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber =
    DEFAULT_REPORT_HEADER_ROW_NUMBER,
): string[] => {
  return readHeaders(
    worksheet,
    headerRowNumber,
  );
};

/**
 * สร้าง Expected Row จาก Test Data
 *
 * กติกา:
 * - แถวที่ไม่มีข้อมูลทุกช่องจะถูกข้าม
 * - แถวที่มีข้อมูล แต่ Transaction ID ว่าง
 *   จะยังถูกนำมาประมวลผล
 * - ถ้า Test No. ว่าง แต่ Transaction ID มีค่า
 *   จะใช้ Transaction ID แสดงใน Test Script No.
 * - ถ้า Test No. และ Transaction ID ว่างทั้งคู่
 *   จะใช้ "Test Data Row N" แสดงใน Test Script No.
 * - Matching Key ยังคงเป็น Transaction ID เท่านั้น
 * - Compare Engine จะสร้างผล FAIL แบบควบคุมได้
 *   เมื่อ Transaction ID ว่าง โดยไม่หยุด Script 3
 */
export const buildExpectedRows = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber =
    DEFAULT_TEST_DATA_HEADER_ROW_NUMBER,
): ExpectedRow[] => {
  if (headerRowNumber < 1) {
    throw new Error(
      `Invalid Test Data header row number: ${headerRowNumber}`,
    );
  }

  const headers =
    readHeaders(
      worksheet,
      headerRowNumber,
    );

  const expectedRows: ExpectedRow[] = [];

  for (
    let rowNumber =
      headerRowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const excelRow =
      worksheet.getRow(
        rowNumber,
      );

    const data =
      mapRowToObject(
        excelRow,
        headers,
      ) as TestDataRow;

    /**
     * ข้ามเฉพาะแถวที่ไม่มีข้อมูลทุกช่อง
     *
     * ถ้าแถวมีข้อมูล แต่ Test No. หรือ
     * Transaction ID ว่าง จะยังไม่ข้าม
     */
    if (!hasRowData(data)) {
      continue;
    }

    /**
     * Matching Key สำหรับจับคู่กับ Report
     *
     * ฝั่ง Test Data:
     * Transaction ID/ Reconcile ID
     *
     * ฝั่ง Report:
     * Ref. TX No.
     */
    const matchingKey =
      buildTestDataMatchingKey(
        data,
      );

    /**
     * ค่าที่ใช้แสดงใน Test Script No.
     *
     * 1. Test No.
     * 2. Transaction ID
     * 3. Test Data Row N
     */
    const testScriptNo =
      buildTestScriptNo(
        data[
          TEST_SCRIPT_NO_HEADER
        ],
        matchingKey,
        rowNumber,
      );

    expectedRows.push({
      rowNumber,
      testScriptNo,
      matchingKey,
      data,
    });
  }

  return expectedRows;
};

/**
 * สร้าง Actual Row จาก Report DS_FTX
 *
 * กติกา:
 * - แถวว่างจะถูกข้าม
 * - แถวที่มีข้อมูลแต่ Matching Key ว่างยังถูกเก็บไว้
 * - Matching Key อ่านจาก Header "Ref. TX No."
 */
export const buildActualRows = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber =
    DEFAULT_REPORT_HEADER_ROW_NUMBER,
): ActualRow[] => {
  if (headerRowNumber < 1) {
    throw new Error(
      `Invalid Report header row number: ${headerRowNumber}`,
    );
  }

  const headers =
    getReportHeaders(
      worksheet,
      headerRowNumber,
    );

  const actualRows: ActualRow[] = [];

  for (
    let rowNumber =
      headerRowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const excelRow =
      worksheet.getRow(
        rowNumber,
      );

    const data =
      mapRowToObject(
        excelRow,
        headers,
      ) as ReportRow;

    if (!hasRowData(data)) {
      continue;
    }

    const matchingKey =
      buildReportMatchingKey(
        data,
      );

    actualRows.push({
      rowNumber,
      matchingKey,
      data,
    });
  }

  return actualRows;
};