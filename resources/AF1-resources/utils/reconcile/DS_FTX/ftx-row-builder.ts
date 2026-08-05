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
 * หน้าที่หลักของไฟล์นี้:
 * - สร้าง Expected Row จาก Test Data
 * - สร้าง Actual Row จาก Report DS_FTX
 * - อ่านและ Normalize Matching Key
 * - แปลงข้อมูล Excel Row เป็น Object
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
 * โดยไม่เปลี่ยนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
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
 * แปลง Excel Row เป็น Object
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

/** ตรวจว่า Object ของแถวมีข้อมูลอย่างน้อยหนึ่ง Field หรือไม่ */
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
 * - แถวว่างจะถูกข้าม
 * - แถวที่มีข้อมูลแต่ Matching Key ว่างยังถูกเก็บไว้
 * - Compare Engine จะเป็นผู้ตัดสินผลของ Matching Key ที่ว่าง
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

    if (!hasRowData(data)) {
      continue;
    }

    const testScriptNo =
      normalizeText(
        data[
          TEST_SCRIPT_NO_HEADER
        ],
      );

    const matchingKey =
      buildTestDataMatchingKey(
        data,
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