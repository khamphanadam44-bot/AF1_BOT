/**
 * ตัวควบคุมการตรวจ Required Field ใน Test Data
 *
 * หน้าที่หลัก:
 * 1. อ่าน Header และช่วงข้อมูล
 * 2. ตรวจ Normal Required Field
 * 3. ตรวจ Fee Group ตาม Report
 * 4. สร้าง Sheet "Field Validation"
 */

import ExcelJS from "exceljs";

import type {
  TestDataReportCode,
} from "../../../config/testdata-config";

import {
  getFeeTypeCount,
} from "../../../config/testdata-helper";

import {
  getHeadersFromRow,
} from "../shared/excel-cell.util";

import {
  createFieldValidationSheet,
} from "../shared/field-validation-sheet";

import {
  validateNormalRequiredFields,
} from "./normal-field-validator";

import {
  validateFeeGroupFields,
} from "./fee-group-validator";

/** ตรวจว่าแถวมีข้อมูลจริงอย่างน้อย 1 Cell หรือไม่ */
const rowHasAnyData = (
  row: ExcelJS.Row,
): boolean => {
  let hasData = false;

  row.eachCell(
    {
      includeEmpty: false,
    },
    (cell) => {
      const value =
        cell.value;

      if (
        value === null ||
        value === undefined
      ) {
        return;
      }

      if (
        typeof value === "string" &&
        value
          .replace(/\u00A0/g, " ")
          .trim() === ""
      ) {
        return;
      }

      hasData = true;
    },
  );

  return hasData;
};

/** หาเลขแถวสุดท้ายที่มีข้อมูลจริง */
const getLastDataRowNumber = (
  worksheet: ExcelJS.Worksheet,
  firstDataRowNumber: number,
): number => {
  for (
    let rowNumber =
      worksheet.rowCount;
    rowNumber >=
      firstDataRowNumber;
    rowNumber -= 1
  ) {
    const row =
      worksheet.getRow(
        rowNumber,
      );

    if (
      rowHasAnyData(
        row,
      )
    ) {
      return rowNumber;
    }
  }

  return firstDataRowNumber - 1;
};

/**
 * ตรวจ Required Field ทุกแถวของ Test Data
 *
 * @returns
 * true  = พบข้อมูลไม่ครบอย่างน้อย 1 จุด
 * false = ไม่พบข้อมูลไม่ครบ
 */
export const validateRequiredFields = async (
  workbook: ExcelJS.Workbook,
  expectedHeaders: string[],
  headerRowNumber: number,
  reportCode: TestDataReportCode,
): Promise<boolean> => {
  console.log(
    "\n===== TEST DATA FIELD VALIDATION =====",
  );

  console.log(
    `Report Code : ${reportCode}`,
  );

  const worksheet =
    workbook.getWorksheet(1);

  if (!worksheet) {
    throw new Error(
      "Worksheet not found",
    );
  }

  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  const resultSheet =
    createFieldValidationSheet(
      workbook,
    );

  const firstDataRowNumber =
    headerRowNumber + 1;

  const lastDataRowNumber =
    getLastDataRowNumber(
      worksheet,
      firstDataRowNumber,
    );

  const feeTypeCount =
    getFeeTypeCount(
      reportCode,
    );

  let hasInvalidField =
    false;

  for (
    let rowNumber =
      firstDataRowNumber;
    rowNumber <=
      lastDataRowNumber;
    rowNumber += 1
  ) {
    const row =
      worksheet.getRow(
        rowNumber,
      );

    const hasInvalidNormalField =
      validateNormalRequiredFields(
        row,
        headers,
        expectedHeaders,
        resultSheet,
      );

    let hasInvalidFeeGroup =
      false;

    if (feeTypeCount > 0) {
      hasInvalidFeeGroup =
        validateFeeGroupFields(
          row,
          headers,
          resultSheet,
          feeTypeCount,
          reportCode,
        );
    }

    if (
      hasInvalidNormalField ||
      hasInvalidFeeGroup
    ) {
      hasInvalidField =
        true;
    }
  }

  return hasInvalidField;
};