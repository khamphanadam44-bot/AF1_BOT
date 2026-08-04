/**
 * field-validator.ts
 * ------------------------------------------------------------
 * ตัวควบคุมการตรวจ Required Field ใน Test Data
 *
 * หน้าที่หลัก:
 * 1. อ่าน Header ของ Test Data
 * 2. หาช่วงแถวข้อมูลจริง
 * 3. ตรวจ Normal Required Field
 * 4. ตรวจ Fee Group ตาม Report
 * 5. บันทึกผลลง Sheet "Field Validation"
 *
 * การตรวจ Fee Group:
 * - DS_PTX และ DS_LTX:
 *   ตรวจจำนวน Fee Group จาก Header ในไฟล์จริง
 *
 * - DS_FTX และ DS_FTU:
 *   ไม่มีการตรวจ Fee Group
 * ------------------------------------------------------------
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

/**
 * ตรวจว่าแถวมีข้อมูลจริงอย่างน้อย 1 Cell หรือไม่
 *
 * คืนค่า:
 * true  = แถวนี้มีข้อมูล
 * false = แถวนี้เป็นแถวว่าง
 */
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

/**
 * หาเลขแถวสุดท้ายที่มีข้อมูลจริง
 *
 * เริ่มค้นหาจากแถวล่างสุดของ Worksheet
 * แล้วย้อนขึ้นมาจนกว่าจะพบแถวที่มีข้อมูล
 */
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

  /**
   * หากไม่พบข้อมูลหลังแถว Header
   * คืนค่าก่อนแถวข้อมูลแรก 1 แถว
   *
   * ทำให้ Loop ตรวจข้อมูลไม่ทำงาน
   */
  return firstDataRowNumber - 1;
};

/**
 * ตรวจ Required Field ทุกแถวของ Test Data
 *
 * @param workbook
 * Workbook ของ Test Data ที่กำลังตรวจ
 *
 * @param expectedHeaders
 * รายการ Header ที่ต้องตรวจแบบ Normal Required Field
 *
 * @param headerRowNumber
 * หมายเลขแถวที่เป็น Header ของ Test Data
 *
 * @param reportCode
 * Report ที่กำลังตรวจ เช่น DS_PTX หรือ DS_LTX
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

  /**
   * Test Data ใช้ Worksheet แรก
   */
  const worksheet =
    workbook.getWorksheet(1);

  if (!worksheet) {
    throw new Error(
      "Worksheet not found",
    );
  }

  /**
   * อ่าน Header จริงจากไฟล์ Test Data
   *
   * Header ชุดนี้จะถูกนำไปใช้:
   * - จับคู่ชื่อ Header
   * - ตรวจ Required Field
   * - ตรวจจำนวน Fee Group
   */
  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * สร้าง Sheet สำหรับบันทึกผลการตรวจ Field
   */
  const resultSheet =
    createFieldValidationSheet(
      workbook,
    );

  /**
   * ข้อมูลเริ่มต้นในแถวถัดจาก Header
   */
  const firstDataRowNumber =
    headerRowNumber + 1;

  /**
   * หาแถวสุดท้ายที่มีข้อมูลจริง
   * เพื่อไม่ให้ตรวจแถวว่างท้าย Worksheet
   */
  const lastDataRowNumber =
    getLastDataRowNumber(
      worksheet,
      firstDataRowNumber,
    );

  /**
   * ตรวจจำนวน Fee Group จาก Header จริง
   *
   * DS_PTX และ DS_LTX:
   * - คืนหมายเลข Fee Group สูงสุดที่พบ
   *
   * DS_FTX และ DS_FTU:
   * - คืนค่า 0
   */
  const feeTypeCount =
    getFeeTypeCount(
      reportCode,
      headers,
    );

  console.log(
    `Detected Fee Group Count : ${feeTypeCount}`,
  );

  let hasInvalidField =
    false;

  /**
   * ตรวจข้อมูลทีละแถว
   */
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

    /**
     * ตรวจ Normal Required Field
     *
     * ไม่รวม Fee Group เพราะ Fee Group
     * จะถูกตรวจด้วย Logic แยกต่างหาก
     */
    const hasInvalidNormalField =
      validateNormalRequiredFields(
        row,
        headers,
        expectedHeaders,
        resultSheet,
      );

    let hasInvalidFeeGroup =
      false;

    /**
     * ตรวจ Fee Group เฉพาะ Report
     * ที่มีจำนวน Fee Group มากกว่า 0
     *
     * DS_PTX และ DS_LTX:
     * - เข้าเงื่อนไขนี้
     *
     * DS_FTX และ DS_FTU:
     * - feeTypeCount เป็น 0
     * - ไม่เรียก validateFeeGroupFields()
     */
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

    /**
     * หากพบ Normal Field หรือ Fee Group ไม่ถูกต้อง
     * ให้กำหนดผลรวมของไฟล์ว่าเจอข้อมูลไม่ครบ
     */
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