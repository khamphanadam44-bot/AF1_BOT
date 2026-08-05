/**
 * fee-group-validator.ts
 * ------------------------------------------------------------
 * ตรวจข้อมูล Fee Group ของ Test Data ใน Script 2
 *
 * DS_PTX ตรวจ 4 ช่องหลัก:
 * 1. Fee Type
 * 2. Fee Charge Type
 * 3. Fee Charge Account No.
 * 4. Fee Amount
 *
 * DS_LTX ตรวจ 3 ช่องหลัก:
 * 1. Fee Type
 * 2. Fee Charge Account No.
 * 3. Fee Amount
 *
 * DS_LTX:
 * - ตรวจ 3 ช่องหลักตาม Logic เดิม
 * - Fee Type
 * - Fee Charge Account No.
 * - Fee Amount
 * - ถ้าว่างทั้งกลุ่มให้เป็นสีแดง
 * กฎที่ใช้ร่วมกัน:
 * - ว่างทุกช่องในกลุ่ม: ข้าม ไม่ Highlight และไม่บันทึกผล
 * - มีข้อมูลครบ: Highlight สีเขียว
 * - มีเฉพาะ Fee Amount: Fee Amount สีเขียว ช่องอื่นสีเหลือง
 * - มีข้อมูลบางช่อง: ช่องที่มีข้อมูลสีเขียว ช่องที่ว่างสีแดง
 * ------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  getFeeAmountHeader,
} from "../../../config/testdata-config";

import type {
  TestDataReportCode,
} from "../../../config/testdata-config";

import {
  getCellText,
  isCellEmpty,
  normalizeHeader,
  normalizeValue,
} from "../shared/excel-cell.util";

import {
  CHECK_MESSAGE,
  COLORS,
  REQUIRED_MESSAGE,
} from "../shared/excel-style.util";

import {
  addFieldValidationResult,
} from "../shared/field-validation-sheet";

import {
  findHeaderColumnIndex,
  markCheckCell,
  markRequiredCell,
  markSuccessCell,
} from "./field-helpers.util";

/**
 * Cell และชื่อ Header จริงของ Field
 * ภายใน Fee Group
 */
type FeeCellInfo = {
  actualHeader: string;
  cell: ExcelJS.Cell;
};

/**
 * กฎที่แตกต่างกันของแต่ละ Report
 */
type FeeValidationRule = {
  includeFeeChargeType: boolean;
};

/**
 * คืนกฎการตรวจ Fee Group ตาม Report
 *
 * DS_PTX:
 * ใช้ Fee Charge Type เป็น Field หลัก
 *
 * DS_LTX:
 * ไม่ใช้ Fee Charge Type เป็น Field หลัก
 */
const getFeeValidationRule = (
  reportCode: TestDataReportCode,
): FeeValidationRule => {
  if (reportCode === "DS_PTX") {
    return {
      includeFeeChargeType: true,
    };
  }

  return {
    includeFeeChargeType: false,
  };
};

/**
 * ตรวจว่า Header เป็น Field หลัก
 * ของ Fee Group หรือไม่
 *
 * รองรับหมายเลข Fee Group หลายหลัก
 * เช่น 1, 5 และ 10
 */
export const isFeeGroupHeader = (
  header: string,
): boolean => {
  const normalizedHeader =
    normalizeHeader(
      header,
    );

  return (
    /^fee type \d+$/.test(
      normalizedHeader,
    ) ||
    /^fee charge type \d+$/.test(
      normalizedHeader,
    ) ||
    /^fee charge account no\. type \d+$/.test(
      normalizedHeader,
    ) ||
    /^fee amount type \d+$/.test(
      normalizedHeader,
    ) ||
    /^fee amount \d+$/.test(
      normalizedHeader,
    )
  );
};

/**
 * ค้นหา Cell ของ Fee
 * จากชื่อ Expected Header
 *
 * คืน undefined เมื่อ:
 * - ไม่พบ Header
 * - Header ที่พบไม่ใช่ Field หลักของ Fee Group
 */
const getFeeCellByHeader = (
  row: ExcelJS.Row,
  headers: string[],
  expectedFeeHeader: string,
): FeeCellInfo | undefined => {
  const headerIndex =
    findHeaderColumnIndex(
      headers,
      expectedFeeHeader,
    );

  if (headerIndex === -1) {
    return undefined;
  }

  const actualHeader =
    headers[headerIndex];

  if (
    !isFeeGroupHeader(
      actualHeader,
    )
  ) {
    return undefined;
  }

  return {
    actualHeader,

    cell:
      row.getCell(
        headerIndex + 1,
      ),
  };
};

/**
 * Highlight สีเขียว
 * และบันทึกว่า Cell มีข้อมูล
 */
const markFeeCellAsFound = (
  resultSheet: ExcelJS.Worksheet,
  rowNumber: number,
  item: FeeCellInfo,
  remark: string,
): void => {
  const value =
    normalizeValue(
      getCellText(
        item.cell,
      ),
    );

  markSuccessCell(
    item.cell,
  );

  addFieldValidationResult(
    resultSheet,
    rowNumber,
    item.actualHeader,
    value,
    "FOUND",
    remark,
    COLORS.FIELD_GREEN,
  );
};

/**
 * Highlight สีแดง
 * และใส่ข้อความ "โปรดกรอกข้อมูล"
 */
const markFeeCellAsRequired = (
  resultSheet: ExcelJS.Worksheet,
  rowNumber: number,
  item: FeeCellInfo,
): void => {
  markRequiredCell(
    item.cell,
  );

  addFieldValidationResult(
    resultSheet,
    rowNumber,
    item.actualHeader,
    "",
    "INCOMPLETE",
    REQUIRED_MESSAGE,
    COLORS.RED,
  );
};

/**
 * Highlight สีเหลือง
 * และใส่ข้อความ "โปรดตรวจสอบข้อมูล"
 */
const markFeeCellAsCheck = (
  resultSheet: ExcelJS.Worksheet,
  rowNumber: number,
  item: FeeCellInfo,
): void => {
  markCheckCell(
    item.cell,
  );

  addFieldValidationResult(
    resultSheet,
    rowNumber,
    item.actualHeader,
    "",
    "INCOMPLETE",
    CHECK_MESSAGE,
    COLORS.YELLOW,
  );
};

/**
 * ตรวจ Fee Group ทุกกลุ่ม
 * ใน Test Data หนึ่งแถว
 *
 * @param row
 * แถว Test Data ที่กำลังตรวจ
 *
 * @param headers
 * Header จริงของ Test Data
 *
 * @param resultSheet
 * Sheet สำหรับบันทึกผล Field Validation
 *
 * @param feeTypeCount
 * หมายเลข Fee Group สูงสุดที่ตรวจพบ
 *
 * @param reportCode
 * Report ที่กำลังตรวจ
 *
 * @returns
 * true:
 * พบ Fee Group ที่มีข้อมูลไม่ครบ
 *
 * false:
 * ทุกกลุ่มผ่าน หรือกลุ่มว่างทั้งหมดถูกข้าม
 */
export const validateFeeGroupFields = (
  row: ExcelJS.Row,
  headers: string[],
  resultSheet: ExcelJS.Worksheet,
  feeTypeCount: number,
  reportCode: TestDataReportCode,
): boolean => {
  const rule =
    getFeeValidationRule(
      reportCode,
    );

  let hasInvalidField =
    false;

  /**
   * ตรวจตั้งแต่ Fee Group 1
   * ถึงกลุ่มสูงสุดที่ตรวจพบจาก Header
   */
  for (
    let feeIndex = 1;

    feeIndex <=
      feeTypeCount;

    feeIndex += 1
  ) {
    /**
     * Fee Type
     */
    const feeTypeInfo =
      getFeeCellByHeader(
        row,
        headers,
        `Fee Type ${feeIndex}`,
      );

    /**
     * Fee Charge Type
     *
     * ตรวจเฉพาะ DS_PTX
     * DS_LTX จะได้ค่า undefined
     */
    const feeChargeTypeInfo =
      rule.includeFeeChargeType
        ? getFeeCellByHeader(
            row,
            headers,
            `Fee Charge Type ${feeIndex}`,
          )
        : undefined;

    /**
     * Fee Charge Account No.
     */
    const feeChargeAccountInfo =
      getFeeCellByHeader(
        row,
        headers,
        `Fee Charge Account No. Type ${feeIndex}`,
      );

    /**
     * Fee Amount
     *
     * Fee Group 1:
     * Fee Amount Type 1
     *
     * Fee Group 2 เป็นต้นไป:
     * Fee Amount 2
     * Fee Amount 3
     * ...
     */
    const feeAmountInfo =
      getFeeCellByHeader(
        row,
        headers,
        getFeeAmountHeader(
          feeIndex,
        ),
      );

    /**
     * หาก Header หลักขาด:
     *
     * ให้ Header Validator
     * เป็นผู้แจ้ง Missing Header
     *
     * แล้วข้ามการตรวจข้อมูล
     * ของ Fee Group นี้
     */
    if (
      !feeTypeInfo ||
      !feeChargeAccountInfo ||
      !feeAmountInfo ||
      (
        rule.includeFeeChargeType &&
        !feeChargeTypeInfo
      )
    ) {
      continue;
    }

    /**
     * สร้างรายการ Cell หลักของ Fee Group
     *
     * DS_PTX:
     * มี 4 Cell
     *
     * DS_LTX:
     * มี 3 Cell
     */
    const feeCells:
      FeeCellInfo[] = [
        feeTypeInfo,

        ...(
          feeChargeTypeInfo
            ? [
                feeChargeTypeInfo,
              ]
            : []
        ),

        feeChargeAccountInfo,

        feeAmountInfo,
      ];

    /**
     * นับจำนวน Cell ที่มีข้อมูล
     */
    const valueCount =
      feeCells.filter(
        (item) =>
          !isCellEmpty(
            item.cell,
          ),
      ).length;

    /**
     * Case 1:
     * ว่างทุกช่องใน Fee Group
     *
     * ไม่มีข้อมูลให้ตรวจ
     * จึงข้ามกลุ่มนี้ทันที
     *
     * ผลลัพธ์:
     * - ไม่ Highlight
     * - ไม่ใส่ข้อความ
     * - ไม่บันทึกใน Field Validation
     * - ไม่กำหนดผลเป็น Invalid
     */
    if (valueCount === 0) {
      continue;
    }

    /**
     * Case 2:
     * มีข้อมูลครบทุกช่องหลัก
     *
     * ผลลัพธ์:
     * - ทุกช่องเป็นสีเขียว
     * - บันทึก Status FOUND
     */
    if (
      valueCount ===
      feeCells.length
    ) {
      feeCells.forEach(
        (item) => {
          markFeeCellAsFound(
            resultSheet,
            row.number,
            item,
            "Fee group is complete",
          );
        },
      );

      continue;
    }

    /**
     * ตรวจว่า Fee Amount มีข้อมูลหรือไม่
     */
    const feeAmountHasValue =
      !isCellEmpty(
        feeAmountInfo.cell,
      );

    /**
     * Field ที่ไม่ใช่ Fee Amount
     *
     * DS_PTX:
     * - Fee Type
     * - Fee Charge Type
     * - Fee Charge Account No.
     *
     * DS_LTX:
     * - Fee Type
     * - Fee Charge Account No.
     */
    const nonAmountCells =
      feeCells.slice(
        0,
        -1,
      );

    /**
     * Case 3:
     * มีข้อมูลเฉพาะ Fee Amount
     *
     * เงื่อนไข:
     * - Fee Amount มีข้อมูล
     * - Field หลักอื่นว่างทั้งหมด
     */
    const onlyFeeAmountHasValue =
      feeAmountHasValue &&
      nonAmountCells.every(
        (item) =>
          isCellEmpty(
            item.cell,
          ),
      );

    if (onlyFeeAmountHasValue) {
      /**
       * Field หลักอื่นเป็นสีเหลือง
       * พร้อมข้อความ "โปรดตรวจสอบข้อมูล"
       */
      nonAmountCells.forEach(
        (item) => {
          markFeeCellAsCheck(
            resultSheet,
            row.number,
            item,
          );
        },
      );

      /**
       * Fee Amount มีข้อมูล
       * จึง Highlight สีเขียว
       */
      markFeeCellAsFound(
        resultSheet,
        row.number,
        feeAmountInfo,
        "Fee Amount has value",
      );

      console.log(
        `🟡 ONLY FEE AMOUNT HAS VALUE | Row: ${row.number}, Fee Group: ${feeIndex}`,
      );

      hasInvalidField =
        true;

      continue;
    }

    /**
     * Case 4:
     * มีข้อมูลบางช่องในรูปแบบอื่น
     *
     * ผลลัพธ์:
     * - ช่องที่มีข้อมูลเป็นสีเขียว
     * - ช่องที่ว่างเป็นสีแดง
     * - ช่องสีแดงใส่ข้อความ "โปรดกรอกข้อมูล"
     */
    feeCells.forEach(
      (item) => {
        if (
          !isCellEmpty(
            item.cell,
          )
        ) {
          markFeeCellAsFound(
            resultSheet,
            row.number,
            item,
            "Field has value but fee group is incomplete",
          );

          return;
        }

        markFeeCellAsRequired(
          resultSheet,
          row.number,
          item,
        );
      },
    );

    console.log(
      `🔴 INCOMPLETE FEE GROUP | Row: ${row.number}, Fee Group: ${feeIndex}`,
    );

    hasInvalidField =
      true;
  }

  return hasInvalidField;
};