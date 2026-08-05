/**
 * ตรวจข้อมูล Fee Group ของ Test Data
 *
 * ใช้ไฟล์เดียวร่วมกัน แต่แยก Business Rule ตาม reportCode
 *
 * DS_PTX:
 * - ตรวจ 4 ช่องหลัก
 * - Fee Type
 * - Fee Charge Type
 * - Fee Charge Account No.
 * - Fee Amount
 * - ถ้าว่างทั้งกลุ่มให้เป็นสีแดง
 *
 * DS_LTX:
 * - ตรวจ 3 ช่องหลักตาม Logic เดิม
 * - Fee Type
 * - Fee Charge Account No.
 * - Fee Amount
 * - ถ้าว่างทั้งกลุ่มให้เป็นสีแดง
 */

import ExcelJS from "exceljs";

import type {
  TestDataReportCode,
} from "../../../config/testdata-config";

import {
  getFeeAmountHeader,
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

type FeeCellInfo = {
  actualHeader: string;
  cell: ExcelJS.Cell;
};

type FeeValidationRule = {
  includeFeeChargeType: boolean;
  emptyGroupIsRequired: boolean;
};

/** คืนกฎ Fee ตาม Report โดยไม่ให้ PTX กระทบ LTX */
/** คืนกฎการตรวจ Fee Group แยกตาม Report */
const getFeeValidationRule = (
  reportCode: TestDataReportCode,
): FeeValidationRule => {
  /**
   * PTX:
   * - ตรวจ Fee Type
   * - ตรวจ Fee Charge Type
   * - ตรวจ Fee Charge Account No.
   * - ตรวจ Fee Amount
   * - ถ้าว่างทั้งกลุ่มให้ใส่สีแดง
   */
  if (reportCode === "DS_PTX") {
    return {
      includeFeeChargeType: true,
      emptyGroupIsRequired: true,
    };
  }

  /**
   * LTX:
   * - ตรวจ Fee Type
   * - ไม่รวม Fee Charge Type ในกฎหลัก
   * - ตรวจ Fee Charge Account No.
   * - ตรวจ Fee Amount
   * - ถ้าว่างทั้งกลุ่มให้ใส่สีแดง
   */
  if (reportCode === "DS_LTX") {
    return {
      includeFeeChargeType: false,
      emptyGroupIsRequired: true,
    };
  }

  /**
   * Report อื่นไม่มีการบังคับ Fee Group
   */
  return {
    includeFeeChargeType: false,
    emptyGroupIsRequired: false,
  };
};

/** ตรวจว่า Header เป็นหนึ่งใน Field หลักของ Fee Group หรือไม่ */
export const isFeeGroupHeader = (
  header: string,
): boolean => {
  const normalizedHeader =
    normalizeHeader(
      header,
    );

  return (
    /^fee type \d+$/.test(normalizedHeader) ||
    /^fee charge type \d+$/.test(normalizedHeader) ||
    /^fee charge account no\. type \d+$/.test(normalizedHeader) ||
    /^fee amount type \d+$/.test(normalizedHeader) ||
    /^fee amount \d+$/.test(normalizedHeader)
  );
};

/** ค้นหา Fee Cell จากชื่อ Header */
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
    cell: row.getCell(
      headerIndex + 1,
    ),
  };
};

/** ใส่สีเขียวและบันทึกผลว่า Cell มีข้อมูล */
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

/** ใส่สีแดงและข้อความ "โปรดกรอกข้อมูล" */
const markFeeCellAsRequired = (
  resultSheet: ExcelJS.Worksheet,
  rowNumber: number,
  item: FeeCellInfo,
  status: "EMPTY" | "INCOMPLETE",
): void => {
  markRequiredCell(
    item.cell,
  );

  addFieldValidationResult(
    resultSheet,
    rowNumber,
    item.actualHeader,
    "",
    status,
    REQUIRED_MESSAGE,
    COLORS.RED,
  );
};

/** ใส่สีเหลืองและข้อความ "โปรดตรวจสอบข้อมูล" */
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
 * ตรวจ Fee Group ทุกชุดในข้อมูล 1 แถว
 *
 * @returns
 * true  = พบ Fee Group ที่ไม่ผ่าน
 * false = Fee Group ผ่านทั้งหมด
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

  for (
    let feeIndex = 1;
    feeIndex <= feeTypeCount;
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
     * ตรวจเฉพาะ PTX
     * LTX จะได้ค่า undefined
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
     * ถ้า Header หลักขาด ให้ Header Validator
     * เป็นผู้แจ้ง Missing Header และข้าม Fee Group นี้
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
     * PTX จะมี 4 Cell
     * LTX จะมี 3 Cell
     */
    const feeCells: FeeCellInfo[] = [
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

    const feeAmountHasValue =
      !isCellEmpty(
        feeAmountInfo.cell,
      );

    const valueCount =
      feeCells.filter(
        (item) =>
          !isCellEmpty(
            item.cell,
          ),
      ).length;

    /**
     * Case 1:
     * มีข้อมูลครบทุกช่องหลัก
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
     * Case 2:
     * ว่างทั้งกลุ่ม
     */
    if (valueCount === 0) {
      /**
       * ถ้า Report ไม่บังคับ Fee Group
       * และข้อมูลว่างทั้งกลุ่ม ให้ข้าม
       */
      if (
        !rule.emptyGroupIsRequired
      ) {
        continue;
      }

     /**
      * PTX และ LTX:
      * ถ้า Fee Group เป็นข้อมูลบังคับและว่างทั้งชุด
      * ให้ใส่สีแดงทุกช่องที่ Report กำหนดให้ตรวจ
      */
      feeCells.forEach(
        (item) => {
          markFeeCellAsRequired(
            resultSheet,
            row.number,
            item,
            "EMPTY",
          );
        },
      );

      console.log(
        `🔴 EMPTY FEE GROUP | Row: ${row.number}, Fee Group: ${feeIndex}`,
      );

      hasInvalidField =
        true;

      continue;
    }

    /**
     * Case 3:
     * มีเฉพาะ Fee Amount
     *
     * - Fee Amount เป็นสีเขียว
     * - ช่องหลักอื่นเป็นสีเหลือง
     */
    const nonAmountCells =
      feeCells.slice(
        0,
        -1,
      );

    const onlyFeeAmountHasValue =
      feeAmountHasValue &&
      nonAmountCells.every(
        (item) =>
          isCellEmpty(
            item.cell,
          ),
      );

    if (onlyFeeAmountHasValue) {
      nonAmountCells.forEach(
        (item) => {
          markFeeCellAsCheck(
            resultSheet,
            row.number,
            item,
          );
        },
      );

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
     * - ช่องที่มีข้อมูลเป็นสีเขียว
     * - ช่องที่ว่างเป็นสีแดง
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
          "INCOMPLETE",
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