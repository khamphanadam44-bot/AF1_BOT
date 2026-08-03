/**
 * fee-group-validator.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้ตรวจสอบข้อมูลภายใน Fee Group ของ Test Data
 * โดย Fee Group หลัก 1 กลุ่มประกอบด้วย 3 Field:
 *
 * 1. Fee Type N
 *    = ประเภทค่าธรรมเนียม
 *
 * 2. Fee Charge Account No. Type N
 *    = เลขที่บัญชีที่ใช้เรียกเก็บค่าธรรมเนียม
 *
 * 3. Fee Amount Type N หรือ Fee Amount N
 *    = จำนวนเงินค่าธรรมเนียม
 *
 * N หมายถึงหมายเลขลำดับของ Fee Group เช่น
 * - Fee Type 1
 * - Fee Type 2
 * - Fee Type 3
 *
 * จำนวน Fee Group ที่ต้องตรวจสอบ
 * รับมาจาก Config ของ Report ที่กำลังรัน
 *
 * กติกาการตรวจสอบมีทั้งหมด 4 กรณี
 *
 * Case 1: มีข้อมูลครบทั้ง 3 ช่อง
 * - ใส่สีเขียวทุกช่อง
 * - สถานะ FOUND
 * - Fee Group นี้ผ่าน
 *
 * Case 2: ไม่มีข้อมูลทั้ง 3 ช่อง
 * - ถือว่า Fee Group ลำดับนี้ไม่ได้ถูกใช้งาน
 * - ไม่แก้สีและไม่เขียนข้อความลง Cell
 * - Fee Group นี้ผ่าน
 *
 * Case 3: มีเฉพาะ Fee Amount
 * - Fee Type ใส่สีเหลือง
 * - Fee Charge Account ใส่สีเหลือง
 * - Fee Amount ใส่สีเขียว
 * - Fee Group นี้ไม่ผ่านและต้องตรวจสอบเพิ่มเติม
 *
 * Case 4: มีข้อมูลบางช่องในรูปแบบอื่น
 * - ช่องที่มีข้อมูลใส่สีเขียว
 * - ช่องที่ไม่มีข้อมูลใส่สีแดง
 * - Fee Group นี้ไม่ผ่าน
 *
 * หมายเหตุสำคัญ
 * ถ้าหา Header หลักของ Fee Group ไม่ครบทั้ง 3 ช่อง
 * ฟังก์ชันนี้จะข้าม Fee Group นั้น
 *
 * การตรวจสอบว่า Header ครบหรือไม่
 * เป็นหน้าที่ของ Header Validator
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

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

import {
  getFeeAmountHeader,
} from "../../../config/testdata-config";

/**
 * รูปแบบข้อมูลของ Fee Cell จำนวน 1 ช่อง
 *
 * ประกอบด้วย
 * - actualHeader = ชื่อ Header ที่พบจริงในไฟล์
 * - cell         = Cell ข้อมูลที่อยู่ใต้ Header นั้น
 */
type FeeCellInfo = {
  /**
   * ชื่อ Header ที่พบจริงใน Test Data
   *
   * ตัวอย่าง:
   * "Fee Type 1"
   */
  actualHeader: string;

  /**
   * Cell ที่มีข้อมูลของแถวที่กำลังตรวจสอบ
   */
  cell: ExcelJS.Cell;
};

/**
 * ตรวจสอบว่า Header ที่ได้รับมา
 * เป็นหนึ่งใน 3 Field หลักของ Fee Group หรือไม่
 *
 * Header ที่ถือว่าเป็น Fee Group:
 * - Fee Type N
 * - Fee Charge Account No. Type N
 * - Fee Amount Type N
 * - Fee Amount N
 *
 * ตัวอย่างที่คืนค่า true:
 * - Fee Type 1
 * - Fee Charge Account No. Type 2
 * - Fee Amount Type 1
 * - Fee Amount 3
 *
 * ตัวอย่างที่คืนค่า false:
 * - Fee Charge Type 1
 * - Bank Code Type 1
 * - Branch Code Type 1
 * - Fee Currency Type 1
 *
 * Field ที่ไม่ใช่ 3 ช่องหลักจะไม่ถูกตรวจสอบในฐานะ Fee Group
 * และสามารถนำไปตรวจสอบด้วย Normal Field Validator แทน
 *
 * @param header ชื่อ Header ที่ต้องการตรวจสอบ
 *
 * @returns
 * true  = เป็น Field หลักของ Fee Group
 * false = ไม่ใช่ Field หลักของ Fee Group
 */
export const isFeeGroupHeader = (
  header: string,
): boolean => {
  /**
   * ปรับรูปแบบ Header ก่อนตรวจสอบ เช่น
   * - เปลี่ยนเป็นตัวพิมพ์เล็ก
   * - รวมช่องว่าง
   * - ลบช่องว่างหัวและท้าย
   */
  const normalizedHeader =
    normalizeHeader(header);

  /**
   * ตรวจสอบรูปแบบ Header ด้วย Regular Expression
   *
   * \d+
   * หมายถึง ต้องลงท้ายด้วยตัวเลขอย่างน้อย 1 ตัว
   *
   * ^ และ $
   * หมายถึง ต้องตรงกันทั้งชื่อ Header
   * ไม่ใช่ตรงเพียงบางส่วน
   */
  return (
    /^fee type \d+$/.test(normalizedHeader) ||
    /^fee charge type \d+$/.test(normalizedHeader) ||
    /^fee charge account no\. type \d+$/.test(normalizedHeader) ||
    /^fee amount type \d+$/.test(normalizedHeader) ||
    /^fee amount \d+$/.test(normalizedHeader)
  );
};

/**
 * ค้นหา Cell ข้อมูลของ Fee Header จำนวน 1 ช่อง
 *
 * การทำงาน
 * 1. ค้นหา Array Index ของ Header
 * 2. ถ้าไม่พบ Header ให้คืน undefined
 * 3. ตรวจสอบว่า Actual Header เป็น Fee Group Header จริงหรือไม่
 * 4. ถ้าไม่ใช่ Fee Group Header ให้คืน undefined
 * 5. ถ้าถูกต้อง ให้คืนชื่อ Actual Header และ Cell ข้อมูล
 *
 * ตัวอย่าง
 *
 * expectedFeeHeader:
 * "Fee Type 1"
 *
 * ถ้าพบใน headers ที่ Array Index 5
 * จะอ่านข้อมูลจาก Excel Column 6
 *
 * สาเหตุที่ต้องบวก 1:
 * - Array เริ่มนับตำแหน่งจาก 0
 * - ExcelJS เริ่มนับ Column จาก 1
 *
 * @param row แถวข้อมูลที่กำลังตรวจสอบ
 * @param headers รายการ Header ที่อ่านจาก Excel
 * @param expectedFeeHeader ชื่อ Fee Header ที่ต้องการค้นหา
 *
 * @returns
 * - FeeCellInfo เมื่อพบ Fee Header และ Cell
 * - undefined เมื่อไม่พบหรือไม่ใช่ Fee Group Header
 */
const getFeeCellByHeader = (
  row: ExcelJS.Row,
  headers: string[],
  expectedFeeHeader: string,
): FeeCellInfo | undefined => {
  /**
   * ค้นหาตำแหน่ง Array Index ของ Header
   *
   * ถ้าไม่พบจะได้ค่า -1
   */
  const headerIndex =
    findHeaderColumnIndex(
      headers,
      expectedFeeHeader,
    );

  // ไม่พบ Header ที่ต้องการ
  if (headerIndex === -1) {
    return undefined;
  }

  /**
   * ดึงชื่อ Header จริงจากตำแหน่งที่ค้นพบ
   */
  const actualHeader =
    headers[headerIndex];

  /**
   * ป้องกันไม่ให้ Header ที่ไม่ใช่ 3 ช่องหลัก
   * ถูกนำมาตรวจสอบด้วย Fee Group Logic
   */
  if (!isFeeGroupHeader(actualHeader)) {
    return undefined;
  }

  /**
   * ส่งชื่อ Header และ Cell ข้อมูลกลับไป
   *
   * headerIndex + 1
   * ใช้แปลง Array Index ให้เป็น ExcelJS Column Number
   */
  return {
    actualHeader,
    cell: row.getCell(
      headerIndex + 1,
    ),
  };
};

/**
 * ตรวจสอบ Fee Group ทั้งหมดในแถวข้อมูล 1 แถว
 *
 * จำนวน Fee Group รับมาจาก Config ของ Report
 *
 * @param row แถวข้อมูล Test Data ที่กำลังตรวจสอบ
 * @param headers รายการ Header ทั้งหมดของ Test Data
 * @param resultSheet Sheet "Field Validation" สำหรับบันทึกผล
 * @param feeTypeCount จำนวน Fee Group ของ Report ที่กำลังตรวจ
 *
 * @returns
 * true  = มี Fee Group อย่างน้อย 1 กลุ่มที่ไม่ผ่าน
 * false = ไม่พบ Fee Group ที่ไม่ผ่าน
 *
 * หมายเหตุ
 * ถ้า Header ของกลุ่มใดไม่ครบ ฟังก์ชันนี้จะข้ามกลุ่มนั้น
 * และจะไม่เปลี่ยน hasInvalidField เป็น true
 */
export const validateFeeGroupFields = (
  row: ExcelJS.Row,
  headers: string[],
  resultSheet: ExcelJS.Worksheet,
  feeTypeCount: number,
): boolean => {
  let hasInvalidField = false;

  for (
    let feeIndex = 1;
    feeIndex <= feeTypeCount;
    feeIndex += 1
  ) {
    /**
     * ช่องที่ 1: Fee Type
     */
    const feeTypeInfo =
      getFeeCellByHeader(
        row,
        headers,
        `Fee Type ${feeIndex}`,
      );

    /**
     * ช่องที่ 2: Fee Charge Type
     */
    const feeChargeTypeInfo =
      getFeeCellByHeader(
        row,
        headers,
        `Fee Charge Type ${feeIndex}`,
      );

    /**
     * ช่องที่ 3: Fee Charge Account No.
     */
    const feeChargeAccountInfo =
      getFeeCellByHeader(
        row,
        headers,
        `Fee Charge Account No. Type ${feeIndex}`,
      );

    /**
     * ช่องที่ 4: Fee Amount
     */
    const feeAmountHeader =
      getFeeAmountHeader(
        feeIndex,
      );

    const feeAmountInfo =
      getFeeCellByHeader(
        row,
        headers,
        feeAmountHeader,
      );

    /**
     * ถ้า Header ของ Fee Group ไม่ครบ
     * ให้ Header Validator เป็นผู้แจ้ง Missing Header
     */
    if (
      !feeTypeInfo ||
      !feeChargeTypeInfo ||
      !feeChargeAccountInfo ||
      !feeAmountInfo
    ) {
      continue;
    }

    /**
     * Fee Group ที่ต้องตรวจทั้งหมด 4 ช่อง
     */
    const feeCells = [
      feeTypeInfo,
      feeChargeTypeInfo,
      feeChargeAccountInfo,
      feeAmountInfo,
    ];

    const feeTypeHasValue =
      !isCellEmpty(
        feeTypeInfo.cell,
      );

    const feeChargeTypeHasValue =
      !isCellEmpty(
        feeChargeTypeInfo.cell,
      );

    const feeChargeAccountHasValue =
      !isCellEmpty(
        feeChargeAccountInfo.cell,
      );

    const feeAmountHasValue =
      !isCellEmpty(
        feeAmountInfo.cell,
      );

    /**
     * นับจำนวนช่องที่มีข้อมูล
     *
     * ค่าที่เป็นไปได้:
     * 0 = ว่างทั้งหมด
     * 1 = มีข้อมูล 1 ช่อง
     * 2 = มีข้อมูล 2 ช่อง
     * 3 = มีข้อมูล 3 ช่อง
     * 4 = มีข้อมูลครบ
     */
    const valueCount = [
      feeTypeHasValue,
      feeChargeTypeHasValue,
      feeChargeAccountHasValue,
      feeAmountHasValue,
    ].filter(
      Boolean,
    ).length;

    /**
     * Case 1:
     * มีข้อมูลครบทั้ง 4 ช่อง
     *
     * ใส่สีเขียวทุกช่อง
     */
    if (valueCount === 4) {
      feeCells.forEach(
        (item) => {
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
            row.number,
            item.actualHeader,
            value,
            "FOUND",
            "Fee group is complete",
            COLORS.FIELD_GREEN,
          );
        },
      );

      continue;
    }

    /**
     * Case 2:
     * Fee Group ว่างทั้ง 4 ช่อง
     *
     * ใส่สีแดงและข้อความ
     * "โปรดกรอกข้อมูล" ทุกช่อง
     */
    if (valueCount === 0) {
      feeCells.forEach(
        (item) => {
          markRequiredCell(
            item.cell,
          );

          addFieldValidationResult(
            resultSheet,
            row.number,
            item.actualHeader,
            "",
            "EMPTY",
            REQUIRED_MESSAGE,
            COLORS.RED,
          );
        },
      );

      console.log(
        `🔴 EMPTY FEE GROUP | Row: ${row.number}, Fee Group: ${feeIndex}`,
      );

      hasInvalidField = true;

      continue;
    }

    /**
     * Case 3:
     * มีเฉพาะ Fee Amount
     *
     * Fee Type, Fee Charge Type และ
     * Fee Charge Account No. เป็นสีเหลือง
     *
     * Fee Amount เป็นสีเขียว
     */
    if (
      feeAmountHasValue &&
      !feeTypeHasValue &&
      !feeChargeTypeHasValue &&
      !feeChargeAccountHasValue
    ) {
      const missingFeeCells = [
        feeTypeInfo,
        feeChargeTypeInfo,
        feeChargeAccountInfo,
      ];

      missingFeeCells.forEach(
        (item) => {
          markCheckCell(
            item.cell,
          );

          addFieldValidationResult(
            resultSheet,
            row.number,
            item.actualHeader,
            "",
            "INCOMPLETE",
            CHECK_MESSAGE,
            COLORS.YELLOW,
          );
        },
      );

      const feeAmountValue =
        normalizeValue(
          getCellText(
            feeAmountInfo.cell,
          ),
        );

      markSuccessCell(
        feeAmountInfo.cell,
      );

      addFieldValidationResult(
        resultSheet,
        row.number,
        feeAmountInfo.actualHeader,
        feeAmountValue,
        "FOUND",
        "Fee Amount has value",
        COLORS.FIELD_GREEN,
      );

      console.log(
        `🟡 ONLY FEE AMOUNT HAS VALUE | Row: ${row.number}, Fee Group: ${feeIndex}`,
      );

      hasInvalidField = true;

      continue;
    }

    /**
     * Case 4:
     * มีข้อมูลบางช่องในรูปแบบอื่น
     *
     * ช่องที่มีข้อมูลเป็นสีเขียว
     * ช่องที่ว่างเป็นสีแดง
     */
    feeCells.forEach(
      (item) => {
        const value =
          normalizeValue(
            getCellText(
              item.cell,
            ),
          );

        if (value !== "") {
          markSuccessCell(
            item.cell,
          );

          addFieldValidationResult(
            resultSheet,
            row.number,
            item.actualHeader,
            value,
            "FOUND",
            "Field has value but fee group is incomplete",
            COLORS.FIELD_GREEN,
          );

          return;
        }

        markRequiredCell(
          item.cell,
        );

        addFieldValidationResult(
          resultSheet,
          row.number,
          item.actualHeader,
          "",
          "INCOMPLETE",
          REQUIRED_MESSAGE,
          COLORS.RED,
        );
      },
    );

    console.log(
      `🔴 INCOMPLETE FEE GROUP | Row: ${row.number}, Fee Group: ${feeIndex}`,
    );

    hasInvalidField = true;
  }

  return hasInvalidField;
};
