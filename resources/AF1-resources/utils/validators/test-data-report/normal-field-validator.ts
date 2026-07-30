/**
 * normal-field-validator.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้ตรวจสอบ Required Field หรือช่องข้อมูลบังคับทั่วไป
 * ที่ไม่ใช่ Field หลักของ Fee Group
 *
 * กติกาการตรวจสอบ
 *
 * 1. ถ้า Cell มีข้อมูล
 *    - ใส่พื้นหลังสีเขียวอ่อน
 *    - คงข้อมูลเดิมใน Cell ไว้
 *    - บันทึกสถานะ FOUND
 *
 * 2. ถ้า Cell ไม่มีข้อมูล
 *    - ใส่พื้นหลังสีแดง
 *    - เขียนข้อความ "โปรดกรอกข้อมูล" ลงใน Cell
 *    - บันทึกสถานะ EMPTY
 *    - ถือว่าพบข้อมูลไม่ถูกต้อง
 *
 * 3. ถ้าเป็น Fee Group Header
 *    - ข้ามการตรวจสอบในไฟล์นี้
 *    - ส่งไปตรวจด้วย fee-group-validator.ts
 *
 * 4. ถ้าหา Header ไม่พบ
 *    - ข้ามการตรวจสอบ Field นั้น
 *    - เพราะ Header ที่หายจะถูกตรวจโดย Header Validator
 *
 * คำศัพท์
 * - Normal Field  = ช่องข้อมูลทั่วไป
 * - Required Field = ช่องข้อมูลที่จำเป็นต้องกรอก
 * - Found         = พบข้อมูล
 * - Empty         = ไม่มีข้อมูล
 * - Invalid       = ข้อมูลไม่ถูกต้องหรือไม่ครบ
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  getCellText,
  normalizeValue,
} from "../shared/excel-cell.util";

import {
  COLORS,
  REQUIRED_MESSAGE,
} from "../shared/excel-style.util";

import {
  addFieldValidationResult,
} from "../shared/field-validation-sheet";

import {
  getCellByHeader,
  markRequiredCell,
  markSuccessCell,
} from "./field-helpers.util";

import {
  isFeeGroupHeader,
} from "./fee-group-validator";

/**
 * ตรวจสอบ Normal Required Field ทั้งหมดของข้อมูล 1 แถว
 *
 * ฟังก์ชันจะวนตรวจสอบ Field ทุกตัวที่อยู่ใน expectedHeaders
 * แต่จะข้าม Field ที่เป็น 3 ช่องหลักของ Fee Group ได้แก่
 *
 * - Fee Type N
 * - Fee Charge Account No. Type N
 * - Fee Amount Type N หรือ Fee Amount N
 *
 * @param row
 * แถวข้อมูล Test Data ที่กำลังตรวจสอบ
 *
 * @param headers
 * รายการ Header ที่อ่านได้จากไฟล์ Test Data
 *
 * @param expectedHeaders
 * รายการ Header ที่ต้องการตรวจสอบว่ามีข้อมูลหรือไม่
 *
 * @param resultSheet
 * Sheet "Field Validation" สำหรับบันทึกผลการตรวจสอบ
 *
 * @returns
 * true  = พบ Normal Field ว่างอย่างน้อย 1 ช่อง
 * false = ไม่พบ Normal Field ว่าง
 */
export const validateNormalRequiredFields = (
  row: ExcelJS.Row,
  headers: string[],
  expectedHeaders: string[],
  resultSheet: ExcelJS.Worksheet,
): boolean => {
  /**
   * ตัวแปรเก็บผลรวมของการตรวจสอบแถวปัจจุบัน
   *
   * เริ่มต้นเป็น false
   * เพราะยังไม่พบ Field ที่ไม่ผ่าน
   */
  let hasInvalidField = false;

  /**
   * วนตรวจสอบ Expected Header ทีละรายการ
   */
  expectedHeaders.forEach(
    (expectedHeader) => {
      /**
       * ตรวจสอบว่า Header ปัจจุบัน
       * เป็นหนึ่งใน 3 Field หลักของ Fee Group หรือไม่
       *
       * ถ้าเป็น Fee Group Header ให้ข้าม
       * เพราะจะมี Fee Group Validator ตรวจแยกต่างหาก
       */
      if (
        isFeeGroupHeader(
          expectedHeader,
        )
      ) {
        /**
         * return ภายใน forEach()
         * หมายถึงข้ามเฉพาะ Header ปัจจุบัน
         *
         * ไม่ได้หยุดฟังก์ชัน validateNormalRequiredFields()
         */
        return;
      }

      /**
       * ค้นหา Cell ที่อยู่ใต้ Expected Header
       *
       * getCellByHeader() รองรับการค้นหาผ่าน Alias
       *
       * ตัวอย่าง:
       * Expected Header = "Txn Date"
       * Actual Header   = "Transaction Date"
       *
       * ถ้ามี Alias รองรับ ระบบจะสามารถค้นหา Cell เจอ
       */
      const cell =
        getCellByHeader(
          row,
          headers,
          expectedHeader,
        );

      /**
       * ถ้าหา Header ไม่พบ จะไม่ได้ Cell กลับมา
       *
       * กรณีนี้ให้ข้าม Field ปัจจุบัน
       * เพราะการตรวจสอบ Header ที่หาย
       * เป็นหน้าที่ของ Header Validator
       *
       * หมายเหตุ:
       * กรณีนี้จะไม่เปลี่ยน hasInvalidField เป็น true
       */
      if (!cell) {
        return;
      }

      /**
       * อ่านค่าจาก Cell และปรับรูปแบบข้อความ
       *
       * getCellText()
       * = แปลงค่าภายใน Excel Cell ให้เป็นข้อความธรรมดา
       *
       * normalizeValue()
       * = ตัดช่องว่างและอักขระที่มองไม่เห็น
       */
      const value =
        normalizeValue(
          getCellText(
            cell,
          ),
        );

      /**
       * กรณี Cell ไม่มีข้อมูล
       *
       * normalizeValue() จะคืนข้อความว่าง ""
       * เมื่อ Cell ไม่มีค่าหรือมีแต่ช่องว่าง
       */
      if (value === "") {
        /**
         * ใส่พื้นหลังสีแดง
         * และเขียนข้อความ "โปรดกรอกข้อมูล" ลงใน Cell
         *
         * หมายเหตุ:
         * markRequiredCell() จะเปลี่ยน cell.value
         */
        markRequiredCell(
          cell,
        );

        /**
         * เพิ่มผลการตรวจสอบลงใน Sheet "Field Validation"
         *
         * Row
         * = หมายเลขแถวของข้อมูลต้นทาง
         *
         * Header
         * = Expected Header ที่กำลังตรวจสอบ
         *
         * Value
         * = ข้อความว่าง เพราะไม่พบข้อมูล
         *
         * Status
         * = EMPTY หมายถึงไม่มีข้อมูล
         *
         * Remark
         * = "โปรดกรอกข้อมูล"
         *
         * Color
         * = สีแดง
         */
        addFieldValidationResult(
          resultSheet,
          row.number,
          expectedHeader,
          "",
          "EMPTY",
          REQUIRED_MESSAGE,
          COLORS.RED,
        );

        /**
         * แสดงผลใน Console ว่าพบ Field ว่าง
         *
         * ตัวอย่าง:
         * 🔴 EMPTY FIELD | Row: 6, Header: Currency Id
         */
        console.log(
          `🔴 EMPTY FIELD | Row: ${row.number}, Header: ${expectedHeader}`,
        );

        /**
         * ระบุว่าแถวปัจจุบันมีข้อมูลไม่ถูกต้อง
         */
        hasInvalidField = true;

        /**
         * ข้ามการทำงานที่เหลือของ Header ปัจจุบัน
         * และไปตรวจ Header ตัวถัดไป
         */
        return;
      }

      /**
       * กรณี Cell มีข้อมูล
       *
       * ใส่พื้นหลังสีเขียวอ่อน
       * โดยไม่เปลี่ยนข้อมูลเดิมใน Cell
       */
      markSuccessCell(
        cell,
      );

      /**
       * เพิ่มผลการตรวจสอบลงใน Sheet "Field Validation"
       *
       * Status:
       * FOUND = พบข้อมูล
       *
       * Remark:
       * Field has value = Field มีข้อมูล
       *
       * Color:
       * FIELD_GREEN = สีเขียวอ่อน
       */
      addFieldValidationResult(
        resultSheet,
        row.number,
        expectedHeader,
        value,
        "FOUND",
        "Field has value",
        COLORS.FIELD_GREEN,
      );
    },
  );

  /**
   * คืนผลรวมหลังตรวจสอบ Expected Header ครบทุกตัว
   *
   * true
   * = พบ Field ว่างอย่างน้อย 1 ช่อง
   *
   * false
   * = Field ที่ตรวจสอบมีข้อมูลครบ
   *
   * หมายเหตุ:
   * ถ้าหา Header ไม่พบ Field นั้นจะถูกข้าม
   * และไม่ทำให้ผลลัพธ์เป็น true
   */
  return hasInvalidField;
};