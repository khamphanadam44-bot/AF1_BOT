/**
 * field-validator.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวควบคุมหลักสำหรับตรวจสอบข้อมูล Field ของ Test Data
 *
 * การทำงาน
 * 1. เลือก Worksheet แรกของ Workbook
 * 2. อ่าน Header จากแถวที่กำหนด
 * 3. สร้าง Sheet "Field Validation"
 * 4. หาแถวแรกและแถวสุดท้ายของข้อมูล
 * 5. ตรวจสอบ Normal Field ของทุก Report
 * 6. ตรวจสอบ Fee Group ตาม Config ของ Report
 * 7. สรุปว่าพบข้อมูลไม่ถูกต้องหรือไม่
 *
 * ค่าที่ส่งกลับ
 * - true  = พบข้อมูลไม่ถูกต้องอย่างน้อย 1 จุด
 * - false = ไม่พบข้อมูลไม่ถูกต้อง
 *
 * คำศัพท์
 * - Workbook   = ไฟล์ Excel ทั้งไฟล์
 * - Worksheet  = Sheet ภายในไฟล์ Excel
 * - Normal Field = Field ข้อมูลทั่วไป
 * - Fee Group  = กลุ่มข้อมูลค่าธรรมเนียม
 * - Invalid    = ข้อมูลไม่ถูกต้องหรือไม่ครบ
 * ------------------------------------------------------------------
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
 * ตรวจสอบว่าแถวนี้มีข้อมูลจริงอย่างน้อย 1 Cell หรือไม่
 *
 * ฟังก์ชันนี้ไม่ยึด Field ใดเป็นพิเศษ เช่น
 * - Test No.
 * - Matching Key
 * - Transaction ID
 *
 * ถ้าพบข้อมูลใน Cell ใด Cell หนึ่ง จะถือว่าแถวนั้นมีข้อมูล
 *
 * ค่าที่ไม่ถือว่าเป็นข้อมูล
 * - null
 * - undefined
 * - ข้อความว่าง ""
 * - ข้อความที่มีแต่ช่องว่าง
 * - ข้อความที่มีเฉพาะ Non-breaking Space
 *
 * Non-breaking Space
 * = ช่องว่างพิเศษที่อาจถูก Copy มาจากระบบอื่น
 *
 * @param row แถวที่ต้องการตรวจสอบ
 *
 * @returns
 * true  = แถวนี้มีข้อมูลอย่างน้อย 1 Cell
 * false = แถวนี้ไม่มีข้อมูล
 */
const rowHasAnyData = (
  row: ExcelJS.Row,
): boolean => {
  /**
   * เริ่มต้นถือว่าแถวนี้ยังไม่มีข้อมูล
   */
  let hasData = false;

  /**
   * วนตรวจสอบเฉพาะ Cell ที่ ExcelJS มองว่ามีค่า
   *
   * includeEmpty: false
   * = ไม่ต้องส่ง Cell ว่างเข้ามาใน Callback
   */
  row.eachCell(
    {
      includeEmpty: false,
    },
    (cell) => {
      // อ่านค่าเดิมจาก Cell
      const value =
        cell.value;

      /**
       * ถ้าค่าเป็น null หรือ undefined
       * ให้ข้าม Cell ปัจจุบัน
       */
      if (
        value === null ||
        value === undefined
      ) {
        return;
      }

      /**
       * ถ้าค่าเป็นข้อความ ให้ตรวจสอบเพิ่มเติมว่า
       * เป็นข้อความว่างหรือมีแต่ช่องว่างหรือไม่
       *
       * replace(/\u00A0/g, " ")
       * = เปลี่ยน Non-breaking Space เป็นช่องว่างปกติ
       *
       * trim()
       * = ตัดช่องว่างด้านหน้าและด้านหลัง
       */
      if (
        typeof value === "string" &&
        value
          .replace(/\u00A0/g, " ")
          .trim() === ""
      ) {
        return;
      }

      /**
       * ถ้าผ่านเงื่อนไขด้านบน
       * หมายความว่าพบข้อมูลจริงอย่างน้อย 1 Cell
       */
      hasData = true;
    },
  );

  // ส่งผลว่าแถวนี้มีข้อมูลหรือไม่
  return hasData;
};

/**
 * ค้นหาหมายเลขแถวสุดท้ายที่มีข้อมูลจริง
 *
 * วิธีค้นหา
 * - เริ่มจากแถวล่างสุดของ Worksheet
 * - ตรวจย้อนขึ้นมาด้านบนทีละแถว
 * - หยุดทันทีเมื่อพบแถวที่มีข้อมูล
 *
 * ไม่ใช้ actualRowCount เพราะ
 * actualRowCount คือ "จำนวนแถวที่มีข้อมูล"
 * ไม่ใช่ "หมายเลขของแถวสุดท้าย"
 *
 * ตัวอย่าง:
 * ถ้ามีข้อมูลอยู่ที่แถว 6, 7 และ 20
 *
 * actualRowCount อาจเป็น 3
 * แต่หมายเลขแถวสุดท้ายที่ต้องการคือ 20
 *
 * @param worksheet Worksheet ที่ต้องการค้นหา
 * @param firstDataRowNumber หมายเลขแถวแรกของข้อมูล
 *
 * @returns
 * - หมายเลขแถวสุดท้ายที่มีข้อมูล
 * - firstDataRowNumber - 1 เมื่อไม่พบข้อมูล
 */
const getLastDataRowNumber = (
  worksheet: ExcelJS.Worksheet,
  firstDataRowNumber: number,
): number => {
  /**
   * เริ่มตรวจสอบจาก worksheet.rowCount
   * แล้วย้อนขึ้นไปจนถึงแถวข้อมูลแรก
   */
  for (
    let rowNumber =
      worksheet.rowCount;
    rowNumber >=
      firstDataRowNumber;
    rowNumber -= 1
  ) {
    // ดึงแถวปัจจุบันจาก Worksheet
    const row =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * ถ้าแถวนี้มีข้อมูล
     * ให้คืนหมายเลขแถวทันที
     */
    if (
      rowHasAnyData(
        row,
      )
    ) {
      return rowNumber;
    }
  }

  /**
   * กรณีไม่มีข้อมูลอยู่หลังแถว Header
   *
   * ตัวอย่าง:
   * headerRowNumber = 5
   * firstDataRowNumber = 6
   *
   * ค่าที่คืนคือ 5
   *
   * ทำให้ Loop ตรวจสอบข้อมูลไม่เริ่มทำงาน
   * เพราะแถวแรก 6 มากกว่าแถวสุดท้าย 5
   */
  return firstDataRowNumber - 1;
};

/**
 * ตรวจสอบข้อมูล Field ทุกแถวใน Worksheet แรก
 *
 * ขั้นตอนหลัก
 * 1. เลือก Worksheet แรก
 * 2. อ่าน Header
 * 3. สร้าง Field Validation Sheet
 * 4. หาช่วงแถวข้อมูล
 * 5. ตรวจ Normal Field
 * 6. ตรวจ Fee Group ตามจำนวนที่กำหนดใน Config
 * 7. คืนผลรวมหลังตรวจครบทุกแถว
 *
 * @param workbook ไฟล์ Excel ที่ต้องการตรวจสอบ
 * @param expectedHeaders รายการ Normal Field ที่ต้องตรวจ
 * @param headerRowNumber หมายเลขแถว Header
 * @param reportCode รหัส Report ที่กำลังตรวจสอบ
 *
 * @returns Promise<boolean>
 * - true  = พบ Invalid Field
 * - false = ไม่พบ Invalid Field
 *
 * หมายเหตุ
 * ฟังก์ชันถูกประกาศเป็น async จึงคืนค่าเป็น Promise
 * แม้ปัจจุบันยังไม่มีคำสั่ง await ภายในฟังก์ชัน
 */
export const validateRequiredFields = async (
  workbook: ExcelJS.Workbook,
  expectedHeaders: string[],
  headerRowNumber: number,
  reportCode: TestDataReportCode,
): Promise<boolean> => {
  // แสดงหัวข้อการตรวจสอบใน Console
  console.log(
    "\n===== TEST DATA FIELD VALIDATION =====",
  );

  // แสดงรหัส Report ที่กำลังตรวจสอบ
  console.log(
    `Report Code : ${reportCode}`,
  );

  /**
   * เลือก Worksheet ลำดับแรกของ Workbook
   *
   * getWorksheet(1) หมายถึง Worksheet ลำดับที่ 1
   * ไม่ได้หมายถึง Sheet ที่ชื่อ "1"
   */
  const worksheet =
    workbook.getWorksheet(1);

  /**
   * ถ้าไม่พบ Worksheet แรก
   * โปรแกรมจะหยุดและแจ้ง Error
   */
  if (!worksheet) {
    throw new Error(
      "Worksheet not found",
    );
  }

  /**
   * อ่าน Header ทั้งหมดจากแถวที่กำหนด
   *
   * ตัวอย่าง:
   * headerRowNumber = 5
   * ระบบจะอ่าน Header จากแถวที่ 5
   */
  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * ลบ Sheet "Field Validation" เดิมถ้ามี
   * แล้วสร้าง Sheet ใหม่สำหรับเก็บผลการตรวจสอบ
   */
  const resultSheet =
    createFieldValidationSheet(
      workbook,
    );

  /**
   * กำหนดให้ข้อมูลเริ่มจากแถวถัดจาก Header
   *
   * ตัวอย่าง:
   * Header อยู่แถว 5
   * ข้อมูลจะเริ่มตรวจจากแถว 6
   */
  const firstDataRowNumber =
    headerRowNumber + 1;

  /**
   * ค้นหาแถวสุดท้ายที่มีข้อมูลจริง
   */
  const lastDataRowNumber =
    getLastDataRowNumber(
      worksheet,
      firstDataRowNumber,
    );

  /**
   * เก็บผลรวมของการตรวจสอบทั้งไฟล์
   *
   * เริ่มต้นเป็น false เพราะยังไม่พบข้อมูลผิด
   */
  let hasInvalidField =
    false;

  /**
   * วนตรวจสอบทุกแถวตั้งแต่
   * firstDataRowNumber ถึง lastDataRowNumber
   *
   * หมายเหตุสำคัญ
   * Loop นี้ไม่ได้เรียก rowHasAnyData() เพื่อข้ามแถวว่าง
   *
   * ดังนั้น ถ้ามีแถวว่างคั่นอยู่ระหว่างข้อมูล
   * แถวนั้นยังถูกส่งไปให้ Validator ตรวจสอบ
   */
  for (
    let rowNumber =
      firstDataRowNumber;
    rowNumber <=
      lastDataRowNumber;
    rowNumber += 1
  ) {
    /**
     * ดึงแถวปัจจุบันจาก Worksheet
     */
    const row =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * ตรวจสอบ Normal Required Field
     *
     * Normal Field จะถูกตรวจในทุก Report
     *
     * expectedHeaders ที่ส่งเข้ามา
     * ควรเป็นรายการที่ไม่มี 3 Field หลักของ Fee Group
     * เพราะ Fee Group มี Validator แยกต่างหาก
     *
     * ค่าที่ได้
     * true  = พบ Normal Field ที่ไม่ผ่าน
     * false = Normal Field ผ่านทั้งหมด
     */
    const hasInvalidNormalField =
      validateNormalRequiredFields(
        row,
        headers,
        expectedHeaders,
        resultSheet,
      );

    /**
     * ค่าเริ่มต้นของ Fee Group Validation
     *
     * Report ที่ไม่มี Fee Group ใน Config
     * จะไม่ตรวจ Fee และค่านี้จะคงเป็น false
     */
    let hasInvalidFeeGroup =
      false;

    /**
     * อ่านจำนวน Fee Group ตาม Config ของ Report
     *
     * ตัวอย่าง:
     * - DS_LTX = 5
     * - DS_PTX = 2
     * - DS_FTX / DS_FTU = 0
     */
    const feeTypeCount =
      getFeeTypeCount(
        reportCode,
      );

    /**
     * Report ที่มี Fee Group อย่างน้อย 1 ชุด
     * จะถูกส่งไปตรวจด้วย Fee Group Validator
     */
    if (feeTypeCount > 0) {
      hasInvalidFeeGroup =
        validateFeeGroupFields(
          row,
          headers,
          resultSheet,
          feeTypeCount,
        );
    }

    /**
     * ถ้า Normal Field หรือ Fee Group
     * มีส่วนใดส่วนหนึ่งไม่ผ่าน
     *
     * ให้กำหนดผลรวมของไฟล์เป็น true
     */
    if (
      hasInvalidNormalField ||
      hasInvalidFeeGroup
    ) {
      hasInvalidField =
        true;
    }
  }

  /**
   * คืนผลหลังจากตรวจสอบครบทุกแถวแล้ว
   *
   * true  = พบข้อมูลผิดอย่างน้อย 1 จุด
   * false = ไม่พบข้อมูลผิด
   */
  return hasInvalidField;
};
