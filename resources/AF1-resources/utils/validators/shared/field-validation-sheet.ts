/**
 * field-validation-sheet.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้สำหรับสร้างและบันทึกผลการตรวจสอบข้อมูลลงใน Sheet
 * ที่มีชื่อว่า "Field Validation" (ผลการตรวจสอบ Field)
 *
 * ไฟล์นี้ใช้ร่วมกันระหว่าง
 * 1. Normal Field Validator
 *    หมายถึง การตรวจสอบ Field ข้อมูลทั่วไป
 *
 * 2. Fee Group Validator
 *    หมายถึง การตรวจสอบ Field ที่อยู่ในกลุ่มค่าธรรมเนียม
 *
 * การทำงานหลัก
 * 1. ลบ Sheet "Field Validation" เดิม ถ้ามีอยู่แล้ว
 * 2. สร้าง Sheet "Field Validation" ขึ้นมาใหม่
 * 3. สร้างหัวตารางสำหรับแสดงผลการตรวจสอบ
 * 4. เพิ่มผลการตรวจสอบลงในตารางทีละแถว
 * 5. ใส่สีในช่อง Value และ Status ตามผลการตรวจสอบ
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  applyBoldFill,
  applyFill,
  COLORS,
} from "./excel-style.util";

/**
 * ชื่อของ Sheet ที่ใช้เก็บผลการตรวจสอบ Field
 *
 * ประกาศเป็นค่าคงที่เพื่อให้ทุกจุดในไฟล์ใช้ชื่อเดียวกัน
 * ช่วยลดปัญหาการพิมพ์ชื่อ Sheet ผิด
 */
const FIELD_VALIDATION_SHEET_NAME = "Field Validation";

/**
 * สร้าง Sheet "Field Validation" สำหรับเก็บผลการตรวจสอบ
 *
 * การทำงาน
 * 1. ค้นหา Sheet "Field Validation" เดิมใน Workbook
 * 2. ถ้าพบ Sheet เดิม จะลบ Sheet นั้นออกก่อน
 * 3. สร้าง Sheet "Field Validation" ขึ้นมาใหม่
 * 4. กำหนดหัวตารางและความกว้างของแต่ละ Column
 * 5. ใส่พื้นหลังสีน้ำเงินและทำตัวหนาให้หัวตาราง
 * 6. ส่ง Worksheet ที่สร้างเสร็จแล้วกลับไปให้ส่วนอื่นใช้งาน
 *
 * เหตุผลที่ต้องลบ Sheet เดิม
 * เพื่อป้องกันไม่ให้ผลการตรวจสอบจากการ Run ครั้งก่อน
 * ค้างอยู่หรือปะปนกับผลการตรวจสอบครั้งล่าสุด
 *
 * @param workbook
 * Workbook หมายถึง ไฟล์ Excel ทั้งไฟล์
 *
 * @returns
 * Worksheet หมายถึง Sheet "Field Validation" ที่สร้างใหม่
 */
export const createFieldValidationSheet = (
  workbook: ExcelJS.Workbook,
): ExcelJS.Worksheet => {
  /**
   * ค้นหา Sheet เดิมจากชื่อ "Field Validation"
   *
   * ถ้าไม่พบ Sheet เดิม ค่า oldSheet จะเป็น undefined
   */
  const oldSheet = workbook.getWorksheet(
    FIELD_VALIDATION_SHEET_NAME,
  );

  /**
   * ถ้าพบ Sheet "Field Validation" เดิม
   * ให้ลบออกจาก Workbook ก่อนสร้าง Sheet ใหม่
   */
  if (oldSheet) {
    workbook.removeWorksheet(
      oldSheet.id,
    );
  }

  /**
   * สร้าง Sheet ใหม่ชื่อ "Field Validation"
   *
   * resultSheet จะเป็นตัวแทนของ Sheet ใหม่
   * ที่ใช้เพิ่มผลการตรวจสอบในขั้นตอนถัดไป
   */
  const resultSheet = workbook.addWorksheet(
    FIELD_VALIDATION_SHEET_NAME,
  );

  /**
   * กำหนด Column ของตารางผลการตรวจสอบ
   *
   * header = ข้อความที่แสดงอยู่บนหัวตาราง
   * key    = ชื่อที่ใช้ตอนเพิ่มข้อมูลด้วย Object
   * width  = ความกว้างของ Column
   *
   * ตารางจะประกอบด้วย
   * - Row    = หมายเลขแถวของข้อมูลที่ถูกตรวจสอบ
   * - Header = ชื่อ Field ที่ถูกตรวจสอบ
   * - Value  = ค่าข้อมูลที่พบใน Field
   * - Status = สถานะหรือผลการตรวจสอบ
   * - Remark = รายละเอียดเพิ่มเติม
   */
  resultSheet.columns = [
    {
      header: "Row",
      key: "row",
      width: 10,
    },
    {
      header: "Header",
      key: "header",
      width: 55,
    },
    {
      header: "Value",
      key: "value",
      width: 40,
    },
    {
      header: "Status",
      key: "status",
      width: 20,
    },
    {
      header: "Remark",
      key: "remark",
      width: 50,
    },
  ];

  /**
   * ดึงแถวที่ 1 ซึ่งเป็นแถวหัวตาราง
   *
   * eachCell หมายถึง ให้วนทำงานกับ Cell ทุกช่องในแถวนั้น
   *
   * จากนั้นเรียก applyBoldFill() เพื่อ
   * - ใส่พื้นหลังสีน้ำเงิน
   * - กำหนดสีตัวอักษรให้อ่านง่าย
   * - ทำตัวอักษรเป็นตัวหนา
   */
  resultSheet
    .getRow(1)
    .eachCell((cell) => {
      applyBoldFill(
        cell,
        COLORS.BLUE,
      );
    });

  /**
   * ส่ง Sheet ที่สร้างเสร็จแล้วกลับไป
   *
   * ส่วนที่เรียกใช้ฟังก์ชันนี้สามารถนำ resultSheet
   * ไปเพิ่มผลการตรวจสอบต่อได้
   */
  return resultSheet;
};

/**
 * เพิ่มผลการตรวจสอบ 1 รายการลงใน Sheet "Field Validation"
 *
 * แต่ละครั้งที่เรียกฟังก์ชันนี้ จะเพิ่มข้อมูลใหม่จำนวน 1 แถว
 *
 * ตัวอย่างข้อมูลที่ถูกเพิ่ม
 * - Row    : 6
 * - Header : Currency Id
 * - Value  : THB
 * - Status : PASS
 * - Remark : ข้อมูลถูกต้อง
 *
 * หลังจากเพิ่มข้อมูลแล้ว ฟังก์ชันจะใส่สีให้
 * - Column 3 หรือ Value
 * - Column 4 หรือ Status
 *
 * @param resultSheet
 * Sheet "Field Validation" ที่ต้องการเพิ่มผลลัพธ์
 *
 * @param rowNumber
 * หมายเลขแถวของข้อมูลต้นทางที่ถูกตรวจสอบ
 * ไม่ใช่หมายเลขแถวใน Sheet "Field Validation"
 *
 * @param header
 * ชื่อ Header หรือชื่อ Field ที่กำลังตรวจสอบ
 *
 * @param value
 * ค่าข้อมูลที่พบใน Field นั้น
 *
 * @param status
 * สถานะหรือผลการตรวจสอบ เช่น PASS, FAIL หรือ CHECK
 *
 * @param remark
 * คำอธิบายเพิ่มเติมเกี่ยวกับผลการตรวจสอบ
 *
 * @param color
 * รหัสสี ARGB ที่ต้องการใช้กับช่อง Value และ Status
 *
 * @returns
 * ไม่มีค่าที่ส่งกลับ เพราะฟังก์ชันนี้เพิ่มข้อมูลลงใน Sheet โดยตรง
 */
export const addFieldValidationResult = (
  resultSheet: ExcelJS.Worksheet,
  rowNumber: number,
  header: string,
  value: string,
  status: string,
  remark: string,
  color: string,
): void => {
  /**
   * เพิ่มผลการตรวจสอบลงใน Sheet จำนวน 1 แถว
   *
   * ชื่อ Property ใน Object จะตรงกับ key
   * ที่กำหนดไว้ใน resultSheet.columns
   *
   * row     → Column Row
   * header  → Column Header
   * value   → Column Value
   * status  → Column Status
   * remark  → Column Remark
   */
  const resultRow = resultSheet.addRow({
    row: rowNumber,
    header,
    value,
    status,
    remark,
  });

  /**
   * ใส่สีให้ Column 3 หรือ Column C
   * ซึ่งเป็นช่อง Value
   */
  applyFill(
    resultRow.getCell(3),
    color,
  );

  /**
   * ใส่สีให้ Column 4 หรือ Column D
   * ซึ่งเป็นช่อง Status
   */
  applyFill(
    resultRow.getCell(4),
    color,
  );
};