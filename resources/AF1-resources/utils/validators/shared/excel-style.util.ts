/**
 * excel-style.util.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * รวมค่าคงที่และฟังก์ชันสำหรับจัดรูปแบบ Cell ในไฟล์ Excel เช่น
 * - กำหนดสีพื้นหลังของ Cell
 * - กำหนดสีตัวอักษร
 * - ทำตัวอักษรให้เป็นตัวหนา
 * - คืนรูปแบบของ Header ให้เป็นรูปแบบเริ่มต้น
 *
 * กติกาการใช้สี
 * - เขียว     = ข้อมูลถูกต้อง หรือมีข้อมูล
 * - แดง       = Field ที่จำเป็นต้องมีข้อมูล แต่พบว่าไม่มีข้อมูล
 * - เหลือง    = ใช้เฉพาะ Fee Group ในกรณีที่ต้องตรวจสอบเพิ่มเติม
 * - น้ำเงิน   = ใช้เป็นสีหัวตารางใน Sheet ผลลัพธ์
 * - ขาว/ดำ    = ใช้เป็นสีพื้นฐานของ Cell และตัวอักษร
 *
 * หมายเหตุ
 * - ห้ามใช้สีเหลืองกับ Normal Field
 * - รหัสสีใช้รูปแบบ ARGB จำนวน 8 ตัวอักษร
 * - ตัวอักษร 2 ตัวแรกคือค่าความทึบของสี เช่น FF หมายถึงทึบ 100%
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

/**
 * รวมรหัสสีทั้งหมดที่ใช้ภายในไฟล์ Excel
 *
 * ใช้ `as const` เพื่อกำหนดให้ค่าของสีเป็นค่าคงที่
 * และช่วยป้องกันไม่ให้ Code ส่วนอื่นแก้ไขค่าของสีโดยไม่ตั้งใจ
 */
export const COLORS = {
  /** เขียวเข้ม ใช้ Highlight Header ที่อยู่ใน Requirement */
  GREEN: "FF92D050",

  /** เขียวอ่อน ใช้ Highlight ค่าข้อมูลของ Field ที่ตรวจสอบแล้วว่าถูกต้อง */
  FIELD_GREEN: "FFC6EFCE",

  /** แดง ใช้กับ Required Field ที่ไม่มีข้อมูล */
  RED: "FFFF0000",

  /** เหลือง ใช้เฉพาะ Fee Group ที่ต้องตรวจสอบเพิ่มเติม */
  YELLOW: "FFFFFF00",

  /** น้ำเงิน ใช้เป็นสีพื้นหลังของหัวตารางใน Sheet ผลลัพธ์ */
  BLUE: "FF4472C4",

  /** ขาว ใช้เป็นสีพื้นหลังเริ่มต้น */
  WHITE: "FFFFFFFF",

  /** ดำ ใช้เป็นสีตัวอักษรบนพื้นหลังสีอ่อน */
  BLACK: "FF000000",
} as const;

/**
 * ข้อความที่ใส่ลงใน Cell
 * เมื่อ Field เป็น Required Field แต่ไม่มีข้อมูล
 */
export const REQUIRED_MESSAGE = "โปรดกรอกข้อมูล";

/**
 * ข้อความที่ใส่ลงใน Cell
 * เมื่อ Fee Group อยู่ใน Case 3 และต้องตรวจสอบข้อมูลเพิ่มเติม
 */
export const CHECK_MESSAGE = "โปรดตรวจสอบข้อมูล";

/**
 * สร้างสำเนา Style เดิมของ Cell ก่อนทำการแก้ไข
 *
 * เหตุผลที่ต้อง Clone
 * ExcelJS อาจให้หลาย Cell ใช้ Style Object ตัวเดียวกัน
 * ถ้าแก้ไข Style โดยตรง อาจทำให้ Style ของ Cell อื่นเปลี่ยนตามไปด้วย
 *
 * การทำงาน
 * 1. อ่าน Style ปัจจุบันของ Cell
 * 2. แปลง Style เป็น JSON
 * 3. สร้าง Object ใหม่จาก JSON
 * 4. นำ Style Object ใหม่กลับไปใส่ใน Cell
 *
 * @param cell Cell ที่ต้องการสร้างสำเนา Style
 */
const cloneCellStyle = (cell: ExcelJS.Cell): void => {
  cell.style = JSON.parse(JSON.stringify(cell.style ?? {}));
};

/**
 * ใส่สีพื้นหลังให้ Cell และปรับสีตัวอักษรให้อ่านง่าย
 *
 * กติกาสีตัวอักษร
 * - ถ้าพื้นหลังเป็นสีอ่อน ให้ใช้ตัวอักษรสีดำ
 * - ถ้าพื้นหลังเป็นสีเข้ม ให้ใช้ตัวอักษรสีขาว
 *
 * สีที่ถือว่าเป็นพื้นหลังสีอ่อน
 * - สีเหลือง
 * - สีเขียวอ่อน
 * - สีขาว
 *
 * ตัวอย่าง
 * applyFill(cell, COLORS.RED);
 * ผลลัพธ์: Cell พื้นหลังสีแดง และตัวอักษรสีขาว
 *
 * applyFill(cell, COLORS.FIELD_GREEN);
 * ผลลัพธ์: Cell พื้นหลังสีเขียวอ่อน และตัวอักษรสีดำ
 *
 * @param cell Cell ที่ต้องการใส่สี
 * @param color รหัสสี ARGB ที่ต้องการใช้เป็นสีพื้นหลัง
 */
export const applyFill = (cell: ExcelJS.Cell, color: string): void => {
  // Clone Style ก่อนแก้ไข เพื่อไม่ให้กระทบ Cell อื่น
  cloneCellStyle(cell);

  // กำหนดสีพื้นหลังเป็นสีทึบ
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: color },
  };

  /**
   * ตรวจสอบว่าสีพื้นหลังเป็นสีอ่อนหรือไม่
   *
   * ถ้าเป็นสีอ่อน จะใช้ตัวอักษรสีดำ
   * ถ้าเป็นสีอื่น จะใช้ตัวอักษรสีขาว
   */
  const isLightBackground =
    color === COLORS.YELLOW || color === COLORS.FIELD_GREEN || color === COLORS.WHITE;

  // เก็บ Font Style เดิมไว้ และเปลี่ยนเฉพาะสีตัวอักษร
  cell.font = {
    ...(cell.font ?? {}),
    color: { argb: isLightBackground ? COLORS.BLACK : COLORS.WHITE },
  };
};

/**
 * ใส่สีพื้นหลังและทำตัวอักษรให้เป็นตัวหนา
 *
 * เหมาะสำหรับ
 * - หัวตารางใน Sheet ผลลัพธ์
 * - Header Row ที่ต้องการ Highlight
 *
 * การทำงาน
 * 1. เรียก applyFill() เพื่อใส่สีพื้นหลังและสีตัวอักษร
 * 2. กำหนดตัวอักษรให้เป็นตัวหนา
 *
 * ตัวอย่าง
 * applyBoldFill(cell, COLORS.BLUE);
 * ผลลัพธ์: Cell พื้นหลังสีน้ำเงิน ตัวอักษรสีขาวและเป็นตัวหนา
 *
 * @param cell Cell ที่ต้องการจัดรูปแบบ
 * @param color รหัสสี ARGB ที่ต้องการใช้เป็นสีพื้นหลัง
 */
export const applyBoldFill = (cell: ExcelJS.Cell, color: string): void => {
  // ใส่สีพื้นหลังและกำหนดสีตัวอักษร
  applyFill(cell, color);

  // เก็บ Font Style เดิมไว้ และเพิ่มตัวหนา
  cell.font = {
    ...(cell.font ?? {}),
    bold: true,
  };
};

/**
 * คืนรูปแบบ Header ให้เป็น Style เริ่มต้น
 *
 * Style เริ่มต้นประกอบด้วย
 * - พื้นหลังสีขาว
 * - ตัวอักษรสีดำ
 * - ตัวอักษรไม่เป็นตัวหนา
 *
 * ใช้ในกรณี เช่น
 * Header ที่อยู่ในไฟล์ Report แต่ไม่ได้อยู่ใน Requirement
 *
 * @param cell Header Cell ที่ต้องการคืนค่า Style
 */
export const applyDefaultHeaderStyle = (cell: ExcelJS.Cell): void => {
  // Clone Style ก่อนแก้ไข เพื่อไม่ให้กระทบ Cell อื่น
  cloneCellStyle(cell);

  // กำหนดพื้นหลังของ Cell ให้เป็นสีขาว
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: COLORS.WHITE },
  };

  // กำหนดตัวอักษรให้เป็นสีดำและไม่เป็นตัวหนา
  cell.font = {
    ...(cell.font ?? {}),
    bold: false,
    color: { argb: COLORS.BLACK },
  };
};