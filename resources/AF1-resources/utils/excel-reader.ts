/**
 * excel-reader.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * อ่านไฟล์ Excel และนำข้อมูลจาก Worksheet แรก
 * มาแปลงเป็น Array ของ JavaScript Object
 *
 * ใช้ใน Script 1 เพื่ออ่านจำนวนแถวของ
 * - Downloaded Report
 * - Test Data
 *
 * หมายเหตุสำคัญ
 * 1. Code ปัจจุบันกำหนดให้ Header อยู่ที่แถว 1 เท่านั้น
 * 2. ถ้า Test Data มี Header อยู่ที่แถว 5
 *    จำนวนแถวที่อ่านได้อาจไม่ใช่จำนวน Data Row จริง
 * 3. Code ไม่ได้สร้างข้อความ JSON ด้วย JSON.stringify()
 *    แต่สร้าง Array ของ Object ที่มีรูปแบบคล้าย JSON
 * 4. ใช้เฉพาะ Worksheet แรกของไฟล์ Excel
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

/**
 * อ่านข้อมูลจาก Worksheet แรกของไฟล์ Excel
 * และแปลงแต่ละแถวเป็น JavaScript Object
 *
 * ตัวอย่างไฟล์ Excel:
 *
 * | Transaction ID | Currency |
 * |----------------|----------|
 * | TXN001         | THB      |
 * | TXN002         | USD      |
 *
 * ผลลัพธ์:
 *
 * [
 *   {
 *     "Transaction ID": "TXN001",
 *     "Currency": "THB"
 *   },
 *   {
 *     "Transaction ID": "TXN002",
 *     "Currency": "USD"
 *   }
 * ]
 *
 * @param filePath
 * Path ของไฟล์ Excel ที่ต้องการอ่าน
 *
 * @returns
 * Promise ที่คืน Array ของ Object
 *
 * - Object 1 ตัวแทนข้อมูล Excel 1 แถว
 * - Property Name มาจาก Header แถวที่ 1
 * - Property Value มาจากค่าของ Cell
 * - ถ้าไม่พบ Worksheet จะคืน Array ว่าง
 */
export async function readExcel(
  filePath: string,
): Promise<Record<string, unknown>[]> {
  /**
   * สร้าง Workbook ใหม่ในหน่วยความจำ
   *
   * ตอนนี้ยังไม่มีข้อมูลอยู่ภายใน Workbook
   */
  const workbook =
    new ExcelJS.Workbook();

  /**
   * อ่านไฟล์ Excel จาก filePath
   * และนำข้อมูลเข้ามาเก็บใน Workbook
   *
   * await หมายถึงรอจนกว่า ExcelJS
   * จะอ่านไฟล์เสร็จก่อนทำขั้นตอนถัดไป
   */
  await workbook.xlsx.readFile(
    filePath,
  );

  /**
   * เลือก Worksheet ลำดับแรกของไฟล์
   *
   * getWorksheet(1)
   * หมายถึง Worksheet ลำดับที่ 1
   * ไม่ได้หมายถึง Worksheet ที่ชื่อว่า "1"
   */
  const worksheet =
    workbook.getWorksheet(1);

  /**
   * ถ้าไม่พบ Worksheet
   * ให้คืน Array ว่างโดยไม่ Throw Error
   */
  if (!worksheet) {
    return [];
  }

  /**
   * Array สำหรับเก็บชื่อ Header
   *
   * ตำแหน่งใน Array จะสัมพันธ์กับ Column ของ Excel
   *
   * ตัวอย่าง:
   * headers[0] = Header ของ Column A
   * headers[1] = Header ของ Column B
   * headers[2] = Header ของ Column C
   */
  const headers: string[] = [];

  /**
   * อ่าน Header จากแถวที่ 1
   *
   * หมายเหตุ:
   * เลข 1 ถูกกำหนดตายตัวอยู่ใน Code
   * ฟังก์ชันนี้จึงไม่รองรับ Header ที่อยู่แถวอื่นโดยตรง
   *
   * includeEmpty: true
   * หมายถึงให้วนผ่าน Cell ว่างภายในช่วงของแถวด้วย
   */
  worksheet
    .getRow(1)
    .eachCell(
      {
        includeEmpty: true,
      },
      (
        cell,
        columnNumber,
      ) => {
        /**
         * เก็บชื่อ Header ลงใน Array
         *
         * ExcelJS เริ่มนับ Column จาก 1:
         * - Column A = 1
         * - Column B = 2
         *
         * Array เริ่มนับ Index จาก 0:
         * - ตำแหน่งแรก = 0
         * - ตำแหน่งที่สอง = 1
         *
         * จึงต้องใช้ columnNumber - 1
         */
        headers[
          columnNumber - 1
        ] = String(
          cell.value ?? "",
        );
      },
    );

  /**
   * Array สำหรับเก็บข้อมูลทุกแถว
   *
   * Object 1 ตัวจะแทนข้อมูล Excel 1 แถว
   */
  const rows: Record<
    string,
    unknown
  >[] = [];

  /**
   * วนอ่านข้อมูลแต่ละแถวของ Worksheet
   *
   * eachRow() ที่ไม่ได้กำหนด includeEmpty: true
   * จะไม่วนผ่านแถวที่ ExcelJS มองว่าไม่มีข้อมูล
   */
  worksheet.eachRow(
    (
      row,
      rowNumber,
    ) => {
      /**
       * ข้ามแถวที่ 1
       * เพราะ Code ถือว่าแถวที่ 1 เป็น Header
       *
       * return ตรงนี้ข้ามเฉพาะแถวปัจจุบัน
       * ไม่ได้จบฟังก์ชัน readExcel()
       */
      if (rowNumber === 1) {
        return;
      }

      /**
       * Object สำหรับเก็บข้อมูลของแถวปัจจุบัน
       *
       * ตัวอย่าง:
       *
       * {
       *   "Transaction ID": "TXN001",
       *   "Currency": "THB"
       * }
       */
      const rowData: Record<
        string,
        unknown
      > = {};

      /**
       * วนอ่าน Cell แต่ละ Column ของแถวปัจจุบัน
       *
       * includeEmpty: true
       * หมายถึงรวม Cell ว่างภายในช่วงของแถวด้วย
       */
      row.eachCell(
        {
          includeEmpty: true,
        },
        (
          cell,
          columnNumber,
        ) => {
          /**
           * ดึงชื่อ Property จาก Header
           *
           * headers[columnNumber - 1]
           * ใช้แปลง Excel Column Number
           * ให้ตรงกับ Array Index
           *
           * ถ้า Header ของ Column นั้นว่าง
           * ระบบจะสร้างชื่อสำรอง เช่น
           * - column1
           * - column2
           * - column3
           */
          const key =
            headers[
              columnNumber - 1
            ] ||
            `column${columnNumber}`;

          /**
           * เก็บค่าของ Cell ลงใน Object
           *
           * หมายเหตุ:
           * cell.value อาจไม่ได้เป็น string เสมอไป
           *
           * ค่าที่เป็นไปได้ เช่น
           * - string
           * - number
           * - boolean
           * - Date
           * - Formula Object
           * - Rich Text Object
           * - null
           */
          rowData[key] =
            cell.value;
        },
      );

      /**
       * เพิ่ม Object ของแถวปัจจุบัน
       * เข้าไปใน Array ผลลัพธ์
       */
      rows.push(
        rowData,
      );
    },
  );

  /**
   * คืนข้อมูลทั้งหมดกลับไป
   *
   * Script 1 สามารถใช้ rows.length
   * เพื่อแสดงจำนวนแถวที่ฟังก์ชันอ่านได้
   */
  return rows;
}