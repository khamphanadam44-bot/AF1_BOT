/**
 * excel-cell.util.ts
 * ======================================================
 * รวมฟังก์ชันช่วยสำหรับอ่านและปรับรูปแบบค่าจาก Excel Cell
 *
 * ใช้ร่วมกันในหลายส่วน เช่น:
 * - Report Header Validator
 * - Test Data Header Validator
 * - Field Validator
 * - ฟังก์ชันที่ต้องตรวจว่า Cell ว่างหรือไม่
 *
 * หน้าที่หลัก:
 * 1. แปลงค่าหลายรูปแบบของ ExcelJS ให้เป็นข้อความ
 * 2. จัดรูปแบบข้อความ Header ก่อนตรวจสอบ
 * 3. จัดรูปแบบค่าข้อมูลก่อนนำไปใช้งาน
 * 4. ตรวจสอบว่า Cell ว่างหรือไม่
 * 5. อ่าน Header ทั้งแถวออกมาเป็น Array
 * ======================================================
 */

import ExcelJS from "exceljs";


/**
 * อ่านค่าจาก Excel Cell แล้วแปลงเป็นข้อความธรรมดา
 *
 * เนื่องจาก cell.value ของ ExcelJS ไม่ได้มีเฉพาะ string
 * แต่อาจเป็นข้อมูลหลายประเภท เช่น:
 * - string
 * - number
 * - boolean
 * - Date
 * - Rich Text
 * - Formula Result
 * - Object ที่มี Property ชื่อ text
 *
 * ฟังก์ชันนี้จะพยายามดึงค่าที่สามารถนำไปแสดงผล
 * หรือนำไปเปรียบเทียบได้ออกมาเป็น string
 *
 * หาก Cell ไม่มีค่า จะคืนค่าเป็นข้อความว่าง ""
 */
export const getCellText = (
  cell: ExcelJS.Cell,
): string => {
  // อ่านค่าจริงจาก Cell
  const value = cell.value;

  // Cell ที่ไม่มีค่าจะเป็น null หรือ undefined
  // ให้เปลี่ยนเป็นข้อความว่างเพื่อใช้งานต่อได้ง่าย
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  // หากค่าเป็นข้อมูลพื้นฐาน
  // ให้แปลงเป็น string ได้ทันที
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return String(value);
  }

  // หากค่าเป็นวันที่
  // ให้แปลง Date Object เป็น string
  if (
    value instanceof Date
  ) {
    return String(value);
  }

  // ค่าบางประเภทของ ExcelJS จะอยู่ในรูปแบบ Object
  // เช่น Rich Text, Hyperlink Text หรือ Formula Result
  if (
    typeof value === "object"
  ) {
    /**
     * ตรวจสอบกรณี Cell เป็น Rich Text
     *
     * ตัวอย่างโครงสร้าง:
     * {
     *   richText: [
     *     { text: "Hello " },
     *     { text: "World" }
     *   ]
     * }
     */
    const richTextValue =
      value as {
        richText?: Array<{
          text?: string;
        }>;
      };

    // หากพบ Rich Text ให้นำข้อความแต่ละส่วนมาต่อกัน
    if (
      Array.isArray(
        richTextValue.richText,
      )
    ) {
      return richTextValue
        .richText
        .map(
          (item) =>
            item.text ?? "",
        )
        .join("");
    }

    /**
     * ตรวจสอบกรณี Object มี Property ชื่อ text
     *
     * ตัวอย่าง:
     * {
     *   text: "DS_PTX"
     * }
     */
    const textValue =
      value as {
        text?: string;
      };

    if (
      textValue.text !==
      undefined
    ) {
      return String(
        textValue.text ?? "",
      );
    }

    /**
     * ตรวจสอบกรณี Cell เป็น Formula
     *
     * ExcelJS อาจเก็บข้อมูลเป็น:
     * {
     *   formula: "A1+B1",
     *   result: 100
     * }
     *
     * ฟังก์ชันนี้จะคืนค่าจาก result
     * ไม่ได้คืนข้อความ Formula
     */
    const resultValue =
      value as {
        result?:
          | string
          | number
          | boolean
          | Date
          | null;
      };

    if (
      resultValue.result !==
      undefined
    ) {
      return String(
        resultValue.result ?? "",
      );
    }
  }

  // กรณีไม่ตรงกับรูปแบบที่ตรวจสอบด้านบน
  // ให้ใช้ String() แปลงค่าเป็นข้อความเป็นทางเลือกสุดท้าย
  return String(value);
};


/**
 * จัดรูปแบบข้อความ Header ให้อยู่ในรูปแบบมาตรฐาน
 *
 * ใช้สำหรับ:
 * - ตรวจว่า Header ว่างหรือไม่
 * - เปรียบเทียบ Header แบบทั่วไป
 * - ลดปัญหาจากช่องว่างหรืออักขระที่มองไม่เห็น
 *
 * สิ่งที่ฟังก์ชันนี้ทำ:
 * 1. เปลี่ยน Non-breaking Space ให้เป็นช่องว่างปกติ
 * 2. ลบ Zero-width Character (อักขระพิเศษในรหัสยูนิโค้ด มองไม่เห็นด้วยตาเปล่าและไม่มีพื้นที่ว่างบนหน้าจอ)
 * 3. รวมช่องว่างที่ติดกันหลายช่องให้เหลือหนึ่งช่อง
 * 4. ตัดช่องว่างหัวและท้าย
 * 5. เปลี่ยนข้อความเป็นตัวพิมพ์เล็ก
 *
 * หมายเหตุ:
 * การจับคู่ Header แบบละเอียด เช่น การตรวจ Alias
 * จะใช้ Logic จาก header-matcher.ts
 */
export const normalizeHeader = (
  header: string,
): string => {
  return header

    // เปลี่ยน Non-breaking Space ให้เป็นช่องว่างธรรมดา
    .replace(
      /\u00A0/g,
      " ",
    )

    // ลบอักขระที่มองไม่เห็น เช่น Zero-width Space และ BOM
    .replace(
      /[\u200B-\u200D\uFEFF]/g,
      "",
    )

    // รวมช่องว่างหลายช่องให้เหลือเพียงหนึ่งช่อง
    .replace(
      /\s+/g,
      " ",
    )

    // ลบช่องว่างด้านหน้าและด้านหลังข้อความ
    .trim()

    // เปลี่ยนเป็นตัวพิมพ์เล็ก เพื่อให้เทียบ Header ได้ง่ายขึ้น
    .toLowerCase();
};


/**
 * จัดรูปแบบค่าข้อมูลภายใน Cell ก่อนนำไปตรวจสอบ
 *
 * สิ่งที่ฟังก์ชันนี้ทำ:
 * 1. เปลี่ยน Non-breaking Space เป็นช่องว่างปกติ
 * 2. ลบ Zero-width Character
 * 3. ตัดช่องว่างด้านหน้าและด้านหลัง
 *
 * หมายเหตุ:
 * ฟังก์ชันนี้จะไม่:
 * - เปลี่ยนเป็นตัวพิมพ์เล็ก
 * - รวมช่องว่างที่อยู่ตรงกลางข้อความ
 * - เปลี่ยนรูปแบบวันที่หรือตัวเลข
 */
export const normalizeValue = (
  value: string,
): string => {
  return value
    .replace(
      /\u00A0/g,
      " ",
    )
    .replace(
      /[\u200B-\u200D\uFEFF]/g,
      "",
    )
    .trim();
};


/**
 * ตรวจสอบว่า Excel Cell ว่างหรือไม่
 *
 * ขั้นตอน:
 * 1. ใช้ getCellText() แปลงค่าของ Cell เป็น string
 * 2. ใช้ normalizeValue() ลบช่องว่างและอักขระที่มองไม่เห็น
 * 3. ตรวจว่าค่าที่เหลือเป็นข้อความว่างหรือไม่
 *
 * คืนค่า:
 * - true  = Cell ว่าง
 * - false = Cell มีข้อมูล
 *
 * Cell ที่มีเฉพาะช่องว่างหรือ Zero-width Character
 * จะถือว่าเป็น Cell ว่าง
 */
export const isCellEmpty = (
  cell: ExcelJS.Cell,
): boolean => {
  return (
    normalizeValue(
      getCellText(
        cell,
      ),
    ) === ""
  );
};


/**
 * อ่าน Header ทั้งหมดจากแถวที่กำหนด
 *
 * ตัวอย่าง:
 * หาก Header อยู่ในแถวที่ 1 และมีข้อมูลดังนี้:
 *
 * A1 = Transaction ID
 * B1 = Currency Id
 * C1 = Transaction Amount
 *
 * ผลลัพธ์:
 * [
 *   "Transaction ID",
 *   "Currency Id",
 *   "Transaction Amount"
 * ]
 *
 * การหาจำนวน Column สูงสุดจะพิจารณาจาก:
 * - จำนวน Column ของ Worksheet
 * - จำนวน Cell ที่มีการใช้งานในแถว Header
 * - จำนวน Cell ที่มีค่าจริงในแถว Header
 *
 * Cell ที่ไม่มีข้อมูลจะถูกเก็บเป็นข้อความว่าง ""
 */
export const getHeadersFromRow = (
  worksheet:
    ExcelJS.Worksheet,
  headerRowNumber: number,
): string[] => {
  // อ่านแถวที่กำหนดให้เป็นแถว Header
  const headerRow =
    worksheet.getRow(
      headerRowNumber,
    );

  /**
   * หาจำนวน Column สูงสุดที่ต้องอ่าน
   *
   * ใช้ Math.max() เพื่อป้องกันกรณี:
   * - Worksheet มี Column มากกว่า Cell ใน Header
   * - Header มี Cell ที่ถูกกำหนด Style แต่ไม่มีค่า
   * - Header มีข้อมูลอยู่ใน Column ด้านขวา
   */
  const maxColumn =
    Math.max(
      worksheet.columnCount,
      headerRow.cellCount,
      headerRow.actualCellCount,
    );

  // สร้าง Array สำหรับเก็บชื่อ Header
  const headers: string[] =
    [];

  // วนอ่าน Cell ตั้งแต่ Column 1 ถึง Column สุดท้าย
  for (
    let columnNumber = 1;
    columnNumber <=
    maxColumn;
    columnNumber += 1
  ) {
    /**
     * Array เริ่มตำแหน่งแรกที่ Index 0
     * แต่ ExcelJS เริ่ม Column แรกที่หมายเลข 1
     *
     * จึงเก็บค่าด้วย:
     * headers[columnNumber - 1]
     */
    headers[
      columnNumber - 1
    ] = getCellText(
      headerRow.getCell(
        columnNumber,
      ),
    );
  }

  // คืนรายชื่อ Header ทั้งแถว
  return headers;
};