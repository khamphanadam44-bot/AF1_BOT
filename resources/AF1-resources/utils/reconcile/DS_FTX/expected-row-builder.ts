/**
 * expected-row-builder.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * อ่านข้อมูลจาก Worksheet ของ Test Data
 * แล้วแปลงแต่ละแถวให้เป็น ExpectedRow
 *
 * ExpectedRow จำนวน 1 รายการประกอบด้วย
 * 1. หมายเลขแถวจริงในไฟล์ Test Data
 * 2. Test No.
 * 3. Matching Key
 * 4. ข้อมูลทั้งหมดของ Test Data แถวนั้น
 *
 * ExpectedRow จะถูกใช้เป็นข้อมูลฝั่งที่ระบบคาดหวัง
 * และนำไปเปรียบเทียบกับ ActualRow จาก Report DS_FTX
 *
 * Flow การทำงาน
 *
 * Test Data Worksheet
 * → อ่าน Header
 * → อ่านข้อมูลทีละแถว
 * → แปลงแถวเป็น Object
 * → ข้ามแถวว่าง
 * → อ่าน Test No. และ Matching Key
 * → สร้าง ExpectedRow[]
 *
 * หมายเหตุสำคัญ
 * - Header เริ่มต้นอยู่ที่แถว 5
 * - ผู้เรียกสามารถส่งหมายเลข Header Row อื่นเข้ามาได้
 * - แถวว่างระหว่างข้อมูลจะถูกข้าม
 * - แถวที่มีข้อมูลแต่ Matching Key ว่างยังถูกเก็บไว้
 * - การตัดสินว่า Matching Key ว่างเป็น FAIL
 *   จะทำภายหลังใน compare-validator.ts
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  ExpectedRow,
  TestDataRow,
} from "./compare-types";

import {
  buildTestDataMatchingKey,
} from "./build-matching-key";

import {
  mapRowToObject,
} from "./row-mapper";

import {
  getCellText,
} from "../../validators/shared/excel-cell.util";

/**
 * หมายเลขแถว Header เริ่มต้นของ Test Data
 *
 * ปัจจุบัน Test Data ของโปรเจกต์นี้
 * มี Header อยู่ที่แถว 5
 *
 * ค่านี้เป็นเพียงค่าเริ่มต้น
 * ผู้เรียก buildExpectedRows() สามารถส่งค่าอื่นเข้ามาแทนได้
 */
export const DEFAULT_TEST_DATA_HEADER_ROW_NUMBER =
  5;

/**
 * ชื่อ Header ที่ใช้เก็บหมายเลข Test Case
 *
 * ตัวอย่างค่า:
 * - BOTDMS_001
 * - BOTDMS_002
 *
 * Test No. ใช้สำหรับ
 * - แสดงในไฟล์ผลลัพธ์
 * - ระบุว่าผลลัพธ์มาจาก Test Case ใด
 *
 * Test No. ไม่ได้ใช้เป็น Matching Key
 * เพราะ Report DS_FTX ไม่มี Header Test No.
 *
 * หมายเหตุ
 * การอ่านค่าจะใช้ชื่อ Header แบบตรงตัว
 * จึงต้องมีชื่อว่า "Test No." ตามค่าที่กำหนดไว้
 */
export const TEST_SCRIPT_NO_HEADER =
  "Test No.";

/**
 * แปลงค่าทั่วไปให้อยู่ในรูปแบบข้อความ
 *
 * การทำงาน
 * 1. ถ้าค่าเป็น null หรือ undefined ให้คืนข้อความว่าง
 * 2. แปลงค่าเป็น string
 * 3. ตัดช่องว่างด้านหน้าและด้านหลัง
 *
 * ตัวอย่าง:
 * null        → ""
 * undefined   → ""
 * 123         → "123"
 * " FTX-001 " → "FTX-001"
 *
 * @param value ค่าที่ต้องการแปลงเป็นข้อความ
 * @returns ข้อความที่ตัดช่องว่างแล้ว
 */
const normalizeText = (
  value: unknown,
): string => {
  // ถ้าไม่มีค่า ให้คืนข้อความว่าง
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  // แปลงเป็น string และตัดช่องว่างหน้า–หลัง
  return String(
    value,
  ).trim();
};

/**
 * อ่านชื่อ Header ทั้งหมดจากแถวที่กำหนด
 *
 * ตำแหน่งของ Header ใน Array
 * จะต้องตรงกับตำแหน่ง Column ใน Excel
 *
 * ตัวอย่าง:
 *
 * [
 *   "Test No.",                         // Excel Column A
 *   "Transaction ID/ Reconcile ID",     // Excel Column B
 *   "From Currency (CCY)",              // Excel Column C
 *   "To Currency (CCY)",                // Excel Column D
 *   "Txn Date"                          // Excel Column E
 * ]
 *
 * Header ที่อ่านได้จะถูก
 * - แปลงเป็นข้อความด้วย getCellText()
 * - ตัดช่องว่างด้านหน้าและด้านหลังด้วย trim()
 *
 * แต่จะไม่ถูก
 * - เปลี่ยนเป็นตัวพิมพ์เล็กหรือใหญ่
 * - ลบช่องว่างภายใน
 * - แปลงผ่าน Alias
 *
 * @param worksheet Worksheet ของ Test Data
 * @param headerRowNumber หมายเลขแถว Header
 * @returns รายการ Header เรียงตามตำแหน่ง Column
 */
const readHeaders = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber: number,
): string[] => {
  /**
   * ดึงแถวที่กำหนดให้เป็น Header
   */
  const headerRow =
    worksheet.getRow(
      headerRowNumber,
    );

  /**
   * หาจำนวน Column ที่ต้องอ่าน
   *
   * ใช้ค่าที่มากที่สุดระหว่าง
   *
   * worksheet.actualColumnCount
   * = จำนวน Column ที่ ExcelJS มองว่ามีข้อมูลใน Worksheet
   *
   * headerRow.cellCount
   * = จำนวน Cell ที่แถว Header ครอบคลุม
   *
   * หมายเหตุ:
   * Logic นี้เหมาะกับตารางที่ Column เรียงต่อเนื่องกัน
   * หากมี Column กระโดดห่างมากและ Header ว่าง
   * อาจต้องตรวจสอบพฤติกรรมเพิ่มเติม
   */
  const lastColumnNumber =
    Math.max(
      worksheet.actualColumnCount,
      headerRow.cellCount,
    );

  /**
   * Array สำหรับเก็บชื่อ Header
   */
  const headers: string[] =
    [];

  /**
   * วนอ่าน Header ตั้งแต่ Column 1
   * ไปจนถึง Column สุดท้ายที่คำนวณได้
   */
  for (
    let columnNumber = 1;
    columnNumber <=
    lastColumnNumber;
    columnNumber += 1
  ) {
    /**
     * อ่านชื่อ Header จาก Cell
     *
     * ถ้า Cell ว่าง getCellText() จะคืน ""
     */
    const header =
      getCellText(
        headerRow.getCell(
          columnNumber,
        ),
      ).trim();

    /**
     * เพิ่ม Header เข้า Array ตามลำดับ Column
     *
     * ถ้า Header ว่าง จะยังเพิ่ม ""
     * เพื่อรักษาตำแหน่ง Column ให้ตรงกับ Excel
     */
    headers.push(
      header,
    );
  }

  return headers;
};

/**
 * ตรวจสอบว่า Object ของแถวปัจจุบัน
 * มีข้อมูลจริงอย่างน้อย 1 Field หรือไม่
 *
 * การทำงาน
 * 1. ดึง Value ทั้งหมดจาก Object
 * 2. Normalize Value ทีละรายการ
 * 3. ถ้ามีค่าใดไม่เป็นข้อความว่าง ให้ถือว่าแถวมีข้อมูล
 *
 * ใช้สำหรับข้าม
 * - แถวว่างระหว่างข้อมูล
 * - แถวว่างท้าย Worksheet
 *
 * หมายเหตุ
 * mapRowToObject() ไม่เก็บข้อมูลของ Column ที่ไม่มี Header
 * ดังนั้น ถ้าแถวมีข้อมูลเฉพาะใต้ Column ที่ Header ว่าง
 * ฟังก์ชันนี้จะถือว่าแถวนั้นไม่มีข้อมูล
 *
 * @param data Object ข้อมูล Test Data จำนวน 1 แถว
 *
 * @returns
 * true  = พบข้อมูลอย่างน้อย 1 Field
 * false = ไม่พบข้อมูล
 */
const hasRowData = (
  data: TestDataRow,
): boolean => {
  return Object.values(
    data,
  ).some(
    (
      value,
    ) => {
      return normalizeText(
        value,
      ) !== "";
    },
  );
};

/**
 * สร้าง ExpectedRow ทั้งหมดจาก Worksheet ของ Test Data
 *
 * ลำดับการทำงาน
 * 1. ตรวจสอบหมายเลขแถว Header
 * 2. อ่าน Header
 * 3. เริ่มอ่านข้อมูลจากแถวถัดจาก Header
 * 4. แปลงข้อมูลแต่ละแถวเป็น Object
 * 5. ข้ามแถวที่ไม่มีข้อมูล
 * 6. อ่าน Test No.
 * 7. อ่าน Matching Key
 * 8. สร้าง ExpectedRow
 *
 * @param worksheet
 * Worksheet ของไฟล์ Test Data
 *
 * @param headerRowNumber
 * หมายเลขแถว Header
 *
 * ค่าเริ่มต้นคือแถว 5
 *
 * @returns
 * ExpectedRow[] เรียงตามหมายเลขแถวใน Test Data
 */
export const buildExpectedRows = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber =
    DEFAULT_TEST_DATA_HEADER_ROW_NUMBER,
): ExpectedRow[] => {
  /**
   * ป้องกันหมายเลขแถวที่น้อยกว่า 1
   *
   * ExcelJS เริ่มนับแถวจาก 1
   * จึงไม่สามารถใช้แถว 0 หรือเลขติดลบได้
   *
   * หมายเหตุ:
   * เงื่อนไขนี้ตรวจเฉพาะค่าที่น้อยกว่า 1
   * ไม่ได้ตรวจว่า Header Row เกินจำนวนแถวหรือไม่
   */
  if (
    headerRowNumber < 1
  ) {
    throw new Error(
      `Invalid Test Data header row number: ${headerRowNumber}`,
    );
  }

  /**
   * อ่าน Header จากแถวที่กำหนด
   */
  const headers =
    readHeaders(
      worksheet,
      headerRowNumber,
    );

  /**
   * Array สำหรับเก็บ ExpectedRow ทั้งหมด
   */
  const expectedRows:
  ExpectedRow[] = [];

  /**
   * เริ่มอ่านข้อมูลจากแถวถัดจาก Header
   *
   * ตัวอย่าง:
   * Header อยู่แถว 5
   * ข้อมูลจะเริ่มอ่านจากแถว 6
   */
  for (
    let rowNumber =
      headerRowNumber + 1;

    /**
     * วนอ่านจนถึงหมายเลขแถวสุดท้ายของ Worksheet
     *
     * ใช้ worksheet.rowCount
     * เพื่อไม่ให้ข้อมูลหลังแถวว่างคั่นกลางหายไป
     */
    rowNumber <=
    worksheet.rowCount;
    rowNumber += 1
  ) {
    /**
     * ดึง Excel Row ตามหมายเลขแถวปัจจุบัน
     */
    const excelRow =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * แปลง Excel Row ให้เป็น Object
     * โดยใช้ชื่อ Header เป็น Key
     *
     * ตัวอย่าง:
     *
     * {
     *   "Test No.": "FTX-001",
     *   "Transaction ID/ Reconcile ID": "TX001",
     *   "From Currency (CCY)": "EUR"
     * }
     *
     * as TestDataRow เป็น Type Assertion
     * เพื่อบอก TypeScript ว่า Object นี้
     * ต้องการใช้งานในรูปแบบ TestDataRow
     *
     * ไม่ได้ทำการแปลงข้อมูลเพิ่มเติมตอน Run
     */
    const data =
      mapRowToObject(
        excelRow,
        headers,
      ) as TestDataRow;

    /**
     * ถ้าแถวไม่มีข้อมูล ให้ข้ามแถวนั้น
     *
     * continue หมายถึงจบรอบปัจจุบัน
     * แล้วไปอ่านแถวถัดไป
     */
    if (
      !hasRowData(
        data,
      )
    ) {
      continue;
    }

    /**
     * อ่าน Test No. จาก Header "Test No."
     *
     * ถ้าไม่พบ Header หรือไม่มีค่า
     * testScriptNo จะเป็นข้อความว่าง
     *
     * แถวนี้ยังถูกเพิ่มเป็น ExpectedRow ตามปกติ
     */
    const testScriptNo =
      normalizeText(
        data[
          TEST_SCRIPT_NO_HEADER
        ],
      );

    /**
     * อ่าน Matching Key จาก Header
     * "Transaction ID/ Reconcile ID"
     *
     * buildTestDataMatchingKey()
     * จะตัดช่องว่างด้านหน้าและด้านหลัง
     *
     * ถ้าไม่พบ Header หรือไม่มีค่า
     * matchingKey จะเป็นข้อความว่าง
     *
     * แถวนี้ยังถูกเพิ่มเป็น ExpectedRow
     * และจะถูกตัดสินเป็น FAIL ภายหลัง
     * โดย compare-validator.ts
     */
    const matchingKey =
      buildTestDataMatchingKey(
        data,
      );

    /**
     * รวมข้อมูลของแถวปัจจุบันเป็น ExpectedRow
     */
    expectedRows.push({
      // หมายเลขแถวจริงใน Test Data
      rowNumber,

      // Test No. สำหรับแสดงในผลลัพธ์
      testScriptNo,

      // Matching Key สำหรับจับคู่กับ Report
      matchingKey,

      // ข้อมูลทั้งหมดของ Test Data แถวนี้
      data,
    });
  }

  /**
   * คืน ExpectedRow ทั้งหมด
   * โดยเรียงตามลำดับแถวใน Test Data
   */
  return expectedRows;
};