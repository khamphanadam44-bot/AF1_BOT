/**
 * report-header-validator.ts
 * ======================================================
 * ใช้ตรวจสอบ Header ของ Report ที่ดาวน์โหลดจากระบบ
 *
 * หน้าที่หลักของไฟล์นี้:
 * 1. อ่าน Header จากแถวที่กำหนด
 * 2. ตรวจว่า Header ที่จำเป็นมีครบหรือไม่
 * 3. Highlight Header ที่พบตาม Requirement เป็นสีเขียว
 * 4. คืน Header อื่นให้เป็นรูปแบบเริ่มต้น
 * 5. บันทึกการ Highlight กลับลงไฟล์ Report เดิม
 * 6. Throw Error หากพบ Header ที่ขาด
 *
 * หมายเหตุ:
 * - ไฟล์นี้ไม่สร้าง Sheet "Header Validation"
 * - ใช้ Logic กลางจาก header-validation-sheet.ts
 *   สำหรับจับคู่ Header และตรวจหา Header ที่ขาด
 * ======================================================
 */

import ExcelJS from "exceljs";

import {
  getHeadersFromRow,
  normalizeHeader,
} from "../shared/excel-cell.util";

import {
  applyDefaultHeaderStyle,
  applyFill,
  COLORS,
} from "../shared/excel-style.util";

import {
  findHeaderMatchResults,
  getMissingHeaders,
  isExpectedHeader,
} from "../shared/header-validation-sheet";


/**
 * Highlight Header ภายในแถว Header ของ Report
 *
 * การทำงาน:
 * 1. อ่าน Header จากแถวที่กำหนด
 * 2. คืนรูปแบบของทุก Cell ให้เป็นรูปแบบเริ่มต้นก่อน
 * 3. ข้าม Cell ที่ไม่มีชื่อ Header
 * 4. หาก Header อยู่ใน Requirement หรือ Alias
 *    ให้ Highlight Cell นั้นเป็นสีเขียว
 *
 * Header ที่ไม่อยู่ใน Requirement:
 * - จะไม่ถูก Highlight สีเขียว
 * - จะใช้รูปแบบเริ่มต้นจาก applyDefaultHeaderStyle()
 */
const highlightHeaderRow = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber: number,
  expectedHeaders: string[],
  aliases: Record<string, string[]>,
): void => {
  // อ่านแถวที่กำหนดให้เป็นแถว Header
  const headerRow =
    worksheet.getRow(
      headerRowNumber,
    );

  // อ่านชื่อ Header ทั้งหมดจากแถวดังกล่าว
  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  // ตรวจ Header ทีละช่องจากซ้ายไปขวา
  headers.forEach(
    (
      header,
      index,
    ) => {
      // index ของ Array เริ่มจาก 0
      // แต่เลข Column ของ ExcelJS เริ่มจาก 1
      const cell =
        headerRow.getCell(
          index + 1,
        );

      // คืนรูปแบบของ Cell ให้เป็นค่าเริ่มต้นก่อน
      // เพื่อป้องกันสีเก่าจากการ Validate ครั้งก่อน
      applyDefaultHeaderStyle(
        cell,
      );

      // หากช่อง Header ว่าง ไม่ต้องตรวจและไม่ต้อง Highlight
      if (
        normalizeHeader(
          header,
        ) === ""
      ) {
        return;
      }

      // ตรวจว่า Header อยู่ใน Requirement
      // หรือสามารถจับคู่กับชื่อ Alias ที่กำหนดไว้ได้หรือไม่
      if (
        isExpectedHeader(
          header,
          expectedHeaders,
          aliases,
        )
      ) {
        // Header ที่ตรงกับ Requirement จะถูก Highlight สีเขียว
        applyFill(
          cell,
          COLORS.GREEN,
        );
      }
    },
  );
};


/**
 * ตัวเลือกสำหรับการ Validate Header ของ Report
 */
export type ReportHeaderValidationOptions = {
  /**
   * หมายเลขแถวที่ใช้เป็น Header ของ Report
   *
   * ตัวอย่าง:
   * - DS_PTX ใช้ Header แถวที่ 1
   * - Report อื่นอาจกำหนดเป็นแถวอื่นได้
   */
  headerRowNumber: number;

  /**
   * ชื่อเรียกแทนของ Header ที่ยอมรับได้
   *
   * ตัวอย่าง:
   * Currency Id อาจรองรับ:
   * - Currency Id
   * - Currency ID
   *
   * Alias จะถูกกำหนดแยกตาม Report
   */
  aliases: Record<
    string,
    string[]
  >;
};


/**
 * ตรวจสอบ Header ของ Report ที่ Path ที่กำหนด
 *
 * ขั้นตอนการทำงาน:
 * 1. เปิดไฟล์ Excel Report
 * 2. เลือก Worksheet แรก
 * 3. อ่าน Header จากแถวที่กำหนด
 * 4. Highlight Header ที่ตรงกับ Requirement เป็นสีเขียว
 * 5. จับคู่ Actual Header กับ Expected Header
 * 6. ตรวจหา Expected Header ที่ไม่พบ
 * 7. บันทึกการ Highlight กลับลงไฟล์เดิม
 * 8. Throw Error หากมี Header ขาด
 *
 * หมายเหตุ:
 * - ฟังก์ชันนี้ไม่สร้าง Sheet "Header Validation"
 * - ใช้เฉพาะ Logic กลางจาก header-validation-sheet.ts
 *   เพื่อคำนวณสถานะ FOUND หรือ MISSING
 */
export const validateReportHeader =
  async (
    reportFilePath: string,
    expectedHeaders: string[],
    options:
      ReportHeaderValidationOptions,
  ): Promise<void> => {
    console.log(
      "\n===== REPORT HEADER VALIDATION =====",
    );

    // แยกค่าที่ต้องใช้จาก Options
    const {
      headerRowNumber,
      aliases,
    } = options;

    // สร้าง Workbook สำหรับเปิดไฟล์ Excel
    const workbook =
      new ExcelJS.Workbook();

    // อ่านไฟล์ Report จาก Path ที่ได้รับ
    await workbook.xlsx.readFile(
      reportFilePath,
    );

    // เลือก Worksheet แรกของ Report
    const worksheet =
      workbook.getWorksheet(
        1,
      );

    // หากไม่พบ Worksheet ให้หยุดการทำงาน
    if (!worksheet) {
      throw new Error(
        "Report worksheet not found",
      );
    }

    // อ่าน Header จริงจากแถวที่กำหนด
    const actualHeaders =
      getHeadersFromRow(
        worksheet,
        headerRowNumber,
      );

    // Highlight Header ที่ตรงกับ Requirement
    // และคืน Header อื่นเป็นรูปแบบเริ่มต้น
    highlightHeaderRow(
      worksheet,
      headerRowNumber,
      expectedHeaders,
      aliases,
    );

    // เปรียบเทียบ Expected Header กับ Actual Header
    // โดยรองรับชื่อ Alias ของ Report
    const results =
      findHeaderMatchResults(
        actualHeaders,
        expectedHeaders,
        aliases,
      );

    // ดึงเฉพาะ Expected Header ที่ไม่พบใน Report
    const missingHeaders =
      getMissingHeaders(
        results,
      );

    // บันทึกการ Highlight กลับลงไฟล์ Report เดิม
    // การบันทึกจะเกิดขึ้นก่อน Throw Error
    // เพื่อให้ผู้ใช้เปิดไฟล์ดูจุดที่ตรวจสอบได้
    await workbook.xlsx.writeFile(
      reportFilePath,
    );

    // หากพบ Header ขาด ให้แจ้งจำนวนและรายชื่อ Header
    if (
      missingHeaders.length >
      0
    ) {
      throw new Error(
        `Missing ${missingHeaders.length} Header(s): ${missingHeaders.join(", ")}`,
      );
    }

    // แสดงข้อความเมื่อ Header ครบทุกช่อง
    console.log(
      "✅ Report Header Validation Passed",
    );
  };