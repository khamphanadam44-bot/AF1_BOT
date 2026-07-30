/**
 * ============================================================================
 * ไฟล์: report-row-builder.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้อ่านข้อมูลจาก AF1 Report แล้วแปลงแต่ละแถวเป็น ActualRow
 * ซึ่งหมายถึงข้อมูลจริงที่ระบบนำมาเปรียบเทียบกับ ExpectedRow
 *
 * ขั้นตอนการทำงาน
 * --------------
 * 1. อ่าน Config เพื่อทราบว่า Header ของ Report อยู่แถวใด
 * 2. อ่านชื่อ Header ทั้งหมดจาก Excel
 * 3. อ่านข้อมูลตั้งแต่แถวถัดจาก Header จนถึงแถวสุดท้าย
 * 4. แปลงแต่ละแถวเป็น Object ที่ค้นหาค่าด้วยชื่อ Header ได้
 * 5. อ่าน Reference Transaction Number มาใช้เป็น Matching Key
 * 6. ข้ามแถวที่ไม่มี Matching Key เพราะมักเป็นแถวว่างหรือแถวท้ายไฟล์
 *
 * ผลลัพธ์จากไฟล์นี้จะถูกส่งไปยัง compare-validator.ts
 * เพื่อจับคู่กับข้อมูล Expected
 * ============================================================================
 */


/**
 * ส่วน import ด้านล่าง คือการนำเครื่องมือหรือโครงสร้างข้อมูล
 * จากไฟล์อื่นมาใช้ในไฟล์นี้ เปรียบเหมือนการหยิบอุปกรณ์ที่เตรียมไว้แล้ว
 * มาใช้งาน โดยไม่ต้องเขียนทุกอย่างซ้ำใหม่
 */

import ExcelJS from "exceljs";

import {
  getHeadersFromRow,
} from "../../validators/shared/excel-cell.util";

import {
  getMappingHeaderRowNumber,
  getMappingMatchingKeyHeaders,
} from "../../../config/mapping-helper";

import {
  mapRowToObject,
} from "./row-mapper";

import {
  ActualRow,
  ReportRow,
} from "./compare-types";

/**
 * สร้างแถวข้อมูลจริงจาก Report
 *
 * @param worksheet Worksheet ของ Report
 * @param reportName เช่น DS_PTX
 */
export const buildActualRows = (
  worksheet: ExcelJS.Worksheet,
  reportName: string,
): ActualRow[] => {

  /**
   * --------------------------------------------------------------------------
   * Header Row ของ Report
   * --------------------------------------------------------------------------
   */
  const headerRowNumber =
    getMappingHeaderRowNumber(
      reportName,
    );

  /**
   * อ่าน Header ทั้งหมดจาก Report
   */
  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * Header ที่ใช้เป็น Matching Key
   *
   * DS_PTX:
   * Reference Transaction Number
   */
  const matchingKeyHeader =
    getMappingMatchingKeyHeaders(
      reportName,
    )[0];

  /**
   * ป้องกัน Config ไม่มี Matching Key Header
   */
  if (!matchingKeyHeader) {

    throw new Error(
      `Matching Key Header not found for report: ${reportName}`,
    );

  }

  const actualRows: ActualRow[] = [];

  /**
   * --------------------------------------------------------------------------
   * อ่านข้อมูลหลัง Header Row
   * --------------------------------------------------------------------------
   */
  for (
    let rowNumber = headerRowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {

    const row =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * แปลง Excel Row เป็น Object
     *
     * ตัวอย่าง:
     *
     * {
     *   "Dept Code": "FRO",
     *   "System Id": "GPMH",
     *   "Reference Transaction Number": "...",
     * }
     */
    const rowData =
      mapRowToObject(
        row,
        headers,
      ) as ReportRow;

    /**
     * อ่าน Matching Key
     */
    const matchingKey =
      String(
        rowData[
          matchingKeyHeader
        ] ?? "",
      ).trim();

    /**
     * ข้ามแถวที่ไม่มี Matching Key
     *
     * ป้องกันแถวว่างหรือแถวท้ายไฟล์
     * ถูกนำไปสร้าง ActualRow
     */
    if (
      matchingKey === ""
    ) {

      continue;

    }

    /**
     * สร้าง ActualRow
     */
    actualRows.push({

      /**
       * Row Number จริงใน Report
       */
      rowNumber,

      /**
       * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
       */
      matchingKey,

      /**
       * ข้อมูล Report ทั้งแถว
       */
      data:
        rowData,

    });

  }

  return actualRows;

};