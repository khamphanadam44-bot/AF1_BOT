/**
 * test-data-validator.ts
 * ------------------------------------------------------------
 * ตัวควบคุมหลักของ Test Data Validation ใน Script 2
 *
 * หน้าที่หลัก:
 * 1. ตรวจสอบว่าไฟล์ Test Data มีอยู่จริง
 * 2. เปิดไฟล์ Test Data
 * 3. อ่าน Header จริงจาก Worksheet
 * 4. ตรวจจำนวน Fee Group จาก Header จริง
 * 5. สร้าง Required Header ตาม Report
 * 6. ตรวจข้อมูลใน Required Field
 * 7. ตรวจ Header
 * 8. บันทึกไฟล์ผลลัพธ์
 *
 * การตรวจ Fee Group:
 *
 * DS_PTX และ DS_LTX:
 * - ตรวจจำนวน Fee Group จาก Header จริง
 * - ตรวจข้อมูลภายใน Fee Group
 *
 * DS_FTX และ DS_FTU:
 * - ไม่มีการตรวจ Fee Group
 * - จำนวน Fee Group เป็น 0
 *
 * หมายเหตุ:
 * ทุก Report สามารถใช้ไฟล์ Test Data ตัวเดียวกันได้
 * แต่ Requirement ที่นำมาตรวจจะเลือกตาม reportCode
 * ------------------------------------------------------------
 */

import * as fs from "fs";

import ExcelJS from "exceljs";

import {
  validateTestDataHeader,
} from "./test-data-header-validator";

import {
  validateRequiredFields,
} from "./field-validator";

import {
  getHeadersFromRow,
} from "../shared/excel-cell.util";

import {
  getFeeTypeCount,
  getNormalRequiredTestDataHeaders,
  getRequiredTestDataHeaders,
  getTestDataHeaderRowNumber,
} from "../../../config/testdata-helper";

import type {
  TestDataReportCode,
} from "../../../config/testdata-config";

import {
  buildTempFilePath,
  buildTimestampedFilePath,
  cleanupStaleTempFiles,
  ensureDirectoryExists,
} from "../../file-system.util";

/**
 * ตรวจ Header และข้อมูลทั้งหมดของ Test Data
 *
 * ไฟล์ต้นทางจะไม่ถูกแก้ไขโดยตรง
 * ระบบจะบันทึกผลการตรวจเป็นไฟล์ใหม่
 *
 * @param inputFilePath
 * Path ของไฟล์ Test Data ต้นทาง
 *
 * @param outputDirectory
 * Folder สำหรับเก็บไฟล์ผลลัพธ์
 *
 * @param resultBaseName
 * ชื่อหลักของไฟล์ผลลัพธ์ก่อนเติม Timestamp
 *
 * ตัวอย่าง:
 * DS_PTX_TestData_Validation_Result
 *
 * @param reportCode
 * Report ที่กำลังตรวจ
 *
 * ตัวอย่าง:
 * - DS_PTX
 * - DS_LTX
 * - DS_FTX
 * - DS_FTU
 *
 * @returns
 * Path ของไฟล์ผลลัพธ์
 *
 * หมายเหตุ:
 * - ถ้า Header ขาด ระบบจะบันทึกไฟล์ก่อน Throw Error
 * - ถ้า Field ว่าง ระบบจะไม่ Throw Error
 */
export const validateTestData = async (
  inputFilePath: string,
  outputDirectory: string,
  resultBaseName: string,
  reportCode: TestDataReportCode,
): Promise<string> => {
  console.log(
    "\n===== TEST DATA VALIDATION =====",
  );

  console.log(
    `Report Code : ${reportCode}`,
  );

  console.log(
    `Input File  : ${inputFilePath}`,
  );

  /**
   * ตรวจสอบว่าไฟล์ Test Data ต้นทางมีอยู่จริง
   */
  if (!fs.existsSync(inputFilePath)) {
    throw new Error(
      `Test Data input file not found: ${inputFilePath}`,
    );
  }

  /**
   * สร้าง Folder ผลลัพธ์หากยังไม่มี
   */
  ensureDirectoryExists(
    outputDirectory,
  );

  /**
   * สร้างชื่อไฟล์ผลลัพธ์พร้อม Timestamp
   *
   * ตัวอย่าง:
   * DS_PTX_TestData_Validation_Result_20260804_143000.xlsx
   */
  const finalOutputFilePath =
    buildTimestampedFilePath(
      outputDirectory,
      resultBaseName,
      ".xlsx",
    );

  /**
   * สร้าง Path สำหรับไฟล์ชั่วคราว
   *
   * ระบบจะเขียนไฟล์ชั่วคราวให้เสร็จก่อน
   * แล้วจึงเปลี่ยนชื่อเป็นไฟล์ผลลัพธ์จริง
   */
  const tempOutputFilePath =
    buildTempFilePath(
      finalOutputFilePath,
    );

  /**
   * ลบไฟล์ชั่วคราวเก่าที่อาจค้างจากการ Run ครั้งก่อน
   */
  cleanupStaleTempFiles(
    outputDirectory,
  );

  /**
   * สร้าง Workbook และอ่านไฟล์ Test Data
   */
  const workbook =
    new ExcelJS.Workbook();

  await workbook.xlsx.readFile(
    inputFilePath,
  );

  /**
   * เลือก Worksheet ลำดับแรก
   */
  const worksheet =
    workbook.getWorksheet(1);

  if (!worksheet) {
    throw new Error(
      "Original worksheet not found",
    );
  }

  console.log(
    `Output File : ${finalOutputFilePath}`,
  );

  /**
   * ดึงหมายเลขแถว Header จาก Config ของ Report
   *
   * ไม่กำหนดหมายเลขแถวแบบ Hard code ในไฟล์นี้
   */
  const headerRowNumber =
    getTestDataHeaderRowNumber(
      reportCode,
    );

  /**
   * อ่าน Header จริงจาก Test Data
   *
   * ต้องทำขั้นตอนนี้ก่อนสร้าง Required Header
   * เพราะจำนวน Fee Group จะตรวจจาก Header จริง
   */
  const actualHeaders =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * ตรวจจำนวน Fee Group ตาม Report
   *
   * DS_PTX และ DS_LTX:
   * - ตรวจหมายเลข Fee Group สูงสุดจาก Header จริง
   *
   * DS_FTX และ DS_FTU:
   * - คืนค่า 0
   */
  const feeTypeCount =
    getFeeTypeCount(
      reportCode,
      actualHeaders,
    );

  /**
   * ดึง Required Header ทั้งหมด
   *
   * DS_PTX และ DS_LTX:
   * - รวม Fee Header ตามจำนวนที่ตรวจพบจริง
   *
   * DS_FTX และ DS_FTU:
   * - ไม่มี Fee Header
   */
  const requiredHeaders =
    getRequiredTestDataHeaders(
      reportCode,
      actualHeaders,
    );

  /**
   * ดึงเฉพาะ Header สำหรับตรวจ Normal Field
   *
   * ไม่รวม Fee Group เพราะ Fee Group
   * มี Logic ตรวจแยกต่างหาก
   */
  const normalRequiredHeaders =
    getNormalRequiredTestDataHeaders(
      reportCode,
    );

  console.log(
    `Header Row : ${headerRowNumber}`,
  );

  console.log(
    `Required Header Count : ${requiredHeaders.length}`,
  );

  console.log(
    `Detected Fee Group Count : ${feeTypeCount}`,
  );

  /**
   * ขั้นตอนที่ 1:
   * ตรวจข้อมูลภายใน Required Field
   *
   * ค่าที่ได้:
   * true  = พบข้อมูลไม่ครบอย่างน้อย 1 จุด
   * false = ไม่พบข้อมูลไม่ครบ
   */
  const hasInvalidField =
    await validateRequiredFields(
      workbook,
      normalRequiredHeaders,
      headerRowNumber,
      reportCode,
    );

  /**
   * ขั้นตอนที่ 2:
   * ตรวจ Header ของ Test Data
   *
   * ค่าที่ได้:
   * [] = พบ Header ครบ
   *
   * Array มีข้อมูล = พบ Header ขาด
   */
  const missingHeaders =
    validateTestDataHeader(
      workbook,
      requiredHeaders,
      headerRowNumber,
    );

  /**
   * บันทึก Workbook ลงไฟล์ชั่วคราวก่อน
   *
   * Workbook อาจมีการเปลี่ยนแปลงจาก Field Validation:
   * - Highlight สีเขียว
   * - Highlight สีแดง
   * - Highlight สีเหลือง
   * - ใส่ข้อความ "โปรดกรอกข้อมูล"
   * - ใส่ข้อความ "โปรดตรวจสอบข้อมูล"
   * - สร้าง Sheet "Field Validation"
   */
  await workbook.xlsx.writeFile(
    tempOutputFilePath,
  );

  /**
   * เปลี่ยนชื่อไฟล์ชั่วคราวเป็นไฟล์ผลลัพธ์จริง
   */
  fs.renameSync(
    tempOutputFilePath,
    finalOutputFilePath,
  );

  console.log(
    "✅ Validation result file created successfully",
  );

  /**
   * ถ้าพบ Missing Header:
   *
   * 1. ไฟล์ผลลัพธ์จะถูกบันทึกเรียบร้อยแล้ว
   * 2. ระบบจึง Throw Error
   * 3. Script 2 จะแสดงสถานะ Fail
   */
  if (missingHeaders.length > 0) {
    throw new Error(
      `Header Validation Failed: Missing ${
        missingHeaders.length
      } Header(s): ${missingHeaders.join(", ")}`,
    );
  }

  /**
   * ถ้าพบ Field ที่ข้อมูลไม่ครบ:
   *
   * - แสดงข้อความ Failed ใน Console
   * - ไม่ Throw Error
   * - คืน Path ของไฟล์ผลลัพธ์ตามปกติ
   */
  if (hasInvalidField) {
    console.log(
      "❌ Field Validation Failed",
    );

    return finalOutputFilePath;
  }

  /**
   * Header ครบและไม่พบ Invalid Field
   */
  console.log(
    "✅ Test Data Validation Passed",
  );

  return finalOutputFilePath;
};