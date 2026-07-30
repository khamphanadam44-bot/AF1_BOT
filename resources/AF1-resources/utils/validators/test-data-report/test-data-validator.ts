/**
 * test-data-validator.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวควบคุมหลักของ Test Data Validation ใน Script 2
 *
 * ใช้ตรวจสอบ
 * 1. Header ของไฟล์ Test Data
 * 2. ข้อมูลภายใน Required Field
 * 3. Fee Group เฉพาะ Report ที่รองรับ เช่น DS_PTX
 *
 * ถึงแม้ทุก Report จะใช้ไฟล์ Test Data ตัวเดียวกัน
 * แต่ Requirement ที่นำมาตรวจสอบจะเลือกตาม reportCode
 *
 * ตัวอย่าง:
 * - reportCode = DS_PTX → ใช้ Config ของ DS_PTX
 * - reportCode = DS_FTX → ใช้ Config ของ DS_FTX
 *
 * ลำดับการทำงาน
 * 1. ตรวจสอบว่าไฟล์ Test Data มีอยู่จริง
 * 2. เตรียม Folder และชื่อไฟล์ผลลัพธ์
 * 3. เปิดไฟล์ Test Data
 * 4. อ่าน Config ตาม Report
 * 5. ตรวจสอบข้อมูลภายใน Field
 * 6. ตรวจสอบ Header
 * 7. บันทึกไฟล์ผลลัพธ์
 * 8. สรุปผลการตรวจสอบ
 *
 * คำศัพท์
 * - Input File  = ไฟล์ต้นทาง
 * - Output File = ไฟล์ผลลัพธ์
 * - Temporary File = ไฟล์ชั่วคราว
 * - Validation  = การตรวจสอบ
 * - Invalid     = ข้อมูลไม่ถูกต้องหรือไม่ครบ
 * ------------------------------------------------------------------
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
 * ตรวจสอบ Header และข้อมูลทั้งหมดของไฟล์ Test Data
 *
 * ไฟล์ Test Data ต้นทางจะไม่ถูกแก้ไขโดยตรง
 * เพราะระบบจะบันทึกผลเป็นไฟล์ใหม่ใน outputDirectory
 *
 * @param inputFilePath
 * Path ของไฟล์ Test Data ต้นทาง
 *
 * @param outputDirectory
 * Folder สำหรับเก็บไฟล์ผลการตรวจสอบ
 *
 * @param resultBaseName
 * ชื่อหลักของไฟล์ผลลัพธ์ก่อนเติม Timestamp
 *
 * ตัวอย่าง:
 * DS_PTX_TestData_Validation_Result
 *
 * @param reportCode
 * รหัส Report ที่ใช้เลือก Config และ Logic การตรวจสอบ
 *
 * ตัวอย่าง:
 * - DS_PTX
 * - DS_FTX
 *
 * @returns
 * Promise ที่คืน Path ของไฟล์ผลลัพธ์
 *
 * หมายเหตุ:
 * - ถ้า Header ขาด ฟังก์ชันจะบันทึกไฟล์ก่อนแล้ว Throw Error
 * - ถ้า Field ว่าง ฟังก์ชันจะไม่ Throw Error และยังคืน Path
 */
export const validateTestData = async (
  inputFilePath: string,
  outputDirectory: string,
  resultBaseName: string,
  reportCode: TestDataReportCode,
): Promise<string> => {
  // แสดงหัวข้อการตรวจสอบ
  console.log(
    "\n===== TEST DATA VALIDATION =====",
  );

  // แสดง Report ที่กำลังตรวจสอบ
  console.log(
    `Report Code : ${reportCode}`,
  );

  // แสดงตำแหน่งไฟล์ Test Data ต้นทาง
  console.log(
    `Input File  : ${inputFilePath}`,
  );

  /**
   * ตรวจสอบว่าไฟล์ Test Data ต้นทางมีอยู่จริงหรือไม่
   *
   * ถ้าไม่พบไฟล์ จะหยุดการทำงานทันที
   * และยังไม่มีการสร้างไฟล์ผลลัพธ์
   */
  if (!fs.existsSync(inputFilePath)) {
    throw new Error(
      `Test Data input file not found: ${inputFilePath}`,
    );
  }

  /**
   * ตรวจสอบและสร้าง Folder สำหรับเก็บผลลัพธ์
   *
   * ถ้า Folder มีอยู่แล้วจะใช้งานต่อ
   * ถ้ายังไม่มีจะสร้าง Folder ขึ้นมา
   */
  ensureDirectoryExists(
    outputDirectory,
  );

  /**
   * สร้างชื่อไฟล์ผลลัพธ์พร้อม Timestamp
   *
   * ตัวอย่าง:
   * DS_PTX_TestData_Validation_Result_20260730_143000.xlsx
   */
  const finalOutputFilePath =
    buildTimestampedFilePath(
      outputDirectory,
      resultBaseName,
      ".xlsx",
    );

  /**
   * สร้าง Path ของไฟล์ชั่วคราวจากชื่อไฟล์จริง
   *
   * ตัวอย่าง:
   * ไฟล์จริง:
   * DS_PTX_TestData_Validation_Result_20260730_143000.xlsx
   *
   * ไฟล์ชั่วคราว:
   * DS_PTX_TestData_Validation_Result_20260730_143000_temp.xlsx
   *
   * ระบบจะเขียนผลลงไฟล์ชั่วคราวก่อน
   * เพื่อลดโอกาสได้ไฟล์จริงที่เขียนไม่สมบูรณ์
   */
  const tempOutputFilePath =
    buildTempFilePath(
      finalOutputFilePath,
    );

  /**
   * ลบไฟล์ชั่วคราวเก่าที่อาจค้างอยู่ใน Folder
   *
   * ตัวอย่างสาเหตุที่มีไฟล์ค้าง:
   * - การ Run ครั้งก่อนถูกปิดกลางคัน
   * - Excel เขียนไฟล์ไม่สำเร็จ
   * - Process ถูกหยุดก่อนเปลี่ยนชื่อไฟล์
   */
  cleanupStaleTempFiles(
    outputDirectory,
  );

  /**
   * สร้าง Workbook ใหม่ในหน่วยความจำ
   *
   * ตอนนี้ยังเป็น Workbook ว่าง
   */
  const workbook =
    new ExcelJS.Workbook();

  /**
   * อ่านไฟล์ Test Data ต้นทางเข้ามาใน Workbook
   *
   * await หมายถึงรอจนกว่า ExcelJS
   * จะอ่านไฟล์เสร็จก่อนทำขั้นตอนถัดไป
   */
  await workbook.xlsx.readFile(
    inputFilePath,
  );

  /**
   * ตรวจสอบว่าไฟล์ Test Data มี Worksheet แรกหรือไม่
   *
   * getWorksheet(1)
   * หมายถึง Worksheet ลำดับแรก
   * ไม่ได้หมายถึง Worksheet ที่ชื่อว่า "1"
   */
  if (!workbook.getWorksheet(1)) {
    throw new Error(
      "Original worksheet not found",
    );
  }

  // แสดง Path ของไฟล์ผลลัพธ์
  console.log(
    `Output File : ${finalOutputFilePath}`,
  );

  /**
   * ดึง Required Header ทั้งหมดตาม Report
   *
   * รายการนี้ใช้สำหรับตรวจสอบ Header
   * จึงรวมทั้ง
   * - Normal Required Header
   * - Fee Group Header
   *
   * ตัวอย่าง:
   * reportCode = DS_PTX
   * จะอ่าน Config ของ DS_PTX
   *
   * reportCode = DS_FTX
   * จะอ่าน Config ของ DS_FTX
   */
  const requiredHeaders =
    getRequiredTestDataHeaders(
      reportCode,
    );

  /**
   * ดึงเฉพาะ Header สำหรับตรวจข้อมูลแบบ Normal Field
   *
   * รายการนี้ไม่รวม 3 Field หลักของ Fee Group
   * เพราะ Fee Group จะถูกตรวจแยกใน
   * fee-group-validator.ts
   *
   * ตัวอย่าง Fee Group Header ที่ไม่รวม:
   * - Fee Type N
   * - Fee Charge Account No. Type N
   * - Fee Amount Type N หรือ Fee Amount N
   */
  const normalRequiredHeaders =
    getNormalRequiredTestDataHeaders(
      reportCode,
    );

  /**
   * ดึงหมายเลขแถว Header จาก Config ของ Report
   *
   * ปัจจุบัน DS_PTX และ DS_FTX
   * ใช้ Header Row แถวที่ 5
   *
   * แต่ไม่ได้เขียนเลข 5 ตายตัวในไฟล์นี้
   * ทำให้ Report ในอนาคตสามารถกำหนดแถวอื่นได้
   */
  const headerRowNumber =
    getTestDataHeaderRowNumber(
      reportCode,
    );

  // แสดงหมายเลขแถว Header
  console.log(
    `Header Row  : ${headerRowNumber}`,
  );

  // แสดงจำนวน Required Header ทั้งหมด
  console.log(
    `Required Header Count: ${requiredHeaders.length}`,
  );

  /**
   * ขั้นตอนที่ 1: ตรวจสอบข้อมูลภายใน Field
   *
   * ส่ง normalRequiredHeaders เข้าไป
   * เพื่อให้ Normal Field Validator
   * ตรวจเฉพาะ Field ข้อมูลทั่วไป
   *
   * ส่ง reportCode เข้าไปเพื่อกำหนดว่า
   * ต้องตรวจ Fee Group เพิ่มหรือไม่
   *
   * DS_PTX:
   * - ตรวจ Normal Field
   * - ตรวจ Fee Group
   *
   * DS_FTX:
   * - ตรวจ Normal Field
   * - ไม่ตรวจ Fee Group
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
   * ขั้นตอนที่ 2: ตรวจสอบ Header
   *
   * ใช้ requiredHeaders ซึ่งรวม Header
   * ทุกประเภทที่ Report ต้องมี
   *
   * ค่าที่ได้:
   * [] = พบ Header ครบทั้งหมด
   *
   * Array มีข้อมูล เช่น:
   * ["Currency Id", "Payment Method"]
   * = ไม่พบ Header ตามรายชื่อดังกล่าว
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
   * Workbook นี้อาจมีการเปลี่ยนแปลงจาก Field Validation เช่น
   * - Cell ถูก Highlight สีเขียว
   * - Cell ถูก Highlight สีแดง
   * - Cell ถูกใส่ข้อความ "โปรดกรอกข้อมูล"
   * - มี Sheet "Field Validation"
   */
  await workbook.xlsx.writeFile(
    tempOutputFilePath,
  );

  /**
   * เปลี่ยนชื่อไฟล์ชั่วคราวเป็นไฟล์ผลลัพธ์จริง
   *
   * หลังจากคำสั่งนี้ทำงานสำเร็จ
   * tempOutputFilePath จะไม่มีอยู่แล้ว
   * และจะเหลือ finalOutputFilePath
   */
  fs.renameSync(
    tempOutputFilePath,
    finalOutputFilePath,
  );

  // แจ้งว่าสร้างไฟล์ผลลัพธ์สำเร็จแล้ว
  console.log(
    "✅ Validation result file created successfully",
  );

  /**
   * ตรวจสอบผล Header Validation
   *
   * ถ้าพบ Missing Header:
   * 1. ไฟล์ผลลัพธ์ถูกบันทึกไว้แล้ว
   * 2. จากนั้นจึง Throw Error
   * 3. Script หรือ Test ที่เรียกใช้จะมีสถานะ Fail
   *
   * ตัวอย่าง Error:
   * Header Validation Failed:
   * Missing 2 Header(s): Currency Id, Payment Method
   */
  if (missingHeaders.length > 0) {
    throw new Error(
      `Header Validation Failed: Missing ${
        missingHeaders.length
      } Header(s): ${missingHeaders.join(", ")}`,
    );
  }

  /**
   * ตรวจสอบผล Field Validation
   *
   * ถ้าพบ Field ที่ไม่มีข้อมูล:
   * - แสดงข้อความ Failed ใน Console
   * - ไม่ Throw Error
   * - ยังคืน Path ของไฟล์ผลลัพธ์ตามปกติ
   *
   * ทำให้ Script สามารถนำไฟล์ผลลัพธ์
   * ไปตรวจสอบหรือใช้งานต่อได้
   */
  if (hasInvalidField) {
    console.log(
      "❌ Field Validation Failed",
    );

    return finalOutputFilePath;
  }

  /**
   * ถ้า Header ครบและไม่พบ Invalid Field
   * ให้แสดงว่าการตรวจสอบผ่าน
   */
  console.log(
    "✅ Test Data Validation Passed",
  );

  // คืน Path ของไฟล์ผลลัพธ์
  return finalOutputFilePath;
};