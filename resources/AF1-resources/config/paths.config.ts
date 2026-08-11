/**
 * paths.config.ts
 * ------------------------------------------------------------------
 * รวม Path หลักที่ใช้ทั้งระบบไว้ที่เดียว
 * ------------------------------------------------------------------
 */

import * as path from "path";

/**
 * แปลงชื่อ Report ให้อยู่ในรูปแบบมาตรฐาน
 * และป้องกันการนำอักขระที่ไม่ควรอยู่ใน Path มาใช้งาน
 */
export const normalizeReportCode = (
  reportCode: string,
): string => {
  const normalizedReportCode =
    String(
      reportCode ?? "",
    )
      .trim()
      .toUpperCase();

  if (
    !normalizedReportCode
  ) {
    throw new Error(
      "reportCode is empty.",
    );
  }

  if (
    !/^[A-Z0-9_]+$/.test(
      normalizedReportCode,
    )
  ) {
    throw new Error(
      `Invalid reportCode: "${reportCode}". ` +
      "Only A-Z, 0-9 and underscore are allowed.",
    );
  }

  return normalizedReportCode;
};

/**
 * โฟลเดอร์หลักสำหรับเก็บ Raw Report
 *
 * ตัวอย่าง:
 * test_data/AF1_Report
 */
export const AF1_REPORT_RAW_DIR = path.join(
  "test_data",
  "AF1_Report",
);

/**
 * คืนค่าโฟลเดอร์ Raw Report
 * โดยแยกตามชื่อ Report
 *
 * ตัวอย่าง:
 * test_data/AF1_Report/DS_PTX
 * test_data/AF1_Report/DS_FTX
 */
export const getRawReportDir = (
  reportCode: string,
): string => {
  return path.join(
    AF1_REPORT_RAW_DIR,
    reportCode.toUpperCase(),
  );
};

/**
 * คืนค่าโฟลเดอร์เก็บ Report
 * ที่ผ่านการ Validate Header และ Highlight แล้ว
 *
 * ตัวอย่าง:
 * Test_result/Checked-report-header/DS_PTX
 * Test_result/Checked-report-header/DS_FTX
 */
export const getCheckedReportHeaderDir = (
  reportCode: string,
): string => {
  return path.join(
    "Test_result",
    "Checked-report-header",
    reportCode.toUpperCase(),
  );
};

/**
 * โฟลเดอร์หลักสำหรับเก็บผลตรวจ Test Data
 *
 * Constant นี้เป็นโฟลเดอร์แม่
 * ยังไม่ได้ระบุชื่อ Report
 *
 * ผลลัพธ์:
 * Test_result/Checked-testdata-header
 */
export const CHECKED_TEST_DATA_DIR = path.join(
  "Test_result",
  "Checked-testdata-header",
);

/**
 * คืนค่าโฟลเดอร์เก็บผลตรวจ Test Data
 * โดยแยกตามชื่อ Report
 *
 * ตัวอย่าง:
 * DS_PTX:
 * Test_result/Checked-testdata-header/DS_PTX
 *
 * DS_FTX:
 * Test_result/Checked-testdata-header/DS_FTX
 *
 * Report ใหม่ในอนาคต เช่น DS_FTU:
 * Test_result/Checked-testdata-header/DS_FTU
 */
export const getCheckedTestDataHeaderDir = (
  reportCode: string,
): string => {
  return path.join(
    CHECKED_TEST_DATA_DIR,
    reportCode.toUpperCase(),
  );
};

/**
 * ชื่อ Alias ที่โมดูล Summary ของ DS_LTX เรียกใช้
 */
export const getCheckedTestDataDir = (
  reportCode: string,
): string => {
  return getCheckedTestDataHeaderDir(
    reportCode,
  );
};

/**
 * คืนค่าโฟลเดอร์ผลลัพธ์จาก Script 3
 */
export const getReconcileOutputDir = (
  reportCode: string,
): string => {
  return path.resolve(
    process.cwd(),
    "Test_result",
    "Reconcile-report",
    normalizeReportCode(
      reportCode,
    ),
  );
};

/**
 * คืนค่าโฟลเดอร์ผลลัพธ์จาก Script 4
 */
export const getSummaryOutputDir = (
  reportCode: string,
): string => {
  return path.resolve(
    process.cwd(),
    "Test_result",
    "Summary-report",
    normalizeReportCode(
      reportCode,
    ),
  );
};

/**
 * สร้างชื่อเริ่มต้นของไฟล์ผลตรวจ Test Data
 * ตามชื่อ Report
 *
 * ตัวอย่าง:
 * getTestDataResultBasename("DS_PTX")
 * ผลลัพธ์:
 * DS_PTX_TestData_Validation_Result
 *
 * getTestDataResultBasename("DS_FTX")
 * ผลลัพธ์:
 * DS_FTX_TestData_Validation_Result
 */
export const getTestDataResultBasename = (
  reportCode: string,
): string => {
  return `${reportCode.toUpperCase()}_TestData_Validation_Result`;
};

/**
 * คืน Path โฟลเดอร์ Test Data ตามชื่อ Report
 *
 * โครงสร้าง:
 * AF1_SHAREPATH/af1_test_data/<REPORT>
 *
 * ตัวอย่าง:
 * AF1_SHAREPATH/af1_test_data/DS_PTX
 * AF1_SHAREPATH/af1_test_data/DS_FTX
 *
 * ฟังก์ชันนี้สร้างเฉพาะ Path ของโฟลเดอร์
 * ยังไม่ได้ค้นหาไฟล์ Excel ภายในโฟลเดอร์
 */
export const getTestDataInputDir = (
  sharePath: string,
  reportCode: string,
): string => {
  /**
   * ตัดช่องว่างหน้าและหลังของ Share Path
   */
  const normalizedSharePath =
    String(
      sharePath ?? "",
    ).trim();

  /**
   * ป้องกันกรณีไม่ได้กำหนด Share Path
   */
  if (
    !normalizedSharePath
  ) {
    throw new Error(
      "AF1_SHAREPATH is empty.",
    );
  }

  /**
   * สร้าง Path โดยใช้ชื่อ Report
   * ที่รับเข้ามาจาก Terminal
   *
   * ไม่มีการ Hard code ชื่อ Report
   */
  return path.join(
    normalizedSharePath,
    "af1_test_data",
    normalizeReportCode(
      reportCode,
    ),
  );
};


