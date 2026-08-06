/**
 * ======================================================
 * ไฟล์: script2-validate-report-and-testdata.spec.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * Script 2 ตรวจ Header ของ Raw Report และตรวจ Header กับข้อมูลบังคับใน Test Data
 * ผลลัพธ์จะถูกบันทึกแยกตาม Report เพื่อให้ Script 3 และ Script 4 เลือกไฟล์ที่ตรวจแล้วไปใช้ต่อ
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import "dotenv/config";

import * as path from "path";

import {
  getCheckedReportHeaderDir,
  getCheckedTestDataHeaderDir,
  getRawReportDir,
  getTestDataResultBasename,
} from "../resources/AF1-resources/config/paths.config";

import {
  detectReportCodeFromFileName,
} from "../resources/AF1-resources/config/report-detector";

import {
  TESTDATA_CONFIG,
} from "../resources/AF1-resources/config/testdata-config";

import type {
  TestDataReportCode,
} from "../resources/AF1-resources/config/testdata-config";

import {
  getSelectedReports,
} from "../resources/AF1-resources/config/report-selection";

import {
  getAliases,
  getHeaderRowNumber,
  getRequiredReportHeaders,
} from "../resources/AF1-resources/config/report-helper";

import {
  getTestDataPath,
} from "../resources/AF1-resources/setting/uat/setting";

import {
  copyFileToDirectory,
  getLatestFile,
} from "../resources/AF1-resources/utils/file-system.util";

import {
  validateReportHeader,
} from "../resources/AF1-resources/utils/validators/download-report/report-header-validator";

import {
  validateTestData,
} from "../resources/AF1-resources/utils/validators/test-data-report/test-data-validator";

/**
 * ตรวจสอบ Report 1 รายการ
 *
 * ฟังก์ชันนี้ไม่ได้ล็อกชื่อ Report ไว้กับ DS_PTX หรือ DS_FTX
 * ชื่อ Report จะถูกส่งเข้ามาจากค่า report
 *
 * ตัวอย่าง:
 * validateSelectedReport("DS_PTX");
 * validateSelectedReport("DS_FTX");
 */
const validateSelectedReport =
  async (
    selectedReport: string,
  ): Promise<void> => {
    /**
     * ขั้นตอนที่ 1:
     * หาโฟลเดอร์ Raw Report
     *
     * ตัวอย่าง:
     * test_data/AF1_Report/DS_PTX
     * test_data/AF1_Report/DS_FTX
     */
    const rawReportDirectory =
      getRawReportDir(
        selectedReport,
      );

    console.log(
      "========================================",
    );

    console.log(
      "Selected Report :",
      selectedReport,
    );

    console.log(
      "Raw Report Folder :",
      rawReportDirectory,
    );

    /**
     * ขั้นตอนที่ 2:
     * หาไฟล์ล่าสุดในโฟลเดอร์ของ Report
     */
    const latestRawReportFilePath =
      getLatestFile(
        rawReportDirectory,
      );

    console.log(
      "Latest Raw Report :",
      latestRawReportFilePath,
    );

    /**
     * ขั้นตอนที่ 3:
     * ตรวจจับชื่อ Report จากชื่อไฟล์จริง
     *
     * ตัวอย่างชื่อไฟล์:
     * EXPORT_DS_PTX_20251125_20260722_120000.xlsx
     * EXPORT_DS_FTX_20251125_20260722_120000.xlsx
     */
    const detectedReportCode =
      detectReportCodeFromFileName(
        path.basename(
          latestRawReportFilePath,
        ),
      );

    console.log(
      "Detected Report :",
      detectedReportCode,
    );

    /**
     * ขั้นตอนที่ 4:
     * ป้องกันการนำไฟล์ผิด Report มาตรวจ
     *
     * ตัวอย่าง:
     * ผู้ใช้เลือก DS_FTX
     * แต่ไฟล์ที่พบเป็น DS_PTX
     */
    if (
      detectedReportCode !==
      selectedReport
    ) {
      throw new Error(
        [
          "Report mismatch.",
          `Selected Report: ${selectedReport}`,
          `Detected Report: ${detectedReportCode}`,
          `File: ${latestRawReportFilePath}`,
        ].join("\n"),
      );
    }

    /**
     * ขั้นตอนที่ 5:
     * อ่าน Config ตามชื่อ Report
     *
     * Report แต่ละตัวสามารถมี:
     * - Required Headers ต่างกัน
     * - Header Row ต่างกัน
     * - Aliases ต่างกัน
     */
    const requiredHeaders =
      getRequiredReportHeaders(
        detectedReportCode,
      );

    const headerRowNumber =
      getHeaderRowNumber(
        detectedReportCode,
      );

    const aliases =
      getAliases(
        detectedReportCode,
      );

    console.log(
      "Header Row Number :",
      headerRowNumber,
    );

    console.log(
      "Required Header Count :",
      requiredHeaders.length,
    );

    /**
     * ขั้นตอนที่ 6:
     * Copy Raw Report ไปยังโฟลเดอร์สำหรับตรวจ
     *
     * ตัวอย่าง:
     * Test_result/Checked-report-header/DS_PTX
     * Test_result/Checked-report-header/DS_FTX
     */
    const checkedReportDirectory =
      getCheckedReportHeaderDir(
        detectedReportCode,
      );

    const reportFileForValidation =
      copyFileToDirectory(
        latestRawReportFilePath,
        checkedReportDirectory,
      );

    console.log(
      "Checked Report Folder :",
      checkedReportDirectory,
    );

    console.log(
      "Report File For Validation :",
      reportFileForValidation,
    );

    /**
     * ขั้นตอนที่ 7:
     * ตรวจและ Highlight Header ของ Report
     */
    await validateReportHeader(
      reportFileForValidation,
      requiredHeaders,
      {
        headerRowNumber,
        aliases,
      },
    );

    console.log(
      `Report Header Validation Complete: ${detectedReportCode}`,
    );

    console.log(
      "========================================",
    );
  };
/**
 * ตรวจว่า Report มี Test Data Config หรือไม่
 *
 * ตอนนี้รองรับ:
 * - DS_PTX
 * - DS_FTX
 */
const isTestDataReportCode = (
  reportCode: string,
): reportCode is TestDataReportCode => {
  return Object.prototype.hasOwnProperty.call(
    TESTDATA_CONFIG,
    reportCode,
  );
};
/**
 * Script 2
 *
 * หน้าที่:
 * 1. อ่านรายชื่อ Report จากค่า report
 * 2. ตรวจ Report Header ทีละ Report
 * 3. ตรวจ Test Data และแยกผลลัพธ์ตามแต่ละ Report
 */
describe(
  "Script 2 - Validate Report Header and Test Data",
  function () {
    /**
     * เพิ่มเวลาให้ Script 2 สูงสุด 5 นาที
     * สำหรับกรณีเลือกตรวจหลาย Report พร้อมกัน
     */
    this.timeout(
      300000,
    );

    /**
     * อ่านรายชื่อ Report ที่ผู้ใช้เลือก
     *
     * ตัวอย่าง:
     * report=DS_PTX
     * report=DS_FTX
     * report=DS_PTX,DS_FTX
     */
    const selectedReports =
      getSelectedReports();

    /**
     * ป้องกันกรณีไม่มีชื่อ Report
     */
    if (
      selectedReports.length === 0
    ) {
      throw new Error(
        [
          "No Report was selected.",
          "Please specify at least one Report.",
          "Example: --reports=DS_PTX",
        ].join("\n"),
      );
    }

    console.log(
      "Selected Reports :",
      selectedReports.join(
        ", ",
      ),
    );

    /**
     * สร้าง Test Case สำหรับ Report แต่ละรายการ
     *
     * เมื่อเพิ่ม Report ใหม่ในอนาคต
     * ไม่ต้องเพิ่ม if หรือเพิ่ม it() ในไฟล์นี้
     */
    for (
      const selectedReport
      of selectedReports
    ) {
      it(
        `Validate latest ${selectedReport} Report header`,
        async () => {
          await validateSelectedReport(
            selectedReport,
          );
        },
      );
    }

    /**
 * เลือกเฉพาะ Report ที่มี Test Data Config
 *
 * ตอนนี้คือ:
 * - DS_PTX
 * - DS_FTX
 *
 * Report อื่นยังคงตรวจ Report Header ได้
 * แต่จะยังไม่ตรวจ Test Data
 */
    const selectedTestDataReports =
      selectedReports.filter(
        isTestDataReportCode,
      );

    const reportsWithoutTestDataConfig =
      selectedReports.filter(
        (reportCode) =>
          !isTestDataReportCode(
            reportCode,
          ),
      );

    if (
      reportsWithoutTestDataConfig.length >
      0
    ) {
      console.log(
        "Skip Test Data Validation :",
        reportsWithoutTestDataConfig.join(
          ", ",
        ),
      );

      console.log(
        "Reason: Test Data Config not found",
      );
    }

    /**
     * ตรวจ Test Data แยกตาม Report
     */
    for (
      const selectedReport
      of selectedTestDataReports
    ) {
      it(
        `Validate ${selectedReport} Test Data header and fields`,
        async () => {
          console.log(
            "========================================",
          );

          console.log(
            "Validate Test Data For Report :",
            selectedReport,
          );

          /**
           * ค้นหา Test Data จาก Share Path
           * ตามชื่อ Report ที่กำลังตรวจ
           */
          const testDataPath =
            getTestDataPath(
              selectedReport,
            );

          console.log(
            "Test Data File :",
            testDataPath,
          );

          /**
           * สร้างโฟลเดอร์เก็บผลลัพธ์
           *
           * ตัวอย่าง:
           * Checked-testdata-header/DS_PTX
           * Checked-testdata-header/DS_FTX
           */
          const checkedTestDataDirectory =
            getCheckedTestDataHeaderDir(
              selectedReport,
            );

          /**
           * สร้างชื่อไฟล์ผลลัพธ์
           *
           * ตัวอย่าง:
           * DS_PTX_TestData_Validation_Result
           */
          const testDataResultBasename =
            getTestDataResultBasename(
              selectedReport,
            );

          console.log(
            "Checked Test Data Folder :",
            checkedTestDataDirectory,
          );

          console.log(
            "Test Data Result Basename :",
            testDataResultBasename,
          );

          /**
           * ส่ง Argument ตัวที่ 4 คือ selectedReport
           *
           * เพื่อให้ Test Data Validator
           * เลือก Config และ Logic ถูก Report
           */
          await validateTestData(
            testDataPath,
            checkedTestDataDirectory,
            testDataResultBasename,
            selectedReport,
          );

          console.log(
            `Test Data Validation Complete: ${selectedReport}`,
          );

          console.log(
            "========================================",
          );
        },
      );
    }
  },
);
