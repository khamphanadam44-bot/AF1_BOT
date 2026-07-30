/**
 * Script 4 - Summary Results
 * ------------------------------------------------------------------
 * หน้าที่:
 * 1. อ่าน Compare Result ล่าสุดจาก Script 3
 * 2. อ่าน Original Test Data เพื่อสร้างข้อมูลฝั่ง Test Script Data
 * 3. อ่าน Checked Report และ Checked Test Data จาก Script 2
 * 4. เลือก Template ให้ตรงกับ Report
 * 5. สร้าง Automation Summary แยกตาม Report
 *
 * Report ที่รองรับ:
 * - DS_LTX
 * - DS_PTX
 * - DS_FTX
 * - DS_FTU
 *
 * ตัวอย่างคำสั่ง:
 * npm run test:script4 -- report=DS_PTX
 * npm run test:script4 -- report=DS_FTX
 * npm run test:script4 -- report=DS_LTX
 * npm run test:script4 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU
 * ------------------------------------------------------------------
 */

import path from "path";

import {
  getSelectedReports,
} from "../resources/AF1-resources/config/report-selection";

import {
  createRunId,
  getLatestCheckedReportPath,
  getLatestCheckedTestDataPath,
  getLatestCompareResultPath,
  getSummaryResultOutputPath,
  getSummaryTemplatePath,
  ORIGINAL_TEST_DATA_PATH,
} from "../resources/AF1-resources/utils/summary/summary-file-helper";

import {
  readCompareResultRows,
  writeReportAutomationSummary,
} from "../resources/AF1-resources/utils/summary/automation-summary-writer";

import {
  generateSummaryReport,
} from "../resources/AF1-resources/utils/validators/summary/summary.service";

/**
 * แปลงวันที่เป็นรูปแบบ yyyy-MM-dd
 */
const formatDate = (
  date: Date,
): string => {
  const yyyy = date.getFullYear();
  const MM = String(
    date.getMonth() + 1,
  ).padStart(2, "0");
  const dd = String(
    date.getDate(),
  ).padStart(2, "0");

  return `${yyyy}-${MM}-${dd}`;
};

/**
 * แปลงเวลาเป็นรูปแบบ HH:mm:ss
 */
const formatTime = (
  date: Date,
): string => {
  const HH = String(
    date.getHours(),
  ).padStart(2, "0");
  const mm = String(
    date.getMinutes(),
  ).padStart(2, "0");
  const ss = String(
    date.getSeconds(),
  ).padStart(2, "0");

  return `${HH}:${mm}:${ss}`;
};

/**
 * อ่านรายชื่อ Report จากค่า report
 *
 * ถ้าไม่ส่งค่า report ระบบจะใช้ Report เริ่มต้น
 * จาก dmsReportName ใน setting.ts
 */
const selectedReports =
  getSelectedReports();

describe(
  "Script 4 - Summary Results",
  () => {
    /**
     * สร้าง Test แยกหนึ่งชุดต่อหนึ่ง Report
     *
     * ตัวอย่าง:
     * report=DS_PTX,DS_FTX
     * จะสร้าง Test สำหรับ DS_PTX และ DS_FTX
     */
    for (
      const reportName of
      selectedReports
    ) {
      it(
        `Create ${reportName} Automation Summary`,
        async () => {
          /**
           * DS_LTX และ DS_FTU มีรูปแบบ Reconcile Result และ Template
           * ต่างจาก DS_PTX/DS_FTX จึงใช้ Summary Service ของ LTX
           */
          if (
            reportName === "DS_LTX" ||
            reportName === "DS_FTU"
          ) {
            const summaryResult =
              await generateSummaryReport(
                reportName,
              );

            console.log(
              "Output File       :",
              summaryResult.summaryFilePath,
            );

            return;
          }

          const startedAt =
            new Date();

          /**
           * ขั้นตอนที่ 1:
           * หาไฟล์ต้นทางทั้งหมดที่ Script 4 ต้องใช้
           */
          const compareResultPath =
            getLatestCompareResultPath(
              reportName,
            );

          const checkedReportPath =
            getLatestCheckedReportPath(
              reportName,
            );

          const checkedTestDataPath =
            getLatestCheckedTestDataPath(
              reportName,
            );

          const templatePath =
            getSummaryTemplatePath(
              reportName,
            );

          const outputPath =
            getSummaryResultOutputPath(
              reportName,
            );

          /**
           * ขั้นตอนที่ 2:
           * อ่านผล Compare จาก Script 3
           */
          const compareRows =
            await readCompareResultRows(
              compareResultPath,
            );

          /**
           * ขั้นตอนที่ 3:
           * นับผล PASS / FAIL / SKIP
           * เพื่อนำไปแสดงในส่วนสรุปด้านบนของไฟล์
           */
          const totalPass =
            compareRows.filter(
              (row) =>
                row.status ===
                "PASS",
            ).length;

          const totalFail =
            compareRows.filter(
              (row) =>
                row.status ===
                "FAIL",
            ).length;

          const totalSkip =
            compareRows.filter(
              (row) =>
                row.status ===
                "SKIP",
            ).length;

          /**
           * ขั้นตอนที่ 4:
           * สร้างไฟล์ Automation Summary
           *
           * แหล่งข้อมูลของแต่ละชีท:
           * - Summary Test Results = Compare Result + Original Test Data
           * - <REPORT>_Reconcile  = Compare Result จาก Script 3
           * - <REPORT>            = Checked Report จาก Script 2
           * - Test Data           = Checked Test Data จาก Script 2
           */
          await writeReportAutomationSummary(
            reportName,
            templatePath,
            outputPath,
            compareResultPath,
            ORIGINAL_TEST_DATA_PATH,
            checkedReportPath,
            checkedTestDataPath,
            compareRows,
            {
              reportFileName:
                path.basename(
                  compareResultPath,
                ),

              executionDate:
                formatDate(
                  startedAt,
                ),

              executionTime:
                formatTime(
                  startedAt,
                ),

              runId:
                createRunId(),

              verifiedBy:
                "QAD Automation",

              totalChecked:
                compareRows.length,

              passed:
                totalPass,

              failed:
                totalFail,
            },
          );

          /**
           * ขั้นตอนที่ 5:
           * แสดงผลการทำงานใน Terminal
           */
          console.log("");
          console.log(
            "======================================",
          );
          console.log(
            " Script 4 Summary Result",
          );
          console.log(
            "======================================",
          );
          console.log(
            "Report Name       :",
            reportName,
          );
          console.log(
            "Template File     :",
            templatePath,
          );
          console.log(
            "Compare File      :",
            compareResultPath,
          );
          console.log(
            "Original Test Data:",
            ORIGINAL_TEST_DATA_PATH,
          );
          console.log(
            "Checked Report    :",
            checkedReportPath,
          );
          console.log(
            "Checked Test Data :",
            checkedTestDataPath,
          );
          console.log(
            "Total Checked     :",
            compareRows.length,
          );
          console.log(
            "PASS              :",
            totalPass,
          );
          console.log(
            "FAIL              :",
            totalFail,
          );
          console.log(
            "SKIP              :",
            totalSkip,
          );
          console.log(
            "Output File       :",
            outputPath,
          );
          console.log(
            "======================================",
          );
        },
      );
    }
  },
);
