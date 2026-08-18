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
 * - DF_FXU
 * - DF_OLB
 * - DF_FXM
 *
 * ตัวอย่างคำสั่ง:
 * npm run test:script4 -- report=DS_PTX
 * npm run test:script4 -- report=DF_FXM
 * npm run test:script4 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU,DF_OLB,DF_FXM
 * ------------------------------------------------------------------
 */

import "dotenv/config";

import path from "path";

import {
  getSelectedReports,
} from "../resources/AF1-resources/config/report-selection";

import {
  getTestDataPath,
} from "../resources/AF1-resources/setting/uat/setting";

import {
  getLatestCheckedReportPath,
  getLatestCheckedTestDataPath,
  getLatestCompareResultPath,
  getSummaryResultOutputPath,
  getSummaryTemplatePath,
} from "../resources/AF1-resources/utils/summary/summary-file-helper";

import { getAutomationRunTimingSummary } from "../resources/AF1-resources/utils/summary/automation-run-timing";

import {
  readCompareResultRows,
  writeReportAutomationSummary,
} from "../resources/AF1-resources/utils/summary/automation-summary-writer";

/**
 * อ่านรายชื่อ Report จากค่า report
 *
 * ผู้ใช้ต้องระบุชื่อ Report ทุกครั้ง
 * หากไม่ระบุ ระบบจะแจ้ง Error และหยุดการทำงาน
 */
const selectedReports =
  getSelectedReports();

describe(
  "Script 4 - Summary Results",
  function () {
    /**
     * เพิ่มเวลาสูงสุดเป็น 5 นาที
     * สำหรับการสร้าง Summary หลาย Report
     */
    this.timeout(
      300000,
    );
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
          const startedAt =
            new Date();

          /**
           * อ่านเวลา Script 1-3 ของ Report ปัจจุบัน
           *
           * ถ้าขั้นตอนไหนยังไม่ได้รัน ระบบจะหยุดและแจ้งชื่อขั้นตอนที่ขาด
           */
          const automationTiming =
            getAutomationRunTimingSummary(
              reportName,
            );

          /**
          * ค้นหา Original Test Data จาก Share Path
          * ตามชื่อ Report ที่กำลังสร้าง Summary
          *
          * ภายในโฟลเดอร์ต้องมี Excel เพียง 1 ไฟล์
          * โดยชื่อไฟล์เป็นชื่ออะไรก็ได้
          */
          const originalTestDataPath =
            getTestDataPath(
              reportName,
            );
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
           * ดึง Timestamp จากชื่อไฟล์ Output
           *
           * ตัวอย่างชื่อไฟล์:
           * DS_FTX_Automation_Summary_20260818_144647-Final.xlsx
           *
           * ค่าที่ดึงได้:
           * 20260818_144647
           */
          const outputTimestamp =
            path
              .basename(outputPath)
              .match(
                /\d{8}_\d{6}/,
              )?.[0];

          /**
           * ป้องกันกรณีชื่อไฟล์ Output ไม่มี Timestamp
           */
          if (!outputTimestamp) {
            throw new Error(
              `Timestamp not found in Summary output file: ${outputPath}`,
            );
          }



          /**
           * ขั้นตอนที่ 2: อ่านผล Compare จาก Script 3
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
            originalTestDataPath,
            checkedReportPath,
            checkedTestDataPath,
            compareRows,
            {
              /**
              * ชื่อที่แสดงตรง Report File Name ใน Summary
              *
              * ตัวอย่าง:
              * DS_FTX_Summary_Test_Result_20260818_144647
              */
              reportFileName:
                `${reportName}_Summary_Test_Result_${outputTimestamp}`,

              /** เวลาเริ่มต้นจาก Script 1 */
              automationStartedAt:
                automationTiming
                  .automationStartedAt,

              /** เวลาเริ่ม Script 4 */
              script4StartedAt:
                startedAt,

              /** เวลารวม Script 1-3 */
              completedStageDurationMilliseconds:
                automationTiming
                  .completedStageDurationMilliseconds,

              runId:
                automationTiming.runId,

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
           * ขั้นตอนที่ 5: แสดงผลการทำงานใน Terminal
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
            originalTestDataPath,
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