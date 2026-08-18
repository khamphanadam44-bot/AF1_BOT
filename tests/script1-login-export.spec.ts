/**
 * ======================================================
 * ไฟล์: script1-login-export.spec.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * Script 1 เข้าสู่ระบบ DMS แล้ว Export Report ตามรายชื่อที่ผู้ใช้เลือกจาก Terminal
 * แต่ละ Report จะถูกดาวน์โหลด อ่านจำนวนข้อมูล และบันทึกไว้ในโฟลเดอร์ Raw Report ของ Report นั้น
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import "dotenv/config";

import {
  chromium,
} from "playwright";

import type {
  Browser,
} from "playwright";

import {
  LoginFeature,
} from "../resources/AF1-resources/features/login.feature";

import {
  ExportFeature,
} from "../resources/AF1-resources/features/export.feature";

import {
  getReportDateRange,
  webSetting,
} from "../resources/AF1-resources/setting/uat/setting";

import {
  getSelectedReports,
} from "../resources/AF1-resources/config/report-selection";

import {
  recordAutomationRunStage,
  startAutomationRun,
} from "../resources/AF1-resources/utils/summary/automation-run-timing";

describe(
  "Script 1 - Login + Export Selected Reports",
  function () {
    /**
     * เพิ่มเวลาเป็น 5 นาที
     *
     * เพราะเมื่อเลือกหลาย Report
     * ระยะเวลารวมอาจเกิน 120 วินาที
     */
    this.timeout(
      300000,
    );

    let browser: Browser;

    /**
     * อ่านรายชื่อ Report จาก Terminal
     *
     * ตัวอย่าง:
     * report=DS_PTX,DS_FTX
     */
    const selectedReports =
      getSelectedReports();

    console.log(
      "=== SELECTED REPORTS ===",
      selectedReports,
    );

    before(
      async function () {
        browser =
          await chromium.launch({
            headless: false,
          });
      },
    );

    after(
      async function () {
        if (
          browser
        ) {
          await browser.close();
        }
      },
    );

    it(
      "Login success and Export selected Reports",
      async function () {
        /**
         * วน Export Report ทีละรายการ
         *
         * แต่ละ Report จะ:
         * 1. เปิด Browser Context ใหม่
         * 2. เปิด Page ใหม่
         * 3. Login ใหม่
         * 4. Export Report
         * 5. ปิด Context หลัง Export เสร็จ
         */
        for (
          const selectedReport
          of selectedReports
        ) {
          /**
           * Script 1 เป็นจุดเริ่ม Automation Run ใหม่
           * โดยแยก Timing State ตามชื่อ Report
           */
          const script1StartedAt =
            new Date();

          startAutomationRun(
            selectedReport,
            script1StartedAt,
          );

          const context =
            await browser.newContext({
              acceptDownloads: true,
            });

          const page =
            await context.newPage();

          const loginFeature =
            new LoginFeature(
              page,
            );

          const exportFeature =
            new ExportFeature(
              page,
            );

          try {
            console.log(
              "========================================",
            );

            console.log(
              `Start Export Report: ${selectedReport}`,
            );

            /**
             * เปิดหน้า Login ใหม่
             */
            await page.goto(
              webSetting.url,
              {
                waitUntil:
                  "domcontentloaded",

                timeout:
                  60000,
              },
            );

            /**
             * Login ใหม่สำหรับ Report นี้
             */
            await loginFeature.loginPass(
              webSetting.username,
              webSetting.password,
            );

            /**
             * รอหน้าหลักหลัง Login
             */
            await page.waitForLoadState(
              "domcontentloaded",
            );

            await page.waitForTimeout(
              2000,
            );

            /**
           * อ่านช่วงวันที่ตาม Report ที่กำลัง Export
           *
           * DF_FXM:
           * - Date From = 26/02/2026
           * - Date To   = 27/02/2026
           *
           * Report อื่น:
           * - ใช้ช่วงวันที่เริ่มต้นจาก setting.ts
           */
            const reportDateRange =
              getReportDateRange(
                selectedReport,
              );

            console.log(
              "Export Date From :",
              reportDateRange.dateset,
            );

            console.log(
              "Export Date To   :",
              reportDateRange.dateto,
            );

            /**
             * ส่งออก Report ด้วยช่วงวันที่
             * ที่กำหนดไว้สำหรับ Report นั้น
             */
            const exportResult =
              await exportFeature.exportPass(
                selectedReport,
                reportDateRange.dateset,
                reportDateRange.dateto,
              );

            console.log(
              `Export Complete: ${selectedReport}`,
            );

            console.log(
              "Raw Report Path :",
              exportResult.savePath,
            );

            console.log(
              "Report Rows :",
              exportResult.reportData.length,
            );

          } finally {
            /**
             * ปิดหน้าจอของ Report ปัจจุบัน
             *
             * เมื่อวนรอบถัดไป
             * ระบบจะเปิด Context และ Page ใหม่
             */
            await context.close();

            console.log(
              `Browser Context Closed: ${selectedReport}`,
            );

            /**
             * บันทึกเวลาของ Script 1 เฉพาะ Report ปัจจุบัน
             */
            recordAutomationRunStage(
              selectedReport,
              "SCRIPT_1_EXPORT",
              script1StartedAt,
              new Date(),
            );
          }
        }
      },
    );
  },
);