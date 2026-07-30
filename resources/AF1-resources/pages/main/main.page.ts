/**
 * ======================================================
 * ไฟล์: main.page.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * จัดการการทำงานบนหน้าหลักของ DMS เช่น
 * - เปิดเมนู DMS
 * - เปิดเมนู DATASET EXPORT
 * - กรอกชื่อ Report
 * - เลือก Report
 * - กรอกวันที่
 * - Export และบันทึกไฟล์
 * ======================================================
 */

import {
  existsSync,
} from "fs";

import * as path from "path";

import {
  Page,
} from "playwright";

import {
  AF1_REPORT_RAW_DIR,
} from "../../config/paths.config";

import {
  ensureDirectoryExists,
} from "../../utils/file-system.util";

import {
  mainLocator,
} from "./main.locator";

export class MainPage {
  private readonly page: Page;

  constructor(
    page: Page,
  ) {
    this.page = page;
  }

  /**
   * คลิกเมนู DMS
   */
  async clickDmsMenu(): Promise<void> {
    const dmsMenu =
      this.page.locator(
        mainLocator.dmsMenu,
      );

    await dmsMenu.click();
  }

  /**
   * คลิกเมนู DMM Data Model Export
   */
  async clickDmmMenu(): Promise<void> {
    const dmmMenu =
      this.page.locator(
        mainLocator.dmmDataModelExport,
      );

    await dmmMenu.click();
  }

  /**
   * คลิกเมนู DATASET EXPORT
   */
  async clickDataSetMenu(): Promise<void> {
    const dataSetMenu =
      this.page.locator(
        mainLocator.dataSetExport,
      );

    await dataSetMenu.click();
  }

  /**
   * กรอกชื่อ Report ลงในช่องค้นหา
   *
   * ตัวอย่าง:
   * - DS_PTX
   * - DS_FTX
   * - DS_LTX
   * - DS_FTU
   */
  async inputReportName(
    reportName: string,
  ): Promise<void> {
    const reportNameInput =
      this.page.locator(
        mainLocator.reportNameInput,
      );

    await reportNameInput.fill(
      reportName,
    );
  }

  /**
   * ดับเบิลคลิก Report ตามชื่อที่ได้รับ
   *
   * Report ทั่วไปใช้ reportByName:
   * - DS_PTX
   * - DS_FTX
   * - DS_LTX
   *
   * DS_FTU ใช้ Locator เฉพาะ:
   * //td[contains(text(), 'DS_FTU Transaction')]
   */
  async clickDSReport(
    reportName: string,
  ): Promise<void> {
    /**
     * เลือก Locator ตาม Report
     *
     * ถ้าเป็น DS_FTU:
     * ใช้ dsFtuReport
     *
     * ถ้าเป็น Report อื่น:
     * ใช้ reportByName
     */
    const reportLocator =
      reportName === "DS_FTU"
        ? mainLocator.dsFtuReport
        : mainLocator.reportByName(
            reportName,
          );

    const report =
      this.page.locator(
        reportLocator,
      );

    /**
     * รอจนกระทั่งชื่อ Report
     * แสดงบนหน้าจอ
     */
    await report.waitFor({
      state: "visible",
      timeout: 30000,
    });

    /**
     * ดับเบิลคลิก Report
     */
    await report.dblclick();
  }

  /**
   * กรอก Date From
   */
  async inputDate(
    dateFrom: string,
  ): Promise<void> {
    const dateFromInput =
      this.page.locator(
        mainLocator.dateFromTxt,
      );

    await dateFromInput.fill(
      dateFrom,
    );
  }

  /**
   * กรอก Date To
   */
  async inputDateTo(
    dateTo: string,
  ): Promise<void> {
    const dateToInput =
      this.page.locator(
        mainLocator.dateToTxt,
      );

    await dateToInput.fill(
      dateTo,
    );
  }

  /**
   * คลิกปุ่ม Export และบันทึกไฟล์
   * แยกโฟลเดอร์ตามชื่อ Report
   *
   * ตัวอย่าง:
   * test_data/AF1_Report/DS_PTX/EXPORT_DS_PTX_....xlsx
   * test_data/AF1_Report/DS_FTU/EXPORT_DS_FTU_....xlsx
   */
  async clickExportButton(
    reportName: string,
  ): Promise<string> {
    const normalizedReportName =
      reportName.trim();

    /**
     * ตรวจสอบชื่อ Report
     */
    if (
      !normalizedReportName ||
      !/^[A-Za-z0-9_-]+$/.test(
        normalizedReportName,
      )
    ) {
      throw new Error(
        `Invalid report name: "${reportName}"`,
      );
    }

    const exportButton =
      this.page.locator(
        mainLocator.exportButton,
      );

    /**
     * รอให้ปุ่ม Export แสดงบนหน้าจอ
     */
    await exportButton.waitFor({
      state: "visible",
      timeout: 30000,
    });

    /**
     * ตรวจสอบว่าปุ่ม Export
     * สามารถกดได้หรือไม่
     */
    if (
      !(await exportButton.isEnabled())
    ) {
      throw new Error(
        `Export button is disabled for Report: ${normalizedReportName}`,
      );
    }

    console.log(
      `=== CLICK EXPORT: ${normalizedReportName} ===`,
    );

    /**
     * เริ่มรอ Download ก่อนกดปุ่ม Export
     *
     * Timeout 300000 milliseconds
     * เท่ากับ 5 นาที
     */
    const downloadPromise =
      this.page.waitForEvent(
        "download",
        {
          timeout: 300000,
        },
      );

    /**
     * คลิกปุ่ม Export
     */
    await exportButton.click();

    console.log(
      "Waiting for download...",
    );

    /**
     * รอรับไฟล์ Download
     */
    const download =
      await downloadPromise;

    /**
     * ตรวจสอบว่าการ Download
     * ล้มเหลวหรือไม่
     */
    const downloadFailure =
      await download.failure();

    if (
      downloadFailure
    ) {
      throw new Error(
        `Download Failed: ${downloadFailure}`,
      );
    }

    /**
     * อ่านชื่อไฟล์ที่ Download
     */
    const rawFileName =
      download.suggestedFilename();

    /**
     * ป้องกันไม่ให้ชื่อไฟล์
     * มี Path ติดมาด้วย
     */
    const fileName =
      path.basename(
        rawFileName,
      );

    /**
     * ตรวจสอบชื่อไฟล์
     */
    if (
      !fileName ||
      fileName === "." ||
      fileName === ".."
    ) {
      throw new Error(
        `Invalid download file name: "${rawFileName}"`,
      );
    }

    console.log(
      "=== FILE NAME ===",
      fileName,
    );

    /**
     * สร้างตำแหน่งโฟลเดอร์ของ Report
     *
     * ตัวอย่าง:
     * test_data/AF1_Report/DS_PTX
     * test_data/AF1_Report/DS_FTU
     */
    const reportDirectory =
      path.join(
        AF1_REPORT_RAW_DIR,
        normalizedReportName,
      );

    /**
     * สร้างโฟลเดอร์ ถ้ายังไม่มี
     */
    ensureDirectoryExists(
      reportDirectory,
    );

    /**
     * สร้างตำแหน่งไฟล์ปลายทาง
     */
    const savePath =
      path.join(
        reportDirectory,
        fileName,
      );

    /**
     * บันทึกไฟล์ลงในโฟลเดอร์
     */
    await download.saveAs(
      savePath,
    );

    /**
     * ตรวจสอบว่าไฟล์ถูกบันทึกจริง
     */
    if (
      !existsSync(
        savePath,
      )
    ) {
      throw new Error(
        `Download file was not found: ${savePath}`,
      );
    }

    console.log(
      "Download Success",
    );

    console.log(
      "Report Folder :",
      reportDirectory,
    );

    console.log(
      "=== RAW REPORT SAVE COMPLETE ===",
      savePath,
    );

    return savePath;
  }
}