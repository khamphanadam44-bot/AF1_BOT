/**
 * ReconcileWorkbookPreparer
 * ------------------------------------------------------------------
 * เตรียม Workbook สำหรับ Script 3:
 * 1. หา Checked AF1 Report ล่าสุดตาม Report Code
 * 2. Copy Checked Report ไปยัง Reconcile Output
 * 3. เปิด Workbook ที่ Copy แล้ว
 * 4. สร้าง Worksheet สำหรับเขียนผล Reconcile
 * ------------------------------------------------------------------
 */

import * as fs from "fs";
import ExcelJS from "exceljs";

import {
  getCheckedReportHeaderDir,
  getReconcileOutputDir,
  normalizeReportCode,
} from "../../../config/paths.config";

import {
  buildTimestampedFilePath,
  ensureDirectoryExists,
  getLatestFile,
} from "../../file-system.util";

import { getHeadersFromRow } from "../../validators/shared/excel-cell.util";

import { ReconcileResultSheetWriter } from "./result-writer";

export interface PreparedReconcileWorkbook {
  workbook: ExcelJS.Workbook;
  reportWorksheet: ExcelJS.Worksheet;
  resultSheet: ExcelJS.Worksheet;
  reportHeaders: string[];
  reconcileFilePath: string;
}

export class ReconcileWorkbookPreparer {
  constructor(
    private readonly sheetWriter: ReconcileResultSheetWriter = new ReconcileResultSheetWriter(),
  ) {}

  /**
   * Copy Checked AF1 Report ล่าสุด
   * ไปสร้างเป็นไฟล์ Reconcile Result
   *
   * @param reportCode Report ที่ต้องการประมวลผล เช่น DS_LTX
   * @returns Path ของไฟล์ Reconcile Result
   */
  private copyLatestCheckedReport(reportCode: string): string {
    const normalizedReportCode = normalizeReportCode(reportCode);

    /**
     * หา Checked AF1 Report ล่าสุด
     * เฉพาะ Report Code ที่กำลังรัน
     *
     * ตัวอย่าง:
     * Test_result/Checked-report-header/DS_LTX/EXPORT_DS_LTX_*.xlsx
     */
    const checkedReportDirectory =
      getCheckedReportHeaderDir(normalizedReportCode);

    const latestCheckedReportPath = getLatestFile(
      checkedReportDirectory,
    );

    console.log(
      "Latest Checked AF1 Report :",
      latestCheckedReportPath,
    );

    /**
     * สร้าง Output Folder ตาม Report Code
     *
     * ตัวอย่าง:
     * Test_result/Reconcile-report/DS_LTX
     */
    const reconcileOutputDirectory =
      getReconcileOutputDir(normalizedReportCode);

    ensureDirectoryExists(reconcileOutputDirectory);

    /**
     * สร้างชื่อไฟล์ผลลัพธ์
     *
     * ตัวอย่าง:
     * DS_LTX_Reconcile_YYYYMMDD_HHmmss.xlsx
     */
    const reconcileFilePath = buildTimestampedFilePath(
      reconcileOutputDirectory,
      `${normalizedReportCode}_Reconcile`,
      ".xlsx",
    );

    /**
     * Copy Checked Report ไปเป็นไฟล์ผลลัพธ์ชั่วคราว
     *
     * ไฟล์ Reconcile ของแต่ละ Report จะอ่าน Source Worksheet
     * และลบ Source Worksheet ก่อนบันทึกผลลัพธ์
     */
    fs.copyFileSync(latestCheckedReportPath, reconcileFilePath);

    console.log("Reconcile Result File :", reconcileFilePath);

    return reconcileFilePath;
  }

  /**
   * เตรียม Workbook สำหรับ Reconcile
   */
  async prepare(
    reportCode: string,
    reportHeaderRowNumber: number,
  ): Promise<PreparedReconcileWorkbook> {
    const normalizedReportCode = normalizeReportCode(reportCode);

    /**
     * Copy Checked Report ล่าสุด
     * ไปยัง Reconcile Output Folder
     */
    const reconcileFilePath =
      this.copyLatestCheckedReport(normalizedReportCode);

    /**
     * เปิด Workbook ที่ Copy มา
     */
    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(reconcileFilePath);

    /**
     * Worksheet แรกของ Checked Report
     * ใช้เป็น Source สำหรับเปรียบเทียบข้อมูล
     */
    const reportWorksheet = workbook.getWorksheet(1);

    if (!reportWorksheet) {
      throw new Error(
        `AF1 Report worksheet not found in: ` + reconcileFilePath,
      );
    }

    /**
     * อ่าน Header จาก Checked AF1 Report
     */
    const reportHeaders = getHeadersFromRow(
      reportWorksheet,
      reportHeaderRowNumber,
    );

    /**
     * สร้าง Worksheet ผลลัพธ์
     *
     * ตัวอย่าง:
     * DS_LTX_Reconcile
     */
    const resultSheet = this.sheetWriter.createSheet(
      workbook,
      normalizedReportCode,
    );

    return {
      workbook,
      reportWorksheet,
      resultSheet,
      reportHeaders,
      reconcileFilePath,
    };
  }
}
