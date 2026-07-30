/**
 * ============================================================================
 * ไฟล์: compare-file-helper.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้ดูแลเรื่อง “ตำแหน่งไฟล์” สำหรับ Script 3
 *
 * งานหลักมี 2 ส่วน
 * 1. ค้นหา Report ล่าสุดจากโฟลเดอร์ที่ผ่านการตรวจ Header แล้ว
 * 2. สร้างชื่อและตำแหน่งสำหรับบันทึกไฟล์ผล Compare
 *
 * วิธีเลือก Report ล่าสุด
 * -----------------------
 * โปรแกรมจะอ่านไฟล์ .xlsx ทั้งหมดในโฟลเดอร์ของ Report แล้วเรียงตาม
 * เวลาที่ไฟล์ถูกแก้ไขล่าสุด จากนั้นเลือกไฟล์ใหม่ที่สุดมาใช้งาน
 *
 * วิธีตั้งชื่อ Output
 * -------------------
 * ชื่อไฟล์ผลลัพธ์จะมี Report Name และ Timestamp เพื่อป้องกันไฟล์ใหม่
 * เขียนทับไฟล์เดิม เช่น
 * DS_PTX_Compare_Result_20260715_105317.xlsx
 *
 * หากไม่พบโฟลเดอร์หรือไม่พบไฟล์ Report โปรแกรมจะแจ้ง Error ทันที
 * เพื่อให้ผู้ใช้งานทราบว่าต้องตรวจสอบขั้นตอนก่อนหน้า
 * ============================================================================
 */


/**
 * ส่วน import ด้านล่าง คือการนำเครื่องมือหรือโครงสร้างข้อมูล
 * จากไฟล์อื่นมาใช้ในไฟล์นี้ เปรียบเหมือนการหยิบอุปกรณ์ที่เตรียมไว้แล้ว
 * มาใช้งาน โดยไม่ต้องเขียนทุกอย่างซ้ำใหม่
 */

import fs from "fs";
import path from "path";

import {
  getCheckedReportHeaderDir,
} from "../../../config/paths.config";

/**
 * Folder เก็บ Compare Result
 */
const RECONCILE_REPORT_ROOT_FOLDER = path.resolve(
  process.cwd(),
  "Test_result",
  "Reconcile-report",
);

/**
 * สร้าง Timestamp
 *
 * รูปแบบ:
 * yyyyMMdd_HHmmss
 */
const getTimestamp = (): string => {

  const now =
    new Date();

  const yyyy =
    now.getFullYear();

  const MM =
    String(
      now.getMonth() + 1,
    ).padStart(
      2,
      "0",
    );

  const dd =
    String(
      now.getDate(),
    ).padStart(
      2,
      "0",
    );

  const HH =
    String(
      now.getHours(),
    ).padStart(
      2,
      "0",
    );

  const mm =
    String(
      now.getMinutes(),
    ).padStart(
      2,
      "0",
    );

  const ss =
    String(
      now.getSeconds(),
    ).padStart(
      2,
      "0",
    );

  return (
    `${yyyy}${MM}${dd}_${HH}${mm}${ss}`
  );

};

/**
 * ============================================================================
 * หาไฟล์ Report ล่าสุด
 * ============================================================================
 */
export const getLatestReportPath = (
  reportCode: string,
): string => {

  /**
   * Folder ที่เก็บ Report
   * หลังผ่านการตรวจ Header
   */
  const reportFolder =
    getCheckedReportHeaderDir(
      reportCode,
    );

  if (
    !fs.existsSync(
      reportFolder,
    )
  ) {

    throw new Error(
      `Report folder not found : ${reportFolder}`,
    );

  }

  const files =
    fs
      .readdirSync(
        reportFolder,
      )
      .filter(
        file =>
          file
            .toLowerCase()
            .endsWith(
              ".xlsx",
            ),
      );

  if (
    files.length === 0
  ) {

    throw new Error(
      `No report found in ${reportFolder}`,
    );

  }

  /**
   * เรียงจากไฟล์ใหม่ → เก่า
   */
  files.sort(
    (
      firstFile,
      secondFile,
    ) => {

      const firstModifiedTime =
        fs
          .statSync(
            path.join(
              reportFolder,
              firstFile,
            ),
          )
          .mtime
          .getTime();

      const secondModifiedTime =
        fs
          .statSync(
            path.join(
              reportFolder,
              secondFile,
            ),
          )
          .mtime
          .getTime();

      return (
        secondModifiedTime -
        firstModifiedTime
      );

    },
  );

  return path.join(
    reportFolder,
    files[0],
  );

};


/**
 * ============================================================================
 * สร้าง Output File Path
 * ============================================================================
 */
export const getCompareResultOutputPath = (
  reportName: string,
): string => {
  /**
   * ปรับชื่อ Report สำหรับใช้กับชื่อไฟล์
   *
   * ตัวอย่าง:
   * ds_ptx  -> DS_PTX
   * DS-PTX  -> DS_PTX
   */
  const normalizedReportName = String(reportName)
    .trim()
    .toUpperCase()
    .replace(/-/g, "_");

  /**
   * ใช้ชื่อ Report รูปแบบมาตรฐานเป็นชื่อ Folder
   *
   * ตัวอย่าง:
   * DS_PTX -> DS_PTX
   */
  const reportFolderName =
    normalizedReportName;

  /**
   * Path ของ Folder ปลายทาง
   *
   * ตัวอย่าง:
   * Test_result/Reconcile-report/DS_PTX
   */
  const reconcileReportFolder = path.join(
    RECONCILE_REPORT_ROOT_FOLDER,
    reportFolderName,
  );

  /**
   * สร้าง Folder ทุกชั้นให้อัตโนมัติ
   *
   * หากยังไม่มี Reconcile-report หรือ DS_PTX
   * โปรแกรมจะสร้างให้เอง
   */
  fs.mkdirSync(
    reconcileReportFolder,
    {
      recursive: true,
    },
  );

  const timestamp = getTimestamp();

  const fileName =
    `${normalizedReportName}_Compare_Result_${timestamp}.xlsx`;

  return path.join(
    reconcileReportFolder,
    fileName,
  );
};
