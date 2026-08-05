/**
 * ======================================================
 * ไฟล์: report-selection.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * อ่านรายชื่อ Report ที่ผู้ใช้ส่งมาจาก Terminal เช่น report=DS_PTX,DS_FTX
 * จากนั้นตรวจชื่อ Report และตัดชื่อ Report ที่ซ้ำกัน
 * หากผู้ใช้ไม่ระบุชื่อ Report ระบบจะแจ้ง Error และหยุดทำงานทันที
 *
 * ======================================================
 */
import type {
  ReportName,
} from "./report-config";

import {
  dmsReportNames,
} from "../setting/uat/setting";

/**
 * รายชื่อ Report ที่ระบบรองรับ
 *
 * อ่านจาก dmsReportNames ใน setting.ts
 */
const SUPPORTED_REPORTS:
  readonly ReportName[] =
  dmsReportNames;

/**
 * ตรวจว่าชื่อ Report ที่รับมา
 * อยู่ในรายการที่ระบบรองรับหรือไม่
 */
const isSupportedReport = (
  reportName: string,
): reportName is ReportName => {
  return SUPPORTED_REPORTS.some(
    (
      supportedReport,
    ) =>
      supportedReport ===
      reportName,
  );
};

/**
 * อ่านชื่อ Report จาก Terminal
 *
 * รูปแบบที่รองรับ:
 *
 * report=DS_PTX
 * report=DS_PTX,DS_FTX
 * report DS_PTX
 */
const readReportsArgument =
  (): string | undefined => {
    /**
     * ตรวจหารูปแบบ:
     *
     * report=DS_PTX
     * report=DS_PTX,DS_FTX
     */
    const inlineArgument =
      process.argv.find(
        (
          argument,
        ) =>
          argument.startsWith(
            "report=",
          ),
      );

    if (
      inlineArgument
    ) {
      return inlineArgument.slice(
        "report=".length,
      );
    }

    /**
     * ตรวจหารูปแบบ:
     *
     * report DS_PTX
     * report DS_PTX,DS_FTX
     */
    const argumentIndex =
      process.argv.indexOf(
        "report",
      );

    if (
      argumentIndex >= 0
    ) {
      return process.argv[
        argumentIndex + 1
      ];
    }

    return undefined;
  };

/**
 * คืนค่ารายชื่อ Report ที่ต้องการรัน
 *
 * ลำดับการเลือก:
 *
 * 1. ใช้ Report จาก report
 * 2. ถ้าไม่ได้ระบุ ใช้ dmsReportName จาก setting.ts
 */
/**
 * คืนค่ารายชื่อ Report ที่ผู้ใช้ต้องการรัน
 *
 * ผู้ใช้ต้องระบุชื่อ Report ผ่านค่า report ทุกครั้ง
 *
 * ตัวอย่าง:
 * report=DS_PTX
 * report=DS_PTX,DS_FTX
 *
 * หากไม่ระบุชื่อ Report หรือระบุเป็นค่าว่าง
 * ระบบจะแจ้ง Error และหยุดทำงานทันที
 */
export const getSelectedReports =
  (): ReportName[] => {
    const reportsArgument =
      readReportsArgument();

    console.log(
      "PROCESS ARGV:",
      process.argv,
    );

    console.log(
      "REPORT ARGUMENT:",
      reportsArgument,
    );

    /**
     * ตรวจกรณีผู้ใช้ไม่ได้ระบุชื่อ Report
     *
     * ตัวอย่างที่เข้าเงื่อนไข:
     *
     * npm run test:script1
     * npm run test:script1 -- report=
     * npm run test:script1 -- report
     *
     * เมื่อ throw Error แล้ว
     * ระบบจะหยุดก่อนเปิด Browser หรือทำขั้นตอนถัดไป
     */
    if (
      reportsArgument === undefined ||
      reportsArgument.trim() === ""
    ) {
      throw new Error(
        "กรุณากรอกชื่อรายงานที่ท่านต้องการ",
      );
    }

    /**
     * จัดรูปแบบชื่อ Report
     *
     * 1. แยกชื่อด้วยเครื่องหมาย comma
     * 2. ตัดช่องว่างหน้าและหลัง
     * 3. แปลงเป็นตัวพิมพ์ใหญ่
     * 4. ตัดค่าที่เป็นข้อความว่างออก
     */
    const reportNames =
      reportsArgument
        .split(
          ",",
        )
        .map(
          (
            reportName,
          ) =>
            reportName
              .trim()
              .toUpperCase(),
        )
        .filter(
          Boolean,
        );

    /**
     * ตรวจกรณีผู้ใช้ส่งเฉพาะเครื่องหมาย comma
     *
     * ตัวอย่าง:
     * report=,,,
     */
    if (
      reportNames.length === 0
    ) {
      throw new Error(
        "กรุณากรอกชื่อรายงานที่ท่านต้องการ",
      );
    }

    /**
     * ตรวจสอบว่าชื่อ Report
     * อยู่ในรายการที่ระบบรองรับหรือไม่
     */
    const unsupportedReports =
      reportNames.filter(
        (
          reportName,
        ) =>
          !isSupportedReport(
            reportName,
          ),
      );

    if (
      unsupportedReports.length > 0
    ) {
      throw new Error(
        `ไม่รองรับ Report: ${unsupportedReports.join(", ")}\n` +
        `Report ที่รองรับ: ${SUPPORTED_REPORTS.join(", ")}`,
      );
    }

    /**
     * เก็บเฉพาะชื่อ Report ที่ระบบรองรับ
     */
    const selectedReports =
      reportNames.filter(
        isSupportedReport,
      );

    /**
     * ลบชื่อ Report ที่ซ้ำกัน
     *
     * ตัวอย่าง:
     * report=DS_PTX,DS_PTX,DS_FTX
     *
     * ผลลัพธ์:
     * DS_PTX, DS_FTX
     */
    const uniqueReports = [
      ...new Set(
        selectedReports,
      ),
    ];

    console.log(
      "SELECTED REPORTS:",
      uniqueReports,
    );

    return uniqueReports;
  };