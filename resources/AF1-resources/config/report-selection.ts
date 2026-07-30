/**
 * ======================================================
 * ไฟล์: report-selection.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * อ่านรายชื่อ Report ที่ผู้ใช้ส่งมาจาก Terminal เช่น report=DS_PTX,DS_FTX
 * จากนั้นตรวจชื่อ Report ตัดชื่อซ้ำ และใช้ Report เริ่มต้นจาก setting.ts เมื่อผู้ใช้ไม่ได้ระบุค่า
 *
 * ======================================================
 */
import type {
  ReportName,
} from "./report-config";

import {
  dmsReportName,
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
     * ถ้ามีค่า report จาก Terminal:
     *
     * 1. แยกชื่อ Report ด้วยเครื่องหมาย comma
     * 2. ตัดช่องว่างด้านหน้าและด้านหลัง
     * 3. แปลงเป็นตัวพิมพ์ใหญ่
     * 4. ตัดค่าที่เป็นข้อความว่างออก
     *
     * ถ้าไม่มีค่า report:
     * ใช้ Default Report จาก setting.ts
     */
    const reportNames =
      reportsArgument
        ? reportsArgument
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
          )
        : [
          dmsReportName.reportname,
        ];

    if (
      reportNames.length === 0
    ) {
      throw new Error(
        "ไม่พบ Report ที่ต้องการรัน",
      );
    }

    /**
     * ตรวจสอบว่าชื่อ Report
     * อยู่ใน dmsReportNames หรือไม่
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
     * เก็บเฉพาะชื่อ Report ที่รองรับ
     */
    const selectedReports =
      reportNames.filter(
        isSupportedReport,
      );


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
    /**
     * ลบชื่อ Report ที่ซ้ำกัน
     *
     * ตัวอย่าง:
     *
     * report=DS_PTX,DS_PTX,DS_FTX
     *
     * จะได้:
     *
     * DS_PTX, DS_FTX
     */
    return [
      ...new Set(
        selectedReports,
      ),
    ];
  };