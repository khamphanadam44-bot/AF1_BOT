/**
 * report-runtime.config.ts
 * ------------------------------------------------------------------
 * Config กลางของแต่ละ Report
 *
 * ปัจจุบัน Reconcile ใช้ไฟล์นี้อ่านเลขแถว Header ของ Test Data
 * ผู้ใช้เลือก Report ที่ต้องการรันผ่านค่า report ใน Terminal
 * ------------------------------------------------------------------
 */

import type { ReportCode } from "./report-config";

export interface ReportRuntimeConfig {
  readonly reportCode: ReportCode;

  /** หมายเลขแถว Header ของ Test Data */
  readonly testDataHeaderRowNumber: number;
}



const buildDefaultConfig = (
  reportCode: ReportCode,
): ReportRuntimeConfig => {
  return {
    reportCode,
    testDataHeaderRowNumber: 5,
  };
};

/**
 * Config การทำงานของแต่ละ Report
 *
 * ใช้เก็บค่าที่เกี่ยวข้องกับผล Validation
 * และการสร้าง Summary ของแต่ละ Report
 */
const REPORT_RUNTIME_CONFIG: Record<
  ReportCode,
  ReportRuntimeConfig
> = {
  DS_LTX: buildDefaultConfig("DS_LTX"),
  DS_PTX: buildDefaultConfig("DS_PTX"),
  DS_FTX: buildDefaultConfig("DS_FTX"),
  DS_FTU: buildDefaultConfig("DS_FTU"),
  DF_FXU: buildDefaultConfig("DF_FXU"),
  DF_OLB: buildDefaultConfig("DF_OLB"),
  DF_FXM: buildDefaultConfig("DF_FXM"),
};

export const getReportRuntimeConfig = (
  reportCode: ReportCode,
): ReportRuntimeConfig => {
  const config = REPORT_RUNTIME_CONFIG[reportCode];

  if (!config) {
    throw new Error(
      `ไม่พบ Runtime Config สำหรับ Report "${reportCode}"`,
    );
  }

  return config;
};
