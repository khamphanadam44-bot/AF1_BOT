/**
 * report-runtime.config.ts
 * ------------------------------------------------------------------
 * Config ถาวรของแต่ละ Report ที่ระบบใช้ร่วมกันตั้งแต่ Script 1-4
 *
 * ผู้ใช้งานไม่ต้องแก้ไฟล์นี้ทุกครั้งที่ Run
 * ให้เปลี่ยน Report ที่ RUN_SETTING.reportCode ใน setting/uat/setting.ts เท่านั้น
 * ------------------------------------------------------------------
 */

import * as path from "path";
import type { ReportCode } from "./report-config";
import { TEST_DATA_INPUT_PATH } from "./paths.config";

export interface ReportRuntimeConfig {
  readonly reportCode: ReportCode;

  /** Test Data ต้นฉบับของ Report */
  readonly testDataInputPath: string;

  /** ชื่อไฟล์ผล Validate Test Data ก่อนต่อ Timestamp */
  readonly testDataResultBasename: string;

  /** ข้อมูลที่ใช้โดย Script 4 */
  readonly testDataSheetName: string;
  readonly testDataHeaderRowNumber: number;
  readonly summaryTemplateFilePath: string;
  readonly summaryTemplateSheetName: string;
  readonly verifiedBy: string;
}

const PROJECT_ROOT = process.cwd();

const buildDefaultConfig = (
  reportCode: ReportCode,
): ReportRuntimeConfig => {
  const displayCode = reportCode.replace(/_/g, "-");

  return {
    reportCode,
    testDataInputPath: path.resolve(
      PROJECT_ROOT,
      TEST_DATA_INPUT_PATH,
    ),
    testDataResultBasename: `${reportCode}_TestData_Validation_Result`,
    testDataSheetName: "Test Data",
    testDataHeaderRowNumber: 5,
    summaryTemplateFilePath: path.resolve(
      PROJECT_ROOT,
      "template",
      `${reportCode}_Automation_Summary_Template.xlsx`,
    ),
    summaryTemplateSheetName: `${displayCode} Summary`,
    verifiedBy: "QAD Automation",
  };
};

/**
 * Config ครั้งแรกของแต่ละ Report
 *
 * ทุก Report ใช้ Test Data กลางไฟล์เดียวกัน
 */
const REPORT_RUNTIME_CONFIG: Record<
  ReportCode,
  ReportRuntimeConfig
> = {
  DS_LTX: {
    ...buildDefaultConfig("DS_LTX"),
    summaryTemplateSheetName:
      "DS_LTX_Summary Result",
  },
  DS_FTU: {
    ...buildDefaultConfig("DS_FTU"),
    summaryTemplateSheetName:
      "DS_FTU_Summary Result",
  },
  DS_PTX: buildDefaultConfig("DS_PTX"),
  DS_FTX: buildDefaultConfig("DS_FTX"),
  
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
