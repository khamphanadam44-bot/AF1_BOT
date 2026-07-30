/**
 * ============================================================================
 * script3-compare-report.spec.ts
 * ----------------------------------------------------------------------------
 * Script 3 - Compare Report with Test Data
 *
 * รองรับ:
 * - DS_LTX
 * - DS_PTX
 * - DS_FTX
 * - DS_FTU
 *
 * ตัวอย่าง:
 * npm run test:script3 -- report=DS_PTX
 * npm run test:script3 -- report=DS_FTX
 * npm run test:script3 -- report=DS_LTX
 * npm run test:script3 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU
 * ============================================================================
 */

import "dotenv/config";

import {
  getSelectedReports,
} from "../resources/AF1-resources/config/report-selection";

import {
  testDataPath,
} from "../resources/AF1-resources/setting/uat/setting";

/**
 * ============================================================
 * DS_PTX
 * ============================================================
 */

import {
  compareReportWithTestData as compareDsPtx,
} from "../resources/AF1-resources/utils/reconcile/DS_PTX/compare-validator";

import {
  getLatestReportPath as getLatestDsPtxReportPath,
} from "../resources/AF1-resources/utils/reconcile/DS_PTX/compare-file-helper";

/**
 * ============================================================
 * DS_FTX
 * ============================================================
 *
 * DS_FTX ใช้ Reconciler เป็นตัวควบคุมหลัก
 * ไม่ได้เรียก compare-validator.ts โดยตรง
 */

import {
  reconcileDsFtx,
} from "../resources/AF1-resources/utils/reconcile/DS_FTX/ds-ftx-reconciler";

import {
  prepareFtxCompareFilePaths,
  printFtxCompareFilePaths,
} from "../resources/AF1-resources/utils/reconcile/DS_FTX/compare-file-helper";

/**
 * ============================================================
 * DS_LTX
 * ============================================================
 */

import {
  reconcileReport as reconcileDsLtx,
} from "../resources/AF1-resources/utils/reconcile/DS_LTX/ltx-reconcile";

import {
  reconcileFtuReport,
} from "../resources/AF1-resources/utils/reconcile/DS_FTU/ftu-reconcile";

const SCRIPT_TIMEOUT =
  300000;

/**
 * ทำงานสำหรับ DS_PTX
 */
const runDsPtxCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * หา Checked Report DS_PTX ล่าสุด
   */
  const reportPath =
    getLatestDsPtxReportPath(
      reportName,
    );

  console.log(
    "REPORT PATH",
  );

  console.log(
    reportPath,
  );

  console.log(
    "================================",
  );

  console.log(
    "TEST DATA PATH",
  );

  console.log(
    testDataPath,
  );

  console.log(
    "================================",
  );

  /**
   * เรียก Logic เดิมของ DS_PTX
   */
  await compareDsPtx(
    reportName,
    reportPath,
    testDataPath,
  );
};

/**
 * ทำงานสำหรับ DS_FTX
 */
const runDsFtxCompare = async (): Promise<void> => {
  /**
   * prepareFtxCompareFilePaths() จะ:
   *
   * 1. หา Checked Report DS_FTX ล่าสุด
   * 2. ตรวจสอบ Test Data
   * 3. สร้าง Folder ผลลัพธ์
   * 4. สร้างชื่อ Output File
   */
  const filePaths =
    prepareFtxCompareFilePaths(
      process.cwd(),
      testDataPath,
    );

  /**
   * แสดง Path ที่เลือกใช้
   */
  printFtxCompareFilePaths(
    filePaths,
  );

  /**
   * เรียกตัวควบคุมหลักของ DS_FTX
   *
   * ค่าที่ส่งเข้าไปในฟังก์ชัน (Parameter):
   * 1. Report File
   * 2. Test Data File
   * 3. Output File
   */
  await reconcileDsFtx(
    filePaths.reportFilePath,
    filePaths.testDataFilePath,
    filePaths.outputFilePath,
  );
};

/**
 * ทำงานสำหรับ DS_LTX
 */
const runDsLtxCompare = async (): Promise<void> => {
  await reconcileDsLtx(
    "DS_LTX",
    testDataPath,
  );
};

/**
 * ทำงานสำหรับ DS_FTU
 */
const runDsFtuCompare = async (): Promise<void> => {
  await reconcileFtuReport(
    testDataPath,
  );
};

/**
 * เลือก Logic ตาม Report
 */
const runCompareByReport = async (
  reportName: string,
): Promise<void> => {
  if (
    reportName ===
    "DS_PTX"
  ) {
    await runDsPtxCompare(
      reportName,
    );

    return;
  }

  if (
    reportName ===
    "DS_FTX"
  ) {
    await runDsFtxCompare();

    return;
  }

  if (
    reportName ===
    "DS_LTX"
  ) {
    await runDsLtxCompare();

    return;
  }

  if (
    reportName ===
    "DS_FTU"
  ) {
    await runDsFtuCompare();

    return;
  }

  throw new Error(
    [
      `Script 3 ยังไม่รองรับ Report: ${reportName}`,
      "Report ที่รองรับ: DS_LTX, DS_PTX, DS_FTX, DS_FTU",
    ].join(
      "\n",
    ),
  );
};

describe(
  "Script 3 - Compare Report with Test Data",
  function () {
    this.timeout(
      SCRIPT_TIMEOUT,
    );

    /**
     * อ่าน Report ที่เลือกจาก Terminal
     */
    const selectedReports =
      getSelectedReports();

    console.log(
      "================================",
    );

    console.log(
      "SELECTED REPORTS",
    );

    console.log(
      selectedReports.join(
        ", ",
      ),
    );

    console.log(
      "================================",
    );

    /**
     * สร้าง Test Case ตามจำนวน Report ที่เลือก
     */
    for (
      const selectedReport of
      selectedReports
    ) {
      it(
        `Compare ${selectedReport} Report`,
        async function () {
          console.log("");

          console.log(
            "================================",
          );

          console.log(
            `START COMPARE: ${selectedReport}`,
          );

          console.log(
            "================================",
          );

          /**
           * เลือก Flow ของ Report
           */
          await runCompareByReport(
            selectedReport,
          );

          console.log(
            "================================",
          );

          console.log(
            `COMPARE COMPLETE: ${selectedReport}`,
          );

          console.log(
            "================================",
          );
        },
      );
    }
  },
);
