/**
 * ============================================================================
 * script3-compare-report.spec.ts
 * --------------------------------------------------------------------------
 * Script 3 - Compare Report with Test Data
 *
 * Report ที่รองรับ:
 * - DS_LTX
 * - DS_PTX
 * - DS_FTX
 * - DS_FTU
 * - DF_FXU
<<<<<<< Updated upstream
=======
 * - DF_OLB
 * - DF_FXM
>>>>>>> Stashed changes
 *
 * ตัวอย่างคำสั่ง:
 * npm run test:script3 -- report=DS_PTX
 * npm run test:script3 -- report=DS_FTX
 * npm run test:script3 -- report=DS_LTX
<<<<<<< Updated upstream
 * npm run test:script3 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU
=======
 * npm run test:script3 -- report=DS_FTU
 * npm run test:script3 -- report=DF_FXU
 * npm run test:script3 -- report=DF_OLB
 * npm run test:script3 -- report=DF_FXM
 * npm run test:script3 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU,DF_OLB,DF_FXM
>>>>>>> Stashed changes
 * ============================================================================
 */

import "dotenv/config";

import {
  getSelectedReports,
} from "../resources/AF1-resources/config/report-selection";

import {
  getTestDataPath,
} from "../resources/AF1-resources/setting/uat/setting";

/**
 * ============================================================
 * DS_PTX
 * ============================================================
 *
 * DS_PTX ใช้ ptx-reconcile.ts
 * เป็น Entry Point หลักของ Script 3
 */

import {
  reconcilePtxReport,
} from "../resources/AF1-resources/utils/reconcile/DS_PTX/ptx-reconcile";
/**
 * ============================================================
 * DS_FTX
 * ============================================================
 *
 * DS_FTX ใช้ ftx-reconcile.ts เป็นตัวควบคุมหลัก
 * โดยรวม Flow การจับคู่ การตรวจสอบ
 * และการเปรียบเทียบข้อมูลไว้ภายใน Reconciler
 */
import {
  reconcileDsFtx,
} from "../resources/AF1-resources/utils/reconcile/DS_FTX/ftx-reconcile";

import {
  prepareFtxCompareFilePaths,
  printFtxCompareFilePaths,
} from "../resources/AF1-resources/utils/reconcile/DS_FTX/ftx-file-helper";

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

/**
 * ============================================================
 * DF_FXU
 * ============================================================
 *
 * DF_FXU ใช้ fxu-reconcile.ts
 * เป็นตัวควบคุมหลักของ Script 3
 *
 * Flow:
 * - ตรวจ Presence Rule
 * - Exact Matching
 * - Fallback Matching
 * - ตรวจ Core Fields
 * - สร้าง DF_FXU Reconcile Result
 */
import {
  reconcileFxuReport,
} from "../resources/AF1-resources/utils/reconcile/DF_FXU/fxu-reconcile";

/**
 * DF_FXM ใช้ Reconciler แยกจาก DF_FXU
 *
 * รองรับรายการ FX ที่มี
 * USD Equivalent Amount ตั้งแต่
 * 1,000,000 USD ขึ้นไป
 */
import {
  reconcileFXMReport,
} from "../resources/AF1-resources/utils/reconcile/DF_FXM/fxm-reconcile";

const SCRIPT_TIMEOUT =
  300000;


/**
 * ทำงานสำหรับ DS_PTX
 *
 * รายละเอียดการหา Checked Report
 * และการสร้าง Output Path
 * จะถูกจัดการใน ptx-reconcile.ts
 */
const runDsPtxCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * ค้นหา Test Data ของ Report
   * จาก AF1 Share Path
   */
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

  await reconcilePtxReport(
    reportName,
    testDataFilePath,
  );
};

/**
 * ทำงานสำหรับ DS_FTX
 */
const runDsFtxCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * ค้นหา Test Data ของ Report
   * จาก AF1 Share Path
   */
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

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
      testDataFilePath,
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
   * ค่าที่ส่งเข้าไปในฟังก์ชัน:
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
const runDsLtxCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * ค้นหา Test Data ของ Report
   * จาก AF1 Share Path
   */
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

  await reconcileDsLtx(
    reportName,
    testDataFilePath,
  );
};

/**
 * ทำงานสำหรับ DS_FTU
 */
const runDsFtuCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * ค้นหา Test Data ของ Report
   * จาก AF1 Share Path
   */
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

  await reconcileFtuReport(
    testDataFilePath,
  );
};

/**
 * ทำงานสำหรับ DF_FXU
 */
const runDfFxuCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * ค้นหา Test Data ของ DF_FXU
   * จาก AF1 Share Path
   *
   * ตัวอย่าง Folder:
   * AF1_SHAREPATH/DF_FXU
   */
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

  /**
   * เรียก DF_FXU Reconcile Service
   *
   * ภายใน Service จะ:
   * 1. หา Checked DF_FXU Report ล่าสุด
   * 2. อ่าน Test Data
   * 3. ทำ Exact/Fallback Matching
   * 4. ตรวจ Core Fields
   * 5. เขียน Reconcile Result
   */
  await reconcileFxuReport(
<<<<<<< Updated upstream
=======
    testDataFilePath,
  );
};

/**
* ทำงานสำหรับ DF_OLB
*/
const runDfOlbCompare = async (
  reportName: string,
): Promise<void> => {
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

  await reconcileOlbReport(
>>>>>>> Stashed changes
    testDataFilePath,
  );
};

/**
 * ทำงานสำหรับ DF_FXM
 *
 * DF_FXM ใช้ Test Data จาก:
 * AF1_SHAREPATH/af1_test_data/DF_FXM
 *
 * Business Rule หลัก:
 * - FX Conversion
 * - USD Equivalent Amount ตั้งแต่
 *   1,000,000 USD ขึ้นไป
 */
const runDfFxmCompare = async (
  reportName: string,
): Promise<void> => {
  /**
   * ค้นหา Test Data ของ DF_FXM
   * จาก AF1 Share Path
   */
  const testDataFilePath =
    getTestDataPath(
      reportName,
    );

  /**
   * เรียก DF_FXM Reconcile Service
   */
  await reconcileFXMReport(
    testDataFilePath,
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
    await runDsFtxCompare(
      reportName,
    );

    return;
  }

  if (
    reportName ===
    "DS_LTX"
  ) {
    await runDsLtxCompare(
      reportName,
    );

    return;
  }

  if (
    reportName ===
    "DS_FTU"
  ) {
    await runDsFtuCompare(
      reportName,
    );

    return;
  }

  /**
   * DF_FXU:
   *
   * FX Trading Transaction
   * Under 1,000,000 USD Summary
   */
  if (
    reportName ===
    "DF_FXU"
  ) {
    await runDfFxuCompare(
      reportName,
    );

    return;
  }
<<<<<<< Updated upstream
=======

  /**
   * DF_FXU
  */

  if (
    reportName ===
    "DF_OLB"
  ) {
    await runDfOlbCompare(
      reportName,
    );

    return;
  }

  /**
 * DF_FXM:
 *
 * FX Trading Transaction
 * ตั้งแต่ 1,000,000 USD ขึ้นไป
 */
  if (
    reportName ===
    "DF_FXM"
  ) {
    await runDfFxmCompare(
      reportName,
    );

    return;
  }
>>>>>>> Stashed changes

  throw new Error(


    [
      `Script 3 ยังไม่รองรับ Report: ${reportName}`,
<<<<<<< Updated upstream
            "Report ที่รองรับ: DS_LTX, DS_PTX, DS_FTX, DS_FTU, DF_FXU",
=======
      "Report ที่รองรับ: DS_LTX, DS_PTX, DS_FTX, DS_FTU, DF_FXU, DF_OLB, DF_FXM",

>>>>>>> Stashed changes
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
