/**
 * ======================================================
 * ไฟล์: summary-file-helper.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * จัดการ Path และชื่อไฟล์ทั้งหมดที่ Script 4 ต้องใช้
 * ใช้ค้นหาไฟล์ Excel ล่าสุดของแต่ละแหล่งข้อมูล สร้างชื่อไฟล์ Summary และป้องกันการเลือกไฟล์ชั่วคราวของ Excel
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import fs from "fs";
import path from "path";

/**
 * Folder หลักสำหรับเก็บ Compare Result จาก Script 3
 *
 * ตัวอย่าง:
 * Test_result/Reconcile-report/DS_PTX
 */
const RECONCILE_REPORT_ROOT_FOLDER = path.resolve(
  process.cwd(),
  "Test_result",
  "Reconcile-report",
);

/**
 * Folder หลักสำหรับเก็บ Checked Report จาก Script 2
 *
 * ตัวอย่าง:
 * Test_result/Checked-report-header/DS_PTX
 */
const CHECKED_REPORT_HEADER_ROOT_FOLDER = path.resolve(
  process.cwd(),
  "Test_result",
  "Checked-report-header",
);

/**
 * Folder สำหรับเก็บ Checked Test Data จาก Script 2
 *
 * ตัวอย่าง:
 * Test_result/Checked-testdata-header
 */
const CHECKED_TEST_DATA_FOLDER = path.resolve(
  process.cwd(),
  "Test_result",
  "Checked-testdata-header",
);

/**
 * Folder หลักสำหรับเก็บผลลัพธ์ Summary จาก Script 4
 *
 * ตัวอย่าง:
 * Test_result/Summary-report/DS_PTX
 */
const SUMMARY_REPORT_ROOT_FOLDER = path.resolve(
  process.cwd(),
  "Test_result",
  "Summary-report",
);

/**
 * Original Test Data กลางของทุก Report
 *
 * Script 4 ใช้ไฟล์นี้สร้างข้อมูลฝั่ง "Test Script Data"
 * ในชีท Summary หลักของทั้ง DS_PTX และ DS_FTX
 */
export const ORIGINAL_TEST_DATA_PATH = path.resolve(
  process.cwd(),
  "test_data",
  "Test_Data_Downstream-for pilot.xlsx",
);

/**
 * แปลงชื่อ Report ให้อยู่ในรูปแบบมาตรฐาน
 *
 * ตัวอย่าง:
 * ds_ptx  -> DS_PTX
 * DS-PTX  -> DS_PTX
 */
const normalizeReportName = (reportName: string): string => {
  return String(reportName).trim().toUpperCase().replace(/-/g, "_");
};

/**
 * เลือก Template ให้ตรงกับ Report
 *
 * DS_PTX -> template/DS_PTX_Automation_Summary.xlsx
 * DS_FTX -> template/DS_FTX_Automation_Summary.xlsx
 */
export const getSummaryTemplatePath = (reportName: string): string => {
  const normalizedReportName = normalizeReportName(reportName);

  const templatePath = path.resolve(
    process.cwd(),
    "template",
    `${normalizedReportName}_Automation_Summary.xlsx`,
  );

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Summary template not found: ${templatePath}`);
  }

  console.log("Summary Template      :", templatePath);

  return templatePath;
};

/**
 * สร้าง Timestamp
 *
 * รูปแบบ:
 * yyyyMMdd_HHmmss
 *
 * ตัวอย่าง:
 * 20260721_153045
 */
const getTimestamp = (): string => {
  const now = new Date();

  const yyyy = now.getFullYear();

  const MM = String(now.getMonth() + 1).padStart(2, "0");

  const dd = String(now.getDate()).padStart(2, "0");

  const HH = String(now.getHours()).padStart(2, "0");

  const mm = String(now.getMinutes()).padStart(2, "0");

  const ss = String(now.getSeconds()).padStart(2, "0");

  return `${yyyy}${MM}${dd}` + `_${HH}${mm}${ss}`;
};

/**
 * อ่าน Timestamp จากชื่อไฟล์
 *
 * รูปแบบที่รองรับ:
 * yyyyMMdd_HHmmss
 *
 * ตัวอย่าง:
 * DS_FTX_Compare_Result_20260724_115044.xlsx
 * จะได้ค่า:
 * 20260724115044
 *
 * ถ้าชื่อไฟล์ไม่มี Timestamp รูปแบบนี้
 * จะคืนค่า 0 แล้วใช้ Modified Time เป็นตัวช่วยแทน
 */
const getTimestampFromFileName = (fileName: string): number => {
  const matchedTimestamp = fileName.match(/(\d{8})_(\d{6})/g);

  if (!matchedTimestamp || matchedTimestamp.length === 0) {
    return 0;
  }

  /**
   * ใช้ Timestamp ชุดสุดท้ายในชื่อไฟล์
   * เพื่อป้องกันกรณีชื่อไฟล์มีวันที่มากกว่าหนึ่งชุด
   */
  const latestTimestampText = matchedTimestamp[
    matchedTimestamp.length - 1
  ].replace("_", "");

  return Number(latestTimestampText);
};

/**
 * หาไฟล์ Excel ล่าสุดจาก Folder
 *
 * @param folderPath Folder ที่ต้องการค้นหา
 * @param fileFilter เงื่อนไขการเลือกชื่อไฟล์
 */
const getLatestExcelFile = (
  folderPath: string,
  fileFilter: (fileName: string) => boolean,
): string => {
  /**
   * ตรวจสอบว่า Folder มีอยู่จริง
   */
  if (!fs.existsSync(folderPath)) {
    throw new Error(`Folder not found: ${folderPath}`);
  }

  /**
   * อ่านไฟล์ทั้งหมดภายใน Folder
   * แล้วเลือกเฉพาะไฟล์ที่ตรงตามเงื่อนไข
   */
  const matchedFiles = fs
    .readdirSync(folderPath)
    .filter((fileName) => {
      const fullPath = path.join(folderPath, fileName);

      /**
       * ข้าม Folder ย่อย
       */
      if (!fs.statSync(fullPath).isFile()) {
        return false;
      }

      return fileFilter(fileName);
    })
    .map((fileName) => {
      const fullPath = path.join(folderPath, fileName);

      return {
        fullPath,
        fileName,
        fileTimestamp: getTimestampFromFileName(fileName),
        modifiedTime: fs.statSync(fullPath).mtimeMs,
      };
    })
    .sort((firstFile, secondFile) => {
      /**
       * เลือกจาก Timestamp ในชื่อไฟล์ก่อน
       *
       * เหตุผล:
       * เวลา Copy Project หรือแตก ZIP
       * Modified Time ของหลายไฟล์อาจกลายเป็นเวลาเดียวกัน
       */
      if (firstFile.fileTimestamp !== secondFile.fileTimestamp) {
        return secondFile.fileTimestamp - firstFile.fileTimestamp;
      }

      /**
       * ถ้าชื่อไฟล์ไม่มี Timestamp
       * หรือ Timestamp เท่ากัน
       * จึงค่อยเปรียบเทียบ Modified Time
       */
      if (firstFile.modifiedTime !== secondFile.modifiedTime) {
        return secondFile.modifiedTime - firstFile.modifiedTime;
      }

      /**
       * เงื่อนไขสุดท้าย:
       * เรียงชื่อไฟล์จากมากไปน้อย
       */
      return secondFile.fileName.localeCompare(firstFile.fileName);
    });

  /**
   * แจ้ง Error หากไม่พบไฟล์
   */
  if (matchedFiles.length === 0) {
    throw new Error(`No matching Excel file found in: ${folderPath}`);
  }

  /**
   * รายการแรกคือไฟล์ที่แก้ไขล่าสุด
   */
  return matchedFiles[0].fullPath;
};

/**
 * หา Compare Result ล่าสุดจาก Script 3
 *
 * อ่านจาก:
 * Test_result/Reconcile-report/DS_PTX
 *
 * ตัวอย่าง:
 * DS_PTX_Compare_Result_20260721_092700.xlsx
 */
export const getLatestCompareResultPath = (reportName: string): string => {
  const normalizedReportName = normalizeReportName(reportName);

  const reconcileReportFolder = path.join(
    RECONCILE_REPORT_ROOT_FOLDER,
    normalizedReportName,
  );

  const latestCompareResult = getLatestExcelFile(
    reconcileReportFolder,
    (fileName) => {
      const upperFileName = fileName.toUpperCase();

      const isExcelFile = upperFileName.endsWith(".XLSX");

      const isCorrectReport = upperFileName.includes(normalizedReportName);

      const isCompareResultFile =
        upperFileName.includes("COMPARE") && upperFileName.includes("RESULT");

      const isReconcileFile = upperFileName.includes("RECONCILE");

      const isScript3Result = isCompareResultFile || isReconcileFile;

      return isExcelFile && isCorrectReport && isScript3Result;
    },
  );

  console.log("Latest Compare Result :", latestCompareResult);

  return latestCompareResult;
};

/**
 * หา Checked Report ล่าสุดจาก Script 2
 *
 * อ่านจาก:
 * Test_result/Checked-report-header/DS_PTX
 *
 * ใช้สำหรับสร้างชีท:
 * DS_PTX
 */
export const getLatestCheckedReportPath = (reportName: string): string => {
  const normalizedReportName = normalizeReportName(reportName);

  const checkedReportFolder = path.join(
    CHECKED_REPORT_HEADER_ROOT_FOLDER,
    normalizedReportName,
  );

  const latestCheckedReport = getLatestExcelFile(
    checkedReportFolder,
    (fileName) => {
      const upperFileName = fileName.toUpperCase();

      const isExcelFile = upperFileName.endsWith(".XLSX");

      const isCorrectReport = upperFileName.includes(normalizedReportName);

      const isExportFile = upperFileName.startsWith("EXPORT_");

      return isExcelFile && isCorrectReport && isExportFile;
    },
  );

  console.log("Latest Checked Report :", latestCheckedReport);

  return latestCheckedReport;
};

/**
 * หา Checked Test Data ล่าสุดจาก Script 2
 *
 * อ่านจาก:
 * Test_result/Checked-testdata-header/DS_PTX
 *
 * ตัวอย่าง:
 * DS_PTX_TestData_Validation_Result_20260721_092659.xlsx
 *
 * ใช้สำหรับสร้างชีท:
 * Test Data
 */
export const getLatestCheckedTestDataPath = (reportName: string): string => {
  const normalizedReportName = normalizeReportName(reportName);

  /**
   * ระบุ Folder ย่อยตามชื่อ Report
   *
   * ตัวอย่าง:
   * Test_result/Checked-testdata-header/DS_PTX
   */
  const checkedTestDataFolder = path.join(
    CHECKED_TEST_DATA_FOLDER,
    normalizedReportName,
  );

  const latestCheckedTestData = getLatestExcelFile(
    checkedTestDataFolder,
    (fileName) => {
      const upperFileName = fileName.toUpperCase();

      const isExcelFile = upperFileName.endsWith(".XLSX");

      const isCorrectReport = upperFileName.includes(normalizedReportName);

      const isTestDataFile = upperFileName.includes("TESTDATA");

      const isValidationFile = upperFileName.includes("VALIDATION");

      const isResultFile = upperFileName.includes("RESULT");

      return (
        isExcelFile &&
        isCorrectReport &&
        isTestDataFile &&
        isValidationFile &&
        isResultFile
      );
    },
  );

  console.log("Latest Checked Test Data :", latestCheckedTestData);

  return latestCheckedTestData;
};

/**
 * สร้าง Path สำหรับบันทึกผลลัพธ์ Script 4
 *
 * Output:
 * Test_result/Summary-report/DS_PTX
 *
 * ตัวอย่างชื่อไฟล์:
 * DS_PTX_Automation_Summary_20260721_153045-Final.xlsx
 */
export const getSummaryResultOutputPath = (reportName: string): string => {
  const normalizedReportName = normalizeReportName(reportName);

  const summaryReportFolder = path.join(
    SUMMARY_REPORT_ROOT_FOLDER,
    normalizedReportName,
  );

  /**
   * สร้าง Folder หากยังไม่มี
   */
  fs.mkdirSync(summaryReportFolder, {
    recursive: true,
  });

  const outputFileName =
    `${normalizedReportName}` +
    `_Automation_Summary_` +
    `${getTimestamp()}` +
    `-Final.xlsx`;

  return path.join(summaryReportFolder, outputFileName);
};

/**
 * สร้าง Run ID สำหรับแสดงใน Summary
 *
 * ตัวอย่าง:
 * RUN_20260721_153045
 */
export const createRunId = (): string => {
  return `RUN_${getTimestamp()}`;
};
