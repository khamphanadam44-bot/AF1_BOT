/**
 * ftx-file-helper.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * จัดเตรียมและตรวจสอบ Path ที่ใช้ใน Script 3 ของ DS_FTX
 *
 * ความสามารถหลัก
 * 1. กำหนด Folder ของ Checked Report
 * 2. กำหนด Folder ของ Reconcile Result
 * 3. ค้นหา Report DS_FTX ที่แก้ไขล่าสุด
 * 4. ตรวจสอบไฟล์ Test Data
 * 5. สร้าง Timestamp
 * 6. สร้างชื่อไฟล์ผลลัพธ์
 * 7. ป้องกันชื่อไฟล์ผลลัพธ์ซ้ำ
 * 8. คืน Path ทั้งหมดให้ Script 3
 *
 * หมายเหตุ
 * - ไฟล์นี้ไม่ได้อ่านข้อมูลใน Excel
 * - ไฟล์นี้ไม่ได้ Compare ข้อมูล
 * - ไฟล์นี้ไม่ได้เขียนไฟล์ผลลัพธ์
 * - ไฟล์นี้ไม่ได้สร้างหรือย้ายไฟล์ชั่วคราว
 * ------------------------------------------------------------------
 */

import fs from "fs";
import path from "path";

/**
 * Path แบบ Relative ของ Folder ที่เก็บ Report DS_FTX
 * หลังจากผ่าน Script 2 แล้ว
 *
 * เมื่อนำไปรวมกับ Project Root จะได้:
 *
 * <Project Root>/
 * Test_result/
 * Checked-report-header/
 * DS_FTX/
 */
const CHECKED_REPORT_FOLDER =
  path.join(
    "Test_result",
    "Checked-report-header",
    "DS_FTX",
  );

/**
 * Path แบบ Relative ของ Folder ผลลัพธ์ Script 3
 *
 * เมื่อนำไปรวมกับ Project Root จะได้:
 *
 * <Project Root>/
 * Test_result/
 * Reconcile-report/
 * DS_FTX/
 */
const RECONCILE_RESULT_FOLDER =
  path.join(
    "Test_result",
    "Reconcile-report",
    "DS_FTX",
  );

/**
 * Prefix หรือข้อความส่วนต้นของชื่อไฟล์ผลลัพธ์
 *
 * ตัวอย่าง:
 * DS_FTX_Compare_Result_20260730_180000.xlsx
 */
const COMPARE_RESULT_FILE_PREFIX =
  "DS_FTX_Compare_Result";

/**
 * นามสกุลไฟล์ Excel ที่ระบบรองรับ
 *
 * Code นี้รองรับเฉพาะ .xlsx
 * และไม่รองรับไฟล์ .xls
 */
const EXCEL_FILE_EXTENSION =
  ".xlsx";

/**
 * จำนวนลำดับสูงสุดที่ใช้ค้นหาชื่อไฟล์ใหม่
 * เมื่อพบชื่อไฟล์ผลลัพธ์ซ้ำ
 *
 * ระบบจะลองตั้งแต่
 * - _01
 * - _02
 * - ...
 * - _9999
 */
const MAX_DUPLICATE_FILE_NUMBER =
  9999;

/**
 * รูปแบบข้อมูล Path ทั้งหมดที่เตรียมให้ Script 3
 */
export interface FtxCompareFilePaths {
  /**
   * Absolute Path ของ Root Folder โปรเจกต์
   */
  projectRootPath: string;

  /**
   * Absolute Path ของ Folder Checked Report
   */
  checkedReportDirectoryPath: string;

  /**
   * Absolute Path ของ Report DS_FTX ล่าสุดที่เลือก
   */
  reportFilePath: string;

  /**
   * Absolute Path ของ Test Data ที่ใช้เปรียบเทียบ
   */
  testDataFilePath: string;

  /**
   * Absolute Path ของ Folder ผลลัพธ์
   */
  outputDirectoryPath: string;

  /**
   * Absolute Path ของไฟล์ผลลัพธ์ที่ต้องสร้าง
   *
   * ไฟล์ยังไม่ได้ถูกสร้างในขั้นตอนนี้
   */
  outputFilePath: string;

  /**
   * Timestamp ที่ใช้ในชื่อไฟล์ผลลัพธ์
   */
  timestamp: string;
}

/**
 * รูปแบบข้อมูลของไฟล์ Excel
 * ที่พบภายใน Checked Report Folder
 */
interface ExcelFileCandidate {
  /**
   * ชื่อไฟล์ ไม่รวม Folder
   */
  fileName: string;

  /**
   * Full Path ของไฟล์
   */
  filePath: string;

  /**
   * เวลาแก้ไขไฟล์ล่าสุดในหน่วยมิลลิวินาที
   */
  modifiedTime: number;
}

/**
 * แปลงค่าทั่วไปให้เป็นข้อความ
 *
 * การทำงาน
 * - null หรือ undefined → ""
 * - ค่าอื่น → แปลงเป็น string และ trim()
 */
const toText = (
  value: unknown,
): string => {
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(
    value,
  ).trim();
};

/**
 * เติมเลข 0 ด้านหน้าให้มีอย่างน้อย 2 หลัก
 *
 * ตัวอย่าง:
 * 1  → "01"
 * 9  → "09"
 * 10 → "10"
 *
 * ใช้สำหรับสร้างเดือน วัน ชั่วโมง นาที และวินาที
 */
const padTwoDigits = (
  value: number,
): string => {
  return String(
    value,
  ).padStart(
    2,
    "0",
  );
};

/**
 * ตรวจสอบว่า Path ไม่เป็นข้อความว่าง
 *
 * ถ้า Path ว่างจะ Throw Error
 *
 * @param targetPath Path ที่ต้องการตรวจสอบ
 * @param description ชื่อที่ใช้แสดงใน Error
 */
const validateNonEmptyPath = (
  targetPath: string,
  description: string,
): void => {
  if (
    toText(
      targetPath,
    ) === ""
  ) {
    throw new Error(
      `${description} path is empty`,
    );
  }
};

/**
 * ตรวจสอบว่า Path ที่กำหนดเป็น Folder จริง
 *
 * ขั้นตอน
 * 1. ตรวจว่า Path ไม่ว่าง
 * 2. ตรวจว่า Path มีอยู่จริง
 * 3. ตรวจว่า Path เป็น Directory
 *
 * @throws Error เมื่อไม่พบ Path หรือ Path ไม่ใช่ Folder
 */
const validateDirectoryExists = (
  directoryPath: string,
  description: string,
): void => {
  validateNonEmptyPath(
    directoryPath,
    description,
  );

  if (
    !fs.existsSync(
      directoryPath,
    )
  ) {
    throw new Error(
      `${description} not found: ${directoryPath}`,
    );
  }

  const directoryStatus =
    fs.statSync(
      directoryPath,
    );

  if (
    !directoryStatus.isDirectory()
  ) {
    throw new Error(
      `${description} is not a directory: ${directoryPath}`,
    );
  }
};

/**
 * ตรวจสอบว่า Path ที่กำหนดเป็นไฟล์จริง
 *
 * ขั้นตอน
 * 1. ตรวจว่า Path ไม่ว่าง
 * 2. ตรวจว่า Path มีอยู่จริง
 * 3. ตรวจว่า Path เป็น File
 *
 * @throws Error เมื่อไม่พบ Path หรือ Path ไม่ใช่ไฟล์
 */
const validateFileExists = (
  filePath: string,
  description: string,
): void => {
  validateNonEmptyPath(
    filePath,
    description,
  );

  if (
    !fs.existsSync(
      filePath,
    )
  ) {
    throw new Error(
      `${description} not found: ${filePath}`,
    );
  }

  const fileStatus =
    fs.statSync(
      filePath,
    );

  if (
    !fileStatus.isFile()
  ) {
    throw new Error(
      `${description} is not a file: ${filePath}`,
    );
  }
};

/**
 * ตรวจสอบว่าไฟล์ใช้นามสกุล .xlsx
 *
 * การตรวจสอบไม่สนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 *
 * ตัวอย่างที่ผ่าน:
 * - Report.xlsx
 * - Report.XLSX
 *
 * ตัวอย่างที่ไม่ผ่าน:
 * - Report.xls
 * - Report.csv
 */
const validateExcelExtension = (
  filePath: string,
  description: string,
): void => {
  const extension =
    path
      .extname(
        filePath,
      )
      .toLowerCase();

  if (
    extension !==
    EXCEL_FILE_EXTENSION
  ) {
    throw new Error(
      `${description} must be an .xlsx file: ${filePath}`,
    );
  }
};

/**
 * ตรวจสอบไฟล์ Test Data
 *
 * Test Data ต้อง
 * 1. มี Path
 * 2. ใช้นามสกุล .xlsx
 * 3. มีไฟล์อยู่จริง
 * 4. Path ต้องชี้ไปที่ File ไม่ใช่ Folder
 *
 * หมายเหตุ
 * ฟังก์ชันนี้ยังไม่ได้เปิดอ่านเนื้อหาภายใน Excel
 */
const validateTestDataFile = (
  testDataFilePath: string,
): void => {
  validateNonEmptyPath(
    testDataFilePath,
    "DS_FTX Test Data file",
  );

  validateExcelExtension(
    testDataFilePath,
    "DS_FTX Test Data file",
  );

  validateFileExists(
    testDataFilePath,
    "DS_FTX Test Data file",
  );
};

/**
 * ตรวจสอบว่าชื่อไฟล์ควรถูกมองเป็นไฟล์ชั่วคราว
 * หรือไฟล์ที่ไม่ควรนำมาใช้หรือไม่
 *
 * คืน true เมื่อชื่อขึ้นต้นด้วย
 * - ~$ ซึ่งเป็นไฟล์ชั่วคราวของ Microsoft Excel
 * - . ซึ่งมักเป็น Hidden File
 *
 * ตัวอย่าง:
 * - ~$EXPORT_DS_FTX.xlsx
 * - .DS_FTX.xlsx
 */
const isTemporaryExcelFile = (
  fileName: string,
): boolean => {
  const normalizedFileName =
    toText(
      fileName,
    );

  return (
    normalizedFileName.startsWith(
      "~$",
    ) ||
    normalizedFileName.startsWith(
      ".",
    )
  );
};

/**
 * ตรวจสอบนามสกุลของชื่อไฟล์
 *
 * คืน true เฉพาะไฟล์ .xlsx
 * โดยไม่สนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 */
const isExcelFile = (
  fileName: string,
): boolean => {
  return (
    path
      .extname(
        fileName,
      )
      .toLowerCase() ===
    EXCEL_FILE_EXTENSION
  );
};

/**
 * ตรวจสอบว่าชื่อไฟล์สามารถใช้เป็น Report DS_FTX ได้หรือไม่
 *
 * เงื่อนไขทั้งหมด:
 * 1. ต้องเป็นไฟล์ .xlsx
 * 2. ต้องไม่ใช่ไฟล์ชั่วคราวหรือ Hidden File
 * 3. ชื่อไฟล์ต้องมีคำว่า DS_FTX
 * 4. ชื่อไฟล์ต้องไม่มีคำว่า COMPARE_RESULT
 * 5. ชื่อไฟล์ต้องไม่มีคำว่า AUTOMATION_SUMMARY
 *
 * การตรวจข้อความในชื่อไฟล์
 * ไม่สนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 *
 * หมายเหตุ
 * ไม่ได้บังคับว่าชื่อไฟล์ต้องขึ้นต้นด้วย EXPORT_
 * และยังไม่ได้ตรวจสอบเนื้อหาภายใน Workbook
 */
const isDsFtxReportFile = (
  fileName: string,
): boolean => {
  if (
    !isExcelFile(
      fileName,
    )
  ) {
    return false;
  }

  if (
    isTemporaryExcelFile(
      fileName,
    )
  ) {
    return false;
  }

  const normalizedFileName =
    fileName.toUpperCase();

  if (
    !normalizedFileName.includes(
      "DS_FTX",
    )
  ) {
    return false;
  }

  if (
    normalizedFileName.includes(
      "COMPARE_RESULT",
    )
  ) {
    return false;
  }

  if (
    normalizedFileName.includes(
      "AUTOMATION_SUMMARY",
    )
  ) {
    return false;
  }

  return true;
};

/**
 * อ่านรายชื่อ Report DS_FTX ที่สามารถนำมาใช้ได้
 * จาก Folder ที่กำหนด
 *
 * การทำงาน
 * 1. ตรวจสอบว่า Folder มีอยู่จริง
 * 2. อ่านรายการภายใน Folder
 * 3. สนใจเฉพาะ File ไม่อ่าน Folder ย่อย
 * 4. กรองชื่อด้วย isDsFtxReportFile()
 * 5. เก็บชื่อ Path และเวลาแก้ไขไฟล์
 */
const getDsFtxReportCandidates = (
  directoryPath: string,
): ExcelFileCandidate[] => {
  validateDirectoryExists(
    directoryPath,
    "DS_FTX checked report directory",
  );

  const directoryEntries =
    fs.readdirSync(
      directoryPath,
      {
        withFileTypes:
          true,
      },
    );

  const candidates:
  ExcelFileCandidate[] = [];

  for (
    const directoryEntry of
    directoryEntries
  ) {
    /**
     * สนใจเฉพาะ File ใน Folder ปัจจุบัน
     * ไม่ค้นหา Report ภายใน Folder ย่อย
     */
    if (
      !directoryEntry.isFile()
    ) {
      continue;
    }

    if (
      !isDsFtxReportFile(
        directoryEntry.name,
      )
    ) {
      continue;
    }

    const filePath =
      path.join(
        directoryPath,
        directoryEntry.name,
      );

    const fileStatus =
      fs.statSync(
        filePath,
      );

    candidates.push({
      fileName:
        directoryEntry.name,

      filePath,

      modifiedTime:
        fileStatus.mtimeMs,
    });
  }

  return candidates;
};

/**
 * เรียงไฟล์จากใหม่ที่สุดไปเก่าที่สุด
 *
 * ลำดับการตัดสิน:
 *
 * 1. ใช้เวลาแก้ไขไฟล์ modifiedTime
 * 2. ถ้าเวลาเท่ากัน ใช้ชื่อไฟล์เรียงจากมากไปน้อย
 *
 * ฟังก์ชันสร้าง Array สำเนาก่อน Sort
 * จึงไม่เปลี่ยนลำดับของ candidates ต้นฉบับ
 */
const sortFilesFromNewest = (
  candidates: ExcelFileCandidate[],
): ExcelFileCandidate[] => {
  return [
    ...candidates,
  ].sort(
    (
      left,
      right,
    ) => {
      const modifiedTimeDifference =
        right.modifiedTime -
        left.modifiedTime;

      if (
        modifiedTimeDifference !== 0
      ) {
        return modifiedTimeDifference;
      }

      return right.fileName.localeCompare(
        left.fileName,
        undefined,
        {
          /**
           * เรียงตัวเลขตามค่าตัวเลข
           * เช่น 10 อยู่หลัง 9
           */
          numeric:
            true,

          /**
           * ไม่เน้นความต่างของตัวพิมพ์
           */
          sensitivity:
            "base",
        },
      );
    },
  );
};

/**
 * ค้นหา Report DS_FTX ล่าสุด
 *
 * คำว่า "ล่าสุด" ในฟังก์ชันนี้หมายถึง
 * ไฟล์ที่มีเวลาแก้ไขล่าสุด ไม่ใช่ Timestamp ในชื่อไฟล์
 *
 * ถ้ามีเวลาแก้ไขเท่ากัน
 * จะใช้ชื่อไฟล์เป็นตัวตัดสิน
 *
 * @param checkedReportDirectoryPath
 * Folder Checked Report ของ DS_FTX
 *
 * @returns Full Path ของ Report ที่เลือก
 */
export const getLatestFtxReportPath = (
  checkedReportDirectoryPath: string,
): string => {
  const candidates =
    getDsFtxReportCandidates(
      checkedReportDirectoryPath,
    );

  /**
   * ถ้าไม่พบไฟล์ที่ตรงเงื่อนไข
   * ให้หยุดและแจ้ง Error
   */
  if (
    candidates.length ===
    0
  ) {
    throw new Error(
      [
        "No matching DS_FTX Excel report was found",
        `Directory: ${checkedReportDirectoryPath}`,
        "Expected an .xlsx file whose name contains DS_FTX",
      ].join(
        " | ",
      ),
    );
  }

  const sortedCandidates =
    sortFilesFromNewest(
      candidates,
    );

  /**
   * เลือกไฟล์แรกหลังเรียงลำดับ
   */
  const latestFile =
    sortedCandidates[0];

  /**
   * ป้องกันกรณีไม่สามารถอ่านรายการแรกได้
   *
   * ตาม Flow ปกติ candidates ถูกตรวจแล้วว่าไม่ว่าง
   */
  if (
    latestFile === undefined
  ) {
    throw new Error(
      `Cannot determine latest DS_FTX report: ${checkedReportDirectoryPath}`,
    );
  }

  return latestFile.filePath;
};

/**
 * สร้าง Timestamp สำหรับชื่อไฟล์ผลลัพธ์
 *
 * รูปแบบ:
 * YYYYMMDD_HHmmss
 *
 * ตัวอย่าง:
 * 20260730_181530
 *
 * ใช้วันที่และเวลาตาม Local Time ของเครื่อง
 *
 * @param createdAt
 * วันที่และเวลาที่ต้องการใช้สร้าง Timestamp
 *
 * ค่าเริ่มต้นคือเวลาปัจจุบัน
 */
export const createFtxCompareTimestamp = (
  createdAt: Date = new Date(),
): string => {
  /**
   * ตรวจสอบว่าเป็น Date ที่ถูกต้องหรือไม่
   */
  if (
    Number.isNaN(
      createdAt.getTime(),
    )
  ) {
    throw new Error(
      "Cannot create DS_FTX compare timestamp from an invalid date",
    );
  }

  const year =
    createdAt.getFullYear();

  /**
   * JavaScript เริ่มนับเดือนจาก 0
   * จึงต้องบวก 1
   */
  const month =
    padTwoDigits(
      createdAt.getMonth() +
      1,
    );

  const day =
    padTwoDigits(
      createdAt.getDate(),
    );

  const hour =
    padTwoDigits(
      createdAt.getHours(),
    );

  const minute =
    padTwoDigits(
      createdAt.getMinutes(),
    );

  const second =
    padTwoDigits(
      createdAt.getSeconds(),
    );

  return (
    `${year}${month}${day}` +
    `_${hour}${minute}${second}`
  );
};

/**
 * สร้างชื่อไฟล์ผลลัพธ์จาก Timestamp
 *
 * ตัวอย่าง:
 *
 * Timestamp:
 * 20260730_181530
 *
 * ผลลัพธ์:
 * DS_FTX_Compare_Result_20260730_181530.xlsx
 *
 * @throws Error เมื่อ Timestamp เป็นข้อความว่าง
 */
export const createFtxCompareFileName = (
  timestamp: string,
): string => {
  const normalizedTimestamp =
    toText(
      timestamp,
    );

  if (
    normalizedTimestamp === ""
  ) {
    throw new Error(
      "DS_FTX compare timestamp is empty",
    );
  }

  return (
    `${COMPARE_RESULT_FILE_PREFIX}` +
    `_${normalizedTimestamp}` +
    `${EXCEL_FILE_EXTENSION}`
  );
};

/**
 * สร้าง Folder ผลลัพธ์ถ้ายังไม่มี
 *
 * recursive: true
 * ทำให้สามารถสร้าง Folder หลายระดับได้
 *
 * หลังสร้างแล้วจะตรวจสอบอีกครั้งว่า
 * Path เป็น Directory จริง
 */
const ensureOutputDirectoryExists = (
  outputDirectoryPath: string,
): void => {
  validateNonEmptyPath(
    outputDirectoryPath,
    "DS_FTX reconcile output directory",
  );

  fs.mkdirSync(
    outputDirectoryPath,
    {
      recursive:
        true,
    },
  );

  validateDirectoryExists(
    outputDirectoryPath,
    "DS_FTX reconcile output directory",
  );
};

/**
 * สร้าง Path ของไฟล์ผลลัพธ์ที่ไม่ซ้ำกับไฟล์เดิม
 *
 * ตัวอย่างชื่อเริ่มต้น:
 * DS_FTX_Compare_Result_20260730_181530.xlsx
 *
 * ถ้าชื่อเริ่มต้นมีอยู่แล้ว:
 * DS_FTX_Compare_Result_20260730_181530_01.xlsx
 *
 * ถ้ายังซ้ำ:
 * DS_FTX_Compare_Result_20260730_181530_02.xlsx
 *
 * ระบบจะลองลำดับสูงสุดถึง 9999
 *
 * ฟังก์ชันนี้ตรวจเฉพาะว่ามี Path อยู่แล้วหรือไม่
 * แต่ยังไม่ได้สร้างไฟล์จริง
 */
const createAvailableOutputFilePath = (
  outputDirectoryPath: string,
  timestamp: string,
): string => {
  const defaultFileName =
    createFtxCompareFileName(
      timestamp,
    );

  const defaultFilePath =
    path.join(
      outputDirectoryPath,
      defaultFileName,
    );

  /**
   * ถ้ายังไม่มีชื่อเริ่มต้น
   * ให้ใช้ Path นี้ได้ทันที
   */
  if (
    !fs.existsSync(
      defaultFilePath,
    )
  ) {
    return defaultFilePath;
  }

  /**
   * ลองสร้างชื่อใหม่โดยเติมหมายเลขต่อท้าย
   */
  for (
    let duplicateNumber = 1;
    duplicateNumber <=
    MAX_DUPLICATE_FILE_NUMBER;
    duplicateNumber += 1
  ) {
    /**
     * เติมเลข 0 ให้มีอย่างน้อย 2 หลัก
     *
     * 1   → 01
     * 10  → 10
     * 100 → 100
     */
    const duplicateText =
      String(
        duplicateNumber,
      ).padStart(
        2,
        "0",
      );

    const duplicateFileName =
      (
        `${COMPARE_RESULT_FILE_PREFIX}` +
        `_${timestamp}` +
        `_${duplicateText}` +
        `${EXCEL_FILE_EXTENSION}`
      );

    const duplicateFilePath =
      path.join(
        outputDirectoryPath,
        duplicateFileName,
      );

    /**
     * ถ้ายังไม่มี Path นี้
     * ให้คืนเป็น Output File Path
     */
    if (
      !fs.existsSync(
        duplicateFilePath,
      )
    ) {
      return duplicateFilePath;
    }
  }

  /**
   * ถ้าชื่อซ้ำตั้งแต่ชื่อเริ่มต้นจนถึง _9999
   * ให้หยุดและแจ้ง Error
   */
  throw new Error(
    [
      "Cannot create a unique DS_FTX compare result file name",
      `Directory: ${outputDirectoryPath}`,
      `Timestamp: ${timestamp}`,
    ].join(
      " | ",
    ),
  );
};

/**
 * แปลง Path ของ Test Data เป็น Absolute Path
 *
 * ถ้าเป็น Absolute Path อยู่แล้ว:
 * - ใช้ Path เดิม
 * - Normalize เครื่องหมายแบ่ง Folder
 *
 * ถ้าเป็น Relative Path:
 * - อ้างอิงจาก Project Root
 *
 * @param projectRootPath Root Folder ของโปรเจกต์
 * @param testDataFilePath Path ของ Test Data
 *
 * @returns Absolute Path ของ Test Data
 */
const resolveTestDataFilePath = (
  projectRootPath: string,
  testDataFilePath: string,
): string => {
  if (
    path.isAbsolute(
      testDataFilePath,
    )
  ) {
    return path.normalize(
      testDataFilePath,
    );
  }

  return path.resolve(
    projectRootPath,
    testDataFilePath,
  );
};

/**
 * เตรียม Path ทั้งหมดสำหรับ Script 3 ของ DS_FTX
 *
 * ขั้นตอน
 * 1. แปลง Project Root เป็น Absolute Path
 * 2. ตรวจสอบ Project Root
 * 3. สร้าง Checked Report Folder Path
 * 4. เลือก Report DS_FTX ล่าสุด
 * 5. Resolve และตรวจสอบ Test Data Path
 * 6. สร้าง Folder ผลลัพธ์
 * 7. สร้าง Timestamp
 * 8. สร้าง Output File Path ที่ไม่ซ้ำ
 * 9. คืนข้อมูล Path ทั้งหมด
 *
 * @param projectRootPath
 * Root Folder ของโปรเจกต์
 *
 * ปกติส่ง process.cwd()
 *
 * @param testDataFilePath
 * Path ของ Test Data จาก setting.ts
 *
 * รองรับทั้ง
 * - Absolute Path
 * - Relative Path
 *
 * @param createdAt
 * วันที่และเวลาที่ใช้สร้าง Timestamp
 *
 * ค่าเริ่มต้นคือเวลาปัจจุบัน
 *
 * @returns FtxCompareFilePaths
 */
export const prepareFtxCompareFilePaths = (
  projectRootPath: string,
  testDataFilePath: string,
  createdAt: Date = new Date(),
): FtxCompareFilePaths => {
  /**
   * แปลง Project Root เป็น Absolute Path
   */
  const resolvedProjectRootPath =
    path.resolve(
      projectRootPath,
    );

  /**
   * ตรวจสอบว่า Project Root เป็น Folder จริง
   */
  validateDirectoryExists(
    resolvedProjectRootPath,
    "Project root directory",
  );

  /**
   * สร้าง Absolute Path ของ Checked Report Folder
   */
  const checkedReportDirectoryPath =
    path.resolve(
      resolvedProjectRootPath,
      CHECKED_REPORT_FOLDER,
    );

  /**
   * เลือก Report DS_FTX ล่าสุด
   *
   * Checked Report Folder ต้องมีอยู่ก่อน
   * ฟังก์ชันนี้จะไม่สร้าง Folder ดังกล่าว
   */
  const reportFilePath =
    getLatestFtxReportPath(
      checkedReportDirectoryPath,
    );

  /**
   * Resolve Path ของ Test Data
   */
  const resolvedTestDataFilePath =
    resolveTestDataFilePath(
      resolvedProjectRootPath,
      testDataFilePath,
    );

  /**
   * ตรวจสอบ Test Data
   */
  validateTestDataFile(
    resolvedTestDataFilePath,
  );

  /**
   * สร้าง Absolute Path ของ Output Folder
   */
  const outputDirectoryPath =
    path.resolve(
      resolvedProjectRootPath,
      RECONCILE_RESULT_FOLDER,
    );

  /**
   * สร้าง Output Folder ถ้ายังไม่มี
   */
  ensureOutputDirectoryExists(
    outputDirectoryPath,
  );

  /**
   * สร้าง Timestamp สำหรับชื่อไฟล์
   */
  const timestamp =
    createFtxCompareTimestamp(
      createdAt,
    );

  /**
   * สร้าง Output File Path ที่ไม่ซ้ำ
   *
   * ขั้นตอนนี้ยังไม่ได้สร้างไฟล์จริง
   */
  const outputFilePath =
    createAvailableOutputFilePath(
      outputDirectoryPath,
      timestamp,
    );

  /**
   * คืน Path ทั้งหมดให้ Script 3
   */
  return {
    projectRootPath:
      resolvedProjectRootPath,

    checkedReportDirectoryPath,

    reportFilePath,

    testDataFilePath:
      resolvedTestDataFilePath,

    outputDirectoryPath,

    outputFilePath,

    timestamp,
  };
};

/**
 * แสดง Path หลักที่ Script 3 เลือกใช้ใน Terminal
 *
 * แสดงเฉพาะ
 * - Report File
 * - Test Data
 * - Output File
 *
 * ไม่ได้แสดง
 * - Project Root
 * - Checked Report Folder
 * - Output Folder
 * - Timestamp แยกต่างหาก
 */
export const printFtxCompareFilePaths = (
  filePaths: FtxCompareFilePaths,
): void => {
  console.log(
    "========================================",
  );

  console.log(
    "DS_FTX Compare File Paths",
  );

  console.log(
    `Report File : ${filePaths.reportFilePath}`,
  );

  console.log(
    `Test Data   : ${filePaths.testDataFilePath}`,
  );

  console.log(
    `Output File : ${filePaths.outputFilePath}`,
  );

  console.log(
    "========================================",
  );
};