/**
 * file-system.util.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * รวม Helper Function ที่เกี่ยวข้องกับ File System
 * และใช้ร่วมกันภายในโปรเจกต์
 *
 * ความสามารถหลัก
 * 1. สร้าง Folder
 * 2. ค้นหาไฟล์ Excel ที่แก้ไขล่าสุด
 * 3. คัดลอกไฟล์ไปยัง Folder ปลายทาง
 * 4. ลบไฟล์
 * 5. สร้าง Timestamp
 * 6. สร้าง Path ของไฟล์ผลลัพธ์
 * 7. สร้าง Path ของไฟล์ชั่วคราว
 * 8. ลบไฟล์ชั่วคราวที่ค้างอยู่
 *
 * คำศัพท์
 * - File System = ระบบจัดการไฟล์และ Folder
 * - Path        = ตำแหน่งของไฟล์หรือ Folder
 * - Full Path   = ตำแหน่งแบบเต็มของไฟล์
 * - Timestamp   = วันที่และเวลาที่นำมาต่อท้ายชื่อไฟล์
 * - Temporary File = ไฟล์ชั่วคราว
 * ------------------------------------------------------------------
 */

import * as fs from "fs";
import * as path from "path";

/**
 * ตรวจสอบและสร้าง Folder ถ้ายังไม่มี
 *
 * recursive: true หมายความว่า
 * ถ้า Folder ชั้นก่อนหน้ายังไม่มี ระบบจะสร้างให้ด้วย
 *
 * ตัวอย่าง:
 * Test_result/Reconcile-report/DS_PTX
 *
 * ถ้าไม่มีทั้ง Reconcile-report และ DS_PTX
 * ระบบสามารถสร้าง Folder ทั้งสองระดับได้
 *
 * @param dirPath Path ของ Folder ที่ต้องการสร้าง
 *
 * @returns ไม่มีค่าที่ส่งกลับ
 */
export const ensureDirectoryExists = (
  dirPath: string,
): void => {
  // ตรวจสอบก่อนว่า Path มีอยู่แล้วหรือไม่
  if (!fs.existsSync(dirPath)) {
    /**
     * สร้าง Folder พร้อม Folder ชั้นก่อนหน้า
     * ที่ยังไม่มีอยู่
     */
    fs.mkdirSync(
      dirPath,
      {
        recursive: true,
      },
    );
  }
};

/**
 * ค้นหาไฟล์ Excel ที่มีเวลาแก้ไขล่าสุด
 * ภายใน Folder ที่กำหนด
 *
 * ขั้นตอน
 * 1. ตรวจสอบว่า Folder มีอยู่จริง
 * 2. อ่านรายการภายใน Folder
 * 3. เลือกชื่อที่ลงท้ายด้วย .xlsx หรือ .xls
 * 4. ตัดไฟล์ชั่วคราวของ Excel ที่ขึ้นต้นด้วย ~$ ออก
 * 5. อ่านเวลาแก้ไขล่าสุดของแต่ละไฟล์
 * 6. เรียงจากใหม่ที่สุดไปเก่าที่สุด
 * 7. คืน Full Path ของรายการแรก
 *
 * หมายเหตุสำคัญ
 * - ใช้เวลาแก้ไขไฟล์ หรือ modifiedTime
 * - ไม่ได้เลือกจาก Timestamp ที่อยู่ในชื่อไฟล์
 * - ตรวจนามสกุลแบบตัวพิมพ์เล็กเท่านั้น
 * - ไม่ค้นหาไฟล์ใน Folder ย่อย
 *
 * @param folderPath Folder ที่ต้องการค้นหาไฟล์
 *
 * @returns Full Path ของไฟล์ที่แก้ไขล่าสุด
 *
 * @throws
 * - Folder not found เมื่อไม่พบ Folder
 * - No Excel files found เมื่อไม่พบไฟล์ Excel
 */
export const getLatestFile = (
  folderPath: string,
): string => {
  /**
   * ตรวจสอบว่า Folder ที่ต้องการค้นหามีอยู่จริง
   */
  if (!fs.existsSync(folderPath)) {
    throw new Error(
      `Folder not found: ${folderPath}`,
    );
  }

  /**
   * อ่านรายชื่อภายใน Folder
   * และกรองให้เหลือเฉพาะไฟล์ Excel
   */
  const excelFiles = fs
    .readdirSync(
      folderPath,
    )

    /**
     * เลือกเฉพาะชื่อที่ลงท้ายด้วย
     * - .xlsx
     * - .xls
     *
     * การตรวจสอบนี้เป็นแบบ Case-sensitive
     * จึงไม่รองรับ .XLSX หรือ .XLS
     */
    .filter(
      (fileName) =>
        fileName.endsWith(".xlsx") ||
        fileName.endsWith(".xls"),
    )

    /**
     * ตัดไฟล์ชั่วคราวของ Microsoft Excel ออก
     *
     * เมื่อเปิดไฟล์ Excel อาจพบชื่อประมาณ:
     * ~$DS_PTX_Report.xlsx
     */
    .filter(
      (fileName) =>
        !fileName.startsWith("~$"),
    );

  /**
   * ถ้าไม่พบชื่อไฟล์ Excel
   * ให้หยุดการทำงานและแจ้ง Error
   */
  if (excelFiles.length === 0) {
    throw new Error(
      `No Excel files found in folder: ${folderPath}`,
    );
  }

  /**
   * เปลี่ยนชื่อไฟล์แต่ละรายการให้เป็น Object
   * ที่มีข้อมูลสำหรับนำมาเรียงลำดับ
   *
   * Object ประกอบด้วย
   * - fileName     = ชื่อไฟล์
   * - fullPath     = Path แบบเต็ม
   * - modifiedTime = เวลาแก้ไขไฟล์ หน่วยมิลลิวินาที
   */
  const [latestFile] = excelFiles
    .map(
      (fileName) => {
        // สร้าง Full Path ของรายการ
        const fullPath =
          path.join(
            folderPath,
            fileName,
          );

        return {
          fileName,
          fullPath,

          /**
           * mtimeMs คือเวลาที่ไฟล์ถูกแก้ไขล่าสุด
           * ในรูปแบบมิลลิวินาที
           */
          modifiedTime:
            fs.statSync(
              fullPath,
            ).mtimeMs,
        };
      },
    )

    /**
     * เรียงจากเวลาแก้ไขใหม่ที่สุด
     * ไปหาเวลาแก้ไขเก่าที่สุด
     */
    .sort(
      (a, b) =>
        b.modifiedTime -
        a.modifiedTime,
    );

  // แสดงชื่อไฟล์ล่าสุดใน Terminal
  console.log(
    `Latest File : ${latestFile.fileName}`,
  );

  // คืน Full Path ของไฟล์ล่าสุด
  return latestFile.fullPath;
};

/**
 * คัดลอกไฟล์จากตำแหน่งต้นทาง
 * ไปยัง Folder ปลายทาง
 *
 * ขั้นตอน
 * 1. สร้าง Folder ปลายทางถ้ายังไม่มี
 * 2. ดึงชื่อไฟล์จาก Source Path
 * 3. สร้าง Target Path
 * 4. คัดลอกไฟล์
 * 5. คืน Target Path
 *
 * หมายเหตุสำคัญ
 * ถ้า Folder ปลายทางมีไฟล์ชื่อเดียวกันอยู่แล้ว
 * fs.copyFileSync() จะเขียนทับไฟล์เดิม
 *
 * @param sourceFilePath Path ของไฟล์ต้นทาง
 * @param targetDirectory Folder ปลายทาง
 *
 * @returns Full Path ของไฟล์ที่ถูกคัดลอก
 */
export const copyFileToDirectory = (
  sourceFilePath: string,
  targetDirectory: string,
): string => {
  /**
   * ตรวจสอบและสร้าง Folder ปลายทาง
   */
  ensureDirectoryExists(
    targetDirectory,
  );

  /**
   * ดึงเฉพาะชื่อไฟล์ออกจาก Source Path
   *
   * ตัวอย่าง:
   * C:\Report\EXPORT_DS_PTX.xlsx
   *
   * ผลลัพธ์:
   * EXPORT_DS_PTX.xlsx
   */
  const fileName =
    path.basename(
      sourceFilePath,
    );

  /**
   * รวม Folder ปลายทางกับชื่อไฟล์
   * เพื่อสร้าง Full Path ปลายทาง
   */
  const targetFilePath =
    path.join(
      targetDirectory,
      fileName,
    );

  /**
   * คัดลอกไฟล์ต้นทางไปยังปลายทาง
   *
   * ถ้ามีไฟล์ชื่อเดียวกันอยู่แล้ว
   * ไฟล์เดิมจะถูกเขียนทับ
   */
  fs.copyFileSync(
    sourceFilePath,
    targetFilePath,
  );

  // คืน Path ของไฟล์ที่คัดลอกแล้ว
  return targetFilePath;
};

/**
 * ลบไฟล์ตาม Path ที่กำหนด ถ้าไฟล์นั้นมีอยู่
 *
 * การทำงาน
 * - ถ้าไม่พบไฟล์ จะจบการทำงานโดยไม่แจ้ง Error
 * - ถ้าพบไฟล์ จะพยายามลบด้วย fs.unlinkSync()
 * - ถ้าลบไม่สำเร็จ จะ Throw Error
 *
 * สาเหตุที่อาจลบไม่สำเร็จ เช่น
 * - ไฟล์กำลังเปิดอยู่
 * - ไฟล์ถูก Process อื่นใช้งาน
 * - ไม่มี Permission
 * - Path ที่ได้รับมาไม่ใช่ไฟล์
 *
 * @param filePath Path ของไฟล์ที่ต้องการลบ
 *
 * @returns ไม่มีค่าที่ส่งกลับ
 */
export const deleteFileIfExists = (
  filePath: string,
): void => {
  /**
   * ถ้าไม่พบ Path ไม่ต้องดำเนินการต่อ
   */
  if (!fs.existsSync(filePath)) {
    return;
  }

  try {
    // ลบไฟล์ตาม Path ที่ได้รับมา
    fs.unlinkSync(
      filePath,
    );
  } catch {
    /**
     * แจ้ง Error ถ้าลบไฟล์ไม่สำเร็จ
     *
     * ข้อความ Error ระบุว่าไฟล์อาจเปิดอยู่หรือถูกล็อก
     * แต่สาเหตุจริงอาจเป็น Permission หรือปัญหาอื่นได้เช่นกัน
     */
    throw new Error(
      `Cannot delete file because it may be open or locked: ${filePath}`,
    );
  }
};

/**
 * สร้าง Timestamp จากวันที่และเวลาปัจจุบันของเครื่อง
 *
 * รูปแบบ:
 * YYYYMMDD_HHMMSS
 *
 * ความหมาย:
 * - YYYY = ปี 4 หลัก
 * - MM   = เดือน 2 หลัก
 * - DD   = วันที่ 2 หลัก
 * - HH   = ชั่วโมง 2 หลัก
 * - MM   = นาที 2 หลัก
 * - SS   = วินาที 2 หลัก
 *
 * ตัวอย่าง:
 * 20260730_143025
 *
 * หมายเหตุ
 * - ใช้ Local Time ของเครื่องที่ Run
 * - ความละเอียดของ Timestamp อยู่ที่ระดับวินาที
 *
 * @returns Timestamp ในรูปแบบ YYYYMMDD_HHMMSS
 */
export const getTimestamp = (): string => {
  // อ่านวันที่และเวลาปัจจุบันจากเครื่อง
  const now =
    new Date();

  /**
   * เติมเลข 0 ด้านหน้า
   * เพื่อให้ตัวเลขมีอย่างน้อย 2 หลัก
   *
   * ตัวอย่าง:
   * 7  → 07
   * 12 → 12
   */
  const pad2 = (
    value: number,
  ): string =>
    String(value).padStart(
      2,
      "0",
    );

  // อ่านปีปัจจุบัน
  const year =
    now.getFullYear();

  /**
   * JavaScript เริ่มนับเดือนจาก 0
   *
   * 0  = มกราคม
   * 1  = กุมภาพันธ์
   * 11 = ธันวาคม
   *
   * จึงต้องบวก 1
   */
  const month =
    pad2(
      now.getMonth() + 1,
    );

  // อ่านวันที่
  const day =
    pad2(
      now.getDate(),
    );

  // อ่านชั่วโมง
  const hour =
    pad2(
      now.getHours(),
    );

  // อ่านนาที
  const minute =
    pad2(
      now.getMinutes(),
    );

  // อ่านวินาที
  const second =
    pad2(
      now.getSeconds(),
    );

  /**
   * รวมวันที่และเวลาเป็น Timestamp
   */
  return `${year}${month}${day}_${hour}${minute}${second}`;
};

/**
 * สร้าง Full Path ของไฟล์ที่มี Timestamp ต่อท้ายชื่อ
 *
 * ตัวอย่าง Input:
 *
 * directory:
 * Test_result/Reconcile-report/DS_PTX
 *
 * baseName:
 * DS_PTX_Compare_Result
 *
 * extension:
 * .xlsx
 *
 * ตัวอย่างผลลัพธ์:
 * Test_result/Reconcile-report/DS_PTX/
 * DS_PTX_Compare_Result_20260730_143025.xlsx
 *
 * หมายเหตุสำคัญ
 * ฟังก์ชันนี้สร้างเฉพาะข้อความ Path
 * แต่ไม่ได้สร้าง Folder ให้โดยอัตโนมัติ
 *
 * ผู้เรียกควรเรียก ensureDirectoryExists()
 * ก่อนนำ Path ไปเขียนไฟล์
 *
 * @param directory Folder สำหรับเก็บไฟล์
 * @param baseName ชื่อหลักของไฟล์
 * @param extension นามสกุลไฟล์ เช่น .xlsx
 *
 * @returns Full Path ที่มี Timestamp
 */
export const buildTimestampedFilePath = (
  directory: string,
  baseName: string,
  extension: string,
): string => {
  /**
   * รวม Folder ชื่อไฟล์ Timestamp และนามสกุล
   */
  return path.join(
    directory,
    `${baseName}_${getTimestamp()}${extension}`,
  );
};

/**
 * สร้าง Path ของไฟล์ชั่วคราว
 * จาก Path ของไฟล์ผลลัพธ์จริง
 *
 * ระบบจะเติมคำว่า _temp
 * ไว้ก่อนนามสกุลไฟล์
 *
 * ตัวอย่าง:
 *
 * Input:
 * DS_PTX_Compare_Result.xlsx
 *
 * Output:
 * DS_PTX_Compare_Result_temp.xlsx
 *
 * ฟังก์ชันนี้สร้างเฉพาะข้อความ Path
 * ยังไม่ได้สร้างไฟล์จริง
 *
 * @param finalFilePath Path ของไฟล์ผลลัพธ์จริง
 *
 * @returns Path สำหรับไฟล์ชั่วคราว
 */
export const buildTempFilePath = (
  finalFilePath: string,
): string => {
  /**
   * ดึงตำแหน่ง Folder จาก Path ของไฟล์จริง
   */
  const directory =
    path.dirname(
      finalFilePath,
    );

  /**
   * ดึงนามสกุลของไฟล์
   *
   * ตัวอย่าง:
   * .xlsx
   */
  const extension =
    path.extname(
      finalFilePath,
    );

  /**
   * ดึงชื่อไฟล์โดยไม่รวมนามสกุล
   *
   * ตัวอย่าง:
   * DS_PTX_Compare_Result
   */
  const baseName =
    path.basename(
      finalFilePath,
      extension,
    );

  /**
   * สร้าง Path ใหม่โดยเติม _temp
   * ก่อนนามสกุลไฟล์
   */
  return path.join(
    directory,
    `${baseName}_temp${extension}`,
  );
};

/**
 * ค้นหาและลบไฟล์ชั่วคราวภายใน Folder
 *
 * รูปแบบชื่อไฟล์ที่ถูกลบ:
 * _temp.<นามสกุลไฟล์>
 *
 * ตัวอย่าง:
 * - Result_temp.xlsx
 * - Report_temp.xls
 * - Output_temp.txt
 *
 * หมายเหตุสำคัญ
 * - ไม่ได้ตรวจว่าไฟล์มีอายุเท่าไร
 * - ลบทุกชื่อที่ตรงกับรูปแบบ
 * - ไม่ค้นหาใน Folder ย่อย
 * - ไม่ได้จำกัดเฉพาะไฟล์ Excel
 *
 * @param directory Folder ที่ต้องการค้นหาไฟล์ชั่วคราว
 *
 * @returns ไม่มีค่าที่ส่งกลับ
 */
export const cleanupStaleTempFiles = (
  directory: string,
): void => {
  /**
   * ถ้าไม่พบ Folder
   * ไม่ต้องดำเนินการต่อ
   */
  if (!fs.existsSync(directory)) {
    return;
  }

  /**
   * อ่านชื่อทั้งหมดภายใน Folder ปัจจุบัน
   */
  fs.readdirSync(
    directory,
  )

    /**
     * เลือกชื่อที่ลงท้ายด้วย
     * _temp.<นามสกุล>
     *
     * Regular Expression:
     * _temp  = ต้องมีคำว่า _temp
     * \.     = ตามด้วยเครื่องหมายจุด
     * [^.]+  = ตามด้วยอักขระที่ไม่ใช่จุดอย่างน้อย 1 ตัว
     * $      = ต้องอยู่ท้ายชื่อ
     */
    .filter(
      (fileName) =>
        /_temp\.[^.]+$/.test(
          fileName,
        ),
    )

    /**
     * ลบรายการที่พบทีละไฟล์
     */
    .forEach(
      (fileName) =>
        deleteFileIfExists(
          path.join(
            directory,
            fileName,
          ),
        ),
    );
};