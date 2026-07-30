/**
 * ============================================================================
 * ไฟล์: build-matching-key.ts
 * ============================================================================
 *
 * หน้าที่ของไฟล์นี้
 * -----------------
 * ไฟล์นี้ใช้สร้าง “รหัสสำหรับจับคู่ข้อมูล” หรือ Matching Key ของ DS_PTX
 *
 * ลองนึกภาพว่า Test Data และ Report เป็นเอกสารคนละชุด เราจึงต้องมีรหัสกลาง
 * เพื่อบอกว่าแถวใดใน Test Data ตรงกับแถวใดใน Report รหัสดังกล่าวประกอบจาก
 * Transaction ID, ลำดับของ Fee, System ID, Channel และคำลงท้าย RM
 *
 * ตัวอย่าง
 * --------
 * Transaction ID:
 * KBO3042511250000000081448
 *
 * Fee ลำดับที่ 1:
 * 01
 *
 * Matching Key ที่ได้:
 * KBO3042511250000000081448_01_GPMH_KBO_RM
 *
 * ขั้นตอนการทำงานโดยสรุป
 * ----------------------
 * 1. ทำความสะอาด Transaction ID โดยตัดช่องว่างและแปลงเป็นตัวพิมพ์ใหญ่
 * 2. ตรวจสอบว่า Transaction ID และ Running Number ใช้งานได้
 * 3. ดึง Channel จาก 3 ตัวแรกของ Transaction ID
 * 4. เติมเลข 0 ด้านหน้า Running Number เช่น 1 เป็น 01
 * 5. นำข้อมูลทั้งหมดมาต่อกันด้วยเครื่องหมาย _
 *
 * หากข้อมูลสำคัญไม่ครบ ฟังก์ชันจะหยุดและแจ้ง Error เพื่อไม่ให้สร้าง
 * Matching Key ที่ผิดแล้วนำไปเปรียบเทียบต่อ
 * ============================================================================
 */

/**
 * ============================================================================
 * ค่าคงที่ที่ใช้สร้าง Matching Key
 * ============================================================================
 */

/**
 * รหัสระบบของ PTX (System ID)
 */
const SYSTEM_ID = "GPMH";

/**
 * ข้อความต่อท้าย Matching Key ของ Report
 */
const REPORT_SUFFIX = "RM";

/**
 * จำนวนหลักของเลขลำดับ
 *
 * 1 → 01
 * 2 → 02
 * 3 → 03
 */
const RUNNING_FORMAT_LENGTH = 2;

/**
 * แปลง Running Number
 *
 * 1 → 01
 * 2 → 02
 * 3 → 03
 */
const formatRunningNumber = (
  runningNumber: number,
): string => {

  return String(
    runningNumber,
  ).padStart(
    RUNNING_FORMAT_LENGTH,
    "0",
  );

};

/**
 * หา Channel จาก Transaction ID
 *
 * ตัวอย่าง
 *
 * KBO304... → KBO
 * KMA301... → KMA
 * GPM000... → GPM
 */
const resolveChannelFromTransactionId = (
  transactionId: string,
): string => {

  const normalizedTransactionId =
    transactionId
      .trim()
      .toUpperCase();

  /**
   * Transaction ID ต้องมีอย่างน้อย 3 ตัวอักษร
   */
  if (
    normalizedTransactionId.length < 3
  ) {

    throw new Error(
      `Cannot resolve Channel from Transaction ID: "${transactionId}"`,
    );

  }

  return normalizedTransactionId.slice(
    0,
    3,
  );

};

/**
 * สร้างกุญแจสำหรับจับคู่ข้อมูล
 */
export const buildMatchingKey = (
  transactionId: string,
  runningNumber: number,
): string => {

  const normalizedTransactionId =
    transactionId
      .trim()
      .toUpperCase();

  /**
   * ป้องกันการสร้าง Matching Key
   * จาก Transaction ID ที่ไม่มีข้อมูล
   */
  if (
    normalizedTransactionId === ""
  ) {

    throw new Error(
      "Cannot build Matching Key because Transaction ID is blank",
    );

  }

  /**
   * Running Number ต้องเริ่มจาก 1
   */
  if (
    !Number.isInteger(
      runningNumber,
    ) ||
    runningNumber < 1
  ) {

    throw new Error(
      `Invalid Running Number: ${runningNumber}`,
    );

  }

  const channel =
    resolveChannelFromTransactionId(
      normalizedTransactionId,
    );

  return [

    normalizedTransactionId,

    formatRunningNumber(
      runningNumber,
    ),

    SYSTEM_ID,

    channel,

    REPORT_SUFFIX,

  ].join(
    "_",
  );

};
