/**
 * actual-row-index.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * สร้างตารางค้นหา ActualRow ของ Report DS_FTX
 * โดยใช้ Matching Key เป็นกุญแจในการค้นหา
 *
 * Matching Key ของ Report มาจาก Header:
 * "Ref. TX No."
 *
 * ตารางค้นหาใช้ Map ซึ่งมีรูปแบบดังนี้:
 *
 * Map<
 *   Matching Key,
 *   ActualRow[]
 * >
 *
 * ตัวอย่าง:
 *
 * Map {
 *   "TX001" => [ActualRow],
 *   "TX002" => [ActualRow]
 * }
 *
 * ถ้า Matching Key ซ้ำ ระบบจะเก็บทุกแถวไว้ใน Array:
 *
 * Map {
 *   "TX001" => [
 *     ActualRow จากแถว 2,
 *     ActualRow จากแถว 8
 *   ]
 * }
 *
 * ทำให้ compare-validator.ts สามารถตรวจสอบได้ว่า
 * Matching Key เดียวกันปรากฏใน Report มากกว่า 1 แถวหรือไม่
 *
 * หมายเหตุสำคัญ
 * - Matching Key ว่างจะไม่ถูกเพิ่มลงใน Map
 * - Matching Key เป็นแบบ Case-sensitive
 * - ลบเฉพาะช่องว่างด้านหน้าและด้านหลัง
 * - ไม่ลบช่องว่างภายใน Matching Key
 * - ไม่เปลี่ยนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 * ------------------------------------------------------------------
 */

import {
  ActualRow,
} from "./compare-types";

/**
 * รูปแบบตารางค้นหา ActualRow ตาม Matching Key
 *
 * Key:
 * Matching Key จาก Header "Ref. TX No."
 *
 * Value:
 * Array ของ ActualRow ที่ใช้ Matching Key เดียวกัน
 *
 * ตัวอย่างกรณีปกติ:
 *
 * "TX001" => [ActualRow จากแถว 2]
 *
 * ตัวอย่างกรณี Matching Key ซ้ำ:
 *
 * "TX001" => [
 *   ActualRow จากแถว 2,
 *   ActualRow จากแถว 8
 * ]
 *
 * ปกติ Matching Key หนึ่งค่าควรพบเพียงหนึ่งแถว
 * แต่กำหนด Value เป็น Array เพื่อให้สามารถ
 * เก็บและตรวจสอบกรณีข้อมูลซ้ำได้
 */
export type ActualRowIndex =
  Map<string, ActualRow[]>;

/**
 * ปรับรูปแบบ Matching Key ก่อนนำไปสร้างหรือค้นหา Index
 *
 * การทำงาน
 * 1. ถ้าค่าเป็น null หรือ undefined ให้คืนข้อความว่าง
 * 2. แปลงค่าเป็น string
 * 3. ตัดช่องว่างด้านหน้าและด้านหลัง
 *
 * ตัวอย่าง:
 *
 * " TX001 " → "TX001"
 * 12345     → "12345"
 * null      → ""
 *
 * ฟังก์ชันนี้ไม่ได้
 * - เปลี่ยนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 * - ลบช่องว่างภายใน
 * - ลบเครื่องหมายพิเศษ
 *
 * ดังนั้นค่าต่อไปนี้ถือว่าเป็นคนละ Matching Key:
 *
 * "TX001"
 * "tx001"
 * "TX 001"
 *
 * หมายเหตุ
 * Logic นี้เหมือนกับ normalizeMatchingKey()
 * ที่อยู่ใน build-matching-key.ts
 *
 * แต่ Code ปัจจุบันเขียนแยกกันคนละฟังก์ชัน
 * หากมีการเปลี่ยน Logic ในอนาคตควรตรวจสอบทั้งสองไฟล์
 *
 * @param matchingKey Matching Key ที่ต้องการ Normalize
 *
 * @returns
 * - Matching Key ที่แปลงเป็นข้อความแล้ว
 * - ข้อความว่างเมื่อไม่มีค่า
 */
const normalizeMatchingKey = (
  matchingKey: unknown,
): string => {
  /**
   * ถ้าไม่มีค่า ให้คืนข้อความว่าง
   */
  if (
    matchingKey === null ||
    matchingKey === undefined
  ) {
    return "";
  }

  /**
   * แปลงค่าเป็น string
   * และตัดช่องว่างด้านหน้าและด้านหลัง
   */
  return String(
    matchingKey,
  ).trim();
};

/**
 * สร้างตารางค้นหา ActualRow จากข้อมูล Report ทั้งหมด
 *
 * ขั้นตอน
 * 1. สร้าง Map ว่าง
 * 2. วนอ่าน ActualRow ทีละแถว
 * 3. Normalize Matching Key
 * 4. ข้ามแถวที่ Matching Key ว่าง
 * 5. ถ้ายังไม่พบ Key ให้สร้าง Array ใหม่
 * 6. ถ้าพบ Key เดิมแล้ว ให้เพิ่ม ActualRow ต่อท้าย Array
 * 7. คืน Map ที่สร้างเสร็จแล้ว
 *
 * ตัวอย่าง Input:
 *
 * [
 *   {
 *     rowNumber: 2,
 *     matchingKey: "TX001",
 *     data: {}
 *   },
 *   {
 *     rowNumber: 3,
 *     matchingKey: "TX002",
 *     data: {}
 *   }
 * ]
 *
 * ผลลัพธ์:
 *
 * Map {
 *   "TX001" => [ActualRow จากแถว 2],
 *   "TX002" => [ActualRow จากแถว 3]
 * }
 *
 * ตัวอย่างกรณี Key ซ้ำ:
 *
 * Input:
 *
 * [
 *   {
 *     rowNumber: 2,
 *     matchingKey: "TX001",
 *     data: {}
 *   },
 *   {
 *     rowNumber: 8,
 *     matchingKey: "TX001",
 *     data: {}
 *   }
 * ]
 *
 * ผลลัพธ์:
 *
 * Map {
 *   "TX001" => [
 *     ActualRow จากแถว 2,
 *     ActualRow จากแถว 8
 *   ]
 * }
 *
 * @param actualRows
 * ActualRow ทั้งหมดที่สร้างจาก Report DS_FTX
 *
 * @returns
 * Map สำหรับค้นหา ActualRow ด้วย Matching Key
 */
export const buildActualRowIndex = (
  actualRows: ActualRow[],
): ActualRowIndex => {
  /**
   * สร้าง Map ว่างสำหรับเก็บ Index
   */
  const index:
  ActualRowIndex = new Map();

  /**
   * วนอ่าน ActualRow ทีละรายการ
   * ตามลำดับเดิมใน Report
   */
  for (
    const actualRow of
    actualRows
  ) {
    /**
     * Normalize Matching Key ก่อนนำไปใช้
     */
    const matchingKey =
      normalizeMatchingKey(
        actualRow.matchingKey,
      );

    /**
     * ถ้า Matching Key ว่าง
     * จะไม่เพิ่ม ActualRow ลงใน Index
     *
     * เพราะไม่มี Key สำหรับจับคู่กับ Test Data
     *
     * หมายเหตุ:
     * ActualRow ยังอยู่ใน actualRows เดิม
     * เพียงแต่ไม่อยู่ใน Map นี้
     */
    if (
      matchingKey === ""
    ) {
      continue;
    }

    /**
     * ตรวจสอบว่า Matching Key นี้
     * มีข้อมูลอยู่ใน Map แล้วหรือไม่
     *
     * ถ้าไม่มีจะได้ undefined
     * ถ้ามีจะได้ ActualRow[]
     */
    const existingRows =
      index.get(
        matchingKey,
      );

    /**
     * ถ้ายังไม่เคยพบ Matching Key นี้
     * ให้สร้าง Array ใหม่และเก็บ ActualRow ปัจจุบัน
     */
    if (
      existingRows ===
      undefined
    ) {
      index.set(
        matchingKey,
        [
          actualRow,
        ],
      );

      /**
       * ไปตรวจ ActualRow รายการถัดไป
       */
      continue;
    }

    /**
     * ถ้ามี Matching Key นี้อยู่แล้ว
     * แสดงว่าพบ Key เดิมมากกว่าหนึ่งแถว
     *
     * เพิ่ม ActualRow ปัจจุบันต่อท้าย Array
     * โดยไม่เขียนทับ ActualRow ที่เก็บไว้ก่อนหน้า
     *
     * ไม่ต้องเรียก index.set() ซ้ำ
     * เพราะ existingRows อ้างอิง Array เดียวกับที่อยู่ใน Map
     */
    existingRows.push(
      actualRow,
    );
  }

  /**
   * คืน Map ที่สร้างเสร็จแล้ว
   */
  return index;
};

/**
 * ค้นหา ActualRow ทั้งหมดด้วย Matching Key
 *
 * ขั้นตอน
 * 1. Normalize Matching Key ที่ได้รับมา
 * 2. ถ้า Key ว่าง ให้คืน Array ว่าง
 * 3. ค้นหา Key ภายใน Map
 * 4. ถ้าไม่พบ ให้คืน Array ว่าง
 * 5. ถ้าพบ ให้คืน ActualRow[] ของ Key นั้น
 *
 * ตัวอย่าง:
 *
 * Index:
 *
 * Map {
 *   "TX001" => [
 *     ActualRow จากแถว 2,
 *     ActualRow จากแถว 8
 *   ]
 * }
 *
 * ค้นหา:
 *
 * findActualRowsByMatchingKey(
 *   index,
 *   "TX001",
 * );
 *
 * ผลลัพธ์:
 *
 * [
 *   ActualRow จากแถว 2,
 *   ActualRow จากแถว 8
 * ]
 *
 * @param index Map ที่สร้างจาก buildActualRowIndex()
 * @param matchingKey Matching Key ที่ต้องการค้นหา
 *
 * @returns
 * - ActualRow[] เมื่อพบ Matching Key
 * - Array ว่าง [] เมื่อไม่พบหรือ Matching Key ว่าง
 *
 * หมายเหตุ
 * กรณีพบ Key ฟังก์ชันจะคืน Array ตัวเดียวกับที่เก็บใน Map
 * ถ้าผู้เรียกแก้ไข Array นี้ อาจกระทบข้อมูลภายใน Index
 */
export const findActualRowsByMatchingKey = (
  index: ActualRowIndex,
  matchingKey: unknown,
): ActualRow[] => {
  /**
   * Normalize Matching Key ก่อนค้นหา
   */
  const normalizedKey =
    normalizeMatchingKey(
      matchingKey,
    );

  /**
   * ถ้า Matching Key ว่าง
   * ไม่ต้องค้นหาและคืน Array ว่างทันที
   */
  if (
    normalizedKey === ""
  ) {
    return [];
  }

  /**
   * ค้นหา ActualRow[] จาก Map
   *
   * ถ้าไม่พบ Key
   * index.get() จะคืน undefined
   *
   * ?? [] จะเปลี่ยน undefined เป็น Array ว่าง
   */
  return (
    index.get(
      normalizedKey,
    ) ?? []
  );
};