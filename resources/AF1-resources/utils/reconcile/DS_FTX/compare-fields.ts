/**
 * compare-fields.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * เปรียบเทียบ Core Field ระหว่าง
 * - ExpectedRow จาก Test Data
 * - ActualRow จาก Report DS_FTX
 *
 * Field ที่ตรวจตาม DS_FTX_COMPARE_RULES
 *
 * 1. From Currency (CCY)
 *    เทียบกับ Buy Currency Id
 *    ด้วยวิธี TEXT
 *
 * 2. Settled Currency (CCY)
 *    เทียบกับ Sell Currency Id
 *    ด้วยวิธี TEXT
 *
 * 3. Txn Date
 *    เทียบกับ Transaction Date
 *    ด้วยวิธี DATE
 *
 * ผลลัพธ์ของแต่ละ Field
 * - PASS = ค่าตรงกันหลัง Normalize
 * - FAIL = ค่าไม่ตรงกันหลัง Normalize
 *
 * ExpectedRow และ ActualRow จำนวน 1 คู่
 * จะสร้าง CompareResult ตามจำนวน Rule
 *
 * ปัจจุบันมี 3 Rule
 * จึงสร้าง CompareResult จำนวน 3 รายการต่อคู่
 *
 * หมายเหตุสำคัญ
 * - ชื่อ Header ถูกอ่านแบบตรงตัว
 * - ไม่รองรับ Alias ในขั้นตอนอ่านค่า
 * - ถ้าค่าทั้งสองฝั่งว่างพร้อมกัน อาจได้ PASS
 * - การเปรียบเทียบวันที่ไม่ได้ตรวจความถูกต้องของปฏิทิน
 * ------------------------------------------------------------------
 */

import {
  DS_FTX_COMPARE_RULES,
  resolveExpectedValue,
} from "./ds-ftx-compare-config";

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
} from "./compare-types";

/**
 * ปรับรูปแบบข้อความก่อนนำไปเปรียบเทียบ
 *
 * การทำงาน
 * 1. เปลี่ยน null หรือ undefined เป็นข้อความว่าง
 * 2. แปลงค่าเป็น string
 * 3. ตัดช่องว่างด้านหน้าและด้านหลัง
 * 4. รวมช่องว่างหลายช่องให้เหลือช่องเดียว
 * 5. เปลี่ยนข้อความเป็นตัวพิมพ์ใหญ่
 *
 * ตัวอย่าง:
 *
 * " usd "          → "USD"
 * "United  State"  → "UNITED STATE"
 * "new\nyork"      → "NEW YORK"
 * null             → ""
 * undefined        → ""
 *
 * ผลที่เกิดขึ้น:
 * - ไม่สนตัวพิมพ์เล็กหรือพิมพ์ใหญ่
 * - ไม่สนช่องว่างหน้าและหลัง
 * - ไม่สนจำนวนช่องว่างระหว่างคำ
 *
 * แต่ยังสนใจ:
 * - ตัวอักษร
 * - ตัวเลข
 * - เครื่องหมายพิเศษ
 *
 * @param value ค่าที่ต้องการ Normalize
 * @returns ข้อความตัวพิมพ์ใหญ่ที่ปรับช่องว่างแล้ว
 */
const normalizeText = (
  value: unknown,
): string => {
  /**
   * value ?? ""
   * หมายถึง ถ้าค่าเป็น null หรือ undefined
   * ให้ใช้ข้อความว่างแทน
   */
  return String(
    value ?? "",
  )
    .trim()

    /**
     * รวม Whitespace ที่ต่อกันหลายตัว
     * ให้เหลือช่องว่างปกติเพียง 1 ช่อง
     *
     * รองรับเช่น
     * - ช่องว่าง
     * - Tab
     * - การขึ้นบรรทัดใหม่
     */
    .replace(
      /\s+/g,
      " ",
    )

    // เปลี่ยนเป็นตัวพิมพ์ใหญ่
    .toUpperCase();
};

/**
 * ปรับรูปแบบวันที่ให้อยู่ในรูปแบบเดียวกัน
 *
 * รูปแบบปลายทาง:
 * yyyy-mm-dd
 *
 * รูปแบบที่รองรับโดยตรง:
 * - dd/mm/yyyy
 * - dd-mm-yyyy
 * - yyyy/mm/dd
 * - yyyy-mm-dd
 *
 * รูปแบบปีนำหน้าสามารถมีเวลาต่อท้ายได้ เช่น
 * - 2025-11-25 00:00:00
 * - 2025-11-25T00:00:00
 *
 * ถ้าไม่ตรงกับ Pattern ด้านบน
 * จะลองแปลงด้วย new Date()
 *
 * ถ้ายังแปลงไม่ได้
 * จะนำค่าไป Normalize เป็นข้อความแทน
 *
 * หมายเหตุสำคัญ
 * การ Match ด้วย Regular Expression
 * ตรวจเฉพาะรูปแบบตัวเลข ไม่ได้ตรวจวันที่จริงในปฏิทิน
 *
 * ตัวอย่าง:
 * 31/02/2025 จะถูกแปลงเป็น 2025-02-31
 * แม้เดือนกุมภาพันธ์ไม่มีวันที่ 31
 *
 * @param value ค่าวันที่ที่ต้องการ Normalize
 * @returns วันที่รูปแบบ yyyy-mm-dd หรือข้อความที่ Normalize แล้ว
 */
const normalizeDate = (
  value: unknown,
): string => {
  /**
   * แปลงค่าเป็นข้อความและตัดช่องว่างหน้า–หลัง
   *
   * null หรือ undefined จะกลายเป็น ""
   */
  const text =
    String(
      value ?? "",
    ).trim();

  /**
   * ตรวจวันที่รูปแบบวันนำหน้า
   *
   * รูปแบบที่รองรับ:
   * - dd/mm/yyyy
   * - dd-mm-yyyy
   *
   * ตัวอย่าง:
   * 25/11/2025
   * 25-11-2025
   */
  const dayFirstMatch =
    text.match(
      /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/,
    );

  /**
   * ถ้าตรงกับรูปแบบวันนำหน้า
   * ให้จัดเรียงใหม่เป็น yyyy-mm-dd
   */
  if (
    dayFirstMatch
  ) {
    /**
     * แยกค่าที่ได้จาก Regular Expression
     *
     * ตำแหน่งแรกไม่ใช้
     * เพราะเป็นข้อความที่ Match ทั้งหมด
     */
    const [
      ,
      day,
      month,
      year,
    ] = dayFirstMatch;

    /**
     * padStart(2, "0")
     * เติมเลข 0 ให้วันและเดือนมี 2 หลัก
     *
     * ตัวอย่าง:
     * 5 → 05
     */
    return (
      `${year}-` +
      `${month.padStart(2, "0")}-` +
      `${day.padStart(2, "0")}`
    );
  }

  /**
   * ตรวจวันที่รูปแบบปีนำหน้า
   *
   * รูปแบบที่รองรับ:
   * - yyyy/mm/dd
   * - yyyy-mm-dd
   *
   * Regular Expression ไม่ได้กำหนด $ ไว้ท้าย Pattern
   * จึงรองรับข้อความที่มีเวลาต่อท้าย เช่น
   * - 2025-11-25 00:00:00
   * - 2025-11-25T00:00:00
   */
  const isoMatch =
    text.match(
      /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/,
    );

  /**
   * ถ้าตรงกับรูปแบบปีนำหน้า
   * ให้จัดรูปแบบเป็น yyyy-mm-dd
   */
  if (
    isoMatch
  ) {
    const [
      ,
      year,
      month,
      day,
    ] = isoMatch;

    return (
      `${year}-` +
      `${month.padStart(2, "0")}-` +
      `${day.padStart(2, "0")}`
    );
  }

  /**
   * ถ้าไม่ตรงกับ Pattern ด้านบน
   * ให้ลองใช้ JavaScript Date Parser
   *
   * ตัวอย่างที่อาจพบจาก ExcelJS:
   * Tue Nov 25 2025 00:00:00 GMT+0700
   */
  const parsedDate =
    new Date(
      text,
    );

  /**
   * ถ้า parsedDate เป็นวันที่ที่ JavaScript อ่านได้
   * ให้ดึงปี เดือน และวันตาม Local Time ของเครื่อง
   */
  if (
    !Number.isNaN(
      parsedDate.getTime(),
    )
  ) {
    const year =
      parsedDate.getFullYear();

    /**
     * JavaScript เริ่มนับเดือนจาก 0
     * จึงต้องบวก 1
     */
    const month =
      String(
        parsedDate.getMonth() + 1,
      ).padStart(
        2,
        "0",
      );

    const day =
      String(
        parsedDate.getDate(),
      ).padStart(
        2,
        "0",
      );

    return `${year}-${month}-${day}`;
  }

  /**
   * ถ้าไม่สามารถแปลงเป็นวันที่ได้
   * ให้เปรียบเทียบในรูปแบบข้อความแทน
   *
   * ตัวอย่าง:
   * "N/A" → "N/A"
   * ""    → ""
   */
  return normalizeText(
    value,
  );
};

/**
 * เปรียบเทียบ Core Field ของ ExpectedRow และ ActualRow
 *
 * การทำงาน
 * 1. วนอ่าน DS_FTX_COMPARE_RULES ทีละ Rule
 * 2. อ่าน Expected Value จาก Test Data
 * 3. อ่าน Actual Value จาก Report
 * 4. เลือกวิธี Normalize ตาม Compare Type
 * 5. เปรียบเทียบค่าด้วย ===
 * 6. สร้าง CompareResult ของแต่ละ Field
 *
 * @param expectedRow
 * ข้อมูลที่ระบบคาดหวังจาก Test Data
 *
 * @param actualRow
 * ข้อมูลจริงที่อ่านจาก Report DS_FTX
 *
 * @returns
 * Array ของ CompareResult ตามจำนวน Rule
 *
 * ปัจจุบันมี 3 Rule
 * จึงคืนผลลัพธ์ 3 รายการ
 */
export const compareCoreFields = (
  expectedRow: ExpectedRow,
  actualRow: ActualRow,
): CompareResult[] => {
  /**
   * map() จะเปลี่ยน CompareRule แต่ละรายการ
   * ให้เป็น CompareResult จำนวน 1 รายการ
   */
  return DS_FTX_COMPARE_RULES.map(
    (
      rule,
    ) => {
      /**
       * อ่านค่าที่คาดหวังจาก Test Data
       *
       * ตัวอย่าง Rule:
       *
       * Report Field:
       * Buy Currency Id
       *
       * Test Data Field:
       * From Currency (CCY)
       *
       * resolveExpectedValue() จะอ่าน:
       * expectedRow.data["From Currency (CCY)"]
       */
      const expectedValue =
        resolveExpectedValue(
          expectedRow,
          rule,
        );

      /**
       * อ่านค่าจริงจาก Report DS_FTX
       *
       * ใช้ reportField เป็นชื่อ Property
       *
       * ตัวอย่าง:
       * actualRow.data["Buy Currency Id"]
       *
       * ถ้าไม่พบ Property
       * ค่าที่ได้จะเป็น undefined
       */
      const actualValue =
        actualRow.data[
          rule.reportField
        ];

      /**
       * เลือกวิธีเปรียบเทียบตาม compareType
       *
       * DATE:
       * Normalize วันที่ทั้งสองฝั่งเป็น yyyy-mm-dd
       *
       * TEXT:
       * Normalize ช่องว่างและตัวพิมพ์
       *
       * หลัง Normalize แล้ว
       * จะเปรียบเทียบด้วย ===
       */
      const isMatched =
        rule.compareType ===
          "DATE"
          ? normalizeDate(
              expectedValue,
            ) ===
            normalizeDate(
              actualValue,
            )
          : normalizeText(
              expectedValue,
            ) ===
            normalizeText(
              actualValue,
            );

      /**
       * สร้าง CompareResult สำหรับ Field ปัจจุบัน
       */
      return {
        /**
         * Matching Key ของรายการ
         */
        matchingKey:
          expectedRow.matchingKey,

        /**
         * Test No. จาก Test Data
         */
        testScriptNo:
          expectedRow.testScriptNo,

        /**
         * หมายเลขแถวจริงใน Test Data
         */
        testDataRowNumber:
          expectedRow.rowNumber,

        /**
         * หมายเลขแถวจริงใน Report
         */
        reportRowNumber:
          actualRow.rowNumber,

        /**
         * ใช้ชื่อ Header ฝั่ง Report
         * เป็นชื่อ Field ในผลลัพธ์
         *
         * ตัวอย่าง:
         * Buy Currency Id
         */
        field:
          rule.reportField,

        /**
         * ค่าเดิมจาก Test Data ก่อน Normalize
         *
         * ทำให้ผู้ใช้งานเห็นค่าต้นทางจริง
         * ในไฟล์ผลลัพธ์
         */
        expected:
          expectedValue,

        /**
         * ค่าเดิมจาก Report ก่อน Normalize
         */
        actual:
          actualValue,

        /**
         * ผลการเปรียบเทียบ
         *
         * true  → PASS
         * false → FAIL
         */
        status:
          isMatched
            ? "PASS"
            : "FAIL",

        /**
         * กรณี PASS:
         * Remark เป็นข้อความว่าง
         *
         * กรณี FAIL:
         * Remark แสดง
         * - ชื่อ Report Field
         * - ค่าจาก Test Data
         * - ค่าจาก DS_FTX
         */
        remark:
          isMatched
            ? ""
            : `${rule.reportField}: Value Mismatch | [Test Data]: ${String(expectedValue ?? "")} | [DS_FTX]: ${String(actualValue ?? "")}`,
      };
    },
  );
};