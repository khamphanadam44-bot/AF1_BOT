/**
 * แปลง รวม และเปรียบเทียบ Amount ของ DS_LTX
 * เพื่อให้ Matching และ Field Validation ใช้กติกาเดียวกัน
 */

import { DEFAULT_AMOUNT_TOLERANCE } from "./ltx-config";

/**
 * ชดเชยความคลาดเคลื่อนระดับ Floating-point ของ JavaScript
 * โดยไม่เพิ่ม Business Tolerance ที่กำหนดไว้
 */
const FLOATING_POINT_SAFETY_FACTOR = 4;

export class AmountComparator {
  /**
   * แปลงข้อความ Amount เป็นตัวเลข
   *
   * คืน null เมื่อข้อมูลว่าง ไม่ใช่ตัวเลข หรือเป็น Infinity
   * เพื่อป้องกันข้อมูลเสียถูกนำไป Match หรือคำนวณต่อ
   */
  parse(value: string): number | null {
    const cleaned = value
      .replace(/,/g, "")
      .trim();

    if (cleaned === "") {
      return null;
    }

    const parsed = Number(cleaned);

    return Number.isFinite(parsed)
      ? parsed
      : null;
  }

  /**
   * รวม Fee Amount จากทุกกลุ่มที่ตรวจพบ
   *
   * Fee บางกลุ่มอาจไม่มีข้อมูล จึงนำมารวมเฉพาะค่าที่
   * สามารถแปลงเป็นตัวเลขได้เท่านั้น
   */
  sum(values: readonly string[]): number {
    return values.reduce(
      (
        total,
        value,
      ) => {
        const parsed =
          this.parse(value);

        return parsed === null
          ? total
          : total + parsed;
      },
      0,
    );
  }

  /**
   * ตรวจว่า Expected Amount และ Actual Amount
   * มีผลต่างไม่เกิน Business Tolerance หรือไม่
   *
   * Floating-point allowance ใช้แก้ข้อจำกัดการเก็บเลขทศนิยม
   * ของ JavaScript เท่านั้น ไม่ได้ขยาย Business Tolerance
   */
  matches(
    expected: string,
    actual: string,
    tolerance: number =
      DEFAULT_AMOUNT_TOLERANCE,
  ): boolean {
    const expectedNumber =
      this.parse(expected);

    const actualNumber =
      this.parse(actual);

    if (
      expectedNumber === null ||
      actualNumber === null
    ) {
      return false;
    }

    const difference =
      Math.abs(
        expectedNumber -
        actualNumber,
      );

    /**
     * ตัวอย่าง:
     * 100.01 - 100 อาจได้ 0.010000000000005116
     * ทั้งที่ผลต่างทางธุรกิจคือ 0.01
     */
    const floatingPointAllowance =
      Number.EPSILON *
      Math.max(
        1,
        Math.abs(expectedNumber),
        Math.abs(actualNumber),
      ) *
      FLOATING_POINT_SAFETY_FACTOR;

    return (
      difference <=
      tolerance +
        floatingPointAllowance
    );
  }
}