/**
 * AmountComparator
 * ------------------------------------------------------------------
 * แปลง รวม และเปรียบเทียบ Amount
 * โดยใช้ DEFAULT_AMOUNT_TOLERANCE จาก ltx-config
 * ------------------------------------------------------------------
 */
import { DEFAULT_AMOUNT_TOLERANCE } from "./ltx-config";

/** ตัวคูณเผื่อ Floating-point error โดยไม่เปลี่ยน Business Tolerance */
const FLOATING_POINT_SAFETY_FACTOR = 4;

export interface AmountCompareResult {
  isMatch: boolean;
}

export class AmountComparator {
  /** แปลง string เป็นตัวเลข (ตัด comma คั่นหลักพัน) คืนค่า null ถ้าแปลงไม่ได้/ว่าง */
  parse(value: string): number | null {
    const cleaned = value.replace(/,/g, "").trim();
    if (cleaned === "") {
      return null;
    }
    const parsed = Number(cleaned);
    return Number.isNaN(parsed) ? null : parsed;
  }

 /**
 * รวมยอด Fee Amount ทุกกลุ่มที่ Builder ตรวจพบจาก Header
 * โดยข้ามค่าที่ว่างหรือแปลงเป็นตัวเลขไม่ได้
 */
  sum(values: string[]): number {
    return values.reduce((total, value) => {
      const parsed = this.parse(value);
      return parsed === null ? total : total + parsed;
    }, 0);
  }

  /** เปรียบเทียบตัวเลขและคืนผลว่าอยู่ภายใน Amount Tolerance หรือไม่ */
  compare(
    expected: string,
    actual: string,
    tolerance: number = DEFAULT_AMOUNT_TOLERANCE,
  ): AmountCompareResult {
    const expectedNumber = this.parse(expected);
    const actualNumber = this.parse(actual);

    if (expectedNumber === null || actualNumber === null) {
      return { isMatch: false };
    }

    const diff = Math.abs(expectedNumber - actualNumber);

    /**
     * ป้องกัน Floating-point error ของ JavaScript เช่น
     * 100.01 - 100 อาจได้ 0.010000000000005116
     * ทั้งที่ผลต่างทางธุรกิจคือ 0.01 และต้องอยู่ใน Tolerance
     */
    const floatingPointAllowance =
      Number.EPSILON *
      Math.max(
        1,
        Math.abs(expectedNumber),
        Math.abs(actualNumber),
      ) *
      FLOATING_POINT_SAFETY_FACTOR;

    return {
      isMatch: diff <= tolerance + floatingPointAllowance,
    };
  }
}