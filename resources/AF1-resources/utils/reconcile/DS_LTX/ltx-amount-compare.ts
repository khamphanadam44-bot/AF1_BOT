/**
 * AmountComparator
 * ------------------------------------------------------------------
 * แปลง รวม และเปรียบเทียบ Amount
 * โดยใช้ DEFAULT_AMOUNT_TOLERANCE จาก ltx-config
 * ------------------------------------------------------------------
 */
import { DEFAULT_AMOUNT_TOLERANCE } from "./ltx-config";

export interface AmountCompareResult {
  isMatch: boolean;
  expectedNumber: number | null;
  actualNumber: number | null;
  diff: number | null;
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

  /** เปรียบเทียบตัวเลขแบบ tolerance — คืนรายละเอียดสำหรับใส่ใน Remark */
  compare(
    expected: string,
    actual: string,
    tolerance: number = DEFAULT_AMOUNT_TOLERANCE,
  ): AmountCompareResult {
    const expectedNumber = this.parse(expected);
    const actualNumber = this.parse(actual);

    if (expectedNumber === null || actualNumber === null) {
      return { isMatch: false, expectedNumber, actualNumber, diff: null };
    }

    const diff = Math.abs(expectedNumber - actualNumber);
    return { isMatch: diff <= tolerance, expectedNumber, actualNumber, diff };
  }
}
