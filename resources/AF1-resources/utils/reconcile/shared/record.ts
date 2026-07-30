/**
 * ReconcileRecord
 * ------------------------------------------------------------------
 * Domain model 1 แถวข้อมูล (AF1 Report หรือ Test Data)
 * รวม logic "ดึงค่า field ตามชื่อ header" ไว้ในตัวเอง (แทน getReconcileField แบบ
 * function เดิมที่กระจายอยู่ทั้งใน reconcile-record-reader.ts และ report-data-reader.ts
 * ซึ่งเป็นไฟล์ก็อปปี้กันเป๊ะ 2 ไฟล์ — เป็นจุดซ้ำซ้อนที่สุดของโค้ดเดิม)
 * ------------------------------------------------------------------
 */
import { canonicalHeader } from "../../validators/shared/header-matcher";

/**
 * ผลลัพธ์จากการหา Identity ของ Test Case
 *
 * displayValue:
 * - ค่าที่ใช้แสดงใน Column "Test Script No."
 * - ใช้ Test No. เมื่อมีค่า
 * - ถ้า Test No. ว่าง ให้แสดงเป็นช่องว่าง
 *
 * matchingReference:
 * - ค่า Matching Reference เช่น Transaction ID/ Reconcile ID
 * - ใช้จับคู่กับ Raw Report
 */
export interface ResolvedRecordIdentity {
  readonly displayValue: string;
  readonly matchingReference: string;
}

export class ReconcileRecord {
  /** เลขแถวจริงใน worksheet (ใช้ trace กลับไป highlight/copy แถวต้นฉบับ) */
  readonly rowNumber: number;

  /** key = canonicalHeader(header name), value = ค่าที่ normalize แล้ว */
  private readonly values: Record<string, string>;

  constructor(rowNumber: number, values: Record<string, string>) {
    this.rowNumber = rowNumber;
    this.values = values;
  }

  /** ดึงค่า field ตามชื่อ header (ใช้ canonicalHeader กัน header สะกด/เว้นวรรคต่างกันเล็กน้อย) */
  get(headerName: string): string {
    return this.values[canonicalHeader(headerName)] ?? "";
  }

  /**
   * หา Identity ของ Test Case โดยใช้กติกากลางร่วมกันทุก Report
   *
   * จุดประสงค์:
   * - Test No. ว่างต้องไม่ทำให้ Test Data ทั้งแถวหายจาก Reconcile
   * - Matching Reference ยังใช้จับคู่กับ Raw Report ได้ตามเดิม
   * - ไม่ผูกกติกานี้ไว้ใน Mapping Config เพื่อลด Merge Conflict
   *
   * @param primaryHeader Header หลักสำหรับแสดง Test Case เช่น "Test No."
   * @param fallbackHeader Header สำรองและใช้จับคู่ เช่น
   * "Transaction ID/ Reconcile ID"
   * @param normalizePrimary ตัวแปลงค่าหลักก่อนแสดง เช่น ตัด "-return"
   */
  resolveIdentity(
    primaryHeader: string,
    fallbackHeader: string,
    normalizePrimary?: (value: string) => string,
  ): ResolvedRecordIdentity {
    const primaryValue = this.get(primaryHeader).trim();

    const fallbackValue = this.get(fallbackHeader).trim();

    const normalizedPrimaryValue =
      primaryValue === ""
        ? ""
        : (normalizePrimary?.(primaryValue) ?? primaryValue).trim();

    return {
      /**
       * Test No. ใช้สำหรับแสดงผลเท่านั้น
       * จึงไม่ใช้ Transaction ID แสดงแทนเมื่อ Test No. ว่าง
       */
      displayValue: normalizedPrimaryValue,

      /**
       * Matching Key ยังคงเป็น Transaction ID/ Reconcile ID
       * และไม่ขึ้นกับการมีหรือไม่มี Test No.
       */
      matchingReference: fallbackValue,
    };
  }
}
