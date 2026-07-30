/**
 * ftx-exclusion-rule.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ตรวจสอบว่าข้อมูล Test Data จำนวน 1 แถว
 * เข้าเงื่อนไข Exclusion Rule ของ DS_FTX หรือไม่
 *
 * Exclusion Rule หมายถึง
 * รายการที่ไม่ควรปรากฏอยู่ใน Report DS_FTX
 *
 * หากเข้า Exclusion Rule:
 * - ไม่ต้องค้นหา Matching Key ใน Report
 * - ไม่ต้องเปรียบเทียบ Core Field
 * - สร้างผลลัพธ์เป็น PASS
 * - เพิ่ม Remark เพื่ออธิบายเหตุผล
 *
 * การ Exclude มีทั้งหมด 2 Rule
 *
 * Rule 1:
 * From Currency (CCY) เป็น THB
 *
 * หรือ
 *
 * Rule 2:
 * ต้องผ่านทุกเงื่อนไขต่อไปนี้พร้อมกัน
 *
 * 1. From Currency มีค่า
 * 2. Settled Currency มีค่า
 * 3. From Currency ตรงกับ Settled Currency
 * 4. ลูกค้าเป็น Resident
 * 5. Settled Currency เป็น USD
 * 6. Settled Amount มีค่าและแปลงเป็นตัวเลขได้
 * 7. Settled Amount อยู่ระหว่าง 0 ถึง 50,000 USD
 *    โดยรวมค่า 0 และ 50,000
 *
 * หมายเหตุสำคัญ
 * - Rule 1 และ Rule 2 เชื่อมกันด้วย OR
 * - เงื่อนไขภายใน Rule 2 เชื่อมกันด้วย AND
 * - Currency และ Customer Type ไม่สนตัวพิมพ์เล็ก–ใหญ่
 * - ชื่อ Header ถูกอ่านแบบตรงตัวและไม่รองรับ Alias
 * ------------------------------------------------------------------
 */

import {
  TestDataRow,
} from "./compare-types";

/**
 * ชื่อ Header ที่ใช้ตรวจสอบ From Currency
 *
 * อ่านค่าจาก Test Data ด้วยชื่อ Header แบบตรงตัว
 */
const FROM_CURRENCY_HEADER =
  "From Currency (CCY)";

/**
 * ชื่อ Header ที่ใช้ตรวจสอบ Settled Currency
 */
const SETTLED_CURRENCY_HEADER =
  "Settled Currency (CCY)";

/**
 * ชื่อ Header ที่ใช้ตรวจสอบ Settled Amount
 */
const SETTLED_AMOUNT_HEADER =
  "Settled Amount (CCY)";

/**
 * ชื่อ Header ที่ใช้ตรวจสอบประเภทลูกค้า
 *
 * ค่าที่ใช้ตรวจ Resident ได้แก่
 * - R
 * - RESIDENT
 */
const CUSTOMER_TYPE_HEADER =
  "From Customer Type Description";

/**
 * Base Currency หรือสกุลเงินฐาน
 *
 * ถ้า From Currency เป็น THB
 * จะเข้า Exclusion Rule 1 ทันที
 */
const BASE_CURRENCY =
  "THB";

/**
 * Settled Currency ที่ใช้ตรวจสอบ Rule 2
 *
 * Rule 2 กำหนดให้ Settled Currency เป็น USD
 *
 * เนื่องจาก Rule 2 กำหนดเพิ่มเติมว่า
 * From Currency ต้องตรงกับ Settled Currency
 *
 * ดังนั้นเมื่อเข้า Rule 2
 * ทั้ง From Currency และ Settled Currencyจะเป็น USD
 */
const USD_CURRENCY =
  "USD";

/**
 * วงเงินสูงสุดสำหรับลูกค้า Resident
 * ที่สามารถเข้า Exclusion Rule 2
 *
 * ค่าปัจจุบัน:
 * 50,000 USD
 *
 * เงื่อนไขใน Code ใช้ <=
 * ดังนั้นจำนวน 50,000 จะเข้า Exclusion ด้วย
 */
const RESIDENT_AMOUNT_LIMIT_USD =
  50000;

/**
 * รูปแบบผลการตรวจสอบ Exclusion Rule
 */
export interface FtxExclusionResult {
  /**
   * ผลการตรวจสอบ
   *
   * true
   * = เข้าเงื่อนไข Exclusion
   *
   * false
   * = ไม่เข้าเงื่อนไข Exclusion
   */
  isExcluded: boolean;

  /**
   * คำอธิบายเหตุผลของผลการตรวจสอบ
   *
   * ถ้าเข้า Exclusion:
   * จะเป็นข้อความอธิบาย Rule ที่เข้า
   *
   * ถ้าไม่เข้า Exclusion:
   * จะเป็นข้อความว่าง ""
   */
  remark: string;
}

/**
 * แปลงค่าทั่วไปให้เป็นข้อความมาตรฐาน
 * สำหรับใช้เปรียบเทียบแบบไม่สนตัวพิมพ์เล็ก–ใหญ่
 *
 * การทำงาน
 * 1. ถ้าค่าเป็น null หรือ undefined ให้คืนข้อความว่าง
 * 2. แปลงค่าเป็น string
 * 3. ตัดช่องว่างด้านหน้าและด้านหลัง
 * 4. เปลี่ยนข้อความเป็นตัวพิมพ์ใหญ่
 *
 * ตัวอย่าง:
 *
 * " usd "     → "USD"
 * "Resident"  → "RESIDENT"
 * null        → ""
 *
 * หมายเหตุ
 * ฟังก์ชันนี้ไม่ลบช่องว่างภายในข้อความ
 *
 * @param value ค่าที่ต้องการ Normalize
 * @returns ข้อความตัวพิมพ์ใหญ่ที่ตัดช่องว่างแล้ว
 */
const normalizeText = (
  value: unknown,
): string => {
  // ถ้าไม่มีค่า ให้คืนข้อความว่าง
  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  /**
   * แปลงเป็น string
   * ตัดช่องว่างหน้า–หลัง
   * และเปลี่ยนเป็นตัวพิมพ์ใหญ่
   */
  return String(
    value,
  )
    .trim()
    .toUpperCase();
};

/**
 * แปลงค่า Settled Amount จาก Excel ให้เป็น Number
 *
 * รองรับตัวเลขที่มีเครื่องหมาย comma คั่นหลักพัน
 *
 * ตัวอย่าง:
 *
 * "50,000"    → 50000
 * "12,500.25" → 12500.25
 * 1000        → 1000
 *
 * ถ้าไม่มีค่าหรือแปลงเป็น Number ไม่ได้
 * จะคืนค่า undefined
 *
 * @param value ค่า Amount ที่อ่านจาก Test Data
 *
 * @returns
 * - number เมื่อแปลงสำเร็จ
 * - undefined เมื่อไม่มีค่าหรือแปลงไม่ได้
 */
const normalizeAmount = (
  value: unknown,
): number | undefined => {
  /**
   * ถ้าไม่มีค่า Amount
   * ให้คืน undefined
   */
  if (
    value === null ||
    value === undefined
  ) {
    return undefined;
  }

  /**
   * แปลงค่าเป็นข้อความ
   * ลบ comma ทั้งหมด
   * และตัดช่องว่างหน้า–หลัง
   */
  const text =
    String(
      value,
    )
      .replace(
        /,/g,
        "",
      )
      .trim();

  /**
   * ถ้าหลัง Normalize แล้วเป็นข้อความว่าง
   * ให้คืน undefined
   */
  if (
    text === ""
  ) {
    return undefined;
  }

  /**
   * แปลงข้อความเป็น Number
   */
  const amount =
    Number(
      text,
    );

  /**
   * ถ้าผลลัพธ์เป็น NaN
   * หมายถึงไม่สามารถแปลงเป็นตัวเลขได้
   */
  if (
    Number.isNaN(
      amount,
    )
  ) {
    return undefined;
  }

  return amount;
};

/**
 * ตรวจสอบว่าประเภทลูกค้าเป็น Resident หรือไม่
 *
 * Requirement กำหนดค่า:
 * - R  = Resident
 * - NR = Non-Resident
 *
 * Code ปัจจุบันรองรับค่า Resident 2 รูปแบบ:
 * - R
 * - RESIDENT
 *
 * เนื่องจากใช้ normalizeText()
 * ค่าต่อไปนี้จึงถือว่าเป็น Resident เช่นกัน:
 * - r
 * - resident
 * - " RESIDENT "
 *
 * ค่าอื่นทั้งหมดจะคืน false
 *
 * @param value ค่า From Customer Type Description
 * @returns true เมื่อเป็น Resident
 */
const isResidentCustomer = (
  value: unknown,
): boolean => {
  /**
   * Normalize ค่า Customer Type
   * ก่อนนำมาเปรียบเทียบ
   */
  const customerType =
    normalizeText(
      value,
    );

  return (
    customerType === "R" ||
    customerType === "RESIDENT"
  );
};

/**
 * ตรวจสอบ Exclusion Rule ของ DS_FTX
 *
 * ลำดับการตรวจสอบ
 *
 * 1. ตรวจ Rule 1: From Currency เป็น THB
 * 2. ถ้าไม่เข้า Rule 1 ให้ตรวจ Rule 2
 * 3. Rule 2 ต้องผ่านเงื่อนไขทุกข้อพร้อมกัน
 * 4. ถ้าไม่เข้า Rule ใดเลย ให้คืน isExcluded: false
 *
 * ถ้าข้อมูลตรงกับหลาย Rule
 * ระบบจะใช้ผลของ Rule แรกที่พบ
 *
 * หมายเหตุเกี่ยวกับ Flow ทั้งระบบ
 * compare-validator.ts จะตรวจ Matching Key ว่างก่อน
 * แล้วจึงเรียกฟังก์ชันนี้
 *
 * ดังนั้น เมื่อใช้งานผ่าน Compare Flow ปกติ
 * แถวที่ Matching Key ว่างจะ FAIL ก่อน
 * และจะไม่ได้เข้ามาตรวจ Exclusion Rule
 *
 * @param testDataRow ข้อมูล Test Data จำนวน 1 แถว
 *
 * @returns ผลการตรวจสอบ Exclusion และ Remark
 */
export const evaluateFtxExclusion = (
  testDataRow: TestDataRow,
): FtxExclusionResult => {
  /**
   * อ่าน From Currency จาก Header
   * "From Currency (CCY)"
   *
   * normalizeText() จะเปลี่ยนค่าเป็นตัวพิมพ์ใหญ่
   */
  const fromCurrency =
    normalizeText(
      testDataRow[
        FROM_CURRENCY_HEADER
      ],
    );

  /**
   * อ่าน Settled Currency จาก Header
   * "Settled Currency (CCY)"
   */
  const settlementCurrency =
    normalizeText(
      testDataRow[
        SETTLED_CURRENCY_HEADER
      ],
    );

  /**
   * อ่านประเภทลูกค้าจาก Header
   * "From Customer Type Description"
   *
   * ยังไม่ Normalize ตรงนี้
   * เพราะจะส่งไปให้ isResidentCustomer() จัดการ
   */
  const customerType =
    testDataRow[
      CUSTOMER_TYPE_HEADER
    ];

  /**
   * อ่าน Settled Amount จาก Header
   * "Settled Amount (CCY)"
   *
   * ค่าจะถูกแปลงเป็น number
   * หรือ undefined ถ้าแปลงไม่ได้
   */
  const settledAmount =
    normalizeAmount(
      testDataRow[
        SETTLED_AMOUNT_HEADER
      ],
    );

  /**
   * Exclusion Rule 1
   *
   * เงื่อนไข:
   * From Currency เป็น THB
   *
   * Rule นี้ตรวจเพียง From Currency
   * ไม่ได้กำหนดเงื่อนไข Customer Type,
   * Settled Currency หรือ Settled Amount เพิ่มเติม
   */
  if (
    fromCurrency ===
    BASE_CURRENCY
  ) {
    return {
      isExcluded: true,

      remark:
        "Excluded: From Currency (CCY) เป็น THB จึงไม่อยู่ในขอบเขต DS_FTX",
    };
  }

  /**
   * Exclusion Rule 2
   *
   * ต้องผ่านทุกเงื่อนไขพร้อมกัน:
   *
   * 1. From Currency มีค่า
   * 2. Settled Currency มีค่า
   * 3. From Currency ตรงกับ Settled Currency
   * 4. ลูกค้าเป็น Resident
   * 5. Settled Currency เป็น USD
   * 6. Settled Amount มีค่าและแปลงเป็นตัวเลขได้
   * 7. Settled Amount มากกว่าหรือเท่ากับ 0
   * 8. Settled Amount น้อยกว่าหรือเท่ากับ 50,000
   *
   * เนื่องจากเงื่อนไขทั้งหมดใช้ &&
   * จึงต้องเป็น true ครบทุกข้อ
   */
  if (
    fromCurrency !== "" &&
    settlementCurrency !== "" &&
    fromCurrency ===
      settlementCurrency &&
    isResidentCustomer(
      customerType,
    ) &&
    settlementCurrency ===
      USD_CURRENCY &&
    settledAmount !==
      undefined &&
    settledAmount >= 0 &&
    settledAmount <=
      RESIDENT_AMOUNT_LIMIT_USD
  ) {
    return {
      isExcluded: true,

      remark:
        "Excluded: From Currency ตรงกับ Settled Currency, ลูกค้าเป็น Resident และ Settled Amount ไม่เกิน 50,000 USD",
    };
  }

  /**
   * ไม่เข้า Exclusion Rule 1 หรือ Rule 2
   *
   * รายการนี้จะถูกส่งกลับไปให้ compare-validator.ts
   * เพื่อดำเนินการต่อ เช่น
   * - ค้นหา Matching Key ใน Report
   * - ตรวจ Matching Key ซ้ำ
   * - เปรียบเทียบ Core Field
   */
  return {
    isExcluded: false,
    remark: "",
  };
};