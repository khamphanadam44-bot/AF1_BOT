// ======================================================
// Config สำหรับ Test Data
//
// ใช้เก็บ Header Requirement ของ Test Data
// โดยแยก Configuration ตาม Report
//
// แม้ DS_PTX และ DS_FTX จะใช้ไฟล์ Test Dataเดียวกัน
// แต่แต่ละ Report ใช้ Header ในการตรวจสอบไม่เหมือนกัน
// ======================================================

/**
 * จำนวนชุด Fee ที่ DS_PTX รองรับ
 *
 * ตอนนี้กำหนดไว้ทั้งหมด 2 ชุด:
 * - Fee ชุดที่ 1
 * - Fee ชุดที่ 2
 */
export const FEE_TYPE_COUNT = 2;

/**
 * จำนวนชุด Fee ที่ DS_LTX รองรับ
 *
 * Test Data จริงของ DS_LTX มี Fee Type 1-5
 * จึงต้องแยกจำนวนออกจาก DS_PTX ที่รองรับ 2 ชุด
 */
export const LTX_FEE_TYPE_COUNT = 5;

/**
 * สร้างรายการ Header ของ Fee
 * ตามจำนวนชุด Fee ที่กำหนด
 *
 * ตัวอย่างเมื่อ count = 2:
 *
 * Fee ชุดที่ 1
 * - Fee Type 1
 * - Fee Charge Account No. Type 1
 * - Fee Amount Type 1
 *
 * Fee ชุดที่ 2
 * - Fee Type 2
 * - Fee Charge Account No. Type 2
 * - Fee Amount Type 2
 */

/**
 * คืนชื่อ Header ของ Fee Amount
 *
 * Requirement:
 * - Fee ชุดที่ 1 ใช้ "Fee Amount Type 1"
 * - Fee ชุดที่ 2 เป็นต้นไปใช้ "Fee Amount 2",
 *   "Fee Amount 3", ...
 */
export const getFeeAmountHeader = (
  feeNumber: number,
): string => {
  if (
    feeNumber < 1
  ) {
    throw new Error(
      `Invalid Fee number: ${feeNumber}`,
    );
  }

  return feeNumber === 1
    ? "Fee Amount Type 1"
    : `Fee Amount ${feeNumber}`;
};


function createFeeHeaders(
  count: number,
): string[] {
  return Array.from(
    { length: count },
    (_, index) => {
      const feeNumber =
        index + 1;

      return [
        `Fee Type ${feeNumber}`,
        `Fee Charge Type ${feeNumber}`,
        `Fee Charge Account No. Type ${feeNumber}`,
        getFeeAmountHeader(
          feeNumber,
        ),
      ];
    },
  ).flat();
}

/**
 * Config สำหรับ Test Data
 *
 * แยก Header ที่ต้องตรวจสอบ
 * ตามชื่อ Report
 */
export const TESTDATA_CONFIG = {
  // ====================================================
  // DS_LTX
  // ====================================================
  DS_LTX: {
    /**
     * Header ของไฟล์ Test Data อยู่ที่แถว 5
     */
    headerRowNumber: 5,

    requiredHeaders: {
      /**
       * ใช้จับคู่ Test Data กับ DS_LTX
       */
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

      /**
       * ข้อมูลหลักที่ DS_LTX ใช้สร้าง Expected Case
       */
      core: [
        "Test No.",
        "Txn Date",
        "From Account (A/C Client/Sender)",
        "From Currency (CCY)",
        "From Debit Amount",
        "From Transfer Amount",
        "From THB Equivalent Transfer Amount",
        "From BOT Purpose code",
        "To Credit Amount (To Amount)",
        "To THB Equivalent Transfer Amount",
        "Payment rail",
        "From Account Type",
        "To Account Type (Beneficiary)",
      ],

      /**
       * ข้อมูลลูกค้าที่ใช้ใน Conditional Field ของ DS_LTX
       */
      customer: [
        "To CIF Name (Beneficiary)",
        "To Region (Bene Country)",
        "From Customer (Resident/Non Resident)",
        "From CIF No. (Client/Sender)",
        "From CIF Name (Client/Sender)",
      ],

      conditional: [],

      reference: [],

      /**
       * DS_LTX ใช้โครงสร้าง Fee เดียวกับ Test Data กลาง
       */
      feeGroup: [
        ...createFeeHeaders(
          LTX_FEE_TYPE_COUNT,
        ),
      ],
    },
  },

  // ====================================================
  // DS_PTX
  // ====================================================
  DS_PTX: {
    /**
     * Header ของไฟล์ Test Data
     * อยู่ที่แถว 5
     */
    headerRowNumber: 5,

    requiredHeaders: {
      /**
       * Header สำหรับใช้จับคู่ข้อมูล
       * ระหว่าง Test Data กับ Report
       */
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

      /**
       * Header ข้อมูลหลักของ DS_PTX
       *
       * Header กลุ่มนี้จะถูกตรวจแบบ Normal Field:
       * - มีข้อมูล = ผ่าน
       * - ไม่มีข้อมูล = ไม่ผ่าน
       *
       * Fee Header จะไม่อยู่ในกลุ่มนี้
       * เพราะ Fee ต้องตรวจร่วมกันเป็นชุด
       */
      core: [
        "Txn Date",
        "From Account (A/C Client/Sender)",
        "From Currency (CCY)",
        "From Debit Amount",
        "From Transfer Amount",
        "From THB Equivalent Transfer Amount",
        "From BOT Purpose code",
        "To Credit Amount (To Amount)",
        "To THB Equivalent Transfer Amount",
        "Payment rail",
        "From Account Type",
        "To Account Type (Beneficiary)",
      ],

      /**
       * Header ข้อมูลลูกค้าของ DS_PTX
       */
      customer: [
        "To CIF Name (Beneficiary)",
        "To Region (Bene Country)",
        "From Customer (Resident/Non Resident)",
        "From CIF No. (Client/Sender)",
        "From CIF Name (Client/Sender)",
      ],

      /**
       * ข้อมูลที่ตรวจตามเงื่อนไข (Conditional Field)
       *
       * ตอนนี้ยังไม่มี Header
       * ที่กำหนดในกลุ่มนี้
       */
      conditional: [],

      /**
       * ข้อมูลสำหรับอ้างอิง (Reference Field)
       *
       * เตรียมไว้รองรับ Header เพิ่มเติม
       * ในอนาคต
       */
      reference: [],

      /**
       * Fee Group ของ DS_PTX
       *
       * แยกออกจาก core เพื่อไม่ให้ Fee
       * ถูกตรวจซ้ำแบบ Normal Field
       *
       * หลักการตรวจ:
       * - ว่างทั้งชุด = ผ่าน
       * - มีข้อมูลครบทั้งชุด = ผ่าน
       * - มีข้อมูลเพียงบางช่อง = ไม่ผ่าน
       */
      feeGroup: [
        ...createFeeHeaders(
          FEE_TYPE_COUNT,
        ),
      ],
    },
  },

  // ====================================================
  // DS_FTX
  // ====================================================
  DS_FTX: {
    /**
     * DS_FTX ใช้ไฟล์ Test Data เดียวกับ DS_PTX
     * และ Header อยู่ที่แถว 5 เหมือนกัน
     */
    headerRowNumber: 5,

    requiredHeaders: {
      /**
       * Header สำหรับใช้จับคู่ข้อมูล
       * ระหว่าง Test Data กับ Report
       */
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

      /**
       * Header ข้อมูลหลักของ DS_FTX
       */
      core: [
        "From Currency (CCY)",
        "To Currency (CCY)",
        "Settled Currency (CCY)",
        "Settled Amount (CCY)",
        "Txn Date",
        "From Customer Type Code",
        "Test No.",
        "From Customer Type Description",
      ],

      /**
       * ตอนนี้ DS_FTX
       * ยังไม่มี Customer Field
       */
      customer: [],

      /**
       * ตอนนี้ DS_FTX
       * ยังไม่มี Conditional Field
       */
      conditional: [],

      /**
       * ตอนนี้ DS_FTX
       * ยังไม่มี Reference Field
       */
      reference: [],

      /**
       * DS_FTX ไม่มี Fee Group
       *
       * แต่ต้องประกาศ Property นี้ไว้
       * เพื่อให้โครงสร้าง Config ของทุก Report
       * เหมือนกัน และไม่เกิด TypeScript Error
       */
      feeGroup: [],
    },
  },

  // ====================================================
  // DS_FTU
  // ====================================================
  DS_FTU: {
    headerRowNumber: 5,

    requiredHeaders: {
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

      core: [
        "Test No.",
        "Txn Date",
        "From Currency (CCY)",
        "To Currency (CCY)",
        "Settled Currency (CCY)",
        "Settled Amount (CCY)",
        "From BOT Purpose code",
      ],

      customer: [],

      conditional: [],

      reference: [],

      feeGroup: [],
    },
  },
} as const;

/**
 * ชื่อ Report ที่รองรับ
 * ใน Test Data Config
 *
 * ผลลัพธ์ของ Type นี้คือ:
 * "DS_LTX" | "DS_PTX" | "DS_FTX" | "DS_FTU"
 */
export type TestDataReportCode =
  keyof typeof TESTDATA_CONFIG;
