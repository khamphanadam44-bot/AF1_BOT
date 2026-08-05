/**
 * Config สำหรับตรวจ Header และข้อมูลใน Test Data
 * โดยแยกการตั้งค่าตาม Report
 */

/**
 * คืนชื่อ Header ของ Fee Amount ตามลำดับ Fee Group
 *
 * Fee Group 1:
 * - Fee Amount Type 1
 *
 * Fee Group 2 เป็นต้นไป:
 * - Fee Amount 2
 * - Fee Amount 3
 * - ...
 */
export const getFeeAmountHeader = (
  feeNumber: number,
): string => {
  if (
    !Number.isInteger(feeNumber) ||
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

/**
 * สร้างรายการ Header ของ Fee Group
 *
 * DS_LTX:
 * - Fee Type
 * - Fee Charge Account No.
 * - Fee Amount
 *
 * DS_PTX:
 * - Fee Type
 * - Fee Charge Type
 * - Fee Charge Account No.
 * - Fee Amount
 *
 * จำนวน Fee Group จะถูกส่งเข้ามาจากผลการตรวจ Header
 * ของไฟล์ Test Data จริง จึงไม่ต้องกำหนดจำนวนแบบ Hard code
 */
const createFeeHeaders = (
  feeTypeCount: number,
  includeFeeChargeType = false,
): string[] => {
  if (
    !Number.isInteger(feeTypeCount) ||
    feeTypeCount < 0
  ) {
    throw new Error(
      `Invalid Fee Type count: ${feeTypeCount}`,
    );
  }

  return Array.from(
    {
      length: feeTypeCount,
    },
    (_, index) => {
      const feeNumber = index + 1;

      const headers = [
        `Fee Type ${feeNumber}`,
      ];

      if (includeFeeChargeType) {
        headers.push(
          `Fee Charge Type ${feeNumber}`,
        );
      }

      headers.push(
        `Fee Charge Account No. Type ${feeNumber}`,
        getFeeAmountHeader(feeNumber),
      );

      return headers;
    },
  ).flat();
};

/**
 * Config สำหรับตรวจ Test Data ของแต่ละ Report
 */
export const TESTDATA_CONFIG = {
  // ====================================================
  // DS_LTX
  // ====================================================
  DS_LTX: {
    headerRowNumber: 5,

    requiredHeaders: {
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

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
       * DS_LTX ใช้ Fee Header หลัก 3 ช่องต่อหนึ่งกลุ่ม:
       * - Fee Type
       * - Fee Charge Account No.
       * - Fee Amount
       *
       * จำนวนกลุ่มจะรับมาจาก Header ใน Test Data จริง
       */
      feeGroup: (
        feeTypeCount: number,
      ): string[] => {
        return createFeeHeaders(
          feeTypeCount,
        );
      },
    },
  },

  // ====================================================
  // DS_PTX
  // ====================================================
  DS_PTX: {
    headerRowNumber: 5,

    requiredHeaders: {
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

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
       * DS_PTX ใช้ Fee Header หลัก 4 ช่องต่อหนึ่งกลุ่ม:
       * - Fee Type
       * - Fee Charge Type
       * - Fee Charge Account No.
       * - Fee Amount
       *
       * จำนวนกลุ่มจะรับมาจาก Header ใน Test Data จริง
       */
      feeGroup: (
        feeTypeCount: number,
      ): string[] => {
        return createFeeHeaders(
          feeTypeCount,
          true,
        );
      },
    },
  },

  // ====================================================
  // DS_FTX
  // ====================================================
  DS_FTX: {
    headerRowNumber: 5,

    requiredHeaders: {
      matchingKey: [
        "Transaction ID/ Reconcile ID",
      ],

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

      customer: [],

      conditional: [],

      reference: [],

      /**
       * DS_FTX ไม่มีการตรวจ Fee Group
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

      /**
       * DS_FTU ไม่มีการตรวจ Fee Group
       */
      feeGroup: [],
    },
  },
} as const;

/**
 * Report Code ที่รองรับใน Test Data Config
 */
export type TestDataReportCode =
  keyof typeof TESTDATA_CONFIG;