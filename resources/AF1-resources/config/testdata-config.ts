/**
 * Config สำหรับตรวจ Header และข้อมูลใน Test Data แยกตาม Report
 */

/** DS_PTX รองรับ Fee Group จำนวน n ชุด */
export const FEE_TYPE_COUNT = 2;

/** DS_LTX รองรับ Fee Group จำนวน n ชุด */
export const LTX_FEE_TYPE_COUNT = 2;

/**
 * คืนชื่อ Header ของ Fee Amount
 * - Fee 1 ใช้ "Fee Amount Type 1"
 * - Fee 2 เป็นต้นไปใช้ "Fee Amount 2", "Fee Amount 3", ...
 */
export const getFeeAmountHeader = (
  feeNumber: number,
): string => {
  if (feeNumber < 1) {
    throw new Error(
      `Invalid Fee number: ${feeNumber}`,
    );
  }

  return feeNumber === 1
    ? "Fee Amount Type 1"
    : `Fee Amount ${feeNumber}`;
};

/**
 * สร้าง Header ของ Fee Group ตามกฎของ Report
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
 */
const createFeeHeaders = (
  count: number,
  includeFeeChargeType = false,
): string[] => {
  return Array.from(
    { length: count },
    (_, index) => {
      const feeNumber =
        index + 1;

      return [
        `Fee Type ${feeNumber}`,

        ...(
          includeFeeChargeType
            ? [
                `Fee Charge Type ${feeNumber}`,
              ]
            : []
        ),

        `Fee Charge Account No. Type ${feeNumber}`,

        getFeeAmountHeader(
          feeNumber,
        ),
      ];
    },
  ).flat();
};

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
       * LTX ใช้ Logic เดิม:
       * - ตรวจ Fee หลัก 3 ช่อง
       * - Fee Group ว่างทั้งชุดให้ข้าม
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
       * PTX ตรวจ Fee หลัก 4 ช่อง:
       * - Fee Type
       * - Fee Charge Type
       * - Fee Charge Account No.
       * - Fee Amount
       */
      feeGroup: [
        ...createFeeHeaders(
          FEE_TYPE_COUNT,
          true,
        ),
      ],
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

export type TestDataReportCode =
  keyof typeof TESTDATA_CONFIG;