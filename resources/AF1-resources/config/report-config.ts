// ======================================================
// Report Configuration
//
// ไฟล์นี้เก็บ Header ที่ไฟล์ Report แต่ละประเภทต้องมี
// Script 2 ใช้ข้อมูลชุดนี้ตรวจว่าไฟล์ที่ดาวน์โหลดมามี Header ครบหรือไม่
//
// แต่ละ Report กำหนดข้อมูล 3 ส่วน:
// - headerRowNumber: เลขแถวที่เก็บ Header
// - aliases: ชื่ออื่นที่ยอมรับแทนชื่อ Header หลัก
// - requiredHeaders: Header ที่ต้องพบใน Report
// ======================================================

export const REPORT_CONFIG = {

  // Requirement ของ Report DS_LTX
  DS_LTX: {

    // Header อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // Header ที่ต้องพบในไฟล์ DS_LTX
    requiredHeaders: [
      "Reference Transaction Number",
      "Transaction Date",
      "Currency Id",
      "FI Arrangement Number",
      "Loan Deposit Transaction Type",
      "Transaction Amount",
      "Outflow Transaction Purpose",
      "Payment Method",
      "From Transaction Type",
      "To Transaction Type",
      "Inflow Transaction Purpose",
      "Installment Number",
      "Beneficiary or Sender Name",
      "Country Id of Beneficiary or Sender",
      "Relationship with Beneficiary or Sender",
      "Approval Document Number",
      "Cust Code",
      "CMF CODE",
      "Cust Name",
      "Data Set Date",
      "Data Submission Period",
    ],

  },

  // Requirement ของ Report DS_PTX
  DS_PTX: {

    // Header อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ชื่อ Header สำรองที่ระบบยอมรับแทนชื่อหลัก
    aliases: {

      // ยอมรับทั้ง "Currency Id" และ "Currency ID"
      "Currency Id": [
        "Currency Id",
        "Currency ID",
      ],

      // ยอมรับทั้ง "CMF CODE" และ "CMF Code"
      "CMF CODE": [
        "CMF CODE",
        "CMF Code",
      ],

    },

    // Header ที่ต้องพบในไฟล์ DS_PTX
    requiredHeaders: [
      "Reference Transaction Number",
      "Receive Payment Transaction Date",
      "Currency Id",
      "Transaction Amount in Foreign Currency",
      "Cust Code",
      "CMF CODE",
      "Cust Name",
      "Involved Party Id",
      "Involved Party Name",
      "Country Id of Involved Party",
    ],

  },

  // Requirement ของ Report DS_FTX
  DS_FTX: {

    // Header อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // Header ที่ต้องพบในไฟล์ DS_FTX
    requiredHeaders: [
      "Ref. TX No.",
      "Buy Currency Id",
      "Sell Currency Id",
      "Transaction Date",
    ],

  },

  // Report ด้านล่างยังไม่ได้เพิ่ม Required Header
  // จึงเก็บ Array ว่างไว้เพื่อให้โครงสร้าง Config รองรับชื่อ Report ก่อน
  DS_FTU: {

    headerRowNumber: 1,

    aliases: {},

    requiredHeaders: [
      "Arr Number",
      "Data Set Date",
      "Leg Type",
      "Inflow Transaction Purpose",
      "Outflow Transaction Purpose",
      "Country Id of Beneficiary Involved Party",
      "Currency Id",
      "Foreign Currency Amount",
    ],

  },

  DF_FXU: {

    headerRowNumber: 1,

    aliases: {},

    requiredHeaders: [],

  },

  DF_OLB: {

    headerRowNumber: 1,

    aliases: {},

    requiredHeaders: [],

  },

  DF_FXM: {

    headerRowNumber: 1,

    aliases: {},

    requiredHeaders: [],

  },

} as const;

/**
 * สร้าง Type ของชื่อ Report จาก Key ใน REPORT_CONFIG
 *
 * ตัวอย่างค่าที่ TypeScript ยอมรับ:
 * "DS_LTX", "DS_PTX", "DS_FTX"
 *
 * หากเพิ่มชื่อ Report ใหม่ใน REPORT_CONFIG
 * Type นี้จะรองรับชื่อใหม่ให้อัตโนมัติ
 */
export type ReportName =
  keyof typeof REPORT_CONFIG;

/**
 * ชื่อ Type กลางที่โมดูล Reconcile และ Summary ใช้งาน
 *
 * เป็น Type เดียวกับ ReportName เพื่อให้ Code เดิมของ
 * DS_PTX และ DS_FTX ยังใช้งานได้เหมือนเดิม
 */
export type ReportCode =
  ReportName;
