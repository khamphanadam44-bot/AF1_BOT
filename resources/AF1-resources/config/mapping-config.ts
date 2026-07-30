// ======================================================
// Report Mapping Configuration
//
// ไฟล์นี้กำหนดว่า Report แต่ละประเภทใช้ Header อะไรในการ Mapping
// และแบ่ง Header ตามหน้าที่ เพื่อให้ระบบเลือกตรวจข้อมูลได้ถูกกลุ่ม
//
// ความหมายของแต่ละกลุ่ม:
// - matchingKey: กุญแจที่ใช้จับคู่แถวใน Test Data กับ Report
// - core: ข้อมูลหลักที่ต้องตรวจและใช้ตัดสิน PASS / FAIL
// - customer: ข้อมูลลูกค้าที่ตรวจเฉพาะเมื่อเข้าเงื่อนไข
// - conditions: ข้อมูลเงื่อนไขอื่นที่ไม่ได้อยู่ในกลุ่มลูกค้า
// - reference: ข้อมูลประกอบที่ไม่ใช้ตัดสิน PASS / FAIL
// ======================================================



export const REPORT_CONFIG_mapping = {

  // ======================================================
  // Report: DS_LTX
  // ======================================================
  DS_LTX: {

    // Header ของ Report อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // รายการ Header ของ DS_LTX แบ่งตามหน้าที่ในการตรวจสอบ
    requiredHeaders: {

      // ใช้ Reference Transaction Number จับคู่ Test Data กับ Report
      matchingKey: [
        "Reference Transaction Number",
      ],

      // ข้อมูลหลักที่ต้องตรวจทุก Matching Key ที่พบ
      core: [
        "Transaction Date",
        "Currency Id",
        "FI Arrangement Number",
        "Loan Deposit Transaction Type",
        "Transaction Amount",
        "Outflow Transaction Purpose",
        "Payment Method",
        "From Transaction Type",
        "To Transaction Type",
      ],

      // ข้อมูลที่ตรวจเฉพาะเมื่อรายการเข้าเงื่อนไขที่เกี่ยวข้อง
      conditions: [
        "Inflow Transaction Purpose",
        "Installment Number",
        "Beneficiary or Sender Name",
        "Country Id of Beneficiary or Sender",
        "Relationship with Beneficiary or Sender",
        "Approval Document Number",
        "Cust Code",
        "CMF CODE",
        "Cust Name",
      ],

      // ตอนนี้ยังไม่ได้แยก Header ไว้ในกลุ่มข้อมูลลูกค้า
      customer: [],

      // ข้อมูลอ้างอิงที่แสดงในผลลัพธ์ แต่ไม่ใช้ตัดสิน PASS / FAIL
      reference: [
        "Data Set Date",
        "Data Submission Period",
      ],

    },

  },
// ======================================================
// Report: DS_PTX
// ======================================================
DS_PTX: {

  // Header ของ Report อยู่ที่แถวที่ 1
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

  requiredHeaders: {

    /**
     * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
     *
     * ใช้ Reference Transaction Number ค้นหาว่า
     * รายการใน Test Data ตรงกับรายการใดใน Report
     */
    matchingKey: [
      "Reference Transaction Number",
    ],

    /**
     * ข้อมูลหลักที่ต้องตรวจ (Core Fields)
     *
     * ตรวจทุก Matching Key ที่พบ
     * และนำผลมาใช้ตัดสิน PASS / FAIL
     */
    core: [

      "Receive Payment Transaction Date",

      "Currency Id",

      "Payment Method",

      "Receive Payment Transaction Type",

      "Receive Payment Item Type",

      "Transaction Amount in Foreign Currency",

    ],

    /**
     * ข้อมูลลูกค้าที่ตรวจตามเงื่อนไข
     *
     * ระบบตรวจเฉพาะ Field ที่จำเป็นสำหรับรายการนั้น
     */
    customer: [

      "Cust Code",

      "CMF CODE",

      "Cust Name",

      "Involved Party Id",

      "Involved Party Name",

      "Country Id of Involved Party",

      "Receive Payment Item Description",

    ],

    /**
     * กลุ่มเงื่อนไขเพิ่มเติม
     *
     * ตอนนี้ DS_PTX ยังไม่มี Header ในกลุ่มนี้
     */
    conditions: [],

    /**
     * ข้อมูลสำหรับใช้อ้างอิง
     *
     * แสดงในไฟล์ผลลัพธ์เพื่อช่วยตรวจสอบที่มาของรายการ
     * แต่ไม่นำค่ามาใช้ตัดสิน PASS / FAIL
     */
    reference: [

      "Data Set Date",

      "Dept Code",

      "System Id",

    ],

  },

},
  // ======================================================
  // Report: DS_FTX
  // ======================================================
  DS_FTX: {

    // Header ของ Report อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // ยังไม่ได้เพิ่ม Header Mapping ของ DS_FTX ใน Config ชุดนี้
    // เมื่อ Requirement พร้อมแล้ว ให้เพิ่ม Header ลงในกลุ่มที่ตรงกับหน้าที่
    requiredHeaders: {

      matchingKey: [],

      core: [],

      customer: [],

      conditions: [],

      reference: [],

    },

  },

  // ======================================================
  // Report: DS_FTU
  // ======================================================
  DS_FTU: {

    // Header ของ Report อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    requiredHeaders: {

      matchingKey: [
        "Arr Number",
      ],

      core: [
        "Data Set Date",
        "Leg Type",
        "Country Id of Beneficiary Involved Party",
        "Currency Id",
        "Foreign Currency Amount",
      ],

      customer: [],

      conditions: [
        "Inflow Transaction Purpose",
        "Outflow Transaction Purpose",
      ],

      reference: [],

    },

  },

  // ======================================================
  // Report: DF_FXU
  // ======================================================
  DF_FXU: {

    // Header ของ Report อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // ยังไม่ได้เพิ่ม Header Mapping ของ DF_FXU ใน Config ชุดนี้
    requiredHeaders: {

      matchingKey: [],

      core: [],

      customer: [],

      conditions: [],

      reference: [],

    },

  },

  // ======================================================
  // Report: DF_OLB
  // ======================================================
  DF_OLB: {

    // Header ของ Report อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // ยังไม่ได้เพิ่ม Header Mapping ของ DF_OLB ใน Config ชุดนี้
    requiredHeaders: {

      matchingKey: [],

      core: [],

      customer: [],

      conditions: [],

      reference: [],

    },

  },

  // ======================================================
  // Report: DF_FXM
  // ======================================================
  DF_FXM: {

    // Header ของ Report อยู่ที่แถวที่ 1
    headerRowNumber: 1,

    // ตอนนี้ยังไม่มีชื่อ Header สำรอง
    aliases: {},

    // ยังไม่ได้เพิ่ม Header Mapping ของ DF_FXM ใน Config ชุดนี้
    requiredHeaders: {

      matchingKey: [],

      core: [],

      customer: [],

      conditions: [],

      reference: [],

    },

  },

// บอก TypeScript ว่า Config ชุดนี้เป็นค่าคงที่
// เพื่อให้ชื่อ Report และค่าภายใน Config มี Type ที่ชัดเจน
} as const;
