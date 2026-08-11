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

    /**
     * Header ของ DF_FXU Report อยู่ที่แถวที่ 1
     *
     * Worksheet ที่พบในไฟล์จริง:
     * DF_FXU Transaction
     */
    headerRowNumber: 1,

    /**
     * ชื่อ Header สำรอง
     *
     * ตอนนี้ Header ในไฟล์ DF_FXU ตัวอย่าง
     * ตรงกับชื่อใน Requirement จึงยังไม่มี Alias
     */
    aliases: {},

    /**
     * Header ของ DF_FXU Report
     * แบ่งกลุ่มตามหน้าที่ของ Script 3
     */
    requiredHeaders: {

      /**
       * Matching Key หลัก
       *
       * Test Data:
       * Transaction ID/ Reconcile ID
       *
       * DF_FXU Report:
       * Arrangement Number
       */
      matchingKey: [
        "Arrangement Number",
      ],

      /**
       * Core Field
       *
       * Data Set Date:
       * - ตรวจเทียบกับ Txn Date
       * - ใช้เป็นข้อมูลสนับสนุนการจับคู่
       *
       * USD Equivalent Amount:
       * - ตรวจเทียบกับ Settled Amount (CCY)
       * - ใช้ตรวจ Threshold 1,000,000 USD
       *
       * Arrangement Type:
       * - Requirement กำหนดค่า 018101
       *
       * Leg Type:
       * - รองรับ 182001 และ 182002
       *
       * Leg Type Name:
       * - ต้องสัมพันธ์กับ Leg Type
       */
      core: [
        "Data Set Date",
        "USD Equivalent Amount",
        "Arrangement Type",
        "Leg Type",
        "Leg Type Name",
      ],

      /**
       * Requirement ยังไม่ได้กำหนด Customer Field
       * ที่ใช้ตัดสิน PASS หรือ FAIL สำหรับ DF_FXU
       */
      customer: [],

      /**
       * ไม่มี Conditional Field ฝั่ง DF_FXU Report
       * ที่ระบุชื่อ Header ชัดเจนใน Requirement
       *
       * Payment Intermediary และ Return/Reversal
       * จะประเมินจากข้อมูลใน Test Data ภายใน Business Rule
       */
      conditions: [],

      /**
       * Reference Field
       *
       * เป็นข้อมูลประกอบที่แสดงใน Report
       * แต่ Requirement ยังไม่ได้กำหนดวิธีนำค่ามา
       * เปรียบเทียบเพื่อใช้ตัดสิน PASS หรือ FAIL
       *
       * Fi Arrangement Type Name:
       * - Script 2 ยังคงตรวจว่า Header ต้องมี
       * - Script 3 ยังไม่นำค่ามาตัดสินจนกว่าจะมี Expected Value
       */
      reference: [
        "DEPT CODE",
        "Cust Code",
        "CMF CODE",
        "Cust Name",
        "Fi Arrangement Type Name",
        "Currency ID",
        "Currency ID Name",
        "Original Amount",
      ],

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

    /**
     * Header ของ DF_FXM Report
     * อยู่ที่แถวที่ 1
     *
     * DF_FXM ใช้โครงสร้างข้อมูลเดียวกับ DF_FXU
     * แต่ใช้กับยอดตั้งแต่ 1,000,000 USD ขึ้นไป
     */
    headerRowNumber: 1,

    /**
     * ตอนนี้ Header ในไฟล์ Report
     * ตรงกับชื่อใน Requirement
     * จึงยังไม่มีชื่อ Headerสำรอง
     */
    aliases: {},

    /**
     * Header ของ DF_FXM Report
     * แบ่งตามหน้าที่ที่ใช้ใน Script 3
     */
    requiredHeaders: {

      /**
       * Matching Key หลัก
       *
       * Test Data:
       * Transaction ID/ Reconcile ID
       *
       * DF_FXM Report:
       * Arrangement Number
       */
      matchingKey: [
        "Arrangement Number",
      ],

      /**
       * Core Field
       *
       * Data Set Date:
       * - ตรวจเทียบกับ Txn Date
       * - ใช้สนับสนุน Fallback Matching
       *
       * USD Equivalent Amount:
       * - ตรวจเทียบกับ Settled Amount (CCY)
       * - ต้องมียอดตั้งแต่ 1,000,000 USD ขึ้นไป
       *
       * Arrangement Type:
       * - Requirement กำหนดค่า 018101
       *
       * Leg Type:
       * - รองรับ 182001 และ 182002
       *
       * Leg Type Name:
       * - ต้องสัมพันธ์กับ Leg Type
       */
      core: [
        "Data Set Date",
        "USD Equivalent Amount",
        "Arrangement Type",
        "Leg Type",
        "Leg Type Name",
      ],

      /**
       * Requirement ยังไม่ได้กำหนด Customer Field
       * ที่ใช้ตัดสิน PASS หรือ FAIL สำหรับ DF_FXM
       */
      customer: [],

      /**
       * ยังไม่มี Conditional Field ฝั่ง DF_FXM Report
       *
       * Payment Intermediary เช่น NIUM
       * จะยังไม่นำมาตัดสินในขั้นตอนนี้
       */
      conditions: [],

      /**
       * Reference Field
       *
       * เป็นข้อมูลที่ต้องมีอยู่ใน Report
       * แต่ยังไม่ใช้ตัดสิน PASS หรือ FAIL
       * จนกว่าจะมี Expected Value ชัดเจน
       */
      reference: [
        "DEPT CODE",
        "Cust Code",
        "CMF CODE",
        "Cust Name",
        "Fi Arrangement Type Name",
        "Currency ID",
        "Currency ID Name",
        "Original Amount",
      ],

    },

  },

  // บอก TypeScript ว่า Config ชุดนี้เป็นค่าคงที่
  // เพื่อให้ชื่อ Report และค่าภายใน Config มี Type ที่ชัดเจน
} as const;
