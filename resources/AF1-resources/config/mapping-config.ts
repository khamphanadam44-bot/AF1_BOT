// ======================================================
// Report Mapping Configuration
//
// ไฟล์นี้ใช้บอกระบบว่า "Report แต่ละแบบ" (เช่น DS_LTX, DS_PTX)
// มีคอลัมน์ (Header) อะไรบ้างในไฟล์ Excel/CSV ที่จะนำมาตรวจสอบ
// และคอลัมน์แต่ละอันมีหน้าที่อะไร เวลาระบบเทียบข้อมูลจริง
// (Test Data) กับข้อมูลใน Report จะได้รู้ว่าต้องเช็คตรงไหนบ้าง
//
// พูดง่ายๆ คือ: Report 1 ไฟล์ = 1 ก้อนข้อมูลข้างล่างนี้
// ในแต่ละก้อนจะบอกว่า "Header ชื่ออะไร" ควรถูกจัดเข้ากลุ่มไหน
//
// ความหมายของแต่ละกลุ่ม (คิดง่ายๆ ว่าเป็น "ป้ายกำกับ" ของ Header):
//
// - matchingKey : ใช้เป็น "กุญแจ" จับคู่แถวข้อมูล
//                 เช่น ถ้า Test Data มีเลข Transaction เดียวกับใน Report
//                 แปลว่าเป็นรายการเดียวกัน ระบบจะเอามาเทียบกันต่อ
//
// - core        : ข้อมูลสำคัญที่ "ต้องตรวจทุกครั้ง" ถ้าค่าผิด
//                 ผลลัพธ์รายการนั้นจะออกมาเป็น FAIL ทันที
//
// - customer    : ข้อมูลของลูกค้า (ชื่อ, รหัสลูกค้า ฯลฯ)
//                 ที่จะตรวจสอบก็ต่อเมื่อรายการนั้นเข้าเงื่อนไขบางอย่าง
//                 ไม่ได้ตรวจทุกรายการ
//
// - conditions  : ข้อมูลอื่นๆ ที่ตรวจแบบมีเงื่อนไขเหมือนกัน
//                 แต่ไม่ใช่ข้อมูลลูกค้า (เช่น วัตถุประสงค์การทำธุรกรรม)
//
// - reference   : ข้อมูลที่เก็บไว้ "โชว์ให้ดูเฉยๆ" เพื่อช่วยอ้างอิง
//                 ไม่ได้ใช้ตัดสินว่า PASS หรือ FAIL
// ======================================================

export const REPORT_CONFIG_mapping = {

  // ------------------------------------------------------
  // Report: DS_LTX
  // ------------------------------------------------------
  DS_LTX: {
    headerRowNumber: 1,   // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header
    aliases: {},          // ชื่อ Header สำรอง (เผื่อบางไฟล์สะกดไม่เหมือนกัน) - ตอนนี้ยังไม่มี

    requiredHeaders: {
      // ใช้คอลัมน์ "Reference Transaction Number" เป็นตัวจับคู่ว่า
      // แถวไหนใน Test Data ตรงกับแถวไหนใน Report
      matchingKey: [
        "Reference Transaction Number",
      ],

      // Header กลุ่มนี้ต้องตรวจทุกรายการที่จับคู่กันได้
      // ถ้าค่าไม่ตรงกัน = FAIL
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

      // Header กลุ่มนี้จะถูกตรวจ "เฉพาะเมื่อ" รายการนั้นเข้าเงื่อนไขที่เกี่ยวข้อง
      // (ไม่ใช่ทุกรายการต้องมีค่าพวกนี้)
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

      // Report นี้ยังไม่มีการแยก Header เข้ากลุ่ม "ข้อมูลลูกค้า" โดยเฉพาะ
      customer: [],

      // แสดงไว้ในผลลัพธ์เพื่อให้คนอ่านเข้าใจง่ายขึ้นเฉยๆ
      // ไม่ได้ถูกนำไปใช้ตัดสิน PASS / FAIL
      reference: [
        "Data Set Date",
        "Data Submission Period",
      ],
    },
  },

  // ------------------------------------------------------
  // Report: DS_PTX
  // ------------------------------------------------------
  DS_PTX: {
    headerRowNumber: 1,   // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header

    // บางไฟล์อาจตั้งชื่อคอลัมน์ไม่เหมือนกันเป๊ะๆ (เช่น พิมพ์เล็ก/ใหญ่ต่างกัน)
    // ตรงนี้บอกระบบว่า ถ้าเจอชื่อไหนในลิสต์ ให้ถือว่าเป็นคอลัมน์เดียวกัน
    aliases: {
      // ถ้าเจอ "Currency Id" หรือ "Currency ID" ให้ถือว่าเป็นคอลัมน์เดียวกัน
      "Currency Id": ["Currency Id", "Currency ID"],
      // ถ้าเจอ "CMF CODE" หรือ "CMF Code" ให้ถือว่าเป็นคอลัมน์เดียวกัน
      "CMF CODE": ["CMF CODE", "CMF Code"],
    },

    requiredHeaders: {
      // ใช้คอลัมน์ "Reference Transaction Number" หาว่ารายการใน Test Data
      // ตรงกับรายการไหนใน Report
      matchingKey: [
        "Reference Transaction Number",
      ],

      // ต้องตรวจทุกรายการที่จับคู่กันได้ และใช้ตัดสิน PASS / FAIL
      core: [
        "Receive Payment Transaction Date",
        "Currency Id",
        "Payment Method",
        "Receive Payment Transaction Type",
        "Receive Payment Item Type",
        "Transaction Amount in Foreign Currency",
      ],

      // ข้อมูลของลูกค้า จะตรวจเฉพาะ Field ที่จำเป็นสำหรับรายการนั้นๆ
      // (ไม่ใช่ทุกรายการต้องมีครบทุกค่า)
      customer: [
        "Cust Code",
        "CMF CODE",
        "Cust Name",
        "Involved Party Id",
        "Involved Party Name",
        "Country Id of Involved Party",
        "Receive Payment Item Description",
      ],

      // Report นี้ยังไม่มี Header ในกลุ่มเงื่อนไขอื่นๆ
      conditions: [],

      // ใช้เก็บไว้เพื่อช่วยให้คนตรวจสอบเข้าใจที่มาของรายการง่ายขึ้น
      // ไม่ได้ใช้ตัดสิน PASS / FAIL
      reference: [
        "Data Set Date",
        "Dept Code",
        "System Id",
      ],
    },
  },

  // ------------------------------------------------------
  // Report: DS_FTX
  // ------------------------------------------------------
  DS_FTX: {
    headerRowNumber: 1,   // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header
    aliases: {},          // ยังไม่มีชื่อ Header สำรอง

    requiredHeaders: {
      // ใช้จับคู่รายการ:
      // ฝั่ง Test Data ใช้ชื่อ "Transaction ID / Reconcile ID"
      // ฝั่ง Report (DS_FTX) ใช้ชื่อ "Ref. TX No."
      // สองอันนี้คือค่าเดียวกัน แค่คนละชื่อคอลัมน์
      matchingKey: [
        "Ref. TX No.",
      ],

      // ข้อมูลหลักที่ต้องตรวจทุกรายการ
      // (ตรวจตามเงื่อนไขที่กำหนดไว้ใน DS_FTX_COMPARE_RULES)
      core: [
        "Buy Currency Id",
        "Sell Currency Id",
        "Transaction Date",
      ],

      // ยังไม่มี Header ในกลุ่มข้อมูลลูกค้า
      customer: [],

      // ยังไม่มี Header ในกลุ่มเงื่อนไขอื่นๆ
      conditions: [],

      // ยังไม่มี Header ในกลุ่มข้อมูลอ้างอิง
      reference: [],
    },
  },

  // ------------------------------------------------------
  // Report: DS_FTU
  // ------------------------------------------------------
  DS_FTU: {
    headerRowNumber: 1,   // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header
    aliases: {},          // ยังไม่มีชื่อ Header สำรอง

    requiredHeaders: {
      // ใช้คอลัมน์นี้จับคู่รายการระหว่าง Test Data กับ Report
      matchingKey: [
        "Arr Number",
      ],

      // ต้องตรวจทุกรายการ และใช้ตัดสิน PASS / FAIL
      core: [
        "Data Set Date",
        "Leg Type",
        "Country Id of Beneficiary Involved Party",
        "Currency Id",
        "Foreign Currency Amount",
      ],

      // ยังไม่มี Header ในกลุ่มข้อมูลลูกค้า
      customer: [],

      // ตรวจเฉพาะเมื่อรายการเข้าเงื่อนไขที่เกี่ยวข้องเท่านั้น
      conditions: [
        "Inflow Transaction Purpose",
        "Outflow Transaction Purpose",
      ],

      // ยังไม่มี Header ในกลุ่มข้อมูลอ้างอิง
      reference: [],
    },
  },

  // ------------------------------------------------------
  // Report: DF_FXU
  // ------------------------------------------------------
  DF_FXU: {
    // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header
    // (ไฟล์จริงที่เจอ อยู่ใน Worksheet ชื่อ "DF_FXU Transaction")
    headerRowNumber: 1,

    // ชื่อคอลัมน์ในไฟล์ตัวอย่างตรงกับที่ Requirement กำหนดไว้อยู่แล้ว
    // เลยยังไม่ต้องตั้งชื่อสำรอง
    aliases: {},

    requiredHeaders: {
      // ใช้จับคู่รายการ:
      // ฝั่ง Test Data ใช้ชื่อ "Transaction ID / Reconcile ID"
      // ฝั่ง Report (DF_FXU) ใช้ชื่อ "Arrangement Number"
      // สองอันนี้คือค่าเดียวกัน แค่คนละชื่อคอลัมน์
      matchingKey: [
        "Arrangement Number",
      ],

      // ข้อมูลหลักที่ต้องตรวจทุกรายการ:
      // - Data Set Date         : เทียบกับ Txn Date ฝั่ง Test Data
      //                           ใช้ช่วยยืนยันว่าจับคู่รายการถูกต้อง
      // - USD Equivalent Amount : เทียบกับ Settled Amount (CCY)
      //                           และใช้เช็คว่ายอดถึงเกณฑ์ 1,000,000 USD หรือไม่
      // - Arrangement Type      : ตาม Requirement ต้องเป็นค่า 018101
      // - Leg Type               : รับได้ 2 ค่าคือ 182001 กับ 182002
      // - Leg Type Name          : ค่าต้องสอดคล้องกับ Leg Type ด้านบน
      core: [
        "Data Set Date",
        "USD Equivalent Amount",
        "Arrangement Type",
        "Leg Type",
        "Leg Type Name",
      ],

      // Requirement ยังไม่ได้ระบุว่าต้องตรวจข้อมูลลูกค้าตัวไหนบ้าง
      customer: [],

      // Report ฝั่งนี้ไม่มีชื่อคอลัมน์ที่ระบุชัดเจนสำหรับกลุ่มเงื่อนไข
      // เรื่อง Payment Intermediary และ Return/Reversal ระบบจะไปเช็ค
      // จากข้อมูลใน Test Data โดยตรงแทน (ผ่าน Business Rule อื่น)
      conditions: [],

      // ข้อมูลที่ต้องมีอยู่ใน Report เพื่อโชว์ประกอบผลลัพธ์
      // แต่ยังไม่ได้ใช้ตัดสิน PASS / FAIL เพราะยังไม่มีค่าที่ถูกต้อง (Expected Value)
      // มากำหนดไว้ชัดเจน
      // (หมายเหตุ: Fi Arrangement Type Name ตอนนี้แค่เช็คว่ามีคอลัมน์นี้อยู่จริง
      // ยังไม่ได้เอาค่าไปเทียบอะไร)
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

  // ------------------------------------------------------
  // Report: DF_OLB
  // ------------------------------------------------------
  DF_OLB: {
    headerRowNumber: 1,   // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header
    aliases: {},          // ยังไม่มีชื่อ Header สำรอง

    // Report นี้ยังไม่ได้ใส่ Header ครบทุกกลุ่ม รอข้อมูลเพิ่มเติม
    requiredHeaders: {
      // ใช้ทั้ง 3 คอลัมน์นี้ร่วมกันเพื่อจับคู่รายการ
      matchingKey: [
        "FI Arrangement Number",
        "Arrangement Contract Date",
        "THB Outstanding Amount",
      ],

      // ต้องตรวจทุกรายการ และใช้ตัดสิน PASS / FAIL
      core: [
        "Cust Code",
        "Cust Name",
      ],

      customer: [],
      conditions: [],
      reference: [],
    },
  },

  // ------------------------------------------------------
  // Report: DF_FXM
  // (ใช้โครงสร้างข้อมูลเหมือน DF_FXU เลย ต่างกันแค่ Report นี้ใช้กับ
  //  รายการที่มียอดตั้งแต่ 1,000,000 USD ขึ้นไป)
  // ------------------------------------------------------
  DF_FXM: {
    headerRowNumber: 1,   // แถวที่ 1 ของไฟล์ คือแถวที่เป็นชื่อ Header

    // ชื่อคอลัมน์ในไฟล์ Report ตรงกับที่ Requirement กำหนดไว้อยู่แล้ว
    // เลยยังไม่ต้องตั้งชื่อสำรอง
    aliases: {},

    requiredHeaders: {
      // ใช้จับคู่รายการ:
      // ฝั่ง Test Data ใช้ชื่อ "Transaction ID / Reconcile ID"
      // ฝั่ง Report (DF_FXM) ใช้ชื่อ "FI Arrangement Number"
      matchingKey: [
        "FI Arrangement Number",
      ],

      // ข้อมูลหลักที่ต้องตรวจทุกรายการ:
      // - Data Set Date         : เทียบกับ Txn Date ฝั่ง Test Data
      //                           ใช้ช่วยกรณีจับคู่รายการหลักไม่ได้ (Fallback)
      // - USD Equivalent Amount : เทียบกับ Settled Amount (CCY)
      //                           ต้องมียอดตั้งแต่ 1,000,000 USD ขึ้นไป
      // - Arrangement Type      : ตาม Requirement ต้องเป็นค่า 018101
      // - Leg Type               : รับได้ 2 ค่าคือ 182001 กับ 182002
      // - Leg Type Name          : ค่าต้องสอดคล้องกับ Leg Type ด้านบน
      core: [
        "Data Set Date",
        "USD Equivalent Amount",
        "Arrangement Type",
        "Leg Type",
        "Leg Type Name",
      ],

      // Requirement ยังไม่ได้ระบุว่าต้องตรวจข้อมูลลูกค้าตัวไหนบ้าง
      customer: [],

      // ยังไม่มีเงื่อนไขที่ระบุชัดเจนสำหรับ Report นี้
      // (เช่นกรณี Payment Intermediary อย่าง NIUM ก็ยังไม่ถูกนำมาตัดสินตอนนี้)
      conditions: [],

      // ข้อมูลที่ต้องมีอยู่ใน Report แต่ยังไม่ได้ใช้ตัดสิน PASS / FAIL
      // เพราะยังไม่มีค่าที่ถูกต้อง (Expected Value) มากำหนดไว้ชัดเจน
      reference: [
        "DEPT CODE",
        "Cust Code",
        "CMF CODE",
        "Cust Name",
        "Arrangement Type Name",
        "Currency Code",
        "Currency Code Name",
        "Original Amount",
      ],
    },
  },

  // หมายเหตุสำหรับคนเขียนโค้ด: "as const" ท้ายไฟล์นี้เป็นคำสั่งของ TypeScript
  // บอกว่าอ็อบเจกต์นี้ห้ามแก้ไขค่าและ "ล็อกชนิดข้อมูล" ไว้แบบตายตัว
  // เพื่อให้เวลาดึงชื่อ Report (เช่น DS_LTX) หรือ field ต่างๆ ไปใช้ที่อื่น
  // จะมีการเช็ค type ให้อัตโนมัติ ป้องกันการพิมพ์ชื่อผิด
} as const;