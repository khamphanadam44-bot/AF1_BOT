/**
 * header-matcher.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้ตรวจสอบและจับคู่ชื่อ Header ระหว่าง
 * 1. Header ที่พบจริงใน Downloaded Report
 * 2. Header ที่พบจริงใน Test Data
 * 3. Header ที่กำหนดไว้ใน Requirement หรือ Config
 *
 * ความสามารถหลัก
 * - ปรับรูปแบบข้อความ Header ก่อนเปรียบเทียบ
 * - ตรวจสอบว่า Header สองชื่อ Match กันหรือไม่
 * - รองรับชื่อ Header ที่เขียนได้หลายแบบด้วย Alias
 * - สร้าง Alias ของ Fee Header ตามจำนวน Fee Type
 * - ค้นหา Header จริงจากรายการ Header ทั้งหมด
 * - ค้นหาตำแหน่ง Index ของ Header
 * - วิเคราะห์สาเหตุที่ Header ไม่ Match
 *
 * คำศัพท์
 * - Actual Header   = Header ที่พบจริงในไฟล์ Excel
 * - Required Header = Header ที่กำหนดไว้ใน Requirement
 * - Alias           = ชื่ออื่นที่อนุญาตให้ใช้แทนชื่อหลักได้
 * - Match           = ตรงกันตามเงื่อนไข
 * - Missing         = ไม่พบ Header
 * - Symbol          = เครื่องหมายพิเศษ เช่น /, -, _, (, )
 * ------------------------------------------------------------------
 */

/**
 * ปรับรูปแบบข้อความ Header ให้อยู่ในรูปแบบมาตรฐานเบื้องต้น
 *
 * การทำงาน
 * 1. ถ้าค่าเป็น null หรือ undefined ให้ใช้ข้อความว่างแทน
 * 2. แปลงค่าให้เป็น string
 * 3. เปลี่ยนตัวอักษรทั้งหมดเป็นตัวพิมพ์เล็ก
 * 4. เปลี่ยนการขึ้นบรรทัดใหม่ให้เป็นช่องว่าง
 * 5. รวมช่องว่างที่ติดกันหลายช่องให้เหลือช่องเดียว
 * 6. ลบช่องว่างด้านหน้าและด้านหลัง
 *
 * ตัวอย่าง
 *
 * " Transaction   Date "
 * จะกลายเป็น
 * "transaction date"
 *
 * "Transaction\nDate"
 * จะกลายเป็น
 * "transaction date"
 *
 * @param value ค่าที่ต้องการปรับรูปแบบ
 * @returns ข้อความที่ผ่านการ Normalize แล้ว
 */
export const normalizeHeader = (
  value: unknown,
): string => {
  return String(value ?? "")
    .toLowerCase()
    .replace(/\r?\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
};

/**
 * ปรับ Header ให้อยู่ในรูปแบบมาตรฐานสำหรับการเปรียบเทียบจริง
 *
 * ฟังก์ชันนี้จะ
 * 1. เรียก normalizeHeader() ก่อน
 * 2. ลบช่องว่างทั้งหมดออกจากข้อความ
 *
 * ตัวอย่าง
 *
 * "Transaction Date"
 * จะกลายเป็น
 * "transactiondate"
 *
 * "TransactionDate"
 * จะกลายเป็น
 * "transactiondate"
 *
 * ดังนั้น Header สองตัวอย่างด้านบนจะถือว่า Match กัน
 *
 * หมายเหตุ
 * ฟังก์ชันนี้ลบเฉพาะช่องว่าง แต่ไม่ได้ลบเครื่องหมายพิเศษ
 *
 * ตัวอย่าง
 * "Transaction ID/ Reconcile ID"
 * จะกลายเป็น
 * "transactionid/reconcileid"
 *
 * เครื่องหมาย / จะยังคงอยู่
 *
 * @param value ค่าที่ต้องการแปลงเป็น Canonical Header
 * @returns Header ที่เป็นตัวพิมพ์เล็กและไม่มีช่องว่าง
 */
export const canonicalHeader = (
  value: unknown,
): string => {
  return normalizeHeader(value)
    .replace(/\s+/g, "");
};

/**
 * ตรวจสอบว่า Actual Header และ Required Header ตรงกันหรือไม่
 *
 * การทำงาน
 * 1. แปลง Actual Header ด้วย canonicalHeader()
 * 2. แปลง Required Header ด้วย canonicalHeader()
 * 3. ถ้าฝั่งใดฝั่งหนึ่งเป็นค่าว่าง ให้ถือว่าไม่ Match
 * 4. ถ้าค่าหลังปรับรูปแบบเหมือนกัน ให้ถือว่า Match
 *
 * สิ่งที่ฟังก์ชันนี้ไม่สนใจ
 * - ตัวพิมพ์เล็กและตัวพิมพ์ใหญ่
 * - ช่องว่าง
 * - การขึ้นบรรทัดใหม่
 *
 * สิ่งที่ฟังก์ชันนี้ยังตรวจสอบ
 * - ตัวอักษร
 * - ตัวเลข
 * - เครื่องหมายพิเศษ เช่น /, -, _, (, )
 *
 * ตัวอย่างที่ Match
 * Actual:   "Currency ID"
 * Required: "currency id"
 *
 * ตัวอย่างที่ไม่ Match
 * Actual:   "Transaction ID Reconcile ID"
 * Required: "Transaction ID/ Reconcile ID"
 *
 * เพราะ Actual Header ไม่มีเครื่องหมาย /
 *
 * @param actualHeader Header ที่พบจริงในไฟล์
 * @param requiredHeader Header ที่กำหนดไว้ใน Requirement
 * @returns true เมื่อ Match และ false เมื่อไม่ Match
 */
export const isHeaderMatch = (
  actualHeader: unknown,
  requiredHeader: unknown,
): boolean => {
  // ปรับ Actual Header ให้อยู่ในรูปแบบมาตรฐาน
  const actual = canonicalHeader(
    actualHeader,
  );

  // ปรับ Required Header ให้อยู่ในรูปแบบมาตรฐาน
  const required = canonicalHeader(
    requiredHeader,
  );

  /**
   * ถ้า Actual Header หรือ Required Header เป็นค่าว่าง
   * จะไม่สามารถนำมา Match กันได้
   */
  if (!actual || !required) {
    return false;
  }

  // ส่งผลลัพธ์ว่าข้อความทั้งสองฝั่งตรงกันหรือไม่
  return actual === required;
};

/**
 * Alias กลางที่สามารถใช้ร่วมกันได้กับทุก Report
 *
 * Key ทางด้านซ้าย
 * คือชื่อ Header หลักที่ใช้ใน Requirement หรือ Config
 *
 * Array ทางด้านขวา
 * คือรายชื่อ Header ทั้งหมดที่อนุญาตให้ใช้แทนชื่อหลักได้
 *
 * ตัวอย่าง
 *
 * Required Header:
 * "Txn Date"
 *
 * Actual Header ที่ยอมรับ:
 * - "Txn Date"
 * - "Transaction Date"
 */
export const COMMON_HEADER_ALIASES: Record<
  string,
  string[]
> = {
  /**
   * รองรับทั้งชื่อย่อ Txn Date
   * และชื่อเต็ม Transaction Date
   */
  "Txn Date": [
    "Txn Date",
    "Transaction Date",
  ],

  /**
   * รองรับกรณีมีหรือไม่มีช่องว่างก่อนเครื่องหมาย /
   */
  "Transaction ID/ Reconcile ID": [
    "Transaction ID/ Reconcile ID",
    "Transaction ID / Reconcile ID",
  ],

  /**
   * รองรับการเขียนคำว่า Id และ ID
   *
   * แม้ isHeaderMatch() จะเปลี่ยนเป็นตัวพิมพ์เล็กอยู่แล้ว
   * แต่การระบุไว้ช่วยให้เห็นชื่อที่ระบบอนุญาตอย่างชัดเจน
   */
  "Currency Id": [
    "Currency Id",
    "Currency ID",
  ],

  /**
   * รองรับการเขียน CODE และ Code
   */
  "CMF CODE": [
    "CMF CODE",
    "CMF Code",
  ],
};

/**
 * สร้าง Alias สำหรับ Header ที่อยู่ใน Fee Group
 *
 * เนื่องจาก Fee Header มีหมายเลขต่อท้ายแบบ Dynamic เช่น
 * - Fee Type 1
 * - Fee Type 2
 * - Fee Type 3
 *
 * จึงสร้าง Alias ตามจำนวน Fee Type ที่ได้รับมา
 * แทนการเขียน Alias ทีละรายการด้วยตนเอง
 *
 * Header ที่สร้างในแต่ละลำดับ
 * 1. Fee Type {ลำดับ}
 * 2. Fee Charge Account No. Type {ลำดับ}
 * 3. Fee Amount Type {ลำดับ}
 *
 * ตัวอย่าง เมื่อ feeTypeCount = 2
 * ระบบจะสร้าง Alias สำหรับ Fee Group ลำดับที่ 1 และ 2
 *
 * @param feeTypeCount จำนวน Fee Type สูงสุดที่ต้องการสร้าง Alias
 * @returns Object ที่รวม Alias ของ Fee Header ทั้งหมด
 */
export const buildFeeHeaderAliases = (
  feeTypeCount: number,
): Record<string, string[]> => {
  /**
   * Object สำหรับเก็บ Alias ที่สร้างขึ้น
   *
   * เริ่มต้นเป็น Object ว่าง
   * และจะเพิ่ม Alias เข้าไปภายใน Loop
   */
  const aliases: Record<string, string[]> = {};

  /**
   * เริ่มสร้าง Fee Header ตั้งแต่หมายเลข 1
   * และทำซ้ำจนถึงหมายเลข feeTypeCount
   *
   * ตัวอย่าง feeTypeCount = 3
   * ค่า index จะเป็น 1, 2 และ 3
   */
  for (
    let index = 1;
    index <= feeTypeCount;
    index += 1
  ) {
    /**
     * สร้าง Alias ของ Fee Type
     *
     * ตัวอย่าง index = 1
     * Key   = Fee Type 1
     * Alias = Fee Type 1
     */
    aliases[`Fee Type ${index}`] = [
      `Fee Type ${index}`,
    ];

    /**
     * สร้าง Alias ของ Fee Charge Account Number
     *
     * รองรับการเขียนได้ 3 แบบ
     * - Fee Charge Account No. Type 1
     * - Fee Charge Account No Type 1
     * - Fee Charge Account Number Type 1
     *
     * ช่วยรองรับกรณี
     * - มีหรือไม่มีเครื่องหมายจุดหลัง No
     * - เขียนคำว่า Number แบบเต็ม
     */
    aliases[
      `Fee Charge Account No. Type ${index}`
    ] = [
      `Fee Charge Account No. Type ${index}`,
      `Fee Charge Account No Type ${index}`,
      `Fee Charge Account Number Type ${index}`,
    ];

    /**
     * สร้าง Alias ของ Fee Amount
     *
     * รองรับการเขียนได้ 2 แบบ
     * - Fee Amount Type 1
     * - Fee Amount 1
     *
     * ทำให้รองรับ Header ที่ตั้งชื่อแตกต่างกัน
     * ระหว่าง Test Data หรือ Report แต่ละรูปแบบ
     */
    const feeAmountAliases = [
      `Fee Amount Type ${index}`,
      `Fee Amount ${index}`,
    ];

    /**
     * ประกาศ Alias ทั้งสอง Key
     *
     * ทำให้ Match ได้ทั้งกรณีที่ Config ใช้:
     * - Fee Amount Type N
     * - Fee Amount N
     *
     * โดยไม่ขึ้นอยู่กับว่าฝั่งใดเป็น Expected Header
     */
    aliases[
      `Fee Amount Type ${index}`
    ] = feeAmountAliases;

    aliases[
      `Fee Amount ${index}`
    ] = feeAmountAliases;
  }

  // ส่ง Alias ของ Fee Header ทั้งหมดกลับไป
  return aliases;
};

/**
 * รวม Alias ทุกประเภทเข้าด้วยกัน
 *
 * Alias ที่ถูกนำมารวมประกอบด้วย
 * 1. COMMON_HEADER_ALIASES
 *    Alias กลางที่ใช้ร่วมกันทุก Report
 *
 * 2. buildFeeHeaderAliases()
 *    Alias ของ Fee Header ที่สร้างตามจำนวน Fee Type
 *
 * 3. customAliases
 *    Alias เฉพาะของแต่ละ Report
 *
 * ลำดับการเขียนมีความสำคัญ
 * เพราะ Object ที่ถูก Spread ทีหลังจะทับค่าของ Key เดิม
 *
 * ลำดับความสำคัญจากน้อยไปมาก
 * 1. Common Alias
 * 2. Fee Header Alias
 * 3. Custom Alias
 *
 * ดังนั้น ถ้า customAliases มีชื่อ Key ซ้ำกับ Common Alias
 * ระบบจะใช้ค่าจาก customAliases
 *
 * @param feeTypeCount จำนวน Fee Type สูงสุด
 * @param customAliases Alias เฉพาะของแต่ละ Report
 * @returns Object ที่รวม Alias ทั้งหมดแล้ว
 */
export const createHeaderAliases = (
  feeTypeCount: number,
  customAliases: Record<string, string[]> = {},
): Record<string, string[]> => {
  return {
    // Alias กลางสำหรับทุก Report
    ...COMMON_HEADER_ALIASES,

    // Alias ของ Fee Header ตามจำนวน Fee Type
    ...buildFeeHeaderAliases(feeTypeCount),

    // Alias เฉพาะ Report มีลำดับความสำคัญสูงสุด
    ...customAliases,
  };
};

/**
 * ค้นหาชื่อ Actual Header ที่ Match กับ Required Header
 *
 * การทำงาน
 * 1. ตรวจสอบว่า Required Header มี Alias หรือไม่
 * 2. ถ้ามี Alias ให้ใช้รายชื่อ Alias ทั้งหมดในการค้นหา
 * 3. ถ้าไม่มี Alias ให้ใช้ Required Header โดยตรง
 * 4. วนค้นหา Actual Header ที่ Match
 * 5. ส่งชื่อ Actual Header ตัวแรกที่พบกลับไป
 *
 * หมายเหตุ
 * - find() จะหยุดค้นหาทันทีเมื่อพบ Header ตัวแรก
 * - ถ้าไม่พบ จะคืนค่า undefined
 *
 * @param actualHeaders รายการ Header ที่พบจริงในไฟล์
 * @param requiredHeader Header ที่ต้องการค้นหา
 * @param aliases Alias ที่ใช้ช่วยในการจับคู่ Header
 * @returns Actual Header ที่พบ หรือ undefined ถ้าไม่พบ
 */
export const findMatchingHeader = (
  actualHeaders: string[],
  requiredHeader: string,
  aliases: Record<string, string[]> = {},
): string | undefined => {
  /**
   * ตรวจสอบว่า Required Header มี Alias หรือไม่
   *
   * ถ้ามี:
   * ใช้ Array ของ Alias ที่กำหนดไว้
   *
   * ถ้าไม่มี:
   * สร้าง Array ที่มี Required Header เพียงรายการเดียว
   */
  const possibleRequiredHeaders =
    aliases[requiredHeader] ?? [requiredHeader];

  /**
   * ค้นหา Actual Header ตัวแรกที่ Match
   *
   * find()
   * ใช้วนค้นหา Actual Header แต่ละตัว
   *
   * some()
   * ใช้ตรวจสอบว่า Actual Header นั้น Match
   * กับ Possible Required Header อย่างน้อย 1 ชื่อหรือไม่
   */
  return actualHeaders.find(
    (actualHeader) =>
      possibleRequiredHeaders.some(
        (possibleRequiredHeader) =>
          isHeaderMatch(
            actualHeader,
            possibleRequiredHeader,
          ),
      ),
  );
};

/**
 * ค้นหาตำแหน่ง Array Index ของ Actual Header
 * ที่ Match กับ Required Header
 *
 * การทำงานคล้ายกับ findMatchingHeader()
 * แต่ฟังก์ชันนี้คืนค่าเป็นตำแหน่ง Index
 * ไม่ได้คืนชื่อ Header
 *
 * ตัวอย่าง
 *
 * actualHeaders:
 * [
 *   "Transaction Date", // Index 0
 *   "Currency Id",      // Index 1
 *   "Payment Method",   // Index 2
 * ]
 *
 * ถ้าค้นหา "Currency Id"
 * ผลลัพธ์จะเป็น 1
 *
 * หมายเหตุ
 * - Array เริ่มนับจาก Index 0
 * - ถ้าไม่พบ Header จะคืนค่า -1
 * - ถ้านำ Index ไปใช้กับ ExcelJS Column ต้องบวก 1
 *
 * ตัวอย่าง
 * Array Index 0 = Excel Column 1 หรือ Column A
 * Array Index 1 = Excel Column 2 หรือ Column B
 *
 * @param actualHeaders รายการ Header ที่พบจริงในไฟล์
 * @param requiredHeader Header ที่ต้องการค้นหา
 * @param aliases Alias ที่ใช้ช่วยในการจับคู่
 * @returns Array Index ที่พบ หรือ -1 เมื่อไม่พบ
 */
export const findMatchingHeaderIndex = (
  actualHeaders: string[],
  requiredHeader: string,
  aliases: Record<string, string[]> = {},
): number => {
  /**
   * ดึง Alias ของ Required Header
   *
   * ถ้าไม่มี Alias จะใช้ Required Header โดยตรง
   */
  const possibleRequiredHeaders =
    aliases[requiredHeader] ?? [requiredHeader];

  /**
   * ค้นหา Index ของ Actual Header ตัวแรกที่ Match
   *
   * findIndex() จะคืนค่า
   * - 0 ขึ้นไป เมื่อพบ Header
   * - -1 เมื่อไม่พบ Header
   */
  return actualHeaders.findIndex(
    (actualHeader) =>
      possibleRequiredHeaders.some(
        (possibleRequiredHeader) =>
          isHeaderMatch(
            actualHeader,
            possibleRequiredHeader,
          ),
      ),
  );
};

/**
 * ประเภทของเหตุผลในการ Match Header
 *
 * ใช้ระบุว่า Header Match หรือไม่ Match เพราะอะไร
 *
 * ความหมายของแต่ละสถานะ
 *
 * MATCH
 * = Actual Header และ Required Header ตรงกัน
 *
 * EMPTY_ACTUAL
 * = Actual Header ไม่มีข้อมูล
 *
 * EMPTY_REQUIRED
 * = Required Header ไม่มีข้อมูล
 *
 * SYMBOL_NOT_MATCH
 * = เครื่องหมายพิเศษไม่ตรงกัน
 *
 * TEXT_NOT_MATCH
 * = ตัวอักษรหรือคำใน Header ไม่ตรงกัน
 *
 * NO_SIMILAR_HEADER
 * = ไม่พบ Header ที่มีข้อความใกล้เคียงกับ Requirement
 */
export type HeaderMatchReason =
  | "MATCH"
  | "EMPTY_ACTUAL"
  | "EMPTY_REQUIRED"
  | "SYMBOL_NOT_MATCH"
  | "TEXT_NOT_MATCH"
  | "NO_SIMILAR_HEADER";

/**
 * ดึงเฉพาะเครื่องหมายพิเศษออกจาก Header
 *
 * ฟังก์ชันจะลบ
 * - ตัวอักษรภาษาอังกฤษ
 * - ตัวอักษรภาษาไทย
 * - ตัวเลข
 *
 * และเก็บเฉพาะเครื่องหมายพิเศษไว้ เช่น
 * - /
 * - -
 * - _
 * - (
 * - )
 *
 * ตัวอย่าง
 *
 * "Transaction ID/ Reconcile ID"
 * ผลลัพธ์คือ "/"
 *
 * "From Account (A/C Client/Sender)"
 * ผลลัพธ์คือ "(/)"
 *
 * @param value Header ที่ต้องการดึงเครื่องหมาย
 * @returns ข้อความที่มีเฉพาะเครื่องหมายพิเศษ
 */
const extractSymbols = (
  value: unknown,
): string => {
  return canonicalHeader(value)
    .replace(/[a-z0-9ก-๙]/gi, "");
};

/**
 * ลบเครื่องหมายพิเศษออกจาก Header
 *
 * หลังจากทำงานแล้วจะเหลือเฉพาะ
 * - ตัวอักษรภาษาอังกฤษ
 * - ตัวอักษรภาษาไทย
 * - ตัวเลข
 *
 * ใช้สำหรับค้นหา Header ที่มีคำเหมือนกัน
 * แต่ใช้เครื่องหมายพิเศษต่างกัน
 *
 * ตัวอย่าง
 *
 * "Transaction ID/ Reconcile ID"
 * และ
 * "Transaction ID Reconcile ID"
 *
 * เมื่อลบ Symbol แล้ว ทั้งสองค่าจะกลายเป็น
 * "transactionidreconcileid"
 *
 * @param value Header ที่ต้องการลบเครื่องหมาย
 * @returns Header ที่เหลือเฉพาะตัวอักษรและตัวเลข
 */
const removeSymbols = (
  value: unknown,
): string => {
  return canonicalHeader(value)
    .replace(/[^a-z0-9ก-๙]/gi, "");
};

/**
 * ค้นหา Actual Header ที่มีข้อความใกล้เคียงกับ Required Header
 *
 * ฟังก์ชันนี้ใช้ในกรณีที่ Header ไม่ Match แบบสมบูรณ์
 * แต่ต้องการหา Header ที่น่าจะเป็น Header เดียวกัน
 *
 * ตัวอย่าง
 *
 * Required:
 * "Transaction ID/ Reconcile ID"
 *
 * Actual:
 * "Transaction ID Reconcile ID"
 *
 * Header ทั้งสองไม่ Match กันแบบปกติ เพราะ Actual ไม่มี /
 * แต่เมื่อลบเครื่องหมายออกแล้ว ข้อความจะเหมือนกัน
 * ฟังก์ชันนี้จึงสามารถหา Actual Header ดังกล่าวเจอ
 *
 * หมายเหตุ
 * ฟังก์ชันนี้คืน Actual Header ตัวแรกที่มีข้อความใกล้เคียง
 * ถ้าไม่พบจะคืนค่า undefined
 *
 * @param actualHeaders รายการ Actual Header ทั้งหมด
 * @param requiredHeader Required Header ที่ต้องการค้นหา
 * @returns Actual Header ที่ใกล้เคียง หรือ undefined
 */
export const findSimilarHeaderForReason = (
  actualHeaders: string[],
  requiredHeader: string,
): string | undefined => {
  /**
   * ลบเครื่องหมายออกจาก Required Header
   * เพื่อให้เหลือเฉพาะตัวอักษรและตัวเลข
   */
  const requiredTextOnly =
    removeSymbols(requiredHeader);

  /**
   * ถ้า Required Header ไม่มีข้อความเหลืออยู่
   * จะไม่สามารถนำไปค้นหา Header ที่ใกล้เคียงได้
   */
  if (!requiredTextOnly) {
    return undefined;
  }

  /**
   * ค้นหา Actual Header ตัวแรก
   * ที่มีข้อความหลังลบเครื่องหมายตรงกับ Required Header
   */
  return actualHeaders.find(
    (actualHeader) => {
      const actualTextOnly =
        removeSymbols(actualHeader);

      return actualTextOnly === requiredTextOnly;
    },
  );
};

/**
 * วิเคราะห์ว่า Actual Header และ Required Header
 * Match หรือไม่ Match เพราะสาเหตุใด
 *
 * ลำดับการตรวจสอบ
 * 1. ตรวจสอบว่า Actual Header ว่างหรือไม่
 * 2. ตรวจสอบว่า Required Header ว่างหรือไม่
 * 3. ตรวจสอบว่า Header Match กันหรือไม่
 * 4. ตรวจสอบว่าเครื่องหมายพิเศษตรงกันหรือไม่
 * 5. ตรวจสอบว่าข้อความตรงกันหรือไม่
 *
 * @param actualHeader Header ที่พบจริง
 * @param requiredHeader Header ที่กำหนดใน Requirement
 * @returns Reason Code ที่อธิบายผลการตรวจสอบ
 */
export const getHeaderMatchReason = (
  actualHeader: unknown,
  requiredHeader: unknown,
): HeaderMatchReason => {
  // ปรับ Actual Header ให้อยู่ในรูปแบบมาตรฐาน
  const actual = canonicalHeader(
    actualHeader,
  );

  // ปรับ Required Header ให้อยู่ในรูปแบบมาตรฐาน
  const required = canonicalHeader(
    requiredHeader,
  );

  // Actual Header ไม่มีข้อมูล
  if (!actual) {
    return "EMPTY_ACTUAL";
  }

  // Required Header ไม่มีข้อมูล
  if (!required) {
    return "EMPTY_REQUIRED";
  }

  // Header ทั้งสองฝั่งตรงกัน
  if (actual === required) {
    return "MATCH";
  }

  /**
   * ดึงเฉพาะเครื่องหมายพิเศษของ Header ทั้งสองฝั่ง
   * เพื่อนำมาเปรียบเทียบ
   */
  const actualSymbols =
    extractSymbols(actualHeader);

  const requiredSymbols =
    extractSymbols(requiredHeader);

  /**
   * ถ้าเครื่องหมายพิเศษไม่เหมือนกัน
   * ให้คืนสถานะ SYMBOL_NOT_MATCH
   */
  if (actualSymbols !== requiredSymbols) {
    return "SYMBOL_NOT_MATCH";
  }

  /**
   * ลบเครื่องหมายออกจาก Header ทั้งสองฝั่ง
   * เพื่อเปรียบเทียบเฉพาะตัวอักษรและตัวเลข
   */
  const actualText =
    removeSymbols(actualHeader);

  const requiredText =
    removeSymbols(requiredHeader);

  /**
   * ถ้าข้อความไม่เหมือนกัน
   * ให้คืนสถานะ TEXT_NOT_MATCH
   */
  if (actualText !== requiredText) {
    return "TEXT_NOT_MATCH";
  }

  /**
   * กรณี Header ไม่ Match แต่ไม่เข้าเงื่อนไขก่อนหน้า
   * ให้ใช้ TEXT_NOT_MATCH เป็นค่าเริ่มต้น
   *
   * หมายเหตุ
   * Code ส่วนนี้คืนค่าเหมือนกับเงื่อนไขด้านบน
   * แต่ยังคงไว้เพื่อไม่เปลี่ยน Logic เดิม
   */
  return "TEXT_NOT_MATCH";
};

/**
 * แปลง Reason Code ให้เป็นข้อความที่อ่านเข้าใจง่าย
 *
 * ใช้สำหรับนำผลการตรวจสอบไปแสดงใน
 * - Remark
 * - Sheet ผลลัพธ์
 * - Log
 *
 * ความหมาย
 * - Reason Code = รหัสสาเหตุ
 * - Reason Text = ข้อความอธิบายสาเหตุ
 *
 * @param reason รหัสสาเหตุที่ต้องการแปลงเป็นข้อความ
 * @returns ข้อความอธิบายสาเหตุ
 */
export const getHeaderMatchReasonText = (
  reason: HeaderMatchReason,
): string => {
  switch (reason) {
    case "MATCH":
      // Matched = Header ตรงกัน
      return "Matched";

    case "EMPTY_ACTUAL":
      // Actual header is empty = Header ที่พบจริงไม่มีข้อมูล
      return "Actual header is empty";

    case "EMPTY_REQUIRED":
      // Required header is empty = Header ใน Requirement ไม่มีข้อมูล
      return "Required header is empty";

    case "SYMBOL_NOT_MATCH":
      // เครื่องหมายพิเศษระหว่าง Actual และ Required ไม่ตรงกัน
      return "Symbol ไม่ตรงกัน เช่น /, -, _, (, ) หายหรือไม่เหมือนกัน";

    case "TEXT_NOT_MATCH":
      // ตัวอักษรหรือคำใน Header ไม่ตรงกัน
      return "คำใน Header ไม่ตรงกัน";

    case "NO_SIMILAR_HEADER":
      // ไม่พบ Actual Header ที่มีข้อความใกล้เคียงกับ Requirement
      return "ไม่พบ Header ที่มีคำใกล้เคียงกับ Requirement";

    default:
      /**
       * Unknown reason = ไม่ทราบสาเหตุ
       *
       * ใช้รองรับกรณีได้รับค่าที่ไม่อยู่ใน Reason Code ที่กำหนด
       */
      return "Unknown reason";
  }
};
