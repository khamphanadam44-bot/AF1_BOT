/**
 * header-validation-sheet.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * ใช้คำนวณผลการตรวจสอบ Header ระหว่าง
 * - Actual Header   = Header ที่พบจริงในไฟล์ Excel
 * - Expected Header = Header ที่ระบบคาดหวังจาก Requirement
 *
 * ความสามารถหลัก
 * 1. ตรวจสอบว่า Expected Header แต่ละรายการพบใน Actual Header หรือไม่
 * 2. รองรับการจับคู่ Header ผ่าน Alias
 * 3. ค้นหา Header ที่มีข้อความใกล้เคียงเมื่อไม่พบ Header
 * 4. ระบุเหตุผลว่า Header ไม่ตรงกันเพราะอะไร
 * 5. คืนรายการ Header ที่ Missing
 * 6. ตรวจสอบว่า Actual Header อยู่ใน Requirement หรือไม่
 *
 * หมายเหตุ
 * ไฟล์นี้ไม่ได้สร้าง Sheet "Header Validation"
 * และไม่ได้เขียนข้อมูลลงในไฟล์ Excel โดยตรง
 *
 * ไฟล์นี้ทำหน้าที่คำนวณผลลัพธ์เท่านั้น
 * จากนั้น Validator อื่นจะนำผลลัพธ์ไปใช้งานต่อ
 *
 * คำศัพท์
 * - Actual Header   = Header ที่พบจริงในไฟล์
 * - Expected Header = Header ที่ระบบคาดหวัง
 * - Found           = พบ Header
 * - Missing         = ไม่พบ Header
 * - Alias           = ชื่ออื่นที่อนุญาตให้ใช้แทนชื่อหลัก
 * - Similar Header  = Header ที่มีข้อความใกล้เคียง
 * - Remark          = คำอธิบายผลการตรวจสอบ
 * ------------------------------------------------------------------
 */

import {
  findMatchingHeader,
  findSimilarHeaderForReason,
  getHeaderMatchReason,
  getHeaderMatchReasonText,
  isHeaderMatch,
} from "./header-matcher";

/**
 * รูปแบบข้อมูลผลการตรวจสอบ Header จำนวน 1 รายการ
 *
 * ตัวอย่างผลลัพธ์เมื่อพบ Header:
 *
 * {
 *   expectedHeader: "Currency Id",
 *   matchedHeader: "Currency ID",
 *   isFound: true,
 *   similarHeader: undefined,
 *   remark: "Matched by Alias"
 * }
 *
 * ตัวอย่างผลลัพธ์เมื่อไม่พบ Header:
 *
 * {
 *   expectedHeader: "Transaction Date",
 *   matchedHeader: undefined,
 *   isFound: false,
 *   similarHeader: "Transaction-Date",
 *   remark: "Symbol ไม่ตรงกัน..."
 * }
 */
export type HeaderMatchResult = {
  /**
   * ชื่อ Header ที่กำหนดไว้ใน Requirement
   * และเป็น Header ที่ระบบกำลังค้นหา
   */
  expectedHeader: string;

  /**
   * ชื่อ Header จริงที่ Match
   *
   * ถ้าพบ Header จะมีค่าเป็นชื่อ Header ที่พบจริง
   * ถ้าไม่พบจะมีค่าเป็น undefined
   */
  matchedHeader: string | undefined;

  /**
   * ผลการค้นหา Header
   *
   * true  = พบ Header
   * false = ไม่พบ Header
   */
  isFound: boolean;

  /**
   * Header ที่มีข้อความใกล้เคียงกับ Expected Header
   *
   * ใช้ช่วยอธิบายสาเหตุเมื่อไม่พบ Header ที่ Match จริง
   * ถ้าไม่พบ Header ใกล้เคียงจะเป็น undefined
   */
  similarHeader: string | undefined;

  /**
   * คำอธิบายผลการตรวจสอบ เช่น
   * - Requirement Found
   * - Matched by Alias
   * - Symbol ไม่ตรงกัน
   * - ไม่พบ Header ที่มีคำใกล้เคียง
   */
  remark: string;
};

/**
 * ค้นหา Actual Header ที่ Match กับ Expected Header
 *
 * ฟังก์ชันนี้เป็นตัวกลางสำหรับเรียกใช้
 * findMatchingHeader() จาก header-matcher.ts
 *
 * การค้นหารองรับ Alias ที่ได้รับเข้ามา
 *
 * @param actualHeaders รายการ Header ที่พบจริงในไฟล์
 * @param expectedHeader Header ที่ระบบต้องการค้นหา
 * @param aliases รายชื่ออื่นที่อนุญาตให้ใช้แทน Header หลัก
 *
 * @returns
 * - ชื่อ Actual Header เมื่อพบ Header ที่ Match
 * - undefined เมื่อไม่พบ Header
 */
const findMatchedHeader = (
  actualHeaders: string[],
  expectedHeader: string,
  aliases: Record<string, string[]>,
): string | undefined => {
  return findMatchingHeader(
    actualHeaders,
    expectedHeader,
    aliases,
  );
};

/**
 * วิเคราะห์สาเหตุที่ Expected Header เป็น Missing
 *
 * ฟังก์ชันนี้จะถูกเรียกเมื่อไม่พบ Header ที่ Match จริง
 *
 * การทำงาน
 * 1. ค้นหา Actual Header ที่มีข้อความใกล้เคียง
 * 2. ถ้าไม่พบ Header ใกล้เคียง ให้คืน NO_SIMILAR_HEADER
 * 3. ถ้าพบ Header ใกล้เคียง ให้วิเคราะห์ว่าไม่ Match เพราะอะไร
 * 4. แปลง Reason Code ให้เป็นข้อความที่อ่านง่าย
 *
 * ตัวอย่าง
 *
 * Expected Header:
 * "Transaction ID/ Reconcile ID"
 *
 * Actual Header:
 * "Transaction ID Reconcile ID"
 *
 * ผลลัพธ์:
 * - similarHeader = "Transaction ID Reconcile ID"
 * - reason = เครื่องหมาย Symbol ไม่ตรงกัน
 *
 * หมายเหตุ
 * การค้นหา Similar Header ในฟังก์ชันนี้
 * เปรียบเทียบกับ expectedHeader โดยตรง ไม่ได้ตรวจสอบผ่าน Alias
 *
 * @param actualHeaders รายการ Header ที่พบจริงในไฟล์
 * @param expectedHeader Header ที่ระบบคาดหวัง
 *
 * @returns Object ที่ประกอบด้วย
 * - similarHeader = Header ที่มีข้อความใกล้เคียง
 * - reason = ข้อความอธิบายสาเหตุ
 */
const getMissingReason = (
  actualHeaders: string[],
  expectedHeader: string,
): {
  similarHeader: string | undefined;
  reason: string;
} => {
  /**
   * ค้นหา Header ที่มีคำหรือตัวอักษรใกล้เคียงกัน
   * แต่อาจมีเครื่องหมายพิเศษแตกต่างกัน
   */
  const similarHeader =
    findSimilarHeaderForReason(
      actualHeaders,
      expectedHeader,
    );

  /**
   * ถ้าไม่พบ Similar Header
   * ให้คืนเหตุผลว่าไม่พบ Header ที่มีคำใกล้เคียง
   */
  if (!similarHeader) {
    return {
      similarHeader: undefined,

      /**
       * แปลง NO_SIMILAR_HEADER
       * ให้เป็นข้อความที่อ่านเข้าใจง่าย
       */
      reason:
        getHeaderMatchReasonText(
          "NO_SIMILAR_HEADER",
        ),
    };
  }

  /**
   * ถ้าพบ Similar Header
   *
   * ขั้นตอนการทำงาน
   * 1. getHeaderMatchReason() วิเคราะห์สาเหตุ
   * 2. getHeaderMatchReasonText() แปลงสาเหตุเป็นข้อความ
   */
  const reason =
    getHeaderMatchReasonText(
      getHeaderMatchReason(
        similarHeader,
        expectedHeader,
      ),
    );

  // คืน Header ใกล้เคียงและเหตุผลที่ไม่ Match
  return {
    similarHeader,
    reason,
  };
};

/**
 * คำนวณผลการตรวจสอบ Header ทั้งหมด
 *
 * ฟังก์ชันนี้เป็นฟังก์ชันหลักของไฟล์
 *
 * การทำงาน
 * 1. วนตรวจสอบ Expected Header ทีละรายการ
 * 2. ค้นหา Actual Header ที่ Match
 * 3. ตรวจสอบว่า Match ด้วยชื่อโดยตรงหรือ Alias
 * 4. ถ้าพบ Header ให้สร้างผลลัพธ์สถานะ Found
 * 5. ถ้าไม่พบ ให้ค้นหา Similar Header และวิเคราะห์เหตุผล
 * 6. คืนผลลัพธ์ทั้งหมดเป็น Array
 *
 * จำนวนผลลัพธ์ที่ได้จะเท่ากับจำนวน Expected Header
 *
 * ตัวอย่าง
 *
 * expectedHeaders มี 10 รายการ
 * ผลลัพธ์ HeaderMatchResult[] จะมี 10 รายการ
 *
 * @param actualHeaders รายการ Header ที่พบจริงในไฟล์
 * @param expectedHeaders รายการ Header จาก Requirement
 * @param aliases รายชื่อ Alias ที่อนุญาตให้ใช้จับคู่
 *
 * @returns Array ผลการตรวจสอบ Header ทุก Requirement
 */
export const findHeaderMatchResults = (
  actualHeaders: string[],
  expectedHeaders: string[],
  aliases: Record<string, string[]>,
): HeaderMatchResult[] => {
  /**
   * map() จะวน Expected Header ทุกตัว
   * และเปลี่ยนแต่ละตัวให้เป็น HeaderMatchResult
   */
  return expectedHeaders.map(
    (expectedHeader) => {
      /**
       * ค้นหา Actual Header ที่ Match
       * รองรับทั้งชื่อโดยตรงและ Alias
       */
      const matchedHeader =
        findMatchedHeader(
          actualHeaders,
          expectedHeader,
          aliases,
        );

      /**
       * แปลง matchedHeader ให้เป็น Boolean
       *
       * ถ้าพบ Header:
       * matchedHeader มีข้อความ → true
       *
       * ถ้าไม่พบ:
       * matchedHeader เป็น undefined → false
       */
      const isFound =
        Boolean(matchedHeader);

      /**
       * กรณีพบ Header ที่ Match
       */
      if (isFound && matchedHeader) {
        /**
         * ตรวจสอบว่า Match ด้วยชื่อโดยตรงหรือผ่าน Alias
         *
         * ถ้า matchedHeader Match กับ expectedHeader โดยตรง
         * aliasMatched จะเป็น false
         *
         * ถ้า matchedHeader ไม่ตรงกับ expectedHeader โดยตรง
         * แต่ค้นหาพบผ่าน Alias
         * aliasMatched จะเป็น true
         */
        const aliasMatched =
          !isHeaderMatch(
            matchedHeader,
            expectedHeader,
          );

        /**
         * กำหนดคำอธิบายผลการค้นหา
         *
         * Matched by Alias
         * = พบ Header จากชื่อ Alias
         *
         * Requirement Found
         * = พบ Header ที่ตรงกับ Requirement โดยตรง
         */
        const remark =
          aliasMatched
            ? "Matched by Alias"
            : "Requirement Found";

        /**
         * สร้างข้อความ Mapping สำหรับแสดงใน Console
         *
         * ถ้า Match ผ่าน Alias:
         * " -> Actual Header"
         *
         * ถ้า Match โดยตรง:
         * ใช้ข้อความว่าง
         */
        const mapping =
          aliasMatched
            ? ` -> ${matchedHeader}`
            : "";

        /**
         * แสดงผลใน Console เมื่อพบ Header
         *
         * ตัวอย่าง Match โดยตรง:
         * 🟢 FOUND : Currency Id
         *
         * ตัวอย่าง Match ผ่าน Alias:
         * 🟢 FOUND : Txn Date -> Transaction Date
         */
        console.log(
          `🟢 FOUND : ${expectedHeader}${mapping}`,
        );

        /**
         * คืนผลลัพธ์กรณีพบ Header
         */
        return {
          expectedHeader,
          matchedHeader,
          isFound: true,

          /**
           * กรณีพบ Header จริงแล้ว
           * จึงไม่ต้องเก็บ Similar Header
           */
          similarHeader: undefined,

          remark,
        };
      }

      /**
       * กรณีไม่พบ Header ที่ Match
       *
       * ค้นหา Header ที่ใกล้เคียง
       * และวิเคราะห์สาเหตุที่ไม่ Match
       */
      const {
        similarHeader,
        reason,
      } = getMissingReason(
        actualHeaders,
        expectedHeader,
      );

      /**
       * แสดงผลใน Console เมื่อไม่พบ Header
       *
       * ตัวอย่าง:
       * 🔴 MISSING : Transaction Date
       * | Similar: Transaction-Date
       * | Reason: Symbol ไม่ตรงกัน
       *
       * ถ้าไม่มี Similar Header จะแสดงเครื่องหมาย -
       */
      console.log(
        `🔴 MISSING : ${expectedHeader} | Similar: ${similarHeader ?? "-"} | Reason: ${reason}`,
      );

      /**
       * คืนผลลัพธ์กรณีไม่พบ Header
       */
      return {
        expectedHeader,
        matchedHeader: undefined,
        isFound: false,
        similarHeader,

        /**
         * ใช้เหตุผลที่วิเคราะห์ได้เป็น Remark
         */
        remark: reason,
      };
    },
  );
};

/**
 * ดึงเฉพาะ Expected Header ที่มีสถานะ Missing
 *
 * การทำงาน
 * 1. filter() เลือกเฉพาะผลลัพธ์ที่ isFound เป็น false
 * 2. map() ดึงเฉพาะชื่อ expectedHeader
 *
 * ตัวอย่าง Input:
 *
 * [
 *   {
 *     expectedHeader: "Currency Id",
 *     isFound: true
 *   },
 *   {
 *     expectedHeader: "Payment Method",
 *     isFound: false
 *   }
 * ]
 *
 * ผลลัพธ์:
 *
 * ["Payment Method"]
 *
 * @param results ผลการตรวจสอบ Header ทั้งหมด
 * @returns รายชื่อ Expected Header ที่ไม่พบ
 */
export const getMissingHeaders = (
  results: HeaderMatchResult[],
): string[] => {
  return results
    /**
     * เลือกเฉพาะรายการที่ไม่พบ Header
     */
    .filter(
      (result) => !result.isFound,
    )

    /**
     * ดึงเฉพาะชื่อ Expected Header
     */
    .map(
      (result) => result.expectedHeader,
    );
};

/**
 * ตรวจสอบว่า Actual Header เป็น Header
 * ที่อยู่ใน Expected Header หรือ Requirement หรือไม่
 *
 * ใช้สำหรับตัดสินใจว่า Header ควรถูก Highlight หรือไม่
 *
 * การทำงาน
 * 1. รับ Actual Header มาเพียง 1 รายการ
 * 2. วนตรวจสอบกับ Expected Header ทุกตัว
 * 3. รองรับการ Match ผ่าน Alias
 * 4. ถ้า Match อย่างน้อย 1 รายการ ให้คืนค่า true
 * 5. ถ้าไม่ Match ทุกตัว ให้คืนค่า false
 *
 * ตัวอย่าง
 *
 * actualHeader:
 * "Transaction Date"
 *
 * expectedHeaders:
 * ["Txn Date", "Currency Id"]
 *
 * aliases:
 * {
 *   "Txn Date": ["Txn Date", "Transaction Date"]
 * }
 *
 * ผลลัพธ์:
 * true
 *
 * เพราะ "Transaction Date" เป็น Alias ของ "Txn Date"
 *
 * @param actualHeader Actual Header ที่ต้องการตรวจสอบ
 * @param expectedHeaders รายการ Header จาก Requirement
 * @param aliases Alias ที่ใช้สำหรับจับคู่ Header
 *
 * @returns
 * true  = Header อยู่ใน Requirement
 * false = Header ไม่ได้อยู่ใน Requirement
 */
export const isExpectedHeader = (
  actualHeader: string,
  expectedHeaders: string[],
  aliases: Record<string, string[]>,
): boolean => {
  /**
   * some() จะหยุดตรวจสอบทันที
   * เมื่อพบ Expected Header อย่างน้อย 1 รายการที่ Match
   */
  return expectedHeaders.some(
    (expectedHeader) =>
      Boolean(
        /**
         * ส่ง Actual Header เพียงตัวเดียวเข้าไปใน Array
         * เพื่อใช้ Logic กลางจาก findMatchedHeader()
         */
        findMatchedHeader(
          [actualHeader],
          expectedHeader,
          aliases,
        ),
      ),
  );
};