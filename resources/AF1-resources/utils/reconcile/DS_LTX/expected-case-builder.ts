/**
 * LtxExpectedCaseBuilder
 * ------------------------------------------------------------------
 * (ไม่มีแถวไหนที่ยอดรวมเท่ากับ SUM Fee ทั้งหมดเลย)
 * จึงเปลี่ยนมาสร้าง ExpectedCase แยก โดยใช้หลักการ:
 *
 * "1 แถว Test Data = 1 ExpectedCase"
 *
 * - ไม่รวมยอด Fee ข้ามระหว่างแถวหลักกับแถว -Return
 * - ยังรวม Fee Amount ทุกช่องภายใน Test Data แถวเดียวกัน
 * - ถ้ามี Transaction ID จะใช้สร้าง Expected Reference
 * - ถ้าไม่มี Transaction ID จะส่ง ExpectedCase ไปให้
 *   ReconcileService ทำ Fallback Matching ด้วยข้อมูล Field และ Amount
 * ------------------------------------------------------------------
 */
import {
  createHeaderAliases,
  findMatchingHeader,
} from "../../validators/shared/header-matcher";
import { AmountComparator } from "./amount-compare";
import { ReconcileReportConfig, RECONCILE_FEE_TYPE_COUNT } from "./ltx-config";
import { ExpectedCase, IExpectedCaseBuilder } from "./expected-case";
import { ReconcileRecord } from "../shared/record";

const TEST_NO_HEADER = "Test No.";

export class LtxExpectedCaseBuilder implements IExpectedCaseBuilder {
  constructor(
    private readonly amountComparator: AmountComparator = new AmountComparator(),
  ) {}

  /** ตัด suffix "-Return" ออกจาก Test No. เพื่อให้แถว Return แสดงผลรวมกลุ่มเดียวกับแถวหลักใน Sheet */
  private toDisplayTestCaseNo(testNo: string): string {
    return testNo.replace(/-return$/i, "").trim();
  }

  /** รวมยอด Fee Amount Type 1-5 ของแถวเดียว (ใช้ตัดสินว่าแถวนี้ควรมีแถว FE คู่กันไหม) */
  private sumFeeAmountsOfRecord(
    headers: string[],
    record: ReconcileRecord,
  ): number {
    const aliases = createHeaderAliases(RECONCILE_FEE_TYPE_COUNT);
    const feeValues: string[] = [];

    for (
      let feeIndex = 1;
      feeIndex <= RECONCILE_FEE_TYPE_COUNT;
      feeIndex += 1
    ) {
      const actualHeader = findMatchingHeader(
        headers,
        `Fee Amount Type ${feeIndex}`,
        aliases,
      );
      if (actualHeader) {
        feeValues.push(record.get(actualHeader));
      }
    }

    return this.amountComparator.sum(feeValues);
  }

  build(
    headers: string[],
    testDataRecords: ReconcileRecord[],
    config: ReconcileReportConfig,
  ): ExpectedCase[] {
    return testDataRecords.map((record): ExpectedCase => {
      /**
       * Test No. ใช้สำหรับแสดงผล ไม่ใช่ Matching Key หลัก
       *
       * กรณี Test No. ว่าง:
       * - Script 2 จะ Highlight ช่อง Test No. เป็นสีแดง
       * - ถ้ามี Transaction ID จะใช้ Transaction ID ทำ Exact Matching
       * - Column Test Script No. จะแสดง Transaction ID แทน
       *
       * กรณี Test No. และ Transaction ID ว่างทั้งคู่:
       * - ยังสร้าง ExpectedCase ต่อ ไม่ข้าม Test Data แถวนี้
       * - ReconcileService จะใช้ Field และ Amount ทำ Fallback Matching
       * - Column Test Script No. จะแสดงเลขแถว Test Data แทน
       */
      const identity = record.resolveIdentity(
        TEST_NO_HEADER,
        config.testDataIdField,
        (testNo) => this.toDisplayTestCaseNo(testNo),
      );

      const transactionId = identity.matchingReference;

      const hasTransactionId = transactionId !== "";

      /**
       * ลำดับค่าที่ใช้แสดงใน Column Test Script No.
       *
       * 1. ถ้ามี Test No. ให้แสดง Test No.
       * 2. ถ้าไม่มี Test No. แต่มี Transaction ID
       *    resolveIdentity จะใช้ Transaction ID แสดงแทน
       * 3. ถ้าไม่มีทั้งสองค่า ให้แสดงเลขแถว Test Data
       *    เพื่อให้ QA ยังระบุแถวต้นทางได้
       */
      const displayTestCaseNo =
        identity.displayValue !== ""
          ? identity.displayValue
          : `Test Data Row ${record.rowNumber}`;

      // มีแถว DR ที่คาดหวัง ถ้า Test Data แถวนี้มียอด Transfer Amount หรือ Debit Amount
      // เป็นตัวเลขที่ > 0 จริง ๆ (ไม่ใช่แค่ "ไม่ใช่ string ว่าง") — Bug fix: เดิมเช็คแค่
      // ".trim() !== \"\"" ทำให้ค่า "0" (ศูนย์ตัวเลขจริง ๆ) ถูกนับว่า "มีค่า" ทั้งที่ควร
      // ถือว่า "ไม่มี DR" เจอจริงกับแถว Test Data "-Return" ที่ From Transfer Amount = 0
      // แต่ From Debit Amount เป็นช่องว่าง (ตัวอักษร non-breaking space) — ถ้าไม่กันไว้
      // ระบบจะไปคาดหวัง DR ที่ไม่มีอยู่จริงของแถว Return ทุกแถว
      const drAmount = this.amountComparator.parse(
        record.get(config.drAmountTestDataField),
      );
      const drAmountFallback = this.amountComparator.parse(
        record.get(config.drAmountFallbackTestDataField),
      );
      const hasDrAmount =
        (drAmount !== null && drAmount > 0.01) ||
        (drAmountFallback !== null && drAmountFallback > 0.01);

      // มีแถว FE ที่คาดหวัง ถ้า Test Data แถวนี้มียอด Fee รวม > 0
      const feeSum = this.sumFeeAmountsOfRecord(headers, record);
      const hasFee = feeSum > 0.01;

      return {
        displayTestCaseNo,

        primaryRecord: record,

        /**
         * มี DR Amount แต่ไม่มี Transaction ID:
         * ไม่สร้าง Expected Reference
         * ReconcileService จะใช้ DR Amount ทำ Fallback Matching
         */
        expectedDrReference:
          hasDrAmount && hasTransactionId
            ? `${transactionId}${config.drSuffixLabel}`
            : undefined,

        /**
         * มี Fee Amount แต่ไม่มี Transaction ID:
         * ไม่สร้าง Expected Reference
         * ReconcileService จะใช้ SUM Fee Amount ทำ Fallback Matching
         */
        expectedFeReference:
          hasFee && hasTransactionId
            ? `${transactionId}${config.feSuffixLabel}`
            : undefined,

        expectedFeAmount: feeSum.toFixed(2),
      };
    });
  }
}
