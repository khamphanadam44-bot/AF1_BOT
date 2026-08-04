/**
 * LtxExpectedCaseBuilder
 * ------------------------------------------------------------------
 * implements IExpectedCaseBuilder — plugin เฉพาะ DS_LTX
 * report อื่นที่ยังไม่มี logic นี้ (เช่น DS_PTX) เขียน class implements
 * IExpectedCaseBuilder ของตัวเองแยกต่างหากได้ โดยไม่ต้องแก้ไฟล์นี้
 *
 * เดิมไฟล์นี้ "รวม" แถวหลักกับแถว -Return เข้าด้วยกัน
 * เป็น ExpectedCase เดียว (SUM ยอด Fee ทุกแถว) แล้วเอายอดรวมไปหาแถว Report ที่ยอดตรงกัน
 * — ตรวจสอบกับไฟล์จริงแล้วพบว่าแถวหลักและแถว -Return แต่ละแถวมี Reference Transaction
 * Number ของตัวเอง (จาก Transaction ID/ Reconcile ID) แยกกันคนละแถวใน AF1 Report จริง ๆ
 * (ไม่มีแถวไหนที่ยอดรวมเท่ากับ SUM Fee ทั้งหมดเลย) จึงเปลี่ยนมาสร้าง ExpectedCase แยก
 * "1 แถว Test Data = 1 ExpectedCase" ใช้ Transaction ID โดยตรงเป็น Reference Number
 * ที่คาดหวัง ไม่ต้องคำนวณ/เดายอดรวมอีกต่อไป
 * ------------------------------------------------------------------
 */
import {
  createHeaderAliases,
  findMatchingHeader,
} from "../../validators/shared/header-matcher";
import { AmountComparator } from "./ltx-amount-compare";
import { ReconcileReportConfig, RECONCILE_FEE_TYPE_COUNT } from "./ltx-config";
import { ExpectedCase, IExpectedCaseBuilder } from "./ltx-expected-case";
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
    return testDataRecords
      .map((record): ExpectedCase | null => {
        /**
         * Test No. ใช้สำหรับแสดงผล แต่ไม่ใช่ Matching Key หลัก
         *
         * ถ้า Test No. ว่าง:
         * - Script 2 จะ Highlight ช่อง Test No. เป็นสีแดง
         * - Script 3 ยัง Reconcile แถวนี้ต่อ โดยใช้ Transaction ID
         * - Column Test Script No. จะแสดง Transaction ID แทน
         */
        const identity = record.resolveIdentity(
          TEST_NO_HEADER,
          config.testDataIdField,
          (testNo) => this.toDisplayTestCaseNo(testNo),
        );

        const transactionId = identity.matchingReference;

        if (transactionId === "") {
          return null; // ไม่มี Transaction ID/ Reconcile ID ให้ผูกกับ Report เลย ข้ามแถวนี้
        }

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
          displayTestCaseNo: identity.displayValue,
          primaryRecord: record,
          expectedDrReference: hasDrAmount
            ? `${transactionId}${config.drSuffixLabel}`
            : undefined,
          expectedFeReference: hasFee
            ? `${transactionId}${config.feSuffixLabel}`
            : undefined,
          expectedFeAmount: feeSum.toFixed(2),
        };
      })
      .filter(
        (expectedCase): expectedCase is ExpectedCase => expectedCase !== null,
      );
  }
}
