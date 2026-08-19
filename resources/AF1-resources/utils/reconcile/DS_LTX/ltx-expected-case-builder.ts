/**
 * LtxExpectedCaseBuilder
 * ------------------------------------------------------------------
 * สร้าง Expected Case ของ DS_LTX จาก Test Data
 *
 * หลักการ:
 * 1 Test Data Row = 1 Expected Case
 *
 * - ไม่รวมยอด Fee ข้ามระหว่างแถวหลักกับแถว -Return
 * - รวม Fee Amount ทุกกลุ่มภายใน Test Data แถวเดียวกัน
 * - จำนวน Fee Group ตรวจจาก Header จริง ไม่กำหนดแบบ Hard code
 * - ถ้ามี Transaction ID จะใช้สร้าง Expected Reference
 * - ถ้าไม่มี Transaction ID จะส่ง Expected Case ต่อไปให้
 *   ReconcileService ทำ Fallback Matching ด้วย Field และ Amount
 *
 * ลำดับค่าที่แสดงใน Column Test Script No.:
 * 1. Test No.
 * 2. Transaction ID/ Reconcile ID
 * 3. Test Data Row Number
 * ------------------------------------------------------------------
 */

import {
  detectFeeTypeCount,
} from "../../../config/testdata-helper";

import {
  createHeaderAliases,
  findMatchingHeader,
} from "../../validators/shared/header-matcher";

import {
  ReconcileRecord,
} from "../shared/record";

import {
  AmountComparator,
} from "./ltx-amount-compare";

import type {
  ReconcileReportConfig,
} from "./ltx-config";

import type {
  ExpectedCase,
  IExpectedCaseBuilder,
} from "./ltx-expected-case";

/** Header ที่ใช้แสดงหมายเลข Test Case */
const TEST_NO_HEADER =
  "Test No.";

/**
 * สร้าง Expected Case ของ DS_LTX
 */
export class LtxExpectedCaseBuilder
  implements IExpectedCaseBuilder {
  constructor(
    private readonly amountComparator:
      AmountComparator =
      new AmountComparator(),
  ) {}

  /**
   * ตัด Suffix "-Return" ออกจาก Test No.
   *
   * ตัวอย่าง:
   * AAA_01-Return → AAA_01
   * AAA_01        → AAA_01
   *
   * ใช้สำหรับแสดงผลเท่านั้น
   * ไม่ได้นำไปใช้เป็น Matching Key
   */
  private toDisplayTestCaseNo(
    testNo: string,
  ): string {
    return testNo
      .replace(
        /-return$/i,
        "",
      )
      .trim();
  }

  /**
   * รวม Fee Amount ทุกกลุ่มของ Test Data หนึ่งแถว
   *
   * จำนวน Fee Group ตรวจจาก Header จริง
   *
   * ตัวอย่าง:
   * - Fee Amount Type 1
   * - Fee Amount 2
   * - Fee Amount 3
   *
   * หาก Header มี Fee ถึงกลุ่ม 10
   * ระบบจะตรวจและรวมยอดถึงกลุ่ม 10 โดยอัตโนมัติ
   */
  private sumFeeAmountsOfRecord(
    headers: string[],
    record: ReconcileRecord,
  ): number {
    /**
     * ตรวจหมายเลข Fee Group สูงสุด
     * จาก Header ที่มีอยู่จริงใน Test Data
     */
    const feeTypeCount =
      detectFeeTypeCount(
        headers,
      );

    /**
     * สร้าง Alias ของ Fee Header
     * ตามจำนวน Fee Group ที่ตรวจพบ
     *
     * รองรับทั้ง:
     * Fee Amount Type 1
     * Fee Amount 2
     * Fee Amount 3
     */
    const aliases =
      createHeaderAliases(
        feeTypeCount,
      );

    const feeValues:
      string[] = [];

    /**
     * อ่าน Fee Amount ตั้งแต่กลุ่ม 1
     * ถึงกลุ่มสูงสุดที่พบจาก Header
     */
    for (
      let feeIndex = 1;
      feeIndex <= feeTypeCount;
      feeIndex += 1
    ) {
      /**
       * ใช้ชื่อมาตรฐาน Fee Amount Type N
       * แล้วให้ findMatchingHeader() หา Header จริงผ่าน Alias
       *
       * ตัวอย่าง:
       * Expected Header = Fee Amount Type 2
       * Actual Header   = Fee Amount 2
       */
      const actualHeader =
        findMatchingHeader(
          headers,
          `Fee Amount Type ${feeIndex}`,
          aliases,
        );

      if (
        !actualHeader
      ) {
        continue;
      }

      feeValues.push(
        record.get(
          actualHeader,
        ),
      );
    }

    /**
     * รวมยอด Fee ทั้งหมดภายใน Test Data แถวเดียวกัน
     */
    return this.amountComparator.sum(
      feeValues,
    );
  }

  /**
   * สร้าง Expected Case จาก Test Data ทุกแถว
   */
  build(
    headers: string[],
    testDataRecords:
      ReconcileRecord[],
    config:
      ReconcileReportConfig,
  ): ExpectedCase[] {
    return testDataRecords.map(
      (
        record,
      ): ExpectedCase => {
        /**
         * แยกหน้าที่ของ Identity:
         *
         * displayValue:
         * ใช้ Test No. สำหรับแสดงผล
         *
         * matchingReference:
         * ใช้ Transaction ID/ Reconcile ID
         * สำหรับจับคู่กับ DS_LTX Report
         */
        const identity =
          record.resolveIdentity(
            TEST_NO_HEADER,
            config.testDataIdField,
            (
              testNo,
            ) =>
              this.toDisplayTestCaseNo(
                testNo,
              ),
          );

        const transactionId =
          identity.matchingReference;

        const hasTransactionId =
          transactionId !== "";

        /**
         * เลือกค่าที่ใช้แสดงใน Column Test Script No.
         *
         * ลำดับ:
         * 1. Test No.
         * 2. Transaction ID/ Reconcile ID
         * 3. Test Data Row Number
         *
         * ตัวอย่าง:
         * Test Data Row 6
         *
         * รูปแบบนี้รองรับโดย Script 4 ปัจจุบัน
         */
        const displayTestCaseNo =
          identity.displayValue !== ""
            ? identity.displayValue
            : transactionId !== ""
              ? transactionId
              : `Test Data Row ${record.rowNumber}`;

        /**
         * อ่านยอดหลักสำหรับแถว DR
         *
         * ยอดหลัก:
         * From Transfer Amount
         *
         * ยอดสำรอง:
         * From Debit Amount
         */
        const drAmount =
          this.amountComparator.parse(
            record.get(
              config
                .drAmountTestDataField,
            ),
          );

        const drAmountFallback =
          this.amountComparator.parse(
            record.get(
              config
                .drAmountFallbackTestDataField,
            ),
          );

        /**
         * Test Data แถวนี้ต้องมี DR
         * เมื่อยอดหลักหรือยอดสำรองมากกว่า 0.01
         *
         * ค่า 0 หรือค่าว่างไม่ถือว่าต้องมี DR
         */
        const hasDrAmount =
          (
            drAmount !== null &&
            drAmount > 0.01
          ) ||
          (
            drAmountFallback !== null &&
            drAmountFallback > 0.01
          );

        /**
         * รวมยอด Fee Amount ทุกกลุ่ม
         * ภายใน Test Data แถวเดียวกัน
         */
        const feeSum =
          this.sumFeeAmountsOfRecord(
            headers,
            record,
          );

        /**
         * Test Data แถวนี้ต้องมี FE
         * เมื่อยอด Fee รวมมากกว่า 0.01
         */
        const hasFee =
          feeSum > 0.01;

        return {
          displayTestCaseNo,

          primaryRecord:
            record,

          /**
           * กรณีมี DR Amount และมี Transaction ID:
           * สร้าง Expected Reference ปกติ
           *
           * ตัวอย่าง:
           * ABC123DR
           *
           * กรณีไม่มี Transaction ID:
           * คืน undefined แล้วให้ ReconcileService
           * ทำ Fallback Matching ด้วย Field และ Amount
           */
          expectedDrReference:
            hasDrAmount &&
            hasTransactionId
              ? (
                  `${transactionId}` +
                  `${config.drSuffixLabel}`
                )
              : undefined,

          /**
           * กรณีมี Fee Amount และมี Transaction ID:
           * สร้าง Expected Reference ปกติ
           *
           * ตัวอย่าง:
           * ABC123FE
           *
           * กรณีไม่มี Transaction ID:
           * คืน undefined แล้วให้ ReconcileService
           * ทำ Fallback Matching ด้วย Field และ SUM Fee Amount
           */
          expectedFeReference:
            hasFee &&
            hasTransactionId
              ? (
                  `${transactionId}` +
                  `${config.feSuffixLabel}`
                )
              : undefined,

          /**
           * ยอด Fee รวมภายใน Test Data แถวนี้
           */
          expectedFeAmount:
            feeSum.toFixed(
              2,
            ),
        };
      },
    );
  }
}