/**
 * ltx-expected-case-builder.ts
 * ------------------------------------------------------------
 * สร้าง Expected Case ของ DS_LTX จาก Test Data
 * ก่อนส่งเข้า Reconcile Engine
 *
 * หน้าที่หลัก:
 * 1. อ่าน Test No. และ Transaction ID
 * 2. ตรวจว่า Test Data แถวนั้นต้องมีแถว DR หรือไม่
 * 3. ตรวจว่า Test Data แถวนั้นต้องมีแถว FE หรือไม่
 * 4. รวมยอด Fee Amount ของแต่ละแถว
 * 5. สร้าง Expected Reference Number ของ DR และ FE
 *
 * หลักการสร้าง Expected Case:
 * 1 Test Data Row = 1 Expected Case
 *
 * ไม่รวมแถวหลักกับแถวที่ลงท้ายด้วย "-Return"
 * เป็น Expected Case เดียวกัน
 *
 * สาเหตุ:
 * แถวหลักและแถว Return มี Transaction ID ของตัวเอง
 * และมี Reference Transaction Number แยกกันใน AF1 Report
 *
 * จำนวน Fee Group:
 * ไม่กำหนดเป็น Hard code
 * แต่ตรวจจาก Header ของ Test Data จริง
 * ------------------------------------------------------------
 */

import {
  detectFeeTypeCount,
} from "../../../config/testdata-helper";

import {
  createHeaderAliases,
  findMatchingHeader,
} from "../../validators/shared/header-matcher";

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

import {
  ReconcileRecord,
} from "../shared/record";

/**
 * Header ที่ใช้แสดงหมายเลข Test Case
 */
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
  ) { }

  /**
   * ตัด Suffix "-Return" ออกจาก Test No.
   *
   * ใช้เฉพาะค่าที่นำไปแสดงในผลลัพธ์
   * ไม่ได้ใช้เป็น Matching Key
   *
   * ตัวอย่าง:
   * AAA_01-Return → AAA_01
   * AAA_01        → AAA_01
   *
   * ทำให้แถวหลักและแถว Return
   * แสดงอยู่ภายใต้ Test Case เดียวกัน
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
   * รวมยอด Fee Amount ของ Test Data หนึ่งแถว
   *
   * จำนวน Fee Group จะตรวจจาก Header จริง
   *
   * ตัวอย่าง Header:
   * - Fee Amount Type 1
   * - Fee Amount 2
   * - Fee Amount 3
   * - Fee Amount 4
   * - Fee Amount 5
   *
   * ระบบจะ:
   * 1. ตรวจหมายเลข Fee Group สูงสุด
   * 2. สร้าง Header Alias ตามจำนวนที่พบ
   * 3. อ่าน Fee Amount ของแต่ละกลุ่ม
   * 4. รวมยอด Fee Amount ทั้งหมด
   *
   * ใช้ยอดรวมเพื่อตัดสินว่า Test Data แถวนี้
   * ควรมีแถว FE ใน AF1 Report หรือไม่
   */
  private sumFeeAmountsOfRecord(
    headers: string[],
    record: ReconcileRecord,
  ): number {
    /**
     * ตรวจจำนวน Fee Group จาก Header จริง
     *
     * ไม่ใช้ RECONCILE_FEE_TYPE_COUNT
     * หรือจำนวนแบบ Hard code
     */
    const feeTypeCount =
      detectFeeTypeCount(
        headers,
      );

    /**
     * สร้าง Alias สำหรับ Fee Header
     *
     * รองรับชื่อ:
     * Fee Amount Type 1
     * Fee Amount 2
     * Fee Amount 3
     * ...
     */
    const aliases =
      createHeaderAliases(
        feeTypeCount,
      );

    const feeValues:
      string[] = [];

    /**
     * อ่าน Fee Amount ตั้งแต่กลุ่ม 1
     * ถึงกลุ่มสูงสุดที่พบจาก Header จริง
     */
    for (
      let feeIndex = 1;

      feeIndex <=
      feeTypeCount;

      feeIndex += 1
    ) {
      /**
       * ใช้ชื่อรูปแบบ Fee Amount Type N
       * เป็น Expected Header
       *
       * findMatchingHeader() จะช่วยจับคู่ผ่าน Alias
       * กับชื่อ Header จริง เช่น Fee Amount 2
       */
      const actualHeader =
        findMatchingHeader(
          headers,
          `Fee Amount Type ${feeIndex}`,
          aliases,
        );

      /**
       * หากพบ Header ของ Fee Amount
       * ให้เก็บค่าเพื่อนำไปรวมยอด
       */
      if (actualHeader) {
        feeValues.push(
          record.get(
            actualHeader,
          ),
        );
      }
    }

    /**
     * รวมค่า Fee Amount ทั้งหมด
     *
     * AmountComparator จะจัดการ:
     * - ค่าว่าง
     * - ข้อความ
     * - ตัวเลข
     * - Decimal
     */
    return this.amountComparator.sum(
      feeValues,
    );
  }

  /**
   * สร้าง Expected Case จาก Test Data ทุกแถว
   *
   * @param headers
   * Header จริงของ Test Data
   *
   * @param testDataRecords
   * ข้อมูล Test Data ที่แปลงเป็น ReconcileRecord แล้ว
   *
   * @param config
   * Config สำหรับ Reconcile DS_LTX
   *
   * @returns
   * รายการ Expected Case ที่พร้อมส่งเข้า Reconcile Engine
   */
  build(
    headers: string[],
    testDataRecords:
      ReconcileRecord[],
    config:
      ReconcileReportConfig,
  ): ExpectedCase[] {
    return testDataRecords
      .map(
        (
          record,
        ): ExpectedCase | null => {
          /**
           * Test No. ใช้สำหรับแสดงผล
           * แต่ไม่ใช่ Matching Key หลัก
           *
           * Matching Key หลักคือ:
           * Transaction ID/ Reconcile ID
           *
           * ถ้า Test No. ว่าง:
           * - Script 2 จะ Highlight ช่อง Test No. สีแดง
           * - Script 3 จะยัง Reconcile ต่อ
           * - ใช้ Transaction ID เป็น Matching Key
           * - Column Test Script No. จะแสดง Transaction ID แทน
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

          /**
           * ถ้าไม่มี Transaction ID
           * จะไม่สามารถผูก Test Data กับ AF1 Report ได้
           *
           * จึงข้ามแถวนี้โดยคืน null
           */
          if (
            transactionId === ""
          ) {
            return null;
          }

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
           * การตรวจแบบตัวเลขป้องกันปัญหาเดิม:
           * ค่า "0" ถูกมองว่าเป็นข้อมูล
           * และทำให้ระบบคาดหวังแถว DR ที่ไม่มีจริง
           */
          const hasDrAmount =
            (
              drAmount !== null &&
              drAmount > 0.01
            ) ||
            (
              drAmountFallback !==
              null &&
              drAmountFallback >
              0.01
            );

          /**
           * รวมยอด Fee Amount ทุกกลุ่ม
           * ที่พบจาก Header จริง
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

          /**
           * สร้าง Expected Reference
           *
           * DR:
           * Transaction ID + DR Suffix
           *
           * FE:
           * Transaction ID + FE Suffix
           */
          return {
            displayTestCaseNo:
              identity.displayValue,

            primaryRecord:
              record,

            expectedDrReference:
              hasDrAmount
                ? `${transactionId}${config.drSuffixLabel}`
                : undefined,

            expectedFeReference:
              hasFee
                ? `${transactionId}${config.feSuffixLabel}`
                : undefined,

            expectedFeAmount:
              feeSum.toFixed(
                2,
              ),
          };
        },
      )
      .filter(
        (
          expectedCase,
        ): expectedCase is ExpectedCase =>
          expectedCase !== null,
      );
  }
}