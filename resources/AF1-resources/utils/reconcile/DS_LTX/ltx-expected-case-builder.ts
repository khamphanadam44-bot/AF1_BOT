/**
 * สร้าง Expected Case ของ DS_LTX จาก Test Data
 *
 * หนึ่ง Test Data Row ต้องเป็นหนึ่ง Expected Case เพื่อไม่ให้รายการปกติ
 * และรายการ -Return ถูกนำมารวมกันข้ามแถว ส่วน Fee จะรวมเฉพาะภายในแถวเดียวกัน
 */

import { detectFeeTypeCount } from "../../../config/testdata-helper";
import {
  createHeaderAliases,
  findMatchingHeader,
} from "../../validators/shared/header-matcher";
import type { ReconcileRecord } from "../shared/record";
import { AmountComparator } from "./ltx-amount-compare";
import type { ReconcileReportConfig } from "./ltx-config";

const FEE_AMOUNT_HEADER_PREFIX = "Fee Amount Type";

export interface ExpectedCase {
  displayTestCaseNo: string;
  primaryRecord: ReconcileRecord;
  expectedDrReference: string | undefined;
  expectedFeReference: string | undefined;
  hasExpectedDr: boolean;
  hasExpectedFe: boolean;
  expectedFeAmount: string;
}

/** 1 Test Data row สร้าง 1 Expected Case; Return row จะไม่ถูกรวมข้ามแถว */
export class LtxExpectedCaseBuilder {
  constructor(
    private readonly amountComparator: AmountComparator = new AmountComparator(),
  ) {}

  /** หา Fee headers เพียงครั้งเดียว เพราะทุก record ใช้โครงสร้างเดียวกัน */
  private findFeeAmountHeaders(headers: string[]): string[] {
    const feeTypeCount = detectFeeTypeCount(headers);
    const aliases = createHeaderAliases(feeTypeCount);
    const feeAmountHeaders: string[] = [];

    for (let feeIndex = 1; feeIndex <= feeTypeCount; feeIndex += 1) {
      const actualHeader = findMatchingHeader(
        headers,
        `${FEE_AMOUNT_HEADER_PREFIX} ${feeIndex}`,
        aliases,
      );

      if (actualHeader) {
        feeAmountHeaders.push(actualHeader);
      }
    }

    return feeAmountHeaders;
  }

  private sumFeeAmounts(
    record: ReconcileRecord,
    feeAmountHeaders: readonly string[],
  ): number {
    // ทุกค่าในรายการนี้มาจาก Test Data แถวเดียวกัน จึงไม่รวม Fee ข้าม Case
    return this.amountComparator.sum(
      feeAmountHeaders.map((header) => record.get(header)),
    );
  }

  /** Amount ต้องมากกว่า 0 จึงถือว่า Test Data คาดหวังรายการใน Report */
  private hasPositiveAmount(record: ReconcileRecord, field: string): boolean {
    const amount = this.amountComparator.parse(record.get(field));
    return amount !== null && amount > 0;
  }

  build(
    headers: string[],
    testDataRecords: ReconcileRecord[],
    config: ReconcileReportConfig,
  ): ExpectedCase[] {
    // Workbook หนึ่งไฟล์ใช้ Header ชุดเดียวกัน จึงค้นหา Fee Header ก่อนวนทุกแถว
    const feeAmountHeaders = this.findFeeAmountHeaders(headers);

    return testDataRecords.map((record) => {
      const identity = record.resolveIdentity(
        config.testDataTestNoField,
        config.testDataIdField,
        // รักษา Test No. ตาม Test Data รวมถึง -Return; ตัดเฉพาะช่องว่างหัวท้าย
        (testNo) => testNo.trim(),
      );
      const transactionId = identity.matchingReference;

      /** ใช้ Test No. ก่อน แล้วจึง fallback ไป Transaction ID หรือเลขแถว */
      const displayTestCaseNo =
        identity.displayValue ||
        transactionId ||
        `Test Data Row ${record.rowNumber}`;
      const hasExpectedDr =
        this.hasPositiveAmount(record, config.drAmountTestDataField) ||
        this.hasPositiveAmount(record, config.drAmountFallbackTestDataField);

      /** Fee มากกว่า 0 หมายถึง Test Data คาดหวังแถว FE */
      const feeSum = this.sumFeeAmounts(record, feeAmountHeaders);
      const hasExpectedFe = feeSum > 0;

      return {
        displayTestCaseNo,
        primaryRecord: record,
        // ไม่มี Transaction ID ให้คืน undefined เพื่อส่งต่อไปทำ Fallback Matching
        expectedDrReference:
          hasExpectedDr && transactionId
            ? `${transactionId}${config.drSuffixLabel}`
            : undefined,
        expectedFeReference:
          hasExpectedFe && transactionId
            ? `${transactionId}${config.feSuffixLabel}`
            : undefined,
        hasExpectedDr,
        hasExpectedFe,
        expectedFeAmount: feeSum.toFixed(2),
      };
    });
  }
}
