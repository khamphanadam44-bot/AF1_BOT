/**
 * ftu-matcher.ts
 * ------------------------------------------------------------------
 * รับผิดชอบเฉพาะการหา AF1 Row สำหรับ Test Data หนึ่งแถว
 *
 * หน้าที่:
 * 1. สร้าง Index จาก Arr Number
 * 2. จอง AF1 Row ที่มี Exact Transaction ID
 * 3. ระบุ Direction จาก From/To Currency
 * 4. หา Exact Match ก่อน แล้วจึงหา Unique Fallback Match
 *
 * ไฟล์นี้ไม่ตัดสิน PASS/FAIL/REVIEW และไม่อ่านหรือเขียน Excel
 * ------------------------------------------------------------------
 */

import type { ReconcileRecord } from "../shared/record";
import {
  isDateMatchWithArrangementFallback,
  parseAmount,
  parseDate,
} from "../shared/reconcile-parse.util";
import {
  FTU_AMOUNT_TOLERANCE,
  FTU_DIRECTIONS,
  FTU_LEG_TYPES,
  FTU_REMARKS,
  FTU_REPORT_FIELDS,
  FTU_TEST_DATA_FIELDS,
  FTU_THB_CURRENCY_CODE,
  normalizeFtuText,
} from "./ftu-config";
import type { FtuDirection } from "./ftu-config";

interface FtuFallbackMatchResult {
  matchedRecord?: ReconcileRecord;
  candidateCount?: number;
  remark: string;
}

export interface FtuMatchResult {
  transactionId: string;
  direction: FtuDirection;
  exactMatchedRecord?: ReconcileRecord;
  fallbackResult?: FtuFallbackMatchResult;
  fallbackUnavailableRemark?: string;
  failureRemark?: string;
}

export class FtuMatcher {
  /** สร้าง Index จาก Arr Number โดยคงกติกาเดิมว่า Arr Number ไม่ซ้ำ */
  indexReportRecords(
    records: ReconcileRecord[],
  ): Map<string, ReconcileRecord> {
    const recordsById = new Map<string, ReconcileRecord>();

    for (const record of records) {
      const arrangementNumber = normalizeFtuText(
        record.get(FTU_REPORT_FIELDS.arrangementNumber),
      );

      if (arrangementNumber !== "") {
        recordsById.set(arrangementNumber, record);
      }
    }

    return recordsById;
  }

  /**
   * จอง AF1 Row ที่ Test Data ระบุ Transaction ID ไว้แล้ว
   * เพื่อไม่ให้ Fallback ของ Test Case อื่นนำ Row นั้นไปใช้ก่อน
   */
  findReservedReportRows(
    testDataRecords: ReconcileRecord[],
    reportRecordsById: ReadonlyMap<string, ReconcileRecord>,
  ): Set<number> {
    const reservedRowNumbers = new Set<number>();

    for (const testDataRecord of testDataRecords) {
      const transactionId = normalizeFtuText(
        testDataRecord.get(FTU_TEST_DATA_FIELDS.transactionId),
      );
      const reservedRecord = reportRecordsById.get(transactionId);

      if (reservedRecord) {
        reservedRowNumbers.add(reservedRecord.rowNumber);
      }
    }

    return reservedRowNumbers;
  }

  /** หา Exact/Fallback Candidate โดยไม่ตัดสินผลของ Test Case */
  findMatch(
    testDataRecord: ReconcileRecord,
    reportRecordsById: ReadonlyMap<string, ReconcileRecord>,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): FtuMatchResult {
    const transactionId = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.transactionId),
    );
    const direction = this.resolveDirection(testDataRecord);
    const exactMatchedRecord = reportRecordsById.get(transactionId);

    if (
      exactMatchedRecord &&
      usedReportRowNumbers.has(exactMatchedRecord.rowNumber)
    ) {
      return {
        transactionId,
        direction,
        failureRemark:
          `พบ ${FTU_REPORT_FIELDS.arrangementNumber} = "${transactionId}" ` +
          `ที่ Report row ${exactMatchedRecord.rowNumber} ` +
          "แต่แถวดังกล่าวถูก Test Case อื่นใช้แล้ว",
      };
    }

    if (exactMatchedRecord) {
      return {
        transactionId,
        direction,
        exactMatchedRecord,
      };
    }

    if (direction === FTU_DIRECTIONS.unknown) {
      return {
        transactionId,
        direction,
        fallbackUnavailableRemark: FTU_REMARKS.fallbackDirectionUnavailable,
      };
    }

    return {
      transactionId,
      direction,
      fallbackResult: this.findUniqueFallbackMatch(
        testDataRecord,
        direction,
        reportRecords,
        usedReportRowNumbers,
        reservedReportRowNumbers,
      ),
    };
  }

  /** ระบุ BUY/SELL/NO_THB_LEG จาก From Currency และ To Currency */
  private resolveDirection(testDataRecord: ReconcileRecord): FtuDirection {
    const fromCurrency = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.fromCurrency),
    );
    const toCurrency = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.toCurrency),
    );

    if (fromCurrency === FTU_THB_CURRENCY_CODE) {
      return FTU_DIRECTIONS.sellForeignCurrency;
    }

    if (toCurrency === FTU_THB_CURRENCY_CODE) {
      return FTU_DIRECTIONS.buyForeignCurrency;
    }

    /** Return Case อาจไม่มี From/To Currency */
    if (fromCurrency === "" || toCurrency === "") {
      return FTU_DIRECTIONS.unknown;
    }

    return FTU_DIRECTIONS.noThbLeg;
  }

  /** หา AF1 Row จาก Field สำรองเมื่อ Exact Transaction ID ไม่พบ */
  private findUniqueFallbackMatch(
    testDataRecord: ReconcileRecord,
    direction: Exclude<FtuDirection, typeof FTU_DIRECTIONS.unknown>,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    reservedReportRowNumbers: ReadonlySet<number>,
  ): FtuFallbackMatchResult {
    const expectedDate = parseDate(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.transactionDate),
    );
    const expectedPurpose = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.purposeCode),
    );
    const expectedCurrency = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledCurrency),
    );
    const expectedAmount = parseAmount(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledAmount),
    );
    const missingFields: string[] = [];

    if (expectedPurpose === "") {
      missingFields.push(FTU_TEST_DATA_FIELDS.purposeCode);
    }
    if (expectedCurrency === "") {
      missingFields.push(FTU_TEST_DATA_FIELDS.settledCurrency);
    }
    if (expectedAmount === null) {
      missingFields.push(FTU_TEST_DATA_FIELDS.settledAmount);
    }

    if (expectedAmount === null || missingFields.length > 0) {
      return {
        remark:
          "Fallback Matching ทำไม่ได้ เพราะ Field ไม่ครบ: " +
          missingFields.join(", "),
      };
    }

    const expectedLegType =
      direction === FTU_DIRECTIONS.buyForeignCurrency
        ? FTU_LEG_TYPES.buyForeignCurrency
        : direction === FTU_DIRECTIONS.sellForeignCurrency
          ? FTU_LEG_TYPES.sellForeignCurrency
          : undefined;
    const purposeFields =
      direction === FTU_DIRECTIONS.buyForeignCurrency
        ? [FTU_REPORT_FIELDS.inflowPurpose]
        : direction === FTU_DIRECTIONS.sellForeignCurrency
          ? [FTU_REPORT_FIELDS.outflowPurpose]
          : [FTU_REPORT_FIELDS.inflowPurpose, FTU_REPORT_FIELDS.outflowPurpose];

    const candidates = reportRecords.filter((record) => {
      if (
        usedReportRowNumbers.has(record.rowNumber) ||
        reservedReportRowNumbers.has(record.rowNumber)
      ) {
        return false;
      }

      const actualAmount = parseAmount(
        record.get(FTU_REPORT_FIELDS.foreignCurrencyAmount),
      );
      const legMatches =
        expectedLegType === undefined ||
        normalizeFtuText(record.get(FTU_REPORT_FIELDS.legType)) ===
          normalizeFtuText(expectedLegType);
      const purposeMatches = purposeFields.some(
        (field) =>
          normalizeFtuText(record.get(field)) === expectedPurpose,
      );

      return (
        legMatches &&
        purposeMatches &&
        normalizeFtuText(record.get(FTU_REPORT_FIELDS.currencyId)) ===
          expectedCurrency &&
        actualAmount !== null &&
        Math.abs(actualAmount - expectedAmount) <= FTU_AMOUNT_TOLERANCE &&
        (!expectedDate ||
          isDateMatchWithArrangementFallback(
            expectedDate,
            record.get(FTU_REPORT_FIELDS.dataSetDate),
            record.get(FTU_REPORT_FIELDS.arrangementNumber),
          ))
      );
    });

    const matchedFields = [
      expectedDate ? "Date" : "",
      expectedLegType ? "Leg" : "",
      "Purpose",
      "Currency",
      "Amount",
    ]
      .filter((field) => field !== "")
      .join(" + ");
    const missingDateRemark = expectedDate
      ? ""
      : `\nTest Data ไม่มี/อ่าน ${FTU_TEST_DATA_FIELDS.transactionDate} ไม่ได้ ` +
        "จึงไม่ใช้ Date ใน Fallback";

    if (candidates.length === 1) {
      const matchedRecord = candidates[0];
      const arrangementNumber = matchedRecord.get(
        FTU_REPORT_FIELDS.arrangementNumber,
      );

      return {
        matchedRecord,
        candidateCount: 1,
        remark:
          `Fallback Matching: ${matchedFields} ` +
          `ตรงเพียง 1 แถว\n${FTU_REPORT_FIELDS.arrangementNumber} = ` +
          `"${arrangementNumber}"`,
      };
    }

    if (candidates.length === 0) {
      return {
        candidateCount: 0,
        remark:
          `Fallback Matching ไม่พบ AF1 Row ที่ ${matchedFields} ตรงกันครบ` +
          missingDateRemark,
      };
    }

    const references = candidates
      .map(
        (record) =>
          `Row ${record.rowNumber}: ` +
          record.get(FTU_REPORT_FIELDS.arrangementNumber),
      )
      .join(", ");

    return {
      candidateCount: candidates.length,
      remark:
        `Fallback Matching พบหลายแถว (${candidates.length} แถว) ` +
        `จาก ${matchedFields} จึงไม่เลือกอัตโนมัติ\n${references}` +
        missingDateRemark,
    };
  }
}