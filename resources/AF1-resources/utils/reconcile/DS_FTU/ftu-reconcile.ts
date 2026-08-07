/**
 * ftu-reconcile.ts
 * ------------------------------------------------------------------
 * Reconcile Test Data กับ AF1 Report สำหรับ DS_FTU
 *
 * Flow:
 * 1. อ่าน AF1 Report และ Test Data
 * 2. ตรวจ Header ที่จำเป็น
 * 3. ตัดสินว่า Record ต้องมีหรือไม่ต้องมีใน DS_FTU
 * 4. หา AF1 Row ด้วย Transaction ID หรือ Fallback
 * 5. ตรวจ Field หลังจับคู่ Record
 * 6. เขียนผลลง Result Sheet
 *
 * หมายเหตุ:
 * - Arr Number ของ DS_FTU ถือว่าไม่ซ้ำ
 * - ใช้ Transaction ID จับคู่ Arr Number ก่อน
 * - ถ้าหา Arr ไม่พบ ใช้ Leg + Purpose + Currency + Amount เป็น Fallback
 * - ถ้า Test Data มี Date ให้ใช้ Date ช่วยกรอง Candidate เพิ่ม
 * - Fallback ต้องพบ Candidate เพียง 1 แถว
 * - Date ไม่ตรงทั้ง Data Set Date และ Arr Number Date ให้ FAIL
 * - Date ใน Test Data ว่าง/อ่านไม่ได้ให้ Review และตรวจ Field อื่นต่อ
 * - Amount สกุลเดียวกันแต่ต่างเกิน Tolerance ให้ FAIL
 * - USD Amount นอกช่วงใช้ตัดสิน Expected Absence ก่อน Matching
 * - ไม่มีการรวม Amount หลายแถว
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { ReconcileRecord } from "../shared/record";
import { formatCompareRemark } from "../shared/remark";
import { ReconcileResultSheetWriter, ResultRow } from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import {
  FTU_LEG_TYPES,
  FTU_REMARKS,
  FTU_REPORT_CODE,
  FTU_REPORT_FIELDS,
  FTU_REPORT_HEADER_ROW,
  FTU_TEST_DATA_HEADER_ROW,
  FTU_TEST_DATA_FIELDS,
  FTU_THB_CURRENCY_CODE,
  FTU_USD_THRESHOLD,
} from "./ftu-config";
import { normalize, parseAmount } from "./ftu-parse.util";

type FtuDirection = "BUY_FCY" | "SELL_FCY" | "UNKNOWN_DIRECTION" | "NO_THB_LEG";

type AddComparisonRemark = (
  reportField: string,
  testDataField: string,
  expected: string,
  actual: string,
) => void;

interface ResolvedCase {
  result: ResultRow;
  matchedRowNumber?: number;
}

interface FallbackMatchResult {
  matchedRecord?: ReconcileRecord;
  candidateCount?: number;
  remark: string;
}

interface ThresholdPresenceDecision {
  amount: number;
  mustNotExist: boolean;
}

interface MatchedComparisonOptions {
  fallbackRemark?: string;
  initialFailureRemark?: string;
  initialFailedHeaders?: string[];
  skipReportThresholdValidation?: boolean;
}

/** ค่าคลาดเคลื่อนของ Amount ที่ใช้ทั้ง Fallback และ Field Validation */
const FTU_AMOUNT_TOLERANCE = 0.01;

const REQUIRED_TEST_DATA_HEADERS = [
  FTU_TEST_DATA_FIELDS.testNo,
  FTU_TEST_DATA_FIELDS.transactionId,
  FTU_TEST_DATA_FIELDS.transactionDate,
  FTU_TEST_DATA_FIELDS.fromCurrency,
  FTU_TEST_DATA_FIELDS.toCurrency,
  FTU_TEST_DATA_FIELDS.purposeCode,
  FTU_TEST_DATA_FIELDS.settledCurrency,
  FTU_TEST_DATA_FIELDS.settledAmount,
];

/** แปลงข้อความวันที่ที่รองรับเป็น Date แบบ UTC */
const parseDate = (value: unknown): Date | null => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(
      Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()),
    );
  }

  const text = String(value ?? "").trim();
  const dayFirst = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  const yearFirst = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);

  let year: number;
  let month: number;
  let day: number;

  if (dayFirst) {
    day = Number(dayFirst[1]);
    month = Number(dayFirst[2]);
    year = Number(dayFirst[3]);
  } else if (yearFirst) {
    year = Number(yearFirst[1]);
    month = Number(yearFirst[2]);
    day = Number(yearFirst[3]);
  } else {
    return null;
  }

  return createUtcDate(year, month, day);
};

/** สร้าง Date และตรวจว่าเป็นวันที่จริง */
const createUtcDate = (
  year: number,
  month: number,
  day: number,
): Date | null => {
  const date = new Date(Date.UTC(year, month - 1, day));
  const isValid =
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;

  return isValid ? date : null;
};

/** อ่านวันที่ตำแหน่ง 7-12 ของ Arr Number ในรูปแบบ YYMMDD */
const extractDateFromArrangementNumber = (
  arrangementNumber: unknown,
): Date | null => {
  const value = String(arrangementNumber ?? "").trim(); // ถ้าไม่มี Arr number จะเป็นค่าว่าง

  if (value.length < 12) {
    return null;
  }

  const dateText = value.slice(6, 12);

  if (!/^\d{6}$/.test(dateText)) {
    return null;
  }

  const year = 2000 + Number(dateText.slice(0, 2));
  const month = Number(dateText.slice(2, 4));
  const day = Number(dateText.slice(4, 6));

  return createUtcDate(year, month, day);
};

/** เปรียบเทียบปี เดือน และวัน */
const isSameDate = (left: Date, right: Date): boolean =>
  left.getUTCFullYear() === right.getUTCFullYear() &&
  left.getUTCMonth() === right.getUTCMonth() &&
  left.getUTCDate() === right.getUTCDate();

const formatDate = (value: Date | null): string => {
  if (!value) {
    return "Invalid date";
  }

  const year = value.getUTCFullYear();
  const month = String(value.getUTCMonth() + 1).padStart(2, "0");
  const day = String(value.getUTCDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
};

/**
 * TODO: FTX Exception
 *
 * Requirement ระบุว่าบางรายการต้องรายงานใน DS_FTX แทน DS_FTU
 * แต่ยังไม่มี Field และ Logic ครบถ้วน จึงยังไม่เปิดใช้ Rule นี้
 */

export class FtuReconcileService {
  async reconcile(testDataFilePath: string): Promise<string> {
    const workbookPreparer = new ReconcileWorkbookPreparer();
    const excelReader = new ReconcileExcelReader();
    const sheetWriter = new ReconcileResultSheetWriter();

    console.log(`\n===== RECONCILE - ${FTU_REPORT_CODE} =====`);

    const prepared = await workbookPreparer.prepare(
      FTU_REPORT_CODE,
      FTU_REPORT_HEADER_ROW,
    );

    const reportData = excelReader.parseWorksheet(
      prepared.reportWorksheet,
      FTU_REPORT_HEADER_ROW,
    );

    const testData = await excelReader.readFile(
      testDataFilePath,
      FTU_TEST_DATA_HEADER_ROW,
    );

    this.validateHeaders(prepared.reportHeaders, testData.headers);

    const reportRecordsById = this.indexReportRecords(reportData.records);

    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const results: ResultRow[] = [];

    /** กัน Fallback ใช้ AF1 Row ที่มี Transaction ID จับคู่ไว้แล้ว */
    const usedReportRowNumbers = new Set<number>();

    for (const testDataRecord of testData.records) {
      const reservedTransactionId = normalize(
        testDataRecord.get(FTU_TEST_DATA_FIELDS.transactionId),
      );

      if (reservedTransactionId === "") {
        continue;
      }

      const reservedRecord = reportRecordsById.get(reservedTransactionId);

      if (reservedRecord) {
        usedReportRowNumbers.add(reservedRecord.rowNumber);
      }
    }

    for (const testDataRecord of testData.records) {
      /** Matching ข้อ 1: Transaction ID/Reconcile ID -> Arr Number */
      const transactionId = normalize(
        testDataRecord.get(FTU_TEST_DATA_FIELDS.transactionId),
      );

      const matchedRecord = reportRecordsById.get(transactionId);

      const resolvedCase = this.resolveRecord(
        testDataRecord,
        transactionId,
        matchedRecord,
        reportData.records,
        usedReportRowNumbers,
      );

      /** ไม่มี Test No. ให้แสดงเลขแถวต้นทางและบันทึก Comment */
      if (normalize(testDataRecord.get(FTU_TEST_DATA_FIELDS.testNo)) === "") {
        resolvedCase.result.remark = this.appendRemark(
          resolvedCase.result.remark,
          `Test Data ไม่มี ${FTU_TEST_DATA_FIELDS.testNo}`,
        );
      }

      /** Transaction ID ว่างต้องมี Comment เสมอ */
      if (transactionId === "") {
        resolvedCase.result.remark = this.appendRemark(
          resolvedCase.result.remark,
          `Test Data ไม่มี ${FTU_TEST_DATA_FIELDS.transactionId}`,
        );
      }

      results.push(resolvedCase.result);

      if (resolvedCase.matchedRowNumber === undefined) {
        unmatchedRows.push(resolvedCase.result);
        continue;
      }

      annotationByRowNumber.set(
        resolvedCase.matchedRowNumber,
        resolvedCase.result,
      );

      usedReportRowNumbers.add(resolvedCase.matchedRowNumber);
    }

    sheetWriter.writeHeaderRow(prepared.resultSheet, prepared.reportHeaders);

    const nextRowNumber = sheetWriter.writeRowsInRequestedOrder(
      prepared.resultSheet,
      prepared.reportWorksheet,
      prepared.reportHeaders,
      FTU_REPORT_HEADER_ROW + 1,
      prepared.reportWorksheet.rowCount,
      annotationByRowNumber,
      unmatchedRows,
    );

    sheetWriter.finalizeAutoFilter(
      prepared.resultSheet,
      prepared.reportHeaders,
      nextRowNumber - 1,
    );

    prepared.workbook.removeWorksheet(prepared.reportWorksheet.id);
    await prepared.workbook.xlsx.writeFile(prepared.reconcileFilePath);

    this.logSummary(prepared.reconcileFilePath, results);
    return prepared.reconcileFilePath;
  }

  /** ตรวจ Header ที่จำเป็นก่อนเริ่ม Reconcile */
  private validateHeaders(
    reportHeaders: string[],
    testDataHeaders: string[],
  ): void {
    const reportName = requireMappingReportName(FTU_REPORT_CODE);
    const requiredReportHeaders = getUniqueMappingHeaders(reportName);

    this.assertHeaders(reportHeaders, requiredReportHeaders, "Raw Report");
    this.assertHeaders(
      testDataHeaders,
      REQUIRED_TEST_DATA_HEADERS,
      "Test Data",
    );
  }

  private assertHeaders(
    actualHeaders: string[],
    requiredHeaders: readonly string[],
    sourceName: string,
  ): void {
    const actualHeaderSet = new Set(
      actualHeaders
        .filter((header) => header.trim() !== "")
        .map(canonicalHeader),
    );

    const missingHeaders = requiredHeaders.filter(
      (header) => !actualHeaderSet.has(canonicalHeader(header)),
    );

    if (missingHeaders.length > 0) {
      throw new Error(
        `[${FTU_REPORT_CODE}] ` +
          `${sourceName} missing header(s): ` +
          missingHeaders.join(", "),
      );
    }
  }

  /** ต่อ Remark โดยไม่เพิ่มข้อความซ้ำ */
  private appendRemark(current: string, next: string): string {
    if (next.trim() === "" || current.includes(next)) {
      return current;
    }

    return [current, next]
      .filter((message) => message.trim() !== "")
      .join("\n");
  }

  /**
   * สร้าง Index จาก Arr Number
   */
  private indexReportRecords(
    records: ReconcileRecord[],
  ): Map<string, ReconcileRecord> {
    const recordsById = new Map<string, ReconcileRecord>();

    for (const record of records) {
      const arrangementNumber = normalize(
        record.get(FTU_REPORT_FIELDS.arrangementNumber),
      );

      if (arrangementNumber === "") {
        continue;
      }

      recordsById.set(arrangementNumber, record);
    }

    return recordsById;
  }

  /** ตัดสิน Presence, จับคู่ Record และตรวจ Test Data หนึ่งแถว */
  private resolveRecord(
    testDataRecord: ReconcileRecord,
    transactionId: string,
    matchedRecord: ReconcileRecord | undefined,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
  ): ResolvedCase {
    const testCaseNo =
      normalize(testDataRecord.get(FTU_TEST_DATA_FIELDS.testNo)) ||
      `TEST DATA ROW ${testDataRecord.rowNumber}`;

    const direction = this.resolveDirection(testDataRecord);

    if (direction === "NO_THB_LEG") {
      return this.resolveNoThbLeg(
        testCaseNo,
        testDataRecord,
        transactionId,
        matchedRecord,
        reportRecords,
        usedReportRowNumbers,
      );
    }

    /**
     * USD Amount ที่ไม่เข้าเกณฑ์ต้องไม่มีใน DS_FTU
     * จึงตัดสิน Expected Absence ก่อนตัดสินว่า Arr Number หายเป็น FAIL
     */
    const thresholdDecision =
      this.resolveThresholdPresenceDecision(testDataRecord);

    if (thresholdDecision?.mustNotExist) {
      return this.resolveThresholdExpectedAbsence(
        testCaseNo,
        testDataRecord,
        matchedRecord,
        direction,
        reportRecords,
        usedReportRowNumbers,
        thresholdDecision.amount,
      );
    }

    let resolvedMatchedRecord = matchedRecord;
    let fallbackRemark = "";

    /**
     * ถ้า Matching Key หลักหา Arr Number ไม่พบ ให้ลองจับคู่จาก Field อื่นต่อ
     * ใช้ได้ทั้งกรณี Transaction ID ว่างและมีค่าแต่ Report ไม่มี Arr ที่ตรงกัน
     */
    if (!resolvedMatchedRecord) {
      if (direction === "UNKNOWN_DIRECTION") {
        fallbackRemark =
          "Fallback Matching ทำไม่ได้ เพราะ From/To Currency ไม่ครบ";
      } else {
        const fallbackResult = this.findUniqueFallbackMatch(
          testDataRecord,
          direction,
          reportRecords,
          usedReportRowNumbers,
        );

        resolvedMatchedRecord = fallbackResult.matchedRecord;
        fallbackRemark =
          transactionId === ""
            ? fallbackResult.remark
            : resolvedMatchedRecord
              ? this.appendRemark(
                  `ควรพบ ${FTU_REPORT_FIELDS.arrangementNumber} = ` +
                    `"${transactionId}" แต่ Matching Key หลักไม่พบ`,
                  fallbackResult.remark,
                )
              : fallbackResult.remark;
      }
    }

    if (!resolvedMatchedRecord) {
      return {
        result: {
          testCaseNo,
          status: "FAIL",
          remark: this.appendRemark(
            formatCompareRemark(
              FTU_REPORT_CODE,
              FTU_TEST_DATA_FIELDS.transactionId,
              transactionId,
              FTU_REPORT_FIELDS.arrangementNumber,
              "ไม่พบข้อมูล",
            ),
            fallbackRemark,
          ),
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
          reviewFieldHeaders: [],
          isExpectedAbsence: false,
        },
      };
    }

    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      resolvedMatchedRecord,
      direction,
      {
        fallbackRemark,
      },
    );
  }

  /** ระบุ BUY/SELL จาก From Currency และ To Currency
   * และแยกกรณี Currency ไม่ครบหรือไม่มีขา THB */
  private resolveDirection(testDataRecord: ReconcileRecord): FtuDirection {
    const fromCurrency = normalize(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.fromCurrency),
    );

    const toCurrency = normalize(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.toCurrency),
    );

    if (fromCurrency === FTU_THB_CURRENCY_CODE) {
      return "SELL_FCY";
    }

    if (toCurrency === FTU_THB_CURRENCY_CODE) {
      return "BUY_FCY";
    }

    /**
     * Return Case อาจไม่มี From/To Currency
     * จึงไม่ใช้ตัดสินว่าเป็นรายการไม่มีขา THB
     */
    if (fromCurrency === "" || toCurrency === "") {
      return "UNKNOWN_DIRECTION";
    }

    return "NO_THB_LEG";
  }

  /** แสดงคู่ From/To Currency เพื่ออธิบายเหตุผลของ NO_THB_LEG */
  private formatTestDataCurrencyPairRemark(
    testDataRecord: ReconcileRecord,
  ): string {
    const fromCurrency = testDataRecord
      .get(FTU_TEST_DATA_FIELDS.fromCurrency)
      .trim();
    const toCurrency = testDataRecord
      .get(FTU_TEST_DATA_FIELDS.toCurrency)
      .trim();

    return (
      `[TS] : ${FTU_TEST_DATA_FIELDS.fromCurrency}/` +
      `${FTU_TEST_DATA_FIELDS.toCurrency} = ` +
      `"${fromCurrency}/${toCurrency}"`
    );
  }

  /** หา AF1 Row จาก Field สำรองเมื่อ Matching Key หลักหา Arr Number ไม่พบ */
  private findUniqueFallbackMatch(
    testDataRecord: ReconcileRecord,
    direction: Exclude<FtuDirection, "UNKNOWN_DIRECTION">,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
  ): FallbackMatchResult {
    const expectedDate = parseDate(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.transactionDate),
    );
    const expectedPurpose = normalize(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.purposeCode),
    );
    const expectedCurrency = normalize(
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
      direction === "BUY_FCY"
        ? FTU_LEG_TYPES.buyForeignCurrency
        : direction === "SELL_FCY"
          ? FTU_LEG_TYPES.sellForeignCurrency
          : undefined;
    const purposeFields =
      direction === "BUY_FCY"
        ? [FTU_REPORT_FIELDS.inflowPurpose]
        : direction === "SELL_FCY"
          ? [FTU_REPORT_FIELDS.outflowPurpose]
          : [FTU_REPORT_FIELDS.inflowPurpose, FTU_REPORT_FIELDS.outflowPurpose];

    const candidates = reportRecords.filter((record) => {
      if (usedReportRowNumbers.has(record.rowNumber)) {
        return false;
      }

      const actualAmount = parseAmount(
        record.get(FTU_REPORT_FIELDS.foreignCurrencyAmount),
      );
      const legMatches =
        expectedLegType === undefined ||
        normalize(record.get(FTU_REPORT_FIELDS.legType)) ===
          normalize(expectedLegType);
      const purposeMatches = purposeFields.some(
        (field) => normalize(record.get(field)) === expectedPurpose,
      );

      return (
        legMatches &&
        purposeMatches &&
        normalize(record.get(FTU_REPORT_FIELDS.currencyId)) ===
          expectedCurrency &&
        actualAmount !== null &&
        Math.abs(actualAmount - expectedAmount) <= FTU_AMOUNT_TOLERANCE &&
        (!expectedDate || this.isReportDateMatch(expectedDate, record))
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

  /** Helper กลาง: วันที่ต้องตรงกับ Data Set Date หรือวันที่ใน Arr Number */
  private isReportDateMatch(
    expectedDate: Date,
    reportRecord: ReconcileRecord,
  ): boolean {
    const dataSetDate = parseDate(
      reportRecord.get(FTU_REPORT_FIELDS.dataSetDate),
    );

    if (dataSetDate && isSameDate(expectedDate, dataSetDate)) {
      return true;
    }

    const arrangementDate = extractDateFromArrangementNumber(
      reportRecord.get(FTU_REPORT_FIELDS.arrangementNumber),
    );

    return Boolean(
      arrangementDate && isSameDate(expectedDate, arrangementDate),
    );
  }

  /** สูตรกลางสำหรับตรวจช่วง Amount ของ DS_FTU */
  private isWithinUsdThreshold(amount: number): boolean {
    return amount > 0 && amount < FTU_USD_THRESHOLD;
  }

  /** ใช้ Test Data ตัดสินเกณฑ์ Presence สำหรับ Amount สกุล USD */
  private resolveThresholdPresenceDecision(
    testDataRecord: ReconcileRecord,
  ): ThresholdPresenceDecision | undefined {
    const currency = normalize(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledCurrency),
    );
    const amount = parseAmount(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledAmount),
    );

    /** สกุลอื่นยังไม่มี USD Equivalent/Exchange Rate ที่ยืนยันแล้ว */
    if (currency !== "USD" || amount === null) {
      return undefined;
    }

    return {
      amount,
      mustNotExist: !this.isWithinUsdThreshold(amount),
    };
  }

  /** USD Amount นอกเกณฑ์ต้องไม่พบ Record ใน DS_FTU */
  private resolveThresholdExpectedAbsence(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord | undefined,
    direction: Exclude<FtuDirection, "NO_THB_LEG">,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
    amount: number,
  ): ResolvedCase {
    const expectationRemark =
      `Test Data Amount = ${amount} USD ไม่เข้าเกณฑ์ ` +
      `0 < Amount < ${FTU_USD_THRESHOLD} USD ` +
      `จึงไม่ควรพบรายการใน ${FTU_REPORT_CODE}`;

    let detectedRecord = matchedRecord;
    let detectionRemark = "";
    let fallbackVerificationUnavailable = false;

    /** Exact Match ไม่พบ ให้ใช้ Fallback ตรวจหารายการที่ไม่ควรมีด้วย */
    if (!detectedRecord) {
      if (direction === "UNKNOWN_DIRECTION") {
        detectionRemark = "ตรวจ Fallback ไม่ได้ เพราะ From/To Currency ไม่ครบ";
        fallbackVerificationUnavailable = true;
      } else {
        const fallbackResult = this.findUniqueFallbackMatch(
          testDataRecord,
          direction,
          reportRecords,
          usedReportRowNumbers,
        );

        detectedRecord = fallbackResult.matchedRecord;
        detectionRemark = fallbackResult.remark;
        fallbackVerificationUnavailable =
          fallbackResult.candidateCount === undefined;

        /** พบหลาย Candidate แปลว่าพบรายการที่ไม่ควรมี แม้เลือกแถวเดียวไม่ได้ */
        if (
          !detectedRecord &&
          fallbackResult.candidateCount !== undefined &&
          fallbackResult.candidateCount > 1
        ) {
          return {
            result: {
              testCaseNo,
              status: "FAIL",
              remark: this.appendRemark(
                `${expectationRemark}\nแต่ Fallback พบรายการใน Report`,
                detectionRemark,
              ),
              matchedRowNumber: undefined,
              failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
              reviewFieldHeaders: [],
              isExpectedAbsence: false,
            },
          };
        }
      }
    }

    if (!detectedRecord) {
      const verificationRemark = fallbackVerificationUnavailable
        ? "ไม่สามารถยืนยัน Expected Absence ด้วย Fallback ได้ครบถ้วน"
        : "ไม่พบรายการใน Report ตามที่คาดหวัง";

      return {
        result: {
          testCaseNo,
          status: "PASS",
          remark: this.appendRemark(
            `${expectationRemark}\n${verificationRemark}`,
            detectionRemark,
          ),
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: fallbackVerificationUnavailable
            ? [FTU_REPORT_FIELDS.arrangementNumber]
            : [],
          isExpectedAbsence: true,
        },
      };
    }

    const arrangementNumber = detectedRecord.get(
      FTU_REPORT_FIELDS.arrangementNumber,
    );

    /** พบ Record ที่ไม่ควรมี: บังคับ FAIL แต่ยังตรวจ Field อื่นเพื่อเก็บ Remark */
    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      detectedRecord,
      direction,
      {
        initialFailureRemark: this.appendRemark(
          `${expectationRemark}\nแต่พบรายการใน Report: ` +
            `${FTU_REPORT_FIELDS.arrangementNumber} = ` +
            `"${arrangementNumber}"`,
          detectionRemark,
        ),
        initialFailedHeaders: [
          FTU_REPORT_FIELDS.arrangementNumber,
          FTU_REPORT_FIELDS.foreignCurrencyAmount,
        ],
        skipReportThresholdValidation: true,
      },
    );
  }

  /** ไม่มีขา THB จึงไม่ควรพบ Record ใน DS_FTU */
  private resolveNoThbLeg(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    transactionId: string,
    matchedRecord: ReconcileRecord | undefined,
    reportRecords: ReconcileRecord[],
    usedReportRowNumbers: ReadonlySet<number>,
  ): ResolvedCase {
    let detectedRecord = matchedRecord;
    let detectionRemark = "";
    let fallbackVerificationUnavailable = false;
    const testDataCurrencyPairRemark =
      this.formatTestDataCurrencyPairRemark(testDataRecord);

    /** Exact Match ไม่พบ ให้ค้นหา Unexpected Presence ด้วย Field สำรอง */
    if (!detectedRecord) {
      const fallbackResult = this.findUniqueFallbackMatch(
        testDataRecord,
        "NO_THB_LEG",
        reportRecords,
        usedReportRowNumbers,
      );

      detectedRecord = fallbackResult.matchedRecord;
      detectionRemark = fallbackResult.remark;
      fallbackVerificationUnavailable =
        fallbackResult.candidateCount === undefined;

      if (
        !detectedRecord &&
        fallbackResult.candidateCount !== undefined &&
        fallbackResult.candidateCount > 1
      ) {
        return {
          result: {
            testCaseNo,
            status: "FAIL",
            remark: this.appendRemark(
              `${FTU_REMARKS.noThbLegUnexpectedPresence}\n` +
                testDataCurrencyPairRemark,
              detectionRemark,
            ),
            matchedRowNumber: undefined,
            failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
            reviewFieldHeaders: [],
            isExpectedAbsence: false,
          },
        };
      }
    }

    if (!detectedRecord) {
      const verificationRemark = fallbackVerificationUnavailable
        ? "ไม่สามารถยืนยัน Expected Absence ด้วย Fallback ได้ครบถ้วน"
        : "ไม่พบรายการใน DS_FTU ตามที่คาดหวัง";

      /**
       * ไม่ใส่ข้อความ "Fallback Matching ไม่พบ..." เมื่อค้นหาแล้วไม่พบ Candidate
       * เพราะ NO_THB_LEG ต้องไม่พบใน DS_FTU อยู่แล้ว และถือเป็น Expected Absence
       */
      const fallbackReviewRemark = fallbackVerificationUnavailable
        ? detectionRemark
        : "";

      return {
        result: {
          testCaseNo,
          status: "PASS",
          remark: this.appendRemark(
            `${FTU_REMARKS.noThbLegExpectedAbsence}\n` +
              `${verificationRemark}\n${testDataCurrencyPairRemark}`,
            fallbackReviewRemark,
          ),
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: fallbackVerificationUnavailable
            ? [FTU_REPORT_FIELDS.arrangementNumber]
            : [],
          isExpectedAbsence: true,
        },
      };
    }

    const keyRemark =
      matchedRecord && transactionId !== ""
        ? formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_DATA_FIELDS.transactionId,
            transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            detectedRecord.get(FTU_REPORT_FIELDS.arrangementNumber),
          )
        : detectionRemark;

    /** พบ Record ที่ไม่ควรมี: บังคับ FAIL แต่ยังตรวจ Field อื่นเพื่อเก็บ Remark */
    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      detectedRecord,
      "NO_THB_LEG",
      {
        initialFailureRemark: this.appendRemark(
          FTU_REMARKS.noThbLegUnexpectedPresence,
          keyRemark,
        ),
        initialFailedHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
        skipReportThresholdValidation: true,
      },
    );
  }

  /** เปรียบเทียบ Field เมื่อพบ Record ที่จับคู่กัน */
  private compareMatchedRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: FtuDirection,
    options: MatchedComparisonOptions = {},
  ): ResolvedCase {
    const failedHeaders = new Set<string>(options.initialFailedHeaders ?? []);
    const reviewHeaders = new Set<string>();
    const remarks: string[] = options.initialFailureRemark
      ? [options.initialFailureRemark]
      : [];

    const addFailure: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      failedHeaders.add(reportField);
      remarks.push(
        formatCompareRemark(
          FTU_REPORT_CODE,
          testDataField,
          expected,
          reportField,
          actual,
        ),
      );
    };

    /**
     * ใช้สำหรับข้อมูลที่ต้องแสดงเพื่อ Review
     * แต่ไม่ใช้ตัดสิน PASS/FAIL
     */
    const addReview: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      reviewHeaders.add(reportField);
      remarks.push(
        formatCompareRemark(
          FTU_REPORT_CODE,
          testDataField,
          expected,
          reportField,
          actual,
        ),
      );
    };

    /** Fallback ที่ Map สำเร็จให้ Highlight Arr Number เพื่อ Review */
    if (options.fallbackRemark) {
      reviewHeaders.add(FTU_REPORT_FIELDS.arrangementNumber);
      remarks.push(options.fallbackRemark);
    }

    /** Date ว่างให้ Review; มีค่าแต่ไม่ตรงทั้ง 2 แหล่งให้ FAIL */
    this.compareDate(testDataRecord, matchedRecord, addFailure, addReview);

    /** Matching ข้อ 4: From/To Currency -> Leg Type
     *  หากระบุ Direction ไม่ได้ ให้ Review เท่านั้น */
    this.compareLegType(
      testDataRecord,
      matchedRecord,
      direction,
      addFailure,
      addReview,
    );

    /** Matching ข้อ 5: BOT Purpose -> Inflow/Outflow Purpose (Review) */
    this.comparePurpose(testDataRecord, matchedRecord, direction, addReview);

    /** Matching ข้อ 6: To Currency 2 ตัวแรก -> Country Id (Review) */
    this.compareCountry(testDataRecord, matchedRecord, addReview);

    /** Matching ข้อ 7: Settled Currency -> Currency Id */
    this.compareCurrency(testDataRecord, matchedRecord, addFailure);

    /** Currency เดียวกันแต่ Amount ต่างเกิน Tolerance ให้ FAIL */
    this.compareAmount(testDataRecord, matchedRecord, addFailure);

    /** Expected Absence ตรวจ Threshold จาก Test Data ไปแล้ว จึงไม่เพิ่ม Remark ซ้ำ */
    if (!options.skipReportThresholdValidation) {
      this.evaluateAmountThreshold(matchedRecord, addFailure, addReview);
    }

    const status = failedHeaders.size === 0 ? "PASS" : "FAIL";

    const successRemark =
      direction === "BUY_FCY"
        ? FTU_REMARKS.buyForeignCurrency
        : direction === "SELL_FCY"
          ? FTU_REMARKS.sellForeignCurrency
          : "";
    /**
     * ถ้ามี Field ที่ต้อง Review
     * ให้เพิ่ม Please review เป็นบรรทัดสุดท้าย
     */
    const pleaseReviewRemark =
      reviewHeaders.size > 0 ? FTU_REMARKS.pleaseReview : "";

    const finalRemark = [
      status === "PASS" ? successRemark : "",
      ...remarks,
      pleaseReviewRemark,
    ]
      .map((message) => message.trim())
      .filter((message) => message !== "")
      .join("\n");

    return {
      result: {
        testCaseNo,
        status,
        remark: finalRemark,
        matchedRowNumber: matchedRecord.rowNumber,
        failedKeyFieldHeaders: [...failedHeaders],
        reviewFieldHeaders: [...reviewHeaders],
        isExpectedAbsence: false,
      },
      matchedRowNumber: matchedRecord.rowNumber,
    };
  }

  /**
   * Matching ข้อ 3: ตรวจ Txn Date กับ Data Set Date ก่อน
   *
   * หากไม่ตรงกัน ให้ดึงวันที่ตำแหน่ง 7-12 ของ Arr Number มาเทียบอีกครั้ง
   * - Test Data Date ว่าง/อ่านไม่ได้: Review และตรวจ Field อื่นต่อ
   * - Date มีค่าแต่ไม่ตรงทั้งสองแหล่ง: FAIL
   */
  private compareDate(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const expectedText = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.transactionDate,
    );

    const actualText = matchedRecord.get(FTU_REPORT_FIELDS.dataSetDate);

    const expectedDate = parseDate(expectedText);
    if (!expectedDate) {
      addReview(
        FTU_REPORT_FIELDS.dataSetDate,
        FTU_TEST_DATA_FIELDS.transactionDate,
        expectedText,
        actualText,
      );
      return;
    }

    if (this.isReportDateMatch(expectedDate, matchedRecord)) {
      return;
    }

    const arrangementNumber = matchedRecord.get(
      FTU_REPORT_FIELDS.arrangementNumber,
    );
    const arrangementDate = extractDateFromArrangementNumber(arrangementNumber);

    addFailure(
      FTU_REPORT_FIELDS.dataSetDate,
      FTU_TEST_DATA_FIELDS.transactionDate,
      expectedText,
      `${actualText} | ` + `Arr Number Date: ` + formatDate(arrangementDate),
    );
  }

  /** Matching ข้อ 4: BUY/SELL ต้องตรงกับ Leg Type */
  private compareLegType(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: FtuDirection,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const actualLegType = matchedRecord.get(FTU_REPORT_FIELDS.legType);
    /** Direction ไม่ชัดหรือไม่มีขา THB ให้แสดง Leg Type เพื่อวิเคราะห์เท่านั้น */
    if (direction === "UNKNOWN_DIRECTION" || direction === "NO_THB_LEG") {
      const fromCurrency = testDataRecord.get(
        FTU_TEST_DATA_FIELDS.fromCurrency,
      );
      const toCurrency = testDataRecord.get(FTU_TEST_DATA_FIELDS.toCurrency);
      const expectedDirection = `${fromCurrency}/${toCurrency}`;

      addReview(
        FTU_REPORT_FIELDS.legType,
        `${FTU_TEST_DATA_FIELDS.fromCurrency}/` +
          FTU_TEST_DATA_FIELDS.toCurrency,
        expectedDirection,
        actualLegType,
      );
      return;
    }
    const expectedLegType =
      direction === "BUY_FCY"
        ? FTU_LEG_TYPES.buyForeignCurrency
        : FTU_LEG_TYPES.sellForeignCurrency;

    if (normalize(actualLegType) !== normalize(expectedLegType)) {
      addFailure(
        FTU_REPORT_FIELDS.legType,
        `${FTU_TEST_DATA_FIELDS.fromCurrency}/` +
          FTU_TEST_DATA_FIELDS.toCurrency,
        expectedLegType,
        actualLegType,
      );
    }
  }

  /** Matching ข้อ 5: Purpose ไม่ตรงให้ Review โดยไม่ตัดสิน FAIL */
  private comparePurpose(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: FtuDirection,
    addReview: AddComparisonRemark,
  ): void {
    if (direction === "UNKNOWN_DIRECTION" || direction === "NO_THB_LEG") {
      return;
    }
    const expectedPurpose = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.purposeCode,
    );
    const reportPurposeField =
      direction === "BUY_FCY"
        ? FTU_REPORT_FIELDS.inflowPurpose
        : FTU_REPORT_FIELDS.outflowPurpose;
    const actualPurpose = matchedRecord.get(reportPurposeField);

    if (normalize(expectedPurpose) !== normalize(actualPurpose)) {
      addReview(
        reportPurposeField,
        FTU_TEST_DATA_FIELDS.purposeCode,
        expectedPurpose,
        actualPurpose,
      );
    }
  }

  /**
   * Matching ข้อ 6: เทียบ Country ระหว่าง Test Data กับ AF1
   *
   * Test Data: 2 ตัวแรกของ To Currency (CCY)
   * AF1: Country Id of Beneficiary Involved Party
   *
   * ตัวอย่าง: AUD -> AU, THB -> TH, USD -> US
   * หากไม่ตรงกันให้ Review โดยไม่ตัดสิน FAIL
   */
  private compareCountry(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const toCurrency = normalize(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.toCurrency),
    );

    const expectedCountryFromToCurrency = toCurrency.slice(0, 2);

    const actualCountry = matchedRecord.get(
      FTU_REPORT_FIELDS.beneficiaryCountry,
    );

    if (expectedCountryFromToCurrency === "") {
      addReview(
        FTU_REPORT_FIELDS.beneficiaryCountry,
        FTU_TEST_DATA_FIELDS.toCurrency,
        expectedCountryFromToCurrency,
        actualCountry,
      );
      return;
    }

    if (normalize(actualCountry) !== expectedCountryFromToCurrency) {
      addReview(
        FTU_REPORT_FIELDS.beneficiaryCountry,
        FTU_TEST_DATA_FIELDS.toCurrency,
        expectedCountryFromToCurrency,
        actualCountry,
      );
    }
  }

  /** Matching ข้อ 7: Settled Currency ต้องตรงกับ Currency Id */
  private compareCurrency(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
  ): void {
    const expectedCurrency = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.settledCurrency,
    );
    const actualCurrency = matchedRecord.get(FTU_REPORT_FIELDS.currencyId);

    if (normalize(expectedCurrency) !== normalize(actualCurrency)) {
      addFailure(
        FTU_REPORT_FIELDS.currencyId,
        FTU_TEST_DATA_FIELDS.settledCurrency,
        expectedCurrency,
        actualCurrency,
      );
    }
  }

  /**
   * Matching ข้อ 8: Settled Amount กับ Foreign Currency Amount
   *
   * ตรวจเฉพาะเมื่อ Currency ของทั้งสองฝั่งตรงกัน
   * หาก Amount ต่างเกิน Tolerance ให้ FAIL และไม่มีการรวมหลาย Record
   */
  private compareAmount(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
  ): void {
    const expectedCurrency = normalize(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledCurrency),
    );
    const actualCurrency = normalize(
      matchedRecord.get(FTU_REPORT_FIELDS.currencyId),
    );

    /** Currency ไม่ตรงถูกตัดสิน FAIL ใน compareCurrency() แล้ว */
    if (
      expectedCurrency === "" ||
      actualCurrency === "" ||
      expectedCurrency !== actualCurrency
    ) {
      return;
    }

    const expectedText = testDataRecord.get(FTU_TEST_DATA_FIELDS.settledAmount);

    const actualText = matchedRecord.get(
      FTU_REPORT_FIELDS.foreignCurrencyAmount,
    );

    const expectedAmount = parseAmount(expectedText);
    const actualAmount = parseAmount(actualText);

    if (
      expectedAmount === null ||
      actualAmount === null ||
      Math.abs(expectedAmount - actualAmount) > FTU_AMOUNT_TOLERANCE
    ) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_TEST_DATA_FIELDS.settledAmount,
        expectedText,
        actualText,
      );
    }
  }

  /**
   * ตรวจ 0 < FTU Amount < 50,000 USD
   *
   * ขณะนี้ตรวจได้โดยตรงเฉพาะ Currency Id = USD
   * สกุลเงินอื่นรอแหล่ง Exchange Rate หรือ USD Equivalent
   * ที่ยืนยันจาก Requirement
   */
  private evaluateAmountThreshold(
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const amountText = matchedRecord.get(
      FTU_REPORT_FIELDS.foreignCurrencyAmount,
    );

    const currencyId = matchedRecord.get(FTU_REPORT_FIELDS.currencyId);

    const amount = parseAmount(amountText);
    const normalizedCurrency = normalize(currencyId);

    if (amount === null) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        "FTU Amount Threshold",
        `0 < Amount < ${FTU_USD_THRESHOLD} USD`,
        amountText,
      );
      return;
    }

    if (normalizedCurrency !== "USD") {
      /**
       * TODO: แปลง Amount เป็น USD เมื่อได้รับ
       * Exchange Rate หรือ USD Equivalent Field
       */
      addReview(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        "FTU Amount Threshold",
        `0 < USD Equivalent < ` + `${FTU_USD_THRESHOLD}`,
        `${amount} ${currencyId} ` + "(รอข้อมูลแปลงเป็น USD)",
      );
      return;
    }

    const isWithinThreshold = this.isWithinUsdThreshold(amount);

    if (!isWithinThreshold) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        "FTU Amount Threshold",
        `0 < Amount < ${FTU_USD_THRESHOLD} USD`,
        `${amount} USD`,
      );
    }
  }

  private logSummary(outputPath: string, results: ResultRow[]): void {
    const passCount = results.filter(
      (result) => result.status === "PASS",
    ).length;

    const failCount = results.length - passCount;

    console.log(`Output File : ${outputPath}`);
    console.log(
      `Test Case : ${results.length} | ` +
        `Pass : ${passCount} | ` +
        `Fail : ${failCount}`,
    );
  }
}

export const reconcileFtuReport = (testDataFilePath: string): Promise<string> =>
  new FtuReconcileService().reconcile(testDataFilePath);