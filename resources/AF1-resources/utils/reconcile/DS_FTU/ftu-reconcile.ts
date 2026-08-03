/**
 * ftu-reconcile.ts
 * ------------------------------------------------------------------
 * Reconcile Test Data กับ AF1 Report สำหรับ DS_FTU
 *
 * Flow:
 * 1. อ่าน AF1 Report และ Test Data
 * 2. ตรวจ Header ที่จำเป็น
 * 3. ตรวจ Matching Field
 * 4. ตรวจช่วงมูลค่า FTU เมื่อ Amount เป็น USD
 * 5. เขียนผลลง Result Sheet
 *
 * หมายเหตุ:
 * - Arr Number ของ DS_FTU ถือว่าไม่ซ้ำ
 * - Amount ระหว่าง Test Data กับ AF1 ใช้เพื่อ Review เท่านั้น
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
const isSameDate = (
  left: Date,
  right: Date,
): boolean =>
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
      );

      results.push(resolvedCase.result);

      if (resolvedCase.matchedRowNumber === undefined) {
        unmatchedRows.push(resolvedCase.result);
        continue;
      }

      annotationByRowNumber.set(
        resolvedCase.matchedRowNumber,
        resolvedCase.result,
      );
    }

    sheetWriter.writeHeaderRow(
      prepared.resultSheet,
      prepared.reportHeaders,
    );

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

  /** ตรวจข้อมูลหลังจากจับคู่ Transaction ID กับ Arr Number แล้ว */
  private resolveRecord(
    testDataRecord: ReconcileRecord,
    transactionId: string,
    matchedRecord: ReconcileRecord | undefined,
  ): ResolvedCase {
    const testCaseNo =
      normalize(testDataRecord.get(FTU_TEST_DATA_FIELDS.testNo)) ||
      `Row ${testDataRecord.rowNumber}`;

    const direction = this.resolveDirection(testDataRecord);

    if (direction === "NO_THB_LEG") {
      return this.resolveNoThbLeg(testCaseNo, transactionId, matchedRecord);
    }

    if (!matchedRecord) {
      return {
        result: {
          testCaseNo,
          status: "FAIL",
          remark: formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_DATA_FIELDS.transactionId,
            transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            "ไม่พบข้อมูล",
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
      matchedRecord,
      direction,
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

  /** ไม่มีขา THB จึงไม่ควรพบ Record ใน DS_FTU */
  private resolveNoThbLeg(
    testCaseNo: string,
    transactionId: string,
    matchedRecord: ReconcileRecord | undefined,
  ): ResolvedCase {
    if (!matchedRecord) {
      return {
        result: {
          testCaseNo,
          status: "PASS",
          remark: FTU_REMARKS.noThbLegExpectedAbsence,
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: [],
          isExpectedAbsence: true,
        },
      };
    }

    return {
      result: {
        testCaseNo,
        status: "FAIL",
        remark:
          `${FTU_REMARKS.noThbLegUnexpectedPresence}\n` +
          formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_DATA_FIELDS.transactionId,
            transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            matchedRecord.get(FTU_REPORT_FIELDS.arrangementNumber),
          ),
        matchedRowNumber: matchedRecord.rowNumber,
        failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      },
      matchedRowNumber: matchedRecord.rowNumber,
    };
  }

  /** เปรียบเทียบ Field เมื่อพบ Record ที่จับคู่กัน */
  private compareMatchedRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: Exclude<FtuDirection, "NO_THB_LEG">,
  ): ResolvedCase {
    const failedHeaders = new Set<string>();
    const reviewHeaders = new Set<string>();
    const remarks: string[] = [];

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

    /** Matching ข้อ 3: Txn Date -> Data Set Date/วันที่ใน Arr Number (Review) */
    this.compareDate(testDataRecord, matchedRecord, addReview);

    /** Matching ข้อ 4: From/To Currency -> Leg Type 
     *  หากระบุ Direction ไม่ได้ ให้ Review เท่านั้น */
    this.compareLegType(testDataRecord, matchedRecord, direction, addFailure, addReview);

    /** Matching ข้อ 5: BOT Purpose -> Inflow/Outflow Purpose (Review) */
    this.comparePurpose(testDataRecord, matchedRecord, direction, addReview);

    /** Matching ข้อ 6: To Currency 2 ตัวแรก -> Country Id (Review) */
    this.compareCountry(testDataRecord, matchedRecord, addReview);

    /** Matching ข้อ 7: Settled Currency -> Currency Id */
    this.compareCurrency(testDataRecord, matchedRecord, addFailure);

    /** Matching ข้อ 8: Settled Amount -> Foreign Currency Amount */
    this.compareAmount(testDataRecord, matchedRecord, addReview);

    /** Business Rule: 0 < FTU Amount < 50,000 USD */
    this.evaluateAmountThreshold(matchedRecord, addFailure, addReview);

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
    reviewHeaders.size > 0
      ? FTU_REMARKS.pleaseReview
      : "";

    const finalRemark = [status === "PASS" ? successRemark : "", ...remarks, pleaseReviewRemark]
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
   * หากไม่ตรงกัน ให้ดึงวันที่ตำแหน่ง 7-12
   * ของ Arr Number ในรูปแบบ YYMMDD มาเทียบอีกครั้ง
   */
  private compareDate(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const expectedText = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.transactionDate,
    );

    const actualText = matchedRecord.get(FTU_REPORT_FIELDS.dataSetDate);

    const expectedDate = parseDate(expectedText);
    const actualDate = parseDate(actualText);

    if (!expectedDate) {
      addReview(
        FTU_REPORT_FIELDS.dataSetDate,
        FTU_TEST_DATA_FIELDS.transactionDate,
        expectedText,
        actualText,
      );
      return;
    }

    if (actualDate && isSameDate(expectedDate, actualDate)) {
      return;
    }

    const arrangementNumber = matchedRecord.get(
      FTU_REPORT_FIELDS.arrangementNumber,
    );

    const arrangementDate = extractDateFromArrangementNumber(arrangementNumber);

    if (arrangementDate && isSameDate(expectedDate, arrangementDate)) {
      addReview(
        FTU_REPORT_FIELDS.dataSetDate,
        FTU_TEST_DATA_FIELDS.transactionDate,
        `${expectedText} ` + `(ตรงกับวันที่ใน Arr Number)`,
        actualText,
      );
      return;
    }

    addReview(
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
    direction: Exclude<FtuDirection, "NO_THB_LEG">,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const actualLegType = matchedRecord.get(
    FTU_REPORT_FIELDS.legType,
    );
     /**
   * Return Case ที่ Currency ว่าง
   * ไม่สามารถระบุ Expected Leg Type ได้
   * จึงแสดงเป็น Review โดยไม่ตัดสิน FAIL
   */
  if (direction === "UNKNOWN_DIRECTION") {
    const fromCurrency = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.fromCurrency,
    );

    addReview(
      FTU_REPORT_FIELDS.legType,
      FTU_TEST_DATA_FIELDS.fromCurrency,
      fromCurrency,
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
    direction: Exclude<FtuDirection, "NO_THB_LEG">,
    addReview: AddComparisonRemark,
  ): void {
     if (direction === "UNKNOWN_DIRECTION") {
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
   * Amount ที่ไม่ตรงกันใช้เพื่อ Review เท่านั้น
   * และไม่มีการรวม Amount หลาย Record
   */
  private compareAmount(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const expectedText = testDataRecord.get(FTU_TEST_DATA_FIELDS.settledAmount);

    const actualText = matchedRecord.get(
      FTU_REPORT_FIELDS.foreignCurrencyAmount,
    );

    const expectedAmount = parseAmount(expectedText);
    const actualAmount = parseAmount(actualText);

    if (
      expectedAmount === null ||
      actualAmount === null ||
      expectedAmount !== actualAmount
    ) {
      addReview(
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

    const isWithinThreshold = amount > 0 && amount < FTU_USD_THRESHOLD;

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