/**
 * วิเคราะห์ Test Data ของ DS_FTU จาก Candidate ที่ FtuMatcher เลือกไว้
 *
 * 1. ตัดสิน Expected Presence/Absence
 * 2. ตรวจ Direction, Date, Purpose, Country, Currency และ Amount
 * 3. สร้างผล PASS, FAIL หรือ PASS พร้อม Review
 *
 * ไฟล์นี้ไม่ค้นหา Candidate และไม่อ่านหรือเขียน Excel Workbook
 */

import type { ReconcileRecord } from "../shared/record";
import { formatCompareRemark } from "../shared/remark";
import {
  extractDateFromArrangementNumber,
  formatDate,
  isSameDate,
  parseAmount,
  parseDate,
} from "../shared/reconcile-parse.util";
import type { ResultRow } from "../shared/result-writer";
import {
  FTU_AMOUNT_THRESHOLD_LABEL,
  FTU_AMOUNT_TOLERANCE,
  FTU_COUNTRY_CODE_LENGTH,
  FTU_DIRECTIONS,
  FTU_LEG_TYPES,
  FTU_REMARKS,
  FTU_REPORT_CODE,
  FTU_REPORT_FIELDS,
  FTU_TEST_DATA_FIELDS,
  FTU_USD_CURRENCY_CODE,
  FTU_USD_THRESHOLD,
  normalizeFtuText,
} from "./ftu-config";
import type { FtuDirection } from "./ftu-config";
import type { FtuMatchResult } from "./ftu-matcher";

type AddComparisonRemark = (
  reportField: string,
  testDataField: string,
  expected: string,
  actual: string,
) => void;

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

interface DetectionResult {
  detectedRecord?: ReconcileRecord;
  detectionRemark: string;
  fallbackVerificationUnavailable: boolean;
  candidateCount?: number;
}

export class FtuAnalyzer {
  /** สร้างผลลัพธ์ของ Test Data หนึ่งแถวจากผล Matching */
  analyze(
    testDataRecord: ReconcileRecord,
    matchResult: FtuMatchResult,
  ): ResultRow {
    const testCaseNo =
      normalizeFtuText(testDataRecord.get(FTU_TEST_DATA_FIELDS.testNo)) ||
      `TEST DATA ROW ${testDataRecord.rowNumber}`;
    const result = this.resolveRecord(
      testCaseNo,
      testDataRecord,
      matchResult,
    );

    const testNo = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.testNo),
    );

    if (testNo === "") {
      result.remark = this.appendRemark(
        result.remark,
        `Test Data ไม่มี ${FTU_TEST_DATA_FIELDS.testNo}`,
      );
    }

    if (matchResult.transactionId === "") {
      result.remark = this.appendRemark(
        result.remark,
        `Test Data ไม่มี ${FTU_TEST_DATA_FIELDS.transactionId}`,
      );
    }

    return result;
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

  /** แปลงผล Matcher เป็นข้อมูลตรวจ Presence ที่ Analyzer ใช้ร่วมกัน */
  private resolveDetection(matchResult: FtuMatchResult): DetectionResult {
    if (matchResult.exactMatchedRecord) {
      return {
        detectedRecord: matchResult.exactMatchedRecord,
        detectionRemark: "",
        fallbackVerificationUnavailable: false,
      };
    }

    if (matchResult.fallbackUnavailableRemark) {
      return {
        detectionRemark: matchResult.fallbackUnavailableRemark,
        fallbackVerificationUnavailable: true,
      };
    }

    const fallbackResult = matchResult.fallbackResult;

    return {
      detectedRecord: fallbackResult?.matchedRecord,
      detectionRemark: fallbackResult?.remark ?? "",
      fallbackVerificationUnavailable:
        fallbackResult?.candidateCount === undefined,
      candidateCount: fallbackResult?.candidateCount,
    };
  }

  /** ตัดสิน Presence และตรวจ Record ที่ Matcher หาให้ */
  private resolveRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchResult: FtuMatchResult,
  ): ResultRow {
    const detection = this.resolveDetection(matchResult);

    if (matchResult.direction === FTU_DIRECTIONS.noThbLeg) {
      return this.resolveNoThbLeg(
        testCaseNo,
        testDataRecord,
        matchResult,
        detection,
      );
    }

    const thresholdDecision =
      this.resolveThresholdPresenceDecision(testDataRecord);

    if (thresholdDecision?.mustNotExist) {
      return this.resolveThresholdExpectedAbsence(
        testCaseNo,
        testDataRecord,
        matchResult,
        detection,
        thresholdDecision.amount,
      );
    }

    const matchedRecord = detection.detectedRecord;
    let fallbackRemark = matchResult.failureRemark ?? "";

    if (!matchResult.exactMatchedRecord) {
      const detectionRemark =
        matchResult.transactionId === ""
          ? detection.detectionRemark
          : matchedRecord
            ? this.appendRemark(
                `ควรพบ ${FTU_REPORT_FIELDS.arrangementNumber} = ` +
                  `"${matchResult.transactionId}" แต่ Matching Key หลักไม่พบ`,
                detection.detectionRemark,
              )
            : detection.detectionRemark;
      fallbackRemark = this.appendRemark(fallbackRemark, detectionRemark);
    }

    if (!matchedRecord) {
      return {
        testCaseNo,
        status: "FAIL",
        remark: this.appendRemark(
          formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_DATA_FIELDS.transactionId,
            matchResult.transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            "ไม่พบข้อมูล",
          ),
          fallbackRemark,
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      matchedRecord,
      matchResult.direction,
      { fallbackRemark },
    );
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

  /** สูตรกลางสำหรับตรวจช่วง Amount ของ DS_FTU */
  private isWithinUsdThreshold(amount: number): boolean {
    return amount > 0 && amount < FTU_USD_THRESHOLD;
  }

  /** ใช้ Test Data ตัดสินเกณฑ์ Presence สำหรับ Amount สกุล USD */
  private resolveThresholdPresenceDecision(
    testDataRecord: ReconcileRecord,
  ): ThresholdPresenceDecision | undefined {
    const currency = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledCurrency),
    );
    const amount = parseAmount(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledAmount),
    );

    if (currency !== FTU_USD_CURRENCY_CODE || amount === null) {
      return undefined;
    }

    return {
      amount,
      mustNotExist: amount >= FTU_USD_THRESHOLD,
    };
  }

  /** USD Amount นอกเกณฑ์ต้องไม่พบ Record ใน DS_FTU */
  private resolveThresholdExpectedAbsence(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchResult: FtuMatchResult,
    detection: DetectionResult,
    amount: number,
  ): ResultRow {
    const expectationRemark =
      `Test Data Amount = ${amount} USD ไม่เข้าเกณฑ์ ` +
      `0 < Amount < ${FTU_USD_THRESHOLD} USD ` +
      `จึงไม่ควรพบรายการใน ${FTU_REPORT_CODE}`;

    if (!detection.detectedRecord && (detection.candidateCount ?? 0) > 1) {
      return {
        testCaseNo,
        status: "FAIL",
        remark: this.appendRemark(
          `${expectationRemark}\nแต่ Fallback พบรายการใน Report`,
          detection.detectionRemark,
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    if (!detection.detectedRecord) {
      const verificationRemark = detection.fallbackVerificationUnavailable
        ? FTU_REMARKS.expectedAbsenceVerificationUnavailable
        : FTU_REMARKS.expectedAbsenceReportNotFound;

      return {
        testCaseNo,
        status: "PASS",
        remark: this.appendRemark(
          `${expectationRemark}\n${verificationRemark}`,
          detection.detectionRemark,
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: detection.fallbackVerificationUnavailable
          ? [FTU_REPORT_FIELDS.arrangementNumber]
          : [],
        isExpectedAbsence: true,
      };
    }

    const arrangementNumber = detection.detectedRecord.get(
      FTU_REPORT_FIELDS.arrangementNumber,
    );

    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      detection.detectedRecord,
      matchResult.direction,
      {
        initialFailureRemark: this.appendRemark(
          `${expectationRemark}\nแต่พบรายการใน Report: ` +
            `${FTU_REPORT_FIELDS.arrangementNumber} = ` +
            `"${arrangementNumber}"`,
          detection.detectionRemark,
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
    matchResult: FtuMatchResult,
    detection: DetectionResult,
  ): ResultRow {
    const testDataCurrencyPairRemark =
      this.formatTestDataCurrencyPairRemark(testDataRecord);

    if (!detection.detectedRecord && (detection.candidateCount ?? 0) > 1) {
      return {
        testCaseNo,
        status: "FAIL",
        remark: this.appendRemark(
          `${FTU_REMARKS.noThbLegUnexpectedPresence}\n` +
            testDataCurrencyPairRemark,
          detection.detectionRemark,
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [FTU_REPORT_FIELDS.arrangementNumber],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    if (!detection.detectedRecord) {
      const verificationRemark = detection.fallbackVerificationUnavailable
        ? FTU_REMARKS.expectedAbsenceVerificationUnavailable
        : FTU_REMARKS.noThbLegReportNotFound;
      const fallbackReviewRemark = detection.fallbackVerificationUnavailable
        ? detection.detectionRemark
        : "";

      return {
        testCaseNo,
        status: "PASS",
        remark: this.appendRemark(
          `${FTU_REMARKS.noThbLegExpectedAbsence}\n` +
            `${verificationRemark}\n${testDataCurrencyPairRemark}`,
          fallbackReviewRemark,
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: detection.fallbackVerificationUnavailable
          ? [FTU_REPORT_FIELDS.arrangementNumber]
          : [],
        isExpectedAbsence: true,
      };
    }

    const keyRemark =
      matchResult.exactMatchedRecord && matchResult.transactionId !== ""
        ? formatCompareRemark(
            FTU_REPORT_CODE,
            FTU_TEST_DATA_FIELDS.transactionId,
            matchResult.transactionId,
            FTU_REPORT_FIELDS.arrangementNumber,
            detection.detectedRecord.get(FTU_REPORT_FIELDS.arrangementNumber),
          )
        : detection.detectionRemark;

    return this.compareMatchedRecord(
      testCaseNo,
      testDataRecord,
      detection.detectedRecord,
      FTU_DIRECTIONS.noThbLeg,
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

  /** สร้าง Remark เปรียบเทียบรูปแบบกลางของ FTU */
  private buildCompareRemark(
    reportField: string,
    testDataField: string,
    expected: string,
    actual: string,
  ): string {
    return formatCompareRemark(
      FTU_REPORT_CODE,
      testDataField,
      expected,
      reportField,
      actual,
    );
  }

  /** เปรียบเทียบทุก Field เมื่อพบ Record ที่จับคู่กัน */
  private compareMatchedRecord(
    testCaseNo: string,
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: FtuDirection,
    options: MatchedComparisonOptions = {},
  ): ResultRow {
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
        this.buildCompareRemark(
          reportField,
          testDataField,
          expected,
          actual,
        ),
      );
    };
    const addReview: AddComparisonRemark = (
      reportField,
      testDataField,
      expected,
      actual,
    ): void => {
      reviewHeaders.add(reportField);
      remarks.push(
        this.buildCompareRemark(
          reportField,
          testDataField,
          expected,
          actual,
        ),
      );
    };

    if (options.fallbackRemark) {
      reviewHeaders.add(FTU_REPORT_FIELDS.arrangementNumber);
      remarks.push(options.fallbackRemark);
    }

    this.compareDate(testDataRecord, matchedRecord, addFailure, addReview);
    this.compareLegType(
      testDataRecord,
      matchedRecord,
      direction,
      addFailure,
      addReview,
    );
    this.comparePurpose(testDataRecord, matchedRecord, direction, addReview);
    this.compareCountry(testDataRecord, matchedRecord, addReview);
    this.compareCurrency(testDataRecord, matchedRecord, addFailure);
    this.compareAmount(testDataRecord, matchedRecord, addFailure);

    if (!options.skipReportThresholdValidation) {
      this.evaluateAmountThreshold(matchedRecord, addFailure, addReview);
    }

    const status = failedHeaders.size === 0 ? "PASS" : "FAIL";
    const successRemark =
      direction === FTU_DIRECTIONS.buyForeignCurrency
        ? FTU_REMARKS.buyForeignCurrency
        : direction === FTU_DIRECTIONS.sellForeignCurrency
          ? FTU_REMARKS.sellForeignCurrency
          : "";
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
      testCaseNo,
      status,
      remark: finalRemark,
      matchedRowNumber: matchedRecord.rowNumber,
      failedKeyFieldHeaders: [...failedHeaders],
      reviewFieldHeaders: [...reviewHeaders],
      isExpectedAbsence: false,
    };
  }

  /** Date ว่างให้ Review; ไม่ตรงทั้ง Report Date/Arr Date ให้ FAIL */
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

    const actualDate = parseDate(actualText);

    if (actualDate && isSameDate(expectedDate, actualDate)) {
      return;
    }

    const arrangementDate = extractDateFromArrangementNumber(
      matchedRecord.get(FTU_REPORT_FIELDS.arrangementNumber),
    );

    if (arrangementDate && isSameDate(expectedDate, arrangementDate)) {
      addReview(
        FTU_REPORT_FIELDS.dataSetDate,
        FTU_TEST_DATA_FIELDS.transactionDate,
        expectedText,
        `${actualText} | Arr Number Date: ${formatDate(arrangementDate)}`,
      );
      return;
    }

    addFailure(
      FTU_REPORT_FIELDS.dataSetDate,
      FTU_TEST_DATA_FIELDS.transactionDate,
      expectedText,
      `${actualText} | Arr Number Date: ${formatDate(arrangementDate)}`,
    );
  }

  /** BUY/SELL ต้องตรงกับ Leg Type; Direction ไม่ชัดให้ Review */
  private compareLegType(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: FtuDirection,
    addFailure: AddComparisonRemark,
    addReview: AddComparisonRemark,
  ): void {
    const actualLegType = matchedRecord.get(FTU_REPORT_FIELDS.legType);

    if (
      direction === FTU_DIRECTIONS.unknown ||
      direction === FTU_DIRECTIONS.noThbLeg
    ) {
      const fromCurrency = testDataRecord.get(
        FTU_TEST_DATA_FIELDS.fromCurrency,
      );
      const toCurrency = testDataRecord.get(FTU_TEST_DATA_FIELDS.toCurrency);

      addReview(
        FTU_REPORT_FIELDS.legType,
        `${FTU_TEST_DATA_FIELDS.fromCurrency}/` +
          FTU_TEST_DATA_FIELDS.toCurrency,
        `${fromCurrency}/${toCurrency}`,
        actualLegType,
      );
      return;
    }

    const expectedLegType =
      direction === FTU_DIRECTIONS.buyForeignCurrency
        ? FTU_LEG_TYPES.buyForeignCurrency
        : FTU_LEG_TYPES.sellForeignCurrency;

    if (normalizeFtuText(actualLegType) !== normalizeFtuText(expectedLegType)) {
      addFailure(
        FTU_REPORT_FIELDS.legType,
        `${FTU_TEST_DATA_FIELDS.fromCurrency}/` +
          FTU_TEST_DATA_FIELDS.toCurrency,
        expectedLegType,
        actualLegType,
      );
    }
  }

  /** Purpose ไม่ตรงให้ Review โดยไม่ตัดสิน FAIL */
  private comparePurpose(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    direction: FtuDirection,
    addReview: AddComparisonRemark,
  ): void {
    if (
      direction === FTU_DIRECTIONS.unknown ||
      direction === FTU_DIRECTIONS.noThbLeg
    ) {
      return;
    }

    const expectedPurpose = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.purposeCode,
    );
    const reportPurposeField =
      direction === FTU_DIRECTIONS.buyForeignCurrency
        ? FTU_REPORT_FIELDS.inflowPurpose
        : FTU_REPORT_FIELDS.outflowPurpose;
    const actualPurpose = matchedRecord.get(reportPurposeField);

    if (
      normalizeFtuText(expectedPurpose) !== normalizeFtuText(actualPurpose)
    ) {
      addReview(
        reportPurposeField,
        FTU_TEST_DATA_FIELDS.purposeCode,
        expectedPurpose,
        actualPurpose,
      );
    }
  }

  /** To Currency 2 ตัวแรกต้องตรงกับ Country Id */
  private compareCountry(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addReview: AddComparisonRemark,
  ): void {
    const toCurrency = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.toCurrency),
    );
    const expectedCountry = toCurrency.slice(0, FTU_COUNTRY_CODE_LENGTH);
    const actualCountry = matchedRecord.get(
      FTU_REPORT_FIELDS.beneficiaryCountry,
    );

    if (
      expectedCountry === "" ||
      normalizeFtuText(actualCountry) !== expectedCountry
    ) {
      addReview(
        FTU_REPORT_FIELDS.beneficiaryCountry,
        FTU_TEST_DATA_FIELDS.toCurrency,
        expectedCountry,
        actualCountry,
      );
    }
  }

  /** Settled Currency ต้องตรงกับ Currency Id */
  private compareCurrency(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
  ): void {
    const expectedCurrency = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.settledCurrency,
    );
    const actualCurrency = matchedRecord.get(FTU_REPORT_FIELDS.currencyId);

    if (
      normalizeFtuText(expectedCurrency) !== normalizeFtuText(actualCurrency)
    ) {
      addFailure(
        FTU_REPORT_FIELDS.currencyId,
        FTU_TEST_DATA_FIELDS.settledCurrency,
        expectedCurrency,
        actualCurrency,
      );
    }
  }

  /** Currency ตรงกันแล้วจึงเทียบ Settled/Foreign Currency Amount */
  private compareAmount(
    testDataRecord: ReconcileRecord,
    matchedRecord: ReconcileRecord,
    addFailure: AddComparisonRemark,
  ): void {
    const expectedCurrency = normalizeFtuText(
      testDataRecord.get(FTU_TEST_DATA_FIELDS.settledCurrency),
    );
    const actualCurrency = normalizeFtuText(
      matchedRecord.get(FTU_REPORT_FIELDS.currencyId),
    );

    if (
      expectedCurrency === "" ||
      actualCurrency === "" ||
      expectedCurrency !== actualCurrency
    ) {
      return;
    }

    const expectedText = testDataRecord.get(
      FTU_TEST_DATA_FIELDS.settledAmount,
    );
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

  /** ตรวจ 0 < FTU Amount < 50,000 USD */
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
    const normalizedCurrency = normalizeFtuText(currencyId);

    if (amount === null) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_AMOUNT_THRESHOLD_LABEL,
        `0 < Amount < ${FTU_USD_THRESHOLD} USD`,
        amountText,
      );
      return;
    }

    if (normalizedCurrency !== FTU_USD_CURRENCY_CODE) {
      addReview(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_AMOUNT_THRESHOLD_LABEL,
        `0 < USD Equivalent < ${FTU_USD_THRESHOLD}`,
        `${amount} ${currencyId} (รอข้อมูลแปลงเป็น USD)`,
      );
      return;
    }

    if (!this.isWithinUsdThreshold(amount)) {
      addFailure(
        FTU_REPORT_FIELDS.foreignCurrencyAmount,
        FTU_AMOUNT_THRESHOLD_LABEL,
        `0 < Amount < ${FTU_USD_THRESHOLD} USD`,
        `${amount} USD`,
      );
    }
  }
}