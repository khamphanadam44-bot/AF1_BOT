/**
 * ltx-reconcile
 *
 * โครงสร้าง Class ที่เกี่ยวข้อง:
 *   getReconcileConfig          -> อ่านและตรวจ Config ของ LTX
 *   ReconcileWorkbookPreparer   -> Copy Report + สร้าง Sheet ผลลัพธ์
 *   ReconcileExcelReader        -> อ่าน Excel (Report/Test Data) เป็น ReconcileRecord[]
 *   IExpectedCaseBuilder        -> Test Data -> Expected Case (strategy ต่อ report)
 *   ReconcileMatcher            -> จับคู่ Test Case กับ Report row (Reference Number ตรง ๆ)
 *   FieldRuleValidatorSet       -> เทียบ field อื่นที่ไม่ใช่ 3 หัวข้อหลัก (ผล = ไฮไลท์เหลืองเท่านั้น)
 *   ReconcileResultSheetWriter  -> เขียนผลลง Sheet (copy ทุกแถวจาก AF1 Report + แปะ annotation)
 *
 *  Test_Result Pass/Fail ตัดสินจาก "3 หัวข้อ" เท่านั้น:
 *   1. Reference Transaction Number (Report) vs Transaction ID/ Reconcile ID (Test Data)
 *   2. FI Arrangement Number (Report) vs From Account (A/C Client/Sender) (Test Data)
 *   3. Transaction Amount (Report) vs SUM ของ Fee Amount ทุกรายการ (แถว FE) หรือ
 *      From Transfer Amount/Debit Amount (แถว DR)
 * ครบทั้ง 3 ข้อ = Pass, ข้อใดข้อหนึ่งไม่ตรง = Fail (checkKeyConditions ด้านล่าง)
 *
 * ส่วน field อื่นทั้งหมด (Payment Method, Currency, Beneficiary Name, ฯลฯ) ยังคงเทียบตาม
 * Requirement เดิมผ่าน FieldRuleValidatorSet เหมือนเดิมทุกอย่าง แต่ผลลัพธ์ตอนนี้ไม่ว่าจะ
 * เป็น FAIL หรือ REVIEW (ตามตรรกะเดิมของ FieldComparer) จะถูกปฏิบัติเหมือนกันหมด คือ
 * "ไฮไลท์เหลือง + ใส่ข้อความใน Remark" เท่านั้น ไม่กระทบ Pass/Fail ของแถวอีกต่อไป
 *
 *  Fallback Matching ของ LTX:
 * - ใช้ Exact Reference Matching เป็นทางหลักเหมือนเดิม
 * - ถ้า Test Data ไม่มี Transaction ID หรือค้น Reference ไม่พบ ให้ค้นคู่สำรองจาก
 *   FI Arrangement Number/Account + Currency + DR/FE suffix + Amount
 * - DR Amount เทียบ From Transfer Amount หรือ Debit Amount
 * - FE Amount เทียบ SUM Fee Amount ที่ ExpectedCaseBuilder คำนวณไว้
 * - ต้องพบ Report row เพียง 1 แถวเท่านั้นจึงถือว่า Mapping เจอ
 * - แม้ Fallback จะ Mapping เจอ แต่ Transaction ID ที่ว่าง/ไม่ตรงยังคงทำให้ Key เป็น FAIL
 *   และระบบจะตรวจ Account, Amount และ Field อื่นต่อ พร้อมบันทึก Remark
 * ------------------------------------------------------------------
 */
import { getReconcileConfig, ReconcileReportConfig } from "./ltx-config";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { ReconcileMatcher } from "../shared/record-matcher";
import { AmountComparator } from "./amount-compare";
import { FieldRuleValidatorSet } from "./field-validator";
import {
  ReconcileResultSheetWriter,
  ResultRow,
  RowStatus,
} from "../shared/result-writer";
import { ExpectedCase, IExpectedCaseBuilder } from "./expected-case";
import { LtxExpectedCaseBuilder } from "./expected-case-builder";
import { ReconcileRecord } from "../shared/record";
import {
  getPresenceRuleFields,
  ReportPresenceRuleEvaluator,
} from "../shared/presence-rule";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import {
  getMappingHeaderRowNumber,
  getMappingMatchingKeyHeaders,
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import type { ReportCode } from "../../../config/report-config";

const TEST_NO_HEADER = "Test No.";

/**
 * เลือก ExpectedCaseBuilder ที่ถูกต้องให้แต่ละ report — จุดเดียวที่ต้องแก้เมื่อเพิ่ม report ใหม่
 */
const EXPECTED_CASE_BUILDERS: Partial<
  Record<ReportCode, IExpectedCaseBuilder>
> = {
  DS_LTX: new LtxExpectedCaseBuilder(),
};

export class ReconcileService {
  constructor(
    private readonly workbookPreparer: ReconcileWorkbookPreparer = new ReconcileWorkbookPreparer(),
    private readonly excelReader: ReconcileExcelReader = new ReconcileExcelReader(),
    private readonly matcher: ReconcileMatcher = new ReconcileMatcher(),
    private readonly fieldValidatorSet: FieldRuleValidatorSet = new FieldRuleValidatorSet(),
    private readonly sheetWriter: ReconcileResultSheetWriter = new ReconcileResultSheetWriter(),
    private readonly amountComparator: AmountComparator = new AmountComparator(),
    private readonly presenceRuleEvaluator: ReportPresenceRuleEvaluator = new ReportPresenceRuleEvaluator(),
  ) {}

  private validateMappingConfiguration(
    reportCode: string,
    config: ReconcileReportConfig,
  ): {
    mappingHeaderRowNumber: number;
    matchingKeyHeader: string;
  } {
    const mappingReportName = requireMappingReportName(reportCode);
    const mappingHeaderRowNumber = getMappingHeaderRowNumber(mappingReportName);
    const [matchingKeyHeader] = getMappingMatchingKeyHeaders(mappingReportName);

    if (!matchingKeyHeader) {
      throw new Error(
        `[${reportCode}] Mapping Config ไม่มี Matching Key Header.`,
      );
    }

    if (mappingHeaderRowNumber !== config.headerRowNumber) {
      throw new Error(
        `[${reportCode}] Header Row Config mismatch.\n` +
          `mapping-config.ts = ${mappingHeaderRowNumber}\n` +
          `ltx-config.ts = ${config.headerRowNumber}`,
      );
    }

    if (
      canonicalHeader(matchingKeyHeader) !==
      canonicalHeader(config.referenceNumberReportField)
    ) {
      throw new Error(
        `[${reportCode}] Matching Key Config mismatch.\n` +
          `mapping-config.ts = "${matchingKeyHeader}"\n` +
          `ltx-config.ts = "${config.referenceNumberReportField}"`,
      );
    }

    const configuredHeaders = new Set(
      getUniqueMappingHeaders(mappingReportName).map((header) =>
        canonicalHeader(header),
      ),
    );

    const unknownRuleHeaders = config.fieldRules
      .map((rule) => rule.reportField)
      .filter(
        (reportField) => !configuredHeaders.has(canonicalHeader(reportField)),
      );

    if (unknownRuleHeaders.length > 0) {
      throw new Error(
        `[${reportCode}] Reconcile Config อ้างถึง Report Header ` +
          "ที่ไม่มีใน Mapping Config: " +
          [...new Set(unknownRuleHeaders)].join(", "),
      );
    }

    return {
      mappingHeaderRowNumber,
      matchingKeyHeader,
    };
  }

  private getExpectedCaseBuilder(reportCode: string): IExpectedCaseBuilder {
    const builder = EXPECTED_CASE_BUILDERS[reportCode as ReportCode];
    if (!builder) {
      throw new Error(
        `ยังไม่มี ExpectedCaseBuilder สำหรับ report "${reportCode}" (ดู EXPECTED_CASE_BUILDERS)`,
      );
    }
    return builder;
  }

  private formatRuleRemark(remark: string, reportCode: string): string {
    return remark.replaceAll("{REPORT_CODE}", reportCode);
  }

  private validatePresenceRuleHeaders(
    reportCode: string,
    actualHeaders: string[],
    config: ReconcileReportConfig,
  ): void {
    const requiredRuleFields = getPresenceRuleFields(
      config.reportPresenceRules,
    );

    if (requiredRuleFields.length === 0) {
      return;
    }

    const normalizedActualHeaders = new Set(
      actualHeaders
        .filter((header) => header.trim() !== "")
        .map((header) => canonicalHeader(header)),
    );

    const missingHeaders = requiredRuleFields.filter(
      (requiredHeader) =>
        !normalizedActualHeaders.has(canonicalHeader(requiredHeader)),
    );

    if (missingHeaders.length === 0) {
      return;
    }

    throw new Error(
      `[${reportCode}] Test Data ไม่มี Header ` +
        "ที่จำเป็นสำหรับ Reconcile Presence Rule: " +
        missingHeaders.join(", "),
    );
  }

  /**
   * Normalize ค่าที่ใช้ Fallback Matching
   * - ตัดช่องว่างหน้าและหลัง
   * - ไม่สนตัวพิมพ์เล็ก/ใหญ่
   */
  private normalizeFallbackValue(value: string): string {
    return value.trim().toUpperCase();
  }

  /**
   * ตรวจว่าค่าจาก Test Data และ AF1 Report ตรงกันหรือไม่
   * ค่าว่างจะไม่ถือว่าตรงกัน เพื่อป้องกันการ Mapping จากค่าว่างสองฝั่ง
   */
  private isSameFallbackValue(
    expectedValue: string,
    actualValue: string,
  ): boolean {
    const normalizedExpected = this.normalizeFallbackValue(expectedValue);
    const normalizedActual = this.normalizeFallbackValue(actualValue);

    return (
      normalizedExpected !== "" &&
      normalizedActual !== "" &&
      normalizedExpected === normalizedActual
    );
  }

  /**
   * ตรวจว่า Report Row เป็นขา DR หรือ FE
   * ที่กำลังทำ Fallback Matching หรือไม่
   *
   * ตัวอย่าง:
   * - suffix = "-DR" ใช้ Reference ที่ลงท้ายด้วย -DR
   * - suffix = "-FE" ใช้ Reference ที่ลงท้ายด้วย -FE
   */
  private isSameReferenceSlot(
    record: ReconcileRecord,
    config: ReconcileReportConfig,
    suffix: string,
  ): boolean {
    const actualReference = this.normalizeFallbackValue(
      record.get(config.referenceNumberReportField),
    );

    const normalizedSuffix = this.normalizeFallbackValue(suffix);

    return (
      actualReference !== "" &&
      normalizedSuffix !== "" &&
      actualReference.endsWith(normalizedSuffix)
    );
  }

  /**
   * ตรวจ Amount ตาม Logic เดิมของ LTX
   * - FE: Report Transaction Amount ต้องเท่ากับ SUM Fee Amount
   * - DR: Report Transaction Amount ต้องเท่ากับ From Transfer Amount
   *       หรือ Debit Amount ซึ่งเป็นค่า Fallback
   */
  private isFallbackAmountMatch(
    record: ReconcileRecord,
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): boolean {
    const actualAmount = record.get(config.transactionAmountReportField);
    const isFeSlot =
      this.normalizeFallbackValue(suffix) ===
      this.normalizeFallbackValue(config.feSuffixLabel);

    if (isFeSlot) {
      return this.amountComparator.compare(
        expectedCase.expectedFeAmount,
        actualAmount,
      ).isMatch;
    }

    const matchesPrimary = this.amountComparator.compare(
      expectedCase.primaryRecord.get(config.drAmountTestDataField),
      actualAmount,
    ).isMatch;
    const matchesFallback = this.amountComparator.compare(
      expectedCase.primaryRecord.get(config.drAmountFallbackTestDataField),
      actualAmount,
    ).isMatch;

    return matchesPrimary || matchesFallback;
  }

  /**
   * Fallback Matching ของ LTX เมื่อ Exact Reference Matching ใช้งานไม่ได้
   *
   * ลำดับการกรอง:
   * 1. FI Arrangement Number ต้องตรงกับ Account ใน Test Data
   * 2. Currency Id ต้องตรงกับ Currency ใน Test Data
   * 3. Reference ต้องเป็นขา DR/FE ที่กำลังตรวจ
   * 4. Amount ต้องตรงตาม Logic ของ DR หรือ FE
   *
   * ไม่เลือก .first() เพราะถ้าพบหลายแถวจะเป็น Mapping ที่กำกวม
   */
  private findLtxFallbackCandidates(
    reportRecords: ReconcileRecord[],
    expectedCase: ExpectedCase,
    config: ReconcileReportConfig,
    suffix: string,
  ): ReconcileRecord[] {
    const expectedAccount = expectedCase.primaryRecord
      .get(config.groupKeyFields.testDataAccountField)
      .trim();
    const expectedCurrency = expectedCase.primaryRecord
      .get(config.groupKeyFields.testDataCurrencyField)
      .trim();

    // Account และ Currency เป็นข้อมูลขั้นต่ำของ Fallback Mapping
    // ถ้าขาดอย่างใดอย่างหนึ่งจะไม่เดา Report row จาก Amount เพียงอย่างเดียว
    if (expectedAccount === "" || expectedCurrency === "") {
      return [];
    }

    return reportRecords.filter((record) => {
      const actualAccount = record
        .get(config.groupKeyFields.reportAccountField)
        .trim();
      const actualCurrency = record
        .get(config.groupKeyFields.reportCurrencyField)
        .trim();

      return (
        this.isSameFallbackValue(expectedAccount, actualAccount) &&
        this.isSameFallbackValue(expectedCurrency, actualCurrency) &&
        this.isSameReferenceSlot(record, config, suffix) &&
        this.isFallbackAmountMatch(record, expectedCase, config, suffix)
      );
    });
  }

  private checkKeyConditions(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    suffix: string,
    matchedRecord: ReconcileRecord,
  ): {
    status: RowStatus;
    failedKeyFieldHeaders: string[];
    failRemarks: string[];
  } {
    const failedKeyFieldHeaders: string[] = [];
    const failRemarks: string[] = [];

    const expectedId = expectedCase.primaryRecord
      .get(config.testDataIdField)
      .trim()
      .toUpperCase();
    const actualReference = matchedRecord
      .get(config.referenceNumberReportField)
      .trim()
      .toUpperCase();
    const actualId = actualReference.endsWith(suffix.toUpperCase())
      ? actualReference.slice(0, actualReference.length - suffix.length)
      : actualReference;

    if (expectedId === "" || actualId !== expectedId) {
      failedKeyFieldHeaders.push(config.referenceNumberReportField);
      failRemarks.push(
        `[TS] : ${config.testDataIdField} = "${expectedId}" | ` +
          `[${reportCode}] : ${config.referenceNumberReportField} = "${actualReference}"`,
      );
    }

    const expectedAccount = expectedCase.primaryRecord
      .get(config.groupKeyFields.testDataAccountField)
      .trim();
    const actualAccount = matchedRecord
      .get(config.groupKeyFields.reportAccountField)
      .trim();

    if (
      expectedAccount === "" ||
      actualAccount === "" ||
      expectedAccount.toLowerCase() !== actualAccount.toLowerCase()
    ) {
      failedKeyFieldHeaders.push(config.groupKeyFields.reportAccountField);
      failRemarks.push(
        `[TS] : ${config.groupKeyFields.testDataAccountField} = "${expectedAccount}" | ` +
          `[${reportCode}] : ${config.groupKeyFields.reportAccountField} = "${actualAccount}"`,
      );
    }

    const actualAmount = matchedRecord.get(config.transactionAmountReportField);
    let isAmountOk: boolean;

    if (suffix === config.feSuffixLabel) {
      isAmountOk = this.amountComparator.compare(
        expectedCase.expectedFeAmount,
        actualAmount,
      ).isMatch;
    } else {
      const matchesPrimary = this.amountComparator.compare(
        expectedCase.primaryRecord.get(config.drAmountTestDataField),
        actualAmount,
      ).isMatch;
      const matchesFallback = this.amountComparator.compare(
        expectedCase.primaryRecord.get(config.drAmountFallbackTestDataField),
        actualAmount,
      ).isMatch;
      isAmountOk = matchesPrimary || matchesFallback;
    }

    if (!isAmountOk) {
      failedKeyFieldHeaders.push(config.transactionAmountReportField);
      if (suffix === config.feSuffixLabel) {
        failRemarks.push(
          `[TS] : SUM(Fee Amount) = "${expectedCase.expectedFeAmount}" | ` +
            `[${reportCode}] : ${config.transactionAmountReportField} = "${actualAmount}"`,
        );
      } else {
        const expectedTransferAmount = expectedCase.primaryRecord
          .get(config.drAmountTestDataField)
          .trim();
        const expectedDebitAmount = expectedCase.primaryRecord
          .get(config.drAmountFallbackTestDataField)
          .trim();
        failRemarks.push(
          `[TS] : ${config.drAmountTestDataField} = "${expectedTransferAmount}" ` +
            `/ ${config.drAmountFallbackTestDataField} = "${expectedDebitAmount}" | ` +
            `[${reportCode}] : ${config.transactionAmountReportField} = "${actualAmount}"`,
        );
      }
    }

    return {
      status: failedKeyFieldHeaders.length === 0 ? "PASS" : "FAIL",
      failedKeyFieldHeaders,
      failRemarks,
    };
  }

  private buildReviewSection(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    matchedRecord: ReconcileRecord,
    suffix: string,
  ): { reviewFieldHeaders: string[]; remark: string } {
    const fieldResults = this.fieldValidatorSet.validateAll(
      reportCode,
      config.fieldRules,
      expectedCase.primaryRecord,
      matchedRecord,
      suffix,
    );

    const mismatches = fieldResults.filter(
      (result) => result.status !== "PASS",
    );

    if (mismatches.length === 0) {
      return { reviewFieldHeaders: [], remark: "" };
    }

    const remark = `${mismatches.map((result) => result.remark).join("\n")}\nPlease review`;

    return {
      reviewFieldHeaders: mismatches.map((result) => result.fieldHeader),
      remark,
    };
  }

  private resolveSlot(
    reportCode: string,
    config: ReconcileReportConfig,
    expectedCase: ExpectedCase,
    expectedReference: string,
    suffix: string,
    reportRecords: ReconcileRecord[],
  ): ResultRow {
    const presenceDecision = this.presenceRuleEvaluator.evaluate(
      expectedCase.primaryRecord,
      config.reportPresenceRules,
    );

    const expectedId = expectedCase.primaryRecord
      .get(config.testDataIdField)
      .trim();
    const normalizedExpectedReference = expectedReference.trim();

    const testNo = expectedCase.primaryRecord.get(TEST_NO_HEADER).trim();

    const missingTestNoRemark =
      testNo === "" ? `Test Data ไม่มี ${TEST_NO_HEADER}` : "";

    /**
     * ใส่ Comment กรณี Test No. ว่างให้ Result ทุกเส้นทาง
     * โดยไม่กระทบผล Matching และ PASS/FAIL
     */
    const withMissingTestNoRemark = (...remarks: string[]): string =>
      [missingTestNoRemark, ...remarks]
        .filter((remark) => remark.trim() !== "")
        .join("\n");

    /**
     * Primary Matching:
     * ถ้ามี Expected Reference ให้ค้นแบบ Exact Reference ก่อนเสมอ
     */
    let matchedRecord =
      normalizedExpectedReference !== ""
        ? this.matcher.findByExactReference(
            reportRecords,
            config.referenceNumberReportField,
            normalizedExpectedReference,
          )
        : undefined;

    let fallbackRemark = "";

    /**
     * Fallback Matching:
     * ทำเมื่อ Transaction ID ว่าง หรือ Exact Reference Matching ไม่พบคู่
     */
    if (!matchedRecord) {
      const fallbackCandidates = this.findLtxFallbackCandidates(
        reportRecords,
        expectedCase,
        config,
        suffix,
      );

      if (fallbackCandidates.length > 1) {
        const candidateRows = fallbackCandidates
          .map((record) => record.rowNumber)
          .join(", ");
        const fallbackReason =
          expectedId === ""
            ? `Test Data ไม่มี ${config.testDataIdField}`
            : `ไม่พบ Exact Reference = "${normalizedExpectedReference}"`;

        return {
          testCaseNo: expectedCase.displayTestCaseNo,
          status: "FAIL",
          remark: withMissingTestNoRemark(
            fallbackReason,
            `LTX Fallback Matching พบ ${fallbackCandidates.length} แถว ` +
              `(Report row: ${candidateRows}) จึงไม่สามารถระบุคู่ที่แน่นอนได้`,
          ),
          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: [],
          isExpectedAbsence: false,
        };
      }

      if (fallbackCandidates.length === 1) {
        matchedRecord = fallbackCandidates[0];

        const matchedReference = matchedRecord
          .get(config.referenceNumberReportField)
          .trim();
        const fallbackReason =
          expectedId === ""
            ? `Test Data ไม่มี ${config.testDataIdField}`
            : `ไม่พบ Exact Reference = "${normalizedExpectedReference}"`;

        fallbackRemark =
          `${fallbackReason}\n` +
          "Mapping ด้วย LTX Fallback: Account + Currency + DR/FE + Amount " +
          `พบคู่กับ ${config.referenceNumberReportField} = "${matchedReference}"`;
      }
    }

    /**
     * Presence Rule กำหนดว่า Test Data รายการนี้
     * ต้องไม่มีอยู่ใน AF1 Report
     */
    if (presenceDecision.expectation === "MUST_NOT_EXIST") {
      /**
       * ไม่พบ Report Row ตามที่ Business Rule คาดไว้
       * จึงถือว่า PASS แบบ Expected Absence
       *
       * ถ้า Transaction ID ว่าง
       * ให้บันทึก Comment เพิ่ม แต่ไม่เปลี่ยนผล
       * ของ Expected Absence
       */
      if (!matchedRecord) {
        return {
          testCaseNo: expectedCase.displayTestCaseNo,

          status: "PASS",

          remark: withMissingTestNoRemark(
            this.formatRuleRemark(presenceDecision.passRemark, reportCode),
            expectedId === ""
              ? `Test Data ไม่มี ${config.testDataIdField}`
              : "",
          ),

          matchedRowNumber: undefined,
          failedKeyFieldHeaders: [],
          reviewFieldHeaders: [],
          isExpectedAbsence: true,
        };
      }

      /**
       * พบ Report Row ทั้งที่ Business Rule กำหนด
       * ว่ารายการนี้ต้องไม่มี จึงถือว่า FAIL
       */
      return {
        testCaseNo: expectedCase.displayTestCaseNo,

        status: "FAIL",

        remark: withMissingTestNoRemark(
          this.formatRuleRemark(presenceDecision.failRemark, reportCode),

          `[${reportCode}] : ${config.referenceNumberReportField} = ` +
            `"${matchedRecord.get(config.referenceNumberReportField).trim()}"`,

          fallbackRemark !== ""
            ? fallbackRemark
            : expectedId === ""
              ? `Test Data ไม่มี ${config.testDataIdField}`
              : "",
        ),
        matchedRowNumber: matchedRecord.rowNumber,

        failedKeyFieldHeaders: [config.referenceNumberReportField],

        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    if (!matchedRecord) {
      const fallbackReason =
        expectedId === ""
          ? `Test Data ไม่มี ${config.testDataIdField}`
          : `ไม่พบ Exact Reference = "${normalizedExpectedReference}"`;

      return {
        testCaseNo: expectedCase.displayTestCaseNo,

        status: "FAIL",

        remark: withMissingTestNoRemark(
          fallbackReason,
          `ไม่พบแถว ${suffix} ใน ${reportCode} ทั้งจาก Exact Reference ` +
            "และ LTX Fallback Matching " +
            "(Account + Currency + DR/FE + Amount)",
        ),
        matchedRowNumber: undefined,
        failedKeyFieldHeaders: [],
        reviewFieldHeaders: [],
        isExpectedAbsence: false,
      };
    }

    const { status, failedKeyFieldHeaders, failRemarks } =
      this.checkKeyConditions(
        reportCode,
        config,
        expectedCase,
        suffix,
        matchedRecord,
      );
    const { reviewFieldHeaders, remark: reviewRemark } =
      this.buildReviewSection(
        reportCode,
        config,
        expectedCase,
        matchedRecord,
        suffix,
      );
    const remark = withMissingTestNoRemark(
      fallbackRemark,
      ...failRemarks,
      reviewRemark,
    );

    return {
      testCaseNo: expectedCase.displayTestCaseNo,
      status,
      remark,
      matchedRowNumber: matchedRecord.rowNumber,
      failedKeyFieldHeaders,
      reviewFieldHeaders,
      isExpectedAbsence: false,
    };
  }

  async reconcile(
    reportCode: string,
    testDataFilePath: string,
  ): Promise<string> {
    console.log(`\n===== RECONCILE - ${reportCode} =====`);

    const config = getReconcileConfig(reportCode);
    const { mappingHeaderRowNumber } = this.validateMappingConfiguration(
      reportCode,
      config,
    );
    const expectedCaseBuilder = this.getExpectedCaseBuilder(reportCode);

    const {
      workbook,
      reportWorksheet,
      resultSheet,
      reportHeaders,
      reconcileFilePath,
    } = await this.workbookPreparer.prepare(reportCode, mappingHeaderRowNumber);

    const { records: reportRecords } = this.excelReader.parseWorksheet(
      reportWorksheet,
      mappingHeaderRowNumber,
    );
    const testData = await this.excelReader.readFile(
      testDataFilePath,
      config.testDataHeaderRowNumber,
    );

    this.validatePresenceRuleHeaders(reportCode, testData.headers, config);

    const expectedCases = expectedCaseBuilder.build(
      testData.headers,
      testData.records,
      config,
    );

    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];

    const collect = (row: ResultRow): void => {
      if (row.matchedRowNumber !== undefined) {
        annotationByRowNumber.set(row.matchedRowNumber, row);
      } else {
        unmatchedRows.push(row);
      }
    };

    expectedCases.forEach((expectedCase) => {
      const expectedDrReference = expectedCase.expectedDrReference ?? "";

      const expectedFeReference = expectedCase.expectedFeReference ?? "";

      /**
       * Parse DR Amount ด้วย Logic กลางเดียวกับ
       * LtxExpectedCaseBuilder
       */
      const drPrimaryAmount = this.amountComparator.parse(
        expectedCase.primaryRecord.get(config.drAmountTestDataField),
      );

      const drFallbackAmount = this.amountComparator.parse(
        expectedCase.primaryRecord.get(config.drAmountFallbackTestDataField),
      );

      /**
       * Expected FE Amount เป็นผลรวม Fee Amount
       * ที่ LtxExpectedCaseBuilder คำนวณไว้แล้ว
       */
      const expectedFeAmount = this.amountComparator.parse(
        String(expectedCase.expectedFeAmount ?? ""),
      );

      /**
       * Transaction ID ว่างอาจทำให้ Expected Reference ว่าง
       * จึงใช้ Amount ตรวจว่ายังมี DR Slot
       * ที่ต้องนำไป Fallback Matching หรือไม่
       */
      const shouldResolveDr =
        expectedDrReference.trim() !== "" ||
        (drPrimaryAmount !== null && drPrimaryAmount > 0.01) ||
        (drFallbackAmount !== null && drFallbackAmount > 0.01);

      /**
       * มี FE Slot เมื่อมี Expected Reference
       * หรือ SUM Fee Amount มากกว่า 0.01
       */
      const shouldResolveFe =
        expectedFeReference.trim() !== "" ||
        (expectedFeAmount !== null && expectedFeAmount > 0.01);

      if (shouldResolveDr) {
        collect(
          this.resolveSlot(
            reportCode,
            config,
            expectedCase,
            expectedDrReference,
            config.drSuffixLabel,
            reportRecords,
          ),
        );
      }

      if (shouldResolveFe) {
        collect(
          this.resolveSlot(
            reportCode,
            config,
            expectedCase,
            expectedFeReference,
            config.feSuffixLabel,
            reportRecords,
          ),
        );
      }
    });

    const unexpectedMissingRows = unmatchedRows.filter(
      (row) => row.status === "FAIL",
    );

    if (unexpectedMissingRows.length > 0) {
      console.warn(
        `⚠️ พบ ${unexpectedMissingRows.length} แถวที่หาคู่กันใน ${reportCode} ` +
          `ไม่เจอและไม่เข้าเงื่อนไข Expected Absence`,
      );
    }

    this.sheetWriter.writeHeaderRow(resultSheet, reportHeaders);

    const lastRowNumber = this.sheetWriter.writeRowsInRequestedOrder(
      resultSheet,
      reportWorksheet,
      reportHeaders,
      mappingHeaderRowNumber + 1,
      reportWorksheet.rowCount,
      annotationByRowNumber,
      unmatchedRows,
    );

    this.sheetWriter.finalizeAutoFilter(
      resultSheet,
      reportHeaders,
      lastRowNumber - 1,
    );

    workbook.removeWorksheet(reportWorksheet.id);

    await workbook.xlsx.writeFile(reconcileFilePath);

    const allRows = [...annotationByRowNumber.values(), ...unmatchedRows];
    const totalPass = allRows.filter((row) => row.status === "PASS").length;
    const totalFail = allRows.filter((row) => row.status === "FAIL").length;
    const totalWithReview = allRows.filter(
      (row) => row.reviewFieldHeaders.length > 0,
    ).length;

    console.log(`Output File : ${reconcileFilePath}`);
    console.log(
      `Test Case ที่ประมวลผล : ${expectedCases.length} | แถวผลลัพธ์ทั้งหมด : ${allRows.length} | ` +
        `Pass : ${totalPass} | Fail : ${totalFail} | (ในจำนวนนี้มี field สีเหลืองต้อง review : ${totalWithReview} แถว)`,
    );

    return reconcileFilePath;
  }
}

/**
 * Backward-compatible function wrapper — ให้ tests/script3-reconcile.spec.ts เดิม
 * เรียกใช้ได้โดยไม่ต้องแก้ไฟล์ test เลย
 */
export const reconcileReport = (
  reportCode: string,
  testDataFilePath: string,
): Promise<string> => {
  return new ReconcileService().reconcile(reportCode, testDataFilePath);
};
