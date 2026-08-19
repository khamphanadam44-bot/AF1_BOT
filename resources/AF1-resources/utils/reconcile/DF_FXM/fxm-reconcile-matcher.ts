/**
 * fxm-reconcile-matcher.ts
 * ------------------------------------------------------------------
 * รับผิดชอบการอ่านข้อมูล ตรวจ Header ตัดสิน Presence Rule
 * และจับคู่ Test Data กับ DF_FXM Report
 * ------------------------------------------------------------------
 */

import {
    FXM_REPORT_CODE,
    FXM_REPORT_FIELDS,
    FXM_REPORT_HEADER_ROW,
    FXM_TEST_DATA_FIELDS,
    FXM_TEST_DATA_HEADER_ROW,
    FXM_USD_CURRENCY_CODE,
    FXM_USD_THRESHOLD,
} from "./fxm-config";

import {
    FXMDirection,
    FXMRuleEvaluator,
    normalizeFXMValue,
    parseFXMAmount,
} from "./fxm-rules";

import {
    FXM_AMOUNT_TOLERANCE,
    FXM_REQUIRED_TEST_DATA_HEADERS,
    FXMFallbackResult,
    FXMLoadedData,
    FXMMatchResolution,
    FXMReconcileDecision,
    FXMReturnTestInfo,
    parseFXMDateKey,
    parseFXMReturnTestInfo,
} from "./fxm-reconcile-model";

import {
    getUniqueMappingHeaders,
    requireMappingReportName,
} from "../../../config/mapping-helper";

import {
    canonicalHeader,
} from "../../validators/shared/header-matcher";

import {
    ReconcileExcelReader,
} from "../shared/excel-reader";

import {
    ReconcileRecord,
} from "../shared/record";

import {
    ReconcileWorkbookPreparer,
} from "../shared/workbook-preparer";

/**
 * Base class สำหรับขั้นตอน Load, Validate และ Matching
 * ใช้ protected เฉพาะ method ที่ class ลูกจำเป็นต้องเรียก
 */
export class FXMReconcileMatcher {
    /**
     * ใช้ Rule จาก FXM-rules.ts
     * เพื่อตรวจ Field และหาทิศทางธุรกรรม
     */
    private readonly ruleEvaluator =
        new FXMRuleEvaluator();

    /**
     * อ่าน Checked Report และ Test Data
       * พร้อมตรวจสอบ Header ก่อนเริ่ม Reconcile
       *
       * ขั้นตอน:
       * 1. Copy Checked Report ล่าสุดไปยัง Reconcile Folder
       * 2. เปิด Report Worksheet
       * 3. อ่านข้อมูล Report
       * 4. อ่านข้อมูล Test Data
       * 5. ตรวจ Header ทั้งสองฝั่ง
       * 6. สร้าง Arrangement Number Index
       * 7. จอง Report Row ที่มี Exact Transaction ID
       */
    async loadAndValidate(
        testDataFilePath: string,
    ): Promise<FXMLoadedData> {
        const workbookPreparer =
            new ReconcileWorkbookPreparer();

        const excelReader =
            new ReconcileExcelReader();

        console.log(
            `\n===== LOAD RECONCILE DATA - ${FXM_REPORT_CODE} =====`,
        );

        /**
         * หา Checked Report ล่าสุด
         * และ Copy ไปสร้างไฟล์ Reconcile Result
         */
        const prepared =
            await workbookPreparer.prepare(
                FXM_REPORT_CODE,
                FXM_REPORT_HEADER_ROW,
            );

        /**
         * อ่านข้อมูลจาก Report Worksheet
         *
         * Header ของ DF_FXM อยู่แถวที่ 1
         */
        const reportData =
            excelReader.parseWorksheet(
                prepared.reportWorksheet,
                FXM_REPORT_HEADER_ROW,
            );

        /**
         * อ่าน Test Data
         *
         * Header ของ Test Data อยู่แถวที่ 5
         */
        const testData =
            await excelReader.readFile(
                testDataFilePath,
                FXM_TEST_DATA_HEADER_ROW,
            );

        /**
         * ตรวจสอบ Header ก่อนเริ่ม Matching
         *
         * ถ้า Header หาย ระบบจะหยุดพร้อมแสดง
         * รายการ Header ที่ไม่พบ
         */
        this.validateHeaders(
            prepared.reportHeaders,
            testData.headers,
        );

        /**
         * สร้าง Index:
         *
         * Arrangement Number
         * → Report Record[]
         *
         * ใช้ Array เพราะ Arrangement Number
         * อาจเกิดค่าซ้ำใน Report
         */
        const reportIndex =
            this.indexReportRecords(
                reportData.records,
            );

        /**
         * จอง Report Row ที่สามารถ Exact Match ได้
         *
         * ป้องกันไม่ให้ Fallback Matching
         * ของ Test Case ก่อนหน้าใช้ Report Row
         * ที่เป็นของ Test Case ถัดไป
         */
        const reservedExactRowNumbers =
            this.reserveExactRows(
                testData.records,
                reportIndex,
            );

        console.log(
            `Report Record Count : ${reportData.records.length}`,
        );

        console.log(
            `Test Data Record Count : ${testData.records.length}`,
        );

        console.log(
            `Reserved Exact Row Count : ${reservedExactRowNumbers.size}`,
        );

        return {
            prepared,

            reportRecords:
                reportData.records,

            testDataRecords:
                testData.records,

            reportIndex,

            reservedExactRowNumbers,
        };
    }

    /**
     * ตัดสินว่า Test Data หนึ่งแถว
     * ต้องมีหรือไม่ต้องมีใน DF_FXM
     */
    evaluatePresence(
        testDataRecord: ReconcileRecord,
    ): FXMReconcileDecision {
        /**
         * ระบุทิศทางของธุรกรรมจาก:
         * - From Currency
         * - To Currency
         */
        const direction =
            this.ruleEvaluator.getDirection(
                testDataRecord,
            );

        /**
         * ตรวจ Field ที่จำเป็นสำหรับ Rule
         *
         * Test No. ไม่ใช้ตัดสิน Business Rule
         * จึงไม่รวมอยู่ในการตรวจส่วนนี้
         *
         * Transaction ID สามารถใช้ Fallback Matching ได้
         * จึงไม่ใช้ตัดสิน Presence Rule เช่นกัน
         */
        const missingFields =
            this.ruleEvaluator
                .validateRequiredRuleFields(
                    testDataRecord,
                );

        /**
         * อ่าน Settled Amount
         *
         * Requirement ใช้ Amount นี้
         * เทียบกับ USD Equivalent Amount ใน Report
         */
        const settledAmount =
            parseFXMAmount(
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .settledAmount,
                ),
            );
        /**
         * อ่าน Settled Currency และ Normalize เป็นตัวพิมพ์ใหญ่
         *
         * ตัวอย่าง:
         * "usd", " USD ", "Usd"
         * จะถูกแปลงเป็น "USD"
         */
        const settledCurrency =
            normalizeFXMValue(
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .settledCurrency,
                ),
            );

        /**
         * กรณี Field ว่าง
         * หรือ Settled Amount ไม่สามารถแปลงเป็นตัวเลขได้
         *
         * ระบบจะไม่เดาว่ารายการต้องมีหรือไม่มี
         */
        if (
            missingFields.length >
            0 ||
            settledAmount ===
            null
        ) {
            const invalidFields = [
                ...missingFields,

                ...(settledAmount ===
                    null
                    ? [
                        FXM_TEST_DATA_FIELDS
                            .settledAmount,
                    ]
                    : []),
            ].filter(
                (
                    field,
                    index,
                    allFields,
                ) =>
                    allFields.indexOf(
                        field,
                    ) ===
                    index,
            );

            return {
                direction,

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    "ข้อมูลไม่ครบหรือรูปแบบไม่ถูกต้อง: " +
                    invalidFields.join(
                        ", ",
                    ),
                ],

                requiresReview:
                    true,
            };
        }

        /**
         * From Currency และ To Currency เหมือนกัน
         *
         * ไม่ถือเป็นรายการแลกเปลี่ยนเงินตรา
         * จึงต้องไม่พบใน DF_FXM
         */
        if (
            direction ===
            "NOT_FX"
        ) {
            return {
                direction,

                expectation:
                    "MUST_NOT_EXIST",

                reasons: [
                    "From Currency และ To Currency เป็นสกุลเดียวกัน",
                ],

                requiresReview:
                    false,
            };
        }

        /**
 * ปัจจุบัน DF_FXM รองรับเฉพาะ
 * Settled Currency = USD
 *
 * ถ้าเป็นสกุลอื่น จะยังไม่นำ Settled Amount
 * ไปเปรียบเทียบกับ USD Equivalent Amount
 * เพราะยังไม่มี Currency Conversion Rule
 */
        if (
            settledCurrency !==
            FXM_USD_CURRENCY_CODE
        ) {
            return {
                direction,

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    `รองรับเฉพาะ ${FXM_USD_CURRENCY_CODE} แต่พบ ` +
                    `${FXM_TEST_DATA_FIELDS.settledCurrency} = ` +
                    `"${settledCurrency}"`,
                ],

                requiresReview:
                    true,
            };
        }

        /**
 * Amount ต่ำกว่า 1,000,000 USD
 *
 * ยอดไม่เข้าเงื่อนไขของ DF_FXM
 * จึงต้องไม่พบใน DF_FXM
 * และต้องพิจารณารายงานใน DF_FXU
 */
        if (
            settledAmount <
            FXM_USD_THRESHOLD
        ) {
            return {
                direction,

                expectation:
                    "MUST_NOT_EXIST",

                reasons: [
                    `Settled Amount ${settledAmount} < ` +
                    `${FXM_USD_THRESHOLD}`,
                ],

                requiresReview:
                    false,
            };
        }

        /**
         * Cross Currency:
         *
         * Source Currency และ Destination Currency ต่างกัน
         * แต่ไม่มีขา THB
         *
         * Requirement ระบุว่าต้องใช้ Settlement Currency
         * และ Payment Intermediary เช่น NIUM
         * ช่วยตัดสินว่าใครเป็นผู้ทำ FX Conversion
         *
         * ขณะนี้ Requirement ยังไม่มี Use Case ที่ยืนยัน
         * จึงห้ามสรุปว่ารายการต้องพบหรือไม่ต้องพบใน DF_FXM
         */
        if (
            direction ===
            "CROSS_CURRENCY"
        ) {
            return {
                direction,

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    "Cross Currency: ยังไม่สามารถตัดสินได้ เนื่องจาก " +
                    "Requirement ของ Settlement Currency และ " +
                    "Payment Intermediary เช่น NIUM ยังไม่ครบถ้วน",
                ],

                requiresReview:
                    true,
            };
        }


        /**
 * BUY_FCY หรือ SELL_FCY
 * และ Amount ตั้งแต่ 1,000,000 USD ขึ้นไป
 *
 * ต้องพบใน DF_FXM
 */
        return {
            direction,

            expectation:
                "MUST_EXIST",

            reasons: [],

            requiresReview:
                false,
        };

    }

    /**
 * ตัดสิน Presence Rule สำหรับ Return/Reversal
 * โดยอ้างอิงผลของ Original Transaction
 */
    private evaluateReturnPresence(
        returnRecord:
            ReconcileRecord,

        returnInfo:
            FXMReturnTestInfo,

        testDataIndex:
            ReadonlyMap<
                string,
                readonly ReconcileRecord[]
            >,
    ): FXMReconcileDecision {
        const originalTestNo =
            normalizeFXMValue(
                returnInfo.originalTestNo,
            );

        const originalCandidates =
            testDataIndex.get(
                originalTestNo,
            ) ?? [];

        /**
         * ไม่พบ Original Transaction
         */
        if (
            originalCandidates.length ===
            0
        ) {
            return {
                direction:
                    "UNKNOWN",

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    `Return/Reversal: ไม่พบ Original Test Case ` +
                    `"${returnInfo.originalTestNo}"`,
                ],

                requiresReview:
                    true,
            };
        }

        /**
         * Original Test No. ซ้ำ
         * ระบบจึงไม่เลือกแถวใดแถวหนึ่งเอง
         */
        if (
            originalCandidates.length >
            1
        ) {
            return {
                direction:
                    "UNKNOWN",

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    `Return/Reversal: พบ Original Test Case ` +
                    `"${returnInfo.originalTestNo}" ซ้ำ ` +
                    `${originalCandidates.length} แถว`,
                ],

                requiresReview:
                    true,
            };
        }

        const originalRecord =
            originalCandidates[0];

        const originalDecision =
            this.evaluatePresence(
                originalRecord,
            );

        /**
         * ถ้า Original ยังตัดสินไม่ได้
         * Return/Reversal ก็ยังตัดสินไม่ได้เช่นกัน
         */
        if (
            originalDecision.expectation ===
            "CANNOT_DECIDE"
        ) {
            return {
                direction:
                    originalDecision.direction,

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    `Return/Reversal: Original Test Case ` +
                    `"${returnInfo.originalTestNo}" ` +
                    "ยังไม่สามารถตัดสินได้",
                    ...originalDecision.reasons,
                ],

                requiresReview:
                    true,
            };
        }

        /**
         * Return/Reversal ใช้วันที่และ Amount
         * จากแถว Return/Reversal ของตัวเอง
         *
         * From/To Currency สามารถว่างได้
         * เพราะทิศทางจะย้อนกลับจาก Original
         */
        const requiredReturnFields = [
            FXM_TEST_DATA_FIELDS
                .transactionDate,
            FXM_TEST_DATA_FIELDS
                .settledCurrency,
            FXM_TEST_DATA_FIELDS
                .settledAmount,
        ];

        const invalidReturnFields =
            requiredReturnFields.filter(
                (
                    field,
                ) =>
                    normalizeFXMValue(
                        returnRecord.get(
                            field,
                        ),
                    ) === "",
            );

        const returnAmount =
            parseFXMAmount(
                returnRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .settledAmount,
                ),
            );

        if (
            returnAmount ===
            null &&
            !invalidReturnFields.includes(
                FXM_TEST_DATA_FIELDS
                    .settledAmount,
            )
        ) {
            invalidReturnFields.push(
                FXM_TEST_DATA_FIELDS
                    .settledAmount,
            );
        }

        if (
            invalidReturnFields.length >
            0
        ) {
            return {
                direction:
                    originalDecision.direction,

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    "Return/Reversal: ข้อมูลไม่ครบหรือรูปแบบไม่ถูกต้อง: " +
                    invalidReturnFields.join(
                        ", ",
                    ),
                ],

                requiresReview:
                    true,
            };
        }

        /**
         * ถ้า Original ไม่ต้องรายงานใน DF_FXM
         * Return/Reversal ก็ไม่ต้องรายงานตาม
         */
        if (
            originalDecision.expectation ===
            "MUST_NOT_EXIST"
        ) {
            return {
                direction:
                    originalDecision.direction,

                expectation:
                    "MUST_NOT_EXIST",

                reasons: [
                    `Return/Reversal: Original Test Case ` +
                    `"${returnInfo.originalTestNo}" ` +
                    "ไม่ต้องรายงานใน DF_FXM " +
                    `จึงไม่ต้องรายงานรายการ ${returnInfo.transactionType} ตาม`,
                ],

                requiresReview:
                    false,
            };
        }

        /**
         * Original ต้องรายงานใน DF_FXM
         * Return/Reversal ต้องรายงานเช่นกัน
         * แต่ต้องกลับทิศทาง FX
         *
         * Original BUY_FCY  → Return SELL_FCY
         * Original SELL_FCY → Return BUY_FCY
         */
        const reverseDirection:
            FXMDirection =
            originalDecision.direction ===
                "BUY_FCY"
                ? "SELL_FCY"
                : originalDecision.direction ===
                    "SELL_FCY"
                    ? "BUY_FCY"
                    : "UNKNOWN";

        if (
            reverseDirection ===
            "UNKNOWN"
        ) {
            return {
                direction:
                    reverseDirection,

                expectation:
                    "CANNOT_DECIDE",

                reasons: [
                    `Return/Reversal: ไม่สามารถกลับทิศทางของ ` +
                    `Original Test Case "${returnInfo.originalTestNo}" ได้`,
                ],

                requiresReview:
                    true,
            };
        }

        return {
            direction:
                reverseDirection,

            expectation:
                "MUST_EXIST",

            reasons: [],

            requiresReview:
                false,
        };
    }
    /**
     * จับคู่ Test Data หนึ่งแถวกับ DF_FXM Report
     *
     * ลำดับการทำงาน:
     * 1. ประเมิน Presence Rule
     * 2. ค้นหา Exact Match
     * 3. ตรวจ Duplicate Exact Match
     * 4. ถ้า Exact ไม่พบ ให้ลอง Fallback
     * 5. ตรวจ Ambiguous Fallback Match
     *
     * @param testDataIndex Index สำหรับค้นหา Original Test Case
     * @param testDataRecord Test Data หนึ่งแถว
     * @param reportRecords Report Records ทั้งหมด
     * @param reportIndex Arrangement Number Index
     * @param reservedExactRowNumbers Report Row ที่จองไว้
     * @param usedReportRowNumbers Report Row ที่ใช้ไปแล้ว
     */
    resolveMatch(
        testDataRecord:
            ReconcileRecord,

        /**
         * Index ของ Test Data ตาม Test No.
         *
         * ใช้ค้นหา Original Transaction
         * ของ Return/Reversal
         */
        testDataIndex:
            ReadonlyMap<
                string,
                readonly ReconcileRecord[]
            >,

        reportRecords:
            ReconcileRecord[],

        reportIndex:
            ReadonlyMap<
                string,
                ReconcileRecord[]
            >,

        reservedExactRowNumbers:
            ReadonlySet<number>,

        usedReportRowNumbers:
            ReadonlySet<number>,
    ): FXMMatchResolution {
        /**
         * อ่านและ Normalize Matching Key
         *
         * ตัวอย่าง:
         * " tx001 "
         * → "TX001"
         */
        const transactionId =
            normalizeFXMValue(
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .transactionId,
                ),
            );

        /**
         * หาเลขที่ใช้แสดงในผลลัพธ์
         */
        const testCaseNo =
            this.getTestCaseNo(
                testDataRecord,
                transactionId,
            );

        /**
         * ตัดสินก่อนว่ารายการควรมี
         * หรือไม่ควรมีใน DF_FXM
         */
        const returnInfo =
            parseFXMReturnTestInfo(
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .testNo,
                ),
            );

        const decision =
            returnInfo ===
                null
                ? this.evaluatePresence(
                    testDataRecord,
                )
                : this.evaluateReturnPresence(
                    testDataRecord,
                    returnInfo,
                    testDataIndex,
                );

        /**
         * Exact Matching:
         *
         * Test Data:
         * Transaction ID/ Reconcile ID
         *
         * Report:
         * Arrangement Number
         */
        const exactCandidates =
            transactionId ===
                ""
                ? []
                : (
                    reportIndex.get(
                        transactionId,
                    ) ??
                    []
                );

        /**
         * Exact Match พบมากกว่าหนึ่งแถว
         *
         * ไม่เลือกแถวแรกให้อัตโนมัติ
         * เพราะอาจจับคู่ผิด Transaction
         */
        if (
            exactCandidates.length >
            1
        ) {
            return {
                testCaseNo,

                transactionId,

                decision,

                matchMode:
                    "AMBIGUOUS",

                candidateCount:
                    exactCandidates.length,

                remark:
                    `พบ ${FXM_REPORT_FIELDS.arrangementNumber} ` +
                    `= "${transactionId}" ซ้ำ ` +
                    `${exactCandidates.length} แถวใน DF_FXM Report`,
            };
        }

        /**
         * Exact Match พบหนึ่งแถว
         */
        const exactMatchedRecord =
            exactCandidates[0];

        if (
            exactMatchedRecord
        ) {
            /**
             * Report Row นี้ถูกใช้กับ Test Case ก่อนหน้าแล้ว
             *
             * ระบบจะไม่ใช้แถวเดิมซ้ำ
             */
            if (
                usedReportRowNumbers.has(
                    exactMatchedRecord.rowNumber,
                )
            ) {
                return {
                    testCaseNo,

                    transactionId,

                    decision,

                    matchMode:
                        "ROW_ALREADY_USED",

                    candidateCount:
                        1,

                    remark:
                        `Report Row ${exactMatchedRecord.rowNumber} ` +
                        "ถูกจับคู่กับ Test Case ก่อนหน้าแล้ว",
                };
            }

            return {
                testCaseNo,

                transactionId,

                decision,

                matchMode:
                    "EXACT",

                matchedRecord:
                    exactMatchedRecord,

                candidateCount:
                    1,

                remark:
                    `Exact Match: ` +
                    `${FXM_TEST_DATA_FIELDS.transactionId} ` +
                    `ตรงกับ ${FXM_REPORT_FIELDS.arrangementNumber}`,
            };
        }

        /**
         * ถ้าข้อมูลไม่เพียงพอสำหรับตัดสิน Presence Rule
         * จะไม่ทำ Fallback ต่อ
         *
         * เพราะ Fallback ต้องใช้:
         * - Txn Date
         * - Settled Amount
         */
        if (
            decision.expectation ===
            "CANNOT_DECIDE"
        ) {
            return {
                testCaseNo,

                transactionId,

                decision,

                matchMode:
                    "FALLBACK_UNAVAILABLE",

                candidateCount:
                    0,

                remark: [
                    "ไม่สามารถทำ Fallback Matching ได้",
                    ...decision.reasons,
                ].join(
                    "\n",
                ),
            };
        }

        /**
         * Exact Match ไม่พบ
         * จึงค้นหาด้วย Matching Support
         */
        const fallbackResult =
            this.findUniqueFallbackMatch(
                testDataRecord,
                reportRecords,
                reservedExactRowNumbers,
                usedReportRowNumbers,
            );

        /**
         * Fallback พบ Candidate เพียงหนึ่งแถว
         */
        if (
            fallbackResult
                .matchedRecord
        ) {
            return {
                testCaseNo,

                transactionId,

                decision,

                matchMode:
                    "FALLBACK",

                matchedRecord:
                    fallbackResult
                        .matchedRecord,

                candidateCount:
                    1,

                remark:
                    fallbackResult.remark,
            };
        }

        /**
         * Fallback พบ Candidate มากกว่าหนึ่งแถว
         */
        if (
            (
                fallbackResult
                    .candidateCount ??
                0
            ) >
            1
        ) {
            return {
                testCaseNo,

                transactionId,

                decision,

                matchMode:
                    "AMBIGUOUS",

                candidateCount:
                    fallbackResult
                        .candidateCount ??
                    0,

                remark:
                    fallbackResult.remark,
            };
        }

        /**
         * Fallback ไม่สามารถทำงานได้
         * เพราะ Txn Date หรือ Settled Amount ไม่ถูกต้อง
         */
        if (
            fallbackResult
                .candidateCount ===
            undefined
        ) {
            return {
                testCaseNo,

                transactionId,

                decision,

                matchMode:
                    "FALLBACK_UNAVAILABLE",

                candidateCount:
                    0,

                remark:
                    fallbackResult.remark,
            };
        }

        /**
         * ไม่พบทั้ง Exact Match และ Fallback Match
         */
        return {
            testCaseNo,

            transactionId,

            decision,

            matchMode:
                "NOT_FOUND",

            candidateCount:
                0,

            remark:
                fallbackResult.remark,
        };
    }

    /**
     * ค้นหา Report Row ด้วย Matching Support
     *
     * Test Data:
     * - Txn Date
     * - Settled Amount (CCY)
     *
     * Report:
     * - Data Set Date
     * - USD Equivalent Amount
     *
     * Fallback จะสำเร็จเมื่อพบ Candidate
     * เพียงหนึ่งแถวเท่านั้น
     */
    private findUniqueFallbackMatch(
        testDataRecord:
            ReconcileRecord,

        reportRecords:
            ReconcileRecord[],

        reservedExactRowNumbers:
            ReadonlySet<number>,

        usedReportRowNumbers:
            ReadonlySet<number>,
    ): FXMFallbackResult {
        /**
         * อ่านวันที่จาก Test Data
         */
        const expectedDateText =
            testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .transactionDate,
            );

        const expectedDate =
            parseFXMDateKey(
                expectedDateText,
            );

        /**
         * อ่าน Amount จาก Test Data
         */
        const expectedAmountText =
            testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .settledAmount,
            );

        const expectedAmount =
            parseFXMAmount(
                expectedAmountText,
            );

        const invalidFields:
            string[] =
            [];

        if (
            expectedDate ===
            null
        ) {
            invalidFields.push(
                FXM_TEST_DATA_FIELDS
                    .transactionDate,
            );
        }

        if (
            expectedAmount ===
            null
        ) {
            invalidFields.push(
                FXM_TEST_DATA_FIELDS
                    .settledAmount,
            );
        }

        /**
         * ไม่มีข้อมูลเพียงพอสำหรับทำ Fallback
         *
         * candidateCount จะเป็น undefined
         * เพื่อแยกจากกรณีค้นหาแล้วไม่พบ Candidate
         */
        if (
            expectedDate ===
            null ||
            expectedAmount ===
            null
        ) {
            return {
                remark:
                    "Fallback Matching ทำไม่ได้ เพราะ Field " +
                    "ว่างหรือรูปแบบไม่ถูกต้อง: " +
                    invalidFields.join(
                        ", ",
                    ),
            };
        }

        /**
         * ค้นหา Candidate ที่มี:
         *
         * Data Set Date = Txn Date
         *
         * และ
         *
         * USD Equivalent Amount = Settled Amount
         */
        const candidates =
            reportRecords.filter(
                (
                    reportRecord,
                ) => {
                    /**
                     * ห้ามใช้แถวที่ถูกจองไว้สำหรับ Exact Match
                     */
                    if (
                        reservedExactRowNumbers.has(
                            reportRecord.rowNumber,
                        )
                    ) {
                        return false;
                    }

                    /**
                     * ห้ามนำ Report Row ที่ใช้แล้วกลับมาใช้ซ้ำ
                     */
                    if (
                        usedReportRowNumbers.has(
                            reportRecord.rowNumber,
                        )
                    ) {
                        return false;
                    }

                    const actualDate =
                        parseFXMDateKey(
                            reportRecord.get(
                                FXM_REPORT_FIELDS
                                    .dataSetDate,
                            ),
                        );

                    const actualAmount =
                        parseFXMAmount(
                            reportRecord.get(
                                FXM_REPORT_FIELDS
                                    .usdEquivalentAmount,
                            ),
                        );

                    if (
                        actualDate !==
                        expectedDate
                    ) {
                        return false;
                    }

                    if (
                        actualAmount ===
                        null
                    ) {
                        return false;
                    }

                    return (
                        Math.abs(
                            actualAmount -
                            expectedAmount,
                        ) <=
                        FXM_AMOUNT_TOLERANCE
                    );
                },
            );

        /**
         * พบ Candidate หนึ่งแถว
         *
         * ถือว่า Fallback Matching สำเร็จ
         */
        if (
            candidates.length ===
            1
        ) {
            return {
                matchedRecord:
                    candidates[0],

                candidateCount:
                    1,

                remark:
                    "Primary Matching Key ไม่พบ แต่ Fallback Match สำเร็จจาก " +
                    `${FXM_TEST_DATA_FIELDS.transactionDate} + ` +
                    FXM_TEST_DATA_FIELDS
                        .settledAmount,
            };
        }

        /**
         * ไม่พบ Candidate
         */
        if (
            candidates.length ===
            0
        ) {
            return {
                candidateCount:
                    0,

                remark:
                    "Fallback Matching ไม่พบข้อมูลจาก " +
                    `${FXM_TEST_DATA_FIELDS.transactionDate} + ` +
                    FXM_TEST_DATA_FIELDS
                        .settledAmount,
            };
        }

        /**
         * พบ Candidate มากกว่าหนึ่งแถว
         *
         * ไม่เลือกแถวแรกให้อัตโนมัติ
         */
        return {
            candidateCount:
                candidates.length,

            remark:
                `Ambiguous Fallback Match: พบ ${candidates.length} แถวจาก ` +
                `${FXM_TEST_DATA_FIELDS.transactionDate} + ` +
                FXM_TEST_DATA_FIELDS
                    .settledAmount,
        };
    }

    /**
     * หาเลขที่ใช้แสดงใน Column Test Script No.
     *
     * ลำดับ:
     * 1. Test No.
     * 2. Transaction ID
     * 3. TEST DATA ROW ตามด้วยเลขแถวจริง
     */
    private getTestCaseNo(
        testDataRecord:
            ReconcileRecord,

        transactionId:
            string,
    ): string {
        const testNo =
            String(
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .testNo,
                ),
            ).trim();

        if (
            testNo !==
            ""
        ) {
            return testNo;
        }

        if (
            transactionId !==
            ""
        ) {
            return transactionId;
        }

        return (
            `TEST DATA ROW ` +
            testDataRecord.rowNumber
        );
    }

    /**
     * ตรวจสอบ Header ของ Report และ Test Data
     */
    private validateHeaders(
        reportHeaders: string[],
        testDataHeaders: string[],
    ): void {
        /**
         * ยืนยันว่า DF_FXM
         * มี Config อยู่ใน mapping-config.ts
         */
        const reportName =
            requireMappingReportName(
                FXM_REPORT_CODE,
            );

        /**
         * อ่าน Header ทุกกลุ่มจาก Mapping Config:
         *
         * - matchingKey
         * - core
         * - customer
         * - conditions
         * - reference
         */
        const requiredReportHeaders =
            getUniqueMappingHeaders(
                reportName,
            );

        this.assertHeaders(
            reportHeaders,
            requiredReportHeaders,
            "Raw Report",
        );

        this.assertHeaders(
            testDataHeaders,
            FXM_REQUIRED_TEST_DATA_HEADERS,
            "Test Data",
        );
    }

    /**
     * ตรวจว่า Header ที่จำเป็นมีอยู่จริงหรือไม่
     *
     * การเปรียบเทียบใช้ canonicalHeader()
     * จึงไม่สนตัวพิมพ์เล็ก/ใหญ่
     * และช่องว่างส่วนเกิน
     */
    private assertHeaders(
        actualHeaders: string[],
        requiredHeaders:
            readonly string[],
        sourceName: string,
    ): void {
        /**
         * สร้าง Set ของ Header ที่พบจริง
         *
         * ตัวอย่าง:
         * "Arrangement Number"
         * → "arrangement number"
         */
        const actualHeaderSet =
            new Set(
                actualHeaders
                    .filter(
                        (
                            header,
                        ) =>
                            header.trim() !==
                            "",
                    )
                    .map(
                        canonicalHeader,
                    ),
            );

        /**
         * ค้นหา Header ที่ Config ต้องการ
         * แต่ไม่พบในไฟล์จริง
         */
        const missingHeaders =
            requiredHeaders.filter(
                (
                    requiredHeader,
                ) =>
                    !actualHeaderSet.has(
                        canonicalHeader(
                            requiredHeader,
                        ),
                    ),
            );

        if (
            missingHeaders.length >
            0
        ) {
            throw new Error(
                `[${FXM_REPORT_CODE}] ` +
                `${sourceName} missing header(s): ` +
                missingHeaders.join(
                    ", ",
                ),
            );
        }
    }

    /**
     * สร้าง Index ของ Test Data จาก Test No.
     *
     * Key:
     * Test No. ที่ Normalize เป็นตัวพิมพ์ใหญ่แล้ว
     *
     * Value:
     * Test Data Record ที่ใช้ Test No. เดียวกัน
     *
     * ใช้ Array เพื่อให้ตรวจพบกรณี
     * Test No. ซ้ำมากกว่าหนึ่งแถวได้
     */
    protected indexTestDataRecordsByTestNo(
        testDataRecords:
            readonly ReconcileRecord[],
    ): Map<
        string,
        ReconcileRecord[]
    > {
        const testDataIndex =
            new Map<
                string,
                ReconcileRecord[]
            >();

        for (
            const testDataRecord of
            testDataRecords
        ) {
            const testNo =
                normalizeFXMValue(
                    testDataRecord.get(
                        FXM_TEST_DATA_FIELDS
                            .testNo,
                    ),
                );

            /**
             * แถวที่ไม่มี Test No.
             * จะไม่สามารถเป็น Original
             * ของ Return/Reversal ได้
             */
            if (
                testNo ===
                ""
            ) {
                continue;
            }

            const existingRecords =
                testDataIndex.get(
                    testNo,
                ) ??
                [];

            existingRecords.push(
                testDataRecord,
            );

            testDataIndex.set(
                testNo,
                existingRecords,
            );
        }

        return testDataIndex;
    }

    /**
     * สร้าง Index ของ Report จาก Arrangement Number
     *
     * ตัวอย่าง:
     *
     * Map {
     *   "TX001" → [Report Row 2],
     *   "TX002" → [Report Row 3, Report Row 4]
     * }
     *
     * ใช้ Array เพื่อให้ตรวจพบ
     * Arrangement Number ซ้ำใน Report ได้
     */
    private indexReportRecords(
        reportRecords:
            ReconcileRecord[],
    ): Map<
        string,
        ReconcileRecord[]
    > {
        const reportIndex =
            new Map<
                string,
                ReconcileRecord[]
            >();

        for (
            const reportRecord of
            reportRecords
        ) {
            /**
             * อ่าน Arrangement Number
             * และ Normalize ก่อนใช้เป็น Map Key
             */
            const arrangementNumber =
                String(
                    reportRecord.get(
                        FXM_REPORT_FIELDS
                            .arrangementNumber,
                    ),
                )
                    .trim()
                    .toUpperCase();

            /**
             * Report Row ที่ไม่มี Arrangement Number
             * จะไม่ถูกเพิ่มเข้า Index
             *
             * แต่ Row ยังอยู่ใน reportRecords
             * เพื่อให้สามารถแสดงในผลลัพธ์ได้
             */
            if (
                arrangementNumber ===
                ""
            ) {
                continue;
            }

            const existingRecords =
                reportIndex.get(
                    arrangementNumber,
                ) ??
                [];

            existingRecords.push(
                reportRecord,
            );

            reportIndex.set(
                arrangementNumber,
                existingRecords,
            );
        }

        return reportIndex;
    }

    /**
     * จอง Report Row ที่มี Exact Transaction ID
     *
     * เหตุผล:
     *
     * สมมุติ Test Case A ไม่มี Transaction ID
     * และ Fallback ของ A สามารถ Match Report Row 10 ได้
     *
     * แต่ Report Row 10 มี Arrangement Number
     * ที่ตรงกับ Transaction ID ของ Test Case B
     *
     * ถ้าไม่ Reserve ไว้ก่อน
     * Test Case A อาจแย่ง Report Row ของ B ไปใช้
     */
    private reserveExactRows(
        testDataRecords:
            ReconcileRecord[],

        reportIndex:
            ReadonlyMap<
                string,
                ReconcileRecord[]
            >,
    ): Set<number> {
        const reservedRowNumbers =
            new Set<number>();

        for (
            const testDataRecord of
            testDataRecords
        ) {
            /**
             * Primary Matching:
             *
             * Test Data:
             * Transaction ID/ Reconcile ID
             *
             * Report:
             * Arrangement Number
             */
            const transactionId =
                String(
                    testDataRecord.get(
                        FXM_TEST_DATA_FIELDS
                            .transactionId,
                    ),
                )
                    .trim()
                    .toUpperCase();

            if (
                transactionId ===
                ""
            ) {
                continue;
            }

            const matchedRecords =
                reportIndex.get(
                    transactionId,
                ) ??
                [];

            /**
             * จองทุก Row ที่มี Arrangement Number ตรงกัน
             *
             * ถ้ามีค่าซ้ำ จะจองทุกแถวไว้ก่อน
             * และตัดสิน Duplicate Matching ในขั้นถัดไป
             */
            for (
                const matchedRecord of
                matchedRecords
            ) {
                reservedRowNumbers.add(
                    matchedRecord.rowNumber,
                );
            }
        }

        return reservedRowNumbers;
    }
}