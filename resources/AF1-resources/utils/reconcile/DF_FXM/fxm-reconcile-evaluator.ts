/**
 * fxm-reconcile-evaluator.ts
 * ------------------------------------------------------------------
 * รับผิดชอบเปรียบเทียบ Field และแปลงผล Matching เป็น ResultRow
 * ------------------------------------------------------------------
 */

import {
    FXM_ARRANGEMENT_TYPE,
    FXM_LEG_TYPES,
    FXM_LEG_TYPE_NAMES,
    FXM_REPORT_CODE,
    FXM_REPORT_FIELDS,
    FXM_TEST_DATA_FIELDS,
    FXM_USD_THRESHOLD,
} from "./fxm-config";

import {
    normalizeFXMValue,
    parseFXMAmount,
} from "./fxm-rules";

import {
    FXM_AMOUNT_TOLERANCE,
    FXMCompareOptions,
    FXMFieldComparisonResult,
    FXMMatchResolution,
    FXMReconcileDecision,
    parseFXMDateKey,
    parseFXMReferenceDateKey,
} from "./fxm-reconcile-model";

import {
    FXMReconcileMatcher,
} from "./fxm-reconcile-matcher";

import {
    ReconcileRecord,
} from "../shared/record";

import {
    formatCompareRemark,
    formatFixedValueRemark,
} from "../shared/remark";

import {
    ResultRow,
} from "../shared/result-writer";

/**
 * สืบทอดขั้นตอน Matching แล้วเพิ่ม Field Comparison และ Result Decision
 */
export class FXMReconcileEvaluator extends FXMReconcileMatcher {

    /**
 * เปรียบเทียบ Core Fields หลังจับคู่
 * Test Data กับ Report ได้แล้ว
 *
 * Field ที่ตรวจ:
 *
 * 1. Transaction ID
 *    → Arrangement Number
 *
 * 2. Txn Date
 *    → Data Set Date
 *
 * 3. Settled Amount
 *    → USD Equivalent Amount
 *
 * 4. Arrangement Type
 *    → ต้องเป็น 018101
 *
 * 5. From/To Currency
 *    → Leg Type
 *
 * 6. From/To Currency
 *    → Leg Type Name
 */
    compareMatchedFields(
        testDataRecord:
            ReconcileRecord,

        reportRecord:
            ReconcileRecord,

        decision:
            FXMReconcileDecision,

        options:
            FXMCompareOptions =
            {},
    ): FXMFieldComparisonResult {
        /**
         * ใช้ Set เพื่อป้องกันชื่อ Header ซ้ำ
         */
        const failedHeaders =
            new Set<string>(
                options
                    .initialFailedHeaders ??
                [],
            );

        const reviewHeaders =
            new Set<string>(
                options
                    .initialReviewHeaders ??
                [],
            );

        const remarks:
            string[] =
            [
                ...(
                    options
                        .initialRemarks ??
                    []
                ),
            ].filter(
                (
                    remark,
                ) =>
                    remark.trim() !==
                    "",
            );

        /**
         * เพิ่มผลต่างของ Field
         *
         * mode = FAIL:
         * - เพิ่ม Header เข้า failedHeaders
         * - มีผลต่อสถานะ Test Case
         *
         * mode = REVIEW:
         * - เพิ่ม Header เข้า reviewHeaders
         * - ไม่มีผลต่อสถานะ Test Case
         */
        const addDifference = (
            mode:
                "FAIL" |
                "REVIEW",

            reportField:
                string,

            testDataField:
                string,

            expectedValue:
                string,

            actualValue:
                string,
        ): void => {
            if (
                mode ===
                "FAIL"
            ) {
                failedHeaders.add(
                    reportField,
                );
            } else {
                reviewHeaders.add(
                    reportField,
                );
            }

            remarks.push(
                formatCompareRemark(
                    FXM_REPORT_CODE,
                    testDataField,
                    expectedValue,
                    reportField,
                    actualValue,
                ),
            );
        };

        /**
         * ====================================================
         * Rule 1:
         * Transaction ID → Arrangement Number
         * ====================================================
         */
        const expectedTransactionId =
            testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .transactionId,
            );

        const actualArrangementNumber =
            reportRecord.get(
                FXM_REPORT_FIELDS
                    .arrangementNumber,
            );

        /**
         * Matching Key ว่างหรือค่าไม่ตรงกัน
         * ให้ Arrangement Number เป็น FAIL
         *
         * กรณี Fallback พบแถวได้
         * แต่ Transaction ID ไม่ตรง
         * Test Case จะยังคงเป็น FAIL
         */
        if (
            normalizeFXMValue(
                expectedTransactionId,
            ) ===
            "" ||
            normalizeFXMValue(
                expectedTransactionId,
            ) !==
            normalizeFXMValue(
                actualArrangementNumber,
            )
        ) {
            addDifference(
                "FAIL",

                FXM_REPORT_FIELDS
                    .arrangementNumber,

                FXM_TEST_DATA_FIELDS
                    .transactionId,

                expectedTransactionId,

                actualArrangementNumber,
            );
        }

        /**
  * ====================================================
  * Rule 2:
  * Txn Date → Data Set Date
  *
  * ลำดับการตรวจ:
  * 1. Txn Date ตรงกับ Data Set Date
  *    → PASS โดยไม่เพิ่ม Remark
  *
  * 2. Txn Date ไม่ตรงกับ Data Set Date
  *    แต่ตรงกับวันที่ใน Transaction ID/Reconcile ID
  *    → PASS พร้อม Remark
  *
  * 3. Txn Date ไม่ตรงทั้ง Data Set Date
  *    และวันที่ใน Transaction ID/Reconcile ID
  *    → FAIL พร้อม Remark
  *
  * 4. ไม่สามารถอ่านวันที่จาก Ref ได้
  *    → FAIL พร้อม Remark
  * ====================================================
  */
        const expectedDateText =
            testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .transactionDate,
            );

        const actualDateText =
            reportRecord.get(
                FXM_REPORT_FIELDS
                    .dataSetDate,
            );

        const expectedDate =
            parseFXMDateKey(
                expectedDateText,
            );

        const actualDate =
            parseFXMDateKey(
                actualDateText,
            );

        /**
         * อ่านวันที่จาก Transaction ID/Reconcile ID
         *
         * ใช้เป็นตัวตรวจสอบสำรองเมื่อ
         * Txn Date ไม่ตรงกับ Data Set Date
         */
        const referenceDate =
            parseFXMReferenceDateKey(
                expectedTransactionId,
            );

        /**
         * Txn Date อ่านไม่ได้
         *
         * ไม่มีวันที่หลักสำหรับนำไปเปรียบเทียบ
         * จึงให้ Data Set Date เป็น FAIL
         */
        if (
            expectedDate ===
            null
        ) {
            failedHeaders.add(
                FXM_REPORT_FIELDS
                    .dataSetDate,
            );

            remarks.push(
                `ไม่สามารถอ่านค่า ` +
                `${FXM_TEST_DATA_FIELDS.transactionDate} ` +
                `จาก Test Data ได้: ` +
                `"${expectedDateText}"`,
            );
        } else if (
            actualDate ===
            expectedDate
        ) {
            /**
             * Txn Date ตรงกับ Data Set Date
             *
             * ถือว่าผ่านและไม่ต้องเพิ่ม Remark
             */
        } else if (
            referenceDate ===
            null
        ) {
            /**
             * Txn Date ไม่ตรงกับ Data Set Date
             * และ Ref ไม่มีวันที่หรืออ่านรูปแบบไม่ได้
             */
            failedHeaders.add(
                FXM_REPORT_FIELDS
                    .dataSetDate,
            );

            remarks.push(
                `${FXM_TEST_DATA_FIELDS.transactionDate} ` +
                `ไม่ตรงกับ ${FXM_REPORT_FIELDS.dataSetDate} ` +
                `และไม่สามารถอ่านวันที่จาก ` +
                `${FXM_TEST_DATA_FIELDS.transactionId} ได้\n` +
                `Txn Date: "${expectedDateText}"\n` +
                `Data Set Date: "${actualDateText}"\n` +
                `Ref: "${expectedTransactionId}"`,
            );
        } else if (
            referenceDate ===
            expectedDate
        ) {
            /**
             * Txn Date ไม่ตรงกับ Data Set Date
             * แต่วันที่จาก Ref ตรงกับ Txn Date
             *
             * ไม่เพิ่ม failedHeaders
             * จึงสามารถเป็น PASS ได้
             * พร้อมแสดง Remark แจ้งความแตกต่าง
             */
            remarks.push(
                `${FXM_TEST_DATA_FIELDS.transactionDate} ` +
                `ไม่ตรงกับ ${FXM_REPORT_FIELDS.dataSetDate} ` +
                `แต่ตรงกับวันที่ใน ` +
                `${FXM_TEST_DATA_FIELDS.transactionId}\n` +
                `Txn Date: "${expectedDateText}"\n` +
                `Data Set Date: "${actualDateText}"\n` +
                `Ref Date: "${referenceDate}"`,
            );
        } else {
            /**
             * Txn Date ไม่ตรงทั้ง Data Set Date
             * และวันที่ที่อ่านได้จาก Ref
             */
            failedHeaders.add(
                FXM_REPORT_FIELDS
                    .dataSetDate,
            );

            remarks.push(
                `${FXM_TEST_DATA_FIELDS.transactionDate} ` +
                `ไม่ตรงทั้ง ${FXM_REPORT_FIELDS.dataSetDate} ` +
                `และวันที่ใน ` +
                `${FXM_TEST_DATA_FIELDS.transactionId}\n` +
                `Txn Date: "${expectedDateText}"\n` +
                `Data Set Date: "${actualDateText}"\n` +
                `Ref Date: "${referenceDate}"`,
            );
        }
        /**
         * ====================================================
         * Rule 3:
         * Settled Amount → USD Equivalent Amount
         * ====================================================
         */
        const expectedAmountText =
            testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .settledAmount,
            );

        const actualAmountText =
            reportRecord.get(
                FXM_REPORT_FIELDS
                    .usdEquivalentAmount,
            );

        const expectedAmount =
            parseFXMAmount(
                expectedAmountText,
            );

        const actualAmount =
            parseFXMAmount(
                actualAmountText,
            );

        /**
         * กรณีต่อไปนี้ให้ FAIL:
         *
         * - Settled Amount อ่านเป็นตัวเลขไม่ได้
         * - USD Equivalent Amount อ่านเป็นตัวเลขไม่ได้
         * - ผลต่างมากกว่า Tolerance 0.01
         */
        if (
            expectedAmount ===
            null ||
            actualAmount ===
            null ||
            Math.abs(
                expectedAmount -
                actualAmount,
            ) >
            FXM_AMOUNT_TOLERANCE
        ) {
            addDifference(
                "FAIL",

                FXM_REPORT_FIELDS
                    .usdEquivalentAmount,

                FXM_TEST_DATA_FIELDS
                    .settledAmount,

                expectedAmountText,

                actualAmountText,
            );
        }

        /**
         * ====================================================
         * Rule 4:
         * Arrangement Type ต้องเป็น 018101
         * ====================================================
         */
        const actualArrangementType =
            reportRecord.get(
                FXM_REPORT_FIELDS
                    .arrangementType,
            );

        if (
            normalizeFXMValue(
                actualArrangementType,
            ) !==
            normalizeFXMValue(
                FXM_ARRANGEMENT_TYPE,
            )
        ) {
            failedHeaders.add(
                FXM_REPORT_FIELDS
                    .arrangementType,
            );

            remarks.push(
                formatFixedValueRemark(
                    FXM_REPORT_CODE,

                    FXM_REPORT_FIELDS
                        .arrangementType,

                    actualArrangementType,

                    FXM_ARRANGEMENT_TYPE,
                ),
            );
        }

        /**
         * ====================================================
         * Rule 5 และ 6:
         * ตรวจ Leg Type และ Leg Type Name
         * ====================================================
         */
        this.compareLegFields(
            testDataRecord,
            reportRecord,
            decision,
            addDifference,
        );

        /**
         * Cross Currency:
         *
         * Requirement ระบุว่าต้องตรวจ
         * Settlement/Intermediary เช่น NIUM
         *
         * แต่ยังไม่มี Field และเงื่อนไขครบถ้วน
         * จึง Highlight Fi Arrangement Type Name เป็น Review
         */
        if (
            decision.direction ===
            "CROSS_CURRENCY"
        ) {
            reviewHeaders.add(
                FXM_REPORT_FIELDS
                    .arrangementTypeName
            );

            remarks.push(
                "Cross Currency: Please review Settlement/Intermediary rule; " +
                "NIUM และ Settlement/Intermediary logic ยังไม่ได้รับการยืนยัน",
            );
        }

        /**
         * กรณีข้อมูลไม่เพียงพอ
         * ให้เพิ่มเหตุผลจาก Presence Decision
         */
        if (
            decision.requiresReview &&
            decision.reasons.length >
            0
        ) {
            remarks.push(
                ...decision.reasons,
            );
        }

        return {
            failedHeaders: [
                ...failedHeaders,
            ],

            reviewHeaders: [
                ...reviewHeaders,
            ],

            /**
             * ตัด Remark ที่ซ้ำกันออก
             */
            remarks: [
                ...new Set(
                    remarks
                        .map(
                            (
                                remark,
                            ) =>
                                remark.trim(),
                        )
                        .filter(
                            Boolean,
                        ),
                ),
            ],
        };
    }

    /**
     * ตรวจ Leg Type และ Leg Type Name
     * จากทิศทางของธุรกรรม
     *
     * BUY_FCY:
     * - Leg Type = 182001
     * - ตัวแทนรับอนุญาตซื้อเงินตราต่างประเทศแลกกับสกุลเงินบาท
     *
     * SELL_FCY:
     * - Leg Type = 182002
     * - ตัวแทนรับอนุญาตขายเงินตราต่างประเทศแลกกับสกุลเงินบาท
     *
     * CROSS_CURRENCY / NOT_FX / UNKNOWN:
     * ยังไม่สามารถระบุ Leg ที่คาดหวังได้แน่นอน
     * จึงให้ Review แทน FAIL
     */
    private compareLegFields(
        testDataRecord:
            ReconcileRecord,

        reportRecord:
            ReconcileRecord,

        decision:
            FXMReconcileDecision,

        addDifference: (
            mode:
                "FAIL" |
                "REVIEW",

            reportField:
                string,

            testDataField:
                string,

            expectedValue:
                string,

            actualValue:
                string,
        ) => void,
    ): void {
        /**
         * ระบุ Leg Type ที่คาดหวัง
         * จากทิศทางของ Transaction
         */
        const expectedLegType =
            decision.direction ===
                "BUY_FCY"
                ? FXM_LEG_TYPES
                    .buyForeignCurrency
                : decision.direction ===
                    "SELL_FCY"
                    ? FXM_LEG_TYPES
                        .sellForeignCurrency
                    : undefined;

        /**
         * ระบุ Leg Type Name ที่คาดหวัง
         *
         * FXM_LEG_TYPE_NAMES ใช้ Leg Type Code เป็น Key
         *
         * ตัวอย่าง:
         * FXM_LEG_TYPE_NAMES["182001"]
         * FXM_LEG_TYPE_NAMES["182002"]
         *
         * จึงใช้ expectedLegType ที่หาได้ด้านบน
         * เป็น Key สำหรับอ่าน Leg Type Name
         */
        const expectedLegTypeName =
            expectedLegType ===
                undefined
                ? undefined
                : FXM_LEG_TYPE_NAMES[
                expectedLegType
                ];

        /**
         * สร้างชื่อ Field สำหรับ Remark
         *
         * ตัวอย่าง:
         * From Currency (CCY)/To Currency (CCY)
         */
        const currencyPairField =
            `${FXM_TEST_DATA_FIELDS.fromCurrency}/` +
            FXM_TEST_DATA_FIELDS
                .toCurrency;

        /**
         * สร้างค่า Currency Pair
         *
         * ตัวอย่าง:
         * THB/USD
         */
        const currencyPairValue =
            `${testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .fromCurrency,
            )}/` +
            testDataRecord.get(
                FXM_TEST_DATA_FIELDS
                    .toCurrency,
            );

        const actualLegType =
            reportRecord.get(
                FXM_REPORT_FIELDS
                    .legType,
            );

        const actualLegTypeName =
            reportRecord.get(
                FXM_REPORT_FIELDS
                    .legTypeName,
            );

        /**
         * ไม่สามารถหา Expected Leg ได้
         *
         * เช่น:
         * - CROSS_CURRENCY
         * - NOT_FX
         * - UNKNOWN
         *
         * จึงให้ Review ทั้งสอง Field
         */
        if (
            expectedLegType ===
            undefined ||
            expectedLegTypeName ===
            undefined
        ) {
            addDifference(
                "REVIEW",

                FXM_REPORT_FIELDS
                    .legType,

                currencyPairField,

                currencyPairValue,

                actualLegType,
            );

            addDifference(
                "REVIEW",

                FXM_REPORT_FIELDS
                    .legTypeName,

                currencyPairField,

                currencyPairValue,

                actualLegTypeName,
            );

            return;
        }

        /**
         * ตรวจ Leg Type
         */
        if (
            normalizeFXMValue(
                actualLegType,
            ) !==
            normalizeFXMValue(
                expectedLegType,
            )
        ) {
            addDifference(
                "FAIL",

                FXM_REPORT_FIELDS
                    .legType,

                currencyPairField,

                expectedLegType,

                actualLegType,
            );
        }

        /**
         * ตรวจ Leg Type Name
         */
        if (
            normalizeFXMValue(
                actualLegTypeName,
            ) !==
            normalizeFXMValue(
                expectedLegTypeName,
            )
        ) {
            addDifference(
                "FAIL",

                FXM_REPORT_FIELDS
                    .legTypeName,

                currencyPairField,

                expectedLegTypeName,

                actualLegTypeName,
            );
        }
    }

    /**
     * สร้าง ResultRow จากผล Matching
     *
     * ฟังก์ชันนี้เป็นจุดรวมของ:
     * - Presence Rule
     * - Exact/Fallback Matching
     * - Core Field Comparison
     *
     * ผลลัพธ์ที่คืนสามารถส่งให้
     * ReconcileResultSheetWriter ได้โดยตรง
     */
    resolveResultRow(
        testDataRecord:
            ReconcileRecord,

        matchResolution:
            FXMMatchResolution,
    ): ResultRow {
        const {
            decision,
            matchedRecord,
        } =
            matchResolution;

        /**
         * รายการไม่ควรพบใน DF_FXM เช่น:
         * - From Currency = To Currency
         * - Amount < 1,000,000 USD
         */
        if (
            decision.expectation ===
            "MUST_NOT_EXIST"
        ) {
            return this.resolveMustNotExistResult(
                testDataRecord,
                matchResolution,
            );
        }

        /**
         * ====================================================
         * CANNOT_DECIDE
         * ====================================================
         *
         * ข้อมูล Test Data ไม่เพียงพอ
         * สำหรับตัดสิน Presence Rule
         */
        if (
            decision.expectation ===
            "CANNOT_DECIDE"
        ) {
            const cannotDecideRemark = [
                "ไม่สามารถตัดสิน Presence Rule ของ DF_FXM ได้",
                ...decision.reasons,
                matchResolution.remark,
            ]
                .map(
                    (
                        remark,
                    ) =>
                        remark.trim(),
                )
                .filter(
                    Boolean,
                )
                .join(
                    "\n",
                );

            /**
             * ถ้าพบ Report Record
             * ให้ตรวจ Field เท่าที่ตรวจได้
             *
             * แต่สถานะรวมยังเป็น FAIL
             * เพราะ Presence Rule ตัดสินไม่ได้
             */
            if (
                matchedRecord
            ) {
                return this.createMatchedResultRow(
                    testDataRecord,
                    matchResolution,
                    {
                        initialRemarks: [
                            cannotDecideRemark,
                        ],
                    },
                    true,
                );
            }

            /**
             * ไม่พบ Report Record
             * และตัดสิน Presence Rule ไม่ได้
             */
            return this.createUnmatchedFailureResult(
                testDataRecord,
                matchResolution,
                cannotDecideRemark,
            );
        }

        /**
         * ====================================================
         * MUST_EXIST
         * ====================================================
         *
         * รายการต้องพบใน DF_FXM
         */
        if (
            !matchedRecord
        ) {
            return this.createUnmatchedFailureResult(
                testDataRecord,
                matchResolution,
                [
                    "รายการเข้าเงื่อนไขที่ต้องพบใน DF_FXM",
                    matchResolution.remark,
                ]
                    .filter(
                        Boolean,
                    )
                    .join(
                        "\n",
                    ),
            );
        }

        /**
         * Exact Match:
         *
         * ตรวจ Core Fields ตามปกติ
         */
        if (
            matchResolution
                .matchMode ===
            "EXACT"
        ) {
            return this.createMatchedResultRow(
                testDataRecord,
                matchResolution,
            );
        }

        /**
         * Fallback Match:
         *
         * แม้ Date + Amount จะพบ Report Row
         * แต่ Primary Matching Key ยังไม่ตรง
         *
         * จึงให้ Arrangement Number เป็น FAIL
         */
        return this.createMatchedResultRow(
            testDataRecord,
            matchResolution,
            {
                initialRemarks: [
                    matchResolution.remark,
                ],

                initialFailedHeaders: [
                    FXM_REPORT_FIELDS
                        .arrangementNumber,
                ],
            },
        );
    }

    /**
     * สร้างผลสำหรับรายการที่ต้องไม่พบใน DF_FXM
     */
    private resolveMustNotExistResult(
        testDataRecord:
            ReconcileRecord,

        matchResolution:
            FXMMatchResolution,
    ): ResultRow {
        const {
            decision,
            matchedRecord,
        } =
            matchResolution;

        const presenceReason =
            this.getMustNotExistReason(
                testDataRecord,
                decision,
            );

        /**
         * ไม่พบรายการและระบบทำ Fallback แล้ว
         *
         * ตรงตาม Expected Absence
         * จึงเป็น PASS
         */
        if (
            !matchedRecord &&
            matchResolution
                .matchMode ===
            "NOT_FOUND"
        ) {
            const result: ResultRow = {
                testCaseNo:
                    matchResolution
                        .testCaseNo,

                status:
                    "PASS",

                remark: [
                    presenceReason,
                    "ไม่ควรพบและไม่พบรายการใน DF_FXM",
                    matchResolution.remark,
                ]
                    .map(
                        (
                            remark,
                        ) =>
                            remark.trim(),
                    )
                    .filter(
                        Boolean,
                    )
                    .join(
                        "\n",
                    ),

                matchedRowNumber:
                    undefined,

                failedKeyFieldHeaders:
                    [],

                reviewFieldHeaders:
                    [],

                /**
                 * บอก Result Writer ว่านี่คือ
                 * PASS แบบ Expected Absence
                 */
                isExpectedAbsence:
                    true,
            };

            this.appendDataQualityRemarks(
                testDataRecord,
                result,
            );

            return result;
        }

        /**
         * ไม่พบ Matched Record
         * แต่เป็นกรณีที่ยืนยัน Expected Absence ไม่ได้ เช่น:
         *
         * - Ambiguous
         * - Fallback Unavailable
         * - Report Row ถูกใช้แล้ว
         *
         * จึงให้ FAIL
         */
        if (
            !matchedRecord
        ) {
            return this.createUnmatchedFailureResult(
                testDataRecord,
                matchResolution,
                [
                    presenceReason,
                    "ไม่สามารถยืนยัน Expected Absence ได้",
                    matchResolution.remark,
                ]
                    .map(
                        (
                            remark,
                        ) =>
                            remark.trim(),
                    )
                    .filter(
                        Boolean,
                    )
                    .join(
                        "\n",
                    ),
            );
        }

        /**
         * พบรายการที่ไม่ควรมีใน DF_FXM
         *
         * ให้ FAIL และตรวจ Core Fields ต่อ
         * เพื่อแสดงรายละเอียดใน Remark
         */
        return this.createMatchedResultRow(
            testDataRecord,
            matchResolution,
            {
                initialRemarks: [
                    presenceReason,
                    "พบรายการใน DF_FXM ทั้งที่ไม่ควรพบ",
                    matchResolution.remark,
                ],

                initialFailedHeaders: [
                    FXM_REPORT_FIELDS
                        .arrangementNumber,

                    FXM_REPORT_FIELDS
                        .usdEquivalentAmount,
                ],
            },
            true,
        );
    }

    /**
     * สร้าง ResultRow เมื่อพบ Report Record
     *
     * @param forceFail
     * true = บังคับสถานะ FAIL
     *
     * ใช้กับกรณี:
     * - Expected Absence แต่พบรายการ
     * - Presence Rule ตัดสินไม่ได้
     */
    private createMatchedResultRow(
        testDataRecord:
            ReconcileRecord,

        matchResolution:
            FXMMatchResolution,

        options:
            FXMCompareOptions =
            {},

        forceFail =
            false,
    ): ResultRow {
        const matchedRecord =
            matchResolution
                .matchedRecord;

        /**
         * ป้องกันการเรียก Method ผิด Flow
         */
        if (
            !matchedRecord
        ) {
            return this.createUnmatchedFailureResult(
                testDataRecord,
                matchResolution,
                "ไม่พบ Report Record สำหรับเปรียบเทียบข้อมูล",
            );
        }

        /**
         * ตรวจ Core Fields
         */
        const comparison =
            this.compareMatchedFields(
                testDataRecord,
                matchedRecord,
                matchResolution
                    .decision,
                options,
            );

        /**
         * มี failedHeaders อย่างน้อยหนึ่งรายการ
         * หรือ forceFail = true
         * ให้สถานะรวมเป็น FAIL
         */
        const status =
            forceFail ||
                comparison
                    .failedHeaders
                    .length >
                0
                ? "FAIL"
                : "PASS";


        const result: ResultRow = {
            testCaseNo:
                matchResolution
                    .testCaseNo,

            status,

            /**
             * กฎการแสดง Remark:
             *
             * - PASS ปกติ:
             *   ไม่แสดง Remark
             *
             * - PASS เพราะ Txn Date ตรงกับวันที่ใน Ref:
             *   แสดง Remark จาก Date Rule
             *
             * - FAIL:
             *   แสดงสาเหตุที่ตรวจไม่ผ่าน
             *
             * PASS แบบไม่ต้องรายงาน
             * จะสร้าง Remark แยกใน
             * resolveMustNotExistResult()
             */
            remark:
                comparison
                    .remarks
                    .map(
                        (
                            remark,
                        ) =>
                            remark.trim(),
                    )
                    .filter(
                        Boolean,
                    )
                    .join(
                        "\n",
                    ),

            matchedRowNumber:
                matchedRecord
                    .rowNumber,

            failedKeyFieldHeaders:
                comparison
                    .failedHeaders,

            reviewFieldHeaders:
                comparison
                    .reviewHeaders,

            isExpectedAbsence:
                false,
        };

        this.appendDataQualityRemarks(
            testDataRecord,
            result,
        );

        return result;
    }

    /**
     * สร้าง ResultRow เมื่อไม่มี Report Record
     *
     * รายการที่ควรพบแต่ไม่พบ
     * หรือไม่สามารถระบุ Candidate ได้
     * จะเป็น FAIL
     */
    private createUnmatchedFailureResult(
        testDataRecord:
            ReconcileRecord,

        matchResolution:
            FXMMatchResolution,

        remark:
            string,
    ): ResultRow {
        const result: ResultRow = {
            testCaseNo:
                matchResolution
                    .testCaseNo,

            status:
                "FAIL",

            remark,

            matchedRowNumber:
                undefined,

            failedKeyFieldHeaders: [
                FXM_REPORT_FIELDS
                    .arrangementNumber,
            ],

            reviewFieldHeaders:
                [],

            isExpectedAbsence:
                false,
        };

        this.appendDataQualityRemarks(
            testDataRecord,
            result,
        );

        return result;
    }

    /**
     * สร้างเหตุผลว่าเพราะเหตุใด
     * รายการจึงต้องไม่พบใน DF_FXM
     */
    private getMustNotExistReason(
        testDataRecord:
            ReconcileRecord,

        decision:
            FXMReconcileDecision,
    ): string {

        /**
         * Return/Reversal ใช้เหตุผลที่ได้จาก Original Transaction
         *
         * ไม่ใช้ From/To Currency หรือ Amount ของแถว Return
         * เพราะข้อมูลดังกล่าวอาจว่างหรือแตกต่างจาก Original ได้
         */
        const returnReason =
            decision.reasons.find(
                (
                    reason,
                ) =>
                    reason.startsWith(
                        "Return/Reversal:",
                    ),
            );

        if (
            returnReason !==
            undefined
        ) {
            return returnReason;
        }

        /**
         * From Currency และ To Currency เหมือนกัน
         * จึงไม่ใช่รายการ FX
         */
        if (
            decision.direction ===
            "NOT_FX"
        ) {
            const fromCurrency =
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .fromCurrency,
                );

            const toCurrency =
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .toCurrency,
                );

            return (
                `From Currency = "${fromCurrency}" และ ` +
                `To Currency = "${toCurrency}" ` +
                "เป็นสกุลเดียวกัน จึงไม่ควรพบใน DF_FXM"
            );
        }

        /**
        * Amount ต่ำกว่า 1,000,000 USD
        * ต้องพิจารณารายงานใน DF_FXU
        */
        const amount =
            parseFXMAmount(
                testDataRecord.get(
                    FXM_TEST_DATA_FIELDS
                        .settledAmount,
                ),
            );

        const amountDisplay =
            amount ===
                null
                ? "อ่านค่าไม่ได้"
                : amount.toLocaleString(
                    "en-US",
                    {
                        maximumFractionDigits:
                            2,
                    },
                );

        return (
            `Settled Amount = ${amountDisplay} USD ` +
            `< ${FXM_USD_THRESHOLD.toLocaleString(
                "en-US",
            )} USD ` +
            "จึงต้องตรวจใน DF_FXU และไม่ควรพบใน DF_FXM"
        );
    }

    /**
     * เพิ่ม Data Quality Remark
     *
     * Remark กลุ่มนี้ไม่เปลี่ยน Business PASS/FAIL
     */
    private appendDataQualityRemarks(
        testDataRecord:
            ReconcileRecord,

        result:
            ResultRow,
    ): void {
        /**
         * Test No. ใช้แสดงผลเท่านั้น
         * จึงไม่ใช้ตัดสิน Business PASS/FAIL
         */
        if (
            testDataRecord
                .get(
                    FXM_TEST_DATA_FIELDS
                        .testNo,
                )
                .trim() ===
            ""
        ) {
            result.remark =
                this.appendRemark(
                    result.remark,

                    `Test Data ไม่มี ` +
                    FXM_TEST_DATA_FIELDS
                        .testNo,
                );
        }

        /**
         * Transaction ID ว่าง
         *
         * ถึงแม้ Fallback จะพบ Candidate
         * ต้องบันทึกไว้ใน Remark เสมอ
         */
        if (
            testDataRecord
                .get(
                    FXM_TEST_DATA_FIELDS
                        .transactionId,
                )
                .trim() ===
            ""
        ) {
            result.remark =
                this.appendRemark(
                    result.remark,

                    `Test Data ไม่มี ` +
                    FXM_TEST_DATA_FIELDS
                        .transactionId,
                );
        }
    }

    /**
     * ต่อ Remark ใหม่เข้ากับ Remark เดิม
     * และป้องกันข้อความซ้ำ
     */
    private appendRemark(
        currentRemark:
            string,

        nextRemark:
            string,
    ): string {
        const normalizedNext =
            nextRemark.trim();

        if (
            normalizedNext ===
            "" ||
            currentRemark.includes(
                normalizedNext,
            )
        ) {
            return currentRemark;
        }

        return [
            currentRemark,
            normalizedNext,
        ]
            .map(
                (
                    remark,
                ) =>
                    remark.trim(),
            )
            .filter(
                Boolean,
            )
            .join(
                "\n",
            );
    }
}