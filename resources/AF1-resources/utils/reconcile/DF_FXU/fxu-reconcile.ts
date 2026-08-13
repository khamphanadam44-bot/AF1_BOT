/**
 * fxu-reconcile.ts
 * ------------------------------------------------------------------
 * ตัวควบคุมหลักของ Script 3 สำหรับ DF_FXU
 *
 * Flow ที่ไฟล์นี้จะรับผิดชอบ:
 * 1. อ่าน Checked Report และ Test Data
 * 2. ตรวจ Header ที่จำเป็น
 * 3. ตัดสินว่ารายการต้องมีหรือไม่ต้องมีใน DF_FXU
 * 4. จับคู่ Transaction ID กับ Arrangement Number
 * 5. ใช้ Txn Date และ Settled Amount เป็น Matching Support
 * 6. ตรวจ Core Fields
 * 7. เขียนผลลง DF_FXU_Reconcile
 *
 * Entry Point:
 * - reconcileFxuReport()
 * - ถูกเรียกจาก tests/script3-compare-report.spec.ts
 *
 * การตรวจสอบ:
 * - ใช้ npm run test:script3 -- report=DF_FXU เพื่อรัน Reconcile
 * ------------------------------------------------------------------
 */

import {
    FXU_ARRANGEMENT_TYPE,
    FXU_LEG_TYPES,
    FXU_LEG_TYPE_NAMES,
    FXU_REPORT_CODE,
    FXU_REPORT_FIELDS,
    FXU_REPORT_HEADER_ROW,
    FXU_TEST_DATA_FIELDS,
    FXU_TEST_DATA_HEADER_ROW,
    FXU_USD_THRESHOLD,
} from "./fxu-config";

import {
    FxuDirection,
    FxuRuleEvaluator,
    normalizeFxuValue,
    parseFxuAmount,
} from "./fxu-rules";

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
    formatCompareRemark,
    formatFixedValueRemark,
} from "../shared/remark";

import {
    ReconcileResultSheetWriter,
    ResultRow,
} from "../shared/result-writer";

import {
    PreparedReconcileWorkbook,
    ReconcileWorkbookPreparer,
} from "../shared/workbook-preparer";

/**
 * สถานะที่ใช้ตัดสินว่า Transaction
 * ต้องมีหรือไม่ต้องมีใน Report DF_FXU
 *
 * MUST_EXIST:
 * รายการต้องพบใน DF_FXU
 *
 * MUST_NOT_EXIST:
 * รายการต้องไม่พบใน DF_FXU
 *
 * CANNOT_DECIDE:
 * ข้อมูล Test Data ไม่เพียงพอ
 * จึงยังไม่สามารถตัดสิน Presence Rule ได้
 */
export type FxuPresenceExpectation =
    | "MUST_EXIST"
    | "MUST_NOT_EXIST"
    | "CANNOT_DECIDE";

/**
 * ผลการตัดสิน Rule ของ Test Data หนึ่งแถว
 */
export interface FxuReconcileDecision {
    /**
     * ทิศทางของธุรกรรม FX
     *
     * ตัวอย่าง:
     * - BUY_FCY
     * - SELL_FCY
     * - CROSS_CURRENCY
     * - NOT_FX
     * - UNKNOWN
     */
    direction: FxuDirection;

    /**
     * รายการต้องมีหรือไม่ต้องมีใน DF_FXU
     */
    expectation: FxuPresenceExpectation;

    /**
     * เหตุผลเพิ่มเติมจากการตัดสิน Rule
     *
     * ใช้แสดงใน Remark
     */
    reasons: string[];

    /**
     * true:
     * มีข้อมูลบางส่วนที่ต้องให้ผู้ใช้งานตรวจสอบเพิ่มเติม
     *
     * ปัจจุบันใช้กับ Cross Currency
     * เพราะ Requirement ของ Settlement/Intermediary
     * ยังให้รายละเอียดไม่ครบ
     */
    requiresReview: boolean;
}

/**
 * วิธีที่ระบบใช้จับคู่ Test Data กับ DF_FXU Report
 *
 * EXACT:
 * Transaction ID ตรงกับ Arrangement Number
 *
 * FALLBACK:
 * ใช้ Txn Date และ Settled Amount ช่วยค้นหา
 *
 * NOT_FOUND:
 * ไม่พบทั้ง Exact และ Fallback
 *
 * AMBIGUOUS:
 * พบ Candidate มากกว่าหนึ่งแถว
 *
 * ROW_ALREADY_USED:
 * Report Row ถูกใช้กับ Test Case ก่อนหน้าแล้ว
 *
 * FALLBACK_UNAVAILABLE:
 * ข้อมูลไม่เพียงพอสำหรับทำ Fallback
 */
export type FxuMatchMode =
    | "EXACT"
    | "FALLBACK"
    | "NOT_FOUND"
    | "AMBIGUOUS"
    | "ROW_ALREADY_USED"
    | "FALLBACK_UNAVAILABLE";

/**
 * ผลการจับคู่ Test Data หนึ่งแถว
 */
export interface FxuMatchResolution {
    /**
     * ค่าที่ใช้แสดงใน Column Test Script No.
     *
     * ลำดับการเลือก:
     * 1. Test No.
     * 2. Transaction ID
     * 3. TEST DATA ROW ตามด้วยเลขแถว
     */
    testCaseNo: string;

    /**
     * Transaction ID ที่ผ่านการ Normalize แล้ว
     */
    transactionId: string;

    /**
     * ผล Presence Rule ของ Test Data แถวนี้
     */
    decision:
    FxuReconcileDecision;

    /**
     * วิธีที่ใช้จับคู่ข้อมูล
     */
    matchMode:
    FxuMatchMode;

    /**
     * Report Record ที่จับคู่ได้
     *
     * ไม่มีค่าในกรณี:
     * - NOT_FOUND
     * - AMBIGUOUS
     * - ROW_ALREADY_USED
     * - FALLBACK_UNAVAILABLE
     */
    matchedRecord?:
    ReconcileRecord;

    /**
     * จำนวน Candidate ที่พบ
     */
    candidateCount:
    number;

    /**
     * รายละเอียดของผล Matching
     */
    remark:
    string;
}

/**
 * ผลการค้นหา Fallback ภายในระบบ
 */
interface FxuFallbackResult {
    matchedRecord?:
    ReconcileRecord;

    candidateCount?:
    number;

    remark:
    string;
}

/**
 * ตัวเลือกเริ่มต้นสำหรับการเปรียบเทียบ Field
 *
 * ใช้ในกรณีพิเศษ เช่น:
 * - Fallback Matching
 * - Expected Absence แต่พบรายการ
 * - ไม่สามารถตัดสิน Presence Rule ได้
 */
export interface FxuCompareOptions {
    /**
     * Remark ที่ต้องใส่ก่อนเริ่มตรวจ Field
     */
    initialRemarks?:
    readonly string[];

    /**
     * Header ที่ต้องเริ่มต้นเป็น FAIL
     */
    initialFailedHeaders?:
    readonly string[];

    /**
     * Header ที่ต้องเริ่มต้นเป็น Review
     */
    initialReviewHeaders?:
    readonly string[];
}

/**
 * ผลการเปรียบเทียบ Field
 * ของ Test Data กับ Report หนึ่งคู่
 */
export interface FxuFieldComparisonResult {
    /**
     * Header ที่ไม่ผ่าน Business Rule
     *
     * มีผลต่อสถานะระดับ Test Case
     */
    failedHeaders:
    string[];

    /**
     * Header ที่ต้องตรวจสอบเพิ่มเติม
     *
     * ไม่มีผลต่อสถานะ PASS/FAIL
     */
    reviewHeaders:
    string[];

    /**
     * รายละเอียดค่าที่ไม่ตรงกัน
     */
    remarks:
    string[];
}

/**
 * ข้อมูลที่อ่านและตรวจ Header เรียบร้อยแล้ว
 *
 * prepared:
 * Workbook ที่สร้างจาก Checked Report
 *
 * reportRecords:
 * ข้อมูลทุกแถวจาก DF_FXU Report
 *
 * testDataRecords:
 * ข้อมูลทุกแถวจาก Test Data
 *
 * reportIndex:
 * Index สำหรับค้นหา Arrangement Number
 *
 * reservedExactRowNumbers:
 * Report Row ที่ถูกจองไว้สำหรับ Exact Matching
 */
export interface FxuLoadedData {
    prepared:
    PreparedReconcileWorkbook;

    reportRecords:
    ReconcileRecord[];

    testDataRecords:
    ReconcileRecord[];

    reportIndex:
    Map<
        string,
        ReconcileRecord[]
    >;

    reservedExactRowNumbers:
    Set<number>;
}

/**
 * ค่าคลาดเคลื่อนของ Amount
 *
 * ใช้ป้องกันปัญหาทศนิยมจาก Excel เช่น:
 *
 * Test Data:
 * 999999.999999
 *
 * Report:
 * 1000000.00
 *
 * ถ้าผลต่างไม่เกิน 0.01
 * จะถือว่า Amount เท่ากัน
 */
export const FXU_AMOUNT_TOLERANCE =
    0.01;

/**
 * รายการ Header ของ Test Data
 * ที่ Script 3 สำหรับ DF_FXU ต้องใช้
 *
 * ใช้ Set เพื่อตัด Header ซ้ำ
 * หาก Config มีชื่อ Header เดียวกันมากกว่าหนึ่งตำแหน่ง
 */
export const FXU_REQUIRED_TEST_DATA_HEADERS = [
    ...new Set(
        Object.values(
            FXU_TEST_DATA_FIELDS,
        ),
    ),
];

/**
 * แปลงวันที่ให้อยู่ในรูปแบบ:
 *
 * yyyy-mm-dd
 *
 * รูปแบบที่รองรับ:
 * - yyyy-mm-dd
 * - yyyy/mm/dd
 * - dd/mm/yyyy
 * - dd-mm-yyyy
 * - Excel Serial Date
 *
 * ถ้าค่าว่างหรือรูปแบบวันที่ไม่ถูกต้อง
 * จะคืนค่า null
 */
export const parseFxuDateKey = (
    value: unknown,
): string | null => {
    const text =
        String(
            value ?? "",
        ).trim();

    if (
        text ===
        ""
    ) {
        return null;
    }

    /**
     * รองรับ Excel Serial Date
     *
     * ตัวอย่าง:
     * 45986
     */
    if (
        /^\d+(?:\.\d+)?$/.test(
            text,
        )
    ) {
        const serial =
            Number(
                text,
            );

        if (
            !Number.isFinite(
                serial,
            ) ||
            serial <=
            0
        ) {
            return null;
        }

        /**
         * Excel ใช้วันที่ 1899-12-30
         * เป็นจุดเริ่มต้นสำหรับ Serial Date
         */
        const excelEpoch =
            Date.UTC(
                1899,
                11,
                30,
            );

        const date =
            new Date(
                excelEpoch +
                Math.floor(
                    serial,
                ) *
                86_400_000,
            );

        return [
            date.getUTCFullYear(),

            String(
                date.getUTCMonth() +
                1,
            ).padStart(
                2,
                "0",
            ),

            String(
                date.getUTCDate(),
            ).padStart(
                2,
                "0",
            ),
        ].join(
            "-",
        );
    }

    /**
     * รูปแบบวันขึ้นก่อน:
     *
     * dd/mm/yyyy
     * dd-mm-yyyy
     */
    const dayFirst =
        text.match(
            /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s.*)?$/,
        );

    /**
     * รูปแบบปีขึ้นก่อน:
     *
     * yyyy/mm/dd
     * yyyy-mm-dd
     */
    const yearFirst =
        text.match(
            /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[T\s].*)?$/,
        );

    let year: number;
    let month: number;
    let day: number;

    if (
        dayFirst
    ) {
        day =
            Number(
                dayFirst[1],
            );

        month =
            Number(
                dayFirst[2],
            );

        year =
            Number(
                dayFirst[3],
            );
    } else if (
        yearFirst
    ) {
        year =
            Number(
                yearFirst[1],
            );

        month =
            Number(
                yearFirst[2],
            );

        day =
            Number(
                yearFirst[3],
            );
    } else {
        return null;
    }

    /**
     * สร้าง Date เพื่อยืนยันว่าเป็นวันที่จริง
     *
     * ตัวอย่างวันที่ที่ไม่ผ่าน:
     * - 31/02/2026
     * - 2026-13-01
     */
    const date =
        new Date(
            Date.UTC(
                year,
                month - 1,
                day,
            ),
        );

    const isValid =
        date.getUTCFullYear() ===
        year &&
        date.getUTCMonth() ===
        month - 1 &&
        date.getUTCDate() ===
        day;

    if (
        !isValid
    ) {
        return null;
    }

    return [
        year,

        String(
            month,
        ).padStart(
            2,
            "0",
        ),

        String(
            day,
        ).padStart(
            2,
            "0",
        ),
    ].join(
        "-",
    );
};

/**
 * อ่านวันที่ทำธุรกรรมจาก Transaction ID/Reconcile ID
 *
 * รูปแบบ Ref ที่รองรับ:
 * - ตัวอักษร 3 ตัว
 * - ตัวเลข 3 ตัว
 * - วันที่ YYMMDD
 * - เลขลำดับรายการ
 *
 * ตัวอย่าง:
 * KMA3012511260000000028158
 *       251126 = 26/11/2025
 *
 * ถ้า Ref ไม่มีวันที่ หรือวันที่ไม่ถูกต้อง
 * จะคืนค่า null
 */
export const parseFxuReferenceDateKey = (
    value: unknown,
): string | null => {
    const text =
        String(
            value ?? "",
        ).trim();

    const matchedDate =
        text.match(
            /^[A-Za-z]{3}\d{3}(\d{2})(\d{2})(\d{2})/,
        );

    if (
        !matchedDate
    ) {
        return null;
    }

    /**
     * วันที่ใน Ref ใช้รูปแบบ YYMMDD
     *
     * matchedDate[1] = ปี
     * matchedDate[2] = เดือน
     * matchedDate[3] = วัน
     */
    const year =
        2000 +
        Number(
            matchedDate[1],
        );

    const month =
        Number(
            matchedDate[2],
        );

    const day =
        Number(
            matchedDate[3],
        );

    /**
     * ใช้ Date Parser เดิมตรวจสอบซ้ำว่า
     * วันที่จาก Ref เป็นวันที่จริง
     *
     * ตัวอย่าง Ref ที่มีวันที่ 250231
     * จะไม่ผ่าน เพราะไม่มีวันที่ 31/02/2025
     */
    return parseFxuDateKey(
        `${day}/${month}/${year}`,
    );
};

/**
 * ประเภทของรายการที่ต้องโยงกลับไปหา
 * Original Transaction
 */
export type FxuReturnTransactionType =
    | "RETURN"
    | "REVERSAL";

/**
 * ข้อมูลที่อ่านได้จาก Test No.
 * ของ Return/Reversal
 */
export interface FxuReturnTestInfo {
    /**
     * Test No. ของ Original Transaction
     *
     * ตัวอย่าง:
     * BOTDMS_010-return
     * → BOTDMS_010
     */
    originalTestNo:
    string;

    /**
     * ประเภทของรายการ
     */
    transactionType:
    FxuReturnTransactionType;
}

/**
 * ตรวจว่า Test No. เป็น Return/Reversal หรือไม่
 * และคืน Test No. ของ Original Transaction
 *
 * รูปแบบที่รองรับ:
 * - 2511_004-Return
 * - BOTDMS_010-return
 * - TC001-Reversal
 * - TC001 - reversal
 *
 * การตรวจไม่สนตัวพิมพ์เล็ก/ใหญ่
 *
 * ถ้าไม่ใช่ Return/Reversal
 * หรือไม่มี Original Test No.
 * จะคืนค่า null
 */
export const parseFxuReturnTestInfo = (
    value: unknown,
): FxuReturnTestInfo | null => {
    const text =
        String(
            value ?? "",
        ).trim();

    /**
     * แยกข้อความออกเป็น:
     *
     * matched[1]:
     * Original Test No.
     *
     * matched[2]:
     * return หรือ reversal
     */
    const matched =
        text.match(
            /^(.*?)\s*-\s*(return|reversal)\s*$/i,
        );

    if (
        !matched
    ) {
        return null;
    }

    const originalTestNo =
        matched[1].trim();

    /**
     * ป้องกันกรณีมีเพียง:
     * -return
     * หรือ
     * -reversal
     *
     * แต่ไม่มี Original Test No.
     */
    if (
        originalTestNo ===
        ""
    ) {
        return null;
    }

    return {
        originalTestNo,

        transactionType:
            matched[2]
                .toUpperCase() as
            FxuReturnTransactionType,
    };
};

/**
 * Service หลักสำหรับ Reconcile DF_FXU
 *
 * ขั้นนี้เพิ่มเฉพาะ Presence Decision ก่อน
 * ส่วนการอ่านไฟล์และเขียนผลจะเพิ่มในขั้นถัดไป
 */
export class FxuReconcileService {
    /**
     * ใช้ Rule จาก fxu-rules.ts
     * เพื่อตรวจ Field และหาทิศทางธุรกรรม
     */
    private readonly ruleEvaluator =
        new FxuRuleEvaluator();

    /**
     * Entry Point หลักของ DF_FXU Reconcile
     *
     * ขั้นตอน:
     * 1. อ่าน Checked Report และ Test Data
     * 2. ตรวจ Header
     * 3. วนประมวลผล Test Data ทุกแถว
     * 4. ทำ Exact/Fallback Matching
     * 5. ตัดสิน PASS/FAIL
     * 6. เขียนผลลง DF_FXU_Reconcile
     * 7. ลบ Source Worksheet
     * 8. บันทึกไฟล์ผลลัพธ์
     */
    async reconcile(
        testDataFilePath:
            string,
    ): Promise<string> {
        /**
         * อ่านข้อมูลและตรวจ Header
         */
        const loadedData =
            await this.loadAndValidate(
                testDataFilePath,
            );

        const {
            prepared,
            reportRecords,
            testDataRecords,
            reportIndex,
            reservedExactRowNumbers,
        } =
            loadedData;

        /**
   * สร้าง Index ของ Test Data จาก Test No.
   *
   * ใช้สำหรับเชื่อม Return/Reversal
   * กลับไปหา Original Transaction
   */
        const testDataIndex =
            this.indexTestDataRecordsByTestNo(
                testDataRecords,
            );

        const resultWriter =
            new ReconcileResultSheetWriter();

        /**
         * เก็บ Report Row ที่ถูกใช้ไปแล้ว
         *
         * ป้องกัน Test Case หลายรายการ
         * จับคู่กับ Report Row เดียวกัน
         */
        const usedReportRowNumbers =
            new Set<number>();

        /**
         * Map สำหรับผูก Report Row
         * กับ ResultRow ที่สร้างจาก Test Data
         *
         * Key:
         * เลขแถวจริงใน Report
         *
         * Value:
         * ผล PASS/FAIL ของ Test Case
         */
        const annotationByRowNumber =
            new Map<
                number,
                ResultRow
            >();

        /**
         * ผลลัพธ์ที่ไม่มี Report Row
         *
         * ตัวอย่าง:
         * - Expected Absence
         * - Matching Not Found
         * - Ambiguous Match
         * - Fallback Unavailable
         */
        const unmatchedRows:
            ResultRow[] =
            [];

        /**
         * ผลลัพธ์ทั้งหมด
         *
         * ใช้สรุปจำนวน PASS/FAIL
         */
        const results:
            ResultRow[] =
            [];

        console.log(
            `\n===== RECONCILE - ${FXU_REPORT_CODE} =====`,
        );

        /**
         * ประมวลผล Test Data ทีละหนึ่งแถว
         */
        for (
            const testDataRecord of
            testDataRecords
        ) {
            /**
             * Exact/Fallback Matching
             */
            const matchResolution =
                this.resolveMatch(
                    testDataRecord,
                    testDataIndex,
                    reportRecords,
                    reportIndex,
                    reservedExactRowNumbers,
                    usedReportRowNumbers,
                );


            /**
             * นำ Matching + Presence + Field Comparison
             * มาสร้าง ResultRow
             */
            const result =
                this.resolveResultRow(
                    testDataRecord,
                    matchResolution,
                );

            results.push(
                result,
            );

            /**
             * แสดงผล Matching ใน Terminal
             *
             * ช่วยตรวจสอบว่าแต่ละ Test Case
             * ใช้ Exact, Fallback หรือไม่พบข้อมูล
             */
            console.log(
                [
                    `Test Case: ${result.testCaseNo}`,
                    `Match: ${matchResolution.matchMode}`,
                    `Result: ${result.status}`,
                ].join(
                    " | ",
                ),
            );

            /**
             * ไม่มี Report Row ที่จับคู่ได้
             *
             * ส่งไปเขียนในกลุ่ม Unmatched Rows
             */
            if (
                result.matchedRowNumber ===
                undefined
            ) {
                unmatchedRows.push(
                    result,
                );

                continue;
            }

            /**
             * ผูก ResultRow กับ Report Row
             *
             * Result Writer จะใช้เลขแถวนี้
             * Copy ข้อมูลจาก Source Worksheet
             * และใส่สีตามผล Validation
             */
            annotationByRowNumber.set(
                result.matchedRowNumber,
                result,
            );

            /**
             * บันทึกว่า Report Row นี้ถูกใช้แล้ว
             */
            usedReportRowNumbers.add(
                result.matchedRowNumber,
            );
        }

        /**
         * เขียน Header ของ Result Worksheet
         *
         * Column เริ่มต้น:
         * - Test Script No.
         * - Test Result
         * - Remark
         *
         * หลังจากนั้นเป็น Header จาก DF_FXU Report
         */
        resultWriter.writeHeaderRow(
            prepared.resultSheet,
            prepared.reportHeaders,
        );

        /**
         * เขียนข้อมูลตามลำดับมาตรฐาน:
         *
         * 1. PASS ที่มี Report Row
         * 2. Report Row ที่ไม่มี Test Case
         * 3. FAIL ที่มี Report Row
         * 4. PASS แบบ Expected Absence
         * 5. FAIL ที่ไม่มี Report Row
         */
        const nextOutputRowNumber =
            resultWriter
                .writeRowsInRequestedOrder(
                    prepared.resultSheet,
                    prepared.reportWorksheet,
                    prepared.reportHeaders,
                    FXU_REPORT_HEADER_ROW +
                    1,
                    prepared.reportWorksheet
                        .rowCount,
                    annotationByRowNumber,
                    unmatchedRows,
                );

        /**
         * เพิ่ม AutoFilter ตั้งแต่ Header
         * ถึงแถวข้อมูลสุดท้าย
         */
        resultWriter.finalizeAutoFilter(
            prepared.resultSheet,
            prepared.reportHeaders,
            nextOutputRowNumber -
            1,
        );

        /**
         * ลบ Source Worksheet ออกจากไฟล์ผลลัพธ์
         *
         * เหลือเฉพาะ Worksheet:
         * DF_FXU_Reconcile
         */
        prepared.workbook.removeWorksheet(
            prepared.reportWorksheet.id,
        );

        /**
         * บันทึกไฟล์ Excel
         */
        await prepared.workbook.xlsx.writeFile(
            prepared.reconcileFilePath,
        );

        this.logReconcileSummary(
            prepared.reconcileFilePath,
            results,
        );

        return prepared.reconcileFilePath;
    }

    /**
     * แสดงผลสรุปหลัง Reconcile เสร็จ
     */
    private logReconcileSummary(
        outputFilePath:
            string,

        results:
            ResultRow[],
    ): void {
        const passCount =
            results.filter(
                (
                    result,
                ) =>
                    result.status ===
                    "PASS",
            ).length;

        const failCount =
            results.length -
            passCount;

        const expectedAbsenceCount =
            results.filter(
                (
                    result,
                ) =>
                    result.status ===
                    "PASS" &&
                    result
                        .isExpectedAbsence ===
                    true,
            ).length;

        console.log(
            "========================================",
        );

        console.log(
            `${FXU_REPORT_CODE} Reconcile Completed`,
        );

        console.log(
            `Output File : ${outputFilePath}`,
        );

        console.log(
            `Test Case : ${results.length}`,
        );

        console.log(
            `PASS : ${passCount}`,
        );

        console.log(
            `FAIL : ${failCount}`,
        );

        console.log(
            `Expected Absence : ${expectedAbsenceCount}`,
        );

        console.log(
            "========================================",
        );
    }

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
    ): Promise<FxuLoadedData> {
        const workbookPreparer =
            new ReconcileWorkbookPreparer();

        const excelReader =
            new ReconcileExcelReader();

        console.log(
            `\n===== LOAD RECONCILE DATA - ${FXU_REPORT_CODE} =====`,
        );

        /**
         * หา Checked Report ล่าสุด
         * และ Copy ไปสร้างไฟล์ Reconcile Result
         */
        const prepared =
            await workbookPreparer.prepare(
                FXU_REPORT_CODE,
                FXU_REPORT_HEADER_ROW,
            );

        /**
         * อ่านข้อมูลจาก Report Worksheet
         *
         * Header ของ DF_FXU อยู่แถวที่ 1
         */
        const reportData =
            excelReader.parseWorksheet(
                prepared.reportWorksheet,
                FXU_REPORT_HEADER_ROW,
            );

        /**
         * อ่าน Test Data
         *
         * Header ของ Test Data อยู่แถวที่ 5
         */
        const testData =
            await excelReader.readFile(
                testDataFilePath,
                FXU_TEST_DATA_HEADER_ROW,
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
     * ต้องมีหรือไม่ต้องมีใน DF_FXU
     */
    evaluatePresence(
        testDataRecord: ReconcileRecord,
    ): FxuReconcileDecision {
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
            parseFxuAmount(
                testDataRecord.get(
                    FXU_TEST_DATA_FIELDS
                        .settledAmount,
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
                        FXU_TEST_DATA_FIELDS
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
         * จึงต้องไม่พบใน DF_FXU
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
         * Amount ตั้งแต่ 1,000,000 USD ขึ้นไป
         *
         * ต้องตรวจใน DF_FXM
         * จึงต้องไม่พบใน DF_FXU
         */
        if (
            settledAmount >=
            FXU_USD_THRESHOLD
        ) {
            return {
                direction,

                expectation:
                    "MUST_NOT_EXIST",

                reasons: [
                    `Settled Amount ${settledAmount} >= ` +
                    `${FXU_USD_THRESHOLD}`,
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
         * จึงห้ามสรุปว่ารายการต้องพบหรือไม่ต้องพบใน DF_FXU
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
         * และ Amount ต่ำกว่า 1,000,000 USD
         *
         * ต้องพบใน DF_FXU
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
            FxuReturnTestInfo,

        testDataIndex:
            ReadonlyMap<
                string,
                readonly ReconcileRecord[]
            >,
    ): FxuReconcileDecision {
        const originalTestNo =
            normalizeFxuValue(
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
            FXU_TEST_DATA_FIELDS
                .transactionDate,
            FXU_TEST_DATA_FIELDS
                .settledCurrency,
            FXU_TEST_DATA_FIELDS
                .settledAmount,
        ];

        const invalidReturnFields =
            requiredReturnFields.filter(
                (
                    field,
                ) =>
                    normalizeFxuValue(
                        returnRecord.get(
                            field,
                        ),
                    ) === "",
            );

        const returnAmount =
            parseFxuAmount(
                returnRecord.get(
                    FXU_TEST_DATA_FIELDS
                        .settledAmount,
                ),
            );

        if (
            returnAmount ===
            null &&
            !invalidReturnFields.includes(
                FXU_TEST_DATA_FIELDS
                    .settledAmount,
            )
        ) {
            invalidReturnFields.push(
                FXU_TEST_DATA_FIELDS
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
         * ถ้า Original ไม่ต้องรายงานใน DF_FXU
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
                    "ไม่ต้องรายงานใน DF_FXU " +
                    `จึงไม่ต้องรายงานรายการ ${returnInfo.transactionType} ตาม`,
                ],

                requiresReview:
                    false,
            };
        }

        /**
         * Original ต้องรายงานใน DF_FXU
         * Return/Reversal ต้องรายงานเช่นกัน
         * แต่ต้องกลับทิศทาง FX
         *
         * Original BUY_FCY  → Return SELL_FCY
         * Original SELL_FCY → Return BUY_FCY
         */
        const reverseDirection:
            FxuDirection =
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
     * จับคู่ Test Data หนึ่งแถวกับ DF_FXU Report
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
    ): FxuMatchResolution {
        /**
         * อ่านและ Normalize Matching Key
         *
         * ตัวอย่าง:
         * " tx001 "
         * → "TX001"
         */
        const transactionId =
            normalizeFxuValue(
                testDataRecord.get(
                    FXU_TEST_DATA_FIELDS
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
         * หรือไม่ควรมีใน DF_FXU
         */
        const returnInfo =
            parseFxuReturnTestInfo(
                testDataRecord.get(
                    FXU_TEST_DATA_FIELDS
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
                    `พบ ${FXU_REPORT_FIELDS.arrangementNumber} ` +
                    `= "${transactionId}" ซ้ำ ` +
                    `${exactCandidates.length} แถวใน DF_FXU Report`,
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
                    `${FXU_TEST_DATA_FIELDS.transactionId} ` +
                    `ตรงกับ ${FXU_REPORT_FIELDS.arrangementNumber}`,
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
    ): FxuFallbackResult {
        /**
         * อ่านวันที่จาก Test Data
         */
        const expectedDateText =
            testDataRecord.get(
                FXU_TEST_DATA_FIELDS
                    .transactionDate,
            );

        const expectedDate =
            parseFxuDateKey(
                expectedDateText,
            );

        /**
         * อ่าน Amount จาก Test Data
         */
        const expectedAmountText =
            testDataRecord.get(
                FXU_TEST_DATA_FIELDS
                    .settledAmount,
            );

        const expectedAmount =
            parseFxuAmount(
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
                FXU_TEST_DATA_FIELDS
                    .transactionDate,
            );
        }

        if (
            expectedAmount ===
            null
        ) {
            invalidFields.push(
                FXU_TEST_DATA_FIELDS
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
                        parseFxuDateKey(
                            reportRecord.get(
                                FXU_REPORT_FIELDS
                                    .dataSetDate,
                            ),
                        );

                    const actualAmount =
                        parseFxuAmount(
                            reportRecord.get(
                                FXU_REPORT_FIELDS
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
                        FXU_AMOUNT_TOLERANCE
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
                    `${FXU_TEST_DATA_FIELDS.transactionDate} + ` +
                    FXU_TEST_DATA_FIELDS
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
                    `${FXU_TEST_DATA_FIELDS.transactionDate} + ` +
                    FXU_TEST_DATA_FIELDS
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
                `${FXU_TEST_DATA_FIELDS.transactionDate} + ` +
                FXU_TEST_DATA_FIELDS
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
                    FXU_TEST_DATA_FIELDS
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
            FxuReconcileDecision,

        options:
            FxuCompareOptions =
            {},
    ): FxuFieldComparisonResult {
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
                    FXU_REPORT_CODE,
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
                FXU_TEST_DATA_FIELDS
                    .transactionId,
            );

        const actualArrangementNumber =
            reportRecord.get(
                FXU_REPORT_FIELDS
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
            normalizeFxuValue(
                expectedTransactionId,
            ) ===
            "" ||
            normalizeFxuValue(
                expectedTransactionId,
            ) !==
            normalizeFxuValue(
                actualArrangementNumber,
            )
        ) {
            addDifference(
                "FAIL",

                FXU_REPORT_FIELDS
                    .arrangementNumber,

                FXU_TEST_DATA_FIELDS
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
                FXU_TEST_DATA_FIELDS
                    .transactionDate,
            );

        const actualDateText =
            reportRecord.get(
                FXU_REPORT_FIELDS
                    .dataSetDate,
            );

        const expectedDate =
            parseFxuDateKey(
                expectedDateText,
            );

        const actualDate =
            parseFxuDateKey(
                actualDateText,
            );

        /**
         * อ่านวันที่จาก Transaction ID/Reconcile ID
         *
         * ใช้เป็นตัวตรวจสอบสำรองเมื่อ
         * Txn Date ไม่ตรงกับ Data Set Date
         */
        const referenceDate =
            parseFxuReferenceDateKey(
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
                FXU_REPORT_FIELDS
                    .dataSetDate,
            );

            remarks.push(
                `ไม่สามารถอ่านค่า ` +
                `${FXU_TEST_DATA_FIELDS.transactionDate} ` +
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
                FXU_REPORT_FIELDS
                    .dataSetDate,
            );

            remarks.push(
                `${FXU_TEST_DATA_FIELDS.transactionDate} ` +
                `ไม่ตรงกับ ${FXU_REPORT_FIELDS.dataSetDate} ` +
                `และไม่สามารถอ่านวันที่จาก ` +
                `${FXU_TEST_DATA_FIELDS.transactionId} ได้\n` +
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
                `${FXU_TEST_DATA_FIELDS.transactionDate} ` +
                `ไม่ตรงกับ ${FXU_REPORT_FIELDS.dataSetDate} ` +
                `แต่ตรงกับวันที่ใน ` +
                `${FXU_TEST_DATA_FIELDS.transactionId}\n` +
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
                FXU_REPORT_FIELDS
                    .dataSetDate,
            );

            remarks.push(
                `${FXU_TEST_DATA_FIELDS.transactionDate} ` +
                `ไม่ตรงทั้ง ${FXU_REPORT_FIELDS.dataSetDate} ` +
                `และวันที่ใน ` +
                `${FXU_TEST_DATA_FIELDS.transactionId}\n` +
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
                FXU_TEST_DATA_FIELDS
                    .settledAmount,
            );

        const actualAmountText =
            reportRecord.get(
                FXU_REPORT_FIELDS
                    .usdEquivalentAmount,
            );

        const expectedAmount =
            parseFxuAmount(
                expectedAmountText,
            );

        const actualAmount =
            parseFxuAmount(
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
            FXU_AMOUNT_TOLERANCE
        ) {
            addDifference(
                "FAIL",

                FXU_REPORT_FIELDS
                    .usdEquivalentAmount,

                FXU_TEST_DATA_FIELDS
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
                FXU_REPORT_FIELDS
                    .arrangementType,
            );

        if (
            normalizeFxuValue(
                actualArrangementType,
            ) !==
            normalizeFxuValue(
                FXU_ARRANGEMENT_TYPE,
            )
        ) {
            failedHeaders.add(
                FXU_REPORT_FIELDS
                    .arrangementType,
            );

            remarks.push(
                formatFixedValueRemark(
                    FXU_REPORT_CODE,

                    FXU_REPORT_FIELDS
                        .arrangementType,

                    actualArrangementType,

                    FXU_ARRANGEMENT_TYPE,
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
                FXU_REPORT_FIELDS
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
            FxuReconcileDecision,

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
                ? FXU_LEG_TYPES
                    .buyForeignCurrency
                : decision.direction ===
                    "SELL_FCY"
                    ? FXU_LEG_TYPES
                        .sellForeignCurrency
                    : undefined;

        /**
         * ระบุ Leg Type Name ที่คาดหวัง
         *
         * FXU_LEG_TYPE_NAMES ใช้ Leg Type Code เป็น Key
         *
         * ตัวอย่าง:
         * FXU_LEG_TYPE_NAMES["182001"]
         * FXU_LEG_TYPE_NAMES["182002"]
         *
         * จึงใช้ expectedLegType ที่หาได้ด้านบน
         * เป็น Key สำหรับอ่าน Leg Type Name
         */
        const expectedLegTypeName =
            expectedLegType ===
                undefined
                ? undefined
                : FXU_LEG_TYPE_NAMES[
                expectedLegType
                ];

        /**
         * สร้างชื่อ Field สำหรับ Remark
         *
         * ตัวอย่าง:
         * From Currency (CCY)/To Currency (CCY)
         */
        const currencyPairField =
            `${FXU_TEST_DATA_FIELDS.fromCurrency}/` +
            FXU_TEST_DATA_FIELDS
                .toCurrency;

        /**
         * สร้างค่า Currency Pair
         *
         * ตัวอย่าง:
         * THB/USD
         */
        const currencyPairValue =
            `${testDataRecord.get(
                FXU_TEST_DATA_FIELDS
                    .fromCurrency,
            )}/` +
            testDataRecord.get(
                FXU_TEST_DATA_FIELDS
                    .toCurrency,
            );

        const actualLegType =
            reportRecord.get(
                FXU_REPORT_FIELDS
                    .legType,
            );

        const actualLegTypeName =
            reportRecord.get(
                FXU_REPORT_FIELDS
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

                FXU_REPORT_FIELDS
                    .legType,

                currencyPairField,

                currencyPairValue,

                actualLegType,
            );

            addDifference(
                "REVIEW",

                FXU_REPORT_FIELDS
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
            normalizeFxuValue(
                actualLegType,
            ) !==
            normalizeFxuValue(
                expectedLegType,
            )
        ) {
            addDifference(
                "FAIL",

                FXU_REPORT_FIELDS
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
            normalizeFxuValue(
                actualLegTypeName,
            ) !==
            normalizeFxuValue(
                expectedLegTypeName,
            )
        ) {
            addDifference(
                "FAIL",

                FXU_REPORT_FIELDS
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
            FxuMatchResolution,
    ): ResultRow {
        const {
            decision,
            matchedRecord,
        } =
            matchResolution;

        /**
         * ====================================================
         * MUST_NOT_EXIST
         * ====================================================
         *
         * รายการไม่ควรพบใน DF_FXU เช่น:
         * - From Currency = To Currency
         * - Amount >= 1,000,000 USD
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
                "ไม่สามารถตัดสิน Presence Rule ของ DF_FXU ได้",
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
         * รายการต้องพบใน DF_FXU
         */
        if (
            !matchedRecord
        ) {
            return this.createUnmatchedFailureResult(
                testDataRecord,
                matchResolution,
                [
                    "รายการเข้าเงื่อนไขที่ต้องพบใน DF_FXU",
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
                    FXU_REPORT_FIELDS
                        .arrangementNumber,
                ],
            },
        );
    }

    /**
     * สร้างผลสำหรับรายการที่ต้องไม่พบใน DF_FXU
     */
    private resolveMustNotExistResult(
        testDataRecord:
            ReconcileRecord,

        matchResolution:
            FxuMatchResolution,
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
                    "ไม่ควรพบและไม่พบรายการใน DF_FXU",
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
         * พบรายการที่ไม่ควรมีใน DF_FXU
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
                    "พบรายการใน DF_FXU ทั้งที่ไม่ควรพบ",
                    matchResolution.remark,
                ],

                initialFailedHeaders: [
                    FXU_REPORT_FIELDS
                        .arrangementNumber,

                    FXU_REPORT_FIELDS
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
            FxuMatchResolution,

        options:
            FxuCompareOptions =
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
            FxuMatchResolution,

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
                FXU_REPORT_FIELDS
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
     * รายการจึงต้องไม่พบใน DF_FXU
     */
    private getMustNotExistReason(
        testDataRecord:
            ReconcileRecord,

        decision:
            FxuReconcileDecision,
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
                    FXU_TEST_DATA_FIELDS
                        .fromCurrency,
                );

            const toCurrency =
                testDataRecord.get(
                    FXU_TEST_DATA_FIELDS
                        .toCurrency,
                );

            return (
                `From Currency = "${fromCurrency}" และ ` +
                `To Currency = "${toCurrency}" ` +
                "เป็นสกุลเดียวกัน จึงไม่ควรพบใน DF_FXU"
            );
        }

        /**
         * Amount ตั้งแต่ 1,000,000 USD ขึ้นไป
         * ต้องตรวจใน DF_FXM
         */
        const amount =
            parseFxuAmount(
                testDataRecord.get(
                    FXU_TEST_DATA_FIELDS
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
            `>= ${FXU_USD_THRESHOLD.toLocaleString(
                "en-US",
            )} USD ` +
            "จึงต้องตรวจใน DF_FXM และไม่ควรพบใน DF_FXU"
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
                    FXU_TEST_DATA_FIELDS
                        .testNo,
                )
                .trim() ===
            ""
        ) {
            result.remark =
                this.appendRemark(
                    result.remark,

                    `Test Data ไม่มี ` +
                    FXU_TEST_DATA_FIELDS
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
                    FXU_TEST_DATA_FIELDS
                        .transactionId,
                )
                .trim() ===
            ""
        ) {
            result.remark =
                this.appendRemark(
                    result.remark,

                    `Test Data ไม่มี ` +
                    FXU_TEST_DATA_FIELDS
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

    /**
     * ตรวจสอบ Header ของ Report และ Test Data
     */
    private validateHeaders(
        reportHeaders: string[],
        testDataHeaders: string[],
    ): void {
        /**
         * ยืนยันว่า DF_FXU
         * มี Config อยู่ใน mapping-config.ts
         */
        const reportName =
            requireMappingReportName(
                FXU_REPORT_CODE,
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
            FXU_REQUIRED_TEST_DATA_HEADERS,
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
                `[${FXU_REPORT_CODE}] ` +
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
    private indexTestDataRecordsByTestNo(
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
                normalizeFxuValue(
                    testDataRecord.get(
                        FXU_TEST_DATA_FIELDS
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
                        FXU_REPORT_FIELDS
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
                        FXU_TEST_DATA_FIELDS
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

/**
 * Entry Point สำหรับให้
 * script3-compare-report.spec.ts เรียกใช้งาน
 *
 * @param testDataFilePath
 * Path ของ Test Data สำหรับ DF_FXU
 *
 * @returns
 * Path ของไฟล์ DF_FXU Reconcile Result
 */
export const reconcileFxuReport = (
    testDataFilePath:
        string,
): Promise<string> => {
    const reconcileService =
        new FxuReconcileService();

    return reconcileService.reconcile(
        testDataFilePath,
    );
};