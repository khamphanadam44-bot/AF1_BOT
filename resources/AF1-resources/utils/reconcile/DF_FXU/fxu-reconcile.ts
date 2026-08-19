/**
 * fxu-reconcile.ts
 * ------------------------------------------------------------------
 * Orchestrator หลักของ Script 3 สำหรับ DF_FXU
 * ไฟล์นี้คง Public Entry Point เดิม เพื่อไม่กระทบ Script อื่น
 * ------------------------------------------------------------------
 */

import {
    FXU_REPORT_CODE,
    FXU_REPORT_HEADER_ROW,
} from "./fxu-config";

import {
    FxuReconcileEvaluator,
} from "./fxu-reconcile-evaluator";

import {
    ReconcileResultSheetWriter,
    ResultRow,
} from "../shared/result-writer";

export * from "./fxu-reconcile-model";
export { FxuReconcileMatcher } from "./fxu-reconcile-matcher";
export { FxuReconcileEvaluator } from "./fxu-reconcile-evaluator";

/**
 * ควบคุมลำดับการ Reconcile และเขียนผลลัพธ์ลง Excel
 */
export class FxuReconcileService extends FxuReconcileEvaluator {

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
