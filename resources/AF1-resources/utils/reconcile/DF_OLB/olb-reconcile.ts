/**
 * ควบคุม Flow Reconcile ของ DF_OLB จาก Raw Report และ Test Data
 *
 * 1. ตรวจ Required Header ก่อนนำข้อมูลไปใช้
 * 2. สงวน Exact Row และป้องกันการใช้ Report Row ซ้ำ
 * 3. ให้ Matcher เลือก Candidate และ Analyzer สร้างผลลัพธ์
 * 4. เขียน Result Sheet และบันทึก Workbook สำเนา
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { ReconcileExcelReader } from "../shared/excel-reader";
import type { ReconcileRecord } from "../shared/record";
import { assertRequiredHeaders } from "../shared/required-header-validator";
import {
  ReconcileResultSheetWriter,
  type ResultRow,
} from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import { OlbAnalyzer } from "./olb-analyzer";
import {
  OLB_REPORT_CODE,
  OLB_REPORT_FIELDS,
  OLB_REPORT_HEADER_ROW,
  OLB_REQUIRED_TEST_DATA_HEADERS,
  OLB_TEST_DATA_FIELDS,
  OLB_TEST_DATA_HEADER_ROW,
} from "./olb-config";
import { OlbMatcher } from "./olb-matcher";
import { normalizeOlbText } from "./olb-normalize.util";

class OlbReconcileService {
  async reconcile(testDataFilePath: string): Promise<string> {
    const workbookPreparer = new ReconcileWorkbookPreparer();
    const excelReader = new ReconcileExcelReader();
    const sheetWriter = new ReconcileResultSheetWriter();
    const matcher = new OlbMatcher();
    const analyzer = new OlbAnalyzer();

    console.log(`\n===== RECONCILE - ${OLB_REPORT_CODE} =====`);

    const prepared = await workbookPreparer.prepare(
      OLB_REPORT_CODE,
      OLB_REPORT_HEADER_ROW,
    );

    /** Script 3 อ่าน Raw Report โดยตรง จึงต้อง Fail Fast ก่อน Parse */
    const reportName = requireMappingReportName(OLB_REPORT_CODE);
    assertRequiredHeaders(
      OLB_REPORT_CODE,
      "Raw Report",
      prepared.reportHeaders,
      getUniqueMappingHeaders(reportName),
    );

    const reportData = excelReader.parseWorksheet(
      prepared.reportWorksheet,
      OLB_REPORT_HEADER_ROW,
    );

    const testData = await excelReader.readFile(
      testDataFilePath,
      OLB_TEST_DATA_HEADER_ROW,
    );

    assertRequiredHeaders(
      OLB_REPORT_CODE,
      "Test Data",
      testData.headers,
      OLB_REQUIRED_TEST_DATA_HEADERS,
    );

    const reportRecords = reportData.records.filter(
      (record) =>
        normalizeOlbText(
          record.get(OLB_REPORT_FIELDS.arrangementNumber),
        ) !== "",
    );

    /** สงวน Exact Row ก่อนวน Test Data เพื่อไม่ให้ Fallback แย่งไปใช้ */
    const reservedReportRowNumbers =
      matcher.findReservedReportRows(
        testData.records,
        reportRecords,
      );

    const usedReportRowNumbers = new Set<number>();
    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const results: ResultRow[] = [];

    for (const testDataRecord of testData.records) {
      /**
       * Matcher ต้องใช้ Test Data Row จริงเสมอ
       *
       * กรณี -Return จะไม่เอา Parent Row ไป Match AF1
       * เพราะ Parent ใช้เฉพาะตรวจ Expected Presence / Absence เท่านั้น
       */
      const candidate = matcher.findBestCandidate(
        testDataRecord,
        reportRecords,
        usedReportRowNumbers,
        reservedReportRowNumbers,
      );

      /**
       * หา Context สำหรับ Presence Rule
       *
       * - Test Case ปกติ → ใช้ Row ปัจจุบัน
       * - Test Case ที่ลงท้ายด้วย -Return → ใช้ Parent Test Case
       *   เพื่อตรวจ Currency / Residency
       */
      const presenceContextRecord =
        this.findPresenceContextRecord(
          testDataRecord,
          testData.records,
        );

      const result = analyzer.analyze(
        testDataRecord,
        candidate,
        presenceContextRecord,
      );

      results.push(result);

      if (result.matchedRowNumber === undefined) {
        unmatchedRows.push(result);
        continue;
      }

      annotationByRowNumber.set(
        result.matchedRowNumber,
        result,
      );

      usedReportRowNumbers.add(
        result.matchedRowNumber,
      );
    }

    sheetWriter.writeHeaderRow(
      prepared.resultSheet,
      prepared.reportHeaders,
    );

    const nextRowNumber =
      sheetWriter.writeRowsInRequestedOrder(
        prepared.resultSheet,
        prepared.reportWorksheet,
        prepared.reportHeaders,
        OLB_REPORT_HEADER_ROW + 1,
        prepared.reportWorksheet.rowCount,
        annotationByRowNumber,
        unmatchedRows,
      );

    sheetWriter.finalizeAutoFilter(
      prepared.resultSheet,
      prepared.reportHeaders,
      nextRowNumber - 1,
    );

    prepared.workbook.removeWorksheet(
      prepared.reportWorksheet.id,
    );

    await prepared.workbook.xlsx.writeFile(
      prepared.reconcileFilePath,
    );

    this.logSummary(
      prepared.reconcileFilePath,
      results,
    );

    return prepared.reconcileFilePath;
  }

  /**
   * หา Test Data Row ที่ใช้เป็น Context
   * สำหรับตรวจ Expected Presence / Absence
   *
   * Test Case ปกติ:
   * - ใช้ Test Data Row ปัจจุบัน
   *
   * Test Case ที่ลงท้ายด้วย -Return:
   * - ตัด -Return ออกจาก Test No.
   * - หา Parent Test Case จาก Test Data ทั้งหมด
   * - ใช้ Parent Row ตรวจ Currency / Residency
   *
   * หมายเหตุ:
   * Parent Row ใช้เฉพาะ Presence Rule เท่านั้น
   * ไม่ใช้แทน Return Row ใน Matcher หรือ Field Compare
   */
  private findPresenceContextRecord(
    testDataRecord: ReconcileRecord,
    testDataRecords: ReconcileRecord[],
  ): ReconcileRecord {
    const testNo = normalizeOlbText(
      testDataRecord.get(
        OLB_TEST_DATA_FIELDS.testNo,
      ),
    );

    /**
     * Test Case ปกติ
     * ไม่ต้องหา Parent
     */
    if (!testNo.endsWith("-RETURN")) {
      return testDataRecord;
    }

    const parentTestNo = testNo.replace(
      /-RETURN$/i,
      "",
    );

    const parentRecord = testDataRecords.find(
      (record) =>
        normalizeOlbText(
          record.get(
            OLB_TEST_DATA_FIELDS.testNo,
          ),
        ) === parentTestNo,
    );

    /**
     * ถ้าหา Parent ไม่พบ
     * ให้กลับมาใช้ Return Row ตัวเอง
     *
     * ไม่เดาข้อมูล Currency / Residency
     * เพื่อป้องกันการสรุป Expected Absence ผิด
     */
    return parentRecord ?? testDataRecord;
  }

  private logSummary(
    outputPath: string,
    results: ResultRow[],
  ): void {
    const passCount = results.filter(
      (result) => result.status === "PASS",
    ).length;

    console.log(`Output File : ${outputPath}`);
    console.log(
      `Test Case : ${results.length} | ` +
        `Pass : ${passCount} | ` +
        `Fail : ${results.length - passCount}`,
    );
  }
}

export const reconcileOlbReport = (
  testDataFilePath: string,
): Promise<string> =>
  new OlbReconcileService().reconcile(
    testDataFilePath,
  );