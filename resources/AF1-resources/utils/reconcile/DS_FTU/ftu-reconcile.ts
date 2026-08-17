/**
 * ควบคุม Flow Reconcile ของ DS_FTU
 *
 * 1. เตรียม Raw Report และอ่าน Test Data
 * 2. ตรวจ Header ที่จำเป็นก่อนเริ่ม Matching
 * 3. สงวน Exact Row และป้องกัน Report Row ถูกใช้ซ้ำ
 * 4. ส่งผล Matching ให้ Analyzer ตัดสิน PASS/FAIL/Review
 * 5. เขียน Result Sheet และสรุปผล
 *
 * Matching และ Business Validation แยกอยู่ใน Matcher และ Analyzer
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { assertRequiredHeaders } from "../shared/required-header-validator";
import {
  ReconcileResultSheetWriter,
  type ResultRow,
} from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import { FtuAnalyzer } from "./ftu-analyzer";
import {
  FTU_REPORT_CODE,
  FTU_REPORT_HEADER_ROW,
  FTU_REQUIRED_TEST_DATA_HEADERS,
  FTU_TEST_DATA_HEADER_ROW,
} from "./ftu-config";
import { FtuMatcher } from "./ftu-matcher";

class FtuReconcileService {
  async reconcile(testDataFilePath: string): Promise<string> {
    const workbookPreparer = new ReconcileWorkbookPreparer();
    const excelReader = new ReconcileExcelReader();
    const sheetWriter = new ReconcileResultSheetWriter();
    const matcher = new FtuMatcher();
    const analyzer = new FtuAnalyzer();

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

    const reportRecordsById = matcher.indexReportRecords(reportData.records);
    /** Reserved Row ป้องกัน Fallback แย่ง Exact Row; Used Row กันการใช้ซ้ำ */
    const reservedReportRowNumbers = matcher.findReservedReportRows(
      testData.records,
      reportRecordsById,
    );
    const usedReportRowNumbers = new Set<number>();
    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const results: ResultRow[] = [];

    for (const testDataRecord of testData.records) {
      const matchResult = matcher.findMatch(
        testDataRecord,
        reportRecordsById,
        reportData.records,
        usedReportRowNumbers,
        reservedReportRowNumbers,
      );
      const result = analyzer.analyze(testDataRecord, matchResult);

      results.push(result);

      if (result.matchedRowNumber === undefined) {
        unmatchedRows.push(result);
        continue;
      }

      annotationByRowNumber.set(result.matchedRowNumber, result);
      usedReportRowNumbers.add(result.matchedRowNumber);
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

  /** ใช้ Header Guard กลางเพื่อให้ Raw Report และ Test Data ตรวจแบบเดียวกัน */
  private validateHeaders(
    reportHeaders: string[],
    testDataHeaders: string[],
  ): void {
    const reportName = requireMappingReportName(FTU_REPORT_CODE);

    assertRequiredHeaders(
      FTU_REPORT_CODE,
      "Raw Report",
      reportHeaders,
      getUniqueMappingHeaders(reportName),
    );
    assertRequiredHeaders(
      FTU_REPORT_CODE,
      "Test Data",
      testDataHeaders,
      FTU_REQUIRED_TEST_DATA_HEADERS,
    );
  }

  private logSummary(outputPath: string, results: ResultRow[]): void {
    const passCount = results.filter((result) => result.status === "PASS").length;

    console.log(`Output File : ${outputPath}`);
    console.log(
      `Test Case : ${results.length} | ` +
        `Pass : ${passCount} | ` +
        `Fail : ${results.length - passCount}`,
    );
  }
}

export const reconcileFtuReport = (testDataFilePath: string): Promise<string> =>
  new FtuReconcileService().reconcile(testDataFilePath);