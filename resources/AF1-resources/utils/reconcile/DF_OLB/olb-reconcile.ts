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
        normalizeOlbText(record.get(OLB_REPORT_FIELDS.arrangementNumber)) !==
        "",
    );

    /** สงวน Exact Row ก่อนวน Test Data เพื่อไม่ให้ Fallback แย่งไปใช้ */
    const reservedReportRowNumbers = matcher.findReservedReportRows(
      testData.records,
      reportRecords,
    );
    const usedReportRowNumbers = new Set<number>();
    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const results: ResultRow[] = [];

    for (const testDataRecord of testData.records) {
      const candidate = matcher.findBestCandidate(
        testDataRecord,
        reportRecords,
        usedReportRowNumbers,
        reservedReportRowNumbers,
      );
      const result = analyzer.analyze(testDataRecord, candidate);

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
    prepared.workbook.removeWorksheet(prepared.reportWorksheet.id);
    await prepared.workbook.xlsx.writeFile(prepared.reconcileFilePath);

    this.logSummary(prepared.reconcileFilePath, results);
    return prepared.reconcileFilePath;
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

export const reconcileOlbReport = (
  testDataFilePath: string,
): Promise<string> => new OlbReconcileService().reconcile(testDataFilePath);