/**
 * olb-reconcile.ts
 * ------------------------------------------------------------------
 * Orchestrator ของ DF_OLB Script 3
 *
 * Flow:
 * 1. เตรียม Workbook และอ่านข้อมูล
 * 2. ตรวจ Header
 * 3. ให้ Matcher หา AF1 Row
 * 4. ให้ Analyzer สร้าง PASS/FAIL/Remark
 * 5. เขียนและบันทึก Result Workbook
 * ------------------------------------------------------------------
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { canonicalHeader } from "../../validators/shared/header-matcher";
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
import { ReconcileExcelReader } from "../shared/excel-reader";
import {
  ReconcileResultSheetWriter,
  ResultRow,
} from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";

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
    const reportData = excelReader.parseWorksheet(
      prepared.reportWorksheet,
      OLB_REPORT_HEADER_ROW,
    );
    const testData = await excelReader.readFile(
      testDataFilePath,
      OLB_TEST_DATA_HEADER_ROW,
    );

    this.validateHeaders(
      prepared.reportHeaders, 
      testData.headers);

    const reportRecords = reportData.records.filter(
      (record) =>
        normalizeOlbText(
          record.get(OLB_REPORT_FIELDS.arrangementNumber),
        ) !== "",
    );

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

      annotationByRowNumber.set(
        result.matchedRowNumber,
        result,
      );
      usedReportRowNumbers.add(
        result.matchedRowNumber);
    }

    sheetWriter.writeHeaderRow(
      prepared.resultSheet, 
      prepared.reportHeaders);

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

    prepared.workbook.removeWorksheet(
      prepared.reportWorksheet.id);

    await prepared.workbook.xlsx.writeFile(
      prepared.reconcileFilePath);

    this.logSummary(
      prepared.reconcileFilePath, 
      results
    );
    return prepared.reconcileFilePath;
  }

  private validateHeaders(
    reportHeaders: string[],
    testDataHeaders: string[],
  ): void {
    const reportName = requireMappingReportName(OLB_REPORT_CODE);

    this.assertHeaders(
      reportHeaders,
      getUniqueMappingHeaders(reportName),
      "Raw Report",
    );

    this.assertHeaders(
      testDataHeaders,
      OLB_REQUIRED_TEST_DATA_HEADERS,
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
        `[${OLB_REPORT_CODE}] ${sourceName} missing header(s): ` +
          missingHeaders.join(", "),
      );
    }
  }

  private logSummary(outputPath: string, 
    results: ResultRow[]): void {
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

export const reconcileOlbReport = (
  testDataFilePath: string,
): Promise<string> => 
  new OlbReconcileService().reconcile(testDataFilePath);