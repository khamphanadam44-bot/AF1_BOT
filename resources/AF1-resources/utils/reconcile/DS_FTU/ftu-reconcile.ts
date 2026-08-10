/**
 * ftu-reconcile.ts
 * ------------------------------------------------------------------
 * Entry Point ของ Script 3 สำหรับ DS_FTU
 *
 * หน้าที่:
 * 1. เตรียม Workbook และอ่านข้อมูล
 * 2. ตรวจ Header ที่จำเป็น
 * 3. ส่งแต่ละ Test Case ให้ Matcher และ Analyzer
 * 4. เขียนผล Reconcile และสรุปผล
 *
 * Business Matching อยู่ใน ftu-matcher.ts
 * Business Validation อยู่ใน ftu-analyzer.ts
 * ------------------------------------------------------------------
 */

import {
  getUniqueMappingHeaders,
  requireMappingReportName,
} from "../../../config/mapping-helper";
import { canonicalHeader } from "../../validators/shared/header-matcher";
import { ReconcileExcelReader } from "../shared/excel-reader";
import { ReconcileResultSheetWriter } from "../shared/result-writer";
import type { ResultRow } from "../shared/result-writer";
import { ReconcileWorkbookPreparer } from "../shared/workbook-preparer";
import { FtuAnalyzer } from "./ftu-analyzer";
import {
  FTU_REPORT_CODE,
  FTU_REPORT_HEADER_ROW,
  FTU_REQUIRED_TEST_DATA_HEADERS,
  FTU_TEST_DATA_HEADER_ROW,
} from "./ftu-config";
import { FtuMatcher } from "./ftu-matcher";

export class FtuReconcileService {
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
    const usedReportRowNumbers = matcher.findReservedReportRows(
      testData.records,
      reportRecordsById,
    );
    const annotationByRowNumber = new Map<number, ResultRow>();
    const unmatchedRows: ResultRow[] = [];
    const results: ResultRow[] = [];

    for (const testDataRecord of testData.records) {
      const matchResult = matcher.findMatch(
        testDataRecord,
        reportRecordsById,
        reportData.records,
        usedReportRowNumbers,
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

  /** ตรวจ Header ที่จำเป็นก่อนเริ่ม Reconcile */
  private validateHeaders(
    reportHeaders: string[],
    testDataHeaders: string[],
  ): void {
    const reportName = requireMappingReportName(FTU_REPORT_CODE);
    const requiredReportHeaders = getUniqueMappingHeaders(reportName);

    this.assertHeaders(reportHeaders, requiredReportHeaders, "Raw Report");
    this.assertHeaders(
      testDataHeaders,
      FTU_REQUIRED_TEST_DATA_HEADERS,
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
        `[${FTU_REPORT_CODE}] ` +
          `${sourceName} missing header(s): ` +
          missingHeaders.join(", "),
      );
    }
  }

  private logSummary(outputPath: string, results: ResultRow[]): void {
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

export const reconcileFtuReport = (testDataFilePath: string): Promise<string> =>
  new FtuReconcileService().reconcile(testDataFilePath);