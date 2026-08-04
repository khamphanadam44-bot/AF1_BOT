/**
 * summary.service.ts
 * ------------------------------------------------------------------
 * Script 4 Main Service
 *
 * Flow:
 * 1. หา Reconcile Result ล่าสุดตาม Prefix ของ Report
 * 2. อ่านเฉพาะแถวที่มี Test Script No. และสถานะ Pass/Fail
 * 3. อ่าน Test Data ตาม Header ไม่อิงเลข Column ตายตัว
 * 4. ตรวจ Fee Type 1-N แบบ Dynamic จาก Header และค่าที่ใช้งานจริง
 * 5. จับคู่ Reconcile กับ Test Data ด้วย Reference ก่อน
 *    และใช้ Test Script No. เป็น Fallback
 * 6. ตรวจ Pass + Fail = Total และ Total = จำนวน Summary Details
 * 7. สร้าง Metadata จาก Timestamp ในชื่อ Reconcile Result
 * 8. เรียก SummaryPage เพื่อกรอก Template และ Save แบบ Temp File
 * ------------------------------------------------------------------
 

import * as fs from "fs";
import * as path from "path";
import ExcelJS from "exceljs";

import {
  getCellText,
  normalizeHeader,
  normalizeValue,
} from "../shared/excel-cell.util";
import {
  REQUIRED_MESSAGE,
} from "../shared/excel-style.util";
import type { ReportCode } from "../../../config/report-config";
import {
  getSummaryConfig,
  type HeaderAliases,
  type SummaryFieldMapping,
  type SummaryReportConfig,
} from "./summary-config";
import {
  ReconcileDetailRecord,
  ReconcileSummaryCounts,
  ReconcileSummarySource,
  SummaryDetailRow,
  SummaryFeeGroup,
  SummaryGenerationResult,
  SummaryRunMetadata,
  SummaryTestResult,
  SummaryTestResultColorStyle,
  TestScriptDataRecord,
} from "./summary-model";
import { SummaryPage } from "./summary.page";
import { getLatestFile } from "../../file-system.util";

interface IdentityColumnMap {
  testScriptNo: number;
  referenceTransactionNumber: number;
}

interface ReconcileIdentityColumnMap extends IdentityColumnMap {
  testResult: number;
  reason: number;
}

interface ResolvedSummaryField {
  mapping: SummaryFieldMapping;
  sourceColumn: number;
}

interface FeeColumnGroup {
  feeIndex: number;
  feeTypeColumn?: number;
  feeAccountColumn?: number;
  feeAmountColumn?: number;
  feeCurrencyColumn?: number;
}

const RECONCILE_TIMESTAMP_PATTERN = /_(\d{8})_(\d{6})\.xlsx$/i;

export class SummaryService {
  async generate(
    config: SummaryReportConfig,
  ): Promise<SummaryGenerationResult> {
    console.log(`\n===== SCRIPT 4 SUMMARY - ${config.reportCode} =====`);

    this.validateInputPaths(config);

    const reconcileFilePath = this.findLatestReconcileFile(config);
    const checkedReportFilePath = getLatestFile(
      config.checkedReportDirectory,
    );
    const checkedTestDataFilePath = getLatestFile(
      config.checkedTestDataDirectory,
    );

    const source = await this.readReconcileResult(reconcileFilePath, config);
    const testDataRecords = await this.readTestData(
      checkedTestDataFilePath,
      config,
    );
    const detailRows = this.mapDetails(
      source.details,
      testDataRecords,
      config.referenceSuffixes,
    );

    this.validateSummary(source.counts, detailRows);

    const metadata = this.createMetadata(
      reconcileFilePath,
      config.verifiedBy,
    );
    const summaryFilePath = path.join(
      config.summaryOutputDirectory,
      `${config.summaryFilePrefix}${metadata.runTimestamp}.xlsx`,
    );

    const summaryPage = new SummaryPage();
    await summaryPage.openTemplate(config);
    summaryPage.writeTitleAndReportCode();
    summaryPage.writeMetadata(metadata);
    summaryPage.writeSummaryCounts(source.counts);
    const displayedFeeGroupCount = summaryPage.writeDetails(detailRows);
    /** Copy DS_LTX_Reconcile ทั้ง Tab จาก Reconcile-report 
    await summaryPage.replaceWorksheetFromFile(
      reconcileFilePath,
      config.reconcileSheetName,
      config.reconcileSheetName,
    );

    /** Copy DS_LTX ทั้ง Tab จาก Checked-report-header 
    await summaryPage.replaceWorksheetFromFile(
      checkedReportFilePath,
      config.checkedReportSourceSheetName,
      config.checkedReportTargetSheetName,
    );

    /** Copy Test Data ทั้ง Tab จาก Checked-testdata-header โดยไม่เลื่อน Column 
    await summaryPage.replaceWorksheetFromFile(
      checkedTestDataFilePath,
      config.checkedTestDataSourceSheetName,
      config.checkedTestDataTargetSheetName,
    );

    await summaryPage.save(summaryFilePath);

    this.printExecutionSummary(
      source,
      metadata,
      detailRows.length,
      displayedFeeGroupCount,
      summaryFilePath,
    );

    return {
      summaryFilePath,
      source,
      metadata,
      detailRowCount: detailRows.length,
      displayedFeeGroupCount,
    };
  }

  private validateInputPaths(config: SummaryReportConfig): void {
    const requiredPaths = [
      { label: "Reconcile result folder", value: config.reconcileDirectory },
      { label: "Checked Report folder", value: config.checkedReportDirectory },
      {
        label: "Checked Test Data folder",
        value: config.checkedTestDataDirectory,
      },
      { label: "Summary template file", value: config.templateFilePath },
    ];

    requiredPaths.forEach(({ label, value }) => {
      if (!fs.existsSync(value)) {
        throw new Error(`${label} not found: ${value}`);
      }
    });
  }

  /**
   * เลือกไฟล์จาก Timestamp ในชื่อไฟล์ ไม่ใช้เฉพาะ Modified Time
   * เพื่อให้ผลลัพธ์คงที่แม้มีการ Copy ไฟล์เก่ากลับเข้ามาใหม่
   
  private findLatestReconcileFile(config: SummaryReportConfig): string {
    const escapedPrefix = this.escapeRegExp(config.reconcileFilePrefix);
    const expectedPattern = new RegExp(
      `^${escapedPrefix}(\\d{8})_(\\d{6})\\.xlsx$`,
      "i",
    );

    const candidates = fs
      .readdirSync(config.reconcileDirectory)
      .filter((fileName) => !fileName.startsWith("~$"))
      .map((fileName) => {
        const match = fileName.match(expectedPattern);
        if (!match) {
          return undefined;
        }

        return {
          fileName,
          fullPath: path.join(config.reconcileDirectory, fileName),
          timestampKey: `${match[1]}${match[2]}`,
        };
      })
      .filter(
        (
          candidate,
        ): candidate is {
          fileName: string;
          fullPath: string;
          timestampKey: string;
        } => candidate !== undefined,
      )
      .sort((left, right) => right.timestampKey.localeCompare(left.timestampKey));

    const latestFile = candidates[0];
    if (!latestFile) {
      throw new Error(
        `No valid Reconcile result found in: ${config.reconcileDirectory}\n` +
          `Expected format: ${config.reconcileFilePrefix}YYYYMMDD_HHmmss.xlsx`,
      );
    }

    console.log(`Latest Reconcile Result : ${latestFile.fileName}`);
    return latestFile.fullPath;
  }

  private async readReconcileResult(
    reconcileFilePath: string,
    config: SummaryReportConfig,
  ): Promise<ReconcileSummarySource> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(reconcileFilePath);

    const worksheet = workbook.getWorksheet(config.reconcileSheetName);
    if (!worksheet) {
      throw new Error(
        `Reconcile sheet "${config.reconcileSheetName}" not found in: ` +
          reconcileFilePath,
      );
    }

    const headerMap = this.buildHeaderColumnMap(
      worksheet,
      config.reconcileHeaderRowNumber,
    );
    const columns =
      this.resolveReconcileIdentityColumns(
        headerMap,
        config,
      );
    const resolvedFields =
      this.resolveSummaryFields(
        headerMap,
        config.reconcileFields,
        `Reconcile sheet "${config.reconcileSheetName}"`,
      );
    const details: ReconcileDetailRecord[] = [];

    for (
      let rowNumber = config.reconcileHeaderRowNumber + 1;
      rowNumber <= worksheet.rowCount;
      rowNumber += 1
    ) {
      const row = worksheet.getRow(rowNumber);
      const testScriptNo = this.readDisplayText(row.getCell(columns.testScriptNo));
      const testResultCell =
        row.getCell(
          columns.testResult,
        );
      const testResultText =
        this.readDisplayText(
          testResultCell,
        );

      /**
       * ข้ามเฉพาะแถว AF1 ที่ไม่ได้ผูกกับ Test Case จริง ๆ
       *
       * Test No. ว่างแต่มีผล Pass/Fail ต้องนำเข้า Summary ต่อ
       
      if (
        testScriptNo === "" &&
        testResultText === ""
      ) {
        continue;
      }

      const testResult = this.normalizeTestResult(
        testResultText,
        rowNumber,
        testScriptNo !== ""
          ? testScriptNo
          : `Reconcile row ${rowNumber}`,
      );

      details.push({
        sourceRowNumber: rowNumber,
        testScriptNo,
        testResult,
        testResultColorStyle: this.readTestResultColorStyle(testResultCell),
        reason: this.readDisplayText(row.getCell(columns.reason)),
        referenceTransactionNumber: this.readDisplayText(
          row.getCell(columns.referenceTransactionNumber),
        ),
        fieldValues:
          this.readSummaryFieldValues(
            row,
            resolvedFields,
          ),
      });
    }

    if (details.length === 0) {
      throw new Error(
        `No reconciled Test Case rows found in sheet: ${worksheet.name}`,
      );
    }

    const counts = this.calculateCounts(details);
    this.validateCounts(counts, details.length);

    return {
      reportCode: config.reportCode,
      reconcileFilePath,
      reconcileFileName: path.basename(reconcileFilePath),
      reconcileSheetName: config.reconcileSheetName,
      counts,
      details,
    };
  }

  private async readTestData(
    testDataFilePath: string,
    config: SummaryReportConfig,
  ): Promise<TestScriptDataRecord[]> {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(testDataFilePath);

    const worksheet =
      workbook.getWorksheet(
        config.checkedTestDataSourceSheetName,
      );
    if (!worksheet) {
      throw new Error(
        `Test Data sheet "${config.checkedTestDataSourceSheetName}" not found in: ` +
          testDataFilePath,
      );
    }

    const headerMap = this.buildHeaderColumnMap(
      worksheet,
      config.testDataHeaderRowNumber,
    );
    const columns =
      this.resolveTestDataIdentityColumns(
        headerMap,
        config,
      );
    const resolvedFields =
      this.resolveSummaryFields(
        headerMap,
        config.testDataFields,
        `Test Data sheet "${worksheet.name}"`,
      );
    const feeColumnGroups =
      config.includeFeeGroups
        ? this.detectFeeColumnGroups(
            worksheet,
            config.testDataHeaderRowNumber,
            config,
          )
        : [];
    const records: TestScriptDataRecord[] = [];

    for (
      let rowNumber = config.testDataHeaderRowNumber + 1;
      rowNumber <= worksheet.rowCount;
      rowNumber += 1
    ) {
      const row = worksheet.getRow(rowNumber);
      const rawTestScriptNo =
        this.readDisplayText(
          row.getCell(
            columns.testScriptNo,
          ),
        );
      const testScriptNo =
        this.isValidationPlaceholder(
          rawTestScriptNo,
        )
          ? ""
          : rawTestScriptNo;
      const referenceTransactionNumber = this.readDisplayText(
        row.getCell(columns.referenceTransactionNumber),
      );

      // ข้ามแถว Note / Sum / แถวว่างที่ไม่ใช่ Test Data จริง
      if (testScriptNo === "" && referenceTransactionNumber === "") {
        continue;
      }

      records.push({
        sourceRowNumber: rowNumber,
        testScriptNo,
        referenceTransactionNumber,
        fieldValues:
          this.readSummaryFieldValues(
            row,
            resolvedFields,
          ),
        feeGroups: this.readActiveFeeGroups(row, feeColumnGroups),
      });
    }

    if (records.length === 0) {
      throw new Error(`No Test Data records found in sheet: ${worksheet.name}`);
    }

    return records;
  }

  private resolveReconcileIdentityColumns(
    headerMap: Map<string, number>,
    config: SummaryReportConfig,
  ): ReconcileIdentityColumnMap {
    const headers =
      config.reconcileIdentityHeaders;
    const sourceName = `Reconcile sheet "${config.reconcileSheetName}"`;

    return {
      testScriptNo: this.requireHeaderColumn(
        headerMap,
        headers.testScriptNo,
        sourceName,
      ),
      testResult: this.requireHeaderColumn(
        headerMap,
        headers.testResult,
        sourceName,
      ),
      reason: this.requireHeaderColumn(headerMap, headers.reason, sourceName),
      referenceTransactionNumber: this.requireHeaderColumn(
        headerMap,
        headers.referenceTransactionNumber,
        sourceName,
      ),
    };
  }

  private resolveTestDataIdentityColumns(
    headerMap: Map<string, number>,
    config: SummaryReportConfig,
  ): IdentityColumnMap {
    const headers =
      config.testDataIdentityHeaders;
    const sourceName =
      `Test Data sheet "${config.checkedTestDataSourceSheetName}"`;

    return {
      testScriptNo: this.requireHeaderColumn(
        headerMap,
        headers.testScriptNo,
        sourceName,
      ),
      referenceTransactionNumber: this.requireHeaderColumn(
        headerMap,
        headers.referenceTransactionNumber,
        sourceName,
      ),
    };
  }

  /**
   * Resolve Header ของ Field ที่จะแสดงใน Summary
   * จาก Config ของ Report ปัจจุบัน
   
  private resolveSummaryFields(
    headerMap: Map<string, number>,
    mappings: readonly SummaryFieldMapping[],
    sourceName: string,
  ): ResolvedSummaryField[] {
    return mappings.map(
      (mapping): ResolvedSummaryField => ({
        mapping,
        sourceColumn:
          this.requireHeaderColumn(
            headerMap,
            mapping.sourceHeader,
            sourceName,
          ),
      }),
    );
  }

  /**
   * อ่านค่าตาม Field Mapping
   * และเก็บเป็น Key/Value กลางเพื่อรองรับหลาย Report
   
  private readSummaryFieldValues(
    row: ExcelJS.Row,
    resolvedFields: readonly ResolvedSummaryField[],
  ): Readonly<Record<string, string>> {
    return Object.fromEntries(
      resolvedFields.map(
        ({ mapping, sourceColumn }) => [
          mapping.key,
          mapping.valueType === "number"
            ? this.readNumericText(
                row.getCell(sourceColumn),
              )
            : this.readDisplayText(
                row.getCell(sourceColumn),
              ),
        ],
      ),
    );
  }

  private detectFeeColumnGroups(
    worksheet: ExcelJS.Worksheet,
    headerRowNumber: number,
    config: SummaryReportConfig,
  ): FeeColumnGroup[] {
    const groups = new Map<number, FeeColumnGroup>();
    const headerRow = worksheet.getRow(headerRowNumber);
    const maxColumn = Math.max(
      worksheet.columnCount,
      headerRow.cellCount,
      headerRow.actualCellCount,
    );

    const setColumn = (
      feeIndex: number,
      field: keyof Omit<FeeColumnGroup, "feeIndex">,
      columnNumber: number,
    ): void => {
      const existing = groups.get(feeIndex) ?? { feeIndex };
      existing[field] = columnNumber;
      groups.set(feeIndex, existing);
    };

    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      const header = normalizeHeader(getCellText(headerRow.getCell(columnNumber)));
      if (header === "") {
        continue;
      }

      const matches: Array<{
        pattern: RegExp;
        field: keyof Omit<FeeColumnGroup, "feeIndex">;
      }> = [
        { pattern: config.feeHeaderPatterns.feeType, field: "feeTypeColumn" },
        { pattern: config.feeHeaderPatterns.feeAccount, field: "feeAccountColumn" },
        { pattern: config.feeHeaderPatterns.feeAmount, field: "feeAmountColumn" },
        { pattern: config.feeHeaderPatterns.feeCurrency, field: "feeCurrencyColumn" },
      ];

      for (const { pattern, field } of matches) {
        const match = header.match(pattern);
        if (!match) {
          continue;
        }

        const feeIndex = Number(match[1]);
        if (Number.isInteger(feeIndex) && feeIndex > 0) {
          setColumn(feeIndex, field, columnNumber);
        }
        break;
      }
    }

    const result = [...groups.values()].sort(
      (left, right) => left.feeIndex - right.feeIndex,
    );

    if (result.length === 0) {
      throw new Error(
        `No Fee Type headers found in Test Data sheet: ${worksheet.name}`,
      );
    }

    result.forEach((group) => {
      const missingHeaders = [
        ["Fee Type", group.feeTypeColumn],
        ["Fee Charge Account No.", group.feeAccountColumn],
        ["Fee Amount", group.feeAmountColumn],
        ["Fee Currency", group.feeCurrencyColumn],
      ]
        .filter(([, columnNumber]) => columnNumber === undefined)
        .map(([headerName]) => headerName);

      if (missingHeaders.length > 0) {
        console.warn(
          `Warning: Fee Group ${group.feeIndex} has incomplete headers: ` +
            missingHeaders.join(", "),
        );
      }
    });

    return result;
  }

  /**
   * เก็บเฉพาะ Fee Group ที่มีข้อมูลจริงในแถวนั้น
   * จึงไม่สร้าง Fee Type 3-5 ใน Summary เพียงเพราะ Test Data มี Header ว่างไว้ล่วงหน้า
   
  private readActiveFeeGroups(
    row: ExcelJS.Row,
    feeColumnGroups: FeeColumnGroup[],
  ): SummaryFeeGroup[] {
    return feeColumnGroups
      .map((group): SummaryFeeGroup => ({
        feeIndex: group.feeIndex,
        feeType: this.readOptionalCell(row, group.feeTypeColumn),
        feeChargeAccountNo: this.readOptionalCell(row, group.feeAccountColumn),
        feeAmount: this.readOptionalNumericCell(row, group.feeAmountColumn),
        feeCurrency: this.readOptionalCell(row, group.feeCurrencyColumn),
      }))
      .filter((group) =>
        [
          group.feeType,
          group.feeChargeAccountNo,
          group.feeAmount,
          group.feeCurrency,
        ].some((value) => value !== ""),
      );
  }

  private mapDetails(
    reconcileDetails: ReconcileDetailRecord[],
    testDataRecords: TestScriptDataRecord[],
    referenceSuffixes: readonly string[],
  ): SummaryDetailRow[] {
    const byReference = this.buildReferenceIndex(
      testDataRecords,
      referenceSuffixes,
    );
    const byTestScriptNo = this.buildTestScriptIndex(testDataRecords);
    const missingTestNoRecords =
      testDataRecords.filter(
        (record) =>
          this.normalizeKey(
            record.testScriptNo,
          ) === "",
      );

    return reconcileDetails.map((reconcile) => ({
      reconcile,
      testData: this.findMatchingTestData(
        reconcile,
        byReference,
        byTestScriptNo,
        referenceSuffixes,
        missingTestNoRecords,
      ),
    }));
  }

  private buildReferenceIndex(
    records: TestScriptDataRecord[],
    referenceSuffixes: readonly string[],
  ): Map<string, TestScriptDataRecord> {
    const index = new Map<string, TestScriptDataRecord>();

    records.forEach((record) => {
      const key = this.normalizeReference(
        record.referenceTransactionNumber,
        referenceSuffixes,
      );

      /**
       * ข้ามค่าที่ว่างและข้อความ Validation จาก Script 2
       * เพราะข้อความดังกล่าวใช้แสดงผลเท่านั้น
       * ไม่ใช่ Transaction ID สำหรับจับคู่ข้อมูล
       
      if (
        key === "" ||
        this.isValidationPlaceholder(
          record.referenceTransactionNumber,
        )
      ) {
        return;
      }

      const existing = index.get(key);
      if (existing) {
        throw new Error(
          `Duplicate Transaction ID/ Reconcile ID in Test Data: ` +
            `"${record.referenceTransactionNumber}" at rows ` +
            `${existing.sourceRowNumber} and ${record.sourceRowNumber}.`,
        );
      }

      index.set(key, record);
    });

    return index;
  }

  private buildTestScriptIndex(
    records: TestScriptDataRecord[],
  ): Map<string, TestScriptDataRecord[]> {
    const index = new Map<string, TestScriptDataRecord[]>();

    records.forEach((record) => {
      const testScriptKey =
        this.normalizeKey(
          record.testScriptNo,
        );
      const referenceFallbackKey =
        this.normalizeKey(
          record.referenceTransactionNumber,
        );
      const key =
        testScriptKey !== ""
          ? testScriptKey
          : referenceFallbackKey;

      /**
       * ข้ามค่าที่ว่างและข้อความ Validation จาก Script 2
       * เพื่อไม่ให้ข้อความ "โปรดกรอกข้อมูล"
       * ถูกใช้เป็น Test Script Key
       
      if (
        key === "" ||
        this.isValidationPlaceholder(
          record.testScriptNo,
        ) ||
        this.isValidationPlaceholder(
          record.referenceTransactionNumber,
        )
      ) {
        return;
      }

      const candidates = index.get(key) ?? [];
      candidates.push(record);
      index.set(key, candidates);
    });

    return index;
  }

  private findMatchingTestData(
    reconcile: ReconcileDetailRecord,
    byReference: Map<string, TestScriptDataRecord>,
    byTestScriptNo: Map<string, TestScriptDataRecord[]>,
    referenceSuffixes: readonly string[],
    missingTestNoRecords: TestScriptDataRecord[],
  ): TestScriptDataRecord {
    const referenceKey = this.normalizeReference(
      reconcile.referenceTransactionNumber,
      referenceSuffixes,
    );

    if (referenceKey !== "") {
      const referenceMatch = byReference.get(referenceKey);
      if (referenceMatch) {
        return referenceMatch;
      }
    }

    const testScriptKey = this.normalizeKey(reconcile.testScriptNo);
    const candidates = byTestScriptNo.get(testScriptKey) ?? [];

    if (candidates.length === 1) {
      return candidates[0];
    }

    if (candidates.length > 1) {
      throw new Error(
        `Ambiguous Test Data match for Reconcile row ${reconcile.sourceRowNumber}: ` +
          `Test Script No. = "${reconcile.testScriptNo}" matched ` +
          `${candidates.length} Test Data rows, and Reference ` +
          `"${reconcile.referenceTransactionNumber}" did not identify one row.`,
      );
    }

    /**
     * Expected Absence อาจไม่มีข้อมูล Raw Report
     * ทำให้ Reconcile row ไม่มี Reference สำหรับจับคู่
     *
     * หาก Test Script No. ว่าง และมี Test Data ที่ขาด Test No.
     * เพียงหนึ่งแถว สามารถจับคู่แถวนั้นได้อย่างแน่นอน
     *
     * หากมีหลายแถวจะหยุดด้วย Structural Error
     * เพื่อป้องกันการเลือก Test Data ผิดแถวโดยอาศัยลำดับ
     
    if (
      referenceKey === "" &&
      testScriptKey === ""
    ) {
      if (
        missingTestNoRecords.length === 1
      ) {
        return missingTestNoRecords[0];
      }

      if (
        missingTestNoRecords.length > 1
      ) {
        throw new Error(
          `Ambiguous Test Data match for Reconcile row ` +
            `${reconcile.sourceRowNumber}: Test Script No. is empty ` +
            `and ${missingTestNoRecords.length} Test Data rows ` +
            `have no Test No.`,
        );
      }
    }

    throw new Error(
      `Test Data record not found for Reconcile row ${reconcile.sourceRowNumber}: ` +
        `Test Script No. = "${reconcile.testScriptNo}", Reference = ` +
        `"${reconcile.referenceTransactionNumber}".`,
    );
  }

  private createMetadata(
    reconcileFilePath: string,
    verifiedBy: string,
  ): SummaryRunMetadata {
    const reportFileName = path.basename(reconcileFilePath);
    const match = reportFileName.match(RECONCILE_TIMESTAMP_PATTERN);

    if (!match) {
      throw new Error(
        `Invalid Reconcile file name: "${reportFileName}". ` +
          `Expected: <REPORT>_Reconcile_YYYYMMDD_HHmmss.xlsx`,
      );
    }

    const [, datePart, timePart] = match;
    const runTimestamp = `${datePart}_${timePart}`;

    return {
      reportFileName,
      executionDate:
        `${datePart.slice(0, 4)}-${datePart.slice(4, 6)}-` +
        datePart.slice(6, 8),
      executionTime:
        `${timePart.slice(0, 2)}:${timePart.slice(2, 4)}:` +
        timePart.slice(4, 6),
      runId: `RUN_${runTimestamp}`,
      verifiedBy,
      runTimestamp,
    };
  }

  private calculateCounts(
    details: ReconcileDetailRecord[],
  ): ReconcileSummaryCounts {
    const passed = details.filter((row) => row.testResult === "Pass").length;
    const failed = details.filter((row) => row.testResult === "Fail").length;

    return {
      totalChecked: details.length,
      passed,
      failed,
    };
  }

  private validateSummary(
    counts: ReconcileSummaryCounts,
    detailRows: SummaryDetailRow[],
  ): void {
    this.validateCounts(counts, detailRows.length);

    const duplicateSourceRows = detailRows
      .map((detail) => detail.reconcile.sourceRowNumber)
      .filter((rowNumber, index, rows) => rows.indexOf(rowNumber) !== index);

    if (duplicateSourceRows.length > 0) {
      throw new Error(
        `Duplicate Reconcile rows found in Summary Details: ` +
          [...new Set(duplicateSourceRows)].join(", "),
      );
    }
  }

  private validateCounts(
    counts: ReconcileSummaryCounts,
    detailRowCount: number,
  ): void {
    const calculatedTotal = counts.passed + counts.failed;

    if (calculatedTotal !== counts.totalChecked) {
      throw new Error(
        `Summary count mismatch: Pass (${counts.passed}) + Fail ` +
          `(${counts.failed}) = ${calculatedTotal}, but Total Checked = ` +
          `${counts.totalChecked}.`,
      );
    }

    if (counts.totalChecked !== detailRowCount) {
      throw new Error(
        `Summary detail mismatch: Total Checked = ${counts.totalChecked}, ` +
          `but Summary Details = ${detailRowCount}.`,
      );
    }
  }

  private normalizeTestResult(
    rawStatus: string,
    rowNumber: number,
    testScriptNo: string,
  ): SummaryTestResult {
    const normalizedStatus = normalizeValue(rawStatus).toUpperCase();

    if (normalizedStatus === "PASS") {
      return "Pass";
    }

    if (normalizedStatus === "FAIL") {
      return "Fail";
    }

    throw new Error(
      `Invalid Test Result at Reconcile row ${rowNumber}: "${rawStatus}". ` +
        `Only Pass or Fail is allowed when Test Script No. = ` +
        `"${testScriptNo}".`,
    );
  }

  private buildHeaderColumnMap(
    worksheet: ExcelJS.Worksheet,
    headerRowNumber: number,
  ): Map<string, number> {
    const result = new Map<string, number>();
    const headerRow = worksheet.getRow(headerRowNumber);
    const maxColumn = Math.max(
      worksheet.columnCount,
      headerRow.cellCount,
      headerRow.actualCellCount,
    );

    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      const header = normalizeHeader(getCellText(headerRow.getCell(columnNumber)));

      if (header !== "" && !result.has(header)) {
        result.set(header, columnNumber);
      }
    }

    return result;
  }

  private requireHeaderColumn(
    headerMap: Map<string, number>,
    headerConfig: HeaderAliases,
    sourceName: string,
  ): number {
    for (const alias of headerConfig.aliases) {
      const columnNumber = headerMap.get(normalizeHeader(alias));
      if (columnNumber !== undefined) {
        return columnNumber;
      }
    }

    throw new Error(
      `Required header not found in ${sourceName}: ` +
        headerConfig.aliases.map((alias) => `"${alias}"`).join(" or "),
    );
  }


  /**
   * อ่านเฉพาะสี Fill และสี Font จาก Cell Test Result ใน Reconcile Result
   * ไม่ Copy Border, Number Format หรือ Alignment เพื่อไม่ทำลายรูปแบบ Template
   
  private readTestResultColorStyle(
    cell: ExcelJS.Cell,
  ): SummaryTestResultColorStyle {
    const fillArgb =
      cell.fill.type === "pattern"
        ? cell.fill.fgColor?.argb
        : undefined;

    return {
      fillArgb,
      fontColorArgb: cell.font?.color?.argb,
    };
  }

  private readDisplayText(cell: ExcelJS.Cell): string {
    if (cell.value === null || cell.value === undefined) {
      return "";
    }

    const displayedText = cell.text;
    const rawText = displayedText !== "" ? displayedText : getCellText(cell);

    return normalizeValue(rawText);
  }

  private readNumericText(cell: ExcelJS.Cell): string {
    const value = cell.value;

    if (typeof value === "number") {
      return String(value);
    }

    // บาง Test Data Column ถูก Format เป็นเวลาไว้ล่วงหน้า แม้ข้อมูลจริงเป็น Amount
    // ExcelJS จึงอ่านเลข Excel Serial กลับมาเป็น Date; แปลงกลับเป็นเลข Amount ก่อน
    if (value instanceof Date) {
      const excelEpoch = Date.UTC(1899, 11, 30);
      const excelSerial = (value.getTime() - excelEpoch) / 86_400_000;
      return String(Number(excelSerial.toFixed(10)));
    }

    if (
      value !== null &&
      value !== undefined &&
      typeof value === "object" &&
      "result" in value &&
      typeof value.result === "number"
    ) {
      return String(value.result);
    }

    return this.readDisplayText(cell);
  }

  private readOptionalCell(
    row: ExcelJS.Row,
    columnNumber: number | undefined,
  ): string {
    if (columnNumber === undefined) {
      return "";
    }

    return this.readDisplayText(row.getCell(columnNumber));
  }

  private readOptionalNumericCell(
    row: ExcelJS.Row,
    columnNumber: number | undefined,
  ): string {
    if (columnNumber === undefined) {
      return "";
    }

    return this.readNumericText(row.getCell(columnNumber));
  }

  /**
   * ตรวจว่าค่าใน Cell เป็นข้อความที่ Script 2
   * ใส่แทน Required Field ที่ว่างหรือไม่
   *
   * ข้อความนี้ใช้สำหรับแสดงผล Validation เท่านั้น
   * ห้ามนำไปใช้เป็น Matching Key ของ Script 4
   
  private isValidationPlaceholder(
    value: string,
  ): boolean {
    return (
      normalizeValue(value) ===
      normalizeValue(REQUIRED_MESSAGE)
    );
  }

  private normalizeReference(
    value: string,
    referenceSuffixes: readonly string[],
  ): string {
    let normalized = this.normalizeKey(value);

    const suffix = referenceSuffixes.find((candidate) =>
      normalized.endsWith(candidate.toUpperCase()),
    );

    if (suffix) {
      normalized = normalized.slice(0, -suffix.length);
    }

    return normalized;
  }

  private normalizeKey(value: string): string {
    return normalizeValue(value).replace(/\s+/g, "").toUpperCase();
  }

  private escapeRegExp(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  private printExecutionSummary(
    source: ReconcileSummarySource,
    metadata: SummaryRunMetadata,
    detailRowCount: number,
    feeGroupCount: number,
    summaryFilePath: string,
  ): void {
    console.log(`Report File Name : ${metadata.reportFileName}`);
    console.log(`Execution Date   : ${metadata.executionDate}`);
    console.log(`Execution Time   : ${metadata.executionTime}`);
    console.log(`Run ID           : ${metadata.runId}`);
    console.log(`Verified By      : ${metadata.verifiedBy}`);
    console.log(`Total Checked    : ${source.counts.totalChecked}`);
    console.log(`Passed / Match   : ${source.counts.passed}`);
    console.log(`Failed / Unmatch : ${source.counts.failed}`);
    console.log(`Summary Details  : ${detailRowCount} rows`);
    console.log(`Fee Groups       : ${feeGroupCount}`);
    console.log(`Summary File     : ${summaryFilePath}`);
  }
}

/** Function wrapper สำหรับเรียกจาก Mocha Test หรือ Runner 
export const generateSummaryReport = (
  reportCode: ReportCode,
): Promise<SummaryGenerationResult> => {
  const config = getSummaryConfig(reportCode);
  return new SummaryService().generate(config);
};
*/
