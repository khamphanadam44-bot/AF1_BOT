/**
 * summary.page.ts
 * ------------------------------------------------------------------
 * - openTemplate()       = เปิดหน้าหรือไฟล์ที่ต้องใช้งาน
 * - writeMetadata()      = กรอกข้อมูลส่วนบน
 * - writeSummaryCounts() = กรอก KPI
 * - writeDetails()       = กรอกตารางผลลัพธ์
 * - save()               = บันทึกผลลัพธ์
 *
 * Class นี้ไม่อ่าน Reconcile Result และไม่จับคู่ Test Data
 * จึงไม่มี Business Logic ปะปนกับ Logic การจัดการ Template
 * ------------------------------------------------------------------
 

import * as fs from "fs";
import * as path from "path";
import ExcelJS from "exceljs";

import {
  buildTempFilePath,
  cleanupStaleTempFiles,
  deleteFileIfExists,
  ensureDirectoryExists,
} from "../../file-system.util";
import {
  SummaryReportConfig,
  SummaryTemplateLocator,
} from "./summary-config";
import {
  ReconcileSummaryCounts,
  SummaryDetailRow,
  SummaryRunMetadata,
  SummaryTestResultColorStyle,
} from "./summary-model";

const cloneStyle = (style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> =>
  JSON.parse(JSON.stringify(style)) as Partial<ExcelJS.Style>;

type WorksheetWithDataValidations = ExcelJS.Worksheet & {
  dataValidations: {
    model: unknown;
  };
};
export class SummaryPage {
  private readonly workbook = new ExcelJS.Workbook();
  private worksheet: ExcelJS.Worksheet | undefined;
  private config: SummaryReportConfig | undefined;
  private dynamicDetailEndColumn = 0;
  private displayedFeeGroupCount = 0;

  async openTemplate(config: SummaryReportConfig): Promise<void> {
    if (!fs.existsSync(config.templateFilePath)) {
      throw new Error(`Summary template file not found: ${config.templateFilePath}`);
    }

    this.config = config;
    await this.workbook.xlsx.readFile(config.templateFilePath);

    this.worksheet = this.workbook.getWorksheet(config.templateSheetName);
    if (!this.worksheet) {
      throw new Error(
        `Summary template sheet "${config.templateSheetName}" not found in: ` +
          config.templateFilePath,
      );
    }

    this.dynamicDetailEndColumn = config.locator.fixedDetailEndColumn;
    this.displayedFeeGroupCount = config.locator.templateFeeGroupCount;
  }

  writeTitleAndReportCode(): void {
    const worksheet = this.getWorksheet();
    const config = this.getConfig();
    const displayReportCode = config.reportCode.replace(/_/g, "-");

    worksheet.getCell(config.locator.title).value =
      `${displayReportCode} AUTOMATION VERIFICATION SUMMARY`;
    worksheet.getCell(config.locator.reportCode).value = config.reportCode;
    worksheet.getCell(config.locator.testScriptDataTitle).value = "Test Script Data";
  }

  writeMetadata(metadata: SummaryRunMetadata): void {
    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;

    worksheet.getCell(locator.reportFileName).value = metadata.reportFileName;
    worksheet.getCell(locator.executionDate).value = metadata.executionDate;
    worksheet.getCell(locator.executionTime).value = metadata.executionTime;
    worksheet.getCell(locator.runId).value = metadata.runId;
    worksheet.getCell(locator.verifiedBy).value = metadata.verifiedBy;
  }

  writeSummaryCounts(counts: ReconcileSummaryCounts): void {
    const locator = this.getConfig().locator;

    this.writeCountCell(locator.totalChecked, counts.totalChecked);
    this.writeCountCell(locator.passedMatch, counts.passed);
    this.writeCountCell(locator.failedUnmatch, counts.failed);
  }

  /**
   * เขียน Summary Details โดยใช้ Value/Text เป็นหลัก
   * ยกเว้น Cell Test Result ที่ Copy เฉพาะสี Fill และสี Font จาก Reconcile Result
   * ส่วน Border, Alignment และรูปแบบอื่นยังใช้จาก Summary Template
   
  writeDetails(detailRows: SummaryDetailRow[]): number {
    const locator = this.getConfig().locator;
    const activeFeeIndexes = this.findActiveFeeIndexes(detailRows);
    const maximumActiveFeeIndex = activeFeeIndexes.at(-1) ?? 0;

    this.displayedFeeGroupCount = Math.max(
      locator.templateFeeGroupCount,
      maximumActiveFeeIndex,
    );

    this.unmergeExistingDetailRows();
    this.prepareDynamicFeeColumns(this.displayedFeeGroupCount);
    this.writeFeeHeaders(this.displayedFeeGroupCount);

    const sampleStyles = this.captureTemplateDetailStyles();
    const sampleRowHeight = this.getWorksheet().getRow(locator.detailStartRow).height;

    this.clearDetailValues(detailRows.length);

    detailRows.forEach((detail, index) => {
      const rowNumber = locator.detailStartRow + index;
      const targetRow = this.getWorksheet().getRow(rowNumber);

      this.applyTemplateDetailStyles(rowNumber, sampleStyles);
      targetRow.height = this.calculateDetailRowHeight(
        detail.reconcile.reason,
        detail.testData.fieldValues.testCaseScenario ?? "",
        sampleRowHeight,
      );

      this.writeReconcileValues(rowNumber, detail);
      this.writeTestDataValues(rowNumber, detail);
      this.applyWrappedTextAlignment(rowNumber);
    });

    this.mergeRepeatedTestScriptRows(detailRows);

    return this.displayedFeeGroupCount;
  }

  /**
   * Copy Worksheet ทั้ง Tab จากไฟล์ต้นทางลง Tab ที่มีอยู่ใน Summary Template
   * โดยรักษาตำแหน่ง Row, Column, Style, Merge, Validation และ Filter ตามต้นฉบับ
   
  async replaceWorksheetFromFile(
    sourceFilePath: string,
    sourceSheetName: string,
    targetSheetName: string,
  ): Promise<void> {
    if (!fs.existsSync(sourceFilePath)) {
      throw new Error(`Source workbook not found: ${sourceFilePath}`);
    }

    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(sourceFilePath);

    const sourceWorksheet = sourceWorkbook.getWorksheet(sourceSheetName);
    if (!sourceWorksheet) {
      throw new Error(
        `Source worksheet "${sourceSheetName}" not found in: ${sourceFilePath}`,
      );
    }

    const targetWorksheet = this.workbook.getWorksheet(targetSheetName);
    if (!targetWorksheet) {
      throw new Error(
        `Target worksheet "${targetSheetName}" not found in Summary Template.`,
      );
    }

    [...targetWorksheet.model.merges].forEach((mergeRange) => {
      targetWorksheet.unMergeCells(mergeRange);
    });

    const maximumTargetColumn = targetWorksheet.columnCount;
    for (let rowNumber = 1; rowNumber <= targetWorksheet.rowCount; rowNumber += 1) {
      const targetRow = targetWorksheet.getRow(rowNumber);
      for (let columnNumber = 1; columnNumber <= maximumTargetColumn; columnNumber += 1) {
        const targetCell = targetRow.getCell(columnNumber);
        targetCell.value = null;
        targetCell.style = {};
      }
    }

    const maximumColumnCount = Math.max(
      sourceWorksheet.columnCount,
      targetWorksheet.columnCount,
    );

    for (let columnNumber = 1; columnNumber <= maximumColumnCount; columnNumber += 1) {
      const targetColumn = targetWorksheet.getColumn(columnNumber);

      if (columnNumber > sourceWorksheet.columnCount) {
        targetColumn.width = undefined;
        targetColumn.hidden = false;
        targetColumn.outlineLevel = 0;
        targetColumn.style = {};
        continue;
      }

      const sourceColumn = sourceWorksheet.getColumn(columnNumber);
      targetColumn.width = sourceColumn.width;
      targetColumn.hidden = sourceColumn.hidden;
      targetColumn.outlineLevel = sourceColumn.outlineLevel;
      targetColumn.style = cloneStyle(sourceColumn.style);
    }

    for (let rowNumber = 1; rowNumber <= sourceWorksheet.rowCount; rowNumber += 1) {
      const sourceRow = sourceWorksheet.getRow(rowNumber);
      const targetRow = targetWorksheet.getRow(rowNumber);

      targetRow.height = sourceRow.height;
      targetRow.hidden = sourceRow.hidden;
      targetRow.outlineLevel = sourceRow.outlineLevel;

      for (
        let columnNumber = 1;
        columnNumber <= sourceWorksheet.columnCount;
        columnNumber += 1
      ) {
        const sourceCell = sourceRow.getCell(columnNumber);
        const targetCell = targetRow.getCell(columnNumber);

        targetCell.value = sourceCell.value;
        targetCell.style = cloneStyle(sourceCell.style);

        if (sourceCell.note) {
          targetCell.note = sourceCell.note;
        }
      }
    }

    sourceWorksheet.model.merges.forEach((mergeRange) => {
      targetWorksheet.mergeCells(mergeRange);
    });

    const sourceWithDataValidations =
      sourceWorksheet as WorksheetWithDataValidations;
    const targetWithDataValidations =
      targetWorksheet as WorksheetWithDataValidations;

    targetWithDataValidations.dataValidations.model = JSON.parse(
      JSON.stringify(sourceWithDataValidations.dataValidations.model),
    ) as unknown;

    targetWorksheet.views = JSON.parse(
      JSON.stringify(sourceWorksheet.views ?? []),
    ) as ExcelJS.WorksheetView[];

    targetWorksheet.pageSetup = JSON.parse(
      JSON.stringify(sourceWorksheet.pageSetup),
    ) as Partial<ExcelJS.PageSetup>;

    if (sourceWorksheet.autoFilter) {
      targetWorksheet.autoFilter = JSON.parse(
        JSON.stringify(sourceWorksheet.autoFilter),
      ) as ExcelJS.AutoFilter;
    }

    console.log(
      `Replaced worksheet "${targetSheetName}" from: ${sourceFilePath}`,
    );
  }

  async save(outputFilePath: string): Promise<void> {
    this.getWorksheet();

    ensureDirectoryExists(path.dirname(outputFilePath));
    cleanupStaleTempFiles(path.dirname(outputFilePath));

    const tempFilePath = buildTempFilePath(outputFilePath);
    deleteFileIfExists(tempFilePath);

    try {
      await this.workbook.xlsx.writeFile(tempFilePath);
      deleteFileIfExists(outputFilePath);
      fs.renameSync(tempFilePath, outputFilePath);
    } catch (error) {
      deleteFileIfExists(tempFilePath);
      throw error;
    }
  }

  private writeReconcileValues(
    rowNumber: number,
    detail: SummaryDetailRow,
  ): void {
    const columns = this.getConfig().locator.columns;
    const reconcile = detail.reconcile;

    this.writeText(rowNumber, columns.testResult, reconcile.testResult);
    this.applyTestResultColor(
      rowNumber,
      columns.testResult,
      reconcile.testResultColorStyle,
    );
    this.writeText(rowNumber, columns.reason, reconcile.reason);

    this.getConfig().reconcileFields.forEach(
      (mapping) => {
        this.writeText(
          rowNumber,
          mapping.targetColumn,
          reconcile.fieldValues[mapping.key] ?? "",
        );
      },
    );
  }

  private writeTestDataValues(
    rowNumber: number,
    detail: SummaryDetailRow,
  ): void {
    const locator = this.getConfig().locator;
    const columns = locator.columns;
    const testData = detail.testData;

    this.writeText(
      rowNumber,
      columns.testScriptNo,
      testData.testScriptNo,
    );

    this.getConfig().testDataFields.forEach(
      (mapping) => {
        this.writeText(
          rowNumber,
          mapping.targetColumn,
          testData.fieldValues[mapping.key] ?? "",
        );
      },
    );

    const feeGroupByIndex = new Map(
      testData.feeGroups.map((feeGroup) => [feeGroup.feeIndex, feeGroup]),
    );

    for (let feeIndex = 1; feeIndex <= this.displayedFeeGroupCount; feeIndex += 1) {
      const feeGroup = feeGroupByIndex.get(feeIndex);
      const feeStartColumn =
        columns.firstFeeColumn +
        (feeIndex - 1) * locator.feeColumnsPerGroup;

      this.writeText(rowNumber, feeStartColumn, feeGroup?.feeType ?? "");
      this.writeText(
        rowNumber,
        feeStartColumn + 1,
        feeGroup?.feeChargeAccountNo ?? "",
      );
      this.writeText(rowNumber, feeStartColumn + 2, feeGroup?.feeAmount ?? "");
      this.writeText(rowNumber, feeStartColumn + 3, feeGroup?.feeCurrency ?? "");
    }
  }


  /**
   * ใช้สีของ Test Result จาก Reconcile Result โดย Copy เฉพาะสีเท่านั้น
   * Style ส่วนอื่นของ Cell เช่น Border และ Alignment ยังคงมาจาก Template
   
  private applyTestResultColor(
    rowNumber: number,
    columnNumber: number,
    colorStyle: SummaryTestResultColorStyle,
  ): void {
    const cell = this.getWorksheet().getCell(rowNumber, columnNumber);

    if (colorStyle.fillArgb) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: colorStyle.fillArgb },
      };
    }

    if (colorStyle.fontColorArgb) {
      cell.font = {
        ...(cell.font ?? {}),
        color: { argb: colorStyle.fontColorArgb },
      };
    }
  }

  /**
   * Merge ข้อมูลส่วน Test Script Data เมื่อ Reconcile หลายแถว
   * อ้างถึง Test Data ต้นทางแถวเดียวกัน เช่น DR และ FE ของ Test Case เดียวกัน
   *
   * Merge เฉพาะ Column L ถึง Fee Column สุดท้าย
   * ข้อมูล Reconcile ทางซ้ายยังคงแยกเป็นคนละแถวตามเดิม
   
  private mergeRepeatedTestScriptRows(
    detailRows: SummaryDetailRow[],
  ): void {
    if (detailRows.length < 2) {
      return;
    }

    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;
    let groupStartIndex = 0;

    const mergeGroup = (startIndex: number, endIndex: number): void => {
      if (endIndex <= startIndex) {
        return;
      }

      const startRow = locator.detailStartRow + startIndex;
      const endRow = locator.detailStartRow + endIndex;

      for (
        let columnNumber = locator.columns.testScriptNo;
        columnNumber <= this.dynamicDetailEndColumn;
        columnNumber += 1
      ) {
        worksheet.mergeCells(startRow, columnNumber, endRow, columnNumber);

        const mergedCell = worksheet.getCell(startRow, columnNumber);
        mergedCell.alignment = {
          ...(mergedCell.alignment ?? {}),
          vertical: "middle",
          wrapText:
            columnNumber === locator.columns.testCaseScenario
              ? true
              : mergedCell.alignment?.wrapText,
        };
      }
    };

    for (let index = 1; index <= detailRows.length; index += 1) {
      const isEndOfRows = index === detailRows.length;
      const sameTestDataSource =
        !isEndOfRows &&
        detailRows[index].testData.sourceRowNumber ===
          detailRows[groupStartIndex].testData.sourceRowNumber;

      if (sameTestDataSource) {
        continue;
      }

      mergeGroup(groupStartIndex, index - 1);
      groupStartIndex = index;
    }
  }

  /**
   * ป้องกัน Merge Range เก่าจาก Template หรือไฟล์ที่เคยสร้างไว้
   * โดย Unmerge เฉพาะพื้นที่ Detail ตั้งแต่แถวข้อมูลลงไป
   
  private unmergeExistingDetailRows(): void {
    const worksheet = this.getWorksheet();
    const detailStartRow = this.getConfig().locator.detailStartRow;
    const mergeRanges = [...worksheet.model.merges];

    mergeRanges.forEach((mergeRange) => {
      const match = mergeRange.match(
        /^\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)$/i,
      );

      if (!match) {
        return;
      }

      const startRow = Number(match[1]);
      const endRow = Number(match[2]);

      if (startRow >= detailStartRow || endRow >= detailStartRow) {
        worksheet.unMergeCells(mergeRange);
      }
    });
  }

  private findActiveFeeIndexes(detailRows: SummaryDetailRow[]): number[] {
    const indexes = new Set<number>();

    detailRows.forEach((detail) => {
      detail.testData.feeGroups.forEach((feeGroup) => {
        indexes.add(feeGroup.feeIndex);
      });
    });

    return [...indexes].sort((left, right) => left - right);
  }

  /**
   * ถ้า Report มี Fee มากกว่า Template เดิม 2 Type
   * ให้ต่อ Column ใหม่ทางขวา และ Copy เฉพาะ Style/Width จาก Fee Group ตัวอย่าง
   
  private prepareDynamicFeeColumns(feeGroupCount: number): void {
    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;
    const requiredEndColumn =
      locator.columns.firstFeeColumn +
      feeGroupCount * locator.feeColumnsPerGroup -
      1;

    this.dynamicDetailEndColumn = Math.max(
      locator.fixedDetailEndColumn,
      requiredEndColumn,
    );

    if (requiredEndColumn <= locator.fixedDetailEndColumn) {
      return;
    }

    const sourceFeeGroupStartColumn =
      locator.columns.firstFeeColumn +
      (locator.templateFeeGroupCount - 1) * locator.feeColumnsPerGroup;

    for (
      let targetColumn = locator.fixedDetailEndColumn + 1;
      targetColumn <= requiredEndColumn;
      targetColumn += 1
    ) {
      const offsetWithinGroup =
        (targetColumn - locator.columns.firstFeeColumn) %
        locator.feeColumnsPerGroup;
      const sourceColumn = sourceFeeGroupStartColumn + offsetWithinGroup;

      worksheet.getColumn(targetColumn).width = worksheet.getColumn(sourceColumn).width;
      worksheet.getCell(locator.detailHeaderRow, targetColumn).style = cloneStyle(
        worksheet.getCell(locator.detailHeaderRow, sourceColumn).style,
      );
      worksheet.getCell(locator.detailStartRow, targetColumn).style = cloneStyle(
        worksheet.getCell(locator.detailStartRow, sourceColumn).style,
      );
    }

    this.extendTestScriptDataMerge(requiredEndColumn);
  }

  private extendTestScriptDataMerge(requiredEndColumn: number): void {
    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;
    const titleCell = worksheet.getCell(locator.testScriptDataTitle);
    const titleValue = titleCell.value;
    const titleStyle = cloneStyle(titleCell.style);

    const existingMerge = worksheet.model.merges.find((mergeRange) =>
      mergeRange.startsWith(`${titleCell.address}:`),
    );

    if (existingMerge) {
      worksheet.unMergeCells(existingMerge);
    }

    const titleRowNumber = titleCell.fullAddress.row;
    const titleColumnNumber = titleCell.fullAddress.col;

    worksheet.mergeCells(
      titleRowNumber,
      titleColumnNumber,
      titleRowNumber,
      requiredEndColumn,
    );

    const mergedTitleCell = worksheet.getCell(
      titleRowNumber,
      titleColumnNumber,
    );
    mergedTitleCell.value = titleValue;
    mergedTitleCell.style = titleStyle;
  }

  private writeFeeHeaders(feeGroupCount: number): void {
    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;

    for (let feeIndex = 1; feeIndex <= feeGroupCount; feeIndex += 1) {
      const startColumn =
        locator.columns.firstFeeColumn +
        (feeIndex - 1) * locator.feeColumnsPerGroup;

      worksheet.getCell(locator.detailHeaderRow, startColumn).value =
        `Fee Type ${feeIndex}`;
      worksheet.getCell(locator.detailHeaderRow, startColumn + 1).value =
        `Fee Charge Account No. Type ${feeIndex}`;
      worksheet.getCell(locator.detailHeaderRow, startColumn + 2).value =
        `Fee Amount Type ${feeIndex}`;
      worksheet.getCell(locator.detailHeaderRow, startColumn + 3).value =
        `Fee Currency ${feeIndex}`;
    }
  }

  private clearDetailValues(requiredRowCount: number): void {
    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;
    const lastRequiredRow =
      locator.detailStartRow + Math.max(requiredRowCount, 1) - 1;
    const lastRowNumber = Math.max(worksheet.rowCount, lastRequiredRow);

    for (
      let rowNumber = locator.detailStartRow;
      rowNumber <= lastRowNumber;
      rowNumber += 1
    ) {
      for (
        let columnNumber = locator.detailStartColumn;
        columnNumber <= this.dynamicDetailEndColumn;
        columnNumber += 1
      ) {
        worksheet.getCell(rowNumber, columnNumber).value = null;
      }
    }
  }

  private captureTemplateDetailStyles(): Map<number, Partial<ExcelJS.Style>> {
    const worksheet = this.getWorksheet();
    const locator = this.getConfig().locator;
    const styles = new Map<number, Partial<ExcelJS.Style>>();

    for (
      let columnNumber = locator.detailStartColumn;
      columnNumber <= this.dynamicDetailEndColumn;
      columnNumber += 1
    ) {
      styles.set(
        columnNumber,
        cloneStyle(worksheet.getCell(locator.detailStartRow, columnNumber).style),
      );
    }

    return styles;
  }

  private applyTemplateDetailStyles(
    rowNumber: number,
    styles: Map<number, Partial<ExcelJS.Style>>,
  ): void {
    const worksheet = this.getWorksheet();

    styles.forEach((style, columnNumber) => {
      worksheet.getCell(rowNumber, columnNumber).style = cloneStyle(style);
    });
  }

  private applyWrappedTextAlignment(rowNumber: number): void {
    const worksheet = this.getWorksheet();
    const columns = this.getConfig().locator.columns;

    [columns.reason, columns.testCaseScenario].forEach((columnNumber) => {
      const cell = worksheet.getCell(rowNumber, columnNumber);
      cell.alignment = {
        ...(cell.alignment ?? {}),
        wrapText: true,
        vertical: "top",
      };
    });
  }

  private writeText(rowNumber: number, columnNumber: number, value: string): void {
    const cell = this.getWorksheet().getCell(rowNumber, columnNumber);

    // ใช้ null สำหรับค่าว่าง ป้องกัน Excel Reader บางตัวแสดง Shared String Index
    cell.value = value.trim() === "" ? null : value;
  }

  private calculateDetailRowHeight(
    reason: string,
    testCaseScenario: string,
    templateRowHeight: number | undefined,
  ): number {
    const minimumHeight = templateRowHeight ?? 18;
    const estimatedLines = Math.max(
      this.estimateWrappedLineCount(reason, 85),
      this.estimateWrappedLineCount(testCaseScenario, 50),
    );

    if (estimatedLines <= 1) {
      return minimumHeight;
    }

    return Math.max(minimumHeight, estimatedLines * 15 + 5);
  }

  private estimateWrappedLineCount(value: string, charactersPerLine: number): number {
    if (value.trim() === "") {
      return 1;
    }

    return value.split(/\r?\n/).reduce((total, line) => {
      return total + Math.max(1, Math.ceil(line.length / charactersPerLine));
    }, 0);
  }

  private writeCountCell(address: string, value: number): void {
    const cell = this.getWorksheet().getCell(address);

    cell.value = value;
    cell.numFmt = "0";
    cell.alignment = {
      ...(cell.alignment ?? {}),
      horizontal: "center",
      vertical: "middle",
    };
    cell.font = {
      ...(cell.font ?? {}),
      bold: true,
      size: 18,
    };
  }

  private getWorksheet(): ExcelJS.Worksheet {
    if (!this.worksheet) {
      throw new Error("Summary Template has not been opened. Call openTemplate() first.");
    }

    return this.worksheet;
  }

  private getConfig(): SummaryReportConfig {
    if (!this.config) {
      throw new Error("Summary Config is not available. Call openTemplate() first.");
    }

    return this.config;
  }
} */
