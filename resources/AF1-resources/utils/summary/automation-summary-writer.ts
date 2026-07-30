/**
 * ======================================================
 * ไฟล์: automation-summary-writer.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * อ่านข้อมูลจาก Compare Result, Original Test Data, Checked Report และ Checked Test Data
 * จากนั้นเติมข้อมูลลง Template และสร้างไฟล์ Automation Summary ของ Report ที่เลือกสำหรับ Script 4
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import {
  AutomationSummaryInfo,
  CompareResultRow,
  SummaryStatus,
} from "./summary-types";

const COLORS = {
  PASS_FILL: "FFC6EFCE",
  PASS_TEXT: "FF006100",
  FAIL_FILL: "FFFFC7CE",
  FAIL_TEXT: "FF9C0006",
  SKIP_FILL: "FFFFE699",
  SKIP_TEXT: "FF7F6000",
  BORDER: "FFB7B7B7",
};

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: "thin", color: { argb: COLORS.BORDER } },
  left: { style: "thin", color: { argb: COLORS.BORDER } },
  bottom: { style: "thin", color: { argb: COLORS.BORDER } },
  right: { style: "thin", color: { argb: COLORS.BORDER } },
};

type DataRecord = Record<string, unknown>;

type SupportedSummaryReport =
  | "DS_PTX"
  | "DS_FTX";

type SummaryReportConfig = {
  reportCode: SupportedSummaryReport;
  summarySheetName: string;
  reconcileSheetName: string;
  reportSheetName: string;
  title: string;
  hasDynamicFeeColumns: boolean;
  compareLastColumn: number;
  testDataFirstColumn: number;
};

const SUMMARY_REPORT_CONFIG: Record<
  SupportedSummaryReport,
  SummaryReportConfig
> = {
  DS_PTX: {
    reportCode: "DS_PTX",
    summarySheetName:
      "DS-PTX Summary Test Results",
    reconcileSheetName:
      "DS_PTX_Reconcile",
    reportSheetName:
      "DS_PTX",
    title:
      "DS_PTX AUTOMATION VERIFICATION SUMMARY",
    hasDynamicFeeColumns: true,
    compareLastColumn: 11,
    testDataFirstColumn: 13,
  },
  DS_FTX: {
    reportCode: "DS_FTX",
    summarySheetName:
      "DS-FTX Summary Test Results",
    reconcileSheetName:
      "DS_FTX_Reconcile",
    reportSheetName:
      "DS_FTX",
    title:
      "DS_FTX AUTOMATION VERIFICATION SUMMARY",
    hasDynamicFeeColumns: false,
    compareLastColumn: 11,
    testDataFirstColumn: 13,
  },
};

const normalizeText = (value: unknown): string =>
  String(value ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\r?\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const normalizeHeader = (value: unknown): string =>
  normalizeText(value).toLowerCase();

const normalizeReportName = (
  reportName: string,
): SupportedSummaryReport => {
  const normalizedReportName =
    normalizeText(reportName)
      .toUpperCase()
      .replace(/-/g, "_");

  if (
    normalizedReportName === "DS_PTX" ||
    normalizedReportName === "DS_FTX"
  ) {
    return normalizedReportName;
  }

  throw new Error(
    `Unsupported Summary Report: ${reportName}`,
  );
};

const getSummaryReportConfig = (
  reportName: string,
): SummaryReportConfig =>
  SUMMARY_REPORT_CONFIG[
    normalizeReportName(
      reportName,
    )
  ];

const normalizeStatus = (value: unknown): SummaryStatus => {
  const status = normalizeText(value).toUpperCase();

  if (status === "PASS" || status === "FAIL" || status === "SKIP") {
    return status;
  }

  throw new Error(`Unsupported compare status: ${normalizeText(value)}`);
};

const getCellValue = (cell: ExcelJS.Cell): unknown => {
  const value = cell.value;

  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    const dd = String(value.getDate()).padStart(2, "0");
    const MM = String(value.getMonth() + 1).padStart(2, "0");
    const yyyy = value.getFullYear();

    return `${dd}/${MM}/${yyyy}`;
  }

  if (typeof value === "object") {
    if ("result" in value) {
      return value.result ?? "";
    }

    if ("text" in value) {
      return value.text ?? "";
    }

    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText
        .map((item) => item.text)
        .join("");
    }
  }

  return value;
};

const buildHeaderMap = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber: number,
): Map<string, number> => {
  const map = new Map<string, number>();

  worksheet
    .getRow(headerRowNumber)
    .eachCell({ includeEmpty: true }, (cell, columnNumber) => {
      const normalized = normalizeHeader(getCellValue(cell));

      if (normalized !== "") {
        map.set(normalized, columnNumber);
      }
    });

  return map;
};

const findCompareHeaderRowNumber = (
  worksheet: ExcelJS.Worksheet,
): number => {
  const maxRowToCheck = Math.min(20, worksheet.rowCount);

  for (let rowNumber = 1; rowNumber <= maxRowToCheck; rowNumber += 1) {
    const headerMap = buildHeaderMap(worksheet, rowNumber);

    if (
      headerMap.has("test script no.") &&
      headerMap.has("result") &&
      headerMap.has("remark")
    ) {
      return rowNumber;
    }
  }

  throw new Error(
    'Compare result header not found. Required headers: "Test Script No.", "Result", "Remark"',
  );
};

const findTestDataHeaderRowNumber = (
  worksheet: ExcelJS.Worksheet,
): number => {
  const maxRowToCheck = Math.min(20, worksheet.rowCount);

  for (let rowNumber = 1; rowNumber <= maxRowToCheck; rowNumber += 1) {
    const headerMap = buildHeaderMap(worksheet, rowNumber);

    const hasTestNo =
      headerMap.has("test no.") ||
      headerMap.has("test no") ||
      headerMap.has("test script no.") ||
      headerMap.has("test script no");

    const hasTransactionId =
      headerMap.has("transaction id/ reconcile id") ||
      headerMap.has("transaction id / reconcile id");

    if (hasTestNo && hasTransactionId) {
      return rowNumber;
    }
  }

  throw new Error(
    `Test Data header row not found in worksheet: ${worksheet.name}`,
  );
};

const getRequiredColumn = (
  headerMap: Map<string, number>,
  possibleHeaders: string[],
): number => {
  for (const header of possibleHeaders) {
    const columnNumber = headerMap.get(normalizeHeader(header));

    if (columnNumber) {
      return columnNumber;
    }
  }

  throw new Error(
    `Required header not found: ${possibleHeaders.join(" / ")}`,
  );
};

const getOptionalColumn = (
  headerMap: Map<string, number>,
  possibleHeaders: string[],
): number | undefined => {
  for (const header of possibleHeaders) {
    const columnNumber =
      headerMap.get(
        normalizeHeader(
          header,
        ),
      );

    if (columnNumber) {
      return columnNumber;
    }
  }

  return undefined;
};

const worksheetRowToRecord = (
  worksheet: ExcelJS.Worksheet,
  rowNumber: number,
  headerRowNumber: number,
): DataRecord => {
  const record: DataRecord = {};
  const headerRow = worksheet.getRow(headerRowNumber);
  const dataRow = worksheet.getRow(rowNumber);

  headerRow.eachCell(
    { includeEmpty: true },
    (headerCell, columnNumber) => {
      const header = normalizeText(getCellValue(headerCell));

      if (header !== "") {
        record[normalizeHeader(header)] = getCellValue(
          dataRow.getCell(columnNumber),
        );
      }
    },
  );

  return record;
};

const getRecordValue = (
  record: DataRecord | undefined,
  possibleHeaders: string[],
): unknown => {
  if (!record) {
    return "";
  }

  for (const header of possibleHeaders) {
    const key = normalizeHeader(header);

    if (Object.prototype.hasOwnProperty.call(record, key)) {
      return record[key] ?? "";
    }
  }

  return "";
};

export const readCompareResultRows = async (
  compareResultPath: string,
): Promise<CompareResultRow[]> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(compareResultPath);

  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error(`Worksheet not found: ${compareResultPath}`);
  }

  const headerRowNumber = findCompareHeaderRowNumber(worksheet);
  const headerMap = buildHeaderMap(worksheet, headerRowNumber);

  const testScriptColumn = getRequiredColumn(headerMap, [
    "Test Script No.",
  ]);
  const resultColumn = getRequiredColumn(headerMap, ["Result"]);
  const remarkColumn = getRequiredColumn(headerMap, ["Remark"]);
  const matchingKeyColumn =
    getOptionalColumn(
      headerMap,
      [
        "Reference Transaction Number",
        "Ref. TX No.",
        "Ref TX No",
        "Reference TX No.",
        "Matching Key",
      ],
    );

  const rows: CompareResultRow[] = [];

  for (
    let rowNumber = headerRowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const row = worksheet.getRow(rowNumber);
    const testScriptNo = normalizeText(
      getCellValue(row.getCell(testScriptColumn)),
    );
    const resultValue = normalizeText(
      getCellValue(row.getCell(resultColumn)),
    );

    if (testScriptNo === "" && resultValue === "") {
      continue;
    }

    if (testScriptNo === "") {
      continue;
    }

    const reportValues = worksheetRowToRecord(
      worksheet,
      rowNumber,
      headerRowNumber,
    );

    rows.push({
      testScriptNo,
      matchingKey: matchingKeyColumn
        ? normalizeText(getCellValue(row.getCell(matchingKeyColumn)))
        : "",
      status: normalizeStatus(resultValue),
      remark: normalizeText(getCellValue(row.getCell(remarkColumn))),
      reportValues,
    });
  }

  if (rows.length === 0) {
    throw new Error(`No compare result data found: ${compareResultPath}`);
  }

  return rows;
};

type CheckedTestDataResult = {
  testDataMap: Map<string, DataRecord>;
  maxFeeIndex: number;
};

/**
 * อ่านเลขลำดับ Fee จาก Header
 *
 * รองรับตัวอย่างต่อไปนี้:
 * - Fee Type 1
 * - Fee Charge Account No. Type 1
 * - Fee Charge Account Type 1
 * - Fee Currency 1
 * - Fee Currency Type 1
 * - Fee Amount Type 1
 * - Fee Amount 2
 * - Fee Amount 3
 */
const getFeeIndexFromHeader = (
  header: unknown,
): number | undefined => {
  const normalized = normalizeHeader(header);

  const patterns = [
    /^fee type\s+(\d+)$/,
    /^fee charge account no\. type\s+(\d+)$/,
    /^charge account no\. type\s+(\d+)$/,
    /^fee charge account type\s+(\d+)$/,
    /^charge account type\s+(\d+)$/,
    /^fee currency(?: type)?\s+(\d+)$/,
    /^fee amount(?: type)?\s+(\d+)$/,
  ];

  for (const pattern of patterns) {
    const matched = normalized.match(pattern);

    if (matched) {
      return Number(matched[1]);
    }
  }

  return undefined;
};

const getMaxFeeIndexFromHeaderMap = (
  headerMap: Map<string, number>,
): number => {
  let maxFeeIndex = 0;

  for (const header of headerMap.keys()) {
    const feeIndex = getFeeIndexFromHeader(header);

    if (feeIndex && feeIndex > maxFeeIndex) {
      maxFeeIndex = feeIndex;
    }
  }

  return maxFeeIndex;
};

/**
 * ตรวจสอบว่า Fee ลำดับที่กำหนดมีข้อมูลจริงหรือไม่
 *
 * ตัวอย่าง:
 * Fee 3 จะตรวจ:
 * - Fee Type 3
 * - Fee Charge Account No. Type 3
 * - Fee Charge Account Type 3
 * - Fee Currency 3 / Fee Currency Type 3
 * - Fee Amount 3 / Fee Amount Type 3
 */
const hasFeeDataAtIndex = (
  record: DataRecord,
  feeIndex: number,
): boolean => {
  const possibleHeaders = [
    `Fee Type ${feeIndex}`,

    `Fee Charge Account No. Type ${feeIndex}`,
    `Charge Account No. Type ${feeIndex}`,

    `Fee Charge Account Type ${feeIndex}`,
    `Charge Account Type ${feeIndex}`,

    `Fee Currency ${feeIndex}`,
    `Fee Currency Type ${feeIndex}`,

    feeIndex === 1
      ? "Fee Amount Type 1"
      : `Fee Amount ${feeIndex}`,

    `Fee Amount Type ${feeIndex}`,
  ];

  return possibleHeaders.some((header) => {
    const value = getRecordValue(
      record,
      [header],
    );

    return normalizeText(value) !== "";
  });
};

/**
 * หา Fee สูงสุดจากข้อมูลจริง
 *
 * จะไม่นับ Fee ที่มีแค่ Header แต่ข้อมูลทุกแถวว่าง
 */
const getMaxFeeIndexFromActualData = (
  records: DataRecord[],
  maxFeeIndexFromHeader: number,
): number => {
  let maxUsedFeeIndex = 0;

  for (
    let feeIndex = 1;
    feeIndex <= maxFeeIndexFromHeader;
    feeIndex += 1
  ) {
    const hasActualData = records.some(
      (record) =>
        hasFeeDataAtIndex(
          record,
          feeIndex,
        ),
    );

    if (hasActualData) {
      maxUsedFeeIndex = feeIndex;
    }
  }

  return maxUsedFeeIndex;
};

const readCheckedTestDataRows = async (
  originalTestDataPath: string,
): Promise<CheckedTestDataResult> => {
  const workbook =
    new ExcelJS.Workbook();

  await workbook.xlsx.readFile(
    originalTestDataPath,
  );

  const worksheet =
    workbook.getWorksheet("Test Data") ??
    workbook.worksheets.find(
      (sheet) =>
        normalizeHeader(sheet.name) !==
        "field validation",
    ) ??
    workbook.worksheets[0];

  if (!worksheet) {
    throw new Error(
      `Test Data worksheet not found: ${originalTestDataPath}`,
    );
  }

  const headerRowNumber =
    findTestDataHeaderRowNumber(
      worksheet,
    );

  const headerMap =
    buildHeaderMap(
      worksheet,
      headerRowNumber,
    );

  const testNoColumn =
    getRequiredColumn(
      headerMap,
      [
        "Test No.",
        "Test No",
        "Test Script No.",
        "Test Script No",
      ],
    );

  const testDataMap =
    new Map<string, DataRecord>();

  /*
   * เก็บ Record ทั้งหมดไว้ใช้ตรวจว่า
   * Fee แต่ละลำดับมีข้อมูลจริงหรือไม่
   */
  const allTestDataRecords: DataRecord[] =
    [];

  for (
    let rowNumber =
      headerRowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const testScriptNo =
      normalizeText(
        getCellValue(
          worksheet
            .getRow(rowNumber)
            .getCell(testNoColumn),
        ),
      );

    if (testScriptNo === "") {
      continue;
    }

    const record =
      worksheetRowToRecord(
        worksheet,
        rowNumber,
        headerRowNumber,
      );

    allTestDataRecords.push(record);

    if (
      !testDataMap.has(
        testScriptNo,
      )
    ) {
      testDataMap.set(
        testScriptNo,
        record,
      );
    }
  }

  if (testDataMap.size === 0) {
    throw new Error(
      `No Test Data records found: ${originalTestDataPath}`,
    );
  }

  /*
   * จำนวน Fee สูงสุดที่มีอยู่ใน Header
   *
   * เช่น Original Test Data มี Header ถึง Fee 5
   */
  const maxFeeIndexFromHeader =
    getMaxFeeIndexFromHeaderMap(
      headerMap,
    );

  /*
   * จำนวน Fee สูงสุดที่มีข้อมูลจริง
   *
   * เช่น Header มีถึง Fee 5
   * แต่มีข้อมูลจริงถึง Fee 2
   * ผลลัพธ์จะเป็น 2
   */
  const maxFeeIndexFromActualData =
    getMaxFeeIndexFromActualData(
      allTestDataRecords,
      maxFeeIndexFromHeader,
    );

  console.log(
    "Maximum Fee Header Index :",
    maxFeeIndexFromHeader,
  );

  console.log(
    "Maximum Fee Used Index   :",
    maxFeeIndexFromActualData,
  );

  return {
    testDataMap,
    maxFeeIndex:
      maxFeeIndexFromActualData,
  };
};

const writeSummaryInformation = (
  summarySheet: ExcelJS.Worksheet,
  info: AutomationSummaryInfo,
  config: SummaryReportConfig,
): void => {
  summarySheet.getCell("B2").value =
    config.title;

  summarySheet.getCell("C5").value = info.reportFileName;
  summarySheet.getCell("C6").value = info.executionDate;
  summarySheet.getCell("C7").value = info.executionTime;
  summarySheet.getCell("C8").value = info.runId;
  summarySheet.getCell("C9").value = info.verifiedBy;

  summarySheet.getCell("J7").value = info.totalChecked;
  summarySheet.getCell("K7").value = info.passed;
  summarySheet.getCell("L7").value = info.failed;
};

const applyStatusStyle = (
  cell: ExcelJS.Cell,
  status: SummaryStatus,
): void => {
  const fill =
    status === "PASS"
      ? COLORS.PASS_FILL
      : status === "FAIL"
        ? COLORS.FAIL_FILL
        : COLORS.SKIP_FILL;

  const text =
    status === "PASS"
      ? COLORS.PASS_TEXT
      : status === "FAIL"
        ? COLORS.FAIL_TEXT
        : COLORS.SKIP_TEXT;

  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: fill },
  };

  cell.font = {
    bold: true,
    color: { argb: text },
  };
};

const getCompareValue = (
  header: string,
  compareRow: CompareResultRow,
): unknown => {
  const normalized = normalizeHeader(header);

  if (normalized === "test result") {
    return compareRow.status;
  }

  if (normalized === "reason") {
    if (compareRow.remark !== "") {
      return compareRow.remark;
    }

    if (compareRow.status === "PASS") {
      return "Validation passed.";
    }

    if (compareRow.status === "SKIP") {
      return "Validation skipped.";
    }

    return "Validation failed.";
  }

  if (
    normalized ===
      "reference transaction number" ||
    normalized ===
      "ref. tx no." ||
    normalized ===
      "ref tx no" ||
    normalized ===
      "reference tx no."
  ) {
    if (
      compareRow.matchingKey !== ""
    ) {
      return compareRow.matchingKey;
    }

    return getRecordValue(
      compareRow.reportValues,
      [
        "Reference Transaction Number",
        "Ref. TX No.",
        "Ref TX No",
        "Reference TX No.",
        "Matching Key",
      ],
    );
  }

  return getRecordValue(compareRow.reportValues, [header]);
};

const TEST_DATA_HEADER_ALIASES: Record<string, string[]> = {
  "test script no.": [
    "Test No.",
    "Test No",
    "Test Script No.",
    "Test Script No",
  ],
  "test case / scenario": [
    "Test Case / Scenario",
    "Test Case/Scenario",
    "Test Scenario",
  ],
  "txn date": ["Txn Date", "Transaction Date"],
  "reference transaction number": [
    "Transaction ID/ Reconcile ID",
    "Transaction ID / Reconcile ID",
    "Reference Transaction Number",
  ],
  "from cif no. (client/sender)": [
    "From CIF No. (Client/Sender)",
  ],
  "from account ( a/c client/sender)": [
    "From Account ( A/C Client/Sender)",
    "From Account (A/C Client/Sender)",
  ],
  "from account (a/c client/sender)": [
    "From Account ( A/C Client/Sender)",
    "From Account (A/C Client/Sender)",
  ],
  "from currency (ccy)": ["From Currency (CCY)"],
  "from debit amount": ["From Debit Amount", "From Debit Amount "],
  "from transfer amount": ["From Transfer Amount"],
  "from customer (resident/non resident)": [
    "From Customer (Resident/Non Resident)",
    "From Customer (Resident / Non Resident)",
  ],
  "settled amount (ccy)": [
    "Settled Amount (CCY)",
    "Settled Amount",
  ],
  "settled currency (ccy)": [
    "Settled Currency (CCY)",
    "Settled Currency",
  ],
};

/**
 * สร้าง Alias ของ Fee Header แบบ Dynamic
 * รองรับ Fee 1, 2, 3 ... n
 *
 * Fee Amount มีรูปแบบพิเศษใน Test Data:
 * - Fee Amount Type 1
 * - Fee Amount 2
 * - Fee Amount 3
 * - Fee Amount n
 */
const getDynamicFeeHeaderAliases = (
  header: string,
): string[] | undefined => {
  const normalized = normalizeHeader(header);

  let matched = normalized.match(/^fee type\s+(\d+)$/);

  if (matched) {
    return [`Fee Type ${matched[1]}`];
  }

  matched = normalized.match(
    /^fee charge account no\. type\s+(\d+)$/,
  );

  if (matched) {
    return [
      `Fee Charge Account No. Type ${matched[1]}`,
      `Charge Account No. Type ${matched[1]}`,
    ];
  }

  matched = normalized.match(
    /^fee charge account type\s+(\d+)$/,
  );

  if (matched) {
    return [
      `Fee Charge Account Type ${matched[1]}`,
      `Charge Account Type ${matched[1]}`,
    ];
  }

  matched = normalized.match(/^fee currency(?: type)?\s+(\d+)$/);

  if (matched) {
    return [
      `Fee Currency ${matched[1]}`,
      `Fee Currency Type ${matched[1]}`,
    ];
  }

  matched = normalized.match(/^fee amount(?: type)?\s+(\d+)$/);

  if (matched) {
    const index = Number(matched[1]);

    if (index === 1) {
      return [
        "Fee Amount Type 1",
        "Fee Amount 1",
      ];
    }

    return [
      `Fee Amount ${index}`,
      `Fee Amount Type ${index}`,
    ];
  }

  return undefined;
};

const getTestDataValue = (
  header: string,
  testDataRow: DataRecord | undefined,
): unknown => {
  const normalized = normalizeHeader(header);
  const dynamicFeeAliases = getDynamicFeeHeaderAliases(header);
  const aliases =
    dynamicFeeAliases ??
    TEST_DATA_HEADER_ALIASES[normalized] ??
    [header];

  return getRecordValue(testDataRow, aliases);
};

const getFeeHeaderNames = (
  feeIndex: number,
): string[] => [
  `Fee Type ${feeIndex}`,
  `Fee Charge Account No. Type ${feeIndex}`,
  `Fee Charge Account Type ${feeIndex}`,
  `Fee Currency ${feeIndex}`,
  feeIndex === 1
    ? "Fee Amount Type 1"
    : `Fee Amount ${feeIndex}`,
];

const cloneCellStyle = (
  sourceCell: ExcelJS.Cell,
  targetCell: ExcelJS.Cell,
): void => {
  /*
   * Copy style รวมจาก Source Cell
   *
   * ใช้ JSON clone เพื่อไม่ให้ Target Cell อ้างอิง Object
   * ตัวเดียวกับ Source Cell
   */
  targetCell.style = JSON.parse(
    JSON.stringify(sourceCell.style ?? {}),
  ) as Partial<ExcelJS.Style>;

  /*
   * Copy Number Format
   */
  targetCell.numFmt = sourceCell.numFmt;

  /*
   * ExcelJS ไม่ยอมให้กำหนด undefined ให้ alignment โดยตรง
   * จึงกำหนดเฉพาะกรณีที่ Source มีค่า
   */
  if (sourceCell.alignment) {
    targetCell.alignment = {
      ...sourceCell.alignment,
    };
  }

  /*
   * Copy Border
   */
  if (sourceCell.border) {
    targetCell.border = JSON.parse(
      JSON.stringify(sourceCell.border),
    ) as Partial<ExcelJS.Borders>;
  }

  /*
   * Copy Fill
   */
  if (sourceCell.fill) {
    targetCell.fill = JSON.parse(
      JSON.stringify(sourceCell.fill),
    ) as ExcelJS.Fill;
  }

  /*
   * Copy Font
   */
  if (sourceCell.font) {
    targetCell.font = JSON.parse(
      JSON.stringify(sourceCell.font),
    ) as Partial<ExcelJS.Font>;
  }

  /*
   * ExcelJS ไม่ยอมให้กำหนด undefined ให้ protection โดยตรง
   * จึงกำหนดเฉพาะกรณีที่ Source มีค่า
   */
  if (sourceCell.protection) {
    targetCell.protection = {
      ...sourceCell.protection,
    };
  }
};

/**
 * เลือก Worksheet ต้นทางจากชื่อที่ต้องการก่อน
 * หากไม่พบ จะใช้ Worksheet แรกของไฟล์
 */
const findSourceWorksheet = (
  workbook: ExcelJS.Workbook,
  preferredSheetNames: string[],
): ExcelJS.Worksheet => {
  for (const sheetName of preferredSheetNames) {
    /**
     * ค้นหาชื่อแบบตรงตัวก่อน
     */
    const exactWorksheet =
      workbook.getWorksheet(
        sheetName,
      );

    if (exactWorksheet) {
      return exactWorksheet;
    }

    /**
     * ถ้าไม่พบ ให้ค้นหาแบบตัดช่องว่าง
     * และไม่สนใจตัวพิมพ์เล็ก/ใหญ่
     *
     * ตัวอย่าง:
     * "DS_FTX " จะถือว่าเท่ากับ "DS_FTX"
     */
    const normalizedWorksheet =
      workbook.worksheets.find(
        (worksheet) =>
          normalizeHeader(
            worksheet.name,
          ) ===
          normalizeHeader(
            sheetName,
          ),
      );

    if (normalizedWorksheet) {
      return normalizedWorksheet;
    }
  }

  const firstWorksheet = workbook.worksheets[0];

  if (!firstWorksheet) {
    throw new Error("Source workbook does not contain any worksheet.");
  }

  return firstWorksheet;
};

/**
 * หา Worksheet ปลายทางใน Template
 *
 * รองรับกรณีชื่อชีตใน Template
 * มีช่องว่างซ่อนอยู่ด้านหน้าหรือด้านหลัง
 */
const findTargetWorksheet = (
  workbook: ExcelJS.Workbook,
  targetSheetName: string,
): ExcelJS.Worksheet | undefined => {
  const worksheet =
    workbook.getWorksheet(
      targetSheetName,
    ) ??
    workbook.worksheets.find(
      (candidateWorksheet) =>
        normalizeHeader(
          candidateWorksheet.name,
        ) ===
        normalizeHeader(
          targetSheetName,
        ),
    );

  if (
    worksheet &&
    worksheet.name !== targetSheetName
  ) {
    /**
     * ปรับชื่อชีตในไฟล์ผลลัพธ์ให้ตรงกับชื่อมาตรฐาน
     *
     * ตัวอย่าง:
     * "DS_FTX " -> "DS_FTX"
     */
    worksheet.name =
      targetSheetName;
  }

  return worksheet;
};

/** อ่านรายการ Merge Cell จาก Worksheet */
const getMergedCellRanges = (
  worksheet: ExcelJS.Worksheet,
): string[] => {
  const worksheetModel = worksheet.model as ExcelJS.WorksheetModel & {
    merges?: string[];
  };

  return [...(worksheetModel.merges ?? [])];
};

/** ล้างข้อมูลเดิมของ Worksheet ปลายทาง */
const clearTargetWorksheet = (
  worksheet: ExcelJS.Worksheet,
): void => {
  for (const mergedCellRange of getMergedCellRanges(worksheet)) {
    worksheet.unMergeCells(mergedCellRange);
  }

  if (worksheet.rowCount > 0) {
    worksheet.spliceRows(1, worksheet.rowCount);
  }
};

/** Copy ค่าและรูปแบบ Cell จาก Worksheet ต้นทางไปปลายทาง */
const copyWorksheetCells = (
  sourceWorksheet: ExcelJS.Worksheet,
  targetWorksheet: ExcelJS.Worksheet,
): void => {
  const maximumColumnCount = sourceWorksheet.columnCount;

  for (
    let columnNumber = 1;
    columnNumber <= maximumColumnCount;
    columnNumber += 1
  ) {
    const sourceColumn = sourceWorksheet.getColumn(columnNumber);
    const targetColumn = targetWorksheet.getColumn(columnNumber);

    targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.outlineLevel = sourceColumn.outlineLevel;
  }

  for (
    let rowNumber = 1;
    rowNumber <= sourceWorksheet.rowCount;
    rowNumber += 1
  ) {
    const sourceRow = sourceWorksheet.getRow(rowNumber);
    const targetRow = targetWorksheet.getRow(rowNumber);

    targetRow.height = sourceRow.height;
    targetRow.hidden = sourceRow.hidden;
    targetRow.outlineLevel = sourceRow.outlineLevel;

    for (
      let columnNumber = 1;
      columnNumber <= maximumColumnCount;
      columnNumber += 1
    ) {
      const sourceCell = sourceRow.getCell(columnNumber);
      const targetCell = targetRow.getCell(columnNumber);

      targetCell.value = sourceCell.value;
      cloneCellStyle(sourceCell, targetCell);

      if (sourceCell.note) {
        targetCell.note = sourceCell.note;
      }
    }
  }
};

/**
 * คืน Style จากต้นทางอีกครั้งหลัง Merge
 * เพราะ ExcelJS อาจเปลี่ยน Border ของ Cell ภายในช่วง Merge
 */
const restoreWorksheetStylesAfterMerge = (
  sourceWorksheet: ExcelJS.Worksheet,
  targetWorksheet: ExcelJS.Worksheet,
): void => {
  for (
    let rowNumber = 1;
    rowNumber <= sourceWorksheet.rowCount;
    rowNumber += 1
  ) {
    for (
      let columnNumber = 1;
      columnNumber <= sourceWorksheet.columnCount;
      columnNumber += 1
    ) {
      cloneCellStyle(
        sourceWorksheet.getRow(rowNumber).getCell(columnNumber),
        targetWorksheet.getRow(rowNumber).getCell(columnNumber),
      );
    }
  }
};

/**
 * Copy Worksheet จากไฟล์ Excel ต้นทางลงใน Template
 * โดยรักษาค่า Cell, Highlight, Style, Merge, Row Height และ Column Width
 */
const copyWorksheetFromFile = async (
  sourceFilePath: string,
  targetWorkbook: ExcelJS.Workbook,
  targetSheetName: string,
  preferredSourceSheetNames: string[],
): Promise<void> => {
  if (!fs.existsSync(sourceFilePath)) {
    throw new Error(`Worksheet source file not found: ${sourceFilePath}`);
  }

  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(sourceFilePath);

  const sourceWorksheet = findSourceWorksheet(
    sourceWorkbook,
    preferredSourceSheetNames,
  );

  const targetWorksheet =
    findTargetWorksheet(
      targetWorkbook,
      targetSheetName,
    );

  if (!targetWorksheet) {
    throw new Error(
      `Target worksheet "${targetSheetName}" not found in Summary template.`,
    );
  }

  clearTargetWorksheet(targetWorksheet);
  copyWorksheetCells(sourceWorksheet, targetWorksheet);

  for (const mergedCellRange of getMergedCellRanges(sourceWorksheet)) {
    targetWorksheet.mergeCells(mergedCellRange);
  }

  restoreWorksheetStylesAfterMerge(
    sourceWorksheet,
    targetWorksheet,
  );

  targetWorksheet.properties = {
    ...sourceWorksheet.properties,
  };

  targetWorksheet.pageSetup = {
    ...sourceWorksheet.pageSetup,
  };

  targetWorksheet.headerFooter = {
    ...sourceWorksheet.headerFooter,
  };

  targetWorksheet.views = sourceWorksheet.views.map(
    (view) => ({ ...view }),
  );

  console.log(
    `Copied Worksheet      : ${sourceWorksheet.name} -> ${targetSheetName}`,
  );

  console.log(
    `Copied Worksheet Size : ${sourceWorksheet.rowCount} row(s) x ${sourceWorksheet.columnCount} column(s)`,
  );
};

/**
 * เพิ่ม Fee Columns ใน Template อัตโนมัติ
 *
 * ตัวอย่าง ถ้า Template มีถึง Fee 2 แต่ Test Data มีถึง Fee 4
 * ระบบจะเพิ่ม Fee 3 และ Fee 4 รวมถึง Fee Amount ตามรูปแบบจริง
 */
const ensureDynamicFeeColumns = (
  worksheet: ExcelJS.Worksheet,
  headerRowNumber: number,
  maxFeeIndexFromTestData: number,
): void => {
  if (maxFeeIndexFromTestData <= 0) {
    return;
  }

  const headerMap = buildHeaderMap(worksheet, headerRowNumber);
  const maxFeeIndexInTemplate = getMaxFeeIndexFromHeaderMap(headerMap);

  if (maxFeeIndexFromTestData <= maxFeeIndexInTemplate) {
    return;
  }

  const headerRow = worksheet.getRow(headerRowNumber);
  let insertAfterColumn = headerRow.cellCount;

  for (const [header, columnNumber] of headerMap.entries()) {
    if (getFeeIndexFromHeader(header) !== undefined) {
      insertAfterColumn = Math.max(insertAfterColumn, columnNumber);
    }
  }

  const sourceFeeIndex = Math.max(maxFeeIndexInTemplate, 1);
  const sourceHeaders = getFeeHeaderNames(sourceFeeIndex);
  const sourceColumns = sourceHeaders
    .map((header) => headerMap.get(normalizeHeader(header)))
    .filter((columnNumber): columnNumber is number =>
      columnNumber !== undefined,
    );

  const sourceStartColumn =
    sourceColumns.length > 0
      ? Math.min(...sourceColumns)
      : Math.max(1, insertAfterColumn - 4);

  const columnsPerFee = 5;
  const missingFeeCount =
    maxFeeIndexFromTestData - maxFeeIndexInTemplate;

  worksheet.spliceColumns(
    insertAfterColumn + 1,
    0,
    ...Array.from(
      { length: missingFeeCount * columnsPerFee },
      () => [],
    ),
  );

  let targetColumn = insertAfterColumn + 1;

  for (
    let feeIndex = maxFeeIndexInTemplate + 1;
    feeIndex <= maxFeeIndexFromTestData;
    feeIndex += 1
  ) {
    const newHeaders = getFeeHeaderNames(feeIndex);

    newHeaders.forEach((header, offset) => {
      const sourceColumn = sourceStartColumn + offset;
      const targetWorksheetColumn = worksheet.getColumn(targetColumn);
      const sourceWorksheetColumn = worksheet.getColumn(sourceColumn);

      targetWorksheetColumn.width = sourceWorksheetColumn.width;
      targetWorksheetColumn.hidden = sourceWorksheetColumn.hidden;
      targetWorksheetColumn.outlineLevel =
        sourceWorksheetColumn.outlineLevel;

      for (
        let rowNumber = 1;
        rowNumber <= worksheet.rowCount;
        rowNumber += 1
      ) {
        cloneCellStyle(
          worksheet.getRow(rowNumber).getCell(sourceColumn),
          worksheet.getRow(rowNumber).getCell(targetColumn),
        );
      }

      worksheet
        .getRow(headerRowNumber)
        .getCell(targetColumn).value = header;

      targetColumn += 1;
    });
  }
};

const findTemplateHeaderRowNumber = (
  worksheet: ExcelJS.Worksheet,
): number => {
  const maxRowToCheck = Math.min(30, worksheet.rowCount);

  for (let rowNumber = 1; rowNumber <= maxRowToCheck; rowNumber += 1) {
    const headerMap = buildHeaderMap(worksheet, rowNumber);

    if (
      headerMap.has("test result") &&
      headerMap.has("reason") &&
      headerMap.has("test script no.")
    ) {
      return rowNumber;
    }
  }

  throw new Error(
    `Summary detail header row not found in worksheet: ${worksheet.name}`,
  );
};

export const writeReportAutomationSummary = async (
  reportName: string,
  templatePath: string,
  outputPath: string,
  compareResultPath: string,
  originalTestDataPath: string,
  checkedReportPath: string,
  checkedTestDataPath: string,
  resultRows: CompareResultRow[],
  summaryInfo: AutomationSummaryInfo,
): Promise<void> => {
  const config =
    getSummaryReportConfig(
      reportName,
    );

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Summary template not found: ${templatePath}`);
  }

  if (!fs.existsSync(originalTestDataPath)) {
    throw new Error(`Original Test Data not found: ${originalTestDataPath}`);
  }

  if (!fs.existsSync(compareResultPath)) {
    throw new Error(`Compare Result not found: ${compareResultPath}`);
  }

  if (!fs.existsSync(checkedReportPath)) {
    throw new Error(`Checked Report not found: ${checkedReportPath}`);
  }

  if (!fs.existsSync(checkedTestDataPath)) {
    throw new Error(`Checked Test Data not found: ${checkedTestDataPath}`);
  }

  const checkedTestDataResult = await readCheckedTestDataRows(
    originalTestDataPath,
  );

  const testDataMap = checkedTestDataResult.testDataMap;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const summarySheet =
    workbook.getWorksheet(
      config.summarySheetName,
    );

  if (!summarySheet) {
    throw new Error(
      `Worksheet "${config.summarySheetName}" not found in template: ${templatePath}`,
    );
  }

  writeSummaryInformation(
    summarySheet,
    summaryInfo,
    config,
  );

  const headerRowNumber = findTemplateHeaderRowNumber(summarySheet);

  if (
    config.hasDynamicFeeColumns
  ) {
    ensureDynamicFeeColumns(
      summarySheet,
      headerRowNumber,
      checkedTestDataResult.maxFeeIndex,
    );
  }

  const dataStartRow = headerRowNumber + 1;
  const headerRow = summarySheet.getRow(headerRowNumber);

  if (summarySheet.rowCount >= dataStartRow) {
    summarySheet.spliceRows(
      dataStartRow,
      summarySheet.rowCount - dataStartRow + 1,
    );
  }

  resultRows.forEach((compareRow, rowIndex) => {
    const outputRow = summarySheet.getRow(dataStartRow + rowIndex);
    const testDataRow = testDataMap.get(compareRow.testScriptNo);

    headerRow.eachCell(
      { includeEmpty: true },
      (headerCell, columnNumber) => {
        const header = normalizeText(getCellValue(headerCell));

        if (header === "") {
          return;
        }

        const outputCell = outputRow.getCell(columnNumber);

        if (
          columnNumber <=
          config.compareLastColumn
        ) {
          outputCell.value = getCompareValue(header, compareRow) as ExcelJS.CellValue;
        } else if (
          columnNumber >=
          config.testDataFirstColumn
        ) {
          outputCell.value = getTestDataValue(header, testDataRow) as ExcelJS.CellValue;
        }

        outputCell.border = THIN_BORDER;
        outputCell.alignment = {
          vertical: "top",
          horizontal: header === "Test Result" ? "center" : "left",
          wrapText: true,
        };

        if (normalizeHeader(header) === "test result") {
          applyStatusStyle(outputCell, compareRow.status);
        }
      },
    );

    outputRow.height = 36;
  });

  summarySheet.views = [
    {
      state: "frozen",
      ySplit: headerRowNumber,
    },
  ];

  summarySheet.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    paperSize: 9,
  };

  summarySheet.headerFooter = {
    oddFooter: `Compare source: ${path.basename(compareResultPath)} | Original Test Data source: ${path.basename(originalTestDataPath)}`,
  };

  /*
   * Compare Result จาก Script 3
   * -> DS_PTX_Reconcile / DS_FTX_Reconcile
   */
  await copyWorksheetFromFile(
    compareResultPath,
    workbook,
    config.reconcileSheetName,
    [
      config.reconcileSheetName,
      config.reportCode,
      `${config.reportCode} Compare Result`,
      "Compare Result",
      "Reconcile Result",
    ],
  );

  /*
   * Checked Report จาก Script 2
   * -> DS_PTX / DS_FTX
   */
  await copyWorksheetFromFile(
    checkedReportPath,
    workbook,
    config.reportSheetName,
    [
      config.reportSheetName,
      config.reportCode,
    ],
  );

  /* Checked Test Data จาก Script 2 -> Test Data */
  await copyWorksheetFromFile(
    checkedTestDataPath,
    workbook,
    "Test Data",
    ["Test Data"],
  );

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  await workbook.xlsx.writeFile(outputPath);
};

