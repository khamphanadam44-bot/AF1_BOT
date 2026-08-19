/**
 * summary-result-builder.ts
 * ------------------------------------------------------------------
 * แปลงผล Reconcile ระดับ Record เป็น Summary ระดับ Test Case
 *
 * - DS_LTX: Group DR/FE ตาม Test No.
 * - DS_PTX: Group Fee ตาม Test No.
 * - Report อื่น: 1 Compare Row = 1 Test Case
 * ------------------------------------------------------------------
 */

import type {
  CompareResultRow,
  SummaryStatus,
} from "./summary-types";

export type AutomationSummaryResult = {
  rows: CompareResultRow[];
  totalChecked: number;
  passed: number;
  failed: number;
  skipped: number;
};

type GroupedSummaryReport =
  | "DS_LTX"
  | "DS_PTX";

const toText = (value: unknown): string =>
  String(value ?? "").trim();

/** ไม่ตัด -Return เพราะต้องนับเป็นคนละ Test Case */
const normalizeTestNo = (value: unknown): string =>
  toText(value).toUpperCase();

const normalizeReportName = (value: string): string =>
  toText(value)
    .toUpperCase()
    .replace(/-/g, "_");

/** นับ KPI หลังเตรียม Summary Rows แล้ว */
const buildSummaryResult = (
  rows: CompareResultRow[],
): AutomationSummaryResult => ({
  rows,
  totalChecked: rows.length,
  passed: rows.filter((row) => row.status === "PASS").length,
  failed: rows.filter((row) => row.status === "FAIL").length,
  skipped: rows.filter((row) => row.status === "SKIP").length,
});

/** สถานะระดับ Test Case: FAIL > PASS > SKIP */
const resolveStatus = (
  rows: CompareResultRow[],
): SummaryStatus => {
  if (rows.some((row) => row.status === "FAIL")) {
    return "FAIL";
  }

  if (rows.some((row) => row.status === "PASS")) {
    return "PASS";
  }

  return "SKIP";
};

/**
 * Group ตาม Test No.
 * Test No. ว่างจะแยกแต่ละ Row เพื่อไม่ให้หลาย Case ถูกรวมกัน
 */
const groupByTestNo = (
  reportName: GroupedSummaryReport,
  rows: CompareResultRow[],
): CompareResultRow[][] => {
  const groups =
    new Map<string, CompareResultRow[]>();

  rows.forEach((row, index) => {
    const testNo =
      normalizeTestNo(row.testScriptNo);

    const key =
      testNo ||
      `__${reportName}_ROW_${index}`;

    const group =
      groups.get(key) ?? [];

    group.push(row);
    groups.set(key, group);
  });

  return [...groups.values()];
};

/** สร้างหนึ่ง Summary Row จากหลาย Reconcile Records */
const buildGroupedRow = (
  reportName: GroupedSummaryReport,
  rows: CompareResultRow[],
  remark: string,
): CompareResultRow => {
  const firstRow =
    rows[0];

  if (!firstRow) {
    throw new Error(
      `Cannot build ${reportName} Summary row from empty group.`,
    );
  }

  const status =
    resolveStatus(rows);

  const representativeRow =
    rows.find((row) => row.status === status) ??
    firstRow;

  return {
    ...representativeRow,

    // รักษา Test No. ตามต้นทาง รวมถึง -Return
    testScriptNo: firstRow.testScriptNo,

    status,
    remark,
  };
};

/* ======================================================
 * DS_LTX
 * ====================================================== */

const getLtxRecordLabel = (
  row: CompareResultRow,
): string => {
  const reference =
    toText(row.matchingKey).toUpperCase();

  if (reference.endsWith("DR")) {
    return "DR";
  }

  if (reference.endsWith("FE")) {
    return "FE";
  }

  /**
   * Matching Key อาจว่างเมื่อ Match ไม่สำเร็จ
   * จึงใช้ข้อความ "แถว DR/FE" เป็น fallback
   */
  if (/แถว\s+DR\b/i.test(row.remark)) {
    return "DR";
  }

  if (/แถว\s+FE\b/i.test(row.remark)) {
    return "FE";
  }

  return "Record";
};

/**
 * PASS Review แสดงเฉพาะ Field ที่ไม่ตรงกัน
 * ส่วน Expected/Actual Value ดูจาก Reconcile
 */
const summarizeLtxReview = (
  detail: string,
): string | undefined => {
  if (!/Please review/i.test(detail)) {
    return undefined;
  }

  const matches =
    detail.matchAll(
      /\[TS\]\s*:\s*([^=\r\n|]+?)\s*=\s*[\s\S]*?\|\s*\[AF1-LTX\]\s*:\s*([^=\r\n|]+?)\s*=/gi,
    );

  const fieldPairs =
    [...matches]
      .map((match) => {
        const testField =
          toText(match[1]);

        const reportField =
          toText(match[2]);

        if (!testField || !reportField) {
          return "";
        }

        return (
          `${testField} → ` +
          `${reportField} : ไม่ตรงกัน`
        );
      })
      .filter(
        (value, index, values) =>
          value !== "" &&
          values.indexOf(value) === index,
      );

  if (fieldPairs.length === 0) {
    return undefined;
  }

  return [
    "Review",
    ...fieldPairs,
  ].join("\n");
};

/** ย่อ Remark เฉพาะที่จำเป็นสำหรับ LTX Summary */
const summarizeLtxRemark = (
  row: CompareResultRow,
): string => {
  const detail =
    toText(row.remark);

  if (!detail) {
    return "";
  }

  /**
   * กรณีหา Reference ไม่พบจากทั้ง Exact และ Fallback
   */
  const exactReference =
    /ไม่พบ Exact Reference\s*=\s*"([^"]+)"/i.exec(
      detail,
    )?.[1];

  const notFound =
    /ไม่พบแถว\s+(?:DR|FE)\b.*Exact Reference.*Fallback Matching/i.test(
      detail,
    );

  if (exactReference && notFound) {
    return (
      `ไม่พบ Ref "${exactReference}" ` +
      "จาก Exact/Fallback"
    );
  }

  return (
    summarizeLtxReview(detail) ??
    detail
  );
};

/**
 * Expected Absence เป็นเงื่อนไขระดับ Test Case
 * จึงแสดง Remark ครั้งเดียว ไม่แยก DR/FE
 */
const getLtxExpectedAbsenceRemark = (
  rows: CompareResultRow[],
  details: string[],
): string | undefined => {
  if (
    !rows.every(
      (row) =>
        row.status === "PASS",
    )
  ) {
    return undefined;
  }

  const absenceDetails =
    details.filter(
      (detail) =>
        /รายการต้องไม่แสดงใน\s+DS_LTX/i.test(
          detail,
        ),
    );

  if (absenceDetails.length === 0) {
    return undefined;
  }

  /**
   * DR/FE อาจมี Remark เดียวกันแต่ช่องว่างต่างกัน
   * จึง Normalize เฉพาะเพื่อเช็กว่าเป็นข้อความเดียวกัน
   */
  const uniqueDetails =
    new Set(
      absenceDetails.map(
        (detail) =>
          detail
            .replace(/\s+/g, " ")
            .trim(),
      ),
    );

  if (uniqueDetails.size !== 1) {
    return undefined;
  }

  return absenceDetails[0];
};

const buildLtxSummaryRow = (
  rows: CompareResultRow[],
): CompareResultRow => {
  const details =
    rows.map(summarizeLtxRemark);

  /**
   * Expected Absence:
   * ไม่ต้องแสดง Record/DR/FE ซ้ำ
   */
  const expectedAbsenceRemark =
    getLtxExpectedAbsenceRemark(
      rows,
      details,
    );

  if (expectedAbsenceRemark) {
    return buildGroupedRow(
      "DS_LTX",
      rows,
      expectedAbsenceRemark,
    );
  }

  /**
   * ไม่มี Remark ทุก Record
   * Writer จะเติม Validation passed/failed/skipped ให้เอง
   */
  if (details.every((detail) => detail === "")) {
    return buildGroupedRow(
      "DS_LTX",
      rows,
      "",
    );
  }

  const remark =
    rows
      .map((row, index) => {
        const label =
          getLtxRecordLabel(row);

        const detail =
          details[index] ?? "";

        const result =
          `${label}: ${row.status}`;

        return detail
          ? `${result} - ${detail}`
          : result;
      })
      .join("\n\n");

  return buildGroupedRow(
    "DS_LTX",
    rows,
    remark,
  );
};

/* ======================================================
 * DS_PTX
 * ====================================================== */

/**
 * PTX รวมหลาย Fee Record เป็นหนึ่ง Test Case
 * และเก็บเฉพาะ Remark ที่มีจริงโดยไม่ซ้ำกัน
 */
const buildPtxSummaryRow = (
  rows: CompareResultRow[],
): CompareResultRow => {
  const remarks =
    rows
      .map((row) => toText(row.remark))
      .filter(
        (remark, index, values) =>
          remark !== "" &&
          values.indexOf(remark) === index,
      );

  return buildGroupedRow(
    "DS_PTX",
    rows,
    remarks.join(" | "),
  );
};

/* ======================================================
 * Public API
 * ====================================================== */

export const buildAutomationSummaryResult = (
  reportName: string,
  compareRows: CompareResultRow[],
): AutomationSummaryResult => {
  const normalizedReportName =
    normalizeReportName(reportName);

  /**
   * Report อื่นคงพฤติกรรมเดิม:
   * 1 Compare Row = 1 Test Case
   */
  if (
    normalizedReportName !== "DS_LTX" &&
    normalizedReportName !== "DS_PTX"
  ) {
    return buildSummaryResult(
      compareRows,
    );
  }

  const groupedReportName:
    GroupedSummaryReport =
      normalizedReportName === "DS_LTX"
        ? "DS_LTX"
        : "DS_PTX";

  const groups =
    groupByTestNo(
      groupedReportName,
      compareRows,
    );

  const summaryRows =
    groups.map((rows) =>
      groupedReportName === "DS_LTX"
        ? buildLtxSummaryRow(rows)
        : buildPtxSummaryRow(rows),
    );

  return buildSummaryResult(
    summaryRows,
  );
};