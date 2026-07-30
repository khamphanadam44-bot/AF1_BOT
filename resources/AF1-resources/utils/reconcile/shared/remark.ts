/**
 * Remark Formatter
 * ------------------------------------------------------------------
 * สร้างข้อความเปรียบเทียบ Test Script กับ AF1 Report
 * ให้ทุก Report ใช้ Format เดียวกัน
 * ------------------------------------------------------------------
 */

const displayValue = (value: unknown): string => {
  return String(value ?? "").trim();
};

const getReportLabel = (reportCode: string): string => {
  return String(reportCode ?? "")
    .trim()
    .toUpperCase()
    .replace(/^DS_|^DF_/, "");
};

export const formatCompareRemark = (
  reportCode: string,
  testDataField: string,
  expectedValue: unknown,
  reportField: string,
  actualValue: unknown,
): string => {
  const reportLabel = getReportLabel(reportCode);

  return (
    `[TS] : ${testDataField} = ` +
    `"${displayValue(expectedValue)}" | ` +
    `[AF1-${reportLabel}] : ${reportField} = ` +
    `"${displayValue(actualValue)}"`
  );
};

export const formatFixedValueRemark = (
  reportCode: string,
  reportField: string,
  actualValue: unknown,
  expectedValue: unknown,
): string => {
  const reportLabel = getReportLabel(reportCode);

  return (
    `[AF1-${reportLabel}] : ${reportField} = ` +
    `"${displayValue(actualValue)}" ต้องเป็น ` +
    `"${displayValue(expectedValue)}"`
  );
};
