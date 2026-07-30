/**
 * ======================================================
 * Mapping Helper
 * ------------------------------------------------------
 * Helper สำหรับ Script3 (Data Validation Compare)
 *
 * ใช้งานร่วมกับ mapping-config.ts
 *
 * Script2 จะยังใช้ report-helper.ts ตามเดิม
 * Script3 จะใช้ mapping-helper.ts แทน
 * ======================================================
 */

import { REPORT_CONFIG_mapping } from "./mapping-config";

export type MappingReportName =
  keyof typeof REPORT_CONFIG_mapping;

/**
 * ตรวจว่าชื่อ Report มีอยู่ใน Mapping Config หรือไม่
 */
export const isMappingReportName = (
  reportName: string,
): reportName is MappingReportName => {
  return Object.prototype.hasOwnProperty.call(
    REPORT_CONFIG_mapping,
    reportName,
  );
};

/**
 * ตรวจสอบและคืนชื่อ Report ที่ใช้กับ Mapping Config ได้
 */
export const requireMappingReportName = (
  reportName: string,
): MappingReportName => {
  const normalizedReportName =
    String(
      reportName ?? "",
    )
      .trim()
      .toUpperCase();

  if (
    !isMappingReportName(
      normalizedReportName,
    )
  ) {
    throw new Error(
      `ไม่พบ Mapping Config สำหรับ Report "${reportName}".`,
    );
  }

  return normalizedReportName;
};

/**
 * ดึง Config ของ Report
 */
const getReportConfig = (
  reportName: string,
) => {

  return REPORT_CONFIG_mapping[
    reportName as MappingReportName
  ];

};

/**
 * ======================================================
 * แถวที่เก็บ Header (Header Row)
 * ======================================================
 */

/**
 * ดึง Header Row Number
 */
export const getMappingHeaderRowNumber = (
  reportName: string,
): number => {

  return getReportConfig(
    reportName,
  ).headerRowNumber;

};

/**
 * ======================================================
 * กลุ่มของ Header
 * ======================================================
 */

/**
 * กุญแจสำหรับจับคู่ข้อมูล (Matching Key)
 */
export const getMappingMatchingKeyHeaders = (
  reportName: string,
): string[] => {

  return [

    ...getReportConfig(
      reportName,
    ).requiredHeaders.matchingKey,

  ];

};

/**
 * ข้อมูลหลัก (Core)
 */
export const getMappingCoreHeaders = (
  reportName: string,
): string[] => {

  return [

    ...getReportConfig(
      reportName,
    ).requiredHeaders.core,

  ];

};

/**
 * ข้อมูลลูกค้า (Customer)
 */
export const getMappingCustomerHeaders = (
  reportName: string,
): string[] => {

  return [

    ...getReportConfig(
      reportName,
    ).requiredHeaders.customer,

  ];

};

/**
 * ดึง Header ทุกกลุ่มและตัดค่าซ้ำ
 *
 * ใช้ตรวจสอบว่า Field Rule ของ Reconcile อ้างถึง
 * Header ที่มีอยู่ใน Mapping Config จริง
 */
export const getUniqueMappingHeaders = (
  reportName: string,
): string[] => {
  const requiredHeaders =
    getReportConfig(
      reportName,
    ).requiredHeaders;

  return [
    ...new Set(
      Object.values(
        requiredHeaders,
      )
        .flat()
        .map(
          (
            header,
          ) =>
            String(
              header ?? "",
            ).trim(),
        )
        .filter(
          Boolean,
        ),
    ),
  ];
};
