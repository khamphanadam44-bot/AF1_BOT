/**
 * ptx-row-builder.ts
 * ------------------------------------------------------------------
 * เตรียมข้อมูลของ DS_PTX ก่อนส่งเข้า Compare Engine
 *
 * หน้าที่หลัก:
 * 1. แปลง Excel Row เป็น Object
 * 2. สร้าง Matching Key ของ Expected Row
 * 3. สร้าง Expected Row จาก Test Data
 * 4. สร้าง Actual Row จาก AF1 Report
 *
 * ไฟล์นี้มีหน้าที่เตรียมข้อมูลเท่านั้น
 * ไม่มี Logic ตัดสิน PASS, FAIL
 * และไม่ได้เขียนผลลัพธ์ลงในไฟล์ Excel
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  getMappingHeaderRowNumber,
  getMappingMatchingKeyHeaders,
} from "../../../config/mapping-helper";

import {
  FEE_TYPE_COUNT,
  getFeeAmountHeader,
} from "../../../config/testdata-config";

import {
  getCellText,
  getHeadersFromRow,
} from "../../validators/shared/excel-cell.util";

import type {
  ActualRow,
  ExpectedRow,
  ReportRow,
  TestDataRow,
} from "./ptx-types";

/** แถว Header ของ Test Data */
const TEST_DATA_HEADER_ROW = 5;

/** รหัสระบบที่ใช้สร้าง Matching Key ของ DS_PTX */
const PTX_SYSTEM_ID = "GPMH";

/** ข้อความต่อท้าย Matching Key ของ Report */
const PTX_REPORT_SUFFIX = "RM";

/** จำนวนหลักของ Running Number เช่น 1 จะถูกแปลงเป็น 01 */
const RUNNING_NUMBER_LENGTH = 2;

/** แปลงค่าเป็นข้อความและตัดช่องว่างหัวท้าย */
const toText = (
  value: unknown,
): string => {
  return String(
    value ?? "",
  ).trim();
};

/** แปลงค่าเป็นตัวเลข โดยคืน 0 เมื่อค่าไม่สามารถใช้เป็นตัวเลขได้ */
const toNumber = (
  value: unknown,
): number => {
  if (
    value === undefined ||
    value === null ||
    toText(value) === ""
  ) {
    return 0;
  }

  const result = Number(
    value,
  );

  return Number.isNaN(result)
    ? 0
    : result;
};

/** ตรวจว่าค่าเป็นค่าว่างหรือไม่ */
const isBlank = (
  value: unknown,
): boolean => {
  return (
    value === undefined ||
    value === null ||
    toText(value) === ""
  );
};

/**
 * แปลง Excel Row เป็น Object ที่อ่านค่าด้วยชื่อ Header ได้
 *
 * ตัวอย่าง:
 * Header: Txn Date | Currency Id
 * Data:   25/11/2025 | HKD
 *
 * ผลลัพธ์:
 * {
 *   "Txn Date": "25/11/2025",
 *   "Currency Id": "HKD"
 * }
 */
const mapRowToObject = (
  row: ExcelJS.Row,
  headers: string[],
): Record<string, unknown> => {
  const result: Record<
    string,
    unknown
  > = {};

  headers.forEach(
    (header, index) => {
      if (!header) {
        return;
      }

      result[header] =
        getCellText(
          row.getCell(
            index + 1,
          ),
        );
    },
  );

  return result;
};

/** เติมเลข 0 ด้านหน้า Running Number เช่น 1 เป็น 01 */
const formatRunningNumber = (
  runningNumber: number,
): string => {
  return String(
    runningNumber,
  ).padStart(
    RUNNING_NUMBER_LENGTH,
    "0",
  );
};

/**
 * อ่าน Channel จาก 3 ตัวแรกของ Transaction ID
 *
 * ตัวอย่าง:
 * KBO304... เป็น KBO
 * KMA301... เป็น KMA
 */
const resolveChannelFromTransactionId = (
  transactionId: string,
): string => {
  const normalizedTransactionId =
    transactionId
      .trim()
      .toUpperCase();

  if (
    normalizedTransactionId.length < 3
  ) {
    throw new Error(
      `Cannot resolve Channel from Transaction ID: "${transactionId}"`,
    );
  }

  return normalizedTransactionId.slice(
    0,
    3,
  );
};

/**
 * สร้าง Matching Key สำหรับจับคู่ Expected Row กับ AF1 Report
 *
 * รูปแบบ:
 * TransactionId_RunningNumber_GPMH_Channel_RM
 */
const buildMatchingKey = (
  transactionId: string,
  runningNumber: number,
): string => {
  const normalizedTransactionId =
    transactionId
      .trim()
      .toUpperCase();

  if (
    normalizedTransactionId === ""
  ) {
    throw new Error(
      "Cannot build Matching Key because Transaction ID is blank",
    );
  }

  if (
    !Number.isInteger(
      runningNumber,
    ) ||
    runningNumber < 1
  ) {
    throw new Error(
      `Invalid Running Number: ${runningNumber}`,
    );
  }

  const channel =
    resolveChannelFromTransactionId(
      normalizedTransactionId,
    );

  return [
    normalizedTransactionId,
    formatRunningNumber(
      runningNumber,
    ),
    PTX_SYSTEM_ID,
    channel,
    PTX_REPORT_SUFFIX,
  ].join(
    "_",
  );
};

/** อ่าน Test Script No. โดยรองรับชื่อ Header หลายรูปแบบ */
const getTestScriptNo = (
  rowData: TestDataRow,
): string => {
  const possibleHeaders = [
    "Test No.",
    "Test Script No.",
    "Test No",
    "Test Script No",
  ];

  for (
    const header of possibleHeaders
  ) {
    const value = toText(
      rowData[header],
    );

    if (value !== "") {
      return value;
    }
  }

  return "";
};

/** อ่าน Transaction ID ที่ใช้สร้าง Matching Key */
const getTransactionId = (
  rowData: TestDataRow,
): string => {
  return toText(
    rowData[
      "Transaction ID/ Reconcile ID"
    ],
  );
};

/**
 * ตรวจว่า Test Data เป็นแถวว่างจริงหรือไม่
 *
 * แถวจะถูกข้ามเมื่อ Test No. และ Transaction ID ว่างพร้อมกันเท่านั้น
 * หากมีอย่างใดอย่างหนึ่ง ระบบยังต้องสร้างผลลัพธ์ของ Test Case นั้น
 */
const isActualBlankTestDataRow = (
  rowData: TestDataRow,
): boolean => {
  const testScriptNo =
    getTestScriptNo(
      rowData,
    );

  const transactionId =
    getTransactionId(
      rowData,
    );

  return (
    testScriptNo === "" &&
    transactionId === ""
  );
};

/** สร้าง Expected Row จำนวนหนึ่งรายการ */
const createExpectedRow = (
  rowNumber: number,
  sourceRow: TestDataRow,
  runningNumber: number,
  feeType: string,
  feeAmount: unknown,
  hasFee: boolean,
): ExpectedRow => {
  const transactionId =
    getTransactionId(
      sourceRow,
    );

  const testScriptNo =
    getTestScriptNo(
      sourceRow,
    );

  /**
   * รายการที่มี Fee ใช้ Matching Key จริงเพื่อจับคู่กับ Report
   * รายการที่ไม่มี Fee ใช้ Internal Key เพื่อสร้างผล SKIP
   */
  const matchingKey = hasFee
    ? buildMatchingKey(
        transactionId,
        runningNumber,
      )
    : `NO_FEE_ROW_${rowNumber}`;

  return {
    rowNumber,
    testScriptNo,
    matchingKey,
    runningNumber,

    feeType: toText(
      feeType,
    ),

    feeAmount: toNumber(
      feeAmount,
    ),

    hasFee,
    data: sourceRow,
  };
};

/**
 * สร้าง Expected Row จาก Test Data
 *
 * กติกา:
 * - Test Data หนึ่งแถวสามารถสร้างได้หลาย Expected Row
 * - Fee หนึ่งลำดับสร้าง Expected Row หนึ่งรายการ
 * - หากไม่มี Fee ทุกลำดับ จะสร้างหนึ่งรายการสำหรับผล SKIP
 */
export const buildExpectedRows = (
  worksheet: ExcelJS.Worksheet,
): ExpectedRow[] => {
  const headers =
    getHeadersFromRow(
      worksheet,
      TEST_DATA_HEADER_ROW,
    );

  const expectedRows: ExpectedRow[] = [];

  for (
    let rowNumber =
      TEST_DATA_HEADER_ROW + 1;

    rowNumber <= worksheet.rowCount;

    rowNumber += 1
  ) {
    const row =
      worksheet.getRow(
        rowNumber,
      );

    const rowData =
      mapRowToObject(
        row,
        headers,
      ) as TestDataRow;

    if (
      isActualBlankTestDataRow(
        rowData,
      )
    ) {
      continue;
    }

    let hasAnyFee = false;

    for (
      let feeIndex = 1;

      feeIndex <= FEE_TYPE_COUNT;

      feeIndex += 1
    ) {
      const feeType = toText(
        rowData[
          `Fee Type ${feeIndex}`
        ],
      );

      const feeAmountHeader =
        getFeeAmountHeader(
          feeIndex,
        );

      const feeAmount =
        rowData[
          feeAmountHeader
        ];

      /**
       * Fee Type และ Fee Amount ว่างพร้อมกัน
       * แสดงว่า Fee Slot นี้ไม่มีข้อมูล
       */
      if (
        feeType === "" &&
        isBlank(feeAmount)
      ) {
        continue;
      }

      hasAnyFee = true;

      /**
       * Fee หนึ่งลำดับสร้าง Expected Row หนึ่งรายการ
       */
      expectedRows.push(
        createExpectedRow(
          rowNumber,
          rowData,
          feeIndex,
          feeType,
          feeAmount,
          true,
        ),
      );
    }

    /**
     * Test Case ไม่มี Fee ทุกลำดับ
     * สร้างหนึ่งรายการเพื่อให้ Compare Engine สร้างผล SKIP
     */
    if (!hasAnyFee) {
      expectedRows.push(
        createExpectedRow(
          rowNumber,
          rowData,
          0,
          "",
          0,
          false,
        ),
      );
    }
  }

  return expectedRows;
};

/**
 * สร้าง Actual Row จาก AF1 Report
 *
 * กติกา:
 * - อ่าน Header Row และ Matching Key จาก Mapping Config
 * - ข้ามแถวที่ไม่มี Matching Key
 * - เก็บ Row Number จริงสำหรับใช้แสดงในผล Compare
 */
export const buildActualRows = (
  worksheet: ExcelJS.Worksheet,
  reportName: string,
): ActualRow[] => {
  const headerRowNumber =
    getMappingHeaderRowNumber(
      reportName,
    );

  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  const matchingKeyHeader =
    getMappingMatchingKeyHeaders(
      reportName,
    )[0];

  if (!matchingKeyHeader) {
    throw new Error(
      `Matching Key Header not found for report: ${reportName}`,
    );
  }

  const actualRows: ActualRow[] = [];

  for (
    let rowNumber =
      headerRowNumber + 1;

    rowNumber <= worksheet.rowCount;

    rowNumber += 1
  ) {
    const row =
      worksheet.getRow(
        rowNumber,
      );

    const rowData =
      mapRowToObject(
        row,
        headers,
      ) as ReportRow;

    const matchingKey = toText(
      rowData[
        matchingKeyHeader
      ],
    );

    if (matchingKey === "") {
      continue;
    }

    actualRows.push({
      rowNumber,
      matchingKey,
      data: rowData,
    });
  }

  return actualRows;
};