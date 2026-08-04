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
 * 5. ตรวจจำนวน Fee Group จาก Header ของ Test Data จริง
 *
 * ไฟล์นี้มีหน้าที่เตรียมข้อมูลเท่านั้น
 * ไม่มี Logic ตัดสิน PASS หรือ FAIL
 * และไม่ได้เขียนผลลัพธ์ลงในไฟล์ Excel
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  getMappingHeaderRowNumber,
  getMappingMatchingKeyHeaders,
} from "../../../config/mapping-helper";

import {
  getFeeAmountHeader,
} from "../../../config/testdata-config";

import {
  detectFeeTypeCount,
} from "../../../config/testdata-helper";

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

/**
 * หมายเลขแถว Header ของ Test Data
 *
 * Test Data เริ่ม Header ที่แถว 5
 * และข้อมูลจริงเริ่มที่แถว 6
 */
const TEST_DATA_HEADER_ROW = 5;

/**
 * รหัสระบบที่ใช้สร้าง Matching Key ของ DS_PTX
 */
const PTX_SYSTEM_ID = "GPMH";

/**
 * ข้อความต่อท้าย Matching Key ของ Report
 */
const PTX_REPORT_SUFFIX = "RM";

/**
 * จำนวนหลักของ Running Number
 *
 * ตัวอย่าง:
 * 1  → 01
 * 2  → 02
 * 10 → 10
 */
const RUNNING_NUMBER_LENGTH = 2;

/**
 * แปลงค่าเป็นข้อความและตัดช่องว่างหัวท้าย
 *
 * หากค่าเป็น null หรือ undefined
 * จะคืนค่าเป็นข้อความว่าง
 */
const toText = (
  value: unknown,
): string => {
  return String(
    value ?? "",
  ).trim();
};

/**
 * แปลงค่าเป็นตัวเลข
 *
 * คืนค่า 0 เมื่อ:
 * - ค่าเป็น undefined
 * - ค่าเป็น null
 * - ค่าเป็นข้อความว่าง
 * - ค่าไม่สามารถแปลงเป็นตัวเลขได้
 */
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

  const result =
    Number(value);

  return Number.isNaN(result)
    ? 0
    : result;
};

/**
 * ตรวจว่าค่าเป็นค่าว่างหรือไม่
 *
 * คืนค่า:
 * true  = ค่าว่าง
 * false = มีข้อมูล
 */
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
 * แปลง Excel Row เป็น Object
 * ที่สามารถอ่านค่าด้วยชื่อ Header ได้
 *
 * ตัวอย่างข้อมูล:
 *
 * Header:
 * Txn Date | From Currency (CCY)
 *
 * Data:
 * 25/11/2025 | HKD
 *
 * ผลลัพธ์:
 * {
 *   "Txn Date": "25/11/2025",
 *   "From Currency (CCY)": "HKD"
 * }
 *
 * หมายเหตุ:
 * Array ของ Header เริ่มนับจาก Index 0
 * แต่ ExcelJS เริ่มนับ Column จาก 1
 * จึงต้องใช้ index + 1
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
      /**
       * ข้าม Column ที่ไม่มีชื่อ Header
       */
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

/**
 * เติมเลข 0 ด้านหน้า Running Number
 *
 * ตัวอย่าง:
 * 1 → 01
 * 2 → 02
 */
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
 * KBO304... → KBO
 * KMA301... → KMA
 *
 * หาก Transaction ID มีความยาวน้อยกว่า 3 ตัว
 * ระบบจะไม่สามารถสร้าง Matching Key ได้
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
 * สร้าง Matching Key สำหรับจับคู่
 * Expected Row กับ AF1 Report
 *
 * รูปแบบ:
 * TransactionId_RunningNumber_GPMH_Channel_RM
 *
 * ตัวอย่าง:
 * KBO304001_01_GPMH_KBO_RM
 *
 * Running Number ใช้ลำดับของ Fee Group
 *
 * ตัวอย่าง:
 * Fee Group 1 → 01
 * Fee Group 2 → 02
 * Fee Group 3 → 03
 */
const buildMatchingKey = (
  transactionId: string,
  runningNumber: number,
): string => {
  const normalizedTransactionId =
    transactionId
      .trim()
      .toUpperCase();

  /**
   * รายการที่มี Fee จำเป็นต้องมี Transaction ID
   * เพราะต้องใช้สร้าง Matching Key
   */
  if (
    normalizedTransactionId === ""
  ) {
    throw new Error(
      "Cannot build Matching Key because Transaction ID is blank",
    );
  }

  /**
   * Running Number ของ Fee ต้องเริ่มตั้งแต่ 1
   */
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

/**
 * อ่าน Test Script No.
 * โดยรองรับชื่อ Header หลายรูปแบบ
 *
 * ลำดับการค้นหา:
 * 1. Test No.
 * 2. Test Script No.
 * 3. Test No
 * 4. Test Script No
 *
 * เมื่อพบค่าที่ไม่ว่างจะคืนค่านั้นทันที
 */
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
    const header
    of possibleHeaders
  ) {
    const value =
      toText(
        rowData[header],
      );

    if (value !== "") {
      return value;
    }
  }

  return "";
};

/**
 * อ่าน Transaction ID
 * ที่ใช้สร้าง Matching Key
 */
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
 * ข้ามแถวเฉพาะเมื่อ:
 * - Test No. ว่าง
 * - Transaction ID ว่าง
 *
 * และทั้งสองช่องว่างพร้อมกันเท่านั้น
 *
 * หากมีข้อมูลในช่องใดช่องหนึ่ง
 * ระบบยังต้องประมวลผล Test Case นั้นต่อ
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

/**
 * สร้าง Expected Row หนึ่งรายการ
 *
 * @param rowNumber
 * หมายเลขแถวจริงใน Test Data
 *
 * @param sourceRow
 * ข้อมูลทั้งหมดของ Test Data แถวนั้น
 *
 * @param runningNumber
 * ลำดับของ Fee Group
 *
 * @param feeType
 * ประเภท Fee ของกลุ่มนั้น
 *
 * @param feeAmount
 * จำนวนเงิน Fee ของกลุ่มนั้น
 *
 * @param hasFee
 * ระบุว่า Expected Row นี้มี Fee หรือไม่
 */
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
   * รายการที่มี Fee:
   * ใช้ Matching Key จริงเพื่อจับคู่กับ Report
   *
   * รายการที่ไม่มี Fee:
   * ใช้ Internal Key เพื่อให้ Compare Engine
   * สามารถสร้างผล SKIP ได้
   */
  const matchingKey =
    hasFee
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

    feeType:
      toText(
        feeType,
      ),

    feeAmount:
      toNumber(
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
 * 1. Test Data หนึ่งแถวสามารถสร้างหลาย Expected Row
 * 2. Fee หนึ่งกลุ่มสร้าง Expected Row หนึ่งรายการ
 * 3. Running Number ใช้หมายเลข Fee Group
 * 4. หากไม่มี Fee ทุกกลุ่ม จะสร้างหนึ่งรายการสำหรับผล SKIP
 * 5. จำนวน Fee Group ตรวจจาก Header ของ Test Data จริง
 *
 * ตัวอย่าง:
 *
 * Test Data หนึ่งแถวมี:
 * - Fee Type 1 และ Fee Amount Type 1
 * - Fee Type 2 และ Fee Amount 2
 *
 * ระบบจะสร้าง Expected Row จำนวน 2 รายการ
 */
export const buildExpectedRows = (
  worksheet: ExcelJS.Worksheet,
): ExpectedRow[] => {
  /**
   * อ่าน Header จริงจาก Test Data
   */
  const headers =
    getHeadersFromRow(
      worksheet,
      TEST_DATA_HEADER_ROW,
    );

  /**
   * ตรวจหมายเลข Fee Group สูงสุด
   * จาก Header ที่พบจริงใน Test Data
   *
   * ตัวอย่าง:
   * หากพบ Fee Type 1 ถึง Fee Type 5
   * feeTypeCount จะมีค่าเป็น 5
   *
   * ไม่ใช้ FEE_TYPE_COUNT แบบ Hard code อีกต่อไป
   */
  const feeTypeCount =
    detectFeeTypeCount(
      headers,
    );

  const expectedRows:
    ExpectedRow[] = [];

  /**
   * เริ่มอ่านข้อมูลจากแถวถัดจาก Header
   */
  for (
    let rowNumber =
      TEST_DATA_HEADER_ROW + 1;

    rowNumber <=
      worksheet.rowCount;

    rowNumber += 1
  ) {
    const row =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * แปลง Excel Row เป็น Object
     * เพื่อให้สามารถอ่านข้อมูลด้วยชื่อ Header
     */
    const rowData =
      mapRowToObject(
        row,
        headers,
      ) as TestDataRow;

    /**
     * ข้ามแถวที่ไม่มีทั้ง Test No.
     * และ Transaction ID
     */
    if (
      isActualBlankTestDataRow(
        rowData,
      )
    ) {
      continue;
    }

    let hasAnyFee =
      false;

    /**
     * ตรวจ Fee Group ตั้งแต่ 1
     * ถึงหมายเลขสูงสุดที่พบจาก Header จริง
     */
    for (
      let feeIndex = 1;

      feeIndex <=
        feeTypeCount;

      feeIndex += 1
    ) {
      const feeType =
        toText(
          rowData[
            `Fee Type ${feeIndex}`
          ],
        );

      /**
       * ชื่อ Fee Amount ของกลุ่มแรก
       * แตกต่างจากกลุ่มถัดไป
       *
       * Fee 1:
       * Fee Amount Type 1
       *
       * Fee 2 เป็นต้นไป:
       * Fee Amount 2
       * Fee Amount 3
       * ...
       */
      const feeAmountHeader =
        getFeeAmountHeader(
          feeIndex,
        );

      const feeAmount =
        rowData[
          feeAmountHeader
        ];

      /**
       * หาก Fee Type และ Fee Amount
       * ว่างพร้อมกัน แสดงว่า Fee Group นี้ไม่มีข้อมูล
       *
       * ระบบจะข้าม Fee Group นี้
       * และตรวจกลุ่มถัดไป
       */
      if (
        feeType === "" &&
        isBlank(
          feeAmount,
        )
      ) {
        continue;
      }

      hasAnyFee =
        true;

      /**
       * Fee หนึ่งกลุ่ม
       * สร้าง Expected Row หนึ่งรายการ
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
     * หาก Test Case ไม่มี Fee ทุกกลุ่ม
     * ให้สร้าง Expected Row หนึ่งรายการ
     *
     * Expected Row นี้ใช้สำหรับให้ Compare Engine
     * สร้างผลลัพธ์เป็น SKIP
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
 * 1. อ่านหมายเลขแถว Header จาก Mapping Config
 * 2. อ่านชื่อ Matching Key Header จาก Mapping Config
 * 3. แปลงแต่ละแถวของ Report เป็น Object
 * 4. ข้ามแถวที่ไม่มี Matching Key
 * 5. เก็บ Row Number จริงสำหรับแสดงในผล Compare
 *
 * Report ที่ส่งเข้ามาต้องมี Mapping Config
 * สำหรับ Header Row และ Matching Key
 */
export const buildActualRows = (
  worksheet: ExcelJS.Worksheet,
  reportName: string,
): ActualRow[] => {
  /**
   * อ่านหมายเลขแถว Header ของ Report
   * จาก Mapping Config
   */
  const headerRowNumber =
    getMappingHeaderRowNumber(
      reportName,
    );

  /**
   * อ่าน Header จริงของ AF1 Report
   */
  const headers =
    getHeadersFromRow(
      worksheet,
      headerRowNumber,
    );

  /**
   * อ่านชื่อ Header ที่ใช้เป็น Matching Key
   *
   * DS_PTX ใช้ Matching Key ตัวแรก
   * ที่กำหนดไว้ใน Mapping Config
   */
  const matchingKeyHeader =
    getMappingMatchingKeyHeaders(
      reportName,
    )[0];

  if (!matchingKeyHeader) {
    throw new Error(
      `Matching Key Header not found for report: ${reportName}`,
    );
  }

  const actualRows:
    ActualRow[] = [];

  /**
   * เริ่มอ่านข้อมูลจากแถวถัดจาก Header
   */
  for (
    let rowNumber =
      headerRowNumber + 1;

    rowNumber <=
      worksheet.rowCount;

    rowNumber += 1
  ) {
    const row =
      worksheet.getRow(
        rowNumber,
      );

    /**
     * แปลง Excel Row เป็น ReportRow Object
     */
    const rowData =
      mapRowToObject(
        row,
        headers,
      ) as ReportRow;

    /**
     * อ่าน Matching Key จาก Header
     * ที่กำหนดไว้ใน Mapping Config
     */
    const matchingKey =
      toText(
        rowData[
          matchingKeyHeader
        ],
      );

    /**
     * ข้ามแถวที่ไม่มี Matching Key
     * เพราะไม่สามารถนำไปจับคู่กับ Expected Row ได้
     */
    if (
      matchingKey === ""
    ) {
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