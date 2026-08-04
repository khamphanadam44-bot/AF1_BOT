/**
 * ftx-result-writer.ts
 * ------------------------------------------------------------------
 * หน้าที่ของไฟล์นี้
 *
 * สร้างไฟล์ Excel ผลลัพธ์ของ Script 3 สำหรับ DS_FTX
 *
 * รูปแบบไฟล์ผลลัพธ์
 * - มี 1 Sheet ชื่อ "DS_FTX"
 * - แสดง Header และข้อมูลของ Report DS_FTX
 * - เพิ่ม 3 Column ไว้ด้านหน้า ได้แก่
 *   1. Test Script No.
 *   2. Result
 *   3. Remark
 *
 * การแสดงข้อมูล
 * - Field ที่ตรวจผ่านจะถูก Highlight สีเขียว
 * - Field ที่ตรวจไม่ผ่านจะถูก Highlight สีแดง
 * - Report Row ที่ไม่มี Test Data จับคู่ยังถูกแสดง
 * - Expected Row ที่ไม่พบใน Report จะถูกเพิ่มท้ายไฟล์
 * - Exclusion ที่ไม่มีใน Report จะถูกเพิ่มท้ายไฟล์เป็น PASS
 * - Matching Key เดียวกันใน Test Data สามารถมีหลายแถวได้
 *
 * หมายเหตุสำคัญ
 * จำนวน Output Row ไม่จำเป็นต้องเท่ากับจำนวน Test Data Row
 *
 * ตัวอย่าง:
 * - Report Row ที่ไม่มี Test Data จะสร้าง Output Row เพิ่ม
 * - Matching Key ซ้ำใน Report อาจทำให้ Test Data เดิมแสดงหลายแถว
 * ------------------------------------------------------------------
 */

import ExcelJS from "exceljs";

import {
  ActualRow,
  CompareResult,
  ExpectedRow,
} from "./ftx-types";

/**
 * รหัสสีที่ใช้ภายในไฟล์ผลลัพธ์
 *
 * รหัสสีใช้รูปแบบ ARGB
 */
const COLORS = {
  /** สีน้ำเงินสำหรับ Header */
  REPORT_HEADER:
    "FF4472C4",

  /** สีขาวสำหรับตัวอักษร Header */
  HEADER_TEXT:
    "FFFFFFFF",

  /** สีเขียวอ่อนสำหรับข้อมูล PASS */
  PASS_FILL:
    "FFC6EFCE",

  /** สีเขียวเข้มสำหรับตัวอักษร PASS */
  PASS_TEXT:
    "FF006100",

  /** สีแดงอ่อนสำหรับข้อมูล FAIL */
  FAIL_FILL:
    "FFFFC7CE",

  /** สีแดงเข้มสำหรับตัวอักษร FAIL */
  FAIL_TEXT:
    "FF9C0006",

  /** สีเทาสำหรับเส้นขอบ Cell */
  BORDER:
    "FFB7B7B7",
};

/**
 * ชื่อ Header ที่ใช้เป็น Matching Key ใน Report DS_FTX
 */
const MATCHING_KEY_HEADER =
  "Ref. TX No.";

/**
 * สถานะรวมของ Test Data จำนวน 1 แถว
 *
 * PASS = ทุก CompareResult ในกลุ่มผ่าน
 * FAIL = มี CompareResult อย่างน้อย 1 รายการไม่ผ่าน
 */
type GroupStatus =
  | "PASS"
  | "FAIL";

/**
 * รูปแบบเส้นขอบบางที่ใช้กับ Cell ทุกด้าน
 */
const THIN_BORDER:
Partial<ExcelJS.Borders> = {
  top: {
    style:
      "thin",
    color: {
      argb:
        COLORS.BORDER,
    },
  },

  left: {
    style:
      "thin",
    color: {
      argb:
        COLORS.BORDER,
    },
  },

  bottom: {
    style:
      "thin",
    color: {
      argb:
        COLORS.BORDER,
    },
  },

  right: {
    style:
      "thin",
    color: {
      argb:
        COLORS.BORDER,
    },
  },
};

/**
 * ปรับค่าก่อนเขียนลงใน Excel
 *
 * กรณีต่อไปนี้จะถูกเปลี่ยนเป็น null
 * เพื่อให้ Excel สร้าง Cell ว่างจริง:
 * - undefined
 * - null
 * - ข้อความว่าง
 * - ข้อความที่มีแต่ช่องว่าง
 *
 * ค่าที่ ExcelJS รองรับโดยตรงจะถูกคืนตามเดิม:
 * - string
 * - number
 * - boolean
 * - Date
 *
 * Object ชนิดอื่นจะถูกแปลงเป็น string
 *
 * @param value ค่าที่ต้องการเขียนลง Excel
 * @returns ค่าที่ ExcelJS สามารถนำไปเขียนได้
 */
const normalizeExcelValue = (
  value: unknown,
): string | number | boolean | Date | null => {
  if (
    value === undefined ||
    value === null ||
    String(
      value,
    ).trim() === ""
  ) {
    return null;
  }

  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    value instanceof Date
  ) {
    return value;
  }

  return String(
    value,
  );
};

/**
 * จัดกลุ่ม CompareResult ด้วยหมายเลขแถว Test Data
 *
 * ใช้ testDataRowNumber เป็น Key เพราะ
 * - Matching Key ใน Test Data อาจว่าง
 * - Matching Key เดียวกันอาจอยู่ใน Test Data หลายแถว
 *
 * ถ้าใช้ Matching Key เป็น Key โดยตรง
 * ผลของ Test Data แต่ละแถวอาจรวมกันหรือเขียนทับกัน
 *
 * รูปแบบ Map:
 *
 * Map<
 *   Test Data Row Number,
 *   CompareResult[]
 * >
 *
 * @param results ผลการ Compare ทั้งหมด
 * @returns Map ของผลลัพธ์ แยกตามแถว Test Data
 */
const buildResultByTestDataRowMap = (
  results: CompareResult[],
): Map<number, CompareResult[]> => {
  const resultMap =
    new Map<
      number,
      CompareResult[]
    >();

  for (
    const result of results
  ) {
    /**
     * อ่านผลที่มีอยู่แล้วของ Test Data Row นี้
     *
     * ถ้ายังไม่มี ให้ใช้ Array ว่าง
     */
    const groupResults =
      resultMap.get(
        result.testDataRowNumber,
      ) ?? [];

    // เพิ่มผลลัพธ์ปัจจุบันเข้าไปในกลุ่ม
    groupResults.push(
      result,
    );

    // บันทึกกลุ่มกลับเข้า Map
    resultMap.set(
      result.testDataRowNumber,
      groupResults,
    );
  }

  return resultMap;
};

/**
 * สรุปสถานะรวมของ Test Data จำนวน 1 แถว
 *
 * กติกา:
 * - ถ้ามี Field ใด Field หนึ่ง FAIL → ผลรวมเป็น FAIL
 * - ถ้าไม่มี FAIL → ผลรวมเป็น PASS
 *
 * หมายเหตุ:
 * ถ้า groupResults เป็น Array ว่าง
 * ฟังก์ชันนี้จะคืน PASS
 *
 * แต่ Flow ปกติจะตรวจว่า Array มีข้อมูลก่อนเรียกใช้
 */
const getGroupStatus = (
  groupResults: CompareResult[],
): GroupStatus => {
  return groupResults.some(
    result =>
      result.status ===
      "FAIL",
  )
    ? "FAIL"
    : "PASS";
};

/**
 * รวม Remark ของ CompareResult ในกลุ่มเดียวกัน
 *
 * การทำงาน
 * 1. ดึง Remark ของทุกผลลัพธ์
 * 2. ตัดช่องว่างด้านหน้าและด้านหลัง
 * 3. ตัด Remark ที่ว่างออก
 * 4. ตัด Remark ที่ซ้ำกันออกด้วย Set
 * 5. รวมแต่ละ Remark โดยขึ้นบรรทัดใหม่
 *
 * @returns
 * - ข้อความ Remark เมื่อพบสาเหตุ
 * - null เมื่อไม่มี Remark
 */
const getGroupRemark = (
  groupResults: CompareResult[],
): string | null => {
  const remarks = [
    ...new Set(
      groupResults
        .map(
          result =>
            result.remark.trim(),
        )
        .filter(
          remark =>
            remark !== "",
        ),
    ),
  ];

  return remarks.length > 0
    ? remarks.join(
        "\n",
      )
    : null;
};

/**
 * ตรวจสอบว่ากลุ่มผลลัพธ์มาจาก Exclusion Rule หรือไม่
 *
 * รองรับชื่อ Field 2 รูปแบบ:
 * - FTX Exclusion Rule เป็นชื่อจาก Logic เดิม
 * - Exclusion Rule เป็นชื่อที่ Code ปัจจุบันใช้
 */
const isExclusionGroup = (
  groupResults: CompareResult[],
): boolean => {
  return groupResults.some(
    result =>
      result.field ===
        "FTX Exclusion Rule" ||
      result.field ===
        "Exclusion Rule",
  );
};

/**
 * ดึงชื่อ Field ที่มีสถานะ PASS
 *
 * ใช้สำหรับค้นหา Report Column
 * และ Highlight Cell เป็นสีเขียว
 */
const getPassedFields = (
  groupResults: CompareResult[],
): Set<string> => {
  return new Set(
    groupResults
      .filter(
        result =>
          result.status ===
          "PASS",
      )
      .map(
        result =>
          result.field,
      ),
  );
};

/**
 * ดึงชื่อ Field ที่มีสถานะ FAIL
 *
 * ถ้า Field ใน CompareResult เป็น "Matching Key"
 * จะเปลี่ยนเป็นชื่อ Header "Ref. TX No."
 *
 * เพื่อให้สามารถค้นหา Column ใน Report ได้
 */
const getFailedFields = (
  groupResults: CompareResult[],
): Set<string> => {
  return new Set(
    groupResults
      .filter(
        result =>
          result.status ===
          "FAIL",
      )
      .map(
        result =>
          result.field ===
          "Matching Key"
            ? MATCHING_KEY_HEADER
            : result.field,
      ),
  );
};

/**
 * ใส่ Style ให้แถว Header
 *
 * รูปแบบ:
 * - ความสูง 55
 * - ตัวอักษรหนา
 * - ตัวอักษรสีขาว
 * - พื้นหลังสีน้ำเงิน
 * - จัดกึ่งกลาง
 * - Wrap Text
 * - มีเส้นขอบทุกด้าน
 */
const applyHeaderStyle = (
  row: ExcelJS.Row,
): void => {
  row.height =
    55;

  row.eachCell(
    {
      includeEmpty:
        true,
    },
    cell => {
      cell.font = {
        bold:
          true,
        color: {
          argb:
            COLORS.HEADER_TEXT,
        },
      };

      cell.fill = {
        type:
          "pattern",
        pattern:
          "solid",
        fgColor: {
          argb:
            COLORS.REPORT_HEADER,
        },
      };

      cell.alignment = {
        horizontal:
          "center",
        vertical:
          "middle",
        wrapText:
          true,
      };

      cell.border =
        THIN_BORDER;
    },
  );
};

/**
 * ใส่ Style พื้นฐานให้แถวข้อมูล
 *
 * รูปแบบ:
 * - ความสูง 22
 * - มีเส้นขอบทุกด้าน
 * - ข้อความชิดซ้าย
 * - จัดกึ่งกลางแนวตั้ง
 * - Wrap Text
 */
const applyDataRowStyle = (
  row: ExcelJS.Row,
): void => {
  row.height =
    22;

  row.eachCell(
    {
      includeEmpty:
        true,
    },
    cell => {
      cell.border =
        THIN_BORDER;

      cell.alignment = {
        horizontal:
          "left",
        vertical:
          "middle",
        wrapText:
          true,
      };
    },
  );
};

/**
 * ใส่ Style สีเขียวให้ Cell ที่ผ่าน
 */
const applyPassStyle = (
  cell: ExcelJS.Cell,
): void => {
  cell.fill = {
    type:
      "pattern",
    pattern:
      "solid",
    fgColor: {
      argb:
        COLORS.PASS_FILL,
    },
  };

  cell.font = {
    color: {
      argb:
        COLORS.PASS_TEXT,
    },
  };
};

/**
 * ใส่ Style สีแดงให้ Cell ที่ไม่ผ่าน
 */
const applyFailStyle = (
  cell: ExcelJS.Cell,
): void => {
  cell.fill = {
    type:
      "pattern",
    pattern:
      "solid",
    fgColor: {
      argb:
        COLORS.FAIL_FILL,
    },
  };

  cell.font = {
    color: {
      argb:
        COLORS.FAIL_TEXT,
    },
  };
};

/**
 * ใส่ Style ให้ Cell สถานะ Result
 *
 * PASS:
 * - พื้นหลังสีเขียว
 * - ตัวอักษรสีเขียวเข้ม
 *
 * FAIL:
 * - พื้นหลังสีแดง
 * - ตัวอักษรสีแดงเข้ม
 *
 * ทั้งสองสถานะ:
 * - ตัวอักษรหนา
 * - จัดกึ่งกลาง
 */
const applyStatusStyle = (
  cell: ExcelJS.Cell,
  status: GroupStatus,
): void => {
  if (
    status === "PASS"
  ) {
    applyPassStyle(
      cell,
    );
  } else {
    applyFailStyle(
      cell,
    );
  }

  cell.font = {
    ...cell.font,
    bold:
      true,
  };

  cell.alignment = {
    horizontal:
      "center",
    vertical:
      "middle",
  };
};

/**
 * ใส่สีให้ Output Row ที่มีผลการ Compare
 *
 * Column สำคัญ:
 * A = Test Script No.
 * B = Result
 * C = Remark
 * D เป็นต้นไป = Header ของ Report
 *
 * กติกา:
 * - Test Script No. ใช้สีตามสถานะรวม
 * - Result ใช้สีตามสถานะรวม
 * - Remark ของ FAIL ใช้สีแดง
 * - Remark ของ Exclusion PASS ใช้สีเขียว
 * - Matching Key ที่พบใน Reportใช้สีเขียว
 * - Core Field PASS ใช้สีเขียว
 * - Core Field FAIL ใช้สีแดง
 */
const applyCaseStyle = (
  row: ExcelJS.Row,
  status: GroupStatus,
  remark: string | null,
  groupResults: CompareResult[],
  reportHeaders: string[],
): void => {
  /**
   * Column A: Test Script No.
   */
  if (
    status === "PASS"
  ) {
    applyPassStyle(
      row.getCell(
        1,
      ),
    );
  } else {
    applyFailStyle(
      row.getCell(
        1,
      ),
    );
  }

  /**
   * Column B: Result
   */
  applyStatusStyle(
    row.getCell(
      2,
    ),
    status,
  );

  /**
   * Column C: Remark
   *
   * - FAIL ใช้สีแดง
   * - Exclusion PASS ใช้สีเขียว
   * - PASS ปกติไม่มี Remark จึงไม่ใส่สี
   */
  if (
    remark
  ) {
    if (
      status === "FAIL"
    ) {
      applyFailStyle(
        row.getCell(
          3,
        ),
      );
    } else if (
      isExclusionGroup(
        groupResults,
      )
    ) {
      applyPassStyle(
        row.getCell(
          3,
        ),
      );
    }

    /**
     * จัด Remark ชิดซ้ายและด้านบน
     */
    row.getCell(
      3,
    ).alignment = {
      horizontal:
        "left",
      vertical:
        "top",
      wrapText:
        true,
    };

    /**
     * เพิ่มความสูงตามจำนวนบรรทัดของ Remark
     *
     * อย่างน้อย 22
     * และเพิ่มประมาณ 30 ต่อหนึ่งบรรทัด
     */
    const remarkLineCount =
      remark.split(
        "\n",
      ).length;

    row.height =
      Math.max(
        22,
        remarkLineCount * 30,
      );
  }

  /**
   * หา Column Matching Key ใน Report Header
   */
  const matchingKeyHeaderIndex =
    reportHeaders.indexOf(
      MATCHING_KEY_HEADER,
    );

  /**
   * ถ้าพบ Matching Key Header
   * ให้ใส่สีเขียวก่อน
   *
   * Column Report เริ่มหลัง 3 Column ที่เพิ่มมา
   * จึงใช้ Array Index + 4
   */
  if (
    matchingKeyHeaderIndex !== -1
  ) {
    applyPassStyle(
      row.getCell(
        matchingKeyHeaderIndex +
        4,
      ),
    );
  }

  /**
   * Highlight Field ที่ PASS เป็นสีเขียว
   */
  for (
    const passedField of
    getPassedFields(
      groupResults,
    )
  ) {
    const reportHeaderIndex =
      reportHeaders.indexOf(
        passedField,
      );

    /**
     * ถ้าชื่อ Field ไม่มีใน Report Header
     * ให้ข้าม เช่น Exclusion Rule
     */
    if (
      reportHeaderIndex === -1
    ) {
      continue;
    }

    applyPassStyle(
      row.getCell(
        reportHeaderIndex +
        4,
      ),
    );
  }

  /**
   * Highlight Field ที่ FAIL เป็นสีแดง
   *
   * ทำหลังการใส่สี PASS
   * เพื่อให้สีแดงสามารถทับสีเขียวได้
   *
   * ตัวอย่าง:
   * ถ้า Matching Key FAIL
   * Cell Ref. TX No. จะถูกเปลี่ยนจากเขียวเป็นแดง
   */
  for (
    const failedField of
    getFailedFields(
      groupResults,
    )
  ) {
    const reportHeaderIndex =
      reportHeaders.indexOf(
        failedField,
      );

    if (
      reportHeaderIndex === -1
    ) {
      continue;
    }

    applyFailStyle(
      row.getCell(
        reportHeaderIndex +
        4,
      ),
    );
  }
};

/**
 * กำหนดความกว้างของแต่ละ Column
 *
 * Column A = 18
 * Column B = 12
 * Column C = 72
 *
 * Report Column เริ่มต้นที่ Column D
 * และเลือกความกว้างตามชื่อ Header
 */
const applyColumnWidths = (
  worksheet: ExcelJS.Worksheet,
  outputHeaders: string[],
): void => {
  // Test Script No.
  worksheet.getColumn(
    1,
  ).width =
    18;

  // Result
  worksheet.getColumn(
    2,
  ).width =
    12;

  // Remark
  worksheet.getColumn(
    3,
  ).width =
    72;

  /**
   * กำหนดความกว้างของ Report Column
   */
  for (
    let columnNumber = 4;
    columnNumber <=
      outputHeaders.length;
    columnNumber += 1
  ) {
    const header =
      outputHeaders[
        columnNumber - 1
      ];

    // ความกว้างเริ่มต้น
    let width =
      20;

    if (
      header ===
      MATCHING_KEY_HEADER
    ) {
      // Matching Key ใช้ความกว้าง 45
      width =
        45;
    } else if (
      header.includes(
        "Description",
      ) ||
      header.includes(
        "Name",
      )
    ) {
      width =
        30;
    } else if (
      header.includes(
        "Account Number",
      )
    ) {
      width =
        28;
    } else if (
      header.includes(
        "Amount",
      ) ||
      header.includes(
        "Exchange Rate",
      )
    ) {
      width =
        22;
    } else if (
      header.length > 30
    ) {
      width =
        28;
    } else if (
      header.length > 20
    ) {
      width =
        24;
    }

    worksheet.getColumn(
      columnNumber,
    ).width =
      width;
  }
};

/**
 * สร้างและบันทึกไฟล์ผลลัพธ์ DS_FTX
 *
 * ขั้นตอน
 * 1. สร้าง Workbook และ Sheet "DS_FTX"
 * 2. สร้าง Output Header
 * 3. จัดกลุ่ม CompareResult ตาม Test Data Row
 * 4. จัดกลุ่ม ExpectedRow ตาม Matching Key
 * 5. จัดลำดับ ActualRow
 * 6. เขียนข้อมูล Report ทุกแถว
 * 7. เพิ่ม ExpectedRow ที่ไม่มีใน Report
 * 8. Freeze Header และ 3 Column แรก
 * 9. เพิ่ม Auto Filter
 * 10. กำหนดความกว้าง Column
 * 11. บันทึกไฟล์ Excel
 *
 * @param results ผลการเปรียบเทียบทั้งหมด
 * @param expectedRows ข้อมูลจาก Test Data
 * @param actualRows ข้อมูลจาก Report DS_FTX
 * @param reportHeaders Header ทั้งหมดของ Report
 * @param outputFilePath Path ของไฟล์ผลลัพธ์
 */
export const writeCompareResult = async (
  results: CompareResult[],
  expectedRows: ExpectedRow[],
  actualRows: ActualRow[],
  reportHeaders: string[],
  outputFilePath: string,
): Promise<void> => {
  /**
   * สร้าง Workbook ใหม่
   */
  const workbook =
    new ExcelJS.Workbook();

  /**
   * กำหนด Metadata ของไฟล์
   */
  workbook.creator =
    "AF1 Script 3";

  workbook.created =
    new Date();

  /**
   * สร้าง Worksheet ชื่อ DS_FTX
   */
  const worksheet =
    workbook.addWorksheet(
      "DS_FTX",
    );

  /**
   * สร้าง Header ของไฟล์ผลลัพธ์
   *
   * 3 Column แรกเป็นข้อมูลผลการตรวจสอบ
   * หลังจากนั้นเป็น Header เดิมของ Report
   */
  const outputHeaders = [
    "Test Script No.",
    "Result",
    "Remark",
    ...reportHeaders,
  ];

  /**
   * จัดกลุ่ม CompareResult ตามหมายเลขแถว Test Data
   */
  const resultByTestDataRowMap =
    buildResultByTestDataRowMap(
      results,
    );

  /**
   * จัดกลุ่ม ExpectedRow ตาม Matching Key
   *
   * Matching Key หนึ่งค่าอาจอยู่ใน Test Data หลายแถว
   * จึงต้องเก็บ ExpectedRow เป็น Array
   *
   * ExpectedRow ที่ Matching Key ว่าง
   * จะไม่ถูกเพิ่มเข้า Map นี้
   */
  const expectedRowsByMatchingKey =
    new Map<
      string,
      ExpectedRow[]
    >();

  for (
    const expectedRow of expectedRows
  ) {
    if (
      expectedRow.matchingKey !==
      ""
    ) {
      const matchedExpectedRows =
        expectedRowsByMatchingKey.get(
          expectedRow.matchingKey,
        ) ?? [];

      matchedExpectedRows.push(
        expectedRow,
      );

      expectedRowsByMatchingKey.set(
        expectedRow.matchingKey,
        matchedExpectedRows,
      );
    }
  }

  /**
   * สร้าง Set ของ Matching Key ทั้งหมดที่พบใน Report
   *
   * ใช้ตรวจว่า ExpectedRow ใดไม่มีใน Report
   *
   * หมายเหตุ:
   * Set นี้รวม Matching Key ว่างด้วย
   * ถ้ามี ActualRow ที่ Key ว่าง
   */
  const actualKeySet =
    new Set(
      actualRows.map(
        actualRow =>
          actualRow.matchingKey,
      ),
    );

  /**
   * จัดลำดับ ActualRow สำหรับเขียนลงไฟล์
   *
   * ลำดับ:
   * 1. PASS
   * 2. Report Row ที่ Map กับ Test Data ไม่เจอ
   * 3. FAIL
   *
   * การ Sort ใช้สำเนาของ actualRows
   * จึงไม่เปลี่ยนลำดับ Array ต้นฉบับ
   */
  const sortedActualRows =
    [...actualRows].sort(
      (
        firstRow,
        secondRow,
      ) => {
        /**
         * คืนค่าลำดับของ ActualRow
         *
         * 1 = PASS
         * 2 = Map กับ Test Data ไม่เจอ
         * 3 = FAIL
         */
        const getSortOrder = (
          actualRow: ActualRow,
        ): number => {
          /**
           * ค้นหา ExpectedRow ทั้งหมด
           * ที่ใช้ Matching Key เดียวกัน
           */
          const matchedExpectedRows =
            expectedRowsByMatchingKey.get(
              actualRow.matchingKey,
            ) ?? [];

          /**
           * รวม CompareResult ของ ExpectedRow ทุกแถว
           * ที่ใช้ Matching Key นี้
           */
          const groupResults =
            matchedExpectedRows.flatMap(
              expectedRow =>
                resultByTestDataRowMap.get(
                  expectedRow.rowNumber,
                ) ?? [],
            );

          /**
           * ถ้าไม่มี CompareResult
           * ให้จัดเป็น Report Row ที่ Map ไม่เจอ
           */
          if (
            !groupResults ||
            groupResults.length === 0
          ) {
            return 2;
          }

          /**
           * ถ้ามี FAIL อย่างน้อย 1 รายการ
           * จะได้ลำดับ 3
           *
           * ถ้าไม่มี FAIL
           * จะได้ลำดับ 1
           */
          return getGroupStatus(
            groupResults,
          ) === "PASS"
            ? 1
            : 3;
        };

        return (
          getSortOrder(
            firstRow,
          ) -
          getSortOrder(
            secondRow,
          )
        );
      },
    );

  /**
   * เพิ่ม Header ลงในแถวที่ 1
   */
  const headerRow =
    worksheet.addRow(
      outputHeaders,
    );

  applyHeaderStyle(
    headerRow,
  );

  /**
   * เขียนข้อมูล Report DS_FTX ทุกแถว
   *
   * Output Row เริ่มจากแถวที่ 2
   */
  for (
    const actualRow of
    sortedActualRows
  ) {
    /**
     * ค้นหา ExpectedRow ที่ใช้ Matching Key เดียวกัน
     */
    const matchedExpectedRows =
      expectedRowsByMatchingKey.get(
        actualRow.matchingKey,
      ) ?? [];

    /**
     * กรณี Report Row ไม่มี Test Data จับคู่
     *
     * จะเขียนข้อมูล Report ตามปกติ
     * แต่ 3 Column แรกจะว่าง:
     * - Test Script No.
     * - Result
     * - Remark
     */
    if (
      matchedExpectedRows.length === 0
    ) {
      const row =
        worksheet.addRow([
          null,
          null,
          null,

          /**
           * เขียนข้อมูลตาม Report Header
           */
          ...reportHeaders.map(
            header =>
              normalizeExcelValue(
                actualRow.data[
                  header
                ],
              ),
          ),
        ]);

      /**
       * ใส่เฉพาะ Style พื้นฐาน
       * ไม่ Highlight PASS หรือ FAIL
       */
      applyDataRowStyle(
        row,
      );

      continue;
    }

    /**
     * Matching Key เดียวกันอาจอยู่ใน Test Data หลายแถว
     *
     * จึงเขียน Output Row แยก
     * หนึ่งแถวต่อ ExpectedRow
     *
     * ถ้า Matching Key ซ้ำใน Report ด้วย
     * Loop ภายนอกจะทำงานกับ ActualRow ทุกแถว
     * ทำให้ ExpectedRow เดิมอาจถูกเขียนซ้ำหลาย Output Row
     */
    for (
      const expectedRow of
      matchedExpectedRows
    ) {
      /**
       * อ่าน CompareResult ของ Test Data Row ปัจจุบัน
       */
      const groupResults =
        resultByTestDataRowMap.get(
          expectedRow.rowNumber,
        ) ?? [];

      /**
       * ถ้าไม่มี CompareResult
       * จะข้าม ExpectedRow นี้
       *
       * ตาม Flow ปกติ ftx-reconcile.ts
       * ควรสร้างผลอย่างน้อย 1 รายการ
       */
      if (
        groupResults.length === 0
      ) {
        continue;
      }

      /**
       * สรุปสถานะรวมของ Test Data Row
       */
      const status =
        getGroupStatus(
          groupResults,
        );

      /**
       * รวม Remark ของ Test Data Row
       */
      const remark =
        getGroupRemark(
          groupResults,
        );

      /**
       * สร้าง Output Row
       *
       * Test Script No.:
       * ใช้จาก ExpectedRow ก่อน
       * ถ้าว่างจึงใช้จาก CompareResult แรก
       *
       * หลัง 3 Column แรก:
       * เขียนข้อมูล ActualRow ตาม Report Header
       */
      const row =
        worksheet.addRow([
          expectedRow.testScriptNo ||
            groupResults[0]
              ?.testScriptNo ||
            null,

          status,

          remark,

          ...reportHeaders.map(
            header =>
              normalizeExcelValue(
                actualRow.data[
                  header
                ],
              ),
          ),
        ]);

      applyDataRowStyle(
        row,
      );

      /**
       * ใส่สีตามผลการ Compare
       */
      applyCaseStyle(
        row,
        status,
        remark,
        groupResults,
        reportHeaders,
      );
    }
  }

  /**
   * เพิ่ม ExpectedRow ที่ไม่มี Matching Key ใน Report
   *
   * รวมถึง
   * - Matching Key ว่าง
   * - Matching Key ที่ค้นหาใน Report ไม่พบ
   * - Exclusion ที่ไม่มี Matching Key อยู่ใน Report
   */
  for (
    const expectedRow of expectedRows
  ) {
    /**
     * ถ้า Matching Key ไม่ว่าง
     * และมี Key นี้อยู่ใน Report
     * แสดงว่าเขียนผ่าน Loop ActualRow ไปแล้ว
     */
    if (
      expectedRow.matchingKey !==
        "" &&
      actualKeySet.has(
        expectedRow.matchingKey,
      )
    ) {
      continue;
    }

    const groupResults =
      resultByTestDataRowMap.get(
        expectedRow.rowNumber,
      ) ??
      [];

    /**
     * ถ้าไม่มี CompareResult ให้ข้าม
     */
    if (
      groupResults.length === 0
    ) {
      continue;
    }

    const status =
      getGroupStatus(
        groupResults,
      );

    const remark =
      getGroupRemark(
        groupResults,
      );

    /**
     * กำหนด Matching Key ที่จะแสดงใน Output
     *
     * ใช้ ExpectedRow ก่อน
     * ถ้าว่างจึงลองใช้ CompareResult แรก
     */
    const expectedMatchingKey =
      expectedRow.matchingKey !==
      ""
        ? expectedRow.matchingKey
        : groupResults[0]
            ?.matchingKey ??
          "";

    /**
     * สร้างข้อมูล 3 Column แรก
     */
    const rowValues: Array<
      string |
      number |
      boolean |
      Date |
      null
    > = [
      expectedRow.testScriptNo ||
        groupResults[0]
          ?.testScriptNo ||
        null,

      status,

      remark,
    ];

    /**
     * สร้างข้อมูล Report Column
     *
     * กรณีนี้ไม่มี ActualRow
     * จึงใส่เฉพาะ Expected Matching Key
     * ลงใน Column "Ref. TX No."
     *
     * Report Field อื่นจะเป็น Cell ว่าง
     */
    for (
      const reportHeader of
      reportHeaders
    ) {
      if (
        reportHeader ===
        MATCHING_KEY_HEADER
      ) {
        rowValues.push(
          expectedMatchingKey ||
          null,
        );

        continue;
      }

      /**
       * ไม่ใส่ Expected Value ลงใน Report Field
       * เพราะรายการนี้ไม่มี Actual Data จาก DS_FTX
       */
      rowValues.push(
        null,
      );
    }

    /**
     * เพิ่ม Expected Row ที่ไม่มี Actual Data ลงใน Sheet
     */
    const row =
      worksheet.addRow(
        rowValues,
      );

    applyDataRowStyle(
      row,
    );

    /**
     * ใส่สี Column A และ C ตามสถานะ
     */
    if (
      status === "PASS"
    ) {
      applyPassStyle(
        row.getCell(
          1,
        ),
      );

      applyPassStyle(
        row.getCell(
          3,
        ),
      );
    } else {
      applyFailStyle(
        row.getCell(
          1,
        ),
      );

      applyFailStyle(
        row.getCell(
          3,
        ),
      );
    }

    /**
     * ใส่สี Column B: Result
     */
    applyStatusStyle(
      row.getCell(
        2,
      ),
      status,
    );

    /**
     * จัดตำแหน่ง Remark
     */
    row.getCell(
      3,
    ).alignment = {
      horizontal:
        "left",
      vertical:
        "top",
      wrapText:
        true,
    };

    /**
     * กำหนดความสูงแถวเป็น 40
     */
    row.height =
      40;

    /**
     * ค้นหา Column Matching Key
     */
    const matchingKeyHeaderIndex =
      reportHeaders.indexOf(
        MATCHING_KEY_HEADER,
      );

    if (
      matchingKeyHeaderIndex !== -1
    ) {
      const matchingKeyCell =
        row.getCell(
          matchingKeyHeaderIndex +
          4,
        );

      /**
       * Matching Key ใช้สีตามสถานะรวม
       *
       * PASS เช่น Exclusion
       * → สีเขียว
       *
       * FAIL เช่นไม่พบ Matching Key
       * → สีแดง
       */
      if (
        status === "PASS"
      ) {
        applyPassStyle(
          matchingKeyCell,
        );
      } else {
        applyFailStyle(
          matchingKeyCell,
        );
      }
    }
  }

  /**
   * Freeze Pane
   *
   * ySplit: 1
   * = ตรึง Header แถวที่ 1
   *
   * xSplit: 3
   * = ตรึง 3 Column แรก
   */
  worksheet.views = [
    {
      state:
        "frozen",
      ySplit:
        1,
      xSplit:
        3,
    },
  ];

  /**
   * เพิ่ม Auto Filter ให้ Header ทุก Column
   */
  worksheet.autoFilter = {
    from: {
      row:
        1,
      column:
        1,
    },
    to: {
      row:
        1,
      column:
        outputHeaders.length,
    },
  };

  /**
   * กำหนดความกว้างของ Column
   */
  applyColumnWidths(
    worksheet,
    outputHeaders,
  );

  /**
   * บันทึก Workbook เป็นไฟล์ .xlsx
   *
   * Folder ปลายทางต้องถูกสร้างไว้ก่อน
   * เพราะฟังก์ชันนี้ไม่ได้สร้าง Folder
   */
  await workbook.xlsx.writeFile(
    outputFilePath,
  );
};