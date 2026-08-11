/**
 * ============================================================================
 * ptx-result-writer.ts
 * ----------------------------------------------------------------------------
 * Export Compare Result เป็น Excel
 *
 * รูปแบบ Output อ้างอิงจาก
 * DS_PTX reconcile Test Data mapping field.xlsx
 *
 * โครงสร้าง
 * - Row 1 : Mapping กับ Test Data
 * - Row 2 : Header ของผลลัพธ์
 * - Row 3 เป็นต้นไป : ข้อมูล Compare
 * - 1 Expected Row จาก Test Data = 1 Row
 * - Matching Key ซ้ำกันได้ โดยแต่ละ Test Data Row ต้องแสดงแยกกัน
 * - เพิ่มคอลัมน์ Remark เพื่ออธิบายสาเหตุของ Case ที่ FAIL
 * ============================================================================
 */

import ExcelJS from "exceljs";

import {
    ActualRow,
    CompareResult,
    ExpectedRow,
    GroupedCompareResult,
} from "./ptx-types";


/**
 * ============================================================================
 * จัดกลุ่มผลการเปรียบเทียบ
 * ----------------------------------------------------------------------------
 * แปลง CompareResult[] เป็น GroupedCompareResult[]
 *
 * 1 Test Data Row + 1 Matching Key = 1 Output Row
 * ============================================================================
 */
const buildExpectedIdentity = (
    testDataRowNumber: number,
    matchingKey: string,
): string => {

    return (
        `${testDataRowNumber}` +
        "\u0000" +
        matchingKey
    );

};

const transformCompareResults = (
    results: CompareResult[],
): GroupedCompareResult[] => {

    const resultMap =
        new Map<string, GroupedCompareResult>();

    for (
        const result of results
    ) {

        const expectedIdentity =
            buildExpectedIdentity(
                result.testDataRowNumber,
                result.matchingKey,
            );

        let group =
            resultMap.get(
                expectedIdentity,
            );

        if (!group) {

            group = {

                matchingKey:
                    result.matchingKey,

                testDataRowNumber:
                    result.testDataRowNumber,

                reportRowNumber:
                    result.reportRowNumber,

                fields: {},

            };

            resultMap.set(
                expectedIdentity,
                group,
            );

        }

        group.fields[
            result.field
        ] = result;

    }

    return Array.from(
        resultMap.values(),
    );

};

/**
 * ============================================================================
 * สี
 * ============================================================================
 */
const COLORS = {

    MAPPING_HEADER:
        "FFFFE699",

    REPORT_HEADER:
        "FF4472C4",

    HEADER_TEXT:
        "FFFFFFFF",

    PASS_FILL:
        "FFC6EFCE",

    PASS_TEXT:
        "FF006100",

    FAIL_FILL:
        "FFFFC7CE",

    FAIL_TEXT:
        "FF9C0006",

    WRONG_FILL:
        "FFFFE699",

    WRONG_TEXT:
        "FF7F6000",

    SKIP_FILL:
        "FFFFE699",

    SKIP_TEXT:
        "FF7F6000",

    BORDER:
        "FFB7B7B7",

    WHITE:
        "FFFFFFFF",

};

/**
 * ============================================================================
 * Header ของ DS_PTX
 * ============================================================================
 *
 * ต้องเรียงให้ตรงกับไฟล์ต้นแบบ
 */
const DS_PTX_REPORT_HEADERS: string[] = [

    "Dept Code",

    "System Id",

    "Reference Transaction Number",

    "Cust Code",

    "CMF CODE",

    "Cust Name",

    "Data Provider Branch Number",

    "Data Provider IBF Indicator",

    "Data Set Date",

    "Receive Payment Transaction Type",

    "Receive Payment Item Type",

    "Receive Payment Item Description",

    "Receive Payment Transaction Date",

    "Involved Party Id",

    "Involved Party Name",

    "Country Id of Involved Party",

    "Payment Method",

    "Currency Id",

    "Transaction Amount in Foreign Currency",

    "Debt Instrument Type",

    "ISIN Code",

    "Debt Instrument Name",

    "Issuer or Invested Organization Name",

    "Country Id of Issuer or Invested Organization",

    "Issue Date",

    "Maturity Date",

    "Original Term",

    "Original Term Unit",

    "Coupon Rate",

    "Intention Country Id",

    "Unit of Transaction",

    "Sell Foreign Currency Security Transaction Amount in Baht Equivalent",

    "Defaulted Bill Purchase Date",

];

/**
 * ============================================================================
 * Header แถวที่ 2
 * ============================================================================
 */
const OUTPUT_HEADERS: Array<string | null> = [

    "Test Script No.",

    "Result",

    "Remark",

    ...DS_PTX_REPORT_HEADERS,

];


/**
 * ============================================================================
 * Mapping แถวที่ 1
 * ============================================================================
 *
 * จำนวนคอลัมน์ต้องตรงกับ OUTPUT_HEADERS
 */
const TEST_DATA_MAPPING_HEADERS: Array<string | null> = [

    "#",

    "1:Test No.",

    null,

    null,

    null,

    null,

    "5:Transaction ID/ Reconcile ID",

    null,

    null,

    null,

    null,

    null,

    "3:Txn Date",

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    "61:Fee Currency 1\n70:Fee Currency 2",

    "61:Fee Amount 1\n62:Fee Amount 2\n78:Fee Amount 3",

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

    null,

];

/**
 * ============================================================================
 * ตรวจว่า Report Field มี Mapping กับ Test Data หรือไม่
 * ============================================================================
 *
 * TEST_DATA_MAPPING_HEADERS ยังคงเก็บ Mapping ภายในไว้
 * เพื่อใช้ตัดสินสีของ Field
 *
 * ใน Mapping Config เดิม Report Header เริ่มที่ Index 4
 * ดังนั้นตำแหน่ง Mapping ของ Report Field
 * จะเท่ากับ reportHeaderIndex + 4
 */
const hasTestDataMapping = (
    reportField: string,
): boolean => {

    const reportHeaderIndex =
        DS_PTX_REPORT_HEADERS.indexOf(
            reportField,
        );

    if (
        reportHeaderIndex === -1
    ) {

        return false;

    }

    const mappingHeaderIndex =
        reportHeaderIndex + 4;

    const mappingValue =
        TEST_DATA_MAPPING_HEADERS[
        mappingHeaderIndex
        ];

    return (
        mappingValue !== null &&
        mappingValue !== undefined &&
        String(
            mappingValue,
        ).trim() !== ""
    );

};


/**
 * ============================================================================
 * รูปแบบเส้นขอบของ Cell
 * ============================================================================
 */
const THIN_BORDER: Partial<ExcelJS.Borders> = {

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
 * ============================================================================
 * แปลงค่าเพื่อเขียนลง Excel
 * ============================================================================
 */
const normalizeExcelValue = (
    value: unknown,
): string | number | boolean | null => {

    /**
     * ให้ ExcelJS สร้าง Cell ว่างจริง
     *
     * ไม่ใช้ "" เพราะบางกรณีอาจถูกบันทึก
     * เป็น Shared String Index ใน Excel
     */
    if (
        value === undefined ||
        value === null ||
        String(value).trim() === ""
    ) {

        return null;

    }

    if (
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean"
    ) {

        return value;

    }

    return String(
        value,
    );

};

/**
 * ============================================================================
 * หาผลรวมของ Matching Key
 * ============================================================================
 */
const getGroupStatus = (
    group: GroupedCompareResult,
): "PASS" | "FAIL" | "SKIP" => {

    const fieldResults =
        Object.values(
            group.fields,
        );

    if (
        fieldResults.some(
            result =>
                result.status === "FAIL",
        )
    ) {

        return "FAIL";

    }

    if (
        fieldResults.length > 0 &&
        fieldResults.every(
            result =>
                result.status === "SKIP",
        )
    ) {

        return "SKIP";

    }

    return "PASS";

};

/**
 * ============================================================================
 * หา Field ที่ PASS และมีการ Compare จริง
 * ============================================================================
 *
 * ไม่รวม Field ที่ถูกข้ามด้วยเงื่อนไข เช่น
 * Field ที่ข้ามการตรวจเพราะข้อมูลจริงเป็นค่าว่าง
 */
const getPassedFields = (
    group: GroupedCompareResult,
): Set<string> => {

    const passedFields =
        new Set<string>();

    for (
        const result of Object.values(
            group.fields,
        )
    ) {

        if (
            result.status !== "PASS"
        ) {

            continue;

        }

        /**
         * Conditional Field ที่ไม่ได้ Compare จริง
         * ไม่ต้อง Highlight สีเขียว
         */
        if (
            result.remark
                .trim()
                .toLowerCase()
                .startsWith(
                    "skipped",
                )
        ) {

            continue;

        }

        passedFields.add(
            result.field,
        );

    }

    return passedFields;

};

/**
 * ============================================================================
 * หา Field ที่ FAIL
 * ============================================================================
 */
const getFailedFields = (
    group: GroupedCompareResult,
): Set<string> => {

    const failedFields =
        new Set<string>();

    for (
        const result of Object.values(
            group.fields,
        )
    ) {

        if (
            result.status !== "FAIL"
        ) {

            continue;

        }

        /**
         * Matching Key ใน CompareResult
         * ตรงกับ Reference Transaction Number ใน Output
         */
        if (
            result.field === "Matching Key"
        ) {

            failedFields.add(
                "Reference Transaction Number",
            );

            continue;

        }

        failedFields.add(
            result.field,
        );

    }

    return failedFields;

};


/**
 * ============================================================================
 * แปลงค่าให้อ่านง่ายใน Remark
 * ============================================================================
 */
const formatRemarkValue = (
    value: unknown,
): string => {

    if (
        value === undefined ||
        value === null ||
        String(value).trim() === ""
    ) {

        return "(blank)";

    }

    return String(
        value,
    ).trim();

};

/**
 * ============================================================================
 * สร้างข้อความ Remark สำหรับ Case ที่ FAIL
 * ============================================================================
 *
 * หนึ่ง Case อาจมีหลาย Field ที่ผิด
 * จึงรวมสาเหตุทั้งหมดไว้ใน Cell เดียวและขึ้นบรรทัดใหม่
 */
const getGroupRemark = (
    group: GroupedCompareResult,
): string | null => {

    const failedResults =
        Object.values(
            group.fields,
        ).filter(
            result =>
                result.status === "FAIL",
        );

    if (
        failedResults.length === 0
    ) {

        return null;

    }

    return failedResults
        .map(
            result => {

                const reason =
                    result.remark.trim() !== ""
                        ? result.remark.trim()
                        : "Validation Failed";

                return (
                    `${result.field}: ${reason}` +
                    ` | [TS]: ${formatRemarkValue(result.expected)}` +
                    ` | [DS-PTX]: ${formatRemarkValue(result.actual)}`
                );

            },
        )
        .join(
            "\n",
        );

};

/**
 * ============================================================================
 * จัดรูปแบบ Header ของ Report
 * ============================================================================
 */
const applyReportHeaderStyle = (
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
 * ============================================================================
 * จัดรูปแบบแถวข้อมูล
 * ============================================================================
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
 * ============================================================================
 * จัดรูปแบบช่องผลการตรวจ
 * ============================================================================
 */
const applyResultStyle = (
    cell: ExcelJS.Cell,
    status: "PASS" | "FAIL" | "SKIP",
): void => {

    let fillColor =
        COLORS.PASS_FILL;

    let textColor =
        COLORS.PASS_TEXT;

    if (
        status === "FAIL"
    ) {

        fillColor =
            COLORS.FAIL_FILL;

        textColor =
            COLORS.FAIL_TEXT;

    } else if (
        status === "SKIP"
    ) {

        fillColor =
            COLORS.SKIP_FILL;

        textColor =
            COLORS.SKIP_TEXT;

    }

    cell.font = {

        bold:
            true,

        color: {

            argb:
                textColor,

        },

    };

    cell.fill = {

        type:
            "pattern",

        pattern:
            "solid",

        fgColor: {

            argb:
                fillColor,

        },

    };

    cell.alignment = {

        horizontal:
            "center",

        vertical:
            "middle",

    };

};

/**
 * ============================================================================
 * Style Field ที่ PASS
 * ============================================================================
 */
const applyPassedFieldStyle = (
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
 * ============================================================================
 * Highlight Test Script No. ตามผลระดับ Case
 * ============================================================================
 *
 * PASS = เขียวอ่อน
 * FAIL = แดงอ่อน
 * SKIP = เหลืองอ่อน
 */
const applyTestScriptNoStyle = (
    cell: ExcelJS.Cell,
    status: "PASS" | "FAIL" | "SKIP",
): void => {

    let fillColor =
        COLORS.PASS_FILL;

    let textColor =
        COLORS.PASS_TEXT;

    if (
        status === "FAIL"
    ) {

        fillColor =
            COLORS.FAIL_FILL;

        textColor =
            COLORS.FAIL_TEXT;

    } else if (
        status === "SKIP"
    ) {

        fillColor =
            COLORS.SKIP_FILL;

        textColor =
            COLORS.SKIP_TEXT;

    }

    cell.fill = {

        type:
            "pattern",

        pattern:
            "solid",

        fgColor: {

            argb:
                fillColor,

        },

    };

    cell.font = {

        color: {

            argb:
                textColor,

        },

    };

};

/**
 * ============================================================================
 * Style Field ที่ข้อมูลผิดจริง
 * ============================================================================
 *
 * ใช้สีเหลืองอ่อน เพื่อแยกจาก Field อื่น
 * ที่มีการ Compare ใน Case FAIL
 */
const applyFailedFieldStyle = (
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
 * ============================================================================
 * กำหนดความกว้าง Column
 * ============================================================================
 */
const applyColumnWidths = (
    worksheet: ExcelJS.Worksheet,
): void => {

    /**
     * A: Test Script No.
     */
    worksheet.getColumn(1).width =
        18;

    /**
     * B: Result
     */
    worksheet.getColumn(2).width =
        12;

    /**
     * C: Remark
     */
    worksheet.getColumn(3).width =
        60;

    /**
     * ตั้งแต่ Dept Code เป็นต้นไป
     */
    for (
        let columnIndex = 4;
        columnIndex <= OUTPUT_HEADERS.length;
        columnIndex += 1
    ) {

        const header =
            OUTPUT_HEADERS[
            columnIndex - 1
            ] ?? "";

        let width =
            18;
        if (
            header === "Reference Transaction Number"
        ) {

            width =
                52;

        } else if (
            header ===
            "Transaction Amount in Foreign Currency"
        ) {

            width =
                25;

        } else if (
            header ===
            "Issuer or Invested Organization Name"
        ) {

            width =
                34;

        } else if (
            header ===
            "Country Id of Issuer or Invested Organization"
        ) {

            width =
                30;

        } else if (
            header ===
            "Sell Foreign Currency Security Transaction Amount in Baht Equivalent"
        ) {

            width =
                38;

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

        worksheet
            .getColumn(
                columnIndex,
            )
            .width =
            width;

    }

};

/**
 * ============================================================================
 * สร้างและบันทึกไฟล์ Compare Result
 * ============================================================================
 */
export const writeCompareResult = async (
    results: CompareResult[],
    expectedRows: ExpectedRow[],
    actualRows: ActualRow[],
    outputFile: string,
): Promise<void> => {

    const workbook =
        new ExcelJS.Workbook();

    workbook.creator =
        "AF1 Script 3";

    workbook.created =
        new Date();

    const worksheet =
        workbook.addWorksheet(
            "DS_PTX",
        );

    /**
     * --------------------------------------------------------------------------
     * สร้าง Map สำหรับค้นหาข้อมูล
     * --------------------------------------------------------------------------
     */
    const expectedRowMap =
        new Map<string, ExpectedRow>();

    for (
        const expectedRow of expectedRows
    ) {

        expectedRowMap.set(
            buildExpectedIdentity(
                expectedRow.rowNumber,
                expectedRow.matchingKey,
            ),
            expectedRow,
        );

    }


    /**
     * จัดกลุ่มผลการเปรียบเทียบ
     *
     * 1 Test Data Row + 1 Matching Key = 1 Group
     */
    const groupedResults =
        transformCompareResults(
            results,
        );

    /**
 * Map สำหรับค้นหาผล Compare
 * ด้วยตัวตนของ Expected Row
 *
 * ใช้กับ:
 * - Exact Matching Key
 * - Composite Fallback Internal Key
 * - Expected Row ที่ไม่พบ Report
 * - Expected Row ที่เป็น Ambiguous
 */
    const groupedResultByExpectedIdentity =
        new Map<
            string,
            GroupedCompareResult
        >();

    /**
     * Map สำหรับเชื่อมผล Compare
     * กลับไปยัง Report Row จริง
     *
     * ใช้ reportRowNumber แทน matchingKey
     * เพราะ Composite Fallback ใช้ Internal Key
     * ซึ่งไม่ตรงกับ Matching Key ใน Report
     */
    const groupedResultMapByReportRowNumber =
        new Map<
            number,
            GroupedCompareResult[]
        >();

    for (
        const group of groupedResults
    ) {
        const expectedIdentity =
            buildExpectedIdentity(
                group.testDataRowNumber,
                group.matchingKey,
            );

        groupedResultByExpectedIdentity.set(
            expectedIdentity,
            group,
        );

        /**
         * reportRowNumber เท่ากับ 0 หมายถึง:
         * - ไม่พบ Report Row
         * - Ambiguous
         * - No Fee
         * - Exclusion ที่ไม่พบ Report
         *
         * Case เหล่านี้จะถูกเขียนภายหลัง
         * ในส่วน Expected Row ที่ไม่มี Report Mapping
         */
        if (
            group.reportRowNumber <= 0
        ) {
            continue;
        }

        const reportRowGroups =
            groupedResultMapByReportRowNumber.get(
                group.reportRowNumber,
            ) ?? [];

        reportRowGroups.push(
            group,
        );

        groupedResultMapByReportRowNumber.set(
            group.reportRowNumber,
            reportRowGroups,
        );
    }

    /**
     * จัดลำดับ Output
     *
     * 1. PASS
     * 2. แถว Report ที่ Map ไม่เจอ / Result ว่าง
     * 3. SKIP
     * 4. FAIL
     */
    const sortedActualRows =
        [...actualRows].sort(
            (
                rowA,
                rowB,
            ) => {

                const groupA =
                    groupedResultMapByReportRowNumber.get(
                        rowA.rowNumber,
                    )?.[0];

                const groupB =
                    groupedResultMapByReportRowNumber.get(
                        rowB.rowNumber,
                    )?.[0];

                const getSortOrder = (
                    group:
                        GroupedCompareResult |
                        undefined,
                ): number => {

                    if (
                        !group
                    ) {

                        return 2;

                    }

                    const status =
                        getGroupStatus(
                            group,
                        );

                    if (
                        status === "PASS"
                    ) {

                        return 1;

                    }

                    if (
                        status === "SKIP"
                    ) {

                        return 3;

                    }

                    return 4;

                };

                return (
                    getSortOrder(
                        groupA,
                    ) -
                    getSortOrder(
                        groupB,
                    )
                );

            },
        );

    /**
     * --------------------------------------------------------------------------
     * แถวที่ 1 เป็น Header
     * --------------------------------------------------------------------------
     */
    const headerRow =
        worksheet.addRow(
            OUTPUT_HEADERS,
        );

    applyReportHeaderStyle(
        headerRow,
    );

    /**
     * --------------------------------------------------------------------------
     * Row 2 เป็นต้นไป: แสดงข้อมูล Report ทุกแถว
     * --------------------------------------------------------------------------
     */
    for (
        const actualRow of sortedActualRows
    ) {

        /**
 * ค้นหาผล Compare จาก Report Row Number
 *
 * รองรับทั้ง:
 * - Exact Matching
 * - Composite Fallback Matching
 */
        const groups =
            groupedResultMapByReportRowNumber.get(
                actualRow.rowNumber,
            ) ?? [];
        /**
         * Report Row ที่ Map กับ Test Data ไม่เจอ
         *
         * แสดงข้อมูล Report ตามปกติ
         * แต่เว้น Test Script No., Result และ Remark
         */
        if (
            groups.length === 0
        ) {

            const rowValues: Array<
                string |
                number |
                boolean |
                null
            > = [

                    null,
                    null,
                    null,

                ];

            for (
                const reportHeader of DS_PTX_REPORT_HEADERS
            ) {

                rowValues.push(
                    normalizeExcelValue(
                        actualRow.data[
                        reportHeader
                        ],
                    ),
                );

            }

            const row =
                worksheet.addRow(
                    rowValues,
                );

            applyDataRowStyle(
                row,
            );

            continue;

        }

        /**
         * Matching Key เดียวกันอาจถูกใช้โดย Test Data หลายแถว
         * จึงต้องเขียน Output แยกหนึ่งแถวต่อหนึ่ง Group
         */
        for (
            const group of groups
        ) {

            const expectedRow =
                expectedRowMap.get(
                    buildExpectedIdentity(
                        group.testDataRowNumber,
                        group.matchingKey,
                    ),
                );

            const groupStatus =
                getGroupStatus(
                    group,
                );

            const groupRemark =
                getGroupRemark(
                    group,
                );

            const passedFields =
                getPassedFields(
                    group,
                );

            const failedFields =
                getFailedFields(
                    group,
                );

            const rowValues: Array<
                string |
                number |
                boolean |
                null
            > = [

                    expectedRow?.testScriptNo || null,
                    groupStatus,
                    groupRemark,

                ];

            for (
                const reportHeader of DS_PTX_REPORT_HEADERS
            ) {

                rowValues.push(
                    normalizeExcelValue(
                        actualRow.data[
                        reportHeader
                        ],
                    ),
                );

            }

            const row =
                worksheet.addRow(
                    rowValues,
                );

            applyDataRowStyle(
                row,
            );

            /**
             * Test Script No.
             *
             * PASS = เขียว
             * FAIL = แดง
             * SKIP = เหลือง
             */
            applyTestScriptNoStyle(
                row.getCell(1),
                groupStatus,
            );

            /**
             * Result
             */
            applyResultStyle(
                row.getCell(2),
                groupStatus,
            );

            /**
             * Remark ของ Case FAIL = สีแดง
             */
            if (
                groupStatus === "FAIL" &&
                groupRemark
            ) {

                applyFailedFieldStyle(
                    row.getCell(3),
                );

                row.getCell(3).alignment = {

                    horizontal:
                        "left",

                    vertical:
                        "top",

                    wrapText:
                        true,

                };

                const remarkLineCount =
                    groupRemark.split(
                        "\n",
                    ).length;

                row.height =
                    Math.max(
                        22,
                        remarkLineCount * 30,
                    );

            }

            /**
             * กรณีผลรวมเป็น PASS
             *
             * Highlight สีเขียวเฉพาะ Field ที่ Compare แล้ว PASS
             */
            if (
                groupStatus === "PASS"
            ) {

                const referenceHeaderIndex =
                    DS_PTX_REPORT_HEADERS.indexOf(
                        "Reference Transaction Number",
                    );

                if (
                    referenceHeaderIndex !== -1
                ) {

                    applyPassedFieldStyle(
                        row.getCell(
                            referenceHeaderIndex + 4,
                        ),
                    );

                }

                for (
                    const passedField of passedFields
                ) {

                    const reportHeaderIndex =
                        DS_PTX_REPORT_HEADERS.indexOf(
                            passedField,
                        );

                    if (
                        reportHeaderIndex === -1
                    ) {

                        continue;

                    }

                    applyPassedFieldStyle(
                        row.getCell(
                            reportHeaderIndex + 4,
                        ),
                    );

                }

            }

            /**
             * กรณีผลรวมเป็น FAIL
             *
             * - Field ที่ถูก = สีเขียว
             * - Field ที่ผิด = สีแดง
             * - Test Script No., Result และ Remark = สีแดง
             * - Field ที่ไม่ได้ Compare = ไม่ลงสี
             */
            if (
                groupStatus === "FAIL"
            ) {

                /**
                 * Matching Key ที่หาเจอใน Report ถือว่าผ่าน
                 *
                 * ถ้า Matching Key ผิดจริง
                 * failedFields จะลงสีแดงทับภายหลัง
                 */
                const referenceHeaderIndex =
                    DS_PTX_REPORT_HEADERS.indexOf(
                        "Reference Transaction Number",
                    );

                if (
                    referenceHeaderIndex !== -1
                ) {

                    applyPassedFieldStyle(
                        row.getCell(
                            referenceHeaderIndex + 4,
                        ),
                    );

                }

                /**
                 * Field ที่ Compare แล้ว PASS = สีเขียว
                 */
                for (
                    const passedField of passedFields
                ) {

                    const reportHeaderIndex =
                        DS_PTX_REPORT_HEADERS.indexOf(
                            passedField,
                        );

                    if (
                        reportHeaderIndex === -1
                    ) {

                        continue;

                    }

                    applyPassedFieldStyle(
                        row.getCell(
                            reportHeaderIndex + 4,
                        ),
                    );

                }

                /**
                 * Field ที่ FAIL
                 *
                 * - Field ที่ไม่มี Mapping = สีเหลือง
                 * - Field ที่มี Mapping = สีแดง
                 *
                 * ทำหลังสีเขียว เพื่อให้สี FAIL ทับ
                 */
                for (
                    const failedField of failedFields
                ) {

                    const reportHeaderIndex =
                        DS_PTX_REPORT_HEADERS.indexOf(
                            failedField,
                        );

                    if (
                        reportHeaderIndex === -1
                    ) {

                        continue;

                    }

                    const failedCell =
                        row.getCell(
                            reportHeaderIndex + 4,
                        );

                    /**
                     * Field ที่ไม่มี Mapping ใน Row 1
                     * ถ้าค่าไม่ตรงกันให้เป็นสีเหลือง
                     */
                    if (
                        !hasTestDataMapping(
                            failedField,
                        )
                    ) {

                        failedCell.fill = {

                            type:
                                "pattern",

                            pattern:
                                "solid",

                            fgColor: {

                                argb:
                                    COLORS.WRONG_FILL,

                            },

                        };

                        failedCell.font = {

                            color: {

                                argb:
                                    COLORS.WRONG_TEXT,

                            },

                        };

                        continue;

                    }

                    /**
                     * Field ที่มี Mapping และค่าไม่ตรงกัน
                     * ให้เป็นสีแดง
                     */
                    applyFailedFieldStyle(
                        failedCell,
                    );

                }

            }

        }

    }

    /**
  * ============================================================================
  * เพิ่ม Expected Row ที่ไม่มี Report Mapping
  * ============================================================================
  *
  * ครอบคลุม:
  * - Exact Matching Not Found
  * - Composite Fallback Not Found
  * - Composite Fallback Ambiguous
  * - No Fee
  * - Exclusion Case ที่ไม่พบ Report
  *
  * Status และ Remark ใช้ผลที่สร้างจาก ptx-reconcile.ts
  * โดย Writer จะไม่ตัดสิน Business Rule ซ้ำ
  */
    for (
        const expectedRow of expectedRows
    ) {
        const expectedIdentity =
            buildExpectedIdentity(
                expectedRow.rowNumber,
                expectedRow.matchingKey,
            );

        const group =
            groupedResultByExpectedIdentity.get(
                expectedIdentity,
            );

        /**
         * ปกติ Expected Row ทุกแถว
         * ต้องมี Compare Result
         *
         * หากไม่พบ Group ให้ข้ามอย่างปลอดภัย
         * โดยไม่ทำให้การเขียนไฟล์หยุด
         */
        if (
            !group
        ) {
            continue;
        }

        /**
         * Report Row Number มากกว่า 0
         * หมายถึงผลถูกเขียนพร้อม Actual Row
         * ในรอบหลักเรียบร้อยแล้ว
         */
        if (
            group.reportRowNumber > 0
        ) {
            continue;
        }

        const status =
            getGroupStatus(
                group,
            );

        /**
         * กรณี FAIL:
         * รวมรายละเอียดของทุก Field ที่ผิด
         *
         * กรณี PASS หรือ SKIP:
         * ใช้ Remark จาก Compare Result โดยตรง
         */
        const nonFailRemark =
            Object.values(
                group.fields,
            )
                .map(
                    result =>
                        result.remark.trim(),
                )
                .find(
                    resultRemark =>
                        resultRemark !== "",
                );

        const remark =
            status === "FAIL"
                ? (
                    getGroupRemark(
                        group,
                    ) ??
                    "Validation Failed"
                )
                : (
                    nonFailRemark ??
                    ""
                );


        /**
         * เตรียมข้อมูล 3 Column แรก:
         *
         * A = Test Script No.
         * B = Result
         * C = Remark
         */
        const rowValues: Array<
            string |
            number |
            boolean |
            null
        > = [

                expectedRow.testScriptNo || null,
                status,
                remark,

            ];

        /**
         * เติมข้อมูลตาม Header ของ DS_PTX
         */
        for (
            const reportHeader of DS_PTX_REPORT_HEADERS
        ) {

            /**
             * แสดง Expected Matching Key
             * ในช่อง Reference Transaction Number
             */
            if (
                reportHeader ===
                "Reference Transaction Number"
            ) {

                rowValues.push(
                    expectedRow.matchingKey,
                );

                continue;

            }

            /**
             * Field อื่นไม่มี Actual Data
             * เพราะไม่พบรายการใน Report
             */
            rowValues.push(
                null,
            );

        }

        /**
         * เพิ่ม Row ลงใน Excel
         */
        const row =
            worksheet.addRow(
                rowValues,
            );

        applyDataRowStyle(
            row,
        );

        /**
         * ลงสี Test Script No.
         */
        applyTestScriptNoStyle(
            row.getCell(1),
            status,
        );

        /**
         * ลงสี Result
         */
        applyResultStyle(
            row.getCell(2),
            status,
        );

        /**
 * ลงสี Remark ตาม Status
 *
 * PASS = สีเขียว
 * FAIL = สีแดง
 * SKIP = สีเหลือง
 */
        applyTestScriptNoStyle(
            row.getCell(3),
            status,
        );

        row.getCell(3).alignment = {

            horizontal:
                "left",

            vertical:
                "top",

            wrapText:
                true,

        };

        row.height =
            40;

        /**
         * หา Column ของ Matching Key
         */
        const referenceHeaderIndex =
            DS_PTX_REPORT_HEADERS.indexOf(
                "Reference Transaction Number",
            );

        if (
            referenceHeaderIndex !== -1
        ) {

            const matchingKeyCell =
                row.getCell(
                    referenceHeaderIndex + 4,
                );

            /**
                        /**
             * ลงสี Matching Key ตาม Status
             *
             * PASS = สีเขียว
             * FAIL = สีแดง
             * SKIP = สีเหลือง
             */
            applyTestScriptNoStyle(
                matchingKeyCell,
                status,
            );

        }

    }
    /**
     * Freeze Row 1 และ Column A-C
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
     * Auto Filter ที่ Header Row 1
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
                OUTPUT_HEADERS.length,

        },

    };

    applyColumnWidths(
        worksheet,
    );

    /**
     * บันทึกไฟล์
     */
    await workbook.xlsx.writeFile(
        outputFile,
    );

};