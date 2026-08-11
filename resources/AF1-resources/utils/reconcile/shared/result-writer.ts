/**
 * ReconcileResultSheetWriter
 * ------------------------------------------------------------------
 * Bug fix / Redesign (Code Review) — เปลี่ยนกติกา Test_Result ตามที่ทีมแจ้ง:
 * 1) Test_Result เหลือแค่ 2 สถานะ "Pass" / "Fail" (ตัด "Please Review" ออกจากสถานะแถว
 *    โดยสิ้นเชิง) — ตัดสินจาก "3 หัวข้อ" เท่านั้น (ดู ReconcileService.checkKeyConditions):
 *    Reference Transaction Number, FI Arrangement Number, Transaction Amount
 * 2) Field อื่นทั้งหมด (Payment Method, Currency, Beneficiary Name, ฯลฯ) ไม่มีผลต่อ
 *    Pass/Fail ของแถวอีกต่อไป — ถ้าไม่ตรงกัน (ไม่ว่าจะเป็นแบบที่เคยเรียกว่า FAIL หรือ REVIEW)
 *    จะกลายเป็น "ไฮไลท์สีเหลืองที่ cell ข้อมูลนั้น + ข้อความใน Remark" เท่านั้น ไม่กระทบ
 *    สถานะ Pass/Fail ของทั้งแถว
 * 3) เปลี่ยนชุดสีตามที่กำหนด (fill + text สีเข้มเฉพาะ ไม่ใช่ auto ขาว/ดำแบบเดิม):
 *    PASS: fill FFC6EFCE, text FF006100 | FAIL: fill FFFFC7CE, text FF9C0006 |
 *    REVIEW (เฉพาะ cell ข้อมูล ไม่ใช่ Test_Result): fill FFFFE699, text FF7F6000
 * 4) เพิ่ม AutoFilter กลับมา (ให้เหมือน Tab DS_LTX)
 * ------------------------------------------------------------------
 */
import ExcelJS from "exceljs";
import { COLORS } from "../../validators/shared/excel-style.util";

const META_COLUMN_HEADERS = [
  "Test Script No.",
  "Test Result",
  "Remark",
] as const;
const META_COLUMN_COUNT = META_COLUMN_HEADERS.length;
const REPORT_COLUMN_START_INDEX = META_COLUMN_COUNT + 1;

const HEADER_ROW_NUMBER = 1;
const FIRST_DATA_ROW_NUMBER = 2;

const HEADER_ROW_HEIGHT = 24;
const DATA_ROW_HEIGHT = 20;
const META_COLUMN_WIDTHS: Record<(typeof META_COLUMN_HEADERS)[number], number> =
  {
    "Test Script No.": 18,
    "Test Result": 15,
    Remark: 70,
  };
const REPORT_COLUMN_MIN_WIDTH = 14;
const REPORT_COLUMN_MAX_WIDTH = 32;

/**
 * Data Quality Remark กลางสำหรับทุก Report
 *
 * ข้อความนี้ไม่เปลี่ยน Business Pass/Fail
 * และไม่แสดง Transaction ID แทน Test No.
 */
const MISSING_TEST_NO_REMARK =
  `[TS] : Test No. = "" | ` + "ไม่พบ Test No. ใน Test Data";

/** สถานะระดับแถว (Test_Result) — เหลือแค่ 2 สถานะตามที่ทีมแจ้ง ไม่มี REVIEW แล้ว */
export type RowStatus = "PASS" | "FAIL";

/** ข้อความที่แสดงในคอลัมน์ "Test Result." */
const ROW_STATUS_LABEL: Record<RowStatus, string> = {
  PASS: "Pass",
  FAIL: "Fail",
};

/** ชุดสีตามที่กำหนด — fill + text สีเข้มเฉพาะ (ไม่ใช้ applyFill เดิมที่ auto เลือกขาว/ดำ) */
const STYLE = {
  PASS: { fill: "FFC6EFCE", text: "FF006100" },
  FAIL: { fill: "FFFFC7CE", text: "FF9C0006" },
  REVIEW: { fill: "FFFFE699", text: "FF7F6000" },
} as const;

/** ใส่ fill + text color ตามชุดสีที่กำหนด (แทนที่ applyFill/applyBoldFill เดิมสำหรับ cell กลุ่มนี้) */
const applyStatusStyle = (
  cell: ExcelJS.Cell,
  style: { fill: string; text: string },
): void => {
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: style.fill },
  };
  cell.font = { ...(cell.font ?? {}), color: { argb: style.text } };
};

/** ผลลัพธ์ 1 แถวที่พร้อมเขียนลง Sheet (DR หรือ FE ของ 1 Test Case) */
export interface ResultRow {
  testCaseNo: string;
  /** Pass/Fail ของทั้งแถว ตัดสินจาก "3 หัวข้อ" เท่านั้น (ดู checkKeyConditions ใน ltx-reconcile.ts) */
  status: RowStatus;
  /** ข้อความสรุปรายการที่ต่างกัน (มาจาก field อื่นที่ไม่ใช่ 3 หัวข้อหลัก) ต่อท้ายด้วย "Please review" */
  remark: string;
  /** เลขแถวจริงใน sourceWorksheet ที่ match ได้ — undefined = ไม่เจอแถวคู่กันเลย (คอลัมน์ Report เว้นว่าง) */
  matchedRowNumber: number | undefined;
  /** field header ใน "3 หัวข้อหลัก" ที่ไม่ตรงกัน (ไฮไลท์แดง — เป็นตัวตัดสิน Fail ของแถว) */
  failedKeyFieldHeaders: string[];
  /** field header อื่น (นอก 3 หัวข้อหลัก) ที่ไม่ตรงกัน (ไฮไลท์เหลือง — ไม่กระทบ Pass/Fail) */
  reviewFieldHeaders: string[];

  /**
   * true = ไม่มี AF1 row เพราะ Requirement กำหนดว่ารายการต้องไม่แสดง
   * ใช้แยก Expected Absence ที่เป็น PASS ออกจาก Missing Reference ที่เป็น FAIL
   */
  isExpectedAbsence?: boolean;
}

export class ReconcileResultSheetWriter {
  /**
   * ต่อ Data Quality Remark เมื่อ Test No. ว่าง
   *
   * Business Remark เดิมยังอยู่ครบ
   * และป้องกันไม่ให้ข้อความซ้ำหากถูกเรียกมากกว่าหนึ่งครั้ง
   */
  private appendMissingTestNoRemark(
    testCaseNo: string,
    remark: string,
  ): string {
    if (testCaseNo.trim() !== "" || remark.includes(MISSING_TEST_NO_REMARK)) {
      return remark;
    }

    return [remark, MISSING_TEST_NO_REMARK]
      .map((message) => message.trim())
      .filter((message) => message !== "")
      .join("\n");
  }

  /** สร้าง Sheet ใหม่ชื่อ `${reportCode}_Reconcile` — ลบของเก่าทิ้งก่อนถ้ามีอยู่แล้ว */
  createSheet(
    workbook: ExcelJS.Workbook,
    reportCode: string,
  ): ExcelJS.Worksheet {
    const sheetName = `${reportCode}_Reconcile`;
    const existingSheet = workbook.getWorksheet(sheetName);
    if (existingSheet) {
      workbook.removeWorksheet(existingSheet.id);
    }
    return workbook.addWorksheet(sheetName);
  }

  /** กำหนดความกว้าง column ให้เหมาะสม (Meta 3 คอลัมน์ตายตัว + Report column ตามความยาวชื่อ header) */
  private setColumnWidths(
    resultSheet: ExcelJS.Worksheet,
    reportHeaders: string[],
  ): void {
    META_COLUMN_HEADERS.forEach((header, index) => {
      resultSheet.getColumn(index + 1).width = META_COLUMN_WIDTHS[header];
    });

    reportHeaders.forEach((header, index) => {
      const estimatedWidth = Math.max(
        REPORT_COLUMN_MIN_WIDTH,
        Math.min(REPORT_COLUMN_MAX_WIDTH, header.length + 4),
      );
      resultSheet.getColumn(REPORT_COLUMN_START_INDEX + index).width =
        estimatedWidth;
    });
  }

  /** เขียน header แถวเดียว (row 1) พร้อมตั้งความสูง/กว้างที่เหมาะสม */
  writeHeaderRow(
    resultSheet: ExcelJS.Worksheet,
    reportHeaders: string[],
  ): void {
    this.setColumnWidths(resultSheet, reportHeaders);

    const headerRow = resultSheet.getRow(HEADER_ROW_NUMBER);
    headerRow.height = HEADER_ROW_HEIGHT;

    const writeHeaderCell = (cell: ExcelJS.Cell, header: string): void => {
      cell.value = header;
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: COLORS.BLUE },
      };
      cell.font = { bold: true, color: { argb: COLORS.WHITE } };
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      };
    };

    META_COLUMN_HEADERS.forEach((header, index) =>
      writeHeaderCell(headerRow.getCell(index + 1), header),
    );
    reportHeaders.forEach((header, index) =>
      writeHeaderCell(
        headerRow.getCell(REPORT_COLUMN_START_INDEX + index),
        header,
      ),
    );

    headerRow.commit();
  }

  /**
   * คำนวณ Row Height ของ Result Reconcile ตามข้อความใน Remark
   *
   * ExcelJS ไม่มีคำสั่ง AutoFit Row Height โดยตรง จึงประเมินจาก:
   * - จำนวนบรรทัดที่คั่นด้วย \n
   * - จำนวนบรรทัดที่เกิดจาก Wrap Text
   *
   * Column Remark กว้าง 70 จึงประเมินประมาณ 80 ตัวอักษรต่อบรรทัด
   */
  private calculateRowHeight(remark: string): number {
    if (remark.trim() === "") {
      return DATA_ROW_HEIGHT;
    }

    const estimatedCharactersPerLine = 80;
    const heightPerLine = 20;
    const verticalPadding = 5;

    const visualLineCount = remark.split(/\r?\n/).reduce((totalLines, line) => {
      const wrappedLineCount = Math.max(
        1,
        Math.ceil(line.length / estimatedCharactersPerLine),
      );

      return totalLines + wrappedLineCount;
    }, 0);

    return Math.max(
      DATA_ROW_HEIGHT,
      visualLineCount * heightPerLine + verticalPadding,
    );
  }

  /** ใส่ข้อมูล Test No./Result/Remark 1 แถว พร้อมสี Pass(เขียว)/Fail(แดง) เท่านั้น (ไม่มีเหลืองที่นี่) */
  private writeMetaCells(
    outputRow: ExcelJS.Row,
    testCaseNo: string,
    status: RowStatus,
    remark: string,
  ): void {
    const finalRemark = this.appendMissingTestNoRemark(testCaseNo, remark);

    outputRow.getCell(1).value = testCaseNo;
    outputRow.getCell(2).value = ROW_STATUS_LABEL[status];
    outputRow.getCell(3).value = finalRemark;

    outputRow.getCell(1).alignment = {
      vertical: "middle",
    };
    outputRow.getCell(2).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    outputRow.getCell(3).alignment = {
      wrapText: true,
      vertical: "middle",
    };

    // Auto fit โดยประมาณตามจำนวนบรรทัดและ Wrap Text ของ Remark
    outputRow.height = this.calculateRowHeight(finalRemark);

    const style = STYLE[status];
    applyStatusStyle(outputRow.getCell(1), style);
    applyStatusStyle(outputRow.getCell(2), style);
    applyStatusStyle(outputRow.getCell(3), style);
  }

  /**
   * Cleanup (Code Review): เดิมมี writeDataRows() และ writeUnmatchedCaseRows()
   * แยกต่างหาก แต่ไม่มีจุดใดในระบบเรียกใช้ (Logic ซ้ำกับ writeRowsInRequestedOrder()
   * ด้านล่างเกือบทั้งหมด) — ลบออกแล้ว ให้เหลือ writeRowsInRequestedOrder() เป็นทางเข้า
   * เดียวสำหรับเขียนแถวข้อมูล
   */
  /**
   * เขียน Result ตามลำดับที่ต้องการ:
   * 1) PASS
   * 2) แถว AF1 ที่ไม่ได้ผูกกับ Test Case
   * 3) FAIL
   * 4) Test Case ที่หาแถว AF1 ไม่เจอ
   *
   * ภายในกลุ่ม PASS / FAIL / แถว AF1 ที่ไม่ได้ผูกกับ Test Case
   * จะยังเรียงตามลำดับเดิมใน AF1 Report
   *
   * @returns เลขแถวถัดไปที่ยังว่างอยู่
   */
  writeRowsInRequestedOrder(
    resultSheet: ExcelJS.Worksheet,
    sourceWorksheet: ExcelJS.Worksheet,
    reportHeaders: string[],
    firstSourceRowNumber: number,
    lastSourceRowNumber: number,
    annotationByRowNumber: Map<number, ResultRow>,
    unmatchedRows: ResultRow[],
  ): number {
    const passRowNumbers: number[] = [];
    const noTestCaseRowNumbers: number[] = [];
    const failRowNumbers: number[] = [];

    for (
      let sourceRowNumber = firstSourceRowNumber;
      sourceRowNumber <= lastSourceRowNumber;
      sourceRowNumber += 1
    ) {
      const sourceRow = sourceWorksheet.getRow(sourceRowNumber);

      if (sourceRow.cellCount === 0 && sourceRow.actualCellCount === 0) {
        continue;
      }

      const annotation = annotationByRowNumber.get(sourceRowNumber);

      if (!annotation) {
        noTestCaseRowNumbers.push(sourceRowNumber);
      } else if (annotation.status === "PASS") {
        passRowNumbers.push(sourceRowNumber);
      } else {
        failRowNumbers.push(sourceRowNumber);
      }
    }

    const expectedAbsenceRows = unmatchedRows.filter(
      (row) => row.status === "PASS" && row.isExpectedAbsence === true,
    );
    const unexpectedMissingRows = unmatchedRows.filter(
      (row) => !(row.status === "PASS" && row.isExpectedAbsence === true),
    );

    let outputRowNumber = FIRST_DATA_ROW_NUMBER;

    const writeSourceRow = (sourceRowNumber: number): void => {
      const sourceRow = sourceWorksheet.getRow(sourceRowNumber);
      const annotation = annotationByRowNumber.get(sourceRowNumber);
      const outputRow = resultSheet.getRow(outputRowNumber);

      outputRow.height = DATA_ROW_HEIGHT;

      if (annotation) {
        this.writeMetaCells(
          outputRow,
          annotation.testCaseNo,
          annotation.status,
          annotation.remark,
        );
      }

      reportHeaders.forEach((header, index) => {
        const sourceCell = sourceRow.getCell(index + 1);
        const outputCell = outputRow.getCell(REPORT_COLUMN_START_INDEX + index);

        outputCell.value = sourceCell.value;

        if (header.trim().toLowerCase() === "transaction amount") {
          outputCell.numFmt = "#,##0.00";
        } else {
          outputCell.numFmt = sourceCell.numFmt || "General";
        }

        if (annotation) {
          if (annotation.failedKeyFieldHeaders.includes(header)) {
            applyStatusStyle(outputCell, STYLE.FAIL);
          } else if (annotation.reviewFieldHeaders.includes(header)) {
            applyStatusStyle(outputCell, STYLE.REVIEW);
          } else {
            applyStatusStyle(outputCell, STYLE.PASS);
          }
        }
      });

      outputRow.commit();
      outputRowNumber += 1;
    };

    const writeUnmatchedRow = (row: ResultRow): void => {
      const outputRow = resultSheet.getRow(outputRowNumber);
      outputRow.height = DATA_ROW_HEIGHT;
      this.writeMetaCells(outputRow, row.testCaseNo, row.status, row.remark);
      outputRow.commit();
      outputRowNumber += 1;
    };

    // 1) PASS ที่มีข้อมูลใน Raw Report
    passRowNumbers.forEach(writeSourceRow);

    // 2) Raw Report ที่ไม่เข้าเงื่อนไข
    // หรือไม่ได้ผูกกับ Test Case
    noTestCaseRowNumbers.forEach(writeSourceRow);

    // 3) FAIL ที่มีข้อมูลใน Raw Report
    failRowNumbers.forEach(writeSourceRow);

    // 4) PASS ที่ไม่มีข้อมูลใน Raw Report
    // เป็น Expected Absence ตาม Requirement
    expectedAbsenceRows.forEach(writeUnmatchedRow);

    // 5) FAIL ที่ไม่มีข้อมูลใน Raw Report
    unexpectedMissingRows.forEach(writeUnmatchedRow);

    return outputRowNumber;
  }

  /**
   * ใส่ AutoFilter ให้ครอบคลุมตั้งแต่ header (row 1) ถึงแถวข้อมูลสุดท้าย — เหมือน Tab DS_LTX
   * ต้องเรียกหลังเขียนข้อมูลครบทุกแถวแล้วเท่านั้น เพราะตอน writeHeaderRow() ยังไม่รู้ว่า
   * ข้อมูลจะจบที่แถวไหน
   */
  finalizeAutoFilter(
    resultSheet: ExcelJS.Worksheet,
    reportHeaders: string[],
    lastDataRowNumber: number,
  ): void {
    if (lastDataRowNumber < FIRST_DATA_ROW_NUMBER) {
      return;
    }

    const lastColumnNumber =
      REPORT_COLUMN_START_INDEX + reportHeaders.length - 1;

    resultSheet.autoFilter = {
      from: { row: HEADER_ROW_NUMBER, column: 1 },
      to: { row: lastDataRowNumber, column: lastColumnNumber },
    };
  }
}
