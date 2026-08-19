/**
 * fxm-reconcile-model.ts
 * ------------------------------------------------------------------
 * Type, Interface, Constant และ Parser ที่ใช้ร่วมกันภายใน DF_FXM
 * ไฟล์นี้ไม่มีขั้นตอนอ่าน/เขียนไฟล์ จึงไม่มี Side Effect
 * ------------------------------------------------------------------
 */

import {
    FXM_TEST_DATA_FIELDS,
} from "./fxm-config";

import {
    FXMDirection,
} from "./fxm-rules";

import {
    ReconcileRecord,
} from "../shared/record";

import {
    PreparedReconcileWorkbook,
} from "../shared/workbook-preparer";


/**
 * สถานะที่ใช้ตัดสินว่า Transaction
 * ต้องมีหรือไม่ต้องมีใน Report DF_FXM
 *
 * MUST_EXIST:
 * รายการต้องพบใน DF_FXM
 *
 * MUST_NOT_EXIST:
 * รายการต้องไม่พบใน DF_FXM
 *
 * CANNOT_DECIDE:
 * ข้อมูล Test Data ไม่เพียงพอ
 * จึงยังไม่สามารถตัดสิน Presence Rule ได้
 */
export type FXMPresenceExpectation =
    | "MUST_EXIST"
    | "MUST_NOT_EXIST"
    | "CANNOT_DECIDE";

/**
 * ผลการตัดสิน Rule ของ Test Data หนึ่งแถว
 */
export interface FXMReconcileDecision {
    /**
     * ทิศทางของธุรกรรม FX
     *
     * ตัวอย่าง:
     * - BUY_FCY
     * - SELL_FCY
     * - CROSS_CURRENCY
     * - NOT_FX
     * - UNKNOWN
     */
    direction: FXMDirection;

    /**
     * รายการต้องมีหรือไม่ต้องมีใน DF_FXM
     */
    expectation: FXMPresenceExpectation;

    /**
     * เหตุผลเพิ่มเติมจากการตัดสิน Rule
     *
     * ใช้แสดงใน Remark
     */
    reasons: string[];

    /**
     * true:
     * มีข้อมูลบางส่วนที่ต้องให้ผู้ใช้งานตรวจสอบเพิ่มเติม
     *
     * ปัจจุบันใช้กับ Cross Currency
     * เพราะ Requirement ของ Settlement/Intermediary
     * ยังให้รายละเอียดไม่ครบ
     */
    requiresReview: boolean;
}

/**
 * วิธีที่ระบบใช้จับคู่ Test Data กับ DF_FXM Report
 *
 * EXACT:
 * Transaction ID ตรงกับ Arrangement Number
 *
 * FALLBACK:
 * ใช้ Txn Date และ Settled Amount ช่วยค้นหา
 *
 * NOT_FOUND:
 * ไม่พบทั้ง Exact และ Fallback
 *
 * AMBIGUOUS:
 * พบ Candidate มากกว่าหนึ่งแถว
 *
 * ROW_ALREADY_USED:
 * Report Row ถูกใช้กับ Test Case ก่อนหน้าแล้ว
 *
 * FALLBACK_UNAVAILABLE:
 * ข้อมูลไม่เพียงพอสำหรับทำ Fallback
 */
export type FXMMatchMode =
    | "EXACT"
    | "FALLBACK"
    | "NOT_FOUND"
    | "AMBIGUOUS"
    | "ROW_ALREADY_USED"
    | "FALLBACK_UNAVAILABLE";

/**
 * ผลการจับคู่ Test Data หนึ่งแถว
 */
export interface FXMMatchResolution {
    /**
     * ค่าที่ใช้แสดงใน Column Test Script No.
     *
     * ลำดับการเลือก:
     * 1. Test No.
     * 2. Transaction ID
     * 3. TEST DATA ROW ตามด้วยเลขแถว
     */
    testCaseNo: string;

    /**
     * Transaction ID ที่ผ่านการ Normalize แล้ว
     */
    transactionId: string;

    /**
     * ผล Presence Rule ของ Test Data แถวนี้
     */
    decision:
    FXMReconcileDecision;

    /**
     * วิธีที่ใช้จับคู่ข้อมูล
     */
    matchMode:
    FXMMatchMode;

    /**
     * Report Record ที่จับคู่ได้
     *
     * ไม่มีค่าในกรณี:
     * - NOT_FOUND
     * - AMBIGUOUS
     * - ROW_ALREADY_USED
     * - FALLBACK_UNAVAILABLE
     */
    matchedRecord?:
    ReconcileRecord;

    /**
     * จำนวน Candidate ที่พบ
     */
    candidateCount:
    number;

    /**
     * รายละเอียดของผล Matching
     */
    remark:
    string;
}

/**
 * ผลการค้นหา Fallback ภายในระบบ
 */
export interface FXMFallbackResult {
    matchedRecord?:
    ReconcileRecord;

    candidateCount?:
    number;

    remark:
    string;
}

/**
 * ตัวเลือกเริ่มต้นสำหรับการเปรียบเทียบ Field
 *
 * ใช้ในกรณีพิเศษ เช่น:
 * - Fallback Matching
 * - Expected Absence แต่พบรายการ
 * - ไม่สามารถตัดสิน Presence Rule ได้
 */
export interface FXMCompareOptions {
    /**
     * Remark ที่ต้องใส่ก่อนเริ่มตรวจ Field
     */
    initialRemarks?:
    readonly string[];

    /**
     * Header ที่ต้องเริ่มต้นเป็น FAIL
     */
    initialFailedHeaders?:
    readonly string[];

    /**
     * Header ที่ต้องเริ่มต้นเป็น Review
     */
    initialReviewHeaders?:
    readonly string[];
}

/**
 * ผลการเปรียบเทียบ Field
 * ของ Test Data กับ Report หนึ่งคู่
 */
export interface FXMFieldComparisonResult {
    /**
     * Header ที่ไม่ผ่าน Business Rule
     *
     * มีผลต่อสถานะระดับ Test Case
     */
    failedHeaders:
    string[];

    /**
     * Header ที่ต้องตรวจสอบเพิ่มเติม
     *
     * ไม่มีผลต่อสถานะ PASS/FAIL
     */
    reviewHeaders:
    string[];

    /**
     * รายละเอียดค่าที่ไม่ตรงกัน
     */
    remarks:
    string[];
}

/**
 * ข้อมูลที่อ่านและตรวจ Header เรียบร้อยแล้ว
 *
 * prepared:
 * Workbook ที่สร้างจาก Checked Report
 *
 * reportRecords:
 * ข้อมูลทุกแถวจาก DF_FXM Report
 *
 * testDataRecords:
 * ข้อมูลทุกแถวจาก Test Data
 *
 * reportIndex:
 * Index สำหรับค้นหา Arrangement Number
 *
 * reservedExactRowNumbers:
 * Report Row ที่ถูกจองไว้สำหรับ Exact Matching
 */
export interface FXMLoadedData {
    prepared:
    PreparedReconcileWorkbook;

    reportRecords:
    ReconcileRecord[];

    testDataRecords:
    ReconcileRecord[];

    reportIndex:
    Map<
        string,
        ReconcileRecord[]
    >;

    reservedExactRowNumbers:
    Set<number>;
}

/**
 * ค่าคลาดเคลื่อนของ Amount
 *
 * ใช้ป้องกันปัญหาทศนิยมจาก Excel เช่น:
 *
 * Test Data:
 * 999999.999999
 *
 * Report:
 * 1000000.00
 *
 * ถ้าผลต่างไม่เกิน 0.01
 * จะถือว่า Amount เท่ากัน
 */
export const FXM_AMOUNT_TOLERANCE =
    0.01;

/**
 * รายการ Header ของ Test Data
 * ที่ Script 3 สำหรับ DF_FXM ต้องใช้
 *
 * ใช้ Set เพื่อตัด Header ซ้ำ
 * หาก Config มีชื่อ Header เดียวกันมากกว่าหนึ่งตำแหน่ง
 */
export const FXM_REQUIRED_TEST_DATA_HEADERS = [
    ...new Set(
        Object.values(
            FXM_TEST_DATA_FIELDS,
        ),
    ),
];

/**
 * แปลงวันที่ให้อยู่ในรูปแบบ:
 *
 * yyyy-mm-dd
 *
 * รูปแบบที่รองรับ:
 * - yyyy-mm-dd
 * - yyyy/mm/dd
 * - dd/mm/yyyy
 * - dd-mm-yyyy
 * - Excel Serial Date
 *
 * ถ้าค่าว่างหรือรูปแบบวันที่ไม่ถูกต้อง
 * จะคืนค่า null
 */
export const parseFXMDateKey = (
    value: unknown,
): string | null => {
    const text =
        String(
            value ?? "",
        ).trim();

    if (
        text ===
        ""
    ) {
        return null;
    }

    /**
     * รองรับ Excel Serial Date
     *
     * ตัวอย่าง:
     * 45986
     */
    if (
        /^\d+(?:\.\d+)?$/.test(
            text,
        )
    ) {
        const serial =
            Number(
                text,
            );

        if (
            !Number.isFinite(
                serial,
            ) ||
            serial <=
            0
        ) {
            return null;
        }

        /**
         * Excel ใช้วันที่ 1899-12-30
         * เป็นจุดเริ่มต้นสำหรับ Serial Date
         */
        const excelEpoch =
            Date.UTC(
                1899,
                11,
                30,
            );

        const date =
            new Date(
                excelEpoch +
                Math.floor(
                    serial,
                ) *
                86_400_000,
            );

        return [
            date.getUTCFullYear(),

            String(
                date.getUTCMonth() +
                1,
            ).padStart(
                2,
                "0",
            ),

            String(
                date.getUTCDate(),
            ).padStart(
                2,
                "0",
            ),
        ].join(
            "-",
        );
    }

    /**
     * รูปแบบวันขึ้นก่อน:
     *
     * dd/mm/yyyy
     * dd-mm-yyyy
     */
    const dayFirst =
        text.match(
            /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\s.*)?$/,
        );

    /**
     * รูปแบบปีขึ้นก่อน:
     *
     * yyyy/mm/dd
     * yyyy-mm-dd
     */
    const yearFirst =
        text.match(
            /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[T\s].*)?$/,
        );

    let year: number;
    let month: number;
    let day: number;

    if (
        dayFirst
    ) {
        day =
            Number(
                dayFirst[1],
            );

        month =
            Number(
                dayFirst[2],
            );

        year =
            Number(
                dayFirst[3],
            );
    } else if (
        yearFirst
    ) {
        year =
            Number(
                yearFirst[1],
            );

        month =
            Number(
                yearFirst[2],
            );

        day =
            Number(
                yearFirst[3],
            );
    } else {
        return null;
    }

    /**
     * สร้าง Date เพื่อยืนยันว่าเป็นวันที่จริง
     *
     * ตัวอย่างวันที่ที่ไม่ผ่าน:
     * - 31/02/2026
     * - 2026-13-01
     */
    const date =
        new Date(
            Date.UTC(
                year,
                month - 1,
                day,
            ),
        );

    const isValid =
        date.getUTCFullYear() ===
        year &&
        date.getUTCMonth() ===
        month - 1 &&
        date.getUTCDate() ===
        day;

    if (
        !isValid
    ) {
        return null;
    }

    return [
        year,

        String(
            month,
        ).padStart(
            2,
            "0",
        ),

        String(
            day,
        ).padStart(
            2,
            "0",
        ),
    ].join(
        "-",
    );
};

/**
 * อ่านวันที่ทำธุรกรรมจาก Transaction ID/Reconcile ID
 *
 * รูปแบบ Ref ที่รองรับ:
 * - ตัวอักษร 3 ตัว
 * - ตัวเลข 3 ตัว
 * - วันที่ YYMMDD
 * - เลขลำดับรายการ
 *
 * ตัวอย่าง:
 * KMA3012511260000000028158
 *       251126 = 26/11/2025
 *
 * ถ้า Ref ไม่มีวันที่ หรือวันที่ไม่ถูกต้อง
 * จะคืนค่า null
 */
export const parseFXMReferenceDateKey = (
    value: unknown,
): string | null => {
    const text =
        String(
            value ?? "",
        ).trim();

    const matchedDate =
        text.match(
            /^[A-Za-z]{3}\d{3}(\d{2})(\d{2})(\d{2})/,
        );

    if (
        !matchedDate
    ) {
        return null;
    }

    /**
     * วันที่ใน Ref ใช้รูปแบบ YYMMDD
     *
     * matchedDate[1] = ปี
     * matchedDate[2] = เดือน
     * matchedDate[3] = วัน
     */
    const year =
        2000 +
        Number(
            matchedDate[1],
        );

    const month =
        Number(
            matchedDate[2],
        );

    const day =
        Number(
            matchedDate[3],
        );

    /**
     * ใช้ Date Parser เดิมตรวจสอบซ้ำว่า
     * วันที่จาก Ref เป็นวันที่จริง
     *
     * ตัวอย่าง Ref ที่มีวันที่ 250231
     * จะไม่ผ่าน เพราะไม่มีวันที่ 31/02/2025
     */
    return parseFXMDateKey(
        `${day}/${month}/${year}`,
    );
};

/**
 * ประเภทของรายการที่ต้องโยงกลับไปหา
 * Original Transaction
 */
export type FXMReturnTransactionType =
    | "RETURN"
    | "REVERSAL";

/**
 * ข้อมูลที่อ่านได้จาก Test No.
 * ของ Return/Reversal
 */
export interface FXMReturnTestInfo {
    /**
     * Test No. ของ Original Transaction
     *
     * ตัวอย่าง:
     * BOTDMS_010-return
     * → BOTDMS_010
     */
    originalTestNo:
    string;

    /**
     * ประเภทของรายการ
     */
    transactionType:
    FXMReturnTransactionType;
}

/**
 * ตรวจว่า Test No. เป็น Return/Reversal หรือไม่
 * และคืน Test No. ของ Original Transaction
 *
 * รูปแบบที่รองรับ:
 * - 2511_004-Return
 * - BOTDMS_010-return
 * - TC001-Reversal
 * - TC001 - reversal
 *
 * การตรวจไม่สนตัวพิมพ์เล็ก/ใหญ่
 *
 * ถ้าไม่ใช่ Return/Reversal
 * หรือไม่มี Original Test No.
 * จะคืนค่า null
 */
export const parseFXMReturnTestInfo = (
    value: unknown,
): FXMReturnTestInfo | null => {
    const text =
        String(
            value ?? "",
        ).trim();

    /**
     * แยกข้อความออกเป็น:
     *
     * matched[1]:
     * Original Test No.
     *
     * matched[2]:
     * return หรือ reversal
     */
    const matched =
        text.match(
            /^(.*?)\s*-\s*(return|reversal)\s*$/i,
        );

    if (
        !matched
    ) {
        return null;
    }

    const originalTestNo =
        matched[1].trim();

    /**
     * ป้องกันกรณีมีเพียง:
     * -return
     * หรือ
     * -reversal
     *
     * แต่ไม่มี Original Test No.
     */
    if (
        originalTestNo ===
        ""
    ) {
        return null;
    }

    return {
        originalTestNo,

        transactionType:
            matched[2]
                .toUpperCase() as
            FXMReturnTransactionType,
    };
};