import type { ReportPresenceRule,} from "../shared/presence-rule";
import { RESIDENT_THB_TO_FCD_MUST_NOT_EXIST } from "../shared/presence-rule";
import type { ReportCode } from "../../../config/report-config";
import { getReportRuntimeConfig } from "../../../config/report-runtime.config";

// ======================================================
// Reconcile Configuration (Script 3)
//
// เก็บ Field Mapping + Compare Rule ระหว่าง AF1 Report กับ Test Data
// ไฟล์นี้เก็บว่า "field ไหนของ Report
// ต้องเทียบกับ field ไหนของ Test Data และเทียบแบบไหน" (Script3)
//
// โครงสร้างเดินตาม pattern เดียวกับ REPORT_CONFIG / REPORT_CONFIG_mapping
// เดิม: 1 object เก็บทุก report, key = report code เดียวกับที่ report-detector
// ใช้อยู่แล้ว (DS_LTX, DS_PTX, ...)
// ======================================================

/** วิธีเปรียบเทียบค่าของ 1 field */
export type CompareMode =
  | "exact" // ต้องตรงกันเป๊ะ (หลัง normalize)
  | "amountTolerance" // ตัวเลข ยอมรับส่วนต่าง +-tolerance
  | "fixedValue" // ค่าคงที่ตาม Requirement (ไม่ได้มาจาก Test Data)
  | "dateWithIdFallback"; // เทียบวันที่ ถ้าไม่ตรงให้ดึงจาก Reference Transaction Number แทน

export interface ReconcileFieldRule {
  /** ชื่อ header ฝั่ง AF1 Report */
  reportField: string;

  /** ชื่อ header ฝั่ง Test Data — null ถ้าไม่มี source (ใช้กับ fixedValue) */
  testDataField: string | null;

  compareMode: CompareMode;

  /** ใช้เมื่อ compareMode = "fixedValue" */
  fixedValue?: string;

  /** ใช้เมื่อ compareMode = "amountTolerance" (ไม่ระบุ = ใช้ DEFAULT_AMOUNT_TOLERANCE) */
  tolerance?: number;

  /**
   * ใช้เมื่อ compareMode = "amountTolerance" — field สำรองฝั่ง Test Data ถ้าเทียบกับ
   * testDataField หลักแล้วไม่ผ่าน ตาม Business Rule: "หาก Transaction Amount ไม่ตรงกับ
   * From Transfer Amount ให้ไปดูที่ From Debit Amount"
   */
  fallbackTestDataField?: string;

  /** true = Core Field ต้องตรวจสอบทุก Test Case, false = Conditional Field */
  isRequiredForAllCases: boolean;

  /**
   * ใช้เมื่อ isRequiredForAllCases = false (Conditional Field)
   * ถ้า field ของ Test Data นี้ว่าง จะข้ามการเปรียบเทียบ field นี้ทั้งแถว
   * ไม่ระบุ = ใช้ testDataField ของตัวเองเป็นตัวเช็คเงื่อนไข
   */
  skipWhenTestDataFieldEmpty?: string;

  /**
   * ใช้เมื่อ isRequiredForAllCases = false (Conditional Field) — เงื่อนไขฝั่ง AF1 Report
   * แทนที่จะเป็นฝั่ง Test Data (ต่างจาก skipWhenTestDataFieldEmpty) ถ้า field นี้ฝั่ง
   * AF1 Report ว่าง จะข้ามการเปรียบเทียบ field นี้ทั้งแถวไปเลย
   *
   * ใช้กับ Inflow Transaction Purpose ตาม Requirement: "ตรวจเฉพาะเมื่อ direction
   * เป็น Inflow ใน AF1 Report" — ตรวจ direction จากการที่ field นี้ (ฝั่ง Report เอง)
   * มีค่าอยู่หรือไม่ ไม่ได้ดูฝั่ง Test Data เลย
   */
  onlyWhenReportFieldHasValue?: string;

  /** หมายเหตุจาก Requirement สำหรับอธิบายกติกาของ field นี้ */
  remark?: string;

  /**
   * จำกัดว่า rule นี้ใช้ตรวจกับแถว suffix ไหนบ้าง (เช่น ["DR"])
   * ไม่ระบุ = ใช้ตรวจกับทุก suffix (ค่าเริ่มต้น)
   *
   * จำเป็นสำหรับ field ที่เกี่ยวกับ "ผู้รับ/ลูกค้า" เช่น Beneficiary Name, Cust Code —
   * เพราะแถว FE ใน AF1 Report ใช้ field กลุ่มนี้เก็บข้อมูล "ธนาคารผู้รับค่าธรรมเนียม"
   * (เช่น BANK OF AYUDHYA) ไม่ใช่ข้อมูลลูกค้าตัวจริงจาก Test Data จึงห้ามเอาไป
   * เทียบกับแถว FE เด็ดขาด มิฉะนั้นจะ False Fail ทุกครั้ง
   */
  applicableSuffixes?: string[];
}

/** ชื่อ header (ทั้ง 2 ฝั่ง) ที่ใช้จับกลุ่ม record ของ Test Case เดียวกัน (ก่อนแยก DR/FE) */
export interface ReconcileGroupKeyFields {
  reportAccountField: string;
  testDataAccountField: string;
  reportCurrencyField: string;
  testDataCurrencyField: string;
}

export interface ReconcileReportConfig {
  headerRowNumber: number;
  testDataHeaderRowNumber: number;

  groupKeyFields: ReconcileGroupKeyFields;

  /** ชื่อ header ของ Report ที่บอกประเภทธุรกรรม (ใช้กรอง DR/FE candidate) */
  withdrawTypeReportField: string;
  /** ค่าที่ withdrawTypeReportField ต้องเป็นเสมอสำหรับแถว DR/FE */
  withdrawTransactionTypeCode: string;

  /** ชื่อ header ของ Report ที่เก็บ Purpose Code (ใช้แยก DR จาก FE) */
  outflowPurposeReportField: string;
  /** ชื่อ header ของ Test Data ที่เก็บ Purpose Code ของแถว DR */
  purposeCodeTestDataField: string;
  /** ค่า Purpose Code คงที่ที่ AF1 ใช้แทนแถว Fee เสมอ (ไม่ได้มาจาก Test Data) */
  feeOutflowPurposeCode: string;

  /** ชื่อ header ของ Report ที่เก็บ Reference Transaction Number (ใช้จับคู่ DR<->FE ด้วยกันหลังเจอ DR แล้ว) */
  referenceNumberReportField: string;
  /** suffix ท้าย Reference Transaction Number ของแถว DR / FE ตามลำดับ */
  drSuffixLabel: string;
  feSuffixLabel: string;

  /**
   * ชื่อ header ของ Test Data ที่เก็บ "Transaction ID/ Reconcile ID" — ใช้เป็น "base"
   * ของ Reference Transaction Number โดยตรง (base + drSuffixLabel = Reference ที่คาดหวัง
   * ของแถว DR, base + feSuffixLabel = ของแถว FE)
   */
  testDataIdField: string;

  /** ชื่อ header ของ Report ที่เก็บยอดเงิน (ใช้เทียบทั้งแถว DR และ FE) */
  transactionAmountReportField: string;
  /** ชื่อ header ของ Test Data ที่เป็นยอดคาดหวังของแถว DR (Transfer Amount) */
  drAmountTestDataField: string;
  /**
   * ชื่อ header สำรอง ใช้เมื่อเทียบกับ drAmountTestDataField แล้วไม่ผ่าน
   * ตาม Business Rule: "หาก Transaction Amount ไม่ตรงกับ From Transfer Amount
   * ให้ไปดูที่ From Debit Amount"
   */
  drAmountFallbackTestDataField: string;

  /** Rule ตัดสินว่ารายการต้องมีหรือห้ามมีใน AF1 Report */
  reportPresenceRules: readonly ReportPresenceRule[];

  fieldRules: ReconcileFieldRule[];
}

export const DEFAULT_AMOUNT_TOLERANCE = 0.01;

/**
 * จำนวนกลุ่ม Fee ที่ใช้ตอนรวมยอด (SUM) สำหรับ Fee Amount Aggregation Logic ของ Reconcile
 * (เดิม expected-case-builder.ts ดึงค่านี้จาก shared/testdata-config.ts ซึ่งตั้งไว้
 * ที่ 2 แต่ Test Data จริงมีถึง Fee Amount Type 5 — แยกค่าออกมาเป็นของ Reconcile เอง)
 */
export const RECONCILE_FEE_TYPE_COUNT = 5;

// ------------------------------------------------------------------
// ค่าคงที่ทางธุรกิจของ DS_LTX (โค้ดตาม BOT ที่ AF1 ต้องรายงานเสมอ)
// แยกเป็นชื่อ const ที่สื่อความหมาย แทนการฝัง magic number ปนในกติกา
// ------------ ต้องแก้ ไม่ fixed ค่าแต่จะมองว่า ข้อมูล Raw report ถูกต้องเสมอ
const RELATIONSHIP_OTHER_CODE = "172064"; // Resident/Non Resident -> อื่นๆ
const WITHDRAW_TRANSACTION_TYPE_CODE = "184010"; // Withdraw (ใช้กับทั้งแถว DR และ FE)
const FEE_OUTFLOW_PURPOSE_CODE = "318029"; // ค่าธรรมเนียมและค่านายหน้าทางด้านการเงิน


/**
 *  เดิม testDataHeaderRowNumber = 5 ถูก Hardcode เปลี่ยนมา Derive จาก
 * report-runtime.config.ts (ที่ Script 2/4 ใช้อยู่แล้ว) ให้เป็น Single Source
 * of Truth จุดเดียว
 */
const LTX_TEST_DATA_HEADER_ROW_NUMBER =
  getReportRuntimeConfig("DS_LTX").testDataHeaderRowNumber;

  /**
 * Partial<Record<ReportCode, ...>> เพื่อให้ TypeScript ช่วยจับ Typo
 * ตอนเพิ่ม Report ใหม่ (เช่น พิมพ์ "DS_PTx" ผิด) ตั้งแต่ Compile Time
 */
export const RECONCILE_CONFIG: Partial<Record<ReportCode, ReconcileReportConfig>> = {
  // ======================================================
  // DS_LTX — verify กับไฟล์ EXPORT_DS_LTX
  // ======================================================
  DS_LTX: {
    headerRowNumber: 1,
    testDataHeaderRowNumber: LTX_TEST_DATA_HEADER_ROW_NUMBER,

    groupKeyFields: {
      reportAccountField: "FI Arrangement Number",
      testDataAccountField: "From Account ( A/C Client/Sender)",
      reportCurrencyField: "Currency Id",
      testDataCurrencyField: "From Currency (CCY)",
    },

    withdrawTypeReportField: "Loan Deposit Transaction Type",
    withdrawTransactionTypeCode: WITHDRAW_TRANSACTION_TYPE_CODE,

    outflowPurposeReportField: "Outflow Transaction Purpose",
    purposeCodeTestDataField: "From BOT Purpose code",
    feeOutflowPurposeCode: FEE_OUTFLOW_PURPOSE_CODE,

    referenceNumberReportField: "Reference Transaction Number",
    drSuffixLabel: "DR",
    feSuffixLabel: "FE",
    testDataIdField: "Transaction ID/ Reconcile ID",

    transactionAmountReportField: "Transaction Amount",
    drAmountTestDataField: "From Transfer Amount",
    drAmountFallbackTestDataField: "From Debit Amount",

    reportPresenceRules: [RESIDENT_THB_TO_FCD_MUST_NOT_EXIST],

    fieldRules: [
      // ---------------- Core Fields (ต้องตรวจสอบทุก Test Case) ----------------
      {
        reportField: "Transaction Date",
        testDataField: "Txn Date",
        compareMode: "dateWithIdFallback",
        isRequiredForAllCases: true,
        remark:
          "ต้องตรงกัน หากไม่ตรงให้ดึงจาก Reference Transaction Number (ตำแหน่ง 7-12)",
      },
      {
        reportField: "Currency Id",
        testDataField: "From Currency (CCY)",
        compareMode: "exact",
        isRequiredForAllCases: true,
        remark: "ต้องตรงกัน ห้ามเปลี่ยนสกุลเงิน",
      },
  
      {
      
        // "Outflow Transaction Purpose ↔ From BOT Purpose code ต้องตรงกัน")
        // ตรวจเฉพาะแถว DR เพราะแถว FE ใช้ Purpose Code คงที่ (feeOutflowPurposeCode) เสมอ
        reportField: "Outflow Transaction Purpose",
        testDataField: "From BOT Purpose code",
        compareMode: "exact",
        isRequiredForAllCases: true,
        applicableSuffixes: ["DR"],
        remark:
          "ตรวจเฉพาะแถว DR (แถว FE ใช้ Purpose Code คงที่ 318029 เสมอ ไม่ได้มาจาก Test Data)",
      },

      // ---------------- Conditional Fields ----------------
      {
        reportField: "Beneficiary or Sender Name",
        testDataField: "To CIF Name (Beneficiary)",
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: ["DR"],
        remark:
          "ตรวจเฉพาะเมื่อคู่สัญญาไม่ใช่นิติบุคคลรับอนุญาต และตรวจเฉพาะแถว DR เท่านั้น (แถว FE เป็นข้อมูลธนาคารผู้รับค่าธรรมเนียม ไม่ใช่ข้อมูลลูกค้า)",
      },
      {
        reportField: "Country Id of Beneficiary or Sender",
        testDataField: "To Region (Bene Country)",
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: ["DR"],
        skipWhenTestDataFieldEmpty: "To CIF Name (Beneficiary)",
        remark: "ตรวจเฉพาะเมื่อ Beneficiary Name มีค่า และตรวจเฉพาะแถว DR",
      },
      {
        reportField: "Relationship with Beneficiary or Sender",
        testDataField: null,
        compareMode: "fixedValue",
        fixedValue: RELATIONSHIP_OTHER_CODE,
        isRequiredForAllCases: false,
        applicableSuffixes: ["DR"],
        skipWhenTestDataFieldEmpty: "To CIF Name (Beneficiary)",
        remark: "Resident/Non Resident -> 172064 (อื่นๆ) ตรวจเฉพาะแถว DR",
      },
      {
        reportField: "Cust Code",
        testDataField: "From CIF No. (Client/Sender)",
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: ["DR"],
        remark:
          "ตรวจเฉพาะเมื่อมีค่าใน Test Data และตรวจเฉพาะแถว DR (ตัวเลขล้วน จะถูกตัด leading zero ก่อนเทียบ)",
      },
      {
        reportField: "Cust Name",
        testDataField: "From CIF Name (Client/Sender)",
        compareMode: "exact",
        isRequiredForAllCases: false,
        applicableSuffixes: ["DR"],
        skipWhenTestDataFieldEmpty: "From CIF No. (Client/Sender)",
        remark: "ตรวจเมื่อ Cust Code มีค่า และตรวจเฉพาะแถว DR",
      },
      {
        reportField: "Inflow Transaction Purpose",
        testDataField: "From BOT Purpose code",
        compareMode: "exact",
        isRequiredForAllCases: false,
        onlyWhenReportFieldHasValue: "Inflow Transaction Purpose",
        remark:
          "ตรวจเฉพาะเมื่อ AF1 Report row นี้ direction เป็น Inflow (มีค่าใน Inflow Transaction Purpose)",
      },
    ],
    // หมายเหตุ (ตามที่ทีมยืนยัน): Installment Number และ Approval Document Number
    // เป็น Conditional Field ตาม Requirement แต่ "ไม่ต้อง reconcile กับ Test Data" —
    // Column ทั้ง 2 นี้จะถูก copy ค่าจริงจาก AF1 Report ไปแสดงในไฟล์ผลลัพธ์โดยอัตโนมัติ
    // ไม่มีแถวไหนถูกไฮไลท์จาก 2 field นี้เด็ดขาด
  },

}; 

/**
 * คืน Config ของ Report ที่พร้อม Reconcile
 */
export const getReconcileConfig = (
  reportCode: string,
): ReconcileReportConfig => {
  const config = RECONCILE_CONFIG[reportCode as ReportCode];

  if (!config || config.groupKeyFields.reportAccountField === "") {
    throw new Error(
      `Reconcile Config ของ Report ` + `"${reportCode}" ยังไม่ครบถ้วน`,
    );
  }

  return config;
};
