/**
 * Utility กลางสำหรับแปลงค่าที่ใช้ใน Reconcile หลาย Report
 * ------------------------------------------------------------------
 * รวม Logic ที่ DS_LTX, DS_FTU และ DF_OLB ใช้เหมือนกัน
 * เพื่อไม่ให้เกิด Copy-Paste:
 * - Amount
 * - Date
 * - วันที่ YYMMDD ที่ฝังใน Arrangement Number ตำแหน่ง 7-12
 *
 * ไฟล์นี้ทำเฉพาะการแปลงค่า ไม่เก็บ Business Rule ของ Report ใด Report หนึ่ง
 * ------------------------------------------------------------------
 */

/** แปลง Amount เป็น number โดยไม่สน comma; ว่าง/ไม่ใช่ตัวเลขคืน null */
export const parseAmount = (value: unknown): number | null => {
  const normalized = String(value ?? "")
    .replace(/,/g, "")
    .trim();

  if (normalized === "") {
    return null;
  }

  const amount = Number(normalized);
  return Number.isFinite(amount) ? amount : null;
};

/** สร้าง UTC Date พร้อมตรวจว่าปี/เดือน/วันถูกต้องจริง */
const createUtcDate = (
  year: number,
  month: number,
  day: number,
): Date | null => {
  const date = new Date(Date.UTC(year, month - 1, day));

  const isValid =
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day;

  return isValid ? date : null;
};

/**
 * อ่านวันที่จาก Test Data/AF1
 * รองรับเฉพาะ:
 * - Date Object ที่ ExcelJS อ่านจาก Cell รูปแบบวันที่
 * - dd/MM/yyyy
 * - dd-MM-yyyy
 * - yyyy/MM/dd
 * - yyyy-MM-dd
 *
 * ไม่ใช้ Date.parse() เพราะอาจตีความข้อความวันที่ตาม Runtime/Timezone
 * ไม่เหมือนกันในแต่ละเครื่อง
 */
export const parseDate = (value: unknown): Date | null => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return createUtcDate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate(),
    );
  }

  const text = String(value ?? "").trim();

  if (text === "") {
    return null;
  }

  const dayFirst = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
  const yearFirst = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);

  if (dayFirst) {
    return createUtcDate(
      Number(dayFirst[3]),
      Number(dayFirst[2]),
      Number(dayFirst[1]),
    );
  }

  if (yearFirst) {
    return createUtcDate(
      Number(yearFirst[1]),
      Number(yearFirst[2]),
      Number(yearFirst[3]),
    );
  }

  return null;
};

/** อ่านวันที่ตำแหน่ง 7-12 ของ Arrangement Number ในรูปแบบ YYMMDD */
export const extractDateFromArrangementNumber = (
  arrangementNumber: unknown,
): Date | null => {
  const value = String(arrangementNumber ?? "").trim();

  if (value.length < 12) {
    return null;
  }

  const dateText = value.slice(6, 12);

  if (!/^\d{6}$/.test(dateText)) {
    return null;
  }

  return createUtcDate(
    2000 + Number(dateText.slice(0, 2)),
    Number(dateText.slice(2, 4)),
    Number(dateText.slice(4, 6)),
  );
};

/** เปรียบเทียบเฉพาะปี เดือน วัน */
export const isSameDate = (left: Date, right: Date): boolean =>
  left.getUTCFullYear() === right.getUTCFullYear() &&
  left.getUTCMonth() === right.getUTCMonth() &&
  left.getUTCDate() === right.getUTCDate();

/**
 * เปรียบเทียบวันที่ตามลำดับกลางของ LTX, FTU และ OLB
 * 1. เทียบกับ Date Field ของ Report
 * 2. ถ้าไม่ตรง ให้เทียบวันที่ YYMMDD ตำแหน่ง 7-12 ของ Arrangement/Reference Number
 *
 * ผู้เรียกเป็นผู้เลือก Column และตัดสินผล PASS/FAIL/REVIEW ของ Report เอง
 */
export const isDateMatchWithArrangementFallback = (
  expectedDate: Date,
  reportDateValue: unknown,
  arrangementNumber: unknown,
): boolean => {
  const reportDate = parseDate(reportDateValue);

  if (reportDate && isSameDate(expectedDate, reportDate)) {
    return true;
  }

  const arrangementDate =
    extractDateFromArrangementNumber(arrangementNumber);

  return Boolean(
    arrangementDate && isSameDate(expectedDate, arrangementDate),
  );
};

/** แปลง Date เป็น YYYY-MM-DD สำหรับ Remark */
export const formatDate = (value: Date | null): string => {
  if (!value) {
    return "Invalid date";
  }

  const year = value.getUTCFullYear();
  const month = String(value.getUTCMonth() + 1).padStart(2, "0");
  const day = String(value.getUTCDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
};