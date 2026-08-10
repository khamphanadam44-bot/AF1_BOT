/**
 * olb-parse.util.ts
 * ------------------------------------------------------------------
 * ฟังก์ชัน Normalize สำหรับ DF_OLB
 * - Text/Identifier
 * - Amount
 * - Date
 * - วันที่ YYMMDD ที่ฝังอยู่ใน FI Arrangement Number ตำแหน่ง 7-12
 * ------------------------------------------------------------------
 */

/** Normalize ข้อความทั่วไปสำหรับเปรียบเทียบแบบไม่สนตัวพิมพ์และช่องว่างซ้ำ */
export const normalizeText = (value: unknown): string =>
  String(value ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

/** Normalize รหัสตัวเลข เช่น CIF โดยตัดเลขศูนย์นำหน้า */
export const normalizeId = (value: unknown): string => {
  const normalized = normalizeText(value);

  if (!/^\d+$/.test(normalized)) {
    return normalized;
  }

  const withoutLeadingZero = normalized.replace(/^0+(?=\d)/, "");
  return withoutLeadingZero === "" ? "0" : withoutLeadingZero;
};

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
 * อ่านวันที่จากรูปแบบที่ใช้ใน Test Data/AF1
 * รองรับ dd/MM/yyyy, dd-MM-yyyy, yyyy/MM/dd, yyyy-MM-dd
 * และ Date string ที่ ExcelJS แปลงมาจาก Date Object
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

  const parsedTimestamp = Date.parse(text);

  if (Number.isNaN(parsedTimestamp)) {
    return null;
  }

  const parsedDate = new Date(parsedTimestamp);

  return createUtcDate(
    parsedDate.getFullYear(),
    parsedDate.getMonth() + 1,
    parsedDate.getDate(),
  );
};

/** อ่านวันที่ตำแหน่ง 7-12 ของ FI Arrangement Number ในรูปแบบ YYMMDD */
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

export const formatDate = (value: Date | null): string => {
  if (!value) {
    return "Invalid date";
  }

  const year = value.getUTCFullYear();
  const month = String(value.getUTCMonth() + 1).padStart(2, "0");
  const day = String(value.getUTCDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
};
