/**
 * ftu-parse.util.ts
 * ------------------------------------------------------------------
 * Bug fix (Code Review): เดิม normalize และ parseAmount ถูก Copy-Paste ซ้ำ
 * เป๊ะทั้งใน ftu-rules.ts และ ftu-reconcile.ts — รวมเป็นไฟล์เดียวที่นี่
 * ให้ทั้งสองไฟล์ Import ไปใช้ร่วมกันแทนการเขียนซ้ำ
 * ------------------------------------------------------------------
 */

/** normalize ค่าทั่วไปเป็น string ตัวพิมพ์ใหญ่ ตัดช่องว่างหัวท้าย */
export const normalize = (value: unknown): string =>
  String(value ?? "")
    .trim()
    .toUpperCase();

/** แปลง string เป็นตัวเลข (ตัด comma คั่นหลักพัน) คืนค่า null ถ้าแปลงไม่ได้/ว่าง */
export const parseAmount = (value: unknown): number | null => {
  const normalizedValue = String(value ?? "")
    .replace(/,/g, "")
    .trim();

  if (normalizedValue === "") {
    return null;
  }

  const amount = Number(normalizedValue);
  return Number.isFinite(amount) ? amount : null;
};
