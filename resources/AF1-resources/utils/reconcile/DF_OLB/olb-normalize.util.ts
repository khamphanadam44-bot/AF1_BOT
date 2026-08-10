/**
 * Normalize ที่เป็นกติกาเฉพาะของ DF_OLB
 * แยกจาก Shared Parser เพื่อไม่ให้ Report อื่นได้รับผลกระทบหากกติกาเปลี่ยน
 */

/** เปรียบเทียบข้อความโดยไม่สนตัวพิมพ์และช่องว่างซ้ำ */
export const normalizeOlbText = (value: unknown): string =>
  String(value ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

/** เปรียบเทียบ Cust Code/CIF โดยตัดเลขศูนย์นำหน้าสำหรับค่าตัวเลข */
export const normalizeOlbId = (value: unknown): string => {
  const normalized = normalizeOlbText(value);

  if (!/^\d+$/.test(normalized)) {
    return normalized;
  }

  const withoutLeadingZero = normalized.replace(/^0+(?=\d)/, "");
  return withoutLeadingZero === "" ? "0" : withoutLeadingZero;
};