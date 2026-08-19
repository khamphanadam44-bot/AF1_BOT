/**
 * Normalize ตามกติกาเฉพาะของ DF_OLB
 * แยกไว้เพื่อไม่ให้การเปลี่ยนกติกากระทบ Report อื่น
 */

/** เปรียบเทียบข้อความโดยไม่สนตัวพิมพ์และช่องว่างซ้ำ */
export const normalizeOlbText = (value: unknown): string =>
  String(value ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

/** ตัดเลขศูนย์นำหน้าเมื่อ Cust Code/CIF เป็นตัวเลขทั้งหมด */
export const normalizeOlbId = (value: unknown): string => {
  const normalized = normalizeOlbText(value);

  if (!/^\d+$/.test(normalized)) {
    return normalized;
  }

  return normalized.replace(/^0+(?=\d)/, "");
};