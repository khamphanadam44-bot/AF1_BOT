/**
 * ======================================================
 * ไฟล์: run-selected-test.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * เป็นตัวกลางสำหรับเปิดไฟล์ Test ที่ระบุใน package.json แล้วสั่งให้ Mocha เริ่มทำงาน
 * ช่วยให้ Script 1–3 ใช้รูปแบบการรันเดียวกัน และรองรับการส่งค่า report จาก Terminal ต่อไปยังไฟล์ Test
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import * as path from "path";

import Mocha from "mocha";

/**
 * ฟังก์ชันหลักสำหรับเปิดและรันไฟล์ Test
 *
 * เรานำ await มาไว้ภายใน async function
 * เพื่อให้สามารถทำงานกับโปรเจกต์แบบ CommonJS ได้
 */
const main = async (): Promise<void> => {
  /**
   * รับชื่อไฟล์ Test จาก package.json
   *
   * ตัวอย่าง:
   * tests/script1-login-export.spec.ts
   */
  const testFile =
    process.argv[2];

  /**
   * ตรวจว่ามีชื่อไฟล์ Test ส่งเข้ามาหรือไม่
   */
  if (
    !testFile ||
    testFile.startsWith("--")
  ) {
    throw new Error(
      "ไม่พบชื่อไฟล์ Test ที่ต้องการรัน",
    );
  }

  /**
   * สร้าง Mocha
   */
  const mocha =
    new Mocha({
      color: true,
    });

  /**
   * เพิ่มไฟล์ Test ที่ต้องการรัน
   */
  mocha.addFile(
    path.resolve(testFile),
  );

  /**
   * โหลดไฟล์ Test
   *
   * await อยู่ภายใน async function แล้ว
   * จึงไม่เกิด Top-level await
   */
  await mocha.loadFilesAsync();

  /**
   * เริ่มรัน Test
   */
  mocha.run(
    (failureCount) => {
      process.exitCode =
        failureCount > 0
          ? 1
          : 0;
    },
  );
};

/**
 * เรียกฟังก์ชันหลัก
 *
 * ถ้าเกิด Error จะแสดงรายละเอียดใน Terminal
 * และกำหนด Exit Code เป็น 1
 */
main().catch(
  (error: unknown) => {
    console.error(
      "ไม่สามารถเริ่มการทดสอบได้:",
      error,
    );

    process.exitCode = 1;
  },
);