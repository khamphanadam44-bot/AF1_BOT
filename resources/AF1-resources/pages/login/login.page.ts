/**
 * ======================================================
 * ไฟล์: login.page.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * จัดการการกระทำบนหน้า Login เช่น กรอก Username กรอก Password และกดปุ่ม Sign in
 * ไฟล์นี้ใช้ Locator จาก login.locator.ts และถูกเรียกต่อโดย login.feature.ts
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import {
  Page,
} from "playwright";

import {
  loginLocator,
} from "./login.locator";

export class LoginPage {
  private readonly page: Page;

  constructor(
    page: Page,
  ) {
    this.page = page;
  }

  async inputUsername(
    username: string,
  ): Promise<void> {

    const usernameInput =
      this.page.locator(
        loginLocator.usernametxt,
      );

    await usernameInput.fill(
      username,
    );

  }

  async inputPassword(
    password: string,
  ): Promise<void> {

    const passwordInput =
      this.page.locator(
        loginLocator.passwordtxt,
      );

    await passwordInput.fill(
      password,
    );

  }

  async clickSubmit(): Promise<void> {

    const submitButton =
      this.page.locator(
        loginLocator.submitbtn,
      );

    await submitButton.click();

  }
}