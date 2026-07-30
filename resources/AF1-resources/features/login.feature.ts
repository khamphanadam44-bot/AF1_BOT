/**
 * ======================================================
 * ไฟล์: login.feature.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * รวมขั้นตอน Login ให้ Script 1 เรียกใช้งานได้ด้วยคำสั่งเดียว
 * ไฟล์นี้ส่ง Username และ Password ไปให้ login.page.ts เป็นผู้กรอกและกดปุ่มเข้าสู่ระบบ
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import {
  Page,
} from "playwright";

import {
  LoginPage,
} from "../pages/login/login.page";

export class LoginFeature {
  private readonly loginPage: LoginPage;

  constructor(
    page: Page,
  ) {

    this.loginPage =
      new LoginPage(
        page,
      );

  }

  async loginPass(
    username: string,
    password: string,
  ): Promise<void> {

    await this.loginPage.inputUsername(
      username,
    );

    await this.loginPage.inputPassword(
      password,
    );

    await this.loginPage.clickSubmit();

  }
}