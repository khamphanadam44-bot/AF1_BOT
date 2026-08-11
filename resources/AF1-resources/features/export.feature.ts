/**
 * ======================================================
 * ไฟล์: export.feature.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * ควบคุมลำดับงาน Export Report ตั้งแต่เปิดเมนู เลือก Report ใส่ช่วงวันที่ และดาวน์โหลดไฟล์
 * หลังดาวน์โหลดเสร็จจะอ่าน Report แล้วส่ง Path กับข้อมูลกลับไปให้ Script 1
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import {
  Page,
} from "playwright";

import {
  MainPage,
} from "../pages/main/main.page";

import {
  readExcel,
} from "../utils/excel-reader";


export type ExportResult = {
  savePath: string;
  reportData: unknown[];
};

export class ExportFeature {
  mainPage: MainPage;

  constructor(
    page: Page,
  ) {
    this.mainPage =
      new MainPage(
        page,
      );
  }

  async exportPass(
    dmsReportName: string,
    dateFrom: string,
    dateTo: string,
  ): Promise<ExportResult> {
    await this.mainPage.clickDmsMenu();

    await this.mainPage.clickDmmMenu();

    await this.mainPage.clickDataSetMenu();

    await this.mainPage.inputReportName(
      dmsReportName,
    );

    /**
     * ส่งชื่อ Report เข้าไปให้ main.page.ts
     *
     * ตัวอย่าง:
     * DS_PTX
     * DS_FTX
     */
    await this.mainPage.clickDSReport(
      dmsReportName,
    );

    await this.mainPage.inputDate(
      dateFrom,
    );

    await this.mainPage.inputDateTo(
      dateTo,
    );

    const savePath =
      await this.mainPage.clickExportButton(
        dmsReportName,
      );

    const reportData =
      await readExcel(
        savePath,
      );

    return {
      savePath,
      reportData,
     
    };
  }
}