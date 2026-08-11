/**
 * ======================================================
 * ไฟล์: main.locator.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * เก็บ Locator ของเมนู ช่องกรอกชื่อ Report ช่องวันที่
 * และปุ่ม Export บนหน้าหลัก
 * ======================================================
 */

export const mainLocator = {
  /**
   * เมนู DMS
   */
  dmsMenu:
    '[data-link="dms"]',

  /**
   * เมนู DMM Data Model Export
   */
  dmmDataModelExport:
    "#acc-6",

  /**
   * เมนู DATASET EXPORT
   */
  dataSetExport:
    "div.ds-name:has-text('DATASET EXPORT')",

  /**
   * ช่องกรอกชื่อ Report
   */
  reportNameInput:
    'input[ng-model="reportName"]',

/**
 * Locator กลางสำหรับ Report ที่ชื่อบนหน้าเว็บ
 * ตรงกับ Report Code ที่รับเข้ามา
 *
 * ใช้กับ:
 * - DS_PTX
 * - DS_FTX
 * - DS_LTX
 * - DF_OLB
 * - DF_FXM
 *
 * ตัวอย่าง:
 * reportByName("DF_FXM")
 *
 * จะได้ Locator:
 * td.text-left:text-is("DF_FXM")
 *
 * ไม่ใช้กับ DS_FTU และ DF_FXU
 * เพราะชื่อที่แสดงบนหน้าเว็บมีคำว่า
 * "Transaction" ต่อท้าย
 */
  reportByName: (
    reportName: string,
  ): string =>
    `td.text-left:text-is("${reportName}")`,

  /**
   * Locator เฉพาะของ DS_FTU
   *
   * เนื่องจากชื่อที่แสดงบนหน้าจอคือ:
   * DS_FTU Transaction
   *
   * จึงไม่สามารถใช้ reportByName("DS_FTU")
   * แบบเดียวกับ Report อื่นได้
   */
  dsFtuReport:
    "//td[contains(text(), 'DS_FTU Transaction')]",

  /**
 * Locator เฉพาะของ DF_FXU
 *
 * ชื่อที่แสดงบนหน้าเว็บคือ:
 * DF_FXU Transaction
 */
  dfFxuReport:
    "//td[text()='DF_FXU Transaction']",


  /**
   * ช่อง Date From
   */
  dateFromTxt:
    'input[name="input-1"][placeholder="dd/mm/yyyy"]',

  /**
   * ช่อง Date To
   */
  dateToTxt:
    'input[name="inputT-1"]',

  /**
   * ปุ่ม Export
   */
  exportButton:
    'button[data-ng-click="submitFilter()"]',
};