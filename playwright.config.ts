/**
 * ======================================================
 * ไฟล์: playwright.config.ts
 * ======================================================
 *
 * หน้าที่ของไฟล์นี้
 *
 * เก็บค่ากลางของ Playwright เช่น โฟลเดอร์ Test, Timeout, Browser และรูปแบบ Report
 * ค่าชุดนี้ใช้เมื่อรันงานผ่าน Playwright Test Runner และไม่ได้เป็น Logic ทางธุรกิจของ AF1
 *
 * หมายเหตุ:
 * ไฟล์นี้อธิบายหน้าที่ของ Code เท่านั้น การแก้ Comment ไม่มีผลต่อการทำงานของระบบ
 * ======================================================
 */
import { defineConfig, devices } from "@playwright/test";

/**
 * อ่านค่า Environment Variable จากไฟล์
 * https://github.com/motdotla/dotenv
 */
// import dotenv from "dotenv";
// import path from "path";
// dotenv.config({ path: path.resolve(__dirname, ".env") });

/**
 * See https://playwright.dev/docs/test-configuration.
 */
export default defineConfig({
  testDir: "./tests",

  /**
   * ให้ Playwright อ่านเฉพาะไฟล์ที่ลงท้ายด้วย .pw.spec.ts
   *
   * ตัวอย่าง:
   * login.pw.spec.ts
   *
   * ส่วนไฟล์ Mocha เช่น:
   * script1-login-export.spec.ts
   * script2-validate-report-and-testdata.spec.ts
   * script3-compare-report.spec.ts
   *
   * จะไม่ถูก Playwright Extension นำไปรัน
   */
  testMatch: "**/*.pw.spec.ts",

  /* Run tests in files in parallel */
  fullyParallel: true,

  /* Fail the build on CI if you accidentally left test.only in the source code. */
  forbidOnly: !!process.env.CI,

  /* Retry on CI only */
  retries: process.env.CI ? 2 : 0,

  /* Opt out of parallel tests on CI. */
  workers: process.env.CI ? 1 : undefined,

  /* Reporter to use. See https://playwright.dev/docs/test-reporters */
  reporter: "html",

  /* Shared settings for all the projects below. See https://playwright.dev/docs/api/class-testoptions. */
  use: {
    /* Base URL to use in actions like `await page.goto('')`. */
    // baseURL: "http://localhost:3000",

    /* Collect trace when retrying the failed test. See https://playwright.dev/docs/trace-viewer */
    trace: "on-first-retry",
  },

  /* Configure projects for major browsers */
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },

    {
      name: "firefox",
      use: { ...devices["Desktop Firefox"] },
    },

    {
      name: "webkit",
      use: { ...devices["Desktop Safari"] },
    },

    /* Test against mobile viewports. */
    // {
    //   name: "Mobile Chrome",
    //   use: { ...devices["Pixel 5"] },
    // },
    // {
    //   name: "Mobile Safari",
    //   use: { ...devices["iPhone 12"] },
    // },

    /* Test against branded browsers. */
    // {
    //   name: "Microsoft Edge",
    //   use: { ...devices["Desktop Edge"], channel: "msedge" },
    // },
    // {
    //   name: "Google Chrome",
    //   use: { ...devices["Desktop Chrome"], channel: "chrome" },
    // },
  ],

  /* Run your local dev server before starting the tests */
  // webServer: {
  //   command: "npm run start",
  //   url: "http://localhost:3000",
  //   reuseExistingServer: !process.env.CI,
  // },
});