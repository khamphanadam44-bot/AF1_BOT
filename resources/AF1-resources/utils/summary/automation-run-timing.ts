/**
 * automation-run-timing.ts
 * ------------------------------------------------------------------
 * เก็บเวลาทำงานของ Script 1-4 แยกตาม Report
 *
 * หลักการคำนวณ Duration:
 * - นับเฉพาะเวลาที่แต่ละ Script ทำงานให้ Report นั้นจริง
 * - ไม่รวมเวลาที่ผู้ใช้งานหยุดพักระหว่างคำสั่ง
 * - ไม่รวมเวลาที่ระบบกำลังประมวลผล Report อื่น
 *
 * Timing State จะถูกเก็บที่:
 * Test_result/Automation-run-time/<REPORT>.json
 * ------------------------------------------------------------------
 */

import fs from "fs";
import path from "path";

/** ขั้นตอนที่ต้องจับเวลาใน Automation Pipeline */
export type AutomationRunStage =
  | "SCRIPT_1_EXPORT"
  | "SCRIPT_2_REPORT_HEADER"
  | "SCRIPT_2_TEST_DATA"
  | "SCRIPT_3_COMPARE"
  | "SCRIPT_4_SUMMARY";

/** เวลาเริ่ม จบ และระยะเวลาของหนึ่งขั้นตอน */
type AutomationStageTiming = {
  startedAt: string;
  completedAt: string;
  durationMilliseconds: number;
};

/** Timing State ของหนึ่ง Report */
type AutomationRunTimingState = {
  reportCode: string;
  runId: string;
  automationStartedAt: string;
  stages: Partial<
    Record<
      AutomationRunStage,
      AutomationStageTiming
    >
  >;
  updatedAt: string;
};

/** ข้อมูลที่ Script 4 ใช้สร้าง Summary */
export type AutomationRunTimingSummary = {
  runId: string;
  automationStartedAt: Date;
  completedStageDurationMilliseconds: number;
};

/**
 * Folder สำหรับเก็บไฟล์ JSON จับเวลาของ Script 1-4
 *
 * ตัวอย่าง:
 * resources/AF1-resources/utils/summary/
 * Automation-run-time/DS_FTX.json
 */
const TIMING_ROOT_FOLDER =
  path.resolve(
    process.cwd(),
    "resources",
    "AF1-resources",
    "utils",
    "summary",
    "Automation-run-time",
  );

/** แปลงชื่อ Report เป็นรูปแบบมาตรฐาน */
const normalizeReportCode = (
  reportName: string,
): string =>
  String(reportName)
    .trim()
    .toUpperCase()
    .replace(/-/g, "_");

/** สร้าง Timestamp สำหรับ Run ID */
const formatRunTimestamp = (
  date: Date,
): string => {
  const yyyy = date.getFullYear();
  const MM = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const HH = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");

  return `${yyyy}${MM}${dd}_${HH}${mm}${ss}`;
};

/** Path ของ Timing State แยกตาม Report */
const getTimingFilePath = (
  reportName: string,
): string =>
  path.join(
    TIMING_ROOT_FOLDER,
    `${normalizeReportCode(reportName)}.json`,
  );

/** บันทึก Timing State ลงไฟล์ */
const writeTimingState = (
  reportName: string,
  state: AutomationRunTimingState,
): void => {
  fs.mkdirSync(
    TIMING_ROOT_FOLDER,
    {
      recursive: true,
    },
  );

  fs.writeFileSync(
    getTimingFilePath(reportName),
    JSON.stringify(state, null, 2),
    "utf8",
  );
};

/** อ่าน Timing State และตรวจสอบรูปแบบข้อมูลเบื้องต้น */
const readTimingState = (
  reportName: string,
): AutomationRunTimingState => {
  const timingFilePath =
    getTimingFilePath(reportName);

  if (!fs.existsSync(timingFilePath)) {
    throw new Error(
      `Automation timing not found for ${normalizeReportCode(reportName)}. ` +
      `กรุณารัน Script 1 ก่อน Script 2-4`,
    );
  }

  const state =
    JSON.parse(
      fs.readFileSync(
        timingFilePath,
        "utf8",
      ),
    ) as AutomationRunTimingState;

  if (
    state.reportCode !==
      normalizeReportCode(reportName) ||
    typeof state.runId !== "string" ||
    typeof state.automationStartedAt !== "string" ||
    typeof state.stages !== "object" ||
    state.stages === null
  ) {
    throw new Error(
      `Invalid automation timing file: ${timingFilePath}`,
    );
  }

  return state;
};

/**
 * เริ่ม Automation Run ใหม่สำหรับหนึ่ง Report
 *
 * เรียกจาก Script 1 ก่อนเริ่ม Export
 * การเรียกฟังก์ชันนี้จะ Reset เวลา Run ก่อนหน้าของ Report เดียวกัน
 */
export const startAutomationRun = (
  reportName: string,
  startedAt: Date,
): void => {
  const reportCode =
    normalizeReportCode(reportName);

  const state: AutomationRunTimingState = {
    reportCode,
    runId: `RUN_${formatRunTimestamp(startedAt)}`,
    automationStartedAt: startedAt.toISOString(),
    stages: {},
    updatedAt: startedAt.toISOString(),
  };

  writeTimingState(
    reportCode,
    state,
  );

  console.log(
    `Automation Timer Started: ${reportCode} | ${state.runId}`,
  );
};

/**
 * บันทึกระยะเวลาของหนึ่งขั้นตอน
 *
 * ถ้ารันขั้นตอนเดิมซ้ำ ระบบจะใช้เวลารอบล่าสุดแทนรอบเดิม
 * เพื่อไม่ให้ Duration รวมเวลาจากรอบที่ Error ซ้ำเข้าไป
 */
export const recordAutomationRunStage = (
  reportName: string,
  stage: AutomationRunStage,
  startedAt: Date,
  completedAt: Date,
): void => {
  const state =
    readTimingState(reportName);

  const durationMilliseconds =
    Math.max(
      0,
      completedAt.getTime() -
        startedAt.getTime(),
    );

  state.stages[stage] = {
    startedAt: startedAt.toISOString(),
    completedAt: completedAt.toISOString(),
    durationMilliseconds,
  };

  state.updatedAt =
    completedAt.toISOString();

  writeTimingState(
    reportName,
    state,
  );

  console.log(
    [
      "Automation Stage Time",
      normalizeReportCode(reportName),
      stage,
      `${durationMilliseconds} ms`,
    ].join(" | "),
  );
};

/**
 * อ่านเวลารวมของ Script 1-3 สำหรับส่งให้ Script 4
 *
 * Script 4 จะนำเวลาของตัวเองมาบวกเพิ่มก่อนเขียน Duration ลง Summary
 */
export const getAutomationRunTimingSummary = (
  reportName: string,
): AutomationRunTimingSummary => {
  const state =
    readTimingState(reportName);

  const requiredStages: readonly AutomationRunStage[] = [
    "SCRIPT_1_EXPORT",
    "SCRIPT_2_REPORT_HEADER",
    "SCRIPT_2_TEST_DATA",
    "SCRIPT_3_COMPARE",
  ];

  const missingStages =
    requiredStages.filter(
      (stage) => !state.stages[stage],
    );

  if (missingStages.length > 0) {
    throw new Error(
      `Automation timing incomplete for ${state.reportCode}. ` +
      `Missing: ${missingStages.join(", ")}`,
    );
  }

  const completedStageDurationMilliseconds =
    requiredStages.reduce(
      (total, stage) =>
        total +
        (state.stages[stage]
          ?.durationMilliseconds ?? 0),
      0,
    );

  const automationStartedAt =
    new Date(state.automationStartedAt);

  if (
    Number.isNaN(
      automationStartedAt.getTime(),
    )
  ) {
    throw new Error(
      `Invalid automation start time for ${state.reportCode}`,
    );
  }

  return {
    runId: state.runId,
    automationStartedAt,
    completedStageDurationMilliseconds,
  };
};