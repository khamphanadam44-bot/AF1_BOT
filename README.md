# AF1 Testcase

โปรเจกต์ Automation สำหรับ Export, Validate, Reconcile และสร้าง Summary ของ AF1 Report

Report ที่เชื่อมการทำงาน Script 1-4 แล้ว:

- `DS_LTX`
- `DS_PTX`
- `DS_FTX`
- `DS_FTU`

## การเตรียมโปรเจกต์

```bash
npm install
```

คัดลอก `.env.example` เป็น `.env` แล้วกำหนดข้อมูล UAT ของเครื่องผู้ใช้งาน

ไฟล์ Test Data จริงไม่ถูกเก็บใน Git ให้นำมาวางที่:

```text
test_data/Test_Data_Downstream-for pilot.xlsx
```

## ตัวอย่างการรัน

```bash
npm run test:script1 -- report=DS_LTX
npm run test:script2 -- report=DS_LTX
npm run test:script3 -- report=DS_LTX
npm run test:script4 -- report=DS_LTX
```

สามารถเลือกหลาย Report ได้:

```bash
npm run test:script1 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU
npm run test:script2 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU
npm run test:script3 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU
npm run test:script4 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU
```

โฟลเดอร์ `test_data` และ `Test_result` ถูกตัดออกจาก Git เพื่อป้องกันข้อมูลจริงและไฟล์ผลลัพธ์ถูก Commit โดยไม่ตั้งใจ
