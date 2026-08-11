# AF1 Testcase

โปรเจกต์ Automation สำหรับ Export, Validate, Reconcile และสร้าง Summary ของ AF1 Report

## Report ที่รองรับ

Report ที่เชื่อมการทำงานกับ Script 1–4 แล้ว:

- `DS_LTX`
- `DS_PTX`
- `DS_FTX`
- `DS_FTU`
- `DF_FXU`
- `DF_OLB`
- `DF_FXM`

## การเตรียมโปรเจกต์

ติดตั้ง Dependency:

```bash
npm ci
```

คัดลอก `.env.example` เป็น `.env`:

```powershell
Copy-Item .env.example .env
```

จากนั้นกำหนดค่าที่ใช้กับ Environment ของผู้ใช้งานใน `.env`:

```dotenv
AF1_UAT_URL=
AF1_UAT_USERNAME=
AF1_UAT_PASSWORD=
AF1_SHAREPATH=
```

ห้าม Commit ไฟล์ `.env` เนื่องจากอาจมี URL, Username, Password และตำแหน่ง Network Share จริง

## ตำแหน่ง Test Data

ระบบค้นหา Test Data จาก:

```text
AF1_SHAREPATH/af1_test_data/<REPORT>
```

ตัวอย่างโครงสร้าง:

```text
AF1_SHAREPATH/
  af1_test_data/
    DS_LTX/
    DS_PTX/
    DS_FTX/
    DS_FTU/
    DF_FXU/
    DF_OLB/
    DF_FXM/
```

แต่ละโฟลเดอร์ Report ต้องมีไฟล์ Excel Test Data ที่ต้องการใช้งานเพียง 1 ไฟล์

## การรันแต่ละ Script

ตัวอย่างการรัน Report เดียว:

```bash
npm run test:script1 -- report=DS_LTX
npm run test:script2 -- report=DS_LTX
npm run test:script3 -- report=DS_LTX
npm run test:script4 -- report=DS_LTX
```

สามารถเลือกหลาย Report โดยคั่นชื่อด้วยเครื่องหมายจุลภาค:

```bash
npm run test:script1 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU,DF_OLB,DF_FXM
npm run test:script2 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU,DF_OLB,DF_FXM
npm run test:script3 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU,DF_OLB,DF_FXM
npm run test:script4 -- report=DS_LTX,DS_PTX,DS_FTX,DS_FTU,DF_FXU,DF_OLB,DF_FXM
```

## ไฟล์ผลลัพธ์

ไฟล์ที่ Export จาก Script 1:

```text
test_data/AF1_Report/<REPORT>
```

ไฟล์ที่ตรวจสอบ Header แล้วจาก Script 2:

```text
Test_result/Checked-report-header/<REPORT>
Test_result/Checked-testdata-header/<REPORT>
```

ไฟล์ Reconcile จาก Script 3:

```text
Test_result/Reconcile-report/<REPORT>
```

ไฟล์ Summary จาก Script 4:

```text
Test_result/Summary-report/<REPORT>
```

## ข้อมูลที่ไม่เก็บใน Git

รายการต่อไปนี้ถูกกำหนดไว้ใน `.gitignore`:

- `.env`
- `test_data/`
- `Test_result/`
- ไฟล์ `.zip`

เพื่อป้องกันข้อมูลจริง ข้อมูลสำหรับทดสอบ และไฟล์ผลลัพธ์ถูก Commit โดยไม่ตั้งใจ