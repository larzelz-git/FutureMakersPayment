# Live Sync Setup

ถ้าชีตเป็น private หน้า `file://` จะอ่านจาก Google Sheet ตรงๆ ไม่ได้เพราะไม่มีสิทธิ์ OAuth ใน browser

วิธีที่ใช้ได้จริง:

1. เปิด [Apps Script](https://script.google.com/)
2. สร้างโปรเจกต์ใหม่
3. วางโค้ดจาก `summary-dashboard/apps-script/summary-api.gs`
4. Deploy > New deployment > Web app
5. ตั้งค่า:
   - Execute as: `Me`
   - Who has access: `Anyone with the link` หรือคนในองค์กรของคุณ
6. คัดลอก Web App URL
7. เอา URL ไปใส่ใน `summary-dashboard/summary-data.js` ที่ `liveJsonUrl`

เมื่อใส่แล้ว หน้า Dashboard จะดึงข้อมูลสดผ่าน URL นี้ก่อนเสมอ

## Central Source Config

หลังจาก deploy Apps Script เวอร์ชันนี้แล้ว ปุ่ม `จัดการ Data Source` ในหน้า Dashboard จะบันทึกค่า source กลางไว้ใน Apps Script ให้เอง

- แก้ `Spreadsheet Link` หรือ `Sheet Name` จากหน้าเว็บได้เลย
- ทุกเครื่องที่เปิด dashboard ผ่าน URL เดียวกันจะใช้ source กลางอันเดียวกัน
- ไม่ต้อง deploy หน้าเว็บใหม่ทุกครั้งที่เปลี่ยน Google Sheet
- ถ้า `liveJsonUrl` ยังว่าง ระบบจะ fallback ไป save เฉพาะ browser เดิมด้วย `localStorage`
