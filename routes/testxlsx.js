const XLSX = require("xlsx");

const workbook = XLSX.readFile("C:/Users/srp33/OneDrive - Prince of Songkla University/study_file/66_1/240-401/ฟอร์ม/รายการคำร้องที่ได้รับการอนุมัติแล้ว.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const jsonData = XLSX.utils.sheet_to_json(worksheet);

console.log(jsonData[0].กรรมการจัดซื้อ)

const stringJson = JSON.stringify(jsonData).replace("\" TOR\":", "\"mustTo\":");

const data = JSON.parse(stringJson)

console.log(data[0].mustTo);
