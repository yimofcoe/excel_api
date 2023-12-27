const XLSX = require('xlsx')

const router = require('express').Router()
require('dotenv').config()

const excelToJson = (path) => {
    try {

        const workbook = XLSX.readFile(path)
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]

        const jsonData = XLSX.utils.sheet_to_json(worksheet)

        let stringJson = JSON.stringify(jsonData);

        stringJson = stringJson.replace("\"มีความจำเป็นต้อง\":", "\"mustTo\":");
        stringJson = stringJson.replace("\"ชื่อ -สกุล\":", "\"name\":");
        stringJson = stringJson.replace("\"เบอร์โทร\":", "\"phonenum\":");
        stringJson = stringJson.replace("\"ฝ่าย\":", "\"team\":");
        stringJson = stringJson.replace("\"รายการ\":", "\"item\":");
        stringJson = stringJson.replace("\" วงเงิน \":", "\"value\":");
        stringJson = stringJson.replace("\" เหตผล\":", "\"reason\":");
        stringJson = stringJson.replace("\"กรรมการ  TOR + ราคากลาง\":", "\"TORcommitter\":");
        stringJson = stringJson.replace("\"กรรมการจัดซื้อ\":", "\"purchasecommitter\":");
        stringJson = stringJson.replace("\"กรรมการตรวจรับ\":", "\"receivecommitter\":");
        stringJson = stringJson.replace("\"ใบเสนอราคา\":", "\"offer\":");
        stringJson = stringJson.replace("\" TOR\":", "\"TOR\":");
        stringJson = stringJson.replace("\"แหล่งเงิน\":", "\"moneysource\":");
        stringJson = stringJson.replace("\"ผู้อนุมัติแหล่งเงิน\":", "\"moneycommitter\":");
        stringJson = stringJson.replace("\"อีเมลผู้ส่ง\":", "\"email\":");

        const data = JSON.parse(stringJson);

        // console.log(data)

        return data

    } catch (error) {

        console.error("An error occurred: ", error);
        return null;

    }
}

//incoming
router.get("/incoming", (req, res) => {
    const jsonOutput = excelToJson(process.env.INCOMING)
    res.json(jsonOutput)
})

router.get("/incoming/:index", (req, res) => {
    const index = req.params.index
    const jsonOutput = excelToJson(process.env.INCOMING)
    if(jsonOutput[index - 1]){
        res.json(jsonOutput[index - 1])
    }else{
        res.json({})
    }
})

//confirm
router.get("/confirm", (req, res) => {
    const jsonOutput = excelToJson(process.env.CONFIRM)
    res.json(jsonOutput)
})

router.get("/confirm/:index", (req, res) => {
    const index = req.params.index
    const jsonOutput = excelToJson(process.env.CONFIRM)
    if(jsonOutput[index - 1]){
        res.json(jsonOutput[index - 1])
    }else{
        res.json({})
    }
})

//confirm
router.get("/not_confirm", (req, res) => {
    const jsonOutput = excelToJson(process.env.NOT_CONFIRM)
    res.json(jsonOutput)
})

router.get("/not_confirm/:index", (req, res) => {
    const index = req.params.index
    const jsonOutput = excelToJson(process.env.NOT_CONFIRM)
    if(jsonOutput[index - 1]){
        res.json(jsonOutput[index - 1])
    }else{
        res.json({})
    }
})

//money confirm
router.get("/money_confirm", (req, res) => {
    const jsonOutput = excelToJson(process.env.MONEY_CONFIRM)
    res.json(jsonOutput)
})

router.get("/money_confirm/:index", (req, res) => {
    const index = req.params.index
    const jsonOutput = excelToJson(process.env.MONEY_CONFIRM)
    if(jsonOutput[index - 1]){
        res.json(jsonOutput[index - 1])
    }else{
        res.json({})
    }
})




module.exports = router