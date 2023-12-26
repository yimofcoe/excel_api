const XLSX = require('xlsx')

const router = require('express').Router()
require('dotenv').config()

const excelToJson = (path) => {
    try {

        const workbook = XLSX.readFile(path)
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]

        const jsonData = XLSX.utils.sheet_to_json(worksheet)

        return jsonData

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