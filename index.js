const express = require('express')
const app = express()
const cors = require('cors')


app.use(express.json())
app.use(cors())

//Router
const apiRoute = require("./routes/apiRoute")

app.use("/api" , apiRoute)

app.listen(8080 , () => console.log("Server is Running..."))