// app.js — contractRoutes холбох жишээ

const express = require("express");
const app = express();

app.use(express.json());

const contractRoutes = require("./contractRoutes");
app.use("/api", contractRoutes);

app.listen(3000, () => console.log("Server running on port 3000"));
