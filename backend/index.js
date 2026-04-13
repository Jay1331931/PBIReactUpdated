require("dotenv").config();
const express = require("express");
const cors = require("cors");
const { initCronJobs } = require("./cron/index.js");
const app = express();
app.use(express.json()); 
const PORT = process.env.PORT;
const reportRoutes = require("./routes/reportRoutes.js");
initCronJobs();
app.use(cors({
  origin: "*",
  credentials: true
}));
app.use("/api/reports", reportRoutes);

app.listen(PORT, () => {
  console.log(`Backend running on port ${PORT}`);
});