const express = require("express");
const reportController = require("../controllers/reportController.js");

const router = express.Router();

router.get("/getEmbedToken", reportController.getEmbedToken);
router.post("/exportToAzure", reportController.exportToAzure);

module.exports = router;