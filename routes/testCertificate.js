const express = require("express");
const router = express.Router();
const controller = require("../controllers/certificateController");


router.post("/testGenCertificate", controller.generateCertificate);

module.exports = router