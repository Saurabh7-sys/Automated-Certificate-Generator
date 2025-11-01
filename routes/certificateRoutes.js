const express = require('express');
const router = express.Router();
const controller = require('../controllers/transferCertificateController');

router.post('/generateTransferCertificate', controller.generateTransferCertificate);

module.exports = router;
