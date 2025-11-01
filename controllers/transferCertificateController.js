const fs = require("fs-extra");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const DocxMerger = require("docx-merger");
const toPdf = require("office-to-pdf");

exports.generateTransferCertificate = async (req, res) => {
  try {
    await fs.ensureDir(path.join(__dirname, "../output"));
    const { templatePath, Data } = req.body;

    if (!templatePath || !Data || !Array.isArray(Data)) {
      return res.status(400).json({
        message: "Invalid input. 'templatePath' and 'Data' array are required.",
      });
    }

    // Load the template
    const templateFullPath = path.join(__dirname, "..", templatePath);
    const content = await fs.readFile(templateFullPath, "binary");
    const tempDocs = [];

    // Generate one filled DOCX per student
    for (let i = 0; i < Data.length; i++) {
      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

      doc.setData(Data[i]);
      doc.render();

      const buf = doc.getZip().generate({ type: "nodebuffer" });
      const tempPath = path.join(__dirname, `../output/temp_${i}.docx`);
      await fs.writeFile(tempPath, buf);
      tempDocs.push(await fs.readFile(tempPath));
    }

    // Merge all DOCX files
    const docx = new DocxMerger({}, tempDocs);
    const mergedBuffer = await new Promise((resolve) => {
      docx.save("nodebuffer", (data) => resolve(data));
    });

    const outputDocx = path.join(__dirname, "../output/TransferCertificate.docx");
    const outputPdf = path.join(__dirname, "../output/TransferCertificate.pdf");

    await fs.writeFile(outputDocx, mergedBuffer);

    // Convert DOCX → PDF using office-to-pdf
    const pdfBuffer = await toPdf(mergedBuffer);
    await fs.writeFile(outputPdf, pdfBuffer);

    // Clean up temp DOCX files
    for (const temp of tempDocs) {
      fs.unlink(temp).catch(() => {});
    }

    // Return PDF preview in Postman
    res.setHeader("Content-Type", "application/pdf");
    res.send(pdfBuffer);
  } catch (err) {
    console.error("Error generating transfer certificate:", err);
    res.status(500).json({ message: "Internal server error", error: err.message });
  }
};
