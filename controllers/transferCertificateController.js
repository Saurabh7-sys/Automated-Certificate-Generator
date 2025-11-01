const fs = require("fs-extra");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const DocxMerger = require("docx-merger");
const toPdf = require("office-to-pdf");
const ImageModule = require("docxtemplater-image-module-free");

exports.generateTransferCertificate  = async (req, res) => {
  try {
    await fs.ensureDir(path.join(__dirname, "../output"));
    const { templatePath, Data } = req.body;

    if (!templatePath || !Data || !Array.isArray(Data)) {
      return res.status(400).json({
        message: "Invalid input. 'templatePath' and 'Data' array are required.",
      });
    }

    // Load the DOCX template
    const templateFullPath = path.join(__dirname, "..", templatePath);
    const content = await fs.readFile(templateFullPath, "binary");
    const tempDocs = [];

    // Generate one filled DOCX per record
    for (let i = 0; i < Data.length; i++) {
      const zip = new PizZip(content);

      // Configure image module
      const imageOpts = {
        centered: false,
        getImage: function (tagValue) {
          // tagValue can be a local path or base64 string
          if (!tagValue) return null;
          const imagePath = path.join(__dirname, "..", tagValue);
          return fs.readFileSync(imagePath);
        },
        getSize: function () {
          // width, height in pixels
          return [130, 160];
        },
      };

      const imageModule = new ImageModule(imageOpts);

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        modules: [imageModule],
      });

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

    const outputDocx = path.join(__dirname, "../output/Certificates.docx");
    const outputPdf = path.join(__dirname, "../output/Certificates.pdf");

    await fs.writeFile(outputDocx, mergedBuffer);

    // Convert DOCX → PDF
    const pdfBuffer = await toPdf(mergedBuffer);
    await fs.writeFile(outputPdf, pdfBuffer);

    // Clean up temp DOCX files
    for (let i = 0; i < Data.length; i++) {
      const tempPath = path.join(__dirname, `../output/temp_${i}.docx`);
      fs.unlink(tempPath).catch(() => {});
    }

    // Send generated PDF
    res.setHeader("Content-Type", "application/pdf");
    res.send(pdfBuffer);
  } catch (err) {
    console.error("Error generating certificate:", err);
    res.status(500).json({ message: "Internal server error", error: err.message });
  }
};
