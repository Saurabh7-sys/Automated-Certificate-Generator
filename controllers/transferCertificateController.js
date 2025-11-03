const fs = require("fs-extra");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const DocxMerger = require("docx-merger");
const toPdf = require("office-to-pdf");
const ImageModule = require("docxtemplater-image-module-free");
const request = require("sync-request");

exports.generateTransferCertificate = async (req, res) => {
  try {
    await fs.ensureDir(path.join(__dirname, "../output"));
    const { templatePath } = req.body;
    const Data = req.body.Data || req.body.data;

    if (!templatePath || !Data || !Array.isArray(Data)) {
      return res.status(400).json({
        message: "Invalid input. 'templatePath' and 'Data' array are required.",
      });
    }

    const templateFullPath = path.join(__dirname, "..", templatePath);
    const content = await fs.readFile(templateFullPath, "binary");
    const tempDocs = [];

    // Handle image fetch (URL, base64, or local)
    const getImageBuffer = (tagValue) => {
      if (!tagValue) return null;

      // If tagValue is an array -> [url, width, height] or [url]
      if (Array.isArray(tagValue) && tagValue.length > 0) {
        tagValue = tagValue[0];
      }

      let imageUrl = (typeof tagValue === "object" && tagValue.url) ? tagValue.url : tagValue; // handle if object or string
      try {
        if (typeof imageUrl === "string" && /^https?:\/\//.test(imageUrl)) {
          const resReq = request("GET", imageUrl);
          return resReq.getBody();
        }

        if (typeof imageUrl === "string" && imageUrl.startsWith("data:")) {
          const base64 = imageUrl.split(",")[1];
          return Buffer.from(base64, "base64");
        }

        const absPath = path.resolve(process.cwd(), imageUrl);
        if (fs.existsSync(absPath)) return fs.readFileSync(absPath);

        const altPath = path.join(__dirname, "..", imageUrl);
        if (fs.existsSync(altPath)) return fs.readFileSync(altPath);
      } catch {
        return null;
      }

      return null;
    };

    for (let i = 0; i < Data.length; i++) {
      const record = Data[i];
      const zip = new PizZip(content);

      // Image module options closed over current record
      const imageOpts = {
        centered: false,
        // return Buffer for the image. Accepts string URL or object {url, width, height} or array [url, w, h]
        getImage: (tagValue) => getImageBuffer(tagValue),

        // signature: getSize(imgBuffer, tagValue, tagName)
        getSize: (imgBuffer, tagValue, tagName) => {
          // 1) If tagValue is array [url, width, height]
          if (Array.isArray(tagValue) && tagValue.length >= 3) {
            const w = parseInt(tagValue[1], 10);
            const h = parseInt(tagValue[2], 10);
            if (!isNaN(w) && !isNaN(h)) return [w, h];
          }

          // 2) If tagValue is an object { url, width, height }
          if (tagValue && typeof tagValue === "object") {
            const w = parseInt(tagValue.width || tagValue.w, 10);
            const h = parseInt(tagValue.height || tagValue.h, 10);
            if (!isNaN(w) && !isNaN(h)) return [w, h];
          }

          // 3) If user provided separate fields in record:
          //    for tagName 'student_img' this checks 'student_img_width' and 'student_img_height'
          if (tagName && record) {
            const wField = `${tagName}_width`;
            const hField = `${tagName}_height`;
            const w = parseInt(record[wField], 10);
            const h = parseInt(record[hField], 10);
            if (!isNaN(w) && !isNaN(h)) return [w, h];

            // Also check generic fields: 'img_width' / 'img_height' (if you use same name)
            const gw = parseInt(record["img_width"], 10);
            const gh = parseInt(record["img_height"], 10);
            if (!isNaN(gw) && !isNaN(gh)) return [gw, gh];
          }

          // 4) fallback default (pixels)
          return [130, 160];
        },
      };

      const imageModule = new ImageModule(imageOpts);

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        modules: [imageModule],
      });

      doc.setData(record);
      doc.render();

      const buf = doc.getZip().generate({ type: "nodebuffer" });
      const tempPath = path.join(__dirname, `../output/temp_${i}.docx`);
      await fs.writeFile(tempPath, buf);
      const fileBuffer = await fs.readFile(tempPath);
      if (fileBuffer && fileBuffer.length > 0) {
        tempDocs.push(fileBuffer);
      } else {
        console.warn(`⚠️ Skipping empty DOCX at index ${i}`);
      }
    }

    const docx = new DocxMerger({}, tempDocs);
    const mergedBuffer = await new Promise((resolve) => {
      docx.save("nodebuffer", (data) => resolve(data));
    });

    const outputDocx = path.join(__dirname, "../output/Certificates.docx");
    const outputPdf = path.join(__dirname, "../output/Certificates.pdf");
    await fs.writeFile(outputDocx, mergedBuffer);

    const pdfBuffer = await toPdf(mergedBuffer);
    await fs.writeFile(outputPdf, pdfBuffer);

    for (let i = 0; i < Data.length; i++) {
      const tempPath = path.join(__dirname, `../output/temp_${i}.docx`);
      fs.unlink(tempPath).catch(() => { });
    }

    res.setHeader("Content-Type", "application/pdf");
    res.send(pdfBuffer);
  } catch (err) {
    console.error("Error generating certificate:", err);
    res.status(500).json({ message: "Internal server error", error: err.message });
  }
};
