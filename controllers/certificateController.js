// controllers/certificateController.js
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const { execFile } = require("child_process");
const { promisify } = require("util");
const axios = require("axios");
const { PDFDocument } = require("pdf-lib");

const execFileAsync = promisify(execFile);
// update this path if your LibreOffice is elsewhere
const sofficePath = "C:\\Program Files\\LibreOffice\\program\\soffice.exe";

exports.generateCertificate = async (req, res) => {
  const safeUnlink = (filePath) => {
    try {
      if (filePath && fs.existsSync(filePath)) fs.unlinkSync(filePath);
    } catch (_) { }
  };

  const tempDir = path.join(__dirname, "../temp");
  if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });

  try {
    const payload =
      req.body && Object.keys(req.body).length
        ? req.body
        : JSON.parse(fs.readFileSync(path.join(__dirname, "../data.json"), "utf8"));

    const templatePath = path.resolve(__dirname, "..", payload.templatePath);
    if (!fs.existsSync(templatePath)) throw new Error("Template not found: " + templatePath);
    const templateDir = path.dirname(templatePath);
    const content = fs.readFileSync(templatePath, "binary");
    const rawData = payload.data || payload.Data || [];
    const students = Array.isArray(rawData) ? rawData : [rawData];

    // fetch remote image and convert to data URI
    async function fetchImageAsDataUri(url) {
      const resp = await axios.get(url, { responseType: "arraybuffer", timeout: 15000 });
      const ct = resp.headers["content-type"] || "image/jpeg";
      const b64 = Buffer.from(resp.data, "binary").toString("base64");
      return `data:${ct};base64,${b64}`;
    }

    // create image module that resolves paths relative to template
    const createImageModule = () =>
      new ImageModule({
        getImage(tagValue) {
          if (!tagValue) {
            console.warn("getImage called with empty tagValue");
            return null;
          }

          // If it's already a data URI
          if (typeof tagValue === "string" && tagValue.startsWith("data:image")) {
            const base64 = tagValue.split(",")[1];
            const buf = Buffer.from(base64, "base64");
            console.log("getImage: got data URI, size=", buf.length);
            try { fs.writeFileSync(path.join(tempDir, "debug_datauri.png"), buf); } catch (e) { }
            return buf;
          }

          // If it's an absolute path
          if (typeof tagValue === "string" && path.isAbsolute(tagValue)) {
            if (fs.existsSync(tagValue)) {
              const buf = fs.readFileSync(tagValue);
              console.log("getImage: read absolute path ->", tagValue, "size=", buf.length);
              try { fs.writeFileSync(path.join(tempDir, "debug_abs.png"), buf); } catch (e) { }
              return buf;
            } else {
              console.warn("getImage: absolute path not found:", tagValue);
            }
          }

          // Try relative to template directory (most common)
          try {
            const candidate1 = path.resolve(templateDir, tagValue);
            console.log("getImage: trying templateDir ->", candidate1, "exists=", fs.existsSync(candidate1));
            if (fs.existsSync(candidate1)) {
              const buf = fs.readFileSync(candidate1);
              try { fs.writeFileSync(path.join(tempDir, "debug_template.png"), buf); } catch (e) { }
              return buf;
            }

            // Try explicit templates folder sibling to project
            const candidate2 = path.resolve(__dirname, "..", "templates", tagValue);
            console.log("getImage: trying __dirname../templates ->", candidate2, "exists=", fs.existsSync(candidate2));
            if (fs.existsSync(candidate2)) {
              const buf = fs.readFileSync(candidate2);
              try { fs.writeFileSync(path.join(tempDir, "debug_templates_folder.png"), buf); } catch (e) { }
              return buf;
            }

            // Try current working directory
            const candidate3 = path.resolve(process.cwd(), tagValue);
            console.log("getImage: trying CWD ->", candidate3, "exists=", fs.existsSync(candidate3));
            if (fs.existsSync(candidate3)) {
              const buf = fs.readFileSync(candidate3);
              try { fs.writeFileSync(path.join(tempDir, "debug_cwd.png"), buf); } catch (e) { }
              return buf;
            }
          } catch (e) {
            console.error("getImage error resolving paths:", e);
          }

          console.warn("getImage: image not found for tagValue:", tagValue);
          return null;
        },

        getSize() {
          // width, height in px (tune as needed)
          return [120, 120];
        },
      });

    // render single student to DOCX buffer (async to allow fetching images)
    async function renderCertificateBuffer(studentData) {
      // Pre-fetch remote URL images and convert to data URI
      if (studentData && studentData.STU_PIC && typeof studentData.STU_PIC === "string") {
        const v = studentData.STU_PIC.trim();
        if (v.startsWith("http://") || v.startsWith("https://")) {
          try {
            studentData.STU_PIC = await fetchImageAsDataUri(v);
          } catch (err) {
            console.warn("⚠️ Failed to fetch image URL:", v, err.message);
            // leave as-is; module will log and render fallback
          }
        }
        // else: leave as filename (relative) or absolute path — the image module will try to read it
      }

      // Debug: print resolved checks for this student's STU_PIC
      try {
        const resolvedFromTemplate = path.resolve(templateDir, studentData.STU_PIC || "");
        const resolvedFromCwd = path.resolve(process.cwd(), studentData.STU_PIC || "");
        const altTemplatesPath = path.resolve(__dirname, "..", "templates", studentData.STU_PIC || "");
        console.log("DEBUG image checks for student:", studentData.student_name || "<unknown>");
        console.log(" templateDir =", templateDir);
        console.log(" resolve(templateDir, STU_PIC) =", resolvedFromTemplate, "exists =", fs.existsSync(resolvedFromTemplate));
        console.log(" resolve(CWD, STU_PIC)      =", resolvedFromCwd, "exists =", fs.existsSync(resolvedFromCwd));
        console.log(" __dirname/../templates      =", altTemplatesPath, "exists =", fs.existsSync(altTemplatesPath));
      } catch (e) {
        console.error("Debug path check error:", e);
      }

      const zip = new PizZip(content);
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        modules: [createImageModule()],
      });

      doc.render(studentData);
      return doc.getZip().generate({ type: "nodebuffer" });
    }

    const generatedPdfPaths = [];

    for (let i = 0; i < students.length; i++) {
      const s = students[i] || {};
      const idPart = (s.id || s.student_name || `idx${i}`).toString().replace(/\s+/g, "_");
      const uniqueBase = `certificate_${idPart}_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
      const tmpDocx = path.join(tempDir, `${uniqueBase}.docx`);
      const tmpPdfGuess = path.join(tempDir, `${uniqueBase}.pdf`);

      const docxBuffer = await renderCertificateBuffer(s);
      fs.writeFileSync(tmpDocx, docxBuffer);

      await execFileAsync(sofficePath, [
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        tempDir,
        tmpDocx,
      ]);

      const producedPdf = fs
        .readdirSync(tempDir)
        .map((f) => path.join(tempDir, f))
        .find(
          (f) =>
            f.endsWith(".pdf") &&
            (f.includes(uniqueBase) || f.includes(path.basename(tmpDocx, ".docx")))
        );

      const finalPdfPath = producedPdf || tmpPdfGuess;
      if (!fs.existsSync(finalPdfPath)) {
        const listing = fs.readdirSync(tempDir).join(", ");
        throw new Error(
          `PDF not found for ${tmpDocx}. Checked: ${finalPdfPath}. Temp folder contains: ${listing}`
        );
      }

      generatedPdfPaths.push(finalPdfPath);
      safeUnlink(tmpDocx);
    }

    let outputPdfPath;
    if (generatedPdfPaths.length === 1) {
      outputPdfPath = generatedPdfPaths[0];
    } else {
      const mergedPdf = await PDFDocument.create();
      for (const p of generatedPdfPaths) {
        const bytes = fs.readFileSync(p);
        const donor = await PDFDocument.load(bytes);
        const copied = await mergedPdf.copyPages(donor, donor.getPageIndices());
        copied.forEach((pg) => mergedPdf.addPage(pg));
      }
      const mergedBytes = await mergedPdf.save();
      outputPdfPath = path.join(tempDir, `merged_certificates_${Date.now()}.pdf`);
      fs.writeFileSync(outputPdfPath, Buffer.from(mergedBytes));
      generatedPdfPaths.forEach((p) => safeUnlink(p));
    }

    res.download(outputPdfPath, "student_certificates.pdf", (err) => {
      safeUnlink(outputPdfPath);
      if (err) console.error("❌ Error sending file:", err);
      else console.log("📤 PDF sent successfully!");
    });
  } catch (err) {
    console.error("❌ Error generating certificate:", err);
    res.status(500).json({ error: err.message || "Server error" });
  }
};
