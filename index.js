const express = require("express");
const path = require("path");
const app = express();
const certificateRoutes = require("./routes/certificateRoutes");
const testGenCertificate = require("./routes/testCertificate");

app.use(express.json({ limit: "50mb" }));
app.use("/api", certificateRoutes);
app.use("/api", testGenCertificate);

app.use("/templates", express.static(path.join(__dirname, "templates")));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
