const express = require("express");
const cors = require("cors");
const morgan = require("morgan");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const app = express();
const PORT = process.env.PORT || 3000;

// Middlewares
app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(morgan("dev"));

// دالة توليد البوربوينت من القالب
function generateFromTemplate(data) {
  const templatePath = path.join(__dirname, "templates", "template.pptx");
  console.log("📁 Using template:", templatePath);

  let content;
  try {
    content = fs.readFileSync(templatePath, "binary");
  } catch (err) {
    console.error("❌ Template not found:", err.message);
    throw new Error("TEMPLATE_NOT_FOUND");
  }

  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    fileType: "pptx",
    delimiters: {
      start: "{{",
      end: "}}",
    },
  });

  doc.setData(data);

  try {
    doc.render();
  } catch (error) {
    console.error("❌ Template render error:");
    if (error.properties && error.properties.errors) {
      error.properties.errors.forEach((e) => {
        console.error(JSON.stringify(e.properties, null, 2));
      });
    } else {
      console.error(error);
    }
    throw error;
  }

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  return buf;
}

// فحص سريع إن السيرفر شغال
app.get("/", (req, res) => {
  res.send("✅ shawahid-backend is running");
});

// مسار توليد البوربوينت
app.post("/generate-ppt", (req, res) => {
  const data = req.body || {};
  console.log("📦 BODY:", data);

  try {
    const buffer = generateFromTemplate(data);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="shawahid-template.pptx"'
    );

    res.send(buffer);
  } catch (error) {
    if (error.message === "TEMPLATE_NOT_FOUND") {
      return res.status(500).json({
        message: "Template file not found on server",
      });
    }

    const payload = {
      message: "Template error",
    };

    if (error.properties && error.properties.errors) {
      payload.details = error.properties.errors.map((e) => ({
        id: e.properties.id,
        file: e.properties.file,
        context: e.properties.context,
        explanation: e.properties.explanation,
      }));
    } else {
      payload.details = { message: error.message };
    }

    console.error("❌ Error in /generate-ppt:", payload);
    res.status(500).json(payload);
  }
});

app.get("/debug-template", (req, res) => {
  const templatePath = path.join(__dirname, "templates", "template.pptx");
  const exists = fs.existsSync(templatePath);

  res.json({
    templatePath,
    exists,
  });
});


app.listen(PORT, () => {
  console.log(`🚀 Server listening on port ${PORT}`);
});
