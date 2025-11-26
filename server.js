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

  // أهم شيء: مافي fileType ولا setOptions
  const doc = new Docxtemplater(zip, {
    delimiters: {
      start: "{{",
      end: "}}",
    },
    // لو حاب تضيف خيارات ثانية زي:
    // paragraphLoop: true,
    // linebreaks: true,
  });

  // تقدر تبقيها كذا
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
const body = req.body || {};

const data = {
  id: body.id || "",
  teacher_name: body.teacher_name || "",
  birth: body.birth || "",
  adress: body.adress || "",
  phone: body.phone || "",
  email: body.email || "",
  date: body.date || "",
  dgree: body.dgree || "",
  branch: body.branch || "",
  local: body.local || "",
  moahel: body.moahel || "",
  tahsel: body.tahsel || "",
  one: body.one || "",
  start: body.start || "",
  step: body.step || "",
  teacher: body.teacher || ""
};

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

// ====== Google OAuth2 setup ======
const { google } = require("googleapis");

const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI;

const oauth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);

// تخزين مؤقت للتوكنات أثناء التجربة
let cachedTokens = null;

// بدء تسجيل الدخول عن طريق جوجل
app.get("/auth/google", (req, res) => {
  const scopes = [
    "https://www.googleapis.com/auth/drive.file",
  ];

  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: scopes,
  });

  res.redirect(url);
});

// استقبال Google Callback
app.get("/oauth2callback", async (req, res) => {
  const code = req.query.code;

  if (!code) {
    return res.status(400).send("missing code");
  }

  try {
    const { tokens } = await oauth2Client.getToken(code);
    cachedTokens = tokens;

    console.log("🔥 Google Tokens:", tokens);

    res.send(`
      <h2 style="font-family: sans-serif; direction: rtl; text-align: center; margin-top:40px;">
        🎉 تم ربط حساب Google بنجاح  
      </h2>
      <p style="text-align:center; direction:rtl">تقدر تقفل الصفحة الآن وترجع لنظام الشواهد.</p>
    `);

  } catch (err) {
    console.error("❌ Error exchanging code:", err);
    res.status(500).send("Authentication error with Google");
  }
});

// لمعرفة هل فيه مستخدم مربوط الآن
app.get("/debug/google-tokens", (req, res) => {
  res.json(cachedTokens || { message: "No tokens yet" });
});
