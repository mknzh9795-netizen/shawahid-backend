// server.js

const express = require("express");
const cors = require("cors");
const morgan = require("morgan");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { google } = require("googleapis"); // Google APIs

const app = express();
const PORT = process.env.PORT || 3000;

// ========= Middlewares =========
app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(morgan("dev"));

// ========= Google OAuth2 setup =========
const CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI;

const oauth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
);

// للتجربة فقط: نخزن التوكن في الذاكرة
let cachedTokens = null;

// بدء المصادقة مع جوجل
app.get("/auth/google", (req, res) => {
  console.log("GOOGLE_CLIENT_ID =", CLIENT_ID);
  console.log("GOOGLE_REDIRECT_URI =", REDIRECT_URI);

  const scopes = ["https://www.googleapis.com/auth/drive.file"];

  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: scopes,
    redirect_uri: REDIRECT_URI,
  });

  console.log("Generated auth URL:", url);
  res.redirect(url);
});

// استلام الكود من جوجل
app.get("/oauth2callback", async (req, res) => {
  const code = req.query.code;
  console.log("Callback code =", code);

  if (!code) {
    return res.status(400).send("missing code");
  }

  try {
    const { tokens } = await oauth2Client.getToken({
      code,
      redirect_uri: REDIRECT_URI,
    });

    cachedTokens = tokens;

    console.log("✅ Google tokens saved:", tokens);

    res.send(`
      <html lang="ar" dir="rtl">
      <head><meta charset="utf-8"><title>تم الربط</title></head>
      <body style="font-family: system-ui; text-align: center; padding-top:40px;">
        <h2>✅ تم ربط حساب Google بنجاح</h2>
        <p>يمكنك إغلاق هذه الصفحة والعودة لنظام الشواهد.</p>
      </body>
      </html>
    `);
  } catch (err) {
    console.error("❌ Error exchanging code:", err);
    res.status(500).send("Authentication error with Google");
  }
});

// فحص التوكنات
app.get("/debug/google-tokens", (req, res) => {
  if (!cachedTokens) {
    return res.json({ connected: false, message: "لا يوجد مستخدم مربوط بعد" });
  }

  res.json({
    connected: true,
    tokens: {
      access_token: cachedTokens.access_token,
      refresh_token: cachedTokens.refresh_token,
      expiry_date: cachedTokens.expiry_date,
    },
  });
});

// ========= دالة توليد البوربوينت من القالب =========
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

// ========= Routes عادية =========

// فحص سريع أن السيرفر شغال
app.get("/", (req, res) => {
  res.send("✅ shawahid-backend is running");
});

// توليد البوربوينت من القالب الأساسي
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

// فحص القالب
app.get("/debug-template", (req, res) => {
  const templatePath = path.join(__dirname, "templates", "template.pptx");
  const exists = fs.existsSync(templatePath);

  res.json({
    templatePath,
    exists,
  });
});

// ========= Start server =========
app.listen(PORT, () => {
  console.log(`🚀 Server listening on port ${PORT}`);
});
