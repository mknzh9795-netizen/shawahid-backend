// server.js

const express = require("express");
const cors = require("cors");
const morgan = require("morgan");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { google } = require("googleapis");
const QRCode = require("qrcode");
const ImageModule = require("docxtemplater-image-module-free");

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

// نخزن التوكنات في الذاكرة (للتجربة فقط)
let cachedTokens = null;

// أسماء المجلدات الفرعية لشواهد الأداء الوظيفي
const performanceFolders = [
  "أداء الواجبات الوظيفية",
  "التفاعل مع المجتمع المهني",
  "التفاعل مع أولياء الأمور",
  "التنويع في استراتيجيات التدريس",
  "تحسين نتائج المتعلمين",
  "إعداد وتنفيذ خطة التعلم",
  "توظيف تقنيات ووسائل التعلم المناسبة",
  "تهيئة البيئة التعليمية",
  "الإدارة الصفية",
  "تحليل نتائج المتعلمين وتشخيص مستوياتهم",
  "تنوع أساليب التقويم"
];

// ========= دوال Google Drive =========

function getDriveForCurrentUser() {
  if (!cachedTokens) {
    throw new Error("NO_TOKENS");
  }
  oauth2Client.setCredentials(cachedTokens);
  return google.drive({ version: "v3", auth: oauth2Client });
}

async function createFolder(drive, name, parentId) {
  const fileMetadata = {
    name,
    mimeType: "application/vnd.google-apps.folder"
  };

  if (parentId) {
    fileMetadata.parents = [parentId];
  }

  const res = await drive.files.create({
    requestBody: fileMetadata,
    fields: "id"
  });

  const fileId = res.data.id;

  // جعل المجلد متاحاً لأي شخص معه الرابط (قراءة فقط)
  await drive.permissions.create({
    fileId,
    requestBody: {
      role: "reader",
      type: "anyone"
    }
  });

  const info = await drive.files.get({
    fileId,
    fields: "webViewLink"
  });

  return { id: fileId, link: info.data.webViewLink };
}

async function createTeacherFoldersForUser(teacherName) {
  const drive = getDriveForCurrentUser();

  // المجلد الرئيسي بصيغة:
  // شواهد الأداء الوظيفي أ. (خالد)
  const main = await createFolder(
    drive,
`شواهد الأداء الوظيفي أ. (${teacherName})`
    null
  );

  const folderLinks = [];

  // إنشاء 11 مجلد فرعي بأسمائها الرسمية داخل المجلد الرئيسي
  for (let i = 0; i < performanceFolders.length; i++) {
    const folderName = performanceFolders[i];
    const { link } = await createFolder(drive, folderName, main.id);
    folderLinks.push(link);
  }

  return {
    mainFolderId: main.id,
    links: folderLinks // روابط المجلدات الفرعية فقط
  };
}

// ========= دوال QR + Image Module =========

let qrImages = {}; // نخزن فيها صور الـ QR كـ Buffer

async function generateQrBuffer(url) {
  if (!url) return null;
  const buffer = await QRCode.toBuffer(url, {
    type: "png",
    width: 600,
    margin: 1
  });
  return buffer;
}

function createImageModule() {
  return new ImageModule({
    getImage(tagValue) {
      // tagValue مثل: "qr1", "qr2", ...
      return qrImages[tagValue];
    },
    getSize() {
      // العرض × الارتفاع (سم تقريبياً)
      return [4, 4];
    }
  });
}

// ========= دوال التوليد من القالب =========

// نسخة قديمة (بدون QR) – لو حاب تستخدمها لمسارات أخرى
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
      end: "}}"
    }
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
    compression: "DEFLATE"
  });

  return buf;
}

// نسخة جديدة تدعم حقن صور QR داخل القالب
function generateFromTemplateWithQr(data) {
  const templatePath = path.join(__dirname, "templates", "template.pptx");
  console.log("📁 Using template (QR):", templatePath);

  let content;
  try {
    content = fs.readFileSync(templatePath, "binary");
  } catch (err) {
    console.error("❌ Template not found:", err.message);
    throw new Error("TEMPLATE_NOT_FOUND");
  }

  const zip = new PizZip(content);
  const imageModule = createImageModule();

  const doc = new Docxtemplater(zip, {
    delimiters: { start: "{{", end: "}}" },
    modules: [imageModule]
  });

  doc.setData(data);

  try {
    doc.render();
  } catch (error) {
    console.error("❌ Template render error (QR):");
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
    compression: "DEFLATE"
  });

  return buf;
}

// ========= مسارات Google OAuth =========

// بدء المصادقة مع جوجل
app.get("/auth/google", (req, res) => {
  console.log("GOOGLE_CLIENT_ID =", CLIENT_ID);
  console.log("GOOGLE_REDIRECT_URI =", REDIRECT_URI);

  const scopes = ["https://www.googleapis.com/auth/drive.file"];

  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: scopes,
    redirect_uri: REDIRECT_URI
  });

  console.log("Generated auth URL:", url);
  res.redirect(url);
});

// استقبال الكود من جوجل
app.get("/oauth2callback", async (req, res) => {
  const code = req.query.code;
  console.log("Callback code =", code);

  if (!code) {
    return res.status(400).send("missing code");
  }

  try {
    const { tokens } = await oauth2Client.getToken({
      code,
      redirect_uri: REDIRECT_URI
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

// فحص حالة التوكنات
app.get("/debug/google-tokens", (req, res) => {
  if (!cachedTokens) {
    return res.json({ connected: false, message: "لا يوجد مستخدم مربوط بعد" });
  }

  res.json({
    connected: true,
    tokens: {
      access_token: cachedTokens.access_token,
      refresh_token: cachedTokens.refresh_token,
      expiry_date: cachedTokens.expiry_date
    }
  });
});

// ========= Routes عادية =========

// فحص سريع أن السيرفر شغال
app.get("/", (req, res) => {
  res.send("✅ shawahid-backend is running");
});

// توليد بوربوينت عادي من القالب (بدون QR) – اختياري
app.post("/generate-ppt", (req, res) => {
  const data = req.body || {};
  console.log("📦 BODY (generate-ppt):", data);

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
        message: "Template file not found on server"
      });
    }

    const payload = {
      message: "Template error"
    };

    if (error.properties && error.properties.errors) {
      payload.details = error.properties.errors.map((e) => ({
        id: e.properties.id,
        file: e.properties.file,
        context: e.properties.context,
        explanation: e.properties.explanation
      }));
    } else {
      payload.details = { message: error.message };
    }

    console.error("❌ Error in /generate-ppt:", payload);
    res.status(500).json(payload);
  }
});

// المسار الرئيسي: إنشاء مجلد رئيسي + 11 مجلد فرعي + باركود فقط لكل واحد
app.post("/generate-folders-and-ppt", async (req, res) => {
  if (!cachedTokens) {
    return res.status(401).json({
      message: "اربط حساب Google عبر /auth/google أولاً"
    });
  }

  const body = req.body || {};
  const teacherName = body.teacher_name || "معلم";

  try {
    // 1) إنشاء المجلدات في درايف المستخدم
    const { links } = await createTeacherFoldersForUser(teacherName);
    console.log("📂 Created folders:", links);

    // 2) توليد QR فقط لكل مجلد فرعي
    qrImages = {};
    for (let i = 0; i < links.length; i++) {
      const key = `qr${i + 1}`; // qr1..qr11
      qrImages[key] = await generateQrBuffer(links[i]);
    }

    // 3) المتغيرات الممرّرة للقالب: فقط اسم المعلم + مفاتيح QR
    const data = {
      teacher_name: teacherName,
      qr1: "qr1",
      qr2: "qr2",
      qr3: "qr3",
      qr4: "qr4",
      qr5: "qr5",
      qr6: "qr6",
      qr7: "qr7",
      qr8: "qr8",
      qr9: "qr9",
      qr10: "qr10",
      qr11: "qr11"
    };

    // 4) توليد البوربوينت مع الباركود
    const buffer = generateFromTemplateWithQr(data);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader(
      "Content-Disposition",
      'attachment; filename="shawahid-folders-qr.pptx"'
    );

    res.send(buffer);
  } catch (error) {
    console.error("❌ Error in /generate-folders-and-ppt:", error);
    res.status(500).json({
      message: "Drive or template error",
      details: error.message
    });
  }
});

// فحص القالب
app.get("/debug-template", (req, res) => {
  const templatePath = path.join(__dirname, "templates", "template.pptx");
  const exists = fs.existsSync(templatePath);

  res.json({
    templatePath,
    exists
  });
});

// ========= Start server =========
app.listen(PORT, () => {
  console.log(`🚀 Server listening on port ${PORT}`);
});
