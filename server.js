import express from "express";
import { google } from "googleapis";

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 3000;

const {
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  PUBLIC_BASE_URL,
  API_KEY,
  DEFAULT_SPREADSHEET_ID,
  DEFAULT_RANGE
} = process.env;

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets.readonly",
  "https://www.googleapis.com/auth/gmail.send"
];

const oauth2Client = new google.auth.OAuth2(
  GOOGLE_CLIENT_ID,
  GOOGLE_CLIENT_SECRET,
  `${PUBLIC_BASE_URL}/auth/callback`
);

let tokens = null;

/* ================= HELPERS ================= */

function requireApiKey(req, res, next) {
  if (req.headers["x-api-key"] !== API_KEY) {
    return res.status(401).json({ error: "Unauthorized" });
  }
  next();
}

function ensureAuth(req, res, next) {
  if (!tokens) {
    return res.status(401).json({
      error: "Google not authenticated. Visit /auth/start first."
    });
  }
  oauth2Client.setCredentials(tokens);
  next();
}

function normalizeHeader(h) {
  return String(h || "")
    .toLowerCase()
    .trim()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function addDerivedFields(row) {
  const fullName = String(row.name || "").trim();
  if (fullName && !row.first_name) {
    row.first_name = fullName.split(/\s+/)[0];
  }
  return row;
}

function mergeTemplate(template, row) {
  return template.replace(/\{\{\s*([a-z0-9_]+)\s*\}\}/gi, (_, key) => {
    return row[key.toLowerCase()] || "";
  });
}

function createRawEmail({ to, subject, body, fromName }) {
  const message = [
    `To: ${to}`,
    `Subject: ${subject}`,
    `From: "${fromName}" <me>`,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "",
    body
  ].join("\r\n");

  return Buffer.from(message)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

/* ================= ROUTES ================= */

app.get("/", (req, res) => {
  res.send("Pension Outreach Sender Running");
});

app.get("/auth/start", (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
    prompt: "consent"
  });
  res.redirect(url);
});

app.get("/auth/callback", async (req, res) => {
  const { code } = req.query;
  const { tokens: newTokens } = await oauth2Client.getToken(code);
  tokens = newTokens;
  res.send("Google connected successfully. You can close this tab.");
});

async function readSheet(spreadsheetId, range) {
  const sheets = google.sheets({ version: "v4", auth: oauth2Client });

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range
  });

  const rows = response.data.values || [];
  if (rows.length < 2) return [];

  const headers = rows[0].map(normalizeHeader);

  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = r[i] || ""));
    return addDerivedFields(obj);
  });
}

app.post("/campaign/preview", requireApiKey, ensureAuth, async (req, res) => {
  const spreadsheetId = req.body.spreadsheetId || DEFAULT_SPREADSHEET_ID;
  const range = req.body.range || DEFAULT_RANGE;

  const rows = await readSheet(spreadsheetId, range);

  const counts = { A: 0, B: 0, OTHER: 0 };

  rows.forEach(r => {
    const letter = String(r.letters || "").toUpperCase();
    if (letter === "A") counts.A++;
    else if (letter === "B") counts.B++;
    else counts.OTHER++;
  });

  res.json({
    total: rows.length,
    counts,
    sample: rows.slice(0, 3)
  });
});

app.post("/campaign/send", requireApiKey, ensureAuth, async (req, res) => {
  const {
    spreadsheetId = DEFAULT_SPREADSHEET_ID,
    range = DEFAULT_RANGE,
    templates,
    dryRun = true,
    throttlePerMinute = 20
  } = req.body;

  const rows = await readSheet(spreadsheetId, range);
  const gmail = google.gmail({ version: "v1", auth: oauth2Client });

  const delay = 60000 / throttlePerMinute;
  const results = [];

  for (const r of rows) {
    const to = r.email;
    const letter = String(r.letters || "").toUpperCase();

    const template = templates[letter];
    if (!template) {
      results.push({ email: to, status: "skipped" });
      continue;
    }

    const subject = mergeTemplate(template.subjectTemplate, r);
    const body = mergeTemplate(template.bodyTemplate, r);

    if (dryRun) {
      results.push({ email: to, status: "dry_run" });
      continue;
    }

    const raw = createRawEmail({
      to,
      subject,
      body,
      fromName: "Matthew Mershimer, Pension Advisor"
    });

    try {
      const response = await gmail.users.messages.send({
        userId: "me",
        requestBody: { raw }
      });

      results.push({
        email: to,
        status: "sent",
        id: response.data.id
      });
    } catch (error) {
      results.push({
        email: to,
        status: "failed",
        error: error.message
      });
    }

    await new Promise(resolve => setTimeout(resolve, delay));
  }

  res.json({
    total: results.length,
    results
  });
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
