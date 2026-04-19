// PUT YOUR GROQ KEY HERE
process.env.GROQ_API_KEY = "API_KEY_HERE";

const express = require("express");
const fs      = require("fs");
const path    = require("path");
const JSZip   = require("jszip");

const app = express();
app.use(express.json({ limit: "4mb" }));
app.use(express.static("public"));

const TEMPLATE = path.join(__dirname, "Model1Fin.pptx");
const OUT_DIR  = path.join(__dirname, "output");
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR);

// GOOGLE DRIVE (optional)
let driveClient  = null;
let driveAuthUrl = null;
try {
  const { google } = require("googleapis");
  const creds = JSON.parse(fs.readFileSync(path.join(__dirname, "credentials.json")));
  const { client_id, client_secret, redirect_uris } = creds.installed || creds.web;
  const oauth2 = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  const tokenPath = path.join(__dirname, "token.json");
  if (fs.existsSync(tokenPath)) {
    oauth2.setCredentials(JSON.parse(fs.readFileSync(tokenPath)));
    driveClient = google.drive({ version: "v3", auth: oauth2 });
    console.log("Google Drive connected");
  } else {
    driveAuthUrl = oauth2.generateAuthUrl({ access_type: "offline", scope: ["https://www.googleapis.com/auth/drive.file"] });
    console.log("Google Drive: visit", driveAuthUrl);
    app._oauth2 = oauth2;
  }
} catch { console.log("Google Drive not configured (optional)"); }

// GROQ LLM
async function callLLM(userText) {
  const key = process.env.GROQ_API_KEY;
  if (!key || key === "gsk_YOUR_KEY_HERE") throw new Error("Set your GROQ_API_KEY at the top of server.js");

  const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${key}` },
    body: JSON.stringify({
      model: "llama-3.1-8b-instant",
      temperature: 0.3,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content: `You generate JSON for a 2-axis slide. Return ONLY valid JSON, no markdown, no explanation, and do not state AXE1 and AXE2 explicitly in those titles.

Schema:
{
  "title": "Slide title, max 8 words",
  "shortParagraph": "One compelling intro sentence, max 20 words",
  "axes": [
    {
      "header": "LABEL1",
      "items": [
        { "type": "paragraph", "text": "Optional intro sentence (max 15 words)" },
        { "type": "bullet", "text": "Concise bullet point (max 10 words)" },
        { "type": "bullet", "text": "Concise bullet point (max 10 words)" },
        { "type": "bullet", "text": "Concise bullet point (max 10 words)" }
      ],
      "keywords": ["word1", "word2", "word3"]
    },
    {
      "header": "LABEL2p",
      "items": [
        { "type": "paragraph", "text": "Optional intro sentence (max 15 words)" },
        { "type": "bullet", "text": "Concise bullet point (max 10 words)" },
        { "type": "bullet", "text": "Concise bullet point (max 10 words)" },
        { "type": "bullet", "text": "Concise bullet point (max 10 words)" }
      ],
      "keywords": ["word1", "word2", "word3"]
    }
  ]
}

Rules:
- "paragraph" = plain intro text (gray, no bullet). 0 or 1 per axis.
- "bullet" = bullet point (dark, bullet char). 2 to 4 per axis.
- keywords must appear in the items text of that axis and will be bolded.
- header LABEL must be SHORT (1 to 2 words max), ALL CAPS. Example: "AXE 1 : SANTE", "AXE 2 : DONNEES".
- Use a colon with spaces " : " as separator in headers, NOT an em dash or any dash.
- Keep all text concise, this is a slide.`,
        },
        { role: "user", content: `Summarize this into the JSON schema:\n\n${userText}` },
      ],
    }),
    signal: AbortSignal.timeout(30000),
  });

  const data = await res.json();
  if (!res.ok) throw new Error(JSON.stringify(data.error || data));
  const raw = data.choices[0].message.content.replace(/^```(?:json)?\s*/i,"").replace(/\s*```$/i,"").trim();
  return JSON.parse(raw);
}

// ROUTES

app.post("/api/process", async (req, res) => {
  const { text } = req.body;
  if (!text?.trim()) return res.status(400).json({ error: "No text provided" });
  try { res.json({ ok: true, content: await callLLM(text.trim()) }); }
  catch (err) { res.status(500).json({ error: err.message }); }
});

app.post("/api/generate", async (req, res) => {
  const { content } = req.body;
  if (!content?.axes?.length) return res.status(400).json({ error: "Need content with axes" });
  try { res.json({ ok: true, file: await buildPptx(content) }); }
  catch (err) { res.status(500).json({ error: err.message }); }
});

// One-shot: text to PPTX
app.post("/api/process-and-generate", async (req, res) => {
  const { text } = req.body;
  if (!text?.trim()) return res.status(400).json({ error: "No text provided" });
  try {
    const content = await callLLM(text.trim());
    const file    = await buildPptx(content);
    res.json({ ok: true, file, content });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/api/download/:file", (req, res) => {
  const f = path.join(OUT_DIR, path.basename(req.params.file));
  if (!fs.existsSync(f)) return res.status(404).send("Not found");
  res.download(f);
});

app.get("/api/drive-status", (req, res) => {
  res.json({ connected: !!driveClient, authUrl: driveAuthUrl || null });
});

app.post("/api/upload-drive", async (req, res) => {
  if (!driveClient) return res.status(400).json({ error: "Drive not connected", authUrl: driveAuthUrl });
  const { file } = req.body;
  const f = path.join(OUT_DIR, path.basename(file));
  if (!fs.existsSync(f)) return res.status(404).json({ error: "File not found" });
  try {
    const { data } = await driveClient.files.create({
      requestBody: { name: path.basename(file, ".pptx"), mimeType: "application/vnd.google-apps.presentation" },
      media: { mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", body: fs.createReadStream(f) },
      fields: "id,webViewLink",
    });
    res.json({ ok: true, link: data.webViewLink });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get("/oauth2callback", async (req, res) => {
  const { code } = req.query;
  if (!code || !app._oauth2) return res.send("Error: OAuth not initialized");
  try {
    const { tokens } = await app._oauth2.getToken(code);
    app._oauth2.setCredentials(tokens);
    fs.writeFileSync(path.join(__dirname, "token.json"), JSON.stringify(tokens));
    driveClient  = require("googleapis").google.drive({ version: "v3", auth: app._oauth2 });
    driveAuthUrl = null;
    res.send(`<h2 style="font-family:sans-serif;color:green;padding:40px">Google Drive connected! Close this tab.</h2>`);
  } catch (err) { res.send("Error: " + err.message); }
});

// PPTX BUILDER

async function buildPptx(content) {
  const zip   = await JSZip.loadAsync(fs.readFileSync(TEMPLATE));
  const SLIDE = "ppt/slides/slide1.xml";
  let xml     = await zip.file(SLIDE).async("string");

  xml = injectSlide(xml, content);

  zip.file(SLIDE, xml);
  const buf     = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  const outFile = path.join(OUT_DIR, `slide_${Date.now()}.pptx`);
  fs.writeFileSync(outFile, buf);
  return path.basename(outFile);
}

// XML HELPERS

function esc(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
}

function makeRun(text, { sz="1300", color="3D1A2B", face="Inter", bold=false } = {}) {
  return [
    `<a:r>`,
    `<a:rPr${bold ? ' b="1"' : ' b="0"'} i="0" lang="fr" sz="${sz}" u="none" cap="none" strike="noStrike">`,
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`,
    `<a:latin typeface="${face}"/><a:ea typeface="${face}"/><a:cs typeface="${face}"/><a:sym typeface="${face}"/>`,
    `</a:rPr>`,
    `<a:t>${esc(text)}</a:t>`,
    `</a:r>`,
  ].join("");
}

// Split text by keywords and bold matching runs
function makeRuns(text, keywords = [], opts = {}) {
  if (!keywords.length) return makeRun(text, opts);
  const re = new RegExp(
    `(${keywords.map(k => k.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")).join("|")})`, "gi"
  );
  return text.split(re).filter(Boolean).map(part => {
    const isKw = keywords.some(k => k.toLowerCase() === part.toLowerCase());
    return makeRun(part, { ...opts, bold: isKw });
  }).join("");
}

// Plain intro paragraph: gray, no bullet, left aligned
function makePlainPara(text, keywords) {
  const runs = makeRuns(text, keywords, { sz:"1300", color:"5C4B53", face:"Inter" });
  return [
    `<a:p>`,
    `<a:pPr indent="0" lvl="0" marL="0" marR="0" rtl="0" algn="l">`,
    `<a:lnSpc><a:spcPct val="130000"/></a:lnSpc>`,
    `<a:spcBef><a:spcPts val="0"/></a:spcBef>`,
    `<a:spcAft><a:spcPts val="200"/></a:spcAft>`,
    `<a:buNone/>`,
    `</a:pPr>`,
    runs,
    `<a:endParaRPr b="0" i="0" sz="1300" u="none" cap="none" strike="noStrike">`,
    `<a:solidFill><a:srgbClr val="5C4B53"/></a:solidFill>`,
    `<a:latin typeface="Inter"/><a:ea typeface="Inter"/><a:cs typeface="Inter"/>`,
    `</a:endParaRPr></a:p>`,
  ].join("");
}

// Bullet paragraph: dark, native bullet, left aligned
function makeBulletPara(text, keywords) {
  const runs = makeRuns(text, keywords, { sz:"1300", color:"3D1A2B", face:"Inter" });
  return [
    `<a:p>`,
    `<a:pPr indent="-215900" lvl="0" marL="215900" marR="0" rtl="0" algn="l">`,
    `<a:lnSpc><a:spcPct val="120000"/></a:lnSpc>`,
    `<a:spcBef><a:spcPts val="160"/></a:spcBef>`,
    `<a:spcAft><a:spcPts val="0"/></a:spcAft>`,
    `<a:buClr><a:srgbClr val="3D1A2B"/></a:buClr>`,
    `<a:buSzPts val="1300"/>`,
    `<a:buFont typeface="Arial"/>`,
    `<a:buChar char="•"/>`,
    `</a:pPr>`,
    runs,
    `<a:endParaRPr b="0" i="0" sz="1300" u="none" cap="none" strike="noStrike">`,
    `<a:solidFill><a:srgbClr val="3D1A2B"/></a:solidFill>`,
    `<a:latin typeface="Inter"/><a:ea typeface="Inter"/><a:cs typeface="Inter"/>`,
    `</a:endParaRPr></a:p>`,
  ].join("");
}

// Build all paragraphs for one axis content zone
function buildAxisContent(items, keywords) {
  const kws = (keywords || []).filter(Boolean);
  return (items || []).map(item =>
    item.type === "paragraph"
      ? makePlainPara(item.text, kws)
      : makeBulletPara(item.text, kws)
  ).join("");
}

// CORE INJECTION
// Replaces the entire txBody of shape `id` with paragraphsXml.
function injectTxBody(xml, id, paragraphsXml) {
  const pos = xml.indexOf(`id="${id}"`);
  if (pos === -1) { console.warn(`Shape ${id} not found`); return xml; }

  const sp0   = xml.lastIndexOf("<p:sp>", pos);
  const sp1   = xml.indexOf("</p:sp>", pos) + 7;
  const block = xml.slice(sp0, sp1);

  const tb0    = block.indexOf("<p:txBody>");
  const tb1    = block.indexOf("</p:txBody>") + 11;
  const tbInner = block.slice(tb0 + 10, tb1 - 11);

  // Extract the original bodyPr
  const bodyPrEnd = tbInner.includes("</a:bodyPr>")
    ? tbInner.indexOf("</a:bodyPr>") + 11
    : tbInner.indexOf("/>") + 2;
  const bodyPr = tbInner.slice(0, bodyPrEnd);

  const newTb    = `<p:txBody>${bodyPr}<a:lstStyle/>${paragraphsXml}</p:txBody>`;
  const newBlock = block.slice(0, tb0) + newTb + block.slice(tb1);
  return xml.slice(0, sp0) + newBlock + xml.slice(sp1);
}

// Move and resize a shape
function moveShape(xml, id, x, y, cx, cy) {
  const pos = xml.indexOf(`id="${id}"`);
  if (pos === -1) return xml;
  const sp0 = xml.lastIndexOf("<p:sp>", pos);
  const sp1 = xml.indexOf("</p:sp>", pos) + 7;
  let block = xml.slice(sp0, sp1);
  block = block
    .replace(/<a:off x="[^"]*" y="[^"]*"\//, `<a:off x="${x}" y="${y}"/`)
    .replace(/<a:ext cx="[^"]*" cy="[^"]*"\//, `<a:ext cx="${cx}" cy="${cy}"/`);
  return xml.slice(0, sp0) + block + xml.slice(sp1);
}

// MASTER INJECTION
/*
  Template shape map (Model1Fin.pptx):

  ID 133  Title                    full width, top
  ID 136  Short paragraph          full width, below title
  ID 134  AXE 1 header             left column
  ID 135  AXE 1 content zone       left column, below header
  ID 139  AXE 2 header             right column
  ID 138  AXE 2 content zone       right column, below header

  No shapes 140/141/142 in this template (already removed upstream).
  Positions preserved from the original file; content is fully replaced.
*/
function injectSlide(xml, content) {
  const [ax1, ax2] = content.axes;

  // 1. Title
  xml = injectTxBody(xml, "133",
    `<a:p><a:pPr algn="l"><a:buNone/></a:pPr>` +
    makeRun(content.title || "TITRE", { sz:"2700", color:"3D1A2B", face:"Urbanist", bold:true }) +
    `<a:endParaRPr b="0" i="0" lang="fr" sz="2700" u="none" cap="none" strike="noStrike">` +
    `<a:solidFill><a:srgbClr val="3D1A2B"/></a:solidFill>` +
    `<a:latin typeface="Urbanist"/><a:ea typeface="Urbanist"/><a:cs typeface="Urbanist"/>` +
    `</a:endParaRPr></a:p>`
  );

  // 2. Short paragraph
  xml = injectTxBody(xml, "136",
    `<a:p><a:pPr algn="l"><a:buNone/></a:pPr>` +
    makeRuns(content.shortParagraph || "", [], { sz:"1500", color:"5C4B53", face:"Inter" }) +
    `<a:endParaRPr b="0" i="0" lang="fr" sz="1500" u="none" cap="none" strike="noStrike">` +
    `<a:solidFill><a:srgbClr val="5C4B53"/></a:solidFill>` +
    `<a:latin typeface="Inter"/><a:ea typeface="Inter"/><a:cs typeface="Inter"/>` +
    `</a:endParaRPr></a:p>`
  );

  // 3. AXE 1 header
  xml = injectTxBody(xml, "134",
    `<a:p><a:pPr algn="l"><a:buNone/></a:pPr>` +
    makeRun(ax1.header || "AXE 1", { sz:"2100", color:"3D1A2B", face:"Urbanist", bold:true }) +
    `<a:endParaRPr b="0" i="0" lang="fr" sz="2100" u="none" cap="none" strike="noStrike">` +
    `<a:solidFill><a:srgbClr val="3D1A2B"/></a:solidFill>` +
    `<a:latin typeface="Urbanist"/><a:ea typeface="Urbanist"/><a:cs typeface="Urbanist"/>` +
    `</a:endParaRPr></a:p>`
  );

  // 4. AXE 1 content
  xml = injectTxBody(xml, "135", buildAxisContent(ax1.items, ax1.keywords));

  // 5. AXE 2 header
  xml = injectTxBody(xml, "139",
    `<a:p><a:pPr algn="l"><a:buNone/></a:pPr>` +
    makeRun(ax2.header || "AXE 2", { sz:"2100", color:"3D1A2B", face:"Urbanist", bold:true }) +
    `<a:endParaRPr b="0" i="0" lang="fr" sz="2100" u="none" cap="none" strike="noStrike">` +
    `<a:solidFill><a:srgbClr val="3D1A2B"/></a:solidFill>` +
    `<a:latin typeface="Urbanist"/><a:ea typeface="Urbanist"/><a:cs typeface="Urbanist"/>` +
    `</a:endParaRPr></a:p>`
  );

  // 6. AXE 2 content
  xml = injectTxBody(xml, "138", buildAxisContent(ax2.items, ax2.keywords));

  return xml;
}

// START
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`http://localhost:${PORT}`));