'use strict';

require('dotenv').config();

const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { extractSlidesText, extractTemplateReferences } = require('./parser');
const { analyzePresentation, regenerateSlide, regeneratePresentation } = require('./llmService');
const { buildPptxBuffer } = require('./pptxBuilder');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 75 * 1024 * 1024 }
});

const templatePath = findTemplatePath();
const templateReferences = templatePath ? extractTemplateReferences(templatePath) : [];

app.use(express.json({ limit: '25mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/process', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Veuillez importer un fichier .pptx.' });
  if (!req.file.originalname.toLowerCase().endsWith('.pptx')) {
    return res.status(400).json({ error: 'Seuls les fichiers .pptx sont acceptés.' });
  }

  try {
    const slides = await extractSlidesText(req.file.buffer);
    if (!slides.length) return res.status(400).json({ error: 'Aucune slide lisible n’a été trouvée dans ce PPTX.' });
    const state = await analyzePresentation({
      slides,
      description: req.body.description || '',
      mode: 'equilibre',
      franceDate: getFranceDateContext(),
      templateReferences
    });
    res.json(state);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || 'Impossible de traiter la présentation.' });
  }
});

app.post('/api/regenerate-slide', async (req, res) => {
  const { slide, layout, mode } = req.body || {};
  if (!slide || !layout) return res.status(400).json({ error: 'La slide et le modèle sont requis.' });

  try {
    const updated = await regenerateSlide({ slide, layout, mode: mode || 'equilibre', franceDate: getFranceDateContext() });
    res.json(updated);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || 'Impossible de régénérer la slide.' });
  }
});

app.post('/api/regenerate-presentation', async (req, res) => {
  const { state, mode } = req.body || {};
  if (!state) return res.status(400).json({ error: 'La présentation est requise.' });

  try {
    const updated = await regeneratePresentation({ state, mode: mode || 'equilibre', franceDate: getFranceDateContext() });
    res.json(updated);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || 'Impossible de régénérer la présentation.' });
  }
});

app.post('/api/download', async (req, res) => {
  try {
    const buffer = await buildPptxBuffer(req.body);
    const filename = `Ascence_Rebranded_${Date.now()}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message || 'Impossible de construire le PPTX.' });
  }
});

app.get('/api/health', (req, res) => {
  res.json({
    ok: true,
    template: templatePath ? path.basename(templatePath) : null,
    templateSlides: templateReferences.length
  });
});

function findTemplatePath() {
  const candidates = ['Template So Far.pptx', 'Template_So_Far.pptx', 'charte Ascence Avisory(2).pptx'];
  for (const name of candidates) {
    const fullPath = path.join(__dirname, name);
    if (fs.existsSync(fullPath)) return fullPath;
  }
  return null;
}

function getFranceDateContext() {
  const now = new Date();
  return new Intl.DateTimeFormat('fr-FR', {
    timeZone: 'Europe/Paris',
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    timeZoneName: 'short'
  }).format(now);
}

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`PPTX Rebrander running at http://localhost:${port}`);
  console.log(templatePath ? `Template loaded: ${path.basename(templatePath)}` : 'Template file not found, using coded Ascence style.');
});
