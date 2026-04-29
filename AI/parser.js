'use strict';

const AdmZip = require('adm-zip');
const { parseStringPromise } = require('xml2js');

async function extractSlidesText(input) {
  const zip = new AdmZip(input);
  const slideEntries = zip.getEntries()
    .filter(entry => /^ppt\/slides\/slide\d+\.xml$/i.test(entry.entryName))
    .sort((a, b) => slideNumber(a.entryName) - slideNumber(b.entryName));

  const slides = [];
  for (const entry of slideEntries) {
    const xml = entry.getData().toString('utf8');
    const rawText = normalizeWhitespace(await extractTextFromXml(xml));
    slides.push({ slideIndex: slideNumber(entry.entryName), rawText });
  }
  return slides;
}

function extractTemplateReferences(filePath) {
  const zip = new AdmZip(filePath);
  return zip.getEntries()
    .filter(entry => /^ppt\/slides\/slide\d+\.xml$/i.test(entry.entryName))
    .sort((a, b) => slideNumber(a.entryName) - slideNumber(b.entryName))
    .map(entry => ({
      slideIndex: slideNumber(entry.entryName),
      xml: entry.getData().toString('utf8')
    }));
}

async function extractTextFromXml(xml) {
  try {
    const parsed = await parseStringPromise(xml, { explicitArray: true, preserveChildrenOrder: true });
    const textRuns = [];
    collectText(parsed, textRuns);
    return textRuns.join(' ');
  } catch (_) {
    return xml.replace(/<[^>]*>/g, ' ');
  }
}

function collectText(node, out) {
  if (node == null) return;
  if (typeof node === 'string') {
    out.push(node);
    return;
  }
  if (Array.isArray(node)) {
    node.forEach(item => collectText(item, out));
    return;
  }
  if (typeof node === 'object') {
    Object.entries(node).forEach(([key, value]) => {
      if (key === 'a:t' || key.endsWith(':t')) collectText(value, out);
      else if (key !== '$') collectText(value, out);
    });
  }
}

function slideNumber(name) {
  const match = name.match(/slide(\d+)\.xml/i);
  return match ? Number(match[1]) : 0;
}

function normalizeWhitespace(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

module.exports = { extractSlidesText, extractTemplateReferences };
