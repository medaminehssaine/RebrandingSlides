'use strict';

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
async function buildPptxBuffer(state) {
  const templatePath = findTemplatePath();
  if (!templatePath) throw new Error('Template So Far.pptx est introuvable.');

  const zip = await JSZip.loadAsync(fs.readFileSync(templatePath));
  const template = await readTemplate(zip);
  const plannedSlides = planSlides(state);

  removeGeneratedSlideParts(zip);
  for (let index = 0; index < plannedSlides.length; index += 1) {
    const slideNo = index + 1;
    const planned = plannedSlides[index];
    const renderedXml = renderTemplateSlide(template, planned.templateSlide, planned.data);
    const relsXml = renderSlideRels(template, planned.templateSlide);
    zip.file(`ppt/slides/slide${slideNo}.xml`, renderedXml);
    zip.file(`ppt/slides/_rels/slide${slideNo}.xml.rels`, relsXml);
  }

  zip.file('ppt/presentation.xml', await renderPresentationXml(zip, plannedSlides.length));
  zip.file('ppt/_rels/presentation.xml.rels', await renderPresentationRels(zip, plannedSlides.length));
  zip.file('[Content_Types].xml', await renderContentTypes(zip, plannedSlides.length));
  await updateAppSlideCount(zip, plannedSlides.length);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

function findTemplatePath() {
  const candidates = ['Template So Far.pptx', 'Template_So_Far.pptx'];
  for (const name of candidates) {
    const fullPath = path.join(__dirname, name);
    if (fs.existsSync(fullPath)) return fullPath;
  }
  return null;
}

async function readTemplate(zip) {
  const slideNames = Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => slideNo(a) - slideNo(b));

  const slides = {};
  for (const name of slideNames) {
    const no = slideNo(name);
    const xml = await zip.file(name).async('string');
    const relName = `ppt/slides/_rels/slide${no}.xml.rels`;
    slides[no] = {
      xml,
      rels: zip.file(relName) ? await zip.file(relName).async('string') : emptyRels(),
      shapes: extractTextShapes(xml)
    };
  }
  return { slides };
}

function extractTextShapes(xml) {
  const shapes = [];
  let order = 0;
  for (const match of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
    const shapeXml = match[0];
    if (!shapeXml.includes('<p:txBody>')) {
      order += 1;
      continue;
    }
    const id = (shapeXml.match(/<p:cNvPr\b[^>]*\bid="([^"]+)"/) || [])[1];
    const name = (shapeXml.match(/<p:cNvPr\b[^>]*\bname="([^"]*)"/) || [])[1] || '';
    const texts = [...shapeXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(item => unescapeXml(item[1]));
    const rawText = texts.join('');
    shapes.push({ order, id, name, rawText, xml: shapeXml });
    order += 1;
  }
  return shapes;
}

function planSlides(state) {
  const slides = [
    { templateSlide: 1, data: { type: 'title', state } },
    { templateSlide: 2, data: { type: 'agenda', state } }
  ];

  (state.sections || []).forEach((section, sectionIndex) => {
    slides.push({ templateSlide: 3, data: { type: 'section', section, sectionIndex } });
    (section.slides || []).forEach(slide => {
      slides.push({
        templateSlide: slide.layout === 'A' ? 4 : slide.layout === 'B' ? 5 : 6,
        data: { type: 'content', slide }
      });
    });
  });

  slides.push({ templateSlide: 7, data: { type: 'closing', state } });
  return slides;
}

function renderTemplateSlide(template, templateSlide, data) {
  const slide = template.slides[templateSlide];
  if (!slide) throw new Error(`Slide template ${templateSlide} introuvable.`);
  let xml = slide.xml;
  const shapes = slide.shapes;

  if (data.type === 'title') {
    xml = injectTextPreserve(xml, findShape(shapes, 'Title', 0).id, [fitTitleForPptx(data.state.title, 52, 7)]);
    xml = injectTextPreserve(xml, findShape(shapes, 'Sub-Title', 1).id, [fitTitleForPptx(data.state.subtitle, 48, 7)]);
    xml = injectTextPreserve(xml, findShape(shapes, 'Sub-Sub-Title', 2).id, [fitTitleForPptx(data.state.subsubtitle, 56, 8)]);
    xml = injectTextPreserve(xml, findShape(shapes, '19 février 2026', 3).id, [data.state.date]);
  } else if (data.type === 'agenda') {
    const agendaShape = findByText(shapes, 'Agenda') || shapes[0];
    const listShape = shapes.find(shape => shape.id !== agendaShape.id && shape.rawText.includes('Rubrique')) || shapes[1];
    const items = (data.state.sections || []).slice(0, 4).map(section => fitTitleForPptx(section.name || 'Rubrique', 34, 5));
    while (items.length < 4) items.push('');
    xml = injectTextPreserve(xml, agendaShape.id, ['Agenda']);
    xml = injectTextPreserve(xml, listShape.id, items);
  } else if (data.type === 'section') {
    const nameShape = shapes.find(shape => shape.rawText.includes('Rubrique')) || shapes[0];
    const numberShape = shapes.find(shape => /\d\d\./.test(shape.rawText)) || shapes[shapes.length - 1];
    xml = injectTextPreserve(xml, nameShape.id, [fitTitleForPptx(data.section.name, 34, 5)]);
    xml = injectTextPreserve(xml, numberShape.id, [`${String(data.sectionIndex + 1).padStart(2, '0')}.`]);
  } else if (data.type === 'content') {
    xml = renderContentSlide(xml, shapes, data.slide);
  }

  return stripEmDashes(xml);
}

function renderContentSlide(xml, shapes, slide) {
  const content = slide.content || {};
  if (slide.layout === 'A') {
    const columns = ensureArray(content.columns, 2).map(normalizeAxisColumnForPptx);
    const map = mapLayoutAShapes(shapes);
    xml = injectTextPreserve(xml, map.title.id, [fitTitleForPptx(content.title, 48, 7)]);
    xml = injectTextPreserve(xml, map.leftHeader.id, [columns[0].label || 'AXE']);
    xml = injectAxisContentPreserve(xml, map.leftBody.id, columns[0], '1200');
    xml = injectTextPreserve(xml, map.bridge.id, [fitTextForPptx(content.bridge || bridgeFromColumns(columns), 140, 20)]);
    xml = injectAxisContentPreserve(xml, map.rightBody.id, columns[1], '1200');
    xml = injectTextPreserve(xml, map.rightHeader.id, [columns[1].label || 'AXE']);
    return xml;
  }

  if (slide.layout === 'B') {
    const columns = ensureArray(content.columns, 3).map(normalizeCardColumnForPptx);
    const map = mapLayoutBShapes(shapes);
    xml = injectTextPreserve(xml, map.title.id, [fitTitleForPptx(content.title, 48, 7)]);
    columns.forEach((column, index) => {
      xml = injectTextPreserve(xml, map.headers[index].id, [column.header || `Levier ${index + 1}`]);
      xml = injectTextPreserve(xml, map.bodies[index].id, [column.body || '']);
    });
    return xml;
  }

  const map = mapLayoutCShapes(shapes);
  const paragraph = fitTextForPptx(removeEmoji(content.paragraph || rowsToParagraph(content.rows) || defaultParagraph()), 360, 58);
  const paragraphSize = oneAxisFontSize(paragraph);
  xml = injectTextPreserve(xml, map.title.id, [removeEmoji(fitTitleForPptx(content.title, 48, 7))]);
  if (map.subtitle) xml = injectTextPreserve(xml, map.subtitle.id, [removeEmoji(fitTitleForPptx(content.subtitle || '', 32, 4))]);
  const paragraphShapes = map.paragraphShapes.length ? map.paragraphShapes : [map.subtitle].filter(Boolean);
  if (!paragraphShapes.length) return xml;
  if (paragraphShapes.length > 1) xml = stretchShapeLike(xml, paragraphShapes[0].id, paragraphShapes);
  xml = injectTextPreserve(xml, paragraphShapes[0].id, [paragraph]);
  xml = setShapeTextSize(xml, paragraphShapes[0].id, paragraphSize);
  paragraphShapes.slice(1).forEach(row => {
    xml = injectTextPreserve(xml, row.id, ['']);
  });
  return xml;
}

function mapLayoutAShapes(shapes) {
  return {
    title: shapes[0],
    leftHeader: shapes[1],
    leftBody: shapes[2],
    bridge: shapes[3],
    rightBody: shapes[4],
    rightHeader: shapes[5]
  };
}

function mapLayoutBShapes(shapes) {
  return {
    title: shapes[0],
    bodies: [shapes[2], shapes[4], shapes[6]],
    headers: [shapes[3], shapes[5], shapes[7]]
  };
}

function mapLayoutCShapes(shapes) {
  const title = shapes[0];
  const rest = shapes.slice(1);
  const sortedByY = rest
    .map(shape => ({ shape, box: shapeBox(shape) }))
    .filter(item => item.box)
    .sort((a, b) => a.box.y - b.box.y);
  const subtitle = sortedByY[0]?.shape || shapes[5] || null;
  const paragraphShapes = rest.filter(shape => !subtitle || shape.id !== subtitle.id);
  return {
    title,
    subtitle,
    paragraphShapes
  };
}

function findShape(shapes, text, fallbackIndex) {
  return findByText(shapes, text) || shapes[fallbackIndex] || shapes[0];
}

function findByText(shapes, text) {
  return shapes.find(shape => shape.rawText.trim() === text) || shapes.find(shape => shape.rawText.includes(text));
}

function injectTextPreserve(xml, id, values, keywords = []) {
  return replaceShapeXml(xml, id, shapeXml => {
    const paragraphs = [...shapeXml.matchAll(/<a:p\b[\s\S]*?<\/a:p>/g)].map(match => match[0]);
    if (!paragraphs.length) return shapeXml;
    let paragraphIndex = 0;
    return shapeXml.replace(/<a:p\b[\s\S]*?<\/a:p>/g, paragraphXml => {
      const value = paragraphIndex < values.length ? values[paragraphIndex] : '';
      const next = paragraphs.length === 1 && values.length > 1
        ? replaceParagraphRunsSequential(paragraphXml, values)
        : replaceParagraphText(paragraphXml, value, keywords);
      paragraphIndex += 1;
      return next;
    });
  });
}

function injectAxisContentPreserve(xml, id, column, size) {
  const values = [column?.intro || '', ...ensureArray(column?.bullets, 3).slice(0, 3)];
  const keywords = ensureArray(column?.keywords, 0);
  return replaceShapeXml(xml, id, shapeXml => {
    let index = 0;
    const replaced = shapeXml.replace(/<a:p\b[\s\S]*?<\/a:p>/g, paragraphXml => {
      const value = index < values.length ? values[index] : '';
      index += 1;
      return replaceParagraphText(paragraphXml, value, keywords);
    });
    return size ? setRunSize(replaced, size) : replaced;
  });
}

function replaceShapeXml(xml, id, transform) {
  const pos = xml.indexOf(`id="${id}"`);
  if (pos === -1) return xml;
  const sp0 = xml.lastIndexOf('<p:sp', pos);
  const sp1 = xml.indexOf('</p:sp>', pos) + 7;
  if (sp0 < 0 || sp1 < 7) return xml;
  const block = xml.slice(sp0, sp1);
  return xml.slice(0, sp0) + transform(block) + xml.slice(sp1);
}

function stretchShapeLike(xml, id, shapes) {
  const boxes = shapes.map(shapeBox).filter(Boolean);
  if (!boxes.length) return xml;
  const x = Math.min(...boxes.map(box => box.x));
  const y = Math.min(...boxes.map(box => box.y));
  const right = Math.max(...boxes.map(box => box.x + box.cx));
  const bottom = Math.max(...boxes.map(box => box.y + box.cy));
  return replaceShapeXml(xml, id, shapeXml => shapeXml
    .replace(/<a:off x="[^"]*" y="[^"]*"\//, `<a:off x="${x}" y="${y}"/`)
    .replace(/<a:ext cx="[^"]*" cy="[^"]*"\//, `<a:ext cx="${right - x}" cy="${bottom - y}"/`));
}

function shapeBox(shape) {
  const off = (shape.xml.match(/<a:off x="([^"]+)" y="([^"]+)"/) || []).slice(1).map(Number);
  const ext = (shape.xml.match(/<a:ext cx="([^"]+)" cy="([^"]+)"/) || []).slice(1).map(Number);
  if (off.length < 2 || ext.length < 2 || off.some(Number.isNaN) || ext.some(Number.isNaN)) return null;
  return { x: off[0], y: off[1], cx: ext[0], cy: ext[1] };
}

function setShapeTextSize(xml, id, size) {
  return replaceShapeXml(xml, id, shapeXml => setRunSize(shapeXml, size));
}

function setRunSize(xml, size) {
  return String(xml || '').replace(/\bsz="\d+"/g, `sz="${size}"`);
}

function replaceParagraphText(paragraphXml, value, keywords = []) {
  const cleanKeywords = ensureArray(keywords, 0).filter(keyword => String(value || '').toLowerCase().includes(String(keyword).toLowerCase()));
  if (cleanKeywords.length) return replaceParagraphWithKeywordRuns(paragraphXml, value, cleanKeywords);
  return replaceParagraphRunsSequential(paragraphXml, [value]);
}

function replaceParagraphRunsSequential(paragraphXml, values) {
  let index = 0;
  return paragraphXml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, () => {
    const value = index < values.length ? values[index] : '';
    index += 1;
    return `<a:t>${escapeXml(value)}</a:t>`;
  });
}

function replaceParagraphWithKeywordRuns(paragraphXml, text, keywords) {
  const firstRun = (paragraphXml.match(/<a:r\b[\s\S]*?<\/a:r>/) || [])[0];
  if (!firstRun) return replaceParagraphRunsSequential(paragraphXml, [text]);
  const firstRPr = (firstRun.match(/<a:rPr\b[\s\S]*?<\/a:rPr>/) || [])[0] || '<a:rPr/>';
  const runs = makeRunsFromTemplate(text, keywords, firstRPr);
  const withoutRuns = paragraphXml.replace(/<a:r\b[\s\S]*?<\/a:r>/g, '');
  if (withoutRuns.includes('<a:endParaRPr')) return withoutRuns.replace(/<a:endParaRPr\b/, `${runs}<a:endParaRPr`);
  return withoutRuns.replace('</a:p>', `${runs}</a:p>`);
}

function makeRunsFromTemplate(text, keywords, rPrXml) {
  const cleanKeywords = (keywords || []).map(String).filter(Boolean);
  if (!cleanKeywords.length) return makeTemplateRun(text, rPrXml, false);
  const re = new RegExp(`(${cleanKeywords.map(escapeRegExp).join('|')})`, 'gi');
  return String(text || '').split(re).filter(Boolean).map(part => {
    const bold = cleanKeywords.some(keyword => keyword.toLowerCase() === part.toLowerCase());
    return makeTemplateRun(part, rPrXml, bold);
  }).join('');
}

function makeTemplateRun(text, rPrXml, bold) {
  return `<a:r>${setRunBold(rPrXml, bold)}<a:t>${escapeXml(text)}</a:t></a:r>`;
}

function setRunBold(rPrXml, bold) {
  if (/\sb="[01]"/.test(rPrXml)) return rPrXml.replace(/\sb="[01]"/, ` b="${bold ? 1 : 0}"`);
  return rPrXml.replace('<a:rPr', `<a:rPr b="${bold ? 1 : 0}"`);
}

function renderSlideRels(template, templateSlide) {
  return (template.slides[templateSlide]?.rels || emptyRels())
    .replace(/<Relationship\b(?=[^>]*notesSlide)[^>]*\/>/g, '')
    .replace(/<Relationship\b(?=[^>]*relationships\/slide")[^>]*\/>/g, '');
}

async function renderPresentationXml(zip, count) {
  let xml = await zip.file('ppt/presentation.xml').async('string');
  const ids = Array.from({ length: count }, (_, index) => `<p:sldId id="${256 + index}" r:id="rId${1000 + index}"/>`).join('');
  return xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${ids}</p:sldIdLst>`);
}

async function renderPresentationRels(zip, count) {
  let xml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  xml = xml.replace(/<Relationship\b(?=[^>]*relationships\/slide")[^>]*\/>/g, '');
  const slideRels = Array.from({ length: count }, (_, index) => {
    return `<Relationship Id="rId${1000 + index}" Type="${SLIDE_REL_TYPE}" Target="slides/slide${index + 1}.xml"/>`;
  }).join('');
  return xml.replace('</Relationships>', `${slideRels}</Relationships>`);
}

async function renderContentTypes(zip, count) {
  let xml = await zip.file('[Content_Types].xml').async('string');
  xml = xml.replace(/<Override\b(?=[^>]*presentationml\.slide\+xml)[^>]*\/>/g, '');
  const overrides = Array.from({ length: count }, (_, index) => {
    return `<Override ContentType="${SLIDE_CONTENT_TYPE}" PartName="/ppt/slides/slide${index + 1}.xml"/>`;
  }).join('');
  return xml.replace('</Types>', `${overrides}</Types>`);
}

async function updateAppSlideCount(zip, count) {
  if (!zip.file('docProps/app.xml')) return;
  const xml = await zip.file('docProps/app.xml').async('string');
  zip.file('docProps/app.xml', xml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${count}</Slides>`));
}

function removeGeneratedSlideParts(zip) {
  Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/(slide\d+\.xml|_rels\/slide\d+\.xml\.rels)$/i.test(name))
    .forEach(name => zip.remove(name));
}

function slideNo(name) {
  return Number((name.match(/slide(\d+)\.xml/) || [])[1] || 0);
}

function splitRow(text, index) {
  const value = s(text);
  const colon = value.indexOf(':');
  if (colon > 0 && colon < 34) return [value.slice(0, colon + 1), ` ${value.slice(colon + 1).trim()}`];
  const words = value.split(/\s+/).filter(Boolean);
  const label = words.slice(0, 3).join(' ') || `Point ${index + 1}:`;
  const body = words.slice(3).join(' ');
  return [label.endsWith(':') ? label : `${label}:`, body ? ` ${body}` : ''];
}

function defaultRow(index) {
  return [
    'Cadrage: définir le périmètre et les responsabilités',
    'Exécution: déployer les actions prioritaires avec un pilotage régulier',
    'Suivi: mesurer les résultats et traiter les points de blocage',
    'Ajustement: capitaliser sur les apprentissages et renforcer le dispositif'
  ][index] || 'Action: préciser les responsabilités et les résultats attendus';
}

function rowsToParagraph(rows) {
  return ensureArray(rows, 0)
    .map(row => typeof row === 'string' ? row : row?.text)
    .filter(Boolean)
    .join(' ');
}

function defaultParagraph() {
  return 'Le sujet est présenté comme un axe unique, avec les principaux constats, les implications opérationnelles et les priorités à traiter dans une formulation continue.';
}

function oneAxisFontSize(text) {
  const words = String(text || '').split(/\s+/).filter(Boolean).length;
  const chars = String(text || '').length;
  if (words <= 34 && chars <= 230) return '1700';
  if (words <= 44 && chars <= 290) return '1600';
  if (words <= 52 && chars <= 330) return '1500';
  return '1400';
}

function bridgeFromColumns(columns) {
  const left = columns[0]?.label || 'Premier axe';
  const right = columns[1]?.label || 'second axe';
  return `${left} et ${right} structurent les priorités à traiter et les décisions opérationnelles à engager.`;
}

function normalizeAxisColumnForPptx(column) {
  return {
    ...column,
    label: fitTitleForPptx(column?.label || 'AXE', 14, 2).replace(/[^\p{L}\p{N}\s]/gu, '').toUpperCase() || 'AXE',
    intro: fitTextForPptx(column?.intro || '', 240, 38),
    bullets: ensureArray(column?.bullets, 3).slice(0, 3).map(item => fitTextForPptx(item, 80, 12)),
    keywords: column?.keywords || []
  };
}

function normalizeCardColumnForPptx(column) {
  return {
    ...column,
    header: fitTitleForPptx(column?.header || '', 26, 4),
    body: fitTextForPptx(column?.body || '', 135, 22)
  };
}

function fitTitleForPptx(value, maxChars, maxWords) {
  return fitTextForPptx(value, maxChars, maxWords).replace(/[:;,.]+$/g, '').trim();
}

function fitTextForPptx(value, maxChars, maxWords) {
  const words = s(value).replace(/\s+/g, ' ').trim().split(/\s+/).filter(Boolean);
  let text = words.slice(0, maxWords).join(' ');
  while (text.length > maxChars && text.includes(' ')) text = text.replace(/\s+\S+$/, '');
  if (text.length > maxChars) text = text.slice(0, maxChars).replace(/\s+\S*$/, '').trim();
  return text || (words[0] || '').slice(0, maxChars);
}

function removeEmoji(value) {
  return s(value)
    .replace(/[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\uFE0F]/gu, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function ensureArray(value, min) {
  const arr = Array.isArray(value) ? value.slice() : value == null ? [] : [value];
  while (arr.length < min) arr.push('');
  return arr;
}

function emptyRels() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
}

function stripEmDashes(value) {
  return String(value || '').replace(/[\u2013\u2014]/g, '-');
}

function s(value) {
  return stripEmDashes(value == null ? '' : value);
}

function escapeXml(value) {
  return s(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function unescapeXml(value) {
  return String(value || '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

module.exports = { buildPptxBuffer };
