'use strict';

const OpenAI = require('openai');

const MODEL = process.env.OPENAI_MODEL || 'gpt-4.1';
const MAX_OUTPUT_TOKENS = Number(process.env.OPENAI_MAX_OUTPUT_TOKENS || 32000);
const SLIDE_BATCH_SIZE = Number(process.env.OPENAI_SLIDE_BATCH_SIZE || 2);
const SLIDE_BATCH_CONCURRENCY = Number(process.env.OPENAI_SLIDE_BATCH_CONCURRENCY || 2);

function client() {
  if (!process.env.OPENAI_API_KEY) throw new Error('OPENAI_API_KEY est manquante dans .env.');
  return new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
}

async function analyzePresentation({ slides, description, templateReferences, mode = 'equilibre', franceDate = '' }) {
  const franceDateValue = franceDate || franceDateString();
  const sourceContent = getSourceContentSlides(slides);
  const deckPlan = await planDeckEnvelope({ slides, description, templateReferences, franceDate: franceDateValue });
  const skeleton = buildSectionSkeleton(deckPlan, sourceContent);
  const sectionBySlideIndex = sectionLookup(skeleton.sections);
  const batches = chunk(sourceContent, Math.max(1, SLIDE_BATCH_SIZE));

  const mappedBatches = await mapLimit(batches, Math.max(1, SLIDE_BATCH_CONCURRENCY), async batch => {
    return generateMappedSlideBatch({
      batch,
      allSlides: slides,
      sectionBySlideIndex,
      mode,
      franceDate: franceDateValue,
      description
    });
  });

  const mappedSlides = mappedBatches.flat();
  const mappedByIndex = new Map(mappedSlides.map(slide => [Number(slide.originalSlideIndex), slide]));
  const sections = skeleton.sections.map(section => ({
    name: section.name,
    slides: section.slideIndexes.map(slideIndex => {
      const sourceSlide = sourceContent.find(slide => Number(slide.slideIndex) === Number(slideIndex));
      return mappedByIndex.get(Number(slideIndex)) || fallbackSlide(sourceSlide || { slideIndex, rawText: '' });
    })
  })).filter(section => section.slides.length);

  return sanitizePresentation({
    title: deckPlan.title,
    subtitle: deckPlan.subtitle,
    subsubtitle: deckPlan.subsubtitle,
    date: deckPlan.date,
    sections,
    closingTagline: deckPlan.closingTagline
  }, slides);
}

async function regenerateSlide({ slide, layout, mode = 'equilibre', franceDate = '' }) {
  return withJsonRetry(async () => {
    const result = await completeJson([
      { role: 'system', content: buildSystemPrompt() },
      {
        role: 'user',
        content: [
          `Régénère uniquement cette slide de contenu dans le modèle ${layout}.`,
          `Mode de densité: ${modeInstruction(mode)}`,
          `Date actuelle en France: ${franceDate || franceDateString()}`,
          'Retourne uniquement ce JSON: {"layout":"A|B|C","content":{...}}.',
          'Respecte exactement les nombres de champs du template. Aucun champ visible ne doit être vide.',
          JSON.stringify(slide, null, 2)
        ].join('\n\n')
      }
    ]);
    return sanitizeSlide(result, slide, layout);
  }, 'OpenAI a retourné un JSON invalide pendant la régénération de la slide.');
}

async function regeneratePresentation({ state, mode = 'equilibre', franceDate = '' }) {
  return withJsonRetry(async () => {
    const result = await completeJson([
      { role: 'system', content: buildSystemPrompt() },
      {
        role: 'user',
        content: [
          `Régénère toute la présentation avec ce mode de densité: ${modeInstruction(mode)}`,
          `Date actuelle en France au moment de la régénération: ${franceDate || franceDateString()}`,
          'Garde le même schéma JSON. Tu peux améliorer les sections, les modèles, la formulation et la répartition du contenu.',
          'Préserve le sens du deck actuel. Remplis tous les champs du template. Ne crée jamais de numéros de page pour l’agenda.',
          JSON.stringify(state, null, 2)
        ].join('\n\n')
      }
    ]);
    return sanitizePresentation(result, flattenStateSlides(state));
  }, 'OpenAI a retourné un JSON invalide pendant la régénération de la présentation.');
}

async function completeJson(messages) {
  const response = await client().chat.completions.create({
    model: MODEL,
    temperature: 0.25,
    max_tokens: MAX_OUTPUT_TOKENS,
    response_format: { type: 'json_object' },
    messages
  });
  const content = response.choices?.[0]?.message?.content;
  if (!content) throw new Error('OpenAI a retourné une réponse vide.');
  return JSON.parse(content);
}

async function withJsonRetry(fn, message) {
  try {
    return await fn();
  } catch (firstError) {
    try {
      return await fn();
    } catch (secondError) {
      const detail = secondError.message || firstError.message;
      throw new Error(`${message} ${detail}`);
    }
  }
}

async function planDeckEnvelope({ slides, description, templateReferences, franceDate }) {
  try {
    return await withJsonRetry(async () => completeJson([
      { role: 'system', content: buildDeckPlanPrompt() },
      {
        role: 'user',
        content: [
          `Date actuelle en France: ${franceDate || franceDateString()}`,
          `Description optionnelle utilisateur:\n${description || '(aucune)'}`,
          `Nombre de slides dans le template de référence: ${templateReferences.length}.`,
          'Slides source à organiser. Utilise slideIndex pour construire les sections, sans reformuler le contenu slide par slide ici:',
          JSON.stringify(slides.map(slide => ({
            slideIndex: slide.slideIndex,
            rawText: fitText(slide.rawText, 1200, 180)
          })), null, 2)
        ].join('\n\n')
      }
    ]), 'OpenAI a retourné un JSON invalide pendant la planification du deck.');
  } catch (error) {
    console.warn(error.message);
    return fallbackDeckPlan(slides, franceDate);
  }
}

async function generateMappedSlideBatch({ batch, allSlides, sectionBySlideIndex, mode, franceDate, description }) {
  try {
    const result = await withJsonRetry(async () => completeJson([
      { role: 'system', content: buildSystemPrompt() },
      {
        role: 'user',
        content: [
          'Tu reformules uniquement les slides source listées dans "slides_a_generer".',
          'RÈGLE ABSOLUE: une slide source = une slide JSON de sortie. Ne fusionne jamais deux slides. Ne crée jamais une slide pour un autre slideIndex.',
          'Pour une slide longue, compresse seulement cette slide dans le template. N’utilise les voisines que pour clarifier le vocabulaire, jamais pour ajouter des faits absents.',
          'Retourne uniquement ce JSON: {"slides":[{"originalSlideIndex":2,"layout":"A|B|C","content":{...}}]}.',
          `Mode de densité: ${modeInstruction(mode)}`,
          `Date actuelle en France: ${franceDate || franceDateString()}`,
          `Description optionnelle utilisateur:\n${description || '(aucune)'}`,
          `Sections cibles par slide:\n${JSON.stringify(Object.fromEntries(batch.map(slide => [slide.slideIndex, sectionBySlideIndex.get(Number(slide.slideIndex)) || 'Contenu'])), null, 2)}`,
          `Contexte voisin:\n${JSON.stringify(neighborContext(batch, allSlides), null, 2)}`,
          `slides_a_generer:\n${JSON.stringify(batch, null, 2)}`
        ].join('\n\n')
      }
    ]), 'OpenAI a retourné un JSON invalide pendant la génération slide par slide.');
    return sanitizeMappedBatch(result, batch);
  } catch (error) {
    console.warn(error.message);
    return batch.map(fallbackSlide);
  }
}

function buildDeckPlanPrompt() {
  return `
Tu planifies uniquement la structure globale d’un deck PowerPoint rebrandé.
Retourne uniquement du JSON valide, sans prose, sans markdown.

Objectif:
- Déduire le titre, sous-titre, date et 2 à 4 sections d’agenda.
- Assigner chaque slide source lisible après la slide 1 à une section.
- Ne reformule pas encore le contenu détaillé des slides.
- Ne supprime aucun slideIndex lisible et ne l’assigne pas deux fois.
- L’agenda aura seulement les noms de sections, aucun numéro de page.

JSON obligatoire:
{
  "title": "...",
  "subtitle": "...",
  "subsubtitle": "...",
  "date": "...",
  "sections": [
    { "name": "Nom court", "slideIndexes": [2, 3, 4] }
  ],
  "closingTagline": "ASCENCE ADVISORY"
}

Contraintes:
- title: maximum 52 caractères et 7 mots.
- subtitle: maximum 48 caractères et 7 mots.
- subsubtitle: maximum 56 caractères et 8 mots.
- section name: 14 à 34 caractères, maximum 5 mots.
- Crée 2 à 4 sections maximum.
- Utilise la date actuelle en France si le deck ne contient pas une date explicite plus pertinente.
`.trim();
}

function buildSystemPrompt() {
  return `
Tu es un expert senior en conseil et en reformatage de présentations PowerPoint.
Tu transformes un deck source en contenu prêt à injecter dans le vrai template ASCENCE ADVISORY.
Retourne uniquement du JSON valide, sans prose, sans markdown, sans balises de code.
N’utilise jamais de tiret cadratin. Utilise des virgules, deux-points, points-virgules ou phrases courtes.

LANGUE:
- Si le deck source est majoritairement en français, écris tout en français professionnel.
- Si le deck source est majoritairement dans une autre langue, conserve cette langue.

DATE:
- Utilise la date actuelle en France fournie dans le message utilisateur pour le champ "date", sauf si le deck source impose explicitement une autre date.
- Le champ date doit être court et lisible, par exemple "Avril 2026" ou "29 avril 2026", selon le niveau de précision du deck source.

RÈGLES CRITIQUES DE FIT TEMPLATE:
- Le contenu sera injecté dans le fichier réel "Template So Far.pptx", dans ses vrais emplacements fixes.
- La génération PPTX conserve exactement les styles du template: polices, tailles, couleurs, alignements, interlignes, puces et espacements. Tu dois donc adapter le volume de texte aux zones existantes, jamais demander une mise en forme différente.
- Aucun champ visible ne doit être vide. Pas de point 4 vide, pas de label vide, pas de body vide, pas de "TBD".
- Ne crée jamais de contenu d’agenda. L’agenda est généré uniquement avec les noms de sections, sans numéros de page.
- Garde un volume de mots aussi proche que possible du deck source. Cible plus ou moins 15 pour cent, sauf demande de densité contraire.
- Préserve les chiffres, dates, noms propres, acronymes, constats, risques, actions, priorités et détails opérationnels.
- Supprime les doublons et le remplissage, mais ne supprime pas les idées substantielles.
- Si une slide source est pauvre, enrichis prudemment avec le contexte et les slides voisines.
- Si une slide source est dense, répartis les détails dans les champs exacts du template.

LIMITES VISUELLES STRICTES:
- Tous les titres visibles doivent tenir sur une seule ligne dans le template. Reformule court, jamais de titre sur deux lignes.
- Titre principal du deck: maximum 52 caractères et 7 mots.
- Title de slide contenu: 26 à 48 caractères, maximum 7 mots.
- Nom de section: 14 à 34 caractères, maximum 5 mots.
- Agenda: noms courts uniquement, aucun numéro, aucun leader pointillé.
- Si un texte dépasse la limite, raccourcis en gardant les chiffres, noms propres et verbes d’action.
- Ne mets pas de phrases longues dans les bullets ou labels. Les zones du template sont petites.
- Évite les retours à la ligne manuels, parenthèses longues, listes dans un champ, slashs répétés et formulations avec plusieurs propositions.

JSON obligatoire:
{
  "title": "...",
  "subtitle": "...",
  "subsubtitle": "...",
  "date": "...",
  "sections": [
    {
      "name": "Nom de section",
      "slides": [
        {
          "originalSlideIndex": 2,
          "layout": "A",
          "content": {}
        }
      ]
    }
  ],
  "closingTagline": "..."
}

Modèle A, deux axes:
{ "title": "...", "bridge": "...", "columns": [{ "label": "...", "intro": "...", "bullets": ["...", "...", "..."], "keywords": ["...", "...", "..."] }, { "label": "...", "intro": "...", "bullets": ["...", "...", "..."], "keywords": ["...", "...", "..."] }] }
- Exactement 2 colonnes.
- Ce modèle correspond à la slide template avec deux axes latéraux et un court paragraphe central.
- Title: 26 à 48 caractères, une seule ligne.
- Bridge: 12 à 18 mots, une phrase de synthèse située au centre. Elle explique le lien entre les deux axes sans répéter les intros.
- Label: 1 à 2 mots, en MAJUSCULES, 14 caractères maximum, sans ponctuation.
- Intro: 18 à 26 mots, une phrase détaillée mais compacte, 165 caractères maximum. Elle doit contextualiser l’axe, pas seulement annoncer un thème.
- Exactement 3 bullets par colonne.
- Bullet: 5 à 9 mots, 58 caractères maximum, concret, sans point final.
- Keywords: exactement 2 à 3 mots ou courtes expressions par axe. Chaque keyword doit apparaître tel quel dans l’intro ou les bullets du même axe. Ces keywords seront mis en gras dans PowerPoint.
- À utiliser pour comparaison, deux axes, diagnostic vs cible, risques vs actions, transformation vs fidélisation.
- Ne choisis pas A si le contenu n’a pas deux axes naturels. Utilise B ou C à la place.

Modèle B, trois cartes:
{ "title": "...", "columns": [{ "header": "...", "body": "..." }, { "header": "...", "body": "..." }, { "header": "...", "body": "..." }] }
- Exactement 3 colonnes.
- Title: 26 à 48 caractères, une seule ligne.
- Header: 2 à 4 mots, 26 caractères maximum.
- Body: 14 à 22 mots, 135 caractères maximum, paragraphe compact.
- À utiliser pour trois piliers, trois leviers, trois phases, trois options ou trois constats.

Modèle C, un axe en paragraphe:
{ "title": "...", "subtitle": "...", "paragraph": "..." }
- Ce modèle remplace l’ancienne liste en lignes par un seul grand bloc de texte continu.
- Title: 26 à 48 caractères, une seule ligne.
- Subtitle: 2 à 4 mots, 32 caractères maximum, pas une phrase longue.
- Paragraph: 45 à 70 mots, 430 caractères maximum, une seule idée structurée en prose fluide.
- Le paragraphe doit être continu, sans bullets, sans liste numérotée et sans retours à la ligne.
- Tu peux inclure un seul emoji si cela aide vraiment le sens ou la lisibilité, sinon aucun emoji.
- À utiliser pour un axe unique, un constat dense, une explication narrative, une synthèse opérationnelle ou une slide qui ne se divise pas naturellement.

CHOIX STRUCTURE:
- Crée 2 à 4 sections maximum, car l’agenda du template a quatre lignes visibles.
- Attention: les 2 à 4 sections sont uniquement des regroupements agenda. Elles peuvent contenir beaucoup de slides.
- Ne résume jamais plusieurs slides source en une seule slide de contenu.
- Pour chaque slide source lisible après la slide 1, crée exactement une slide de contenu dans le JSON.
- Conserve le champ originalSlideIndex pour chaque slide de contenu. Il doit correspondre au slideIndex source.
- Si le deck source contient 58 slides et que la première est une page de titre, retourne environ 57 slides de contenu, réparties dans les sections.
- Les slides pauvres, de transition ou très courtes doivent quand même devenir une slide de contenu utile, enrichie avec le contexte voisin si nécessaire.
- Ne fusionne pas, ne saute pas et ne condense pas les slides pour raccourcir la réponse.
- Chaque section doit contenir au moins une slide de contenu.
- Choisis A, B ou C selon la forme du contenu, pas au hasard.
- Si tu choisis C, fournis toujours un paragraphe complet, jamais des lignes séparées.
- Ne rends jamais des champs vides sous prétexte que le contenu source est court.
`.trim();
}

function modeInstruction(mode) {
  const modes = {
    detaille: 'Plus détaillé, préserver toutes les idées source slide par slide. Cible 110 à 125 pour cent du volume de chaque slide source, sans fusion entre slides.',
    equilibre: 'Équilibré, garder un volume très proche de chaque slide source et remplir proprement le template. Cible 95 à 115 pour cent par slide.',
    bref: 'Plus bref, garder les points les plus forts tout en préservant le sens. Cible 70 à 85 pour cent du volume source.',
    concis: 'Très concis, phrases courtes, aucun remplissage, claims précis. Cible 55 à 70 pour cent du volume source.'
  };
  return modes[mode] || modes.equilibre;
}

function franceDateString(granularity = 'full') {
  const options = granularity === 'month'
    ? { timeZone: 'Europe/Paris', month: 'long', year: 'numeric' }
    : { timeZone: 'Europe/Paris', day: 'numeric', month: 'long', year: 'numeric' };
  return new Intl.DateTimeFormat('fr-FR', options).format(new Date());
}

function sanitizePresentation(result, sourceSlides) {
  const clean = removeEmDash(result || {});
  const sections = Array.isArray(clean.sections) && clean.sections.length
    ? clean.sections.slice(0, 4)
    : fallbackSections(sourceSlides);
  const normalizedSections = ensureSourceCoverage(sections.map((section, sectionIndex) => ({
    name: fitTitle(section.name || `Section ${sectionIndex + 1}`, 34, 5),
    slides: normalizeSlides(section.slides || [], sourceSlides)
  })).filter(section => section.slides.length), sourceSlides);

  return {
    title: fitTitle(clean.title || firstWords(sourceSlides[0]?.rawText, 8) || 'ASCENCE ADVISORY', 52, 7),
    subtitle: fitTitle(clean.subtitle || 'Présentation rebrandée', 48, 7),
    subsubtitle: fitTitle(clean.subsubtitle || clean['sub-subtitle'] || 'Synthèse de travail', 56, 8),
    date: clean.date || franceDateString('month'),
    sections: normalizedSections,
    closingTagline: clean.closingTagline || 'ASCENCE ADVISORY'
  };
}

function sanitizeSlide(result, originalSlide, requestedLayout) {
  const layout = ['A', 'B', 'C'].includes(result.layout) ? result.layout : requestedLayout;
  return {
    ...originalSlide,
    layout,
    content: normalizeContent(layout, result.content || originalSlide.content || {})
  };
}

function sanitizeMappedBatch(result, sourceBatch) {
  const generated = Array.isArray(result?.slides) ? result.slides : [];
  return sourceBatch.map(sourceSlide => {
    const match = generated.find(slide => Number(slide.originalSlideIndex) === Number(sourceSlide.slideIndex));
    if (!match) return fallbackSlide(sourceSlide);
    const requestedLayout = ['A', 'B', 'C'].includes(match.layout) ? match.layout : chooseLayoutForSource(sourceSlide.rawText);
    return sanitizeSlide({
      originalSlideIndex: sourceSlide.slideIndex,
      layout: requestedLayout,
      content: match.content || {}
    }, {
      originalSlideIndex: sourceSlide.slideIndex,
      layout: requestedLayout,
      content: {}
    }, requestedLayout);
  });
}

function normalizeSlides(slides, sourceSlides) {
  if (!slides.length) return [];
  return slides.map(slide => {
    const layout = ['A', 'B', 'C'].includes(slide.layout) ? slide.layout : 'C';
    return {
      originalSlideIndex: Number(slide.originalSlideIndex) || null,
      layout,
      content: normalizeContent(layout, slide.content || {})
    };
  });
}

function ensureSourceCoverage(sections, sourceSlides) {
  const sourceContent = getSourceContentSlides(sourceSlides);
  if (!sourceContent.length) return sections;

  const generatedSlides = sections.flatMap(section => section.slides || []);
  const indexed = new Set(generatedSlides.map(slide => Number(slide.originalSlideIndex)).filter(Boolean));
  if (!indexed.size && generatedSlides.length >= sourceContent.length) return sections;

  const missing = sourceContent.filter(slide => !indexed.has(Number(slide.slideIndex)));
  if (!missing.length) return sections;

  const targetSections = sections.length ? sections : [{ name: 'Contenu source', slides: [] }];
  const target = targetSections[targetSections.length - 1];
  missing.forEach(slide => target.slides.push(fallbackSlide(slide)));
  return targetSections.slice(0, 4);
}

function fallbackDeckPlan(sourceSlides, franceDate) {
  const contentSlides = getSourceContentSlides(sourceSlides);
  return {
    title: firstWords(sourceSlides[0]?.rawText, 7) || 'ASCENCE ADVISORY',
    subtitle: 'Présentation rebrandée',
    subsubtitle: 'Synthèse de travail',
    date: franceDate || franceDateString('month'),
    sections: distributeSlideIndexes(contentSlides, ['Contexte', 'Analyse', 'Priorités', 'Plan d’action']),
    closingTagline: 'ASCENCE ADVISORY'
  };
}

function buildSectionSkeleton(deckPlan, sourceSlides) {
  const sourceIndexes = sourceSlides.map(slide => Number(slide.slideIndex));
  const sectionNames = normalizeSectionNames(deckPlan?.sections);
  const seen = new Set();
  const sections = sectionNames.map((name, index) => ({
    name,
    slideIndexes: normalizeSlideIndexes(deckPlan?.sections?.[index]?.slideIndexes, sourceIndexes)
      .filter(slideIndex => {
        if (seen.has(slideIndex)) return false;
        seen.add(slideIndex);
        return true;
      })
  }));

  const assigned = new Set(sections.flatMap(section => section.slideIndexes));
  const missing = sourceIndexes.filter(index => !assigned.has(index));
  if (!sections.some(section => section.slideIndexes.length) || missing.length > sourceIndexes.length / 2) {
    return { sections: distributeSlideIndexes(sourceSlides, sectionNames) };
  }

  missing.forEach(index => {
    const position = sourceIndexes.indexOf(index);
    const sectionIndex = Math.min(sections.length - 1, Math.floor(position / Math.ceil(sourceIndexes.length / sections.length)));
    sections[sectionIndex].slideIndexes.push(index);
  });
  sections.forEach(section => {
    section.slideIndexes = [...new Set(section.slideIndexes)]
      .filter(index => sourceIndexes.includes(index))
      .sort((a, b) => a - b);
  });
  return { sections: sections.filter(section => section.slideIndexes.length) };
}

function normalizeSectionNames(sections) {
  const names = Array.isArray(sections)
    ? sections.slice(0, 4).map((section, index) => fitTitle(section?.name || `Section ${index + 1}`, 34, 5)).filter(Boolean)
    : [];
  return names.length ? names : ['Contexte', 'Analyse', 'Priorités', 'Plan d’action'];
}

function normalizeSlideIndexes(value, sourceIndexes) {
  if (!Array.isArray(value)) return [];
  return value
    .map(Number)
    .filter(index => sourceIndexes.includes(index));
}

function distributeSlideIndexes(sourceSlides, sectionNames) {
  const names = sectionNames.slice(0, 4);
  const indexes = sourceSlides.map(slide => Number(slide.slideIndex));
  const chunkSize = Math.max(1, Math.ceil(indexes.length / names.length));
  return names.map((name, index) => ({
    name: fitTitle(name, 34, 5),
    slideIndexes: indexes.slice(index * chunkSize, (index + 1) * chunkSize)
  })).filter(section => section.slideIndexes.length);
}

function sectionLookup(sections) {
  const lookup = new Map();
  sections.forEach(section => {
    section.slideIndexes.forEach(slideIndex => lookup.set(Number(slideIndex), section.name));
  });
  return lookup;
}

function neighborContext(batch, allSlides) {
  const indexes = new Set(batch.map(slide => Number(slide.slideIndex)));
  const wanted = new Set();
  indexes.forEach(index => {
    wanted.add(index - 1);
    wanted.add(index + 1);
  });
  return allSlides
    .filter(slide => wanted.has(Number(slide.slideIndex)) && !indexes.has(Number(slide.slideIndex)))
    .map(slide => ({
      slideIndex: slide.slideIndex,
      rawText: fitText(slide.rawText, 800, 120)
    }));
}

function chooseLayoutForSource(text) {
  const source = String(text || '').toLowerCase();
  if (/\b(vs|versus|compare|comparaison|axe|axes|actuel|cible|risque|action)\b/.test(source)) return 'A';
  const bullets = (source.match(/[•\-]\s|\n\d+[.)]/g) || []).length;
  if (bullets >= 3 || /\b(trois|3|piliers|leviers|options|phases)\b/.test(source)) return 'B';
  return 'C';
}

function chunk(items, size) {
  const chunks = [];
  for (let index = 0; index < items.length; index += size) chunks.push(items.slice(index, index + size));
  return chunks;
}

async function mapLimit(items, limit, worker) {
  const results = new Array(items.length);
  let nextIndex = 0;
  async function run() {
    while (nextIndex < items.length) {
      const index = nextIndex;
      nextIndex += 1;
      results[index] = await worker(items[index], index);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, run));
  return results;
}

function getSourceContentSlides(sourceSlides) {
  return (sourceSlides || [])
    .slice(1)
    .filter(slide => String(slide.rawText || '').trim());
}

function fallbackSlide(slide) {
  return {
    originalSlideIndex: slide.slideIndex,
    layout: 'C',
    content: normalizeContent('C', {
      title: firstWords(slide.rawText, 8) || `Slide ${slide.slideIndex}`,
      subtitle: 'Points clés',
      paragraph: paragraphFromText(slide.rawText)
    })
  };
}

function normalizeContent(layout, content) {
  const clean = removeEmDash(content || {});
  if (layout === 'A') {
    const columns = Array.isArray(clean.columns) ? clean.columns.slice(0, 2) : [];
    while (columns.length < 2) columns.push({});
    return {
      title: fitTitle(clean.title || 'Analyse structurée des priorités clés', 48, 7),
      bridge: fitSentence(clean.bridge || 'Ces deux axes structurent les priorités de transformation et orientent les décisions opérationnelles à engager.', 125, 18),
      columns: columns.map((column, index) => ({
        label: fitLabel(nonEmpty(column.label, index === 0 ? 'AXE UN' : 'AXE DEUX')),
        intro: fitSentence(nonEmpty(column.intro, 'Cette dimension synthétise les constats, les implications opérationnelles et les décisions à sécuriser rapidement.'), 165, 26),
        bullets: fillList(column.bullets, 3, ['Clarifier les priorités clés', 'Structurer les actions immédiates', 'Suivre les résultats attendus'])
          .slice(0, 3)
          .map(item => fitText(item, 58, 9)),
        keywords: normalizeKeywords(column.keywords, column, 3)
      }))
    };
  }
  if (layout === 'B') {
    const columns = Array.isArray(clean.columns) ? clean.columns.slice(0, 3) : [];
    while (columns.length < 3) columns.push({});
    return {
      title: fitTitle(clean.title || 'Trois leviers clés à activer', 48, 7),
      columns: columns.map((column, index) => ({
        header: fitText(nonEmpty(column.header, `Levier ${index + 1}`), 26, 4),
        body: fitSentence(nonEmpty(column.body, 'Ce levier précise les actions à engager, les responsabilités à clarifier et les effets attendus.'), 135, 22)
      }))
    };
  }
  const paragraph = clean.paragraph || rowsToParagraph(clean.rows);
  return {
    title: fitTitle(clean.title || 'Plan d’action opérationnel', 48, 7),
    subtitle: fitText(clean.subtitle || 'Méthode cible', 32, 4),
    paragraph: fitParagraph(nonEmpty(paragraph, paragraphFromText(defaultRows().join(' '))), 430, 70)
  };
}

function fallbackSections(sourceSlides) {
  const contentSlides = getSourceContentSlides(sourceSlides);
  return [{
    name: 'Synthèse',
    slides: contentSlides.map(slide => ({
      originalSlideIndex: slide.slideIndex,
      layout: 'C',
      content: normalizeContent('C', {
        title: firstWords(slide.rawText, 8) || `Slide ${slide.slideIndex}`,
        subtitle: 'Points clés',
        paragraph: paragraphFromText(slide.rawText)
      })
    }))
  }];
}

function flattenStateSlides(state) {
  const slides = [{ slideIndex: 1, rawText: [state.title, state.subtitle, state.subsubtitle, state.date].join(' ') }];
  let index = 2;
  (state.sections || []).forEach(section => {
    (section.slides || []).forEach(slide => {
      slides.push({ slideIndex: index, rawText: collectText(slide.content).join(' ') });
      index += 1;
    });
  });
  return slides;
}

function collectText(value, out = []) {
  if (typeof value === 'string') out.push(value);
  else if (Array.isArray(value)) value.forEach(item => collectText(item, out));
  else if (value && typeof value === 'object') Object.values(value).forEach(item => collectText(item, out));
  return out;
}

function removeEmDash(value) {
  if (typeof value === 'string') return value.replace(/[\u2013\u2014]/g, '-');
  if (Array.isArray(value)) return value.map(removeEmDash);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, val]) => [key, removeEmDash(val)]));
  }
  return value;
}

function ensureArray(value) {
  if (Array.isArray(value)) return value.map(String).filter(Boolean);
  if (!value) return [];
  return String(value).split(/\n|;/).map(item => item.trim()).filter(Boolean);
}

function fillList(value, count, defaults) {
  const list = ensureArray(value);
  while (list.length < count) list.push(defaults[list.length] || defaults[defaults.length - 1]);
  return list.map(item => nonEmpty(item, defaults[0]));
}

function nonEmpty(value, fallback) {
  const text = String(value || '').trim();
  return text || fallback;
}

function ensureColon(value) {
  const text = String(value || '').trim();
  if (text.includes(':')) return text;
  const words = text.split(/\s+/).filter(Boolean);
  const label = words.slice(0, 2).join(' ') || 'Action';
  const body = words.slice(2).join(' ') || 'préciser les responsabilités et les résultats attendus';
  return `${label}: ${body}`;
}

function fitLabel(value) {
  return fitText(value, 14, 2).replace(/[^\p{L}\p{N}\s]/gu, '').toUpperCase() || 'AXE';
}

function fitSentence(value, maxChars, maxWords) {
  const text = fitText(value, maxChars, maxWords);
  return text.replace(/[.;:,]+$/g, '') + '.';
}

function fitRow(value) {
  const text = fitText(value, 130, 18);
  const colonIndex = text.indexOf(':');
  if (colonIndex < 0) return ensureColon(text);
  const label = fitText(text.slice(0, colonIndex), 24, 3);
  const body = fitText(text.slice(colonIndex + 1).trim(), 100, 15);
  return `${label}: ${body}`;
}

function fitTitle(value, maxChars, maxWords) {
  return fitText(value, maxChars, maxWords)
    .replace(/[:;,.]+$/g, '')
    .trim();
}

function fitParagraph(value, maxChars, maxWords) {
  return fitText(String(value || '').replace(/[\r\n]+/g, ' '), maxChars, maxWords)
    .replace(/\s+/g, ' ')
    .replace(/[;:,]+$/g, '')
    .trim();
}

function fitText(value, maxChars, maxWords) {
  const words = String(value || '').replace(/\s+/g, ' ').trim().split(/\s+/).filter(Boolean);
  let text = words.slice(0, maxWords).join(' ');
  while (text.length > maxChars && text.includes(' ')) text = text.replace(/\s+\S+$/, '');
  if (text.length > maxChars) text = text.slice(0, maxChars).replace(/\s+\S*$/, '').trim();
  return text || (words[0] || '').slice(0, maxChars);
}

function rowsToParagraph(rows) {
  const list = Array.isArray(rows) ? rows : [];
  return list
    .map(row => typeof row === 'string' ? row : row?.text)
    .filter(Boolean)
    .join(' ');
}

function paragraphFromText(text) {
  return fitParagraph(text, 430, 70);
}

function normalizeKeywords(keywords, column, count) {
  const text = [column.intro, ...(Array.isArray(column.bullets) ? column.bullets : [])].join(' ');
  const provided = ensureArray(keywords)
    .map(item => fitText(item, 28, 3))
    .filter(item => item && text.toLowerCase().includes(item.toLowerCase()));
  const derived = deriveKeywords(text);
  const merged = [...provided, ...derived].filter((item, index, arr) => {
    return item && arr.findIndex(other => other.toLowerCase() === item.toLowerCase()) === index;
  });
  return merged.slice(0, count);
}

function deriveKeywords(text) {
  const stop = new Set('avec dans pour les des une aux sur par afin cette leurs sont plus moins vers entre comme ces actions resultats résultats priorites priorités operationnelles opérationnelles'.split(' '));
  return String(text || '')
    .replace(/[^\p{L}\p{N}\s-]/gu, ' ')
    .split(/\s+/)
    .map(word => word.trim())
    .filter(word => word.length > 5 && !stop.has(word.toLowerCase()))
    .slice(0, 6);
}

function defaultRows() {
  return [
    'Cadrage: définir le périmètre, les responsabilités et les priorités de travail',
    'Exécution: déployer les actions clés avec un pilotage régulier',
    'Suivi: mesurer les résultats, les risques et les points de blocage',
    'Ajustement: capitaliser sur les apprentissages et renforcer le dispositif'
  ];
}

function firstWords(text, count) {
  return String(text || '').split(/\s+/).filter(Boolean).slice(0, count).join(' ');
}

function chunkText(text) {
  const words = String(text || '').split(/\s+/).filter(Boolean);
  if (!words.length) return defaultRows();
  const rows = [];
  const size = Math.max(10, Math.ceil(words.length / 4));
  for (let i = 0; i < words.length && rows.length < 4; i += size) {
    rows.push(ensureColon(words.slice(i, i + size).join(' ')));
  }
  while (rows.length < 4) rows.push(defaultRows()[rows.length]);
  return rows;
}

module.exports = { analyzePresentation, regenerateSlide, regeneratePresentation };
