# ASCENCE PPTX Rebrander

Node.js app that converts an uploaded PowerPoint deck into the `Template So Far.pptx` house style. The backend extracts text from the source deck, asks OpenAI to structure it into a normalized presentation JSON, lets the frontend edit/regenerate that JSON, then injects the final content into the real PPTX template.

## Stack

- Backend: Node.js, Express
- Frontend: vanilla HTML, CSS, JS in `public/index.html`
- Uploads: `multer`
- PPTX read: `adm-zip`, `xml2js`
- PPTX write/injection: `jszip`
- LLM: OpenAI SDK, using `OPENAI_API_KEY`

## Main Files

- `server.js`: Express app and API routes
- `parser.js`: extracts text from uploaded PPTX slides
- `llmService.js`: OpenAI calls, prompt, JSON validation/sanitization
- `pptxBuilder.js`: loads `Template So Far.pptx`, reads slide XML with JSZip, injects content while preserving template styling
- `public/index.html`: current frontend
- `Template So Far.pptx`: source template used for output

## Environment

Create `.env`:

```bash
OPENAI_API_KEY=your_key_here
OPENAI_MODEL=gpt-4o
PORT=3000
```

`OPENAI_MODEL` and `PORT` are optional.

Run:

```bash
npm install
npm start
```

## Workflow

1. User uploads a `.pptx` and optional description.
2. `POST /api/process` parses all slide text from the uploaded PPTX.
3. Backend sends extracted text, user description, and current France date to OpenAI.
4. OpenAI returns a normalized `presentationState` JSON.
5. Frontend displays/edit this JSON.
6. User can regenerate one slide or the full deck with density modes.
7. `POST /api/download` sends final JSON to backend.
8. Backend clones/injects content into `Template So Far.pptx` and returns a binary PPTX.

## Presentation State Contract

The frontend should keep one object shaped like this:

```json
{
  "title": "Deck title",
  "subtitle": "Subtitle",
  "subsubtitle": "Context",
  "date": "Avril 2026",
  "sections": [
    {
      "name": "Section name",
      "slides": [
        {
          "originalSlideIndex": 2,
          "layout": "A",
          "content": {}
        }
      ]
    }
  ],
  "closingTagline": "ASCENCE ADVISORY"
}
```

### Layout A, Two Axes

```json
{
  "title": "Transformation et fidélisation",
  "bridge": "Short central synthesis sentence.",
  "columns": [
    {
      "label": "TRANSFO",
      "intro": "Intro paragraph for the left axis.",
      "bullets": ["Bullet one", "Bullet two", "Bullet three"],
      "keywords": ["keyword", "important phrase"]
    },
    {
      "label": "FIDÉLITÉ",
      "intro": "Intro paragraph for the right axis.",
      "bullets": ["Bullet one", "Bullet two", "Bullet three"],
      "keywords": ["keyword", "important phrase"]
    }
  ]
}
```

### Layout B, Three Cards

```json
{
  "title": "Three-part title",
  "columns": [
    { "header": "Header 1", "body": "Body paragraph." },
    { "header": "Header 2", "body": "Body paragraph." },
    { "header": "Header 3", "body": "Body paragraph." }
  ]
}
```

### Layout C, Four Rows

```json
{
  "title": "Action plan",
  "subtitle": "Méthode cible",
  "rows": [
    { "icon": "•", "text": "Cadrage: define scope and owners" },
    { "icon": "•", "text": "Exécution: deploy priority actions" },
    { "icon": "•", "text": "Suivi: track results and blockers" },
    { "icon": "•", "text": "Ajustement: refine and scale" }
  ]
}
```

## API Endpoints

### `POST /api/process`

Accepts multipart form data.

Fields:

- `file`: required `.pptx`
- `description`: optional text

Returns:

```json
{
  "title": "...",
  "subtitle": "...",
  "subsubtitle": "...",
  "date": "...",
  "sections": [],
  "closingTagline": "..."
}
```

Example frontend call:

```js
const form = new FormData();
form.append("file", fileInput.files[0]);
form.append("description", description);

const res = await fetch("/api/process", {
  method: "POST",
  body: form
});
const presentationState = await res.json();
```

### `POST /api/regenerate-slide`

Regenerates one content slide.

Body:

```json
{
  "slide": {
    "layout": "A",
    "content": {}
  },
  "layout": "A",
  "mode": "equilibre"
}
```

Modes:

- `equilibre`
- `detaille`
- `bref`
- `concis`

Returns the updated slide object:

```json
{
  "layout": "A",
  "content": {}
}
```

### `POST /api/regenerate-presentation`

Regenerates the full presentation JSON.

Body:

```json
{
  "state": { "title": "...", "sections": [] },
  "mode": "detaille"
}
```

Returns a full updated `presentationState`.

### `POST /api/download`

Builds the PPTX from the current `presentationState`.

Body:

```json
{
  "title": "...",
  "subtitle": "...",
  "sections": []
}
```

Returns binary `.pptx`.

Example:

```js
const res = await fetch("/api/download", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify(presentationState)
});

const blob = await res.blob();
const url = URL.createObjectURL(blob);
const link = document.createElement("a");
link.href = url;
link.download = "Ascence_Rebrandee.pptx";
link.click();
URL.revokeObjectURL(url);
```

### `GET /api/health`

Returns template status:

```json
{
  "ok": true,
  "template": "Template So Far.pptx",
  "templateSlides": 7
}
```

## Template Injection Notes

`pptxBuilder.js` does not recreate styles manually. It:

1. Loads `Template So Far.pptx` with JSZip.
2. Reads `ppt/slides/slide*.xml`.
3. Extracts real text shapes and IDs from the template XML.
4. Clones template slides into a new deck order.
5. Replaces only `<a:t>` text values inside existing template paragraphs/runs.
6. Preserves template font sizes, fonts, colors, bullets, spacing, masters, images, and geometry.

This is important. If a new frontend is built, it should not try to style the PPTX. It should only edit the `presentationState`; the backend handles injection.

## Building A New Frontend

A new frontend only needs to:

1. Upload a PPTX to `/api/process`.
2. Store the returned `presentationState`.
3. Render/edit that JSON.
4. Call `/api/regenerate-slide` for one-slide rewrites.
5. Call `/api/regenerate-presentation` for full-deck rewrites.
6. Call `/api/download` with final state.

Recommended UI flow:

- Screen 1: upload PPTX and optional context.
- Screen 2: editor with slide list and editable fields.
- Toolbar: density selector plus regenerate slide/deck buttons.
- Final action: download PPTX.

Do not send edited HTML or visual layout data to the backend. Send only the JSON contract above.

## Current Date Handling

`server.js` computes the current France date with `Europe/Paris` and passes it into OpenAI. The prompt tells OpenAI to use that date for the presentation date unless the source deck clearly specifies another date.
