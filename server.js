const express = require('express');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const multer = require('multer');
const { marked } = require('marked');

const app = express();
const PORT = process.env.PORT || 3456;

// --- Middleware ---
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

// --- File upload config ---
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    const name = crypto.randomBytes(8).toString('hex') + ext;
    cb(null, name);
  },
});
const upload = multer({
  storage,
  limits: { fileSize: 20 * 1024 * 1024 }, // 20MB
  fileFilter: (req, file, cb) => {
    const allowed = /\.(png|jpg|jpeg|gif|webp|svg|pdf|csv|xlsx|docx|md|txt|html)$/i;
    if (allowed.test(path.extname(file.originalname))) {
      cb(null, true);
    } else {
      cb(new Error('File type not allowed'));
    }
  },
});

// --- Data store ---
const STORE_PATH = path.join(__dirname, 'data', 'user_content.json');

function ensureDataDir() {
  const dir = path.dirname(STORE_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function loadStore() {
  ensureDataDir();
  if (!fs.existsSync(STORE_PATH)) return { sections: [], content: {} };
  try {
    return JSON.parse(fs.readFileSync(STORE_PATH, 'utf8'));
  } catch {
    return { sections: [], content: {} };
  }
}

function saveStore(data) {
  ensureDataDir();
  fs.writeFileSync(STORE_PATH, JSON.stringify(data, null, 2));
}

function genId() {
  return crypto.randomBytes(6).toString('hex');
}

// ========== SECTIONS API ==========

// Get all custom sections
app.get('/api/sections', (req, res) => {
  const store = loadStore();
  res.json(store.sections);
});

// Create a new section
app.post('/api/sections', (req, res) => {
  const { title, description, icon } = req.body;
  const author = req.headers['x-author'];
  if (!title || !title.trim()) return res.status(400).json({ error: 'Title is required' });
  if (!author) return res.status(400).json({ error: 'Author is required' });

  const store = loadStore();
  const section = {
    id: genId(),
    title: title.trim(),
    description: (description || '').trim(),
    icon: (icon || '📄').trim(),
    createdBy: author.trim(),
    createdAt: new Date().toISOString(),
  };
  store.sections.push(section);
  if (!store.content[section.id]) store.content[section.id] = [];
  saveStore(store);
  res.status(201).json(section);
});

// Update a section (only by creator)
app.put('/api/sections/:id', (req, res) => {
  const author = req.headers['x-author'];
  const store = loadStore();
  const section = store.sections.find(s => s.id === req.params.id);
  if (!section) return res.status(404).json({ error: 'Section not found' });
  if (!author || author.toLowerCase() !== (section.createdBy || '').toLowerCase()) {
    return res.status(403).json({ error: 'You can only edit sections you created' });
  }

  const { title, description, icon } = req.body;
  if (title) section.title = title.trim();
  if (description !== undefined) section.description = description.trim();
  if (icon) section.icon = icon.trim();
  saveStore(store);
  res.json(section);
});

// Delete a section (only by creator)
app.delete('/api/sections/:id', (req, res) => {
  const author = req.headers['x-author'];
  const store = loadStore();
  const idx = store.sections.findIndex(s => s.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Section not found' });
  const section = store.sections[idx];
  if (!author || author.toLowerCase() !== (section.createdBy || '').toLowerCase()) {
    return res.status(403).json({ error: 'You can only delete sections you created' });
  }

  store.sections.splice(idx, 1);
  delete store.content[req.params.id];
  saveStore(store);
  res.json({ success: true });
});

// ========== CONTENT API ==========

// Get content for a section (supports built-in section keys like "updates", "insights", "knowledge")
app.get('/api/content/:section', (req, res) => {
  const store = loadStore();
  const items = store.content[req.params.section] || [];
  res.json(items);
});

// Add content to a section
app.post('/api/content/:section', (req, res) => {
  const { title, body, author, attachments } = req.body;
  if (!title || !title.trim()) return res.status(400).json({ error: 'Title is required' });

  const store = loadStore();
  const sectionKey = req.params.section;
  if (!store.content[sectionKey]) store.content[sectionKey] = [];

  const bodyText = (body || '').trim();
  const item = {
    id: genId(),
    title: title.trim(),
    body: bodyText,
    bodyHtml: bodyText ? marked(bodyText) : '',
    author: (author || 'Anonymous').trim(),
    attachments: attachments || [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  store.content[sectionKey].unshift(item); // newest first
  saveStore(store);
  res.status(201).json(item);
});

// Update content item (only by original author)
app.put('/api/content/:section/:id', (req, res) => {
  const reqAuthor = req.headers['x-author'];
  const store = loadStore();
  const items = store.content[req.params.section];
  if (!items) return res.status(404).json({ error: 'Section not found' });

  const item = items.find(i => i.id === req.params.id);
  if (!item) return res.status(404).json({ error: 'Item not found' });
  if (!reqAuthor || reqAuthor.toLowerCase() !== (item.author || '').toLowerCase()) {
    return res.status(403).json({ error: 'You can only edit your own content' });
  }

  const { title, body, attachments } = req.body;
  if (title) item.title = title.trim();
  if (body !== undefined) {
    item.body = body.trim();
    item.bodyHtml = body.trim() ? marked(body.trim()) : '';
  }
  if (attachments) item.attachments = attachments;
  item.updatedAt = new Date().toISOString();

  saveStore(store);
  res.json(item);
});

// Delete content item (only by original author)
app.delete('/api/content/:section/:id', (req, res) => {
  const reqAuthor = req.headers['x-author'];
  const store = loadStore();
  const items = store.content[req.params.section];
  if (!items) return res.status(404).json({ error: 'Section not found' });

  const idx = items.findIndex(i => i.id === req.params.id);
  if (idx < 0) return res.status(404).json({ error: 'Item not found' });
  const item = items[idx];
  if (!reqAuthor || reqAuthor.toLowerCase() !== (item.author || '').toLowerCase()) {
    return res.status(403).json({ error: 'You can only delete your own content' });
  }

  items.splice(idx, 1);
  saveStore(store);
  res.json({ success: true });
});

// ========== POLISH API (Claude) ==========

app.post('/api/polish', async (req, res) => {
  const { rawText, sectionName, isNewSection, author } = req.body;
  if (!rawText || !rawText.trim()) return res.status(400).json({ error: 'Text is required' });

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    // Fallback: just clean up the text as markdown
    const title = sectionName || 'Update';
    const polished = rawText.trim();
    return res.json({ title, bodyMarkdown: polished, bodyHtml: marked(polished) });
  }

  try {
    const Anthropic = require('@anthropic-ai/sdk');
    const client = new Anthropic({ apiKey });

    const prompt = isNewSection
      ? `You are a professional editor for an internal business intelligence report at Moloco (ad-tech company).
A team member (@${author}) wants to add a NEW SECTION to the weekly report.
They wrote the following in casual/rough language:

---
${rawText}
---

Polish this into a professional report section. Return ONLY a JSON object with:
- "title": A clear, concise section heading (e.g. "APAC Spend Summary", "Campaign Performance Review")
- "bodyMarkdown": The polished content in Markdown format. Use bullet points, bold for key metrics/names, and keep it concise and data-driven. Do NOT include the title in the body.

Keep the same meaning and data points. Be professional but not overly formal. This is an internal team report.`
      : `You are a professional editor for an internal business intelligence report at Moloco (ad-tech company).
A team member (@${author}) wants to add content to the "${sectionName}" section of the weekly report.
They wrote the following in casual/rough language:

---
${rawText}
---

Polish this into professional report content that fits naturally under the "${sectionName}" section. Return ONLY a JSON object with:
- "title": A clear sub-heading for this contribution (concise, 3-8 words)
- "bodyMarkdown": The polished content in Markdown format. Use bullet points, bold for key metrics/names, and keep it concise and data-driven. Do NOT include the title in the body.

Keep the same meaning and data points. Be professional but not overly formal. This is an internal team report.`;

    const message = await client.messages.create({
      model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-20250514',
      max_tokens: 2048,
      messages: [{ role: 'user', content: prompt }],
    });

    const responseText = message.content[0].text;
    // Extract JSON from response
    const jsonMatch = responseText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      return res.status(500).json({ error: 'Failed to parse AI response' });
    }
    const parsed = JSON.parse(jsonMatch[0]);
    const bodyMd = parsed.bodyMarkdown || parsed.body || '';
    res.json({
      title: parsed.title || sectionName || 'Update',
      bodyMarkdown: bodyMd,
      bodyHtml: marked(bodyMd),
    });
  } catch (err) {
    console.error('Polish error:', err.message);
    // Fallback on error
    const title = sectionName || 'Update';
    const polished = rawText.trim();
    res.json({ title, bodyMarkdown: polished, bodyHtml: marked(polished) });
  }
});

// ========== FILE UPLOAD ==========

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  res.json({
    filename: req.file.filename,
    originalName: req.file.originalname,
    size: req.file.size,
    url: '/uploads/' + req.file.filename,
  });
});

// ========== START ==========

app.listen(PORT, () => {
  console.log(`China GDS Repository running at http://localhost:${PORT}`);
});
