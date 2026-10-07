require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const rateLimit = require('express-rate-limit');
const { OpenAI } = require('openai');
const PDFDocument = require('pdfkit');
const PptxGenJS = require('pptxgenjs');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 5000;
const JWT_SECRET = process.env.JWT_SECRET || 'rm_academy_secret_token_key';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || 'RMVisionAdmin@2026';

// Initialize OpenAI client
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Middleware
app.use(cors());
app.use(express.json({ limit: '30mb' }));
app.use(express.urlencoded({ extended: true, limit: '30mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Upload directories
const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname.replace(/\s+/g, '_')}`)
});

const upload = multer({
  storage,
  limits: { fileSize: (parseInt(process.env.MAX_FILE_SIZE_MB) || 15) * 1024 * 1024 }
});

// In-Memory Database Store (Easy to switch to PostgreSQL / SQLite / MongoDB)
const db = {
  users: [
    {
      id: 'admin_1',
      username: 'admin',
      email: 'admin@rmvisionacademy.com',
      passwordHash: bcrypt.hashSync(ADMIN_PASSWORD, 8),
      role: 'admin',
      dailyCount: 0,
      lastReset: new Date().toDateString()
    }
  ],
  knowledgeBase: [
    {
      id: 'kb_default_1',
      title: 'R.M. Vision Academy Core Geography Framework',
      subject: 'General Geography',
      category: 'UG',
      text: 'R.M. Vision Academy specializes in Indian Monsoon mechanisms (Jet Stream theory, Tibetan heating, ITCZ shifts), West Bengal regional geography (Sundarbans mangrove ecology, North Bengal tea & topography), Plate Tectonics, Davis & Penck cycles of erosion, Cartography, and GIS (QGIS, Remote Sensing bands, NDVI).'
    }
  ],
  chats: []
};

// Geography System Prompt
const GEOGRAPHY_SYSTEM_PROMPT = `
You are the dedicated, highly specialized AI Geography Teacher, Academic Tutor, and Senior Research Assistant for R.M. Vision Academy ("🌍 R.M. Vision Academy OpenAI - Ask Me Geography").

Your expertise spans:
1. Physical Geography: Geomorphology (Davis, Penck, King, plate tectonics, coastal, glacial, fluvial processes), Climatology (Koppen, Thornthwaite, Monsoon dynamics, cyclones, pressure belts), Oceanography (currents, tides, ocean floor topography), Biogeography, Soil & Environmental Geography.
2. Human & Economic Geography: Population Geography (Demographic Transition, Malthus), Settlement & Urban Geography (Burgess, Hoyt, Harris-Ullman, Christaller), Agricultural & Industrial location theories (Von Thunen, Weber).
3. Regional Geography: India (physiography, drainage, climate, economy) and West Bengal (Darjeeling hills, Rarh plain, Sundarbans deltaic ecology, minerals, demographic patterns).
4. Technical Geography & Statistics: Cartography, GIS, Remote Sensing (sensors, spectral resolution, NDVI), GPS, Google Earth Engine. Statistics: Mean, Median, Mode, Standard Deviation, Variance, Pearson's Correlation, Linear Regression, Chi-square, t-test, Z-test, ANOVA, Sampling methods.
5. Examinations: ICSE, CBSE, West Bengal Board (WBBSE/WBCHSE), BA/BSc, MA/MSc, UGC NET & JRF, CUET.

STRICT ACADEMIC STANDARDS:
- Maintain an authoritative, encouraging, academic persona.
- When answering academic questions, follow this structure where applicable:
  Definition → Academic Explanation → Key Characteristics / Principles → Processes / Mechanics → Real-World Examples → Diagrammatic Description (textual ASCII/SVG/layout) → Exam / Research Relevance & Conclusion.
- BILINGUAL REQUIREMENT: If the user communicates or requests Bengali, answer primarily in natural Bengali script, while MANDATORILY embedding primary technical English terminology in parentheses. Example: বায়ুমণ্ডলীয় চাপ (Atmospheric Pressure), পলি শঙ্কু (Alluvial Fan), পাতের চলন (Plate Tectonics).
- ACADEMIC RESEARCH & ETHICS: Never invent false citations, research papers, DOIs, or fabricated statistical data. If uncertain, state the limits of current research.
- Always check and verify statistical and mathematical calculations step-by-step.
- Distinguish between AI-generated schematic representations and cartographically rigorous, peer-reviewed GIS cartography.
- Prioritize facts from R.M. Vision Academy knowledge base when explicitly requested.
`;

// Helper: Token Verification
function authMiddleware(req, res, next) {
  const token = req.headers['authorization']?.split(' ')[1];
  if (!token) return res.status(401).json({ error: 'Authentication required' });
  try {
    const decoded = jwt.verify(token, JWT_SECRET);
    req.user = decoded;
    next();
  } catch (err) {
    res.status(403).json({ error: 'Invalid or expired token' });
  }
}

function adminOnly(req, res, next) {
  authMiddleware(req, res, () => {
    if (req.user.role !== 'admin') {
      return res.status(403).json({ error: 'Administrative privileges required' });
    }
    next();
  });
}

// Check Daily Limit per User
function checkRateLimit(user) {
  const today = new Date().toDateString();
  if (user.lastReset !== today) {
    user.dailyCount = 0;
    user.lastReset = today;
  }
  const maxLimit = parseInt(process.env.DAILY_REQUEST_LIMIT_PER_USER) || 100;
  if (user.dailyCount >= maxLimit && user.role !== 'admin') {
    return false;
  }
  user.dailyCount++;
  return true;
}

// -------------------------------------------------------------
// ROUTES: AUTHENTICATION
// -------------------------------------------------------------
app.post('/api/auth/register', (req, res) => {
  const { username, email, password } = req.body;
  if (!email || !password || !username) {
    return res.status(400).json({ error: 'Username, email, and password are required' });
  }
  const exists = db.users.find(u => u.email.toLowerCase() === email.toLowerCase());
  if (exists) return res.status(400).json({ error: 'User already exists' });

  const newUser = {
    id: 'user_' + Date.now(),
    username,
    email: email.toLowerCase(),
    passwordHash: bcrypt.hashSync(password, 8),
    role: 'student',
    dailyCount: 0,
    lastReset: new Date().toDateString()
  };
  db.users.push(newUser);

  const token = jwt.sign({ id: newUser.id, username: newUser.username, role: newUser.role }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: newUser.id, username: newUser.username, email: newUser.email, role: newUser.role } });
});

app.post('/api/auth/login', (req, res) => {
  const { email, password } = req.body;
  const user = db.users.find(u => u.email.toLowerCase() === email?.toLowerCase());
  if (!user || !bcrypt.compareSync(password, user.passwordHash)) {
    return res.status(401).json({ error: 'Invalid email or password' });
  }
  const token = jwt.sign({ id: user.id, username: user.username, role: user.role }, JWT_SECRET, { expiresIn: '7d' });
  res.json({ token, user: { id: user.id, username: user.username, email: user.email, role: user.role } });
});

// -------------------------------------------------------------
// ROUTE: AI CHAT (Text, Multimodal Vision & Document Context)
// -------------------------------------------------------------
app.post('/api/chat', async (req, res) => {
  try {
    const { message, history = [], answerMode = 'UG/PG Level', language = 'Bilingual', imageBase64, docContext } = req.body;

    if (!message && !imageBase64 && !docContext) {
      return res.status(400).json({ error: 'Please provide a message or file to analyze.' });
    }

    // Identify user if token passed (Optional guest fallback)
    let user = { role: 'guest', dailyCount: 0, lastReset: new Date().toDateString() };
    const authHeader = req.headers['authorization'];
    if (authHeader) {
      try {
        const decoded = jwt.verify(authHeader.split(' ')[1], JWT_SECRET);
        const found = db.users.find(u => u.id === decoded.id);
        if (found) user = found;
      } catch (e) { /* guest mode */ }
    }

    if (!checkRateLimit(user)) {
      return res.status(429).json({
        error: 'Daily question limit reached for your account. Please upgrade or return tomorrow!'
      });
    }

    // Retrieve Knowledge Base highlights
    const kbMatches = db.knowledgeBase.map(k => `[Academy Note: ${k.title} (${k.subject})]: ${k.text}`).join('\n\n');

    let modeInstruction = '';
    switch (answerMode) {
      case 'Quick Answer':
        modeInstruction = 'Answer concisely in 3-4 bullet points or high-yield summary.';
        break;
      case 'Detailed Answer':
        modeInstruction = 'Provide a full descriptive breakdown covering definitions, principles, processes, and examples.';
        break;
      case 'Exam Answer':
        modeInstruction = 'Format as a high-scoring university examination answer with headings, flowcharts (in text), diagrams descriptions, and model marks distribution points.';
        break;
      case 'UG/PG Level':
        modeInstruction = 'Deliver an advanced undergraduate/postgraduate level academic response with advanced terminology, theoretical foundations, and critical analysis.';
        break;
      case 'NET/JRF Level':
        modeInstruction = 'Include UGC NET/JRF specific orientation: assertion-reason questions, seminal scholars/geographers, publication dates of theories, and direct syllabus links.';
        break;
      case 'Research Level':
        modeInstruction = 'Provide academic research perspective: methodology, spatial modeling, GIS integration, potential hypotheses, and verified academic references.';
        break;
      default:
        modeInstruction = 'Academic balanced geography answer.';
    }

    const messages = [
      {
        role: 'system',
        content: `${GEOGRAPHY_SYSTEM_PROMPT}\n\n[SELECTED ACADEMIC MODE]: ${answerMode}. ${modeInstruction}\n[SELECTED LANGUAGE MODE]: ${language}\n\n[R.M. VISION ACADEMY KNOWLEDGE BASE]:\n${kbMatches}`
      }
    ];

    // Append conversation history
    history.slice(-8).forEach(msg => {
      messages.push({ role: msg.role === 'user' ? 'user' : 'assistant', content: msg.content });
    });

    // Build current prompt payload
    let currentContent = [];

    if (docContext) {
      currentContent.push({
        type: 'text',
        text: `[ATTACHED DOCUMENT EXCERPT / CONTEXT]:\n"""\n${docContext.slice(0, 8000)}\n"""\n`
      });
    }

    if (message) {
      currentContent.push({
        type: 'text',
        text: message
      });
    }

    if (imageBase64) {
      currentContent.push({
        type: 'image_url',
        image_url: {
          url: imageBase64.startsWith('data:') ? imageBase64 : `data:image/jpeg;base64,${imageBase64}`
        }
      });
    }

    messages.push({
      role: 'user',
      content: currentContent.length === 1 && currentContent[0].type === 'text' ? currentContent[0].text : currentContent
    });

    const completion = await openai.chat.completions.create({
      model: process.env.DEFAULT_MODEL || 'gpt-4o',
      messages,
      temperature: 0.25,
      max_tokens: parseInt(process.env.MAX_TOKENS_PER_RESPONSE) || 3000
    });

    const reply = completion.choices[0].message.content;
    res.json({ reply });
  } catch (error) {
    console.error('Chat API Error:', error);
    res.status(500).json({ error: error.message || 'Error communicating with AI service.' });
  }
});

// -------------------------------------------------------------
// ROUTE: DOCUMENT PARSING (PDF, DOCX, TXT, XLSX)
// -------------------------------------------------------------
app.post('/api/upload-document', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

    const filePath = req.file.path;
    const ext = path.extname(req.file.originalname).toLowerCase();
    let extractedText = '';

    if (ext === '.pdf') {
      const dataBuffer = fs.readFileSync(filePath);
      const parsed = await pdfParse(dataBuffer);
      extractedText = parsed.text;
    } else if (ext === '.docx') {
      const result = await mammoth.extractRawText({ path: filePath });
      extractedText = result.value;
    } else if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      const workbook = xlsx.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      extractedText = xlsx.utils.sheet_to_csv(sheet);
    } else if (ext === '.txt') {
      extractedText = fs.readFileSync(filePath, 'utf8');
    } else {
      extractedText = `Uploaded document: ${req.file.originalname}. (Binary file uploaded - please ask specific questions).`;
    }

    // Clean up temporary disk file
    fs.unlink(filePath, () => {});

    res.json({
      filename: req.file.originalname,
      size: req.file.size,
      textSnippet: extractedText.slice(0, 12000)
    });
  } catch (err) {
    console.error('Upload Error:', err);
    res.status(500).json({ error: 'Failed to extract text from document.' });
  }
});

// -------------------------------------------------------------
// ROUTE: AI DIAGRAM & MAP VISUAL DRAFT GENERATION (DALL-E 3)
// -------------------------------------------------------------
app.post('/api/generate-diagram', async (req, res) => {
  try {
    const { prompt, type = 'diagram' } = req.body;
    if (!prompt) return res.status(400).json({ error: 'Diagram prompt is required.' });

    let enhancedPrompt = '';
    if (type === 'map') {
      enhancedPrompt = `A clean, academic, cartographic-style thematic geography map illustration of: "${prompt}". High contrast, clear geographical boundaries, educational color-coding, legend, compass rose, vector style, white background, no text gibberish, accurate geographical outlines.`;
    } else {
      enhancedPrompt = `A high-resolution scientific educational textbook diagram of: "${prompt}". Labeled cross-section, clean arrows indicating processes, white background, textbook style, professional educational illustration, crisp lines, modern academic look.`;
    }

    const response = await openai.images.generate({
      model: process.env.IMAGE_MODEL || 'dall-e-3',
      prompt: enhancedPrompt,
      n: 1,
      size: '1024x1024',
      quality: 'standard'
    });

    res.json({
      imageUrl: response.data[0].url,
      disclaimer: type === 'map' ? 'Notice: AI-generated map sketch for conceptual educational visualization. For exact cartography, reference SOI / NATMO / QGIS datasets.' : 'Academic diagram sketch generated by R.M. Vision Academy OpenAI.'
    });
  } catch (error) {
    console.error('Image Generation Error:', error);
    res.status(500).json({ error: error.message || 'Failed to generate geographical visual.' });
  }
});

// -------------------------------------------------------------
// ROUTE: PROFESSIONAL PDF GENERATOR (PDFKit)
// -------------------------------------------------------------
app.post('/api/generate-pdf', async (req, res) => {
  try {
    const { title, topic, content, author = 'R.M. Vision Academy' } = req.body;
    if (!title || !content) return res.status(400).json({ error: 'Title and content are required.' });

    const doc = new PDFDocument({ margin: 50, size: 'A4' });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(title.replace(/\s+/g, '_'))}.pdf"`);

    doc.pipe(res);

    // Title Page / Header Banner
    doc.rect(0, 0, doc.page.width, 110).fill('#0B2545'); // Royal Academy Blue
    doc.fillColor('#DAA520').fontSize(22).font('Helvetica-Bold').text('R.M. VISION ACADEMY', 50, 30);
    doc.fillColor('#FFFFFF').fontSize(12).font('Helvetica').text('Department of Geography & Geospatial Studies', 50, 58);
    doc.fillColor('#8DA9C4').fontSize(10).text(`Topic: ${topic || title} | Date: ${new Date().toLocaleDateString('en-GB')}`, 50, 78);

    doc.moveDown(4);
    doc.fillColor('#0B2545').fontSize(18).font('Helvetica-Bold').text(title, { align: 'center' });
    doc.moveDown(1);
    doc.strokeColor('#DAA520').lineWidth(2).moveTo(50, doc.y).lineTo(doc.page.width - 50, doc.y).stroke();
    doc.moveDown(1.5);

    // Content Body
    doc.fillColor('#1A1A1A').fontSize(10).font('Helvetica').text(content, {
      align: 'justify',
      lineGap: 4,
      paragraphGap: 8
    });

    // Footer with Page Numbers
    const range = doc.bufferedPageRange();
    for (let i = 0; i < range.count; i++) {
      doc.switchToPage(i);
      doc.fillColor('#718096').fontSize(8).text(
        `R.M. Vision Academy AI Study Series  •  Page ${i + 1} of ${range.count}`,
        50,
        doc.page.height - 40,
        { align: 'center', width: doc.page.width - 100 }
      );
    }

    doc.end();
  } catch (err) {
    console.error('PDF Generation Error:', err);
    res.status(500).json({ error: 'Failed to generate academic PDF note.' });
  }
});

// -------------------------------------------------------------
// ROUTE: POWERPOINT GENERATOR (pptxgenjs)
// -------------------------------------------------------------
app.post('/api/generate-ppt', async (req, res) => {
  try {
    const { title, subtitle, targetLevel = 'BA/BSc Geography', slides = [] } = req.body;

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    // Slide 1: Title Slide (Royal Blue & Gold theme)
    const slide1 = pptx.addSlide();
    slide1.background = { color: '0B2545' };
    slide1.addText('R.M. VISION ACADEMY', {
      x: 0.8, y: 1.2, w: '85%', fontSize: 24, bold: true, color: 'DAA520', fontFace: 'Calibri'
    });
    slide1.addText(title || 'Geography Academic Presentation', {
      x: 0.8, y: 2.2, w: '85%', fontSize: 36, bold: true, color: 'FFFFFF', fontFace: 'Arial'
    });
    slide1.addText(`${subtitle || 'Academic Study Module'} | ${targetLevel}`, {
      x: 0.8, y: 3.5, w: '85%', fontSize: 18, color: '8DA9C4', fontFace: 'Calibri'
    });
    slide1.addText(`Prepared via R.M. Vision Academy OpenAI  •  ${new Date().toLocaleDateString('en-GB')}`, {
      x: 0.8, y: 6.2, w: '85%', fontSize: 11, color: 'E0E1DD'
    });

    // Content Slides
    if (slides && slides.length > 0) {
      slides.forEach((s) => {
        const slide = pptx.addSlide();
        slide.background = { color: 'F8F9FA' };

        // Banner Header
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: '100%', h: 1.1, fill: { color: '0B2545' } });
        slide.addText(s.slideTitle || 'Geographical Concept', {
          x: 0.8, y: 0.25, fontSize: 22, bold: true, color: 'FFFFFF', fontFace: 'Arial'
        });
        slide.addText('R.M. Vision Academy', {
          x: 10.5, y: 0.35, fontSize: 12, bold: true, color: 'DAA520', align: 'right'
        });

        // Content
        const bulletPoints = Array.isArray(s.points) ? s.points.map(p => ({ text: p, options: { breakLine: true, fontSize: 15, color: '2B2D42' } })) : [{ text: s.content || '', options: { fontSize: 15, color: '2B2D42' } }];

        slide.addText(bulletPoints, {
          x: 0.8, y: 1.6, w: '88%', h: 4.8, lineSpacing: 28, bullet: true, fontFace: 'Calibri'
        });
      });
    }

    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(title.replace(/\s+/g, '_'))}.pptx"`);
    res.send(buffer);
  } catch (err) {
    console.error('PPT Error:', err);
    res.status(500).json({ error: 'Failed to generate presentation.' });
  }
});

// -------------------------------------------------------------
// ROUTES: ADMIN KNOWLEDGE BASE & MANAGEMENT
// -------------------------------------------------------------
app.get('/api/admin/kb', adminOnly, (req, res) => {
  res.json({ knowledgeBase: db.knowledgeBase });
});

app.post('/api/admin/kb', adminOnly, (req, res) => {
  const { title, subject, category, text } = req.body;
  if (!title || !text) return res.status(400).json({ error: 'Title and text are required' });

  const newItem = {
    id: 'kb_' + Date.now(),
    title,
    subject: subject || 'General',
    category: category || 'UG',
    text
  };
  db.knowledgeBase.push(newItem);
  res.json({ success: true, item: newItem });
});

app.delete('/api/admin/kb/:id', adminOnly, (req, res) => {
  db.knowledgeBase = db.knowledgeBase.filter(k => k.id !== req.params.id);
  res.json({ success: true });
});

app.get('/api/admin/stats', adminOnly, (req, res) => {
  res.json({
    totalUsers: db.users.length,
    kbEntries: db.knowledgeBase.length,
    activeModel: process.env.DEFAULT_MODEL || 'gpt-4o',
    dailyLimitPerUser: process.env.DAILY_REQUEST_LIMIT_PER_USER || 100
  });
});

// Start Server
app.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(`🌍 R.M. Vision Academy Geography AI Server Running`);
  console.log(`Port: http://localhost:${PORT}`);
  console.log(`Default Model: ${process.env.DEFAULT_MODEL || 'gpt-4o'}`);
  console.log(`====================================================`);
});