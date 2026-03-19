const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const { PDFDocument, rgb, StandardFonts } = require('pdf-lib');
const archiver = require('archiver');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.json());

// ─── Preview Excel columns ────────────────────────────────────────────
app.post('/preview-columns', upload.fields([
  { name: 'excel', maxCount: 1 }
]), async (req, res) => {
  try {
    const excelFile = req.files['excel']?.[0];
    if (!excelFile) return res.status(400).json({ error: 'Excel file is required.' });

    const workbook = XLSX.readFile(excelFile.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);
    fs.unlinkSync(excelFile.path);

    if (data.length === 0) return res.status(400).json({ error: 'Excel file is empty.' });

    const columns = Object.keys(data[0]);

    res.json({
      columns,
      rowCount: data.length,
      sampleRows: data.slice(0, 3),
      autoDetected: {
        name: findColumn(columns, ['registered name', 'name', 'student name', 'full name', 'candidate name']) || '',
        course: findColumn(columns, ['course', 'program', 'programme', 'department']) || '',
        score: findColumn(columns, ['result', 'score', 'gecet score', 'marks', 'gecet']) || '',
        scholarship: findColumn(columns, ['scholarship', 'scholarship type']) || '',
        academicSession: findColumn(columns, ['academic session', 'session', 'academic_session', 'academ session']) || ''
      }
    });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

// ─────────────────────────────────────────────────────────────────────
// Generic preview (works for XLSX/XLS/CSV).
// ─────────────────────────────────────────────────────────────────────
app.post('/preview-columns-generic', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    if (!file) return res.status(400).json({ error: 'File is required.' });

    const workbook = XLSX.readFile(file.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    fs.unlinkSync(file.path);

    if (!data.length) return res.status(400).json({ error: 'File is empty.' });

    const columns = Object.keys(data[0]);

    res.json({
      columns,
      rowCount: data.length,
      sampleRows: data.slice(0, 3),
      autoDetected: {
        course: findColumn(columns, ['course', 'program', 'programme', 'department']) || '',
        academicSession: findColumn(columns, ['academic session', 'session', 'academic_session']) || '',
      }
    });
  } catch (error) {
    console.error('preview-columns-generic error:', error);
    res.status(500).json({ error: error.message });
  }
});

function normalizeCourseKey(v) {
  return String(v || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// ─────────────────────────────────────────────────────────────────────
// Enrich student Excel using course->academic session mapping
// ─────────────────────────────────────────────────────────────────────
app.post('/enrich-academic-session', upload.fields([
  { name: 'students', maxCount: 1 },
  { name: 'mapping', maxCount: 1 },
]), async (req, res) => {
  let studentsFile, mappingFile;

  try {
    studentsFile = req.files['students']?.[0];
    mappingFile  = req.files['mapping']?.[0];
    if (!studentsFile || !mappingFile) {
      return res.status(400).json({ error: 'Both Student file and Course Mapping file are required.' });
    }

    const { studentsCourseCol, mappingCourseCol, mappingSessionCol } = req.body;
    if (!studentsCourseCol) return res.status(400).json({ error: 'Student Course column is required.' });
    if (!mappingCourseCol)  return res.status(400).json({ error: 'Mapping Course column is required.' });
    if (!mappingSessionCol) return res.status(400).json({ error: 'Mapping Academic Session column is required.' });

    const studentsWb    = XLSX.readFile(studentsFile.path);
    const studentsSheet = studentsWb.Sheets[studentsWb.SheetNames[0]];
    const studentsRows  = XLSX.utils.sheet_to_json(studentsSheet, { defval: '' });
    if (!studentsRows.length) return res.status(400).json({ error: 'Student file is empty.' });

    const mappingWb    = XLSX.readFile(mappingFile.path);
    const mappingSheet = mappingWb.Sheets[mappingWb.SheetNames[0]];
    const mappingRows  = XLSX.utils.sheet_to_json(mappingSheet, { defval: '' });
    if (!mappingRows.length) return res.status(400).json({ error: 'Course Mapping file is empty.' });

    const courseToSession = new Map();
    for (const r of mappingRows) {
      const key = normalizeCourseKey(r[mappingCourseCol]);
      const sessionVal = String(r[mappingSessionCol] || '').trim();
      if (!key || !sessionVal) continue;
      if (!courseToSession.has(key)) courseToSession.set(key, sessionVal);
    }

    let matched = 0, unmatched = 0;
    const outRows = studentsRows.map((r) => {
      const key     = normalizeCourseKey(r[studentsCourseCol]);
      const session = key ? (courseToSession.get(key) || '') : '';
      const status  = session ? 'Matched' : 'Unmatched';
      if (session) matched++; else unmatched++;
      return { ...r, 'Academic Session': session, 'Academic Session Status': status };
    });

    const outWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(outWb, XLSX.utils.json_to_sheet(outRows), 'Enriched');
    const buffer = XLSX.write(outWb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="students_enriched_academic_session.xlsx"');
    res.setHeader('X-Matched', String(matched));
    res.setHeader('X-Unmatched', String(unmatched));
    res.send(buffer);
  } catch (error) {
    console.error('enrich-academic-session error:', error);
    res.status(500).json({ error: error.message });
  } finally {
    try { if (studentsFile?.path) fs.unlinkSync(studentsFile.path); } catch {}
    try { if (mappingFile?.path)  fs.unlinkSync(mappingFile.path);  } catch {}
  }
});

// ─── Generate PDFs ─────────────────────────────────────────────────────
app.post('/generate-mapped', upload.fields([
  { name: 'template', maxCount: 1 },
  { name: 'excel',    maxCount: 1 }
]), async (req, res) => {
  let templateFile, excelFile;

  try {
    templateFile = req.files['template']?.[0];
    excelFile    = req.files['excel']?.[0];
    if (!templateFile || !excelFile) return res.status(400).json({ error: 'Both files are required.' });

    const { nameCol, courseCol, scoreCol, scholarshipCol, academicSessionCol, slab } = req.body;

    if (!nameCol) return res.status(400).json({ error: 'Name column is required.' });

    const slabValue = String(slab || '').trim();
    if (!slabValue || !['50k', '25k'].includes(slabValue)) {
      return res.status(400).json({ error: 'Please select scholarship slab (50k/25k).' });
    }
    if (!courseCol) {
      return res.status(400).json({ error: 'Course column is required when using 50k/25k filtering.' });
    }

    const workbook = XLSX.readFile(excelFile.path);
    const sheet    = workbook.Sheets[workbook.SheetNames[0]];
    const data     = XLSX.utils.sheet_to_json(sheet);
    const templateBytes = fs.readFileSync(templateFile.path);

    const isBtechOrMca = (courseRaw) => {
      const c = String(courseRaw || '').trim().toLowerCase();
      return /^b\s*\.?\s*tech\b/.test(c) || /^mca\b/.test(c);
    };

    const rowsToProcess = data.filter((row) => {
      const match = isBtechOrMca(row[courseCol]);
      return slabValue === '50k' ? match : !match;
    });

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="generated_pdfs.zip"');
    res.setHeader('X-Generated-Count', String(rowsToProcess.length));
    res.setHeader('X-Filter-Slab', slabValue);

    const archive = archiver('zip', { zlib: { level: 5 } });
    archive.on('error', (err) => {
      if (!res.headersSent) res.status(500).json({ error: err.message });
    });
    archive.pipe(res);

    for (let i = 0; i < rowsToProcess.length; i++) {
      const row = rowsToProcess[i];

      const name            = String(row[nameCol]              || '').trim();
      const course          = courseCol          ? String(row[courseCol]          || '').trim() : '';
      const score           = scoreCol           ? String(row[scoreCol]           || '').trim() : '';
      const scholarship     = scholarshipCol     ? String(row[scholarshipCol]     || '').trim() : '';
      const academicSession = academicSessionCol ? String(row[academicSessionCol] || '').trim() : '';

      if (!name) continue;
      console.log(`[${i + 1}/${rowsToProcess.length}] ${name} | Session: ${academicSession || 'N/A'}`);

      const pdfBytes = await generatePersonalizedPdf(templateBytes, {
        name, course, score, scholarship, academicSession
      });
      const safeName = name.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      archive.append(Buffer.from(pdfBytes), { name: `${safeName}.pdf` });
    }

    await archive.finalize();
  } catch (error) {
    console.error('Error:', error);
    if (!res.headersSent) res.status(500).json({ error: error.message });
  } finally {
    try { if (templateFile?.path) fs.unlinkSync(templateFile.path); } catch {}
    try { if (excelFile?.path)    fs.unlinkSync(excelFile.path);    } catch {}
  }
});

// ═══════════════════════════════════════════════════════════════════════
//  PDF GENERATION — white-out + rewrite
//
//  Template page: 1080 × 1445 pts. PDF origin = BOTTOM-LEFT.
//
//  The floating stamp ("2026-2029" on the right side) is part of the
//  original template. We white it out and DO NOT redraw it.
//  The Academic Session value only appears inline in the paragraph:
//  "...for the Academic Session <VALUE>."
// ═══════════════════════════════════════════════════════════════════════

async function generatePersonalizedPdf(templateBytes, { name, course, score, scholarship, academicSession }) {
  const pdfDoc = await PDFDocument.load(templateBytes);
  const page   = pdfDoc.getPages()[0];

  const fontBold    = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);

  const white     = rgb(1, 1, 1);
  const textColor = rgb(0.13, 0.13, 0.13);
  const fontSize  = 20;
  const lineHeight = 24;
  const paraGap   = 18;
  const maxW      = 820;
  const leftX     = 90;

  // Use Excel value if present, otherwise keep template default
  const sessionDisplay = academicSession || '2026\u201328';

  // ─── 1. NAME ────────────────────────────────────────────────────────
  page.drawRectangle({ x: 88, y: 1173, width: 700, height: 30, color: white });
  page.drawText(`Dear ${name},`, {
    x: leftX, y: 1178, size: fontSize, font: fontRegular, color: textColor
  });

  // ─── 2. MAIN BODY BLOCK ─────────────────────────────────────────────
  // White-out covers the full width of the page (x: 88 → 1070) to also
  // erase the floating stamp that the template has on the right side.
  // We do NOT redraw the stamp — it simply disappears.
  page.drawRectangle({ x: 88, y: 860, width: 985, height: 260, color: white });

  let cursorY = 1094;

  // "Based on your performance... <COURSE> at"
  const segments = buildCourseBlock(course, fontSize, fontRegular, fontBold, maxW, leftX);
  for (const seg of segments) {
    for (const part of seg) {
      page.drawText(part.text, {
        x: part.x, y: cursorY, size: fontSize, font: part.font, color: textColor
      });
    }
    cursorY -= lineHeight;
  }

  // "Graphic Era... for the Academic Session <DYNAMIC VALUE>."
  const line2Text = `Graphic Era (Deemed to be University), Dehradun, for the Academic Session ${sessionDisplay}.`;
  if (segments.needsSeparateLine2) {
    page.drawText(line2Text, {
      x: leftX, y: cursorY, size: fontSize, font: fontRegular, color: textColor
    });
    cursorY -= lineHeight;
  }

  // "As a student..." paragraph
  cursorY -= paraGap;
  const para2Lines = wrapText(
    'As a student, you will gain access to experienced faculty, industry-oriented learning, research exposure, and structured career support within a performance-driven academic environment.',
    fontSize, fontRegular, maxW
  );
  for (const line of para2Lines) {
    page.drawText(line, {
      x: leftX, y: cursorY, size: fontSize, font: fontRegular, color: textColor
    });
    cursorY -= lineHeight;
  }

  // GECET Score
  cursorY -= paraGap;
  page.drawText(`GECET Score: ${score}`, {
    x: leftX, y: cursorY, size: fontSize, font: fontBold, color: textColor
  });
  cursorY -= lineHeight;

  // Scholarship
  page.drawText(`Scholarship: ${scholarship}`, {
    x: leftX, y: cursorY, size: fontSize, font: fontBold, color: textColor
  });

  // ── Stamp is intentionally NOT redrawn ──────────────────────────────
  // The white-out rectangle above already erased it from the template.

  return await pdfDoc.save();
}

/**
 * Build the "Based on your performance... <COURSE> at" block.
 * Returns array of line-segments. Sets .needsSeparateLine2.
 */
function buildCourseBlock(course, fontSize, fontRegular, fontBold, maxW, leftX) {
  const prefix = 'Based on your performance, we are pleased to offer you provisional admission in ';
  const suffix = ' at';

  const prefixW = fontRegular.widthOfTextAtSize(prefix, fontSize);
  const courseW = fontBold.widthOfTextAtSize(course, fontSize);
  const suffixW = fontRegular.widthOfTextAtSize(suffix, fontSize);

  const lines = [];

  if (prefixW + courseW + suffixW <= maxW) {
    lines.push([
      { text: prefix, x: leftX,                     font: fontRegular },
      { text: course, x: leftX + prefixW,            font: fontBold    },
      { text: suffix, x: leftX + prefixW + courseW,  font: fontRegular },
    ]);
    lines.needsSeparateLine2 = true;
    return lines;
  }

  // Course wraps to next line
  const remainingL1 = maxW - prefixW;
  const courseWords = course.split(' ');
  let line1Course = '', line2Course = '';

  for (let w = 0; w < courseWords.length; w++) {
    const attempt = courseWords.slice(0, w + 1).join(' ');
    if (fontBold.widthOfTextAtSize(attempt, fontSize) > remainingL1) {
      line1Course = courseWords.slice(0, Math.max(1, w)).join(' ');
      line2Course = courseWords.slice(Math.max(1, w)).join(' ');
      break;
    }
    line1Course = attempt;
    line2Course = '';
  }

  const l1Parts = [{ text: prefix, x: leftX, font: fontRegular }];
  if (line1Course) l1Parts.push({ text: line1Course, x: leftX + prefixW, font: fontBold });
  lines.push(l1Parts);

  if (line2Course) {
    const l2CourseW   = fontBold.widthOfTextAtSize(line2Course, fontSize);
    const afterCourse = leftX + l2CourseW;
    lines.push([
      { text: line2Course, x: leftX,       font: fontBold    },
      { text: suffix,      x: afterCourse, font: fontRegular },
    ]);
    lines.needsSeparateLine2 = true;
  } else {
    lines.needsSeparateLine2 = true;
  }

  return lines;
}

/**
 * Simple word-wrap. Returns array of strings fitting within maxWidth.
 */
function wrapText(text, fontSize, font, maxWidth) {
  const words = text.split(' ');
  const lines = [];
  let currentLine = '';
  for (const word of words) {
    const testLine = currentLine ? `${currentLine} ${word}` : word;
    if (font.widthOfTextAtSize(testLine, fontSize) <= maxWidth) {
      currentLine = testLine;
    } else {
      if (currentLine) lines.push(currentLine);
      currentLine = word;
    }
  }
  if (currentLine) lines.push(currentLine);
  return lines;
}

/**
 * Find a column name in Excel that matches one of the expected names (case-insensitive).
 */
function findColumn(columns, expectedNames) {
  for (const expected of expectedNames) {
    const found = columns.find(c => c.toLowerCase().trim() === expected.toLowerCase());
    if (found) return found;
  }
  for (const expected of expectedNames) {
    const found = columns.find(c => c.toLowerCase().trim().includes(expected.toLowerCase()));
    if (found) return found;
  }
  return null;
}

app.listen(PORT, () => {
  console.log(`\n✅  Server running at http://localhost:${PORT}\n`);
});