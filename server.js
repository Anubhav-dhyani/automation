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
        scholarship: findColumn(columns, ['scholarship', 'scholarship type']) || ''
      }
    });
  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

// ─── Generate PDFs with custom column mapping ─────────────────────────
app.post('/generate-mapped', upload.fields([
  { name: 'template', maxCount: 1 },
  { name: 'excel', maxCount: 1 }
]), async (req, res) => {
  let templateFile, excelFile;

  try {
    templateFile = req.files['template']?.[0];
    excelFile = req.files['excel']?.[0];
    if (!templateFile || !excelFile) return res.status(400).json({ error: 'Both files are required.' });

    const { nameCol, courseCol, scoreCol, scholarshipCol, slab } = req.body;

    if (!nameCol) return res.status(400).json({ error: 'Name column is required.' });

    // Slab is required for filtering; and courseCol is required to apply slab rules
    const slabValue = String(slab || '').trim(); // expected: "50k" or "25k"
    if (!slabValue || !['50k', '25k'].includes(slabValue)) {
      return res.status(400).json({ error: 'Please select scholarship slab (50k/25k).' });
    }
    if (!courseCol) {
      return res.status(400).json({ error: 'Course column is required when using 50k/25k filtering.' });
    }

    const workbook = XLSX.readFile(excelFile.path);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);
    const templateBytes = fs.readFileSync(templateFile.path);

    // --- filter rows based on slab + course ---
    const isBtechOrMca = (courseRaw) => {
      const c = String(courseRaw || '').trim().toLowerCase();

      // matches: "B.Tech", "BTECH", "B Tech", "B. Tech CSE", etc.
      const isBtech = /^b\s*\.?\s*tech\b/.test(c);
      // matches: "MCA", "MCA (AI)", etc.
      const isMca = /^mca\b/.test(c);

      return isBtech || isMca;
    };

    const rowsToProcess = data.filter((row) => {
      const courseVal = row[courseCol];
      const match = isBtechOrMca(courseVal);
      return slabValue === '50k' ? match : !match;
    });

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="generated_pdfs.zip"');
    res.setHeader('X-Generated-Count', String(rowsToProcess.length));
    res.setHeader('X-Filter-Slab', slabValue);

    const archive = archiver('zip', { zlib: { level: 5 } });
    archive.on('error', (err) => {
      console.error('Archive error:', err);
      if (!res.headersSent) res.status(500).json({ error: err.message });
    });
    archive.pipe(res);

    for (let i = 0; i < rowsToProcess.length; i++) {
      const row = rowsToProcess[i];

      const name = String(row[nameCol] || '').trim();
      const course = courseCol ? String(row[courseCol] || '').trim() : '';
      const score = scoreCol ? String(row[scoreCol] || '').trim() : '';
      const scholarship = scholarshipCol ? String(row[scholarshipCol] || '').trim() : '';

      if (!name) continue;
      console.log(`[${i + 1}/${rowsToProcess.length}] ${name} (${slabValue})`);

      const pdfBytes = await generatePersonalizedPdf(templateBytes, { name, course, score, scholarship });
      const safeName = name.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_');
      archive.append(Buffer.from(pdfBytes), { name: `${safeName}.pdf` });
    }

    await archive.finalize();
  } catch (error) {
    console.error('Error:', error);
    if (!res.headersSent) res.status(500).json({ error: error.message });
  } finally {
    // ensure temp files removed even on error
    try { if (templateFile?.path) fs.unlinkSync(templateFile.path); } catch {}
    try { if (excelFile?.path) fs.unlinkSync(excelFile.path); } catch {}
  }
});

// ═══════════════════════════════════════════════════════════════════════
//  PDF GENERATION — white-out + rewrite approach
// ═══════════════════════════════════════════════════════════════════════
//
//  Template page: 1080 × 1445 pts.  PDF origin = bottom-left.
//
//  EXACT original text positions (from pdf-parse extraction):
//  ───────────────────────────────────────────────────────────
//  Y=1178  "Dear Abhinav Shukla,"
//  Y=1136  "Congratulations on qualifying for the Graphic Era Common Entrance Test (GECET) 2026."
//  Y=1094  "Based on your performance, we are pleased to offer you provisional admission in M.Tech CSE at"
//  Y=1070  "Graphic Era (Deemed to be University), Dehradun, for the Academic Session 2026–28."
//  Y=1028  "As a student, you will gain access to experienced faculty, industry-oriented learning, research"
//  Y=1004  "exposure, and structured career support within a performance-driven academic environment."
//  Y= 962  "GECET Score: 67"
//  Y= 934  "Scholarship: One Time 10%"
//
//  Line spacing within a paragraph ≈ 24 pts
//  Paragraph gap ≈ 42 pts
// ═══════════════════════════════════════════════════════════════════════

async function generatePersonalizedPdf(templateBytes, { name, course, score, scholarship }) {
  const pdfDoc = await PDFDocument.load(templateBytes);
  const page = pdfDoc.getPages()[0];

  const fontBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const fontRegular = await pdfDoc.embedFont(StandardFonts.Helvetica);

  const white = rgb(1, 1, 1);
  const textColor = rgb(0.13, 0.13, 0.13);
  const fontSize = 20;
  const lineHeight = 24;
  const paraGap = 18;    // extra gap between paragraphs (on top of lineHeight)
  const maxW = 910;       // usable text width (x: 90 → ~1000)
  const leftX = 90;

  // ─── 1. NAME ─────────────────────────────────────────────────────────
  // Cover "Dear Abhinav Shukla,"
  page.drawRectangle({ x: 88, y: 1173, width: 700, height: 30, color: white });
  page.drawText(`Dear ${name},`, {
    x: leftX, y: 1178, size: fontSize, font: fontRegular, color: textColor
  });

  // ─── 2. COVER ENTIRE BLOCK from "Based on..." through "Scholarship:" ──
  // Y range: 1094 (top of "Based on...") + some above → down past Scholarship
  // "Complete Your Admission Process" is at Y=814, so we're safe above that
  page.drawRectangle({ x: 88, y: 860, width: 920, height: 250, color: white });

  // Now re-draw all text from Y=1094 downward, adjusting for wrapping
  let cursorY = 1094;

  // ── Line: "Based on your performance... <COURSE> at" ────────────────
  const segments = buildCourseBlock(course, fontSize, fontRegular, fontBold, maxW, leftX);
  for (const seg of segments) {
    for (const part of seg) {
      page.drawText(part.text, {
        x: part.x, y: cursorY, size: fontSize, font: part.font, color: textColor
      });
    }
    cursorY -= lineHeight;
  }

  // ── Line: "Graphic Era (Deemed to be University)..." ────────────────
  // Only draw this if it wasn't already included in the last segment
  const line2Text = 'Graphic Era (Deemed to be University), Dehradun, for the Academic Session 2026\u201328.';
  if (!segments.needsSeparateLine2) {
    // It was already drawn inline — skip
  } else {
    page.drawText(line2Text, {
      x: leftX, y: cursorY, size: fontSize, font: fontRegular, color: textColor
    });
    cursorY -= lineHeight;
  }

  // ── Paragraph: "As a student, you will gain access..." ──────────────
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

  // ── GECET Score ─────────────────────────────────────────────────────
  cursorY -= paraGap;
  page.drawText(`GECET Score: ${score}`, {
    x: leftX, y: cursorY, size: fontSize, font: fontBold, color: textColor
  });
  cursorY -= lineHeight;

  // ── Scholarship ─────────────────────────────────────────────────────
  page.drawText(`Scholarship: ${scholarship}`, {
    x: leftX, y: cursorY, size: fontSize, font: fontBold, color: textColor
  });

  return await pdfDoc.save();
}

/**
 * Build the "Based on your performance... <COURSE> at" block.
 * Returns an array of line-segments (each line = array of {text, x, font}).
 * Also returns .needsSeparateLine2 to indicate if "Graphic Era..." needs its own line.
 */
function buildCourseBlock(course, fontSize, fontRegular, fontBold, maxW, leftX) {
  const prefix = 'Based on your performance, we are pleased to offer you provisional admission in ';
  const suffix = ' at';
  const line2Full = 'Graphic Era (Deemed to be University), Dehradun, for the Academic Session 2026\u201328.';

  const prefixW = fontRegular.widthOfTextAtSize(prefix, fontSize);
  const courseW = fontBold.widthOfTextAtSize(course, fontSize);
  const suffixW = fontRegular.widthOfTextAtSize(suffix, fontSize);

  const lines = [];

  // Case 1: Everything fits on one line
  if (prefixW + courseW + suffixW <= maxW) {
    lines.push([
      { text: prefix, x: leftX, font: fontRegular },
      { text: course, x: leftX + prefixW, font: fontBold },
      { text: suffix, x: leftX + prefixW + courseW, font: fontRegular },
    ]);
    // "Graphic Era..." on its own line
    lines.needsSeparateLine2 = true;
    return lines;
  }

  // Case 2: Course wraps — prefix on line 1, course continues on line 2
  // Line 1: prefix + as much of course as fits
  const remainingL1 = maxW - prefixW;
  const courseWords = course.split(' ');
  let line1Course = '';
  let line2Course = '';

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

  // Line 1
  const l1Parts = [{ text: prefix, x: leftX, font: fontRegular }];
  if (line1Course) {
    l1Parts.push({ text: line1Course, x: leftX + prefixW, font: fontBold });
  }
  lines.push(l1Parts);

  // Line 2: remaining course + " at" + possibly "Graphic Era..."
  if (line2Course) {
    const l2CourseW = fontBold.widthOfTextAtSize(line2Course, fontSize);
    const afterCourse = leftX + l2CourseW;
    const afterSuffix = afterCourse + fontRegular.widthOfTextAtSize(suffix + ' ', fontSize);
    const line2Remaining = maxW - (afterSuffix - leftX);
    const line2FullW = fontRegular.widthOfTextAtSize(line2Full, fontSize);

    if (line2FullW <= line2Remaining) {
      // "remaining course" + " at " + "Graphic Era..." all fit on line 2
      lines.push([
        { text: line2Course, x: leftX, font: fontBold },
        { text: suffix + ' ', x: afterCourse, font: fontRegular },
        { text: line2Full, x: afterSuffix, font: fontRegular },
      ]);
      lines.needsSeparateLine2 = false;
    } else {
      // Line 2: remaining course + " at"
      lines.push([
        { text: line2Course, x: leftX, font: fontBold },
        { text: suffix, x: afterCourse, font: fontRegular },
      ]);
      lines.needsSeparateLine2 = true;
    }
  } else {
    // No remaining course text — suffix goes on line 1 end (already handled above)
    lines.needsSeparateLine2 = true;
  }

  return lines;
}

/**
 * Simple word-wrap helper. Returns array of strings that fit within maxWidth.
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
