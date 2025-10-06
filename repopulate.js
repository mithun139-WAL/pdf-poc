/**
 * Step 3 — No-overlap answer injection (page-safe with inserted pages)
 * --------------------------------------------------------------------
 * Installs:
 *   npm i pdf-lib pdfjs-dist@3.11.174 xlsx
 *
 * Run:
 *   node repopulate.js --pdf "./SOC_document_Sample.pdf" --in ./qa_filled.json --out ./SOC_filled.pdf
 *   node repopulate.js --pdf "./SOC_document_Sample.pdf" --in ./qa_filled.xlsx --out ./SOC_filled.pdf
 */

const fs = require("fs");
const path = require("path");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");
const XLSX = require("xlsx");

// ---------- CLI ----------
function getArg(flag, fallback = undefined) {
  const idx = process.argv.indexOf(flag);
  if (idx !== -1 && idx + 1 < process.argv.length) return process.argv[idx + 1];
  return fallback;
}
const pdfPath = getArg("--pdf");
const inputPath = getArg("--in") || "./qa_filled.json";
const outputPath = getArg("--out") || "./SOC_filled.pdf";

if (!pdfPath || !fs.existsSync(pdfPath)) {
  console.error("❌ Missing or invalid PDF path");
  process.exit(1);
}
if (!fs.existsSync(inputPath)) {
  console.error("❌ Missing input Q/A file (.json or .xlsx)");
  process.exit(1);
}

// ---------- Helpers ----------
function loadItems(inputPath) {
  if (inputPath.endsWith(".json")) {
    return JSON.parse(fs.readFileSync(inputPath, "utf8"));
  } else if (inputPath.endsWith(".xlsx")) {
    const wb = XLSX.readFile(inputPath);
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws);
  } else {
    console.error("❌ Unsupported input format (use .json or .xlsx)");
    process.exit(1);
  }
}

function wrapText(text, font, size, maxWidth) {
  const words = (text || "").toString().split(/\s+/);
  let line = "";
  const lines = [];
  for (const word of words) {
    const testLine = line ? line + " " + word : word;
    const width = font.widthOfTextAtSize(testLine, size);
    if (width > maxWidth && line) {
      lines.push(line);
      line = word;
    } else {
      line = testLine;
    }
  }
  if (line) lines.push(line);
  return lines;
}

// Extract all text items (string + position) per page
async function extractTextPositions(pdfPath) {
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const pdfDoc = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1.0 });
    const content = await page.getTextContent();
    // Normalize items (pdf.js coords origin is bottom-left; we’ll use f as y)
    const items = content.items.map((item) => {
      const [a, b, c, d, e, f] = item.transform;
      return {
        str: (item.str || "").trim(),
        x: e,
        y: f, // baseline Y in PDF user space
        fontSize: item.height || 10,
      };
    });

    // Sort items visually: descending y, then ascending x (top-to-bottom, left-to-right)
    items.sort((p, q) => q.y - p.y || p.x - q.x);
    pages.push({
      pageIndex: pageNum - 1,
      width: viewport.width,
      height: viewport.height,
      items,
    });
  }

  return pages;
}

// Find the "target label" position for a given Q/R label.
// Q* → look for "Answer"
// R* → look for "Compliancy" (preferred) but also accept "Answer".
function findTargetPosition(pages, label) {
  let targets;
  if (label.startsWith("R")) {
    targets = ["compliancy", "answer"]; // ✅ handle both
  } else {
    targets = ["answer"];
  }

  for (let p = 0; p < pages.length; p++) {
    const page = pages[p];
    const idxLabel = page.items.findIndex((it) => it.str === label);
    if (idxLabel >= 0) {
      // Same page search from label onward
      for (let j = idxLabel; j < page.items.length; j++) {
        const word = page.items[j].str.toLowerCase();
        if (targets.some((t) => word.startsWith(t))) {
          return {
            pageIndex: p,
            x: page.items[j].x,
            y: page.items[j].y,
            keyword: targets.find((t) => word.startsWith(t)), // "answer" or "compliancy"
          };
        }
      }
      // Not found on same page → check next page for the target
      if (p + 1 < pages.length) {
        const next = pages[p + 1];
        for (let k = 0; k < next.items.length; k++) {
          const word = next.items[k].str.toLowerCase();
          if (targets.some((t) => word.startsWith(t))) {
            return {
              pageIndex: p + 1,
              x: next.items[k].x,
              y: next.items[k].y,
              keyword: targets.find((t) => word.startsWith(t)),
            };
          }
        }
      }
    }
  }
  return null;
}

// Compute Y of the next piece of content below (on same page), to cap how much can fit
function findNextBelowY(pages, pageIndex, yTop) {
  const items = pages[pageIndex].items;
  let nextBelow = null;
  for (const it of items) {
    // pdf.js y grows upwards; "below" means strictly less than our baseline
    if (it.y < yTop - 1) {
      // small epsilon
      if (nextBelow === null || it.y > nextBelow) {
        nextBelow = it.y;
      }
    }
  }
  return nextBelow; // could be null → nothing below
}

// Insert a blank page right *after* pageIndex
function insertBlankAfter(pdfDoc, pageIndex, size) {
  if (typeof pdfDoc.insertPage === "function") {
    return pdfDoc.insertPage(pageIndex + 1, size);
  }
  // Fallback: add at end (older pdf-lib)
  const newPage = pdfDoc.addPage(size);
  return newPage;
}

function writeLinesOnPage(
  page,
  lines,
  kwX,
  startY,
  answerXFirst,
  helv,
  baseFontSize,
  lineHeight
) {
  let y = startY;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const x = i === 0 ? answerXFirst : kwX;
    page.drawText(line, {
      x,
      y,
      size: baseFontSize,
      font: helv,
      color: rgb(0, 0, 0),
    });
    y -= lineHeight;
  }
}

// ---------- New: writeAnswerWithFlow ----------
function writeAnswerWithFlow(
  pdfDoc,
  parsedPages,
  pages,
  pageSizes,
  helv,
  row,
  target,
  baseFontSize,
  lineHeight
) {
  const { pageIndex, x: kwX, y: kwY, keyword } = target;
  const { width: pageW, height: pageH } = pageSizes[pageIndex];

  const keyPrinted = keyword === "answer" ? "Answer" : "Compliancy";
  const keyWidth = helv.widthOfTextAtSize(keyPrinted, baseFontSize);
  const firstLinePadding = 10;
  const rightMargin = 40;
  const bottomMargin = 40;

  const answerXFirst = kwX + keyWidth + firstLinePadding;
  const maxWidth = pageW - answerXFirst - rightMargin;

  const allLines = wrapText(
    (row.Answer || "").toString().trim(),
    helv,
    baseFontSize,
    maxWidth
  );

  let yCursor = kwY;
  const page = pages[pageIndex];
  const parsedPage = parsedPages[pageIndex];

  // --- 1. Calculate how much space we need
  const neededHeight = allLines.length * lineHeight;

  // --- 2. Find what’s below
  const itemsBelow = parsedPage.items.filter((it) => it.y < kwY);
  const lowestAllowed = bottomMargin;

  const available = yCursor - lowestAllowed;

  if (neededHeight <= available) {
    // ✅ Everything fits, just draw
    writeLinesOnPage(
      page,
      allLines,
      kwX,
      kwY,
      answerXFirst,
      helv,
      baseFontSize,
      lineHeight
    );
    return;
  }

  // --- 3. Not enough space → shift items below
  const extraNeeded = neededHeight - available;
  let shifted = false;

  if (itemsBelow.length > 0) {
    for (const it of itemsBelow) {
      it.y -= extraNeeded; // shift down
    }
    shifted = true;
  }

  // Redraw answer after shifting
  writeLinesOnPage(
    page,
    allLines,
    kwX,
    kwY,
    answerXFirst,
    helv,
    baseFontSize,
    lineHeight
  );

  // --- 4. Handle overflow (if any items fell below bottom margin)
  const overflow = itemsBelow.filter((it) => it.y < bottomMargin);
  if (overflow.length > 0) {
    // Create continuation page
    const newPage = insertBlankAfter(pdfDoc, pageIndex, [pageW, pageH]);
    pages.splice(pageIndex + 1, 0, newPage);
    pageSizes.splice(pageIndex + 1, 0, { width: pageW, height: pageH });
    parsedPages.splice(pageIndex + 1, 0, {
      pageIndex: pageIndex + 1,
      width: pageW,
      height: pageH,
      items: [],
    });

    let y = pageH - 80;
    newPage.drawText(`(continued content)`, {
      x: 50,
      y,
      size: baseFontSize,
      font: helv,
      color: rgb(0.4, 0.4, 0.4),
    });
    y -= lineHeight * 2;

    for (const it of overflow) {
      newPage.drawText(it.str, {
        x: it.x,
        y,
        size: it.fontSize || baseFontSize,
        font: helv,
        color: rgb(0, 0, 0),
      });
      y -= lineHeight;
    }
  }

  if (!shifted) {
    // --- 5. As fallback, if nothing could shift → new page for the answer
    const newPage = insertBlankAfter(pdfDoc, pageIndex, [pageW, pageH]);
    pages.splice(pageIndex + 1, 0, newPage);
    pageSizes.splice(pageIndex + 1, 0, { width: pageW, height: pageH });
    parsedPages.splice(pageIndex + 1, 0, {
      pageIndex: pageIndex + 1,
      width: pageW,
      height: pageH,
      items: [],
    });

    let y = pageH - 80;
    newPage.drawText(`(continued: ${row.Label})`, {
      x: 50,
      y,
      size: baseFontSize,
      font: helv,
      color: rgb(0.4, 0.4, 0.4),
    });
    y -= lineHeight * 2;

    for (const line of allLines) {
      newPage.drawText(line, {
        x: 50,
        y,
        size: baseFontSize,
        font: helv,
        color: rgb(0, 0, 0),
      });
      y -= lineHeight;
    }
  }
}

// ---------- Main ----------
(async function main() {
  try {
    const items = loadItems(inputPath);
    console.log(`📋 Loaded ${items.length} items`);

    const parsedPages = await extractTextPositions(pdfPath);

    const srcBytes = fs.readFileSync(pdfPath);
    const pdfDoc = await PDFDocument.load(srcBytes);
    const helv = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const baseFontSize = 8;
    const lineHeight = 10;
    const rightMargin = 40;
    const bottomMargin = 40;
    const firstLinePadding = 10;

    const pages = pdfDoc.getPages();
    const pageSizes = pages.map((p) => p.getSize());

    for (const row of items) {
      const label = (row.Label || "").toString().trim();
      const answerText = (row.Answer || "").toString().trim();
      if (!label || !answerText) continue;

      const target = findTargetPosition(parsedPages, label);
      if (!target) {
        console.warn(`⚠️ Could not map ${label} to Answer/Compliancy`);
        continue;
      }

      const { pageIndex, x: kwX, y: kwY, keyword } = target;
      const { width: pageW, height: pageH } = pageSizes[pageIndex];

      const keyPrinted = keyword === "answer" ? "Answer" : "Compliancy";
      const keyWidth = helv.widthOfTextAtSize(keyPrinted, baseFontSize);
      const answerXFirst = kwX + keyWidth + firstLinePadding;
      const maxWidth = pageW - answerXFirst - rightMargin;

      const allLines = wrapText(answerText, helv, baseFontSize, maxWidth);

      const nextBelowY = findNextBelowY(parsedPages, pageIndex, kwY);
      const lowestAllowed =
        nextBelowY !== null
          ? Math.max(nextBelowY + lineHeight * 0.5, bottomMargin)
          : bottomMargin;

      let fitCount = 0;
      let yCursor = kwY;
      while (fitCount < allLines.length) {
        if (yCursor < lowestAllowed) break;
        fitCount++;
        yCursor -= lineHeight;
      }

      let remaining = [];

      writeAnswerWithFlow(
        pdfDoc,
        parsedPages,
        pages,
        pageSizes,
        helv,
        row,
        target,
        baseFontSize,
        lineHeight
      );

      // 🔄 Spill remaining into continuation pages
      let insertAt = pageIndex;
      while (remaining.length > 0) {
        const newPage = insertBlankAfter(pdfDoc, insertAt, [pageW, pageH]);
        pages.splice(insertAt + 1, 0, newPage);
        pageSizes.splice(insertAt + 1, 0, { width: pageW, height: pageH });
        parsedPages.splice(insertAt + 1, 0, {
          pageIndex: insertAt + 1,
          width: pageW,
          height: pageH,
          items: [],
        });

        let y = pageH - 80;
        newPage.drawText(`(continued: ${label})`, {
          x: 50,
          y,
          size: baseFontSize,
          font: helv,
          color: rgb(0.4, 0.4, 0.4),
        });
        y -= lineHeight * 2;

        while (remaining.length > 0 && y > bottomMargin) {
          const line = remaining.shift();
          newPage.drawText(line, {
            x: 50,
            y,
            size: baseFontSize,
            font: helv,
            color: rgb(0, 0, 0),
          });
          y -= lineHeight;
        }

        insertAt += 1;
      }
    }

    const outBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, outBytes);
    console.log("✅ Saved:", path.resolve(outputPath));
  } catch (err) {
    console.error("❌ Error:", err);
    process.exit(1);
  }
})();
