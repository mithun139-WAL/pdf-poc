const fs = require("fs");
const path = require("path");
const { PDFDocument, rgb, StandardFonts } = require("pdf-lib");
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

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
  console.error("❌ Missing or invalid PDF path (use --pdf)");
  process.exit(1);
}
if (!inputPath || !fs.existsSync(inputPath)) {
  console.error("❌ Missing or invalid Excel path (use --excel)");
  process.exit(1);
}

// ---------- Text Cleaning ----------
function cleanAnswerText(text) {
  if (!text) return "";
  
  return text
    .toString()
    .replace(/\r\n/g, ' ')
    .replace(/\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// ---------- Excel Processing with Images ----------
async function loadExcelWithImages(inputPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(inputPath);
  const worksheet = workbook.worksheets[0];
  
  const items = [];
  const imageMap = new Map();

  try {
    const images = worksheet.getImages();
    for (const img of images) {
      try {
        const imageId = img.imageId;
        const image = workbook.getImage(imageId);
        const range = img.range;
        
        const rowIndex = range.tl.nativeRow || range.tl.row;
        
        if (!imageMap.has(rowIndex)) {
          imageMap.set(rowIndex, []);
        }
        
        imageMap.get(rowIndex).push({
          buffer: image.buffer,
          extension: image.extension,
          range: range
        });
      } catch (err) {
        console.warn("⚠️  Could not extract image:", err.message);
      }
    }
  } catch (err) {
    console.warn("⚠️  No images found in Excel");
  }

  const wb = XLSX.readFile(inputPath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  const textData = XLSX.utils.sheet_to_json(ws);

  textData.forEach((row, index) => {
    const item = {
      Label: (row.Label || "").toString().trim(),
      Type: (row.Type || "").toString().trim(),
      Text: (row.Text || "").toString().trim(),
      Answer: cleanAnswerText(row.Answer),
      images: imageMap.get(index + 1) || []
    };
    items.push(item);
  });

  return items;
}

// ---------- PDF Text Extraction ----------
async function extractTextPositions(pdfPath) {
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const pdfDoc = await pdfjsLib.getDocument({ data }).promise;
  const pages = [];

  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1.0 });
    const content = await page.getTextContent();
    
    const items = content.items.map((item) => {
      const [a, b, c, d, e, f] = item.transform;
      return {
        str: (item.str || "").trim(),
        x: e,
        y: f,
        fontSize: item.height || 10,
        width: item.width || 0
      };
    });

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

// ---------- Text Wrapping ----------
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

// ---------- FIXED: Find Answer Position ----------
function findTargetPosition(pages, label) {
  let targets;
  if (label.startsWith("R")) {
    targets = ["compliancy", "answer"];
  } else {
    targets = ["answer"];
  }

  let bestMatch = null;
  let minDistance = Infinity;
  
  for (let p = 0; p < pages.length; p++) {
    const page = pages[p];
    let labelFound = false;
    let labelIndex = -1;
    
    for (let i = 0; i < page.items.length; i++) {
      if (page.items[i].str === label) {
        labelFound = true;
        labelIndex = i;
        
        // Strategy 1: Search forward on same page
        for (let j = i + 1; j < page.items.length; j++) {
          const word = page.items[j].str.toLowerCase();
          if (targets.some((t) => word.startsWith(t))) {
            const distance = j - i;
            
            if (distance < minDistance) {
              minDistance = distance;
              bestMatch = {
                pageIndex: p,
                x: page.items[j].x,
                y: page.items[j].y,
                width: page.items[j].width,
                fontSize: page.items[j].fontSize,
                keyword: targets.find((t) => word.startsWith(t)),
                labelPageIndex: p,
                labelY: page.items[i].y
              };
            }
            break;
          }
        }
        
        // Strategy 2: If not found on same page, search next pages
        if (!bestMatch || minDistance > 100) {
          for (let nextP = p + 1; nextP < Math.min(p + 11, pages.length); nextP++) {
            const nextPage = pages[nextP];
            const searchLimit = Math.min(150, nextPage.items.length);
            
            for (let k = 0; k < searchLimit; k++) {
              const word = nextPage.items[k].str.toLowerCase();
              if (targets.some((t) => word.startsWith(t))) {
                const pageDiff = nextP - p;
                const distance = pageDiff * 500 + k;
                
                if (distance < minDistance) {
                  minDistance = distance;
                  bestMatch = {
                    pageIndex: nextP,
                    x: nextPage.items[k].x,
                    y: nextPage.items[k].y,
                    width: nextPage.items[k].width,
                    fontSize: nextPage.items[k].fontSize,
                    keyword: targets.find((t) => word.startsWith(t)),
                    labelPageIndex: p,
                    labelY: page.items[i].y
                  };
                }
                break;
              }
            }
            
            if (bestMatch) break;
          }
        }
        
        // If we found a match for this label, stop searching
        if (bestMatch) break;
      }
    }
    
    // If we found the label and a match, stop searching other pages
    if (bestMatch) break;
  }
  
  return bestMatch;
}

// ---------- Find Box Bottom ----------
function findBoxBottom(parsedPages, pageIndex, keywordY, keywordX) {
  const items = parsedPages[pageIndex].items;
  const minBoxHeight = 35;
  const searchThreshold = 12;
  const horizontalTolerance = 150;
  
  let nextContentY = null;
  
  for (const it of items) {
    const verticallyBelow = it.y < keywordY - searchThreshold;
    const hasContent = it.str.trim().length > 0;
    
    if (verticallyBelow && hasContent) {
      const horizontallyAligned = Math.abs(it.x - keywordX) < horizontalTolerance;
      const isLeftAligned = it.x < keywordX + 50;
      
      if (horizontallyAligned || isLeftAligned) {
        if (nextContentY === null || it.y > nextContentY) {
          nextContentY = it.y;
        }
      }
    }
  }
  
  if (nextContentY !== null) {
    return Math.max(nextContentY + 12, keywordY - minBoxHeight);
  }
  
  return Math.max(keywordY - minBoxHeight, 45);
}

// ---------- Embed Images ----------
async function embedImages(pdfDoc, images) {
  const embeddedImages = [];
  
  for (const img of images) {
    try {
      let embeddedImg;
      
      if (img.extension === "png") {
        embeddedImg = await pdfDoc.embedPng(img.buffer);
      } else if (img.extension === "jpeg" || img.extension === "jpg") {
        embeddedImg = await pdfDoc.embedJpg(img.buffer);
      } else {
        console.warn(`⚠️  Unsupported image format: ${img.extension}`);
        continue;
      }
      
      embeddedImages.push(embeddedImg);
    } catch (err) {
      console.warn("⚠️  Could not embed image:", err.message);
    }
  }
  
  return embeddedImages;
}

// ---------- MORE ROBUST: Write Answer ----------
async function writeAnswerInBox(
  pages,
  parsedPages,
  pageSizes,
  helv,
  row,
  target,
  baseFontSize,
  baseLineHeight
) {
  const { pageIndex, x: kwX, y: kwY, keyword, fontSize: detectedFontSize, labelPageIndex } = target;
  const { width: pageW, height: pageH } = pageSizes[pageIndex];

  const useFontSize = detectedFontSize || baseFontSize;
  const useLineHeight = useFontSize * 1.3;
  
  const page = pages[pageIndex];
  const answerText = row.Answer;
  
  const keywordToPrint = keyword === "answer" ? "Answer" : "Compliancy";
  const keyWidth = helv.widthOfTextAtSize(keywordToPrint, useFontSize);
  
  const firstLinePadding = 4;
  const rightMargin = 40;
  const leftMargin = kwX;
  const answerStartX = kwX + keyWidth + firstLinePadding;
  const maxWidth = pageW - answerStartX - rightMargin - 45;
  
  // FIXED: Use a safer default for box bottom
  let boxBottom = findBoxBottom(parsedPages, pageIndex, kwY, kwX);
  const calculatedHeight = kwY - boxBottom;
  
  // Ensure minimum usable height
  if (calculatedHeight < 30) {
    boxBottom = kwY - 50; // Force at least 50pt height
  }
  
  const boxHeight = kwY - boxBottom;
  const verticalOffset = 1.5;

  // Clear the old text
  page.drawText(" ", {
    x: kwX,
    y: kwY + verticalOffset,
    size: useFontSize,
    font: helv,
    color: rgb(0, 0, 0),
  });
  
  const allLines = wrapText(answerText, helv, useFontSize, maxWidth);
  
  const hasImages = row.images && row.images.length > 0;
  const availableHeight = boxHeight - 10; // More padding
  const maxLinesInBox = Math.max(1, Math.floor(availableHeight / useLineHeight)); // At least 1 line
  
  const shouldContinue = allLines.length > 2 || hasImages;
  
  let linesToWrite, remainingLines;
  
  if (shouldContinue) {
    linesToWrite = allLines.slice(0, Math.min(1, maxLinesInBox));
    remainingLines = allLines.slice(linesToWrite.length);
  } else {
    linesToWrite = allLines.slice(0, Math.min(allLines.length, maxLinesInBox));
    remainingLines = [];
  }
  
  let y = kwY + verticalOffset;
  let linesWritten = 0;
  
  for (let i = 0; i < linesToWrite.length; i++) {
    const line = linesToWrite[i];
    const x = i === 0 ? answerStartX : leftMargin + 8;
    
    // Write immediately, adjust Y after
    page.drawText(line, {
      x,
      y,
      size: useFontSize,
      font: helv,
      color: rgb(0, 0, 0),
    });
    
    y -= useLineHeight;
    linesWritten++;
    
    // Check if next line would fit
    if (i < linesToWrite.length - 1 && y - useLineHeight < boxBottom) {
      // Add remaining to continuation
      for (let j = i + 1; j < linesToWrite.length; j++) {
        remainingLines.unshift(linesToWrite[j]);
      }
      break;
    }
  }

  return {
    label: row.Label,
    remainingLines: remainingLines,
    images: row.images || [],
    fontSize: useFontSize,
    lineHeight: useLineHeight,
    pageIndex: pageIndex,
    pageW: pageW,
    pageH: pageH,
    linesWritten: linesWritten,
    totalLines: allLines.length,
    labelPageIndex: labelPageIndex
  };
}


// ---------- Insert Blank Page ----------
function insertBlankAfter(pdfDoc, pageIndex, size) {
  if (typeof pdfDoc.insertPage === "function") {
    return pdfDoc.insertPage(pageIndex + 1, size);
  }
  return pdfDoc.addPage(size);
}

// ---------- Create Continuation Pages ----------
async function createContinuationPages(pdfDoc, pages, pageSizes, parsedPages, helv, continuations) {
  if (continuations.length === 0) return;
  
  const pageGroups = new Map();
  
  for (const cont of continuations) {
    if (cont.remainingLines.length === 0 && cont.images.length === 0) continue;
    
    const sourcePage = cont.pageIndex;
    if (!pageGroups.has(sourcePage)) {
      pageGroups.set(sourcePage, []);
    }
    pageGroups.get(sourcePage).push(cont);
  }
  
  const sortedPages = Array.from(pageGroups.keys()).sort((a, b) => a - b);
  
  console.log(`\n📄 Processing ${sortedPages.length} page groups for continuations\n`);
  
  let globalInsertOffset = 0;
  
  for (const sourcePage of sortedPages) {
    const pageConts = pageGroups.get(sourcePage);
    const adjustedInsertAt = sourcePage + globalInsertOffset;
    
    console.log(`  📄 Page ${sourcePage + 1}: ${pageConts.length} continuations`);
    
    let currentPage = null;
    let currentY = 0;
    let insertAt = adjustedInsertAt;
    
    for (const cont of pageConts) {
      const embeddedImages = await embedImages(pdfDoc, cont.images);
      let remainingLines = [...cont.remainingLines];
      
      while (remainingLines.length > 0 || embeddedImages.length > 0) {
        
        const needNewPage = currentPage === null || currentY < 100;
        
        if (needNewPage) {
          currentPage = insertBlankAfter(pdfDoc, insertAt, [cont.pageW, cont.pageH]);
          pages.splice(insertAt + 1, 0, currentPage);
          pageSizes.splice(insertAt + 1, 0, { width: cont.pageW, height: cont.pageH });
          parsedPages.splice(insertAt + 1, 0, {
            pageIndex: insertAt + 1,
            width: cont.pageW,
            height: cont.pageH,
            items: [],
          });
          
          currentY = cont.pageH - 70;
          insertAt += 1;
          globalInsertOffset += 1;
        }
        
        currentPage.drawText(`(Continuation of ${cont.label})`, {
          x: 50,
          y: currentY,
          size: cont.fontSize,
          font: helv,
          color: rgb(0.5, 0.5, 0.5),
        });
        
        currentY -= cont.lineHeight * 2;
        
        const pageBottomMargin = embeddedImages.length > 0 ? 320 : 50;
        const availableSpace = currentY - pageBottomMargin;
        const maxLinesOnPage = Math.floor(availableSpace / cont.lineHeight);
        const linesToWriteHere = remainingLines.splice(0, maxLinesOnPage);
        
        for (const line of linesToWriteHere) {
          currentPage.drawText(line, {
            x: 50,
            y: currentY,
            size: cont.fontSize,
            font: helv,
            color: rgb(0, 0, 0),
          });
          currentY -= cont.lineHeight;
        }
        
        if (embeddedImages.length > 0) {
          currentY -= 30;
          const img = embeddedImages.shift();
          const maxImgWidth = cont.pageW - 100;
          const maxImgHeight = 240;
          
          let imgDims = img.scale(1);
          
          if (imgDims.width > maxImgWidth || imgDims.height > maxImgHeight) {
            const scale = Math.min(maxImgWidth / imgDims.width, maxImgHeight / imgDims.height);
            imgDims = img.scale(scale);
          }
          
          const imgY = Math.max(currentY - imgDims.height, 50);
          
          currentPage.drawImage(img, {
            x: 50,
            y: imgY,
            width: imgDims.width,
            height: imgDims.height,
          });
          
          currentY = imgY - 30;
        }
        
        currentY -= cont.lineHeight;
        
        if (remainingLines.length === 0 && embeddedImages.length === 0) {
          break;
        }
      }
      
      console.log(`    ↳ ${cont.label}: Continuation added`);
    }
    
    currentPage = null;
  }
}

// ---------- Main ----------
(async function main() {
  try {
    console.log("📋 Loading Excel with images...\n");
    const items = await loadExcelWithImages(inputPath);
    console.log(`📋 Loaded ${items.length} items\n`);

    console.log("🔍 Extracting PDF text positions...\n");
    const parsedPages = await extractTextPositions(pdfPath);
    const srcBytes = fs.readFileSync(pdfPath);
    const pdfDoc = await PDFDocument.load(srcBytes);
    const helv = await pdfDoc.embedFont(StandardFonts.Helvetica);

    const baseFontSize = 8;
    const baseLineHeight = 10;

    const pages = pdfDoc.getPages();
    const pageSizes = pages.map((p) => p.getSize());

    let processedCount = 0;
    let skippedCount = 0;
    const skippedItems = [];
    const continuations = [];

    console.log("📝 Processing items...\n");
    
    for (const row of items) {
      const label = (row.Label || "").toString().trim();
      const answerText = row.Answer;
      
      if (!label || (!answerText && (!row.images || row.images.length === 0))) {
        console.log(`⏭️  ${label || 'Empty'}: Skipping (no content)`);
        skippedCount++;
        continue;
      }

      const target = findTargetPosition(parsedPages, label);
      
      if (!target) {
        console.warn(`⚠️  ${label}: Could not locate Answer/Compliancy position`);
        skippedItems.push(label);
        skippedCount++;
        continue;
      }

      // FIXED: Removed pdfDoc from parameters
      const contData = await writeAnswerInBox(
        pages,
        parsedPages,
        pageSizes,
        helv,
        row,
        target,
        baseFontSize,
        baseLineHeight
      );
      
      if (contData.remainingLines.length > 0 || contData.images.length > 0) {
        continuations.push(contData);
      }
      
      const crossPageNote = contData.labelPageIndex !== contData.pageIndex ? 
        ` [Label pg${contData.labelPageIndex+1} → Answer pg${contData.pageIndex+1}]` : '';
      const imgNote = contData.images.length > 0 ? ` +${contData.images.length} img` : '';
      const contNote = contData.remainingLines.length > 0 ? ` +cont.` : '';
      const statusNote = contData.totalLines <= 2 && contData.images.length === 0 ? ' [in-box]' : '';
      
      console.log(`✅ ${contData.label}: ${contData.linesWritten}/${contData.totalLines} lines (${contData.fontSize}pt)${statusNote}${contNote}${imgNote}${crossPageNote}`);
      
      processedCount++;
    }

    if (continuations.length > 0) {
      console.log(`\n📄 Creating ${continuations.length} continuation pages in order...\n`);
      await createContinuationPages(pdfDoc, pages, pageSizes, parsedPages, helv, continuations);
    }

    console.log("\n💾 Saving PDF...");
    const outBytes = await pdfDoc.save();
    fs.writeFileSync(outputPath, outBytes);
    
    console.log(`\n📄 Summary:`);
    console.log(`   ✅ Processed: ${processedCount} items`);
    console.log(`   ⏭️  Skipped: ${skippedCount} items`);
    
    if (skippedItems.length > 0) {
      console.log(`\n⚠️  Skipped labels: ${skippedItems.join(', ')}`);
    }
    
    console.log(`\n✨ Saved: ${path.resolve(outputPath)}`);
    
  } catch (err) {
    console.error("❌ Error:", err);
    console.error(err.stack);
    process.exit(1);
  }
})();
