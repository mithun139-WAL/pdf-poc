/**
 * Step 1 — Extract Qx and Rx blocks from PDF → Excel (Enhanced)
 * Usage:
 *   npm init -y
 *   npm i pdf-parse xlsx
 *   node extract_to_excel.js --in "./input.pdf" --out "./questions.xlsx"
 *
 * Output: questions.xlsx and questions.json
 * Columns: Label | Type | Text | Answer
 */

const fs = require("fs");
const path = require("path");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx");

function getArg(flag, fallback = undefined) {
  const idx = process.argv.indexOf(flag);
  if (idx !== -1 && idx + 1 < process.argv.length) return process.argv[idx + 1];
  return fallback;
}

function makeTimestampedName(base) {
  const ts = new Date().toISOString().replace(/[-:]/g, "").replace("T", "_").slice(0, 15);
  const ext = path.extname(base) || ".xlsx";
  const name = path.basename(base, ext);
  return `${name}_${ts}${ext}`;
}

const inputPath = getArg("--in") || getArg("-i");
let outputPath = getArg("--out") || getArg("-o") || "./questions.xlsx";
outputPath = makeTimestampedName(outputPath);

if (!inputPath) {
  console.error("❌ Please provide an input PDF: --in ./file.pdf");
  process.exit(1);
}
if (!fs.existsSync(inputPath)) {
  console.error(`❌ File not found: ${inputPath}`);
  process.exit(1);
}

(async function main() {
  try {
    console.log("📄 Reading PDF...", path.resolve(inputPath));
    const buffer = fs.readFileSync(inputPath);
    const data = await pdfParse(buffer);
    let rawText = data.text || "";

    // Normalize text
    const normalizeText = (t) => {
      return (
        t
          .replace(/\r\n/g, "\n")
          .replace(/\r/g, "\n")
          .replace(/-\n(?=\w)/g, "")
          .replace(/\t/g, " ")
          .replace(/\u00A0/g, " ")
          .replace(/[ ]{2,}/g, " ")
      );
    };
    const text = normalizeText(rawText);

    // Find all Q/R labels with better pattern matching
    const labelFinder = /\b([QR])(\d{1,3})\b/g;
    const matches = [];
    let m;
    while ((m = labelFinder.exec(text)) !== null) {
      matches.push({ 
        label: m[1] + m[2], 
        type: m[1],
        index: m.index 
      });
    }

    if (matches.length === 0) {
      console.warn("⚠️ No Q/R labels found. Check the PDF structure.");
    }

    const rows = [];
    const seen = new Set();

    for (let i = 0; i < matches.length; i++) {
      const current = matches[i];
      const next = matches[i + 1];
      const start = current.index + current.label.length;
      const end = next ? next.index : text.length;
      let block = text.slice(start, end).trim();

      // Clean up initial formatting
      block = block.replace(/^[:\.\-\s\n]+/, "");

      // Split at Answer/Compliancy markers
      const splitRegex = /\n+(?:Answer|Complianc(?:y|e)?|Compliant)\b/i;
      const parts = block.split(splitRegex);
      let coreText = parts[0].trim();

      // Clean up core text
      coreText = coreText
        .replace(/\n{3,}/g, "\n\n")
        .replace(/[ \t]*\n[ \t]*/g, " ")
        .replace(/\s+$/g, "")
        .trim();

      // Avoid duplicates (e.g., from TOC)
      if (seen.has(current.label)) continue;
      seen.add(current.label);

      rows.push({ 
        Label: current.label, 
        Type: current.type, 
        Text: coreText, 
        Answer: "" 
      });
    }

    if (rows.length === 0) {
      console.warn("⚠️ No extracted rows after filtering.");
    } else {
      console.log(`✅ Extracted ${rows.length} item(s) (${rows.filter(r => r.Type === 'Q').length} Questions, ${rows.filter(r => r.Type === 'R').length} Requirements).`);
    }

    // Write Excel
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, { header: ["Label", "Type", "Text", "Answer"] });
    
    // Set column widths for better readability
    ws['!cols'] = [
      { wch: 10 },  // Label
      { wch: 8 },   // Type
      { wch: 80 },  // Text
      { wch: 80 }   // Answer
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, "Items");
    XLSX.writeFile(wb, outputPath);
    console.log("💾 Saved:", path.resolve(outputPath));

    // Also save as JSON
    const jsonPath = outputPath.replace(/\.xlsx?$/i, ".json");
    fs.writeFileSync(jsonPath, JSON.stringify(rows, null, 2));
    console.log("💾 Saved:", path.resolve(jsonPath));

    // Preview
    console.log("\n— Preview (first 8 items) —");
    rows.slice(0, 8).forEach((r, i) => {
      const preview = r.Text.replace(/\n/g, " ").slice(0, 120);
      console.log(`${i + 1}. ${r.Label} (${r.Type}) — ${preview}${r.Text.length > 120 ? "..." : ""}`);
    });

  } catch (err) {
    console.error("❌ Error:", err);
    process.exit(1);
  }
})();