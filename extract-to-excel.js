/**
 * Step 1 — Extract Qx and Rx blocks from PDF → Excel
 * Usage:
 *   npm init -y
 *   npm i pdf-parse xlsx
 *   node step1-extract-to-excel.js --in "./input.pdf" --out "./questions.xlsx"
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
  const ts = new Date().toISOString().replace(/[-:]/g, "").replace("T", "_").slice(0, 13);
  const ext = path.extname(base) || ".xlsx";
  const name = path.basename(base, ext);
  return `${name}-${ts}${ext}`;
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

    // normalize
    const normalizeText = (t) => {
      return (
        t
          .replace(/\r\n/g, "\n")
          .replace(/\r/g, "\n")
          .replace(/-\n(?=\w)/g, "")   // un-hyphenate line breaks
          .replace(/\t/g, " ")
          .replace(/\u00A0/g, " ")
          .replace(/[ ]{2,}/g, " ")
      );
    };
    const text = normalizeText(rawText);

    // find all label positions for Qx and Rx
    const labelFinder = /([QR]\d{1,3})\b/g;
    const matches = [];
    let m;
    while ((m = labelFinder.exec(text)) !== null) {
      matches.push({ label: m[1], index: m.index });
    }

    if (matches.length === 0) {
      console.warn("⚠️ No Q/R labels found. Check the PDF structure or adjust the regex.");
    }

    const rows = [];
    const seen = new Set();

    for (let i = 0; i < matches.length; i++) {
      const current = matches[i];
      const next = matches[i + 1];
      const start = current.index + current.label.length; // content after label
      const end = next ? next.index : text.length;
      let block = text.slice(start, end).trim();

      // Trim initial noise
      block = block.replace(/^[:\.\-\s\n]+/, "");

      // split at a marker word if present (Answer or Compliancy/Compliant)
      const splitRegex = /\n+(?:Answer|Complianc(?:y|e)?|Compliant)\b/;
      const parts = block.split(splitRegex);
      let coreText = parts[0].trim();

      // tidy core text
      coreText = coreText
        .replace(/\n{3,}/g, "\n\n")
        .replace(/[ \t]*\n[ \t]*/g, "\n")
        .replace(/\s+$/g, "");

      // keep first occurrence only (filters TOC duplicates)
      if (seen.has(current.label)) continue;
      seen.add(current.label);

      rows.push({ Label: current.label, Type: current.label[0], Text: coreText, Answer: "" });
    }

    if (rows.length === 0) console.warn("⚠️ No extracted rows (after filtering).");
    else console.log(`✅ Extracted ${rows.length} item(s) (Q and R combined).`);

    // write Excel
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, { header: ["Label", "Type", "Text", "Answer"] });
    XLSX.utils.book_append_sheet(wb, ws, "Items");
    XLSX.writeFile(wb, outputPath);
    console.log("💾 Saved:", path.resolve(outputPath));

    // also JSON
    const jsonPath = outputPath.replace(/\.xlsx?$/i, ".json");
    fs.writeFileSync(jsonPath, JSON.stringify(rows, null, 2));
    console.log("💾 Saved:", path.resolve(jsonPath));

    // preview
    console.log("\n— Preview —");
    rows.slice(0, 8).forEach((r, i) => {
      const preview = r.Text.replace(/\n/g, " ").slice(0, 160);
      console.log(`${i + 1}. ${r.Label} (${r.Type}) — ${preview}${r.Text.length > 160 ? "..." : ""}`);
    });

  } catch (err) {
    console.error("❌ Error:", err);
    process.exit(1);
  }
})();