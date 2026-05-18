// make_docx_v2.js — genera mir100_documentazione_v2.docx
// Stesso testo di make_docx.js, palette verde scuro + heading Cambria
"use strict";
const { execFileSync } = require("child_process");
const fs = require("fs");

let code = fs.readFileSync("make_docx.js", "utf8");

// ── Sostituzioni ─────────────────────────────────────────────────────────────
const rep = [

  // Costanti palette: blu → verde scuro
  ['const BLUE1  = "1D4ED8"',  'const BLUE1  = "14532D"'],   // verde foresta H1
  ['const BLUE2  = "2563EB"',  'const BLUE2  = "166534"'],   // verde medio H2
  ['const BLUE3  = "3B82F6"',  'const BLUE3  = "16A34A"'],   // verde accento H3
  ['const DARK   = "1E293B"',  'const DARK   = "1C1917"'],   // quasi-nero caldo
  ['const MUTED  = "64748B"',  'const MUTED  = "78716C"'],   // grigio caldo
  ['const CODE_BG = "F1F5F9"', 'const CODE_BG = "F0FDF4"'],  // verde chiarissimo
  ['const TH_BG   = "DBEAFE"', 'const TH_BG   = "DCFCE7"'],  // intestazione tabella verde

  // Colori hardcoded nelle funzioni helper
  ['color: "DBEAFE"',  'color: "BBF7D0"'],  // divider, bordi header/footer
  ['color: "93C5FD"',  'color: "86EFAC"'],  // infoBox e bordi tabella
  ['fill: "EFF6FF"',   'fill: "F0FDF4"'],   // infoBox background
  ['color: "1E293B"',  'color: "1C1917"'],  // testo nei code block

  // Font heading H1: Calibri → Cambria, size 32 → 34
  ['bold: true, size: 32, color: BLUE1, font: "Calibri"',
   'bold: true, size: 34, color: BLUE1, font: "Cambria"'],

  // Font heading H2: Calibri → Cambria
  ['bold: true, size: 26, color: BLUE2, font: "Calibri"',
   'bold: true, size: 26, color: BLUE2, font: "Cambria"'],

  // Stili paragrafo (usati dalla TOC e per il TOC)
  ['size: 32, bold: true, font: "Calibri", color: BLUE1',
   'size: 34, bold: true, font: "Cambria", color: BLUE1'],
  ['size: 26, bold: true, font: "Calibri", color: BLUE2',
   'size: 26, bold: true, font: "Cambria", color: BLUE2'],

  // File di output
  ['"mir100_documentazione.docx"',       '"mir100_documentazione_v2.docx"'],
  ['mir100_documentazione.docx generato', 'mir100_documentazione_v2.docx generato'],
];

for (const [from, to] of rep) {
  code = code.split(from).join(to);
}

const tmp = "_tmp_docx_v2.js";
fs.writeFileSync(tmp, code);
try {
  execFileSync("node", [tmp], { stdio: "inherit" });
} finally {
  try { fs.unlinkSync(tmp); } catch {}
}
