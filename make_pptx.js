// make_pptx.js — genera mir100_presentazione.pptx
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout  = "LAYOUT_16x9";
pres.title   = "MiR100 — Sistema di Gestione Magazzino";
pres.author  = "ITIS 2025/2026";

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const C = {
  bg:     "111318",
  bgDeep: "0A0D12",
  panel:  "1A1D24",
  card:   "1F2330",
  border: "2A2F3E",
  accent: "3B82F6",
  green:  "22C55E",
  red:    "EF4444",
  yellow: "EAB308",
  orange: "FF6B00",
  purple: "A78BFA",
  teal:   "38BDF8",
  violet: "7C3AED",
  text:   "F1F5F9",
  text2:  "94A3B8",
  muted:  "475569",
};

// ─── HELPERS ─────────────────────────────────────────────────────────────────
const darkBg = (s) => { s.background = { color: C.bg }; };

const slideTitle = (s, t, y = 0.32) => {
  s.addText(t, {
    x: 0.45, y, w: 9.1, h: 0.5,
    fontSize: 21, bold: true, color: C.accent,
    fontFace: "Trebuchet MS", align: "left", margin: 0,
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: y + 0.53, w: 9.1, h: 0.018,
    fill: { color: C.border }, line: { color: C.border },
  });
};

const slideNum = (s, n) =>
  s.addText(`${n} / 15`, {
    x: 8.8, y: 5.38, w: 0.9, h: 0.2,
    fontSize: 8, color: C.muted, align: "right",
  });

const card = (s, x, y, w, h, borderColor = C.border) => {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.card }, line: { color: borderColor, width: 1 },
  });
};

const accentBar = (s, x, y, h, color) => {
  s.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.055, h,
    fill: { color }, line: { color },
  });
};

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — TITOLO
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);

  // Top accent stripe
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.055,
    fill: { color: C.accent }, line: { color: C.accent },
  });

  // ── Right panel: robot illustration
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.3, y: 0.5, w: 3.35, h: 4.55,
    fill: { color: C.panel }, line: { color: C.border, width: 1 },
  });
  // Robot body
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.7, y: 0.85, w: 2.55, h: 1.5,
    fill: { color: C.card }, line: { color: C.accent, width: 2 },
  });
  s.addText("MiR100", {
    x: 6.7, y: 0.85, w: 2.55, h: 0.7,
    fontSize: 28, bold: true, color: C.accent,
    fontFace: "Trebuchet MS", align: "center", valign: "middle",
  });
  // Camera lens
  s.addShape(pres.shapes.OVAL, {
    x: 7.55, y: 1.65, w: 0.45, h: 0.45,
    fill: { color: C.accent }, line: { color: C.accent },
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.63, y: 1.73, w: 0.29, h: 0.29,
    fill: { color: C.bg }, line: { color: C.bg },
  });
  // Wheels
  for (let wx = 0; wx < 4; wx++) {
    s.addShape(pres.shapes.OVAL, {
      x: 6.82 + wx * 0.55, y: 2.42, w: 0.45, h: 0.22,
      fill: { color: C.muted }, line: { color: C.muted },
    });
  }
  // RPi on robot
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.75, y: 2.8, w: 1.3, h: 0.7,
    fill: { color: "1A2A1A" }, line: { color: C.green, width: 1 },
  });
  s.addText("🍓 RPi", {
    x: 6.75, y: 2.8, w: 1.3, h: 0.7,
    fontSize: 11, color: C.green, align: "center", valign: "middle",
  });
  // Webcam
  s.addShape(pres.shapes.RECTANGLE, {
    x: 8.2, y: 2.85, w: 1.1, h: 0.6,
    fill: { color: C.card }, line: { color: C.teal, width: 1 },
  });
  s.addText("📷 Cam", {
    x: 8.2, y: 2.85, w: 1.1, h: 0.6,
    fontSize: 10, color: C.teal, align: "center", valign: "middle",
  });
  s.addText("192.168.12.x  ·  WiFi", {
    x: 6.35, y: 3.65, w: 3.2, h: 0.28,
    fontSize: 9, color: C.muted, align: "center",
  });
  s.addText("REST API  ·  MJPEG Stream  ·  SSE", {
    x: 6.35, y: 3.93, w: 3.2, h: 0.28,
    fontSize: 9, color: C.muted, align: "center",
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 6.55, y: 4.28, w: 2.8, h: 0.45,
    fill: { color: C.bg }, line: { color: C.green, width: 1 },
  });
  s.addText("⚡ RPi alimentato via USB del robot", {
    x: 6.55, y: 4.28, w: 2.8, h: 0.45,
    fontSize: 9, color: C.green, align: "center", valign: "middle",
  });

  // ── Left: main title block
  s.addText("MiR100", {
    x: 0.45, y: 0.7, w: 5.55, h: 0.85,
    fontSize: 52, bold: true, color: C.accent,
    fontFace: "Trebuchet MS", align: "left", margin: 0,
  });
  s.addText("Sistema di Gestione Magazzino", {
    x: 0.45, y: 1.62, w: 5.55, h: 0.52,
    fontSize: 22, bold: true, color: C.text,
    fontFace: "Trebuchet MS", align: "left", margin: 0,
  });
  s.addText("Automazione della logistica interna\ncon robot mobile autonomo MiR100", {
    x: 0.45, y: 2.25, w: 5.55, h: 0.75,
    fontSize: 13, color: C.text2, fontFace: "Calibri",
  });

  // Tag badge
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 3.15, w: 2.5, h: 0.36,
    fill: { color: C.card }, line: { color: C.accent, width: 1 },
  });
  s.addText("Progetto ITIS  ·  2025/2026", {
    x: 0.45, y: 3.15, w: 2.5, h: 0.36,
    fontSize: 10, color: C.accent, align: "center", valign: "middle",
    fontFace: "Trebuchet MS",
  });

  // Bottom tech bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.28, w: 10, h: 0.345,
    fill: { color: C.panel }, line: { color: C.border },
  });
  s.addText(
    "Python  ·  Flask  ·  SQLite  ·  OpenCV  ·  face_recognition  ·  MiR100 REST API  ·  Raspberry Pi  ·  SSE",
    {
      x: 0.3, y: 5.29, w: 9.4, h: 0.32,
      fontSize: 9, color: C.muted, align: "center",
    }
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — OBIETTIVO
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Obiettivo del Progetto");
  slideNum(s, 2);

  s.addText(
    "Automatizzare il flusso di ricezione e distribuzione dei pacchi in un magazzino tramite il robot mobile MiR100, eliminando il trasporto manuale dei carichi.",
    {
      x: 0.45, y: 1.0, w: 4.3, h: 0.95,
      fontSize: 12, color: C.text2, fontFace: "Calibri",
    }
  );

  const bullets = [
    "Ricezione automatizzata dei carichi dai camion in arrivo",
    "Autenticazione biometrica degli operatori (face recognition)",
    "Tracciamento QR code: baia di scarico → zona destinazione",
    "Controllo MiR100 via API REST",
    "Interfaccia web real-time accessibile da tablet",
  ];
  s.addText(
    bullets.map((t, i) => ({
      text: t,
      options: { bullet: true, breakLine: i < bullets.length - 1, fontSize: 13, color: C.text, paraSpaceAfter: 7 },
    })),
    { x: 0.45, y: 2.05, w: 4.3, h: 3.0, fontFace: "Calibri", valign: "top" }
  );

  // Right: 2×3 feature cards
  const feats = [
    { e: "📦", l: "Tracciamento", s: "QR · atteso→consegnato" },
    { e: "🤖", l: "Automazione", s: "Missioni autonome" },
    { e: "👁️", l: "Biometria", s: "face_recognition · dlib" },
    { e: "📡", l: "Real-time", s: "Server-Sent Events" },
    { e: "🔒", l: "Sicurezza", s: "bcrypt · rate limiting" },
    { e: "⚙️", l: "Admin Panel", s: "config · storico · QR" },
  ];
  feats.forEach((f, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 5.05 + col * 2.42, y = 1.0 + row * 1.55;
    card(s, x, y, 2.25, 1.4);
    s.addText(f.e,  { x, y: y + 0.1,  w: 2.25, h: 0.42, fontSize: 22, align: "center" });
    s.addText(f.l,  { x, y: y + 0.57, w: 2.25, h: 0.3,  fontSize: 12, bold: true, color: C.accent, align: "center", fontFace: "Trebuchet MS" });
    s.addText(f.s,  { x, y: y + 0.9,  w: 2.25, h: 0.28, fontSize: 9.5, color: C.text2, align: "center" });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — ARCHITETTURA
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Architettura di Sistema");
  slideNum(s, 3);

  s.addText("Tutti i componenti comunicano sulla rete locale dedicata del MiR  ·  192.168.12.x", {
    x: 0.45, y: 0.97, w: 9.1, h: 0.25,
    fontSize: 10.5, color: C.muted, align: "center",
  });

  // Nodes: [label, x, y, w, h, fill, textColor]
  const nodes = [
    { l: "CAMION\nIn arrivo",            x: 0.25, y: 1.55, w: 1.55, h: 0.8, f: C.muted,   tc: C.text },
    { l: "TABLET BALIA\n/ (main)",       x: 2.2,  y: 1.55, w: 1.9,  h: 0.8, f: C.violet,  tc: "FFFFFF" },
    { l: "SERVER FLASK\n:5000",          x: 4.5,  y: 1.55, w: 1.95, h: 0.8, f: "1D4ED8",   tc: "FFFFFF" },
    { l: "MiR100\n192.168.12.20",        x: 4.5,  y: 3.1,  w: 1.95, h: 0.8, f: C.orange,  tc: "FFFFFF" },
    { l: "RASPBERRY PI\n192.168.12.242", x: 7.15, y: 1.55, w: 2.4,  h: 0.8, f: C.green,   tc: "111318" },
    { l: "WEBCAM USB\nMJPEG :8081",      x: 7.15, y: 3.1,  w: 2.4,  h: 0.8, f: C.card,    tc: C.teal, border: C.green },
    { l: "TABLET SCARICO\n/scarico",     x: 4.5,  y: 4.35, w: 1.95, h: 0.8, f: C.accent,  tc: "FFFFFF" },
  ];

  nodes.forEach((n) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: n.x, y: n.y, w: n.w, h: n.h,
      fill: { color: n.f }, line: { color: n.border || n.f, width: 1 },
    });
    s.addText(n.l, {
      x: n.x, y: n.y, w: n.w, h: n.h,
      fontSize: 9.5, bold: true, color: n.tc,
      align: "center", valign: "middle", fontFace: "Trebuchet MS",
    });
  });

  // Arrows + labels
  const ln = (x1, y1, x2, y2) =>
    s.addShape(pres.shapes.LINE, {
      x: Math.min(x1,x2), y: Math.min(y1,y2),
      w: Math.abs(x2-x1) || 0.01, h: Math.abs(y2-y1) || 0.01,
      line: { color: C.muted, width: 1.2 },
    });

  ln(1.8, 1.95, 2.2, 1.95);
  ln(4.1, 1.95, 4.5, 1.95);
  ln(5.48, 2.35, 5.48, 3.1);
  ln(6.45, 1.95, 7.15, 1.95);
  ln(5.48, 3.9,  5.48, 4.35);
  ln(8.35, 2.35, 8.35, 3.1);

  const lbl = (t, x, y) =>
    s.addText(t, { x, y, w: 0.9, h: 0.2, fontSize: 7.5, color: C.text2, align: "center" });

  lbl("QR scan",   1.85, 1.72);
  lbl("SSE/HTTP",  4.12, 1.72);
  lbl("REST API",  5.52, 2.62);
  lbl("MJPEG",     6.73, 1.72);
  lbl("SSE",       5.52, 4.0);
  lbl("USB cam",   8.38, 2.62);

  // RPi note
  s.addShape(pres.shapes.RECTANGLE, {
    x: 7.15, y: 4.12, w: 2.4, h: 0.38,
    fill: { color: C.panel }, line: { color: C.green, width: 1 },
  });
  s.addText("⚡ Alimentato via USB del MiR · viaggia col robot", {
    x: 7.15, y: 4.12, w: 2.4, h: 0.38,
    fontSize: 8, color: C.green, align: "center", valign: "middle",
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — HARDWARE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Componenti Hardware");
  slideNum(s, 4);

  const hw = [
    { e: "🤖", n: "MiR100",        c: C.orange, d: "Robot mobile autonomo. Naviga autonomamente, gestisce missioni via API REST su rete WiFi interna. IP: 192.168.12.20" },
    { e: "🍓", n: "Raspberry Pi",  c: C.green,  d: "Alimentato via porta USB del MiR (5V). Si muove fisicamente col robot. Espone stream MJPEG sulla rete interna (porta 8081)." },
    { e: "📷", n: "Webcam USB",    c: C.teal,   d: "Collegata al Raspberry Pi, montata sul MiR100. Usata per scansione QR code e riconoscimento facciale." },
    { e: "📱", n: "Tablet Balia",  c: C.violet, d: "Interfaccia operatore baia di scarico. Gestisce il flusso di ricezione camion. Pagina principale → /" },
    { e: "📟", n: "Tablet Scarico",c: C.accent, d: "Interfaccia zona destinazione. Conferma ricezione dei pacchi consegnati dal MiR. → /scarico" },
    { e: "💻", n: "Server / PC",   c: C.muted,  d: "Esegue la WebApp Flask principale. Gestisce database, SSE, computer vision, comunicazione MiR. Porta 5000." },
  ];

  hw.forEach((h, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.28 + col * 3.18, y = 1.02 + row * 2.25;
    card(s, x, y, 3.02, 2.08);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.02, h: 0.055, fill: { color: h.c }, line: { color: h.c } });
    s.addText(`${h.e}  ${h.n}`, {
      x: x + 0.12, y: y + 0.1, w: 2.78, h: 0.38,
      fontSize: 15, bold: true, color: h.c, fontFace: "Trebuchet MS",
    });
    s.addText(h.d, {
      x: x + 0.12, y: y + 0.55, w: 2.78, h: 1.45,
      fontSize: 11, color: C.text2, fontFace: "Calibri", valign: "top",
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — RASPBERRY PI
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Raspberry Pi — Telecamera sul Robot");
  slideNum(s, 5);

  // Left: diagram MiR box containing RPi
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.05, w: 4.3, h: 4.2,
    fill: { color: C.panel }, line: { color: C.orange, width: 2 },
  });
  s.addText("MiR100", {
    x: 0.35, y: 1.05, w: 4.3, h: 0.38,
    fontSize: 12, bold: true, color: C.orange, fontFace: "Trebuchet MS",
    align: "center", valign: "middle",
  });
  // USB Port
  s.addShape(pres.shapes.RECTANGLE, {
    x: 2.55, y: 1.6, w: 1.7, h: 0.75,
    fill: { color: C.card }, line: { color: C.yellow, width: 1 },
  });
  s.addText("USB Port\n(5V · 500mA)", {
    x: 2.55, y: 1.6, w: 1.7, h: 0.75,
    fontSize: 9.5, color: C.yellow, align: "center", valign: "middle",
  });
  // Power arrow
  s.addShape(pres.shapes.LINE, {
    x: 1.7, y: 1.97, w: 0.85, h: 0,
    line: { color: C.yellow, width: 2 },
  });
  s.addText("⚡ alimenta", { x: 1.7, y: 1.75, w: 0.9, h: 0.2, fontSize: 8, color: C.yellow, align: "center" });

  // RPi box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 1.55, w: 1.9, h: 1.5,
    fill: { color: "0D1A0D" }, line: { color: C.green, width: 2 },
  });
  s.addText("🍓\nRaspberry Pi\nFlask :8081", {
    x: 0.5, y: 1.55, w: 1.9, h: 1.5,
    fontSize: 11, color: C.green, align: "center", valign: "middle",
  });

  // Webcam box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.2, w: 1.9, h: 0.75,
    fill: { color: C.card }, line: { color: C.teal, width: 1 },
  });
  s.addText("📷  Webcam USB", {
    x: 0.5, y: 3.2, w: 1.9, h: 0.75,
    fontSize: 10, color: C.teal, align: "center", valign: "middle",
  });
  // USB line webcam→RPi
  s.addShape(pres.shapes.LINE, { x: 1.45, y: 3.05, w: 0, h: 0.15, line: { color: C.teal, width: 1.5 } });

  // MJPEG stream arrow out
  s.addShape(pres.shapes.LINE, { x: 2.4, y: 3.55, w: 1.4, h: 0, line: { color: C.green, width: 2 } });
  s.addText("MJPEG /stream", { x: 2.42, y: 3.3, w: 1.4, h: 0.22, fontSize: 8.5, color: C.green, align: "center" });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 4.1, w: 3.9, h: 0.32,
    fill: { color: C.bg }, line: { color: C.muted, width: 0 },
  });
  s.addText("Nessun SSH necessario — autostart all'avvio", {
    x: 0.5, y: 4.1, w: 3.9, h: 0.32,
    fontSize: 9, color: C.muted, align: "center",
  });

  // Right: feature rows
  const pts = [
    { i: "⚡", t: "Alimentato via porta USB del MiR (5V) — nessun alimentatore, nessun cavo aggiuntivo" },
    { i: "🚀", t: "Avvio automatico all'accensione del robot — zero intervento manuale" },
    { i: "📡", t: "Espone stream MJPEG via Flask minimale — endpoint /stream porta 8081" },
    { i: "🔗", t: "WebApp si connette a http://192.168.12.242:8081/stream" },
    { i: "🔀", t: "MJPEGBroadcaster: UNA sola connessione HTTP al Pi → fan-out a tutti i consumer via threading.Condition" },
    { i: "🔍", t: "Consumer del broadcaster: QR Scanner · Face Recognition · Camera feed browser" },
  ];
  pts.forEach((p, i) => {
    const y = 1.05 + i * 0.73;
    card(s, 4.95, y, 4.7, 0.63);
    s.addText(p.i, { x: 5.03, y, w: 0.48, h: 0.63, fontSize: 16, align: "center", valign: "middle" });
    s.addText(p.t, {
      x: 5.57, y: y + 0.08, w: 3.98, h: 0.5,
      fontSize: 11, color: C.text, fontFace: "Calibri", valign: "middle",
    });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — STACK TECNOLOGICO
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Stack Tecnologico");
  slideNum(s, 6);

  const mkCol = (s, x, w, title, titleColor, items) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.02, w, h: 4.3,
      fill: { color: C.panel }, line: { color: C.border, width: 1 },
    });
    s.addText(title, {
      x, y: 1.02, w, h: 0.4,
      fontSize: 12, bold: true, color: titleColor, fontFace: "Trebuchet MS",
      align: "center", valign: "middle",
    });
    items.forEach((it, i) => {
      const y = 1.52 + i * 0.63;
      accentBar(s, x + 0.12, y + 0.04, 0.42, it.c);
      s.addText(it.n, { x: x + 0.26, y: y + 0.02, w: w - 0.35, h: 0.24, fontSize: 12, bold: true, color: C.text, fontFace: "Calibri" });
      s.addText(it.d, { x: x + 0.26, y: y + 0.26, w: w - 0.35, h: 0.24, fontSize: 9.5, color: C.text2, fontFace: "Calibri" });
    });
  };

  mkCol(s, 0.35, 4.4, "BACKEND", C.accent, [
    { n: "Python 3.11 + Flask",          d: "Web framework, routing, SSE endpoint", c: C.accent },
    { n: "SQLAlchemy Core + SQLite",      d: "Database locale, StaticPool thread-safe", c: C.teal },
    { n: "OpenCV + face_recognition",     d: "Computer vision, dlib, face encoding 128-float", c: C.green },
    { n: "pyzbar / cv2",                  d: "QR code detection e decoding real-time", c: C.yellow },
    { n: "requests",                      d: "Comunicazione REST API con MiR100", c: C.orange },
    { n: "threading · queue · RLock",     d: "Operazioni async, SSE queue, thread safety", c: C.purple },
  ]);

  mkCol(s, 5.1, 4.55, "FRONTEND", C.teal, [
    { n: "HTML5 + CSS3 + JS Vanilla",     d: "Nessun framework — zero dipendenze esterne", c: C.accent },
    { n: "EventSource API (SSE)",         d: "Aggiornamenti real-time server→client", c: C.green },
    { n: "Barlow / Barlow Condensed",     d: "Tipografia industriale, ottimizzata per display", c: C.yellow },
    { n: "Design tablet-first",           d: "Layout ottimizzato per schermo touch industriale", c: C.orange },
    { n: "CSS custom properties",         d: "Tema scuro, variabili colore, animazioni pulse", c: C.purple },
    { n: "Fetch API + async/await",       d: "Chiamate REST agli endpoint Flask", c: C.teal },
  ]);
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — FLUSSO
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Flusso di Lavoro");
  slideNum(s, 7);

  const mkPanel = (s, x, w, title, titleColor, steps, dotColor) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.02, w, h: 4.28,
      fill: { color: C.panel }, line: { color: titleColor, width: 1 },
    });
    s.addText(title, {
      x, y: 1.02, w, h: 0.38,
      fontSize: 11.5, bold: true, color: titleColor, fontFace: "Trebuchet MS",
      align: "center", valign: "middle",
    });
    steps.forEach((st, i) => {
      const y = 1.52 + i * 0.74;
      s.addShape(pres.shapes.OVAL, { x: x + 0.14, y: y + 0.06, w: 0.34, h: 0.34, fill: { color: dotColor }, line: { color: dotColor } });
      s.addText(st.n, { x: x + 0.14, y: y + 0.06, w: 0.34, h: 0.34, fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
      s.addText(st.t, { x: x + 0.6, y: y + 0.03, w: w - 0.72, h: 0.56, fontSize: 11, color: C.text, fontFace: "Calibri", valign: "middle" });
    });
  };

  mkPanel(s, 0.3, 4.45, "🖥  TABLET BALIA  ·  /", C.violet,
    [
      { n: "1", t: "Scansione QR code del camion con la telecamera montata sul MiR" },
      { n: "2", t: "Riconoscimento facciale — autenticazione operatore" },
      { n: "3", t: "Scansione QR di ogni pacco → marcati \"in_transito\" nel database" },
      { n: "4", t: "Seconda autenticazione facciale (conferma identità)" },
      { n: "5", t: "Click \"Avvia MiR\" → missione accodata → pagina torna allo stato base" },
    ], C.violet
  );

  mkPanel(s, 5.05, 4.6, "📱  TABLET SCARICO  ·  /scarico", C.accent,
    [
      { n: "6", t: "SSE \"mir_arrivato\" → griglia pacchi si aggiorna in tempo reale" },
      { n: "7", t: "Pulsante \"Conferma ricezione\" si abilita sul pannello alert" },
      { n: "8", t: "Operatore conferma ogni pacco (card → modal → \"Ricevuto\")" },
      { n: "9", t: "All'ultimo pacco confermato: MiR parte automaticamente in ritorno" },
      { n:"10", t: "SSE \"tutti_consegnati\" → banner verde \"Tutti i pacchi ricevuti!\"" },
    ], C.accent
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — DATABASE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Database & Tracciamento");
  slideNum(s, 8);

  const tables = [
    { n: "camion",              f: "targa · fornitore · data_arrivo · qr_code · stato",                 c: C.orange },
    { n: "pacchi",              f: "qr_code JSON · camion_id · destinazione_id · peso_kg · stato",      c: C.accent },
    { n: "operatori",           f: "nome · cognome · ruolo · face_encoding (128 float JSON)",            c: C.green  },
    { n: "destinazioni",        f: "nome zona · tipo (scarico/smistamento) · attiva",                    c: C.teal   },
    { n: "missioni_mir",        f: "tipo (consegna/ritorno) · stato · inizio_ts · fine_ts",              c: C.purple },
    { n: "operazioni_scarico",  f: "operatore_id · camion_id · stato · inizio_ts · fine_ts",             c: C.yellow },
    { n: "config",              f: "chiave = valore  ·  IP dispositivi · credenziali · soglie",          c: C.muted  },
  ];

  // 4 + 3 layout
  tables.forEach((t, i) => {
    const col = i < 4 ? i % 2 : (i - 4) % 2;
    const row = i < 4 ? Math.floor(i / 2) : Math.floor((i - 4) / 2);
    const baseX = i < 4 ? 0.3  : 5.2;
    const baseY = i < 4 ? 1.02 : 1.02;
    const x = baseX + col * 2.32 + (i >= 4 ? col * 0.1 : 0);
    const y = baseY + row * 1.2;
    // Actually, let me simplify to 2-column layout
    const colIdx = i % 2;
    const rowIdx = Math.floor(i / 2);
    const cx = 0.3 + colIdx * 4.85;
    const cy = 1.02 + rowIdx * 1.12;

    card(s, cx, cy, 4.65, 0.98, C.border);
    accentBar(s, cx, cy, 0.98, t.c);
    s.addText(t.n, { x: cx + 0.18, y: cy + 0.06, w: 4.35, h: 0.3, fontSize: 13, bold: true, color: t.c, fontFace: "Trebuchet MS" });
    s.addText(t.f, { x: cx + 0.18, y: cy + 0.42, w: 4.35, h: 0.46, fontSize: 10, color: C.text2, fontFace: "Calibri" });
  });

  // State flow
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.3, y: 4.7, w: 9.45, h: 0.58,
    fill: { color: C.panel }, line: { color: C.accent, width: 1 },
  });
  s.addText("Stato pacco:", { x: 0.45, y: 4.76, w: 1.2, h: 0.28, fontSize: 10, color: C.text2 });
  const stati = [
    { l: "atteso",       c: C.muted  },
    { l: " →",           c: C.border },
    { l: " in_transito", c: C.yellow },
    { l: " →",           c: C.border },
    { l: " consegnato",  c: C.green  },
  ];
  let sx = 1.65;
  stati.forEach((st) => {
    s.addText(st.l, { x: sx, y: 4.76, w: 1.1, h: 0.3, fontSize: 13, bold: true, color: st.c, fontFace: "Calibri" });
    sx += 1.08;
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — API MIR100
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Integrazione API MiR100");
  slideNum(s, 9);

  // Left info
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.02, w: 4.2, h: 2.1,
    fill: { color: C.panel }, line: { color: C.border, width: 1 },
  });
  s.addText("Autenticazione", { x: 0.5, y: 1.08, w: 3.9, h: 0.3, fontSize: 12, bold: true, color: C.accent, fontFace: "Trebuchet MS" });
  s.addText("HTTP Basic Auth\nPassword hashata SHA-256\nBase64(\"user:sha256(password)\")", {
    x: 0.5, y: 1.45, w: 3.9, h: 0.82, fontSize: 11.5, color: C.text, fontFace: "Calibri",
  });

  card(s, 0.35, 3.22, 4.2, 0.7, C.yellow);
  s.addText("Credenziali configurabili da pannello admin.\nOverride tramite variabile ambiente MIR_AUTH_PASS", {
    x: 0.5, y: 3.28, w: 3.9, h: 0.58, fontSize: 10, color: C.yellow, fontFace: "Calibri",
  });

  card(s, 0.35, 4.05, 4.2, 1.2, C.border);
  s.addText("Posizioni missioni:", { x: 0.5, y: 4.12, w: 3.9, h: 0.28, fontSize: 11, bold: true, color: C.text2, fontFace: "Calibri" });
  s.addText("Partenza   →  \"area B1\"\nDestinazione →  \"area A1\"", {
    x: 0.5, y: 4.45, w: 3.9, h: 0.65, fontSize: 13, bold: true, color: C.text, fontFace: "Calibri",
  });

  // Right: API flow steps
  const steps = [
    { n: "1", a: "GET /api/v2.0.0/status",          d: "Verifica stato robot. Blocca se state_id 10 (Emergency) o 12 (Error)." },
    { n: "2", a: "GET /api/v2.0.0/missions",         d: "Lista missioni sul robot. Cerca per nome contenente la posizione destinazione." },
    { n: "3", a: "POST /api/v2.0.0/mission_queue",   d: "Accoda la missione con mission_id. Ottiene queue_entry_id di riferimento." },
    { n: "4", a: "GET /mission_queue/{id}",           d: "Polling ogni 5s (max 10 min). Attende stato \"Done\" o \"Aborted\"." },
    { n: "5", a: "SSE push → mir_arrivato",           d: "Robot arrivato: salva _mir_sessione, notifica tablet scarico." },
    { n: "6", a: "POST /mission_queue (ritorno)",     d: "Dopo conferma tutti i pacchi: accoda missione di ritorno verso area B1." },
  ];
  steps.forEach((st, i) => {
    const y = 1.02 + i * 0.78;
    card(s, 4.85, y, 4.8, 0.67);
    s.addShape(pres.shapes.OVAL, { x: 4.92, y: y + 0.17, w: 0.34, h: 0.34, fill: { color: C.accent }, line: { color: C.accent } });
    s.addText(st.n, { x: 4.92, y: y + 0.17, w: 0.34, h: 0.34, fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    s.addText(st.a, { x: 5.34, y: y + 0.05, w: 4.2, h: 0.27, fontSize: 11, bold: true, color: C.accent, fontFace: "Trebuchet MS" });
    s.addText(st.d, { x: 5.34, y: y + 0.34, w: 4.2, h: 0.28, fontSize: 9.5, color: C.text2, fontFace: "Calibri" });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — ADMIN
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Pannello Amministratore");
  slideNum(s, 10);

  const feats = [
    { e: "👥", t: "Gestione Operatori",   d: "Creazione, modifica, abilit./disabil. Registrazione biometrica da webcam browser." },
    { e: "📍", t: "Destinazioni & Zone",  d: "Gestione zone magazzino. Assegnazione operatori a zone specifiche." },
    { e: "🚚", t: "Camion & Pacchi",      d: "Registrazione camion e pacchi. QR code generati server-side in PNG." },
    { e: "🖨️", t: "Stampa Etichette",    d: "Pagina dedicata stampa QR camion e pacchi. Layout ottimizzato per stampa." },
    { e: "⚙️", t: "Configurazione",       d: "IP MiR, IP RPi, credenziali, soglia face recognition. Salvati su SQLite." },
    { e: "📊", t: "Storico",              d: "Log operazioni scarico e missioni MiR. Ultimi 100 record con timestamp." },
    { e: "🔐", t: "Sicurezza Admin",      d: "Cambio password con verifica bcrypt. Sessioni 8h. Rate limiting login." },
    { e: "📈", t: "Dashboard Stats",      d: "Pacchi oggi · Camion attivi · Missioni · Pacchi in transito · Scarichi." },
  ];

  feats.forEach((f, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.28 + col * 2.39, y = 1.02 + row * 2.25;
    card(s, x, y, 2.22, 2.08);
    s.addText(f.e, { x, y: y + 0.1, w: 2.22, h: 0.46, fontSize: 22, align: "center" });
    s.addText(f.t, { x: x + 0.08, y: y + 0.62, w: 2.06, h: 0.34, fontSize: 11, bold: true, color: C.accent, fontFace: "Trebuchet MS", align: "center" });
    s.addText(f.d, { x: x + 0.08, y: y + 1.0,  w: 2.06, h: 1.0,  fontSize: 9.5, color: C.text2, fontFace: "Calibri", align: "center" });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — SICUREZZA
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Sicurezza");
  slideNum(s, 11);

  const sec = [
    { t: "bcrypt + salt",        d: "Hashing password admin con bcrypt — rainbow table resistant, salt univoco per utente",     c: C.red    },
    { t: "Rate Limiting",        d: "Max 5 tentativi login per IP in 60 secondi → risposta HTTP 429 automatica",                c: C.orange },
    { t: "Face Recognition",     d: "Soglia configurabile (default 0.55). Doppia autenticazione: inizio operazione + conferma", c: C.green  },
    { t: "Sessioni Flask",       d: "Durata 8 ore. SECRET_KEY configurabile via variabile d'ambiente (warning se assente)",     c: C.accent },
    { t: "Thread Safety",        d: "threading.RLock (_stato_lock) su tutto lo stato globale. StaticPool SQLAlchemy per SQLite multi-thread", c: C.yellow },
    { t: "Credenziali MiR",      d: "Override via env var MIR_AUTH_PASS. SHA-256 per autenticazione HTTP Basic MiR100",        c: C.purple },
  ];

  sec.forEach((it, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i < 3 ? i : i - 3;
    const x = col === 0 ? 0.35 : 5.15;
    const y = 1.02 + row * 1.52;
    card(s, x, y, 4.55, 1.38);
    accentBar(s, x, y, 1.38, it.c);
    s.addText("🔒  " + it.t, { x: x + 0.18, y: y + 0.1, w: 4.25, h: 0.34, fontSize: 14, bold: true, color: it.c, fontFace: "Trebuchet MS" });
    s.addText(it.d, { x: x + 0.18, y: y + 0.52, w: 4.25, h: 0.76, fontSize: 11.5, color: C.text, fontFace: "Calibri" });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 12 — SSE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Aggiornamenti Real-Time — Server-Sent Events");
  slideNum(s, 12);

  // Left explanation
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.35, y: 1.02, w: 4.05, h: 2.25,
    fill: { color: C.panel }, line: { color: C.border, width: 1 },
  });
  s.addText("Come funziona SSE", { x: 0.5, y: 1.08, w: 3.75, h: 0.3, fontSize: 12, bold: true, color: C.accent, fontFace: "Trebuchet MS" });
  const info = [
    "Connessione HTTP persistente unidirezionale server→client",
    "Più leggera di WebSocket per stream di soli dati",
    "Endpoint: GET /events  ·  heartbeat ogni 5 secondi",
    "Queue thread-safe (maxsize=200) per eventi da thread secondari",
  ];
  s.addText(
    info.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < info.length - 1, fontSize: 11, color: C.text, paraSpaceAfter: 5 } })),
    { x: 0.5, y: 1.47, w: 3.75, h: 1.72, fontFace: "Calibri", valign: "top" }
  );

  // Mini flow diagram
  const flow = [
    { l: "Thread\nworker", c: C.accent },
    { l: "SSE\nQueue", c: C.yellow },
    { l: "Browser\nEventSource", c: C.green },
  ];
  flow.forEach((f, i) => {
    const x = 0.38 + i * 1.38;
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 3.48, w: 1.18, h: 0.58,
      fill: { color: C.card }, line: { color: f.c, width: 1 },
    });
    s.addText(f.l, { x, y: 3.48, w: 1.18, h: 0.58, fontSize: 8.5, color: f.c, align: "center", valign: "middle" });
    if (i < 2) {
      s.addShape(pres.shapes.LINE, { x: x + 1.18, y: 3.77, w: 0.2, h: 0, line: { color: C.border, width: 1.5 } });
    }
  });
  s.addText("→ /events (SSE stream)", { x: 0.38, y: 4.18, w: 4.0, h: 0.22, fontSize: 9.5, color: C.muted, align: "center" });

  // Right: event types
  const events = [
    { n: "stato",            c: C.accent, d: "Aggiornamento fase workflow (0-4) · inviato ad ogni cambio di fase" },
    { n: "pacco_letto",      c: C.yellow, d: "Pacco scansionato: qr_code, id, destinazione_id, peso · scatta durante la scansione" },
    { n: "mir_arrivato",     c: C.green,  d: "Robot arrivato a destinazione · abilita conferme sul tablet scarico" },
    { n: "mir_ritorno",      c: C.teal,   d: "MiR in rientro verso area B1 · aggiorna stato UI pannello scarico" },
    { n: "tutti_consegnati", c: C.green,  d: "Tutti i pacchi della destinazione confermati · mostra banner successo" },
    { n: "notifica",         c: C.orange, d: "Messaggio info / ok / errore per l'operatore con timestamp HH:MM:SS" },
  ];
  events.forEach((ev, i) => {
    const y = 1.02 + i * 0.77;
    card(s, 4.68, y, 4.97, 0.66);
    s.addShape(pres.shapes.RECTANGLE, { x: 4.68, y, w: 0.05, h: 0.66, fill: { color: ev.c }, line: { color: ev.c } });
    s.addText(ev.n,  { x: 4.8, y: y + 0.07, w: 1.55, h: 0.26, fontSize: 11, bold: true, color: ev.c, fontFace: "Trebuchet MS" });
    s.addText(ev.d,  { x: 4.8, y: y + 0.35, w: 4.72, h: 0.27, fontSize: 9.5, color: C.text2, fontFace: "Calibri" });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 13 — SFIDE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Sfide Tecniche e Soluzioni");
  slideNum(s, 13);

  const ch = [
    { p: "Webcam lenta — 10+ refresh per apparire",        sol: "MJPEGBroadcaster: connessione HTTP singola con fan-out via threading.Condition + thread decode dedicato JPEG→numpy" },
    { p: "QR code con caratteri speciali rompevano onclick",sol: "Sostituzione con attributi data-* e funzione escAttr() per escape corretto — nessun JS inline" },
    { p: "MiR non tornava dopo la consegna",               sol: "_mir_sessione + trigger automatico al conferma ultimo pacco + pulsante manuale \"Rimanda MiR\" come fallback" },
    { p: "Face recognition non funzionava dall'admin",     sol: "Guard su navigator.mediaDevices (richiede HTTPS/localhost) + check videoWidth > 0 prima della cattura frame" },
    { p: "Thread safety sul dizionario stato globale",     sol: "threading.RLock (_stato_lock) su tutte le operazioni read-modify-write + StaticPool SQLAlchemy" },
    { p: "Raspberry Pi senza accesso SSH",                 sol: "Alimentazione USB dal MiR (5V) + script Flask con autostart — zero configurazione post-deploy" },
  ];

  ch.forEach((item, i) => {
    const col = i < 3 ? 0 : 1;
    const row = i < 3 ? i : i - 3;
    const x = col === 0 ? 0.35 : 5.15;
    const y = 1.02 + row * 1.52;

    card(s, x, y, 4.55, 1.38);
    // Problem row (dark red bg)
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.55, h: 0.58, fill: { color: "1A0808" }, line: { color: C.border, width: 0 } });
    s.addText("✗  " + item.p, { x: x + 0.1, y: y + 0.06, w: 4.35, h: 0.46, fontSize: 10.5, color: C.red, fontFace: "Calibri" });
    s.addText("✓  " + item.sol, { x: x + 0.1, y: y + 0.65, w: 4.35, h: 0.66, fontSize: 10.5, color: C.green, fontFace: "Calibri" });
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 14 — RISULTATI
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  darkBg(s);
  slideTitle(s, "Risultati");
  slideNum(s, 14);

  const res = [
    { e: "✅", t: "Sistema completamente funzionante end-to-end con hardware reale MiR100" },
    { e: "🤖", t: "Integrazione hardware reale: MiR100 + Raspberry Pi + webcam fisica sul robot" },
    { e: "☁️", t: "Zero dipendenze cloud — tutto gira in locale sulla rete interna del MiR" },
    { e: "📱", t: "Interfaccia intuitiva ottimizzata per tablet industriali — nessun training necessario" },
    { e: "📦", t: "Tracciabilità completa: atteso → in_transito → consegnato, con storico persistito" },
    { e: "⚙️", t: "Architettura scalabile: nuove zone e operatori configurabili senza modifiche al codice" },
  ];

  res.forEach((r, i) => {
    const y = 1.02 + i * 0.73;
    card(s, 0.45, y, 9.1, 0.62);
    s.addText(r.e, { x: 0.55, y, w: 0.52, h: 0.62, fontSize: 20, align: "center", valign: "middle" });
    s.addText(r.t, { x: 1.18, y: y + 0.1, w: 8.2, h: 0.44, fontSize: 14, color: C.text, fontFace: "Calibri", valign: "middle" });
  });

  // Tech bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 5.08, w: 9.1, h: 0.38,
    fill: { color: C.panel }, line: { color: C.border, width: 1 },
  });
  s.addText(
    "Python 3.11  ·  Flask  ·  SQLite  ·  OpenCV  ·  face_recognition  ·  MiR100 REST API  ·  Raspberry Pi  ·  SSE",
    { x: 0.45, y: 5.08, w: 9.1, h: 0.38, fontSize: 10, color: C.accent, align: "center", valign: "middle", fontFace: "Trebuchet MS" }
  );
}

// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE 15 — GRAZIE
// ═══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bgDeep };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.055, fill: { color: C.accent }, line: { color: C.accent } });

  s.addText("Grazie per l'attenzione", {
    x: 0.8, y: 1.15, w: 8.4, h: 0.95,
    fontSize: 42, bold: true, color: C.text,
    fontFace: "Trebuchet MS", align: "center",
  });
  s.addText("Domande?", {
    x: 0.8, y: 2.2, w: 8.4, h: 0.62,
    fontSize: 30, color: C.accent,
    fontFace: "Trebuchet MS", align: "center",
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.2, y: 3.05, w: 3.6, h: 0.025,
    fill: { color: C.border }, line: { color: C.border },
  });
  s.addText("ITIS — Progetto Robot Mobile Autonomo  ·  2025/2026", {
    x: 0.8, y: 3.2, w: 8.4, h: 0.32,
    fontSize: 12.5, color: C.text2, fontFace: "Trebuchet MS", align: "center",
  });

  const techs = ["Python 3.11", "Flask", "SQLite", "OpenCV", "face_recognition", "MiR100 API", "Raspberry Pi", "SSE"];
  techs.forEach((t, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = 0.9 + col * 2.1, y = 3.75 + row * 0.6;
    card(s, x, y, 1.92, 0.44);
    s.addText(t, { x, y, w: 1.92, h: 0.44, fontSize: 11, color: C.text2, align: "center", valign: "middle", fontFace: "Trebuchet MS" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.56, w: 10, h: 0.065, fill: { color: C.accent }, line: { color: C.accent } });
}

// ─── WRITE ────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "mir100_presentazione.pptx" })
  .then(() => console.log("✅  mir100_presentazione.pptx generato"))
  .catch((e) => { console.error("❌", e); process.exit(1); });
