// make_docx.js — genera mir100_documentazione.docx
"use strict";
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
  LevelFormat, ExternalHyperlink, Bookmark, InternalHyperlink,
  TabStopType, TabStopPosition,
} = require("docx");

// ─── PALETTE ─────────────────────────────────────────────────────────────────
const BLUE1  = "1D4ED8";   // H1
const BLUE2  = "2563EB";   // H2
const BLUE3  = "3B82F6";   // H3 / accent
const DARK   = "1E293B";   // body text
const MUTED  = "64748B";   // captions / muted
const CODE_BG = "F1F5F9";  // code block background
const TH_BG   = "DBEAFE";  // table header background
const TR_ALT  = "F8FAFC";  // table alt row

// DXA helpers (1440 DXA = 1 inch, 567 DXA ≈ 1 cm)
const CM  = 567;
const PG_W  = 11906;           // A4 width
const PG_H  = 16838;           // A4 height
const MAR   = Math.round(2.5 * CM);  // 2.5 cm margins
const TW    = PG_W - MAR * 2;        // text width ≈ 9526 DXA

// ─── NUMBERING CONFIG ────────────────────────────────────────────────────────
const numberingConfig = [
  {
    reference: "bullets",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "•",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 440, hanging: 220 } } },
    }],
  },
  {
    reference: "bullets2",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "–",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 220 } } },
    }],
  },
];

// ─── STYLE HELPERS ───────────────────────────────────────────────────────────
const sp = (before = 0, after = 0, line) => ({
  before, after, ...(line ? { line, lineRule: "auto" } : {}),
});

function h1(text, bookmarkId) {
  const run = new TextRun({ text, bold: true, size: 32, color: BLUE1, font: "Calibri" });
  const children = bookmarkId
    ? [new Bookmark({ id: bookmarkId, children: [run] })]
    : [run];
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: sp(360, 160),
    pageBreakBefore: false,
    children,
  });
}

function h2(text, bookmarkId) {
  const run = new TextRun({ text, bold: true, size: 26, color: BLUE2, font: "Calibri" });
  const children = bookmarkId
    ? [new Bookmark({ id: bookmarkId, children: [run] })]
    : [run];
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: sp(280, 100),
    children,
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: sp(200, 80),
    children: [new TextRun({ text, bold: true, size: 22, color: BLUE3, font: "Calibri" })],
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: sp(0, 120, 276),
    children: [new TextRun({ text, size: 22, color: DARK, font: "Calibri", ...opts })],
  });
}

function bodyRuns(runs) {
  return new Paragraph({
    spacing: sp(0, 120, 276),
    children: runs.map(r =>
      new TextRun({ size: 22, color: DARK, font: "Calibri", ...r })
    ),
  });
}

function bullet(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: sp(0, 80, 276),
    children: [new TextRun({ text, size: 22, color: DARK, font: "Calibri" })],
  });
}

function codeBlock(lines) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" };
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [TW],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: TW, type: WidthType.DXA },
            shading: { fill: CODE_BG, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 200, right: 200 },
            borders: {
              top: border, bottom: border, left: border, right: border,
            },
            children: lines.map(l =>
              new Paragraph({
                spacing: sp(0, 0, 240),
                children: [new TextRun({
                  text: l, font: "Courier New", size: 18, color: "1E293B",
                })],
              })
            ),
          }),
        ],
      }),
    ],
  });
}

function spacer(pt = 120) {
  return new Paragraph({ spacing: sp(0, 0), children: [new TextRun({ text: "" })] });
}

function divider() {
  return new Paragraph({
    spacing: sp(200, 200),
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "DBEAFE", space: 1 } },
    children: [new TextRun({ text: "" })],
  });
}

function infoBox(label, text) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "93C5FD" };
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [TW],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: TW, type: WidthType.DXA },
            shading: { fill: "EFF6FF", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 180, right: 180 },
            borders: { top: border, bottom: border, left: { style: BorderStyle.SINGLE, size: 12, color: BLUE3 }, right: border },
            children: [
              new Paragraph({
                spacing: sp(0, 60),
                children: [new TextRun({ text: label, bold: true, size: 20, color: BLUE3, font: "Calibri" })],
              }),
              new Paragraph({
                spacing: sp(0, 0),
                children: [new TextRun({ text, size: 20, color: DARK, font: "Calibri" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// ─── TABLE BUILDER ───────────────────────────────────────────────────────────
function mkTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  const thBorder = { style: BorderStyle.SINGLE, size: 1, color: "93C5FD" };
  const tdBorder = { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" };

  const mkCell = (text, isHeader, width, bold = false) =>
    new TableCell({
      width: { size: width, type: WidthType.DXA },
      shading: { fill: isHeader ? TH_BG : "FFFFFF", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      verticalAlign: VerticalAlign.CENTER,
      borders: {
        top: isHeader ? thBorder : tdBorder,
        bottom: isHeader ? thBorder : tdBorder,
        left: isHeader ? thBorder : tdBorder,
        right: isHeader ? thBorder : tdBorder,
      },
      children: [new Paragraph({
        spacing: sp(0, 0, 240),
        children: [new TextRun({
          text, size: isHeader ? 20 : 20,
          bold: isHeader || bold, color: isHeader ? BLUE1 : DARK, font: "Calibri",
        })],
      })],
    });

  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => mkCell(h, true, colWidths[i])),
      }),
      ...rows.map((row, ri) =>
        new TableRow({
          children: row.map((cell, i) => {
            const tc = mkCell(cell, false, colWidths[i]);
            if (ri % 2 === 1) {
              tc.options = tc.options || {};
            }
            return tc;
          }),
        })
      ),
    ],
  });
}

// ─── COVER PAGE ──────────────────────────────────────────────────────────────
function coverPage() {
  return [
    spacer(2000),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0, 200),
      children: [new TextRun({ text: "MiR100", size: 72, bold: true, color: BLUE1, font: "Calibri" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0, 100),
      children: [new TextRun({ text: "Sistema di Gestione Magazzino", size: 44, bold: true, color: DARK, font: "Calibri" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0, 600),
      children: [new TextRun({ text: "Automazione della logistica interna con robot mobile autonomo", size: 26, color: MUTED, font: "Calibri", italics: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0, 80),
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "DBEAFE", space: 1 } },
      children: [new TextRun({ text: "" })],
    }),
    spacer(),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(60, 60),
      children: [new TextRun({ text: "Documentazione Tecnica del Progetto", size: 30, bold: true, color: BLUE2, font: "Calibri" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(60, 60),
      children: [new TextRun({ text: "ITIS  ·  Anno scolastico 2025/2026", size: 24, color: MUTED, font: "Calibri" })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ─── TOC PAGE ────────────────────────────────────────────────────────────────
function tocPage() {
  return [
    h1("Indice"),
    new TableOfContents("Indice", { hyperlink: true, headingStyleRange: "1-3" }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ─── SECTION 1 ───────────────────────────────────────────────────────────────
function section1() {
  return [
    h1("1. Introduzione e Obiettivi", "s1"),
    h2("1.1 Descrizione del Progetto", "s11"),
    body("Il progetto MiR100 Warehouse Management System è un sistema di automazione per la gestione della logistica interna di un magazzino. Il sistema utilizza un robot mobile autonomo MiR100 per il trasporto di pacchi dalla baia di scarico dei camion alle zone di destinazione interne al magazzino, eliminando la necessità di trasporto manuale dei carichi."),
    body("Il sistema integra riconoscimento facciale per l'autenticazione degli operatori, tracciamento con QR code per ogni singolo pacco e comunicazione in tempo reale tramite Server-Sent Events (SSE) tra tutti i dispositivi coinvolti."),
    spacer(),
    h2("1.2 Obiettivi", "s12"),
    bullet("Automatizzare il flusso di ricezione e distribuzione dei pacchi"),
    bullet("Implementare autenticazione biometrica degli operatori tramite riconoscimento facciale"),
    bullet("Tracciare ogni pacco dal camion alla destinazione finale tramite QR code"),
    bullet("Coordinare il robot MiR100 tramite la sua API REST"),
    bullet("Fornire interfaccia web real-time accessibile da tablet industriali"),
    bullet("Garantire sicurezza e tracciabilità completa delle operazioni"),
    spacer(),
    h2("1.3 Ambito di Applicazione", "s13"),
    body("Il sistema è progettato per operare interamente in locale sulla rete interna del robot MiR100 (192.168.12.x), senza dipendenze da servizi cloud. Tutti i componenti hardware e software comunicano sulla stessa rete locale dedicata, garantendo autonomia operativa completa anche in assenza di connessione Internet."),
    divider(),
  ];
}

// ─── SECTION 2 ───────────────────────────────────────────────────────────────
function section2() {
  return [
    h1("2. Architettura di Sistema", "s2"),
    h2("2.1 Panoramica dell'Architettura", "s21"),
    body("Il sistema è composto da cinque componenti principali che comunicano tramite protocolli HTTP/REST e Server-Sent Events (SSE) sulla rete locale del MiR:"),
    bullet("Server Flask principale (WebApp) — cervello del sistema"),
    bullet("Tablet Balia — interfaccia baia di scarico camion"),
    bullet("Tablet Scarico — interfaccia zona di destinazione pacchi"),
    bullet("Robot MiR100 — trasporto autonomo dei carichi"),
    bullet("Raspberry Pi con webcam — telecamera mobile montata sul robot"),
    spacer(),
    h2("2.2 Flusso di Comunicazione", "s22"),
    body("Il flusso principale di comunicazione segue questo schema:"),
    spacer(60),
    codeBlock([
      "Camion",
      "  ↓ QR scan",
      "Tablet Balia  ↔  Server Flask (:5000)  ↔  MiR100 (REST API)",
      "                      ↓                     ↓",
      "              Raspberry Pi          Tablet Scarico",
      "              (MJPEG :8081)         (SSE ← mir_arrivato)",
      "",
      "Rete locale: 192.168.12.x",
    ]),
    spacer(),
    h2("2.3 Indirizzi di Rete", "s23"),
    spacer(60),
    mkTable(
      ["Componente", "Indirizzo IP", "Porta", "Protocollo"],
      [
        ["Server Flask",   "rete locale (DHCP)", "5000", "HTTP / SSE"],
        ["MiR100",         "192.168.12.20",      "80",   "REST API"],
        ["Raspberry Pi",   "192.168.12.242",     "8081", "MJPEG stream"],
        ["Tablet Balia",   "DHCP",               "—",    "HTTP / SSE client"],
        ["Tablet Scarico", "DHCP",               "—",    "HTTP / SSE client"],
      ],
      [2800, 2400, 1200, 3126]
    ),
    divider(),
  ];
}

// ─── SECTION 3 ───────────────────────────────────────────────────────────────
function section3() {
  return [
    h1("3. Componenti Hardware", "s3"),
    h2("3.1 Robot Mobile MiR100", "s31"),
    body("Il MiR100 (Mobile Industrial Robots) è un robot autonomo collaborativo progettato per ambienti industriali. Caratteristiche principali:"),
    bullet("Navigazione autonoma con rilevamento e aggiramento ostacoli dinamici"),
    bullet("Connettività WiFi sulla rete interna dedicata (IP: 192.168.12.20)"),
    bullet("API REST v2.0.0 per gestione missioni e monitoraggio stato in tempo reale"),
    bullet("Missioni configurabili con posizioni nominate nel software MiR"),
    bullet("Capacità di carico fino a 100 kg"),
    spacer(),
    h2("3.2 Raspberry Pi — Telecamera Mobile", "s32"),
    body("Il Raspberry Pi rappresenta uno degli elementi più innovativi del sistema: una telecamera che si sposta fisicamente insieme al robot."),
    spacer(60),
    infoBox("Alimentazione", "Il Raspberry Pi è alimentato direttamente dalla porta USB del MiR100 (5V, 500mA). Non richiede alimentazione esterna, si accende e spegne automaticamente con il robot, e non necessita di accesso SSH per la configurazione post-deploy."),
    spacer(),
    h3("Funzionamento"),
    body("All'avvio, il Raspberry Pi esegue automaticamente un server Flask minimale che:"),
    bullet("Acquisisce il feed video dalla webcam USB collegata"),
    bullet("Espone lo stream MJPEG sulla rete interna all'endpoint /stream (porta 8081)"),
    bullet("La WebApp principale si connette a http://192.168.12.242:8081/stream"),
    spacer(),
    h2("3.3 Webcam USB", "s33"),
    body("La webcam USB è collegata al Raspberry Pi e montata fisicamente sul robot MiR100. Viene utilizzata per:"),
    bullet("Scansione dei QR code dei camion e dei pacchi in arrivo"),
    bullet("Acquisizione dei frame per il riconoscimento facciale degli operatori"),
    bullet("Camera feed in tempo reale visualizzabile nel pannello di controllo"),
    spacer(),
    h2("3.4 Tablet Balia e Tablet Scarico", "s34"),
    body("Due tablet industriali accedono alla WebApp tramite browser sulla rete locale:"),
    bullet("Tablet Balia: posizionato alla baia di scarico camion, accede alla pagina principale (/)"),
    bullet("Tablet Scarico: posizionato nella zona di destinazione, accede alla pagina /scarico"),
    spacer(),
    h2("3.5 Server / PC", "s35"),
    body("Il PC server esegue la WebApp Flask principale sulla porta 5000, gestendo database SQLite, logica di business, computer vision (QR e face recognition) e comunicazione REST con il MiR100."),
    divider(),
  ];
}

// ─── SECTION 4 ───────────────────────────────────────────────────────────────
function section4() {
  return [
    h1("4. Architettura Software", "s4"),
    h2("4.1 Stack Tecnologico Backend", "s41"),
    spacer(60),
    mkTable(
      ["Libreria / Tecnologia", "Versione", "Utilizzo"],
      [
        ["Python",                "3.11+",  "Linguaggio principale del server"],
        ["Flask",                 "3.x",    "Web framework WSGI, routing, sessioni, SSE"],
        ["SQLAlchemy Core",       "2.x",    "Query builder SQLite, StaticPool thread-safe"],
        ["SQLite",                "—",      "Database locale (file magazzino.db)"],
        ["OpenCV (cv2)",          "4.x",    "Decodifica JPEG, elaborazione immagini"],
        ["face_recognition",      "1.3+",   "Riconoscimento facciale basato su dlib"],
        ["pyzbar",                "0.1+",   "Decodifica QR code real-time"],
        ["requests",              "2.x",    "Client HTTP per API REST MiR100"],
        ["bcrypt",                "4.x",    "Hashing password con salt"],
        ["qrcode",                "7.x",    "Generazione QR code PNG lato server"],
        ["threading / queue",     "stdlib", "Gestione concorrenza, SSE queue, RLock"],
      ],
      [3600, 1400, 4526]
    ),
    spacer(),
    h2("4.2 Stack Tecnologico Frontend", "s42"),
    bullet("HTML5 + CSS3 + JavaScript vanilla — zero dipendenze da framework esterni"),
    bullet("EventSource API — client SSE nativo del browser per aggiornamenti real-time"),
    bullet("Font Barlow / Barlow Condensed — tipografia industriale ottimizzata per display"),
    bullet("Fetch API + async/await — chiamate REST agli endpoint Flask"),
    bullet("CSS custom properties — tema scuro con variabili colore, animazioni pulse"),
    bullet("Design tablet-first — layout ottimizzato per schermi touch industriali"),
    spacer(),
    h2("4.3 Struttura del Progetto", "s43"),
    spacer(60),
    codeBlock([
      "mir100_webapp_v2/",
      "  app.py                  # Server Flask principale, routing, workflow",
      "  reset_config.py         # Utility aggiornamento configurazione DB",
      "  magazzino.db            # Database SQLite (auto-creato al primo avvio)",
      "  moduli/",
      "    camera_manager.py     # MJPEGBroadcaster + CameraManager",
      "    database.py           # Schema SQLAlchemy, CRUD, seed dati default",
      "    qr_scanner.py         # QRScanner con timeout configurabile",
      "    face_module.py        # RiconoscimentoFacciale (face_recognition + dlib)",
      "  templates/",
      "    tablet_balia.html     # Interfaccia baia di scarico",
      "    tablet_scarico.html   # Interfaccia zona destinazione",
      "    admin_dashboard.html  # Pannello amministratore",
      "    admin_login.html      # Login admin",
      "    qr_stampa.html        # Pagina stampa etichette QR",
    ]),
    divider(),
  ];
}

// ─── SECTION 5 ───────────────────────────────────────────────────────────────
function section5() {
  return [
    h1("5. Database", "s5"),
    h2("5.1 Schema del Database", "s51"),
    body("Il database SQLite (magazzino.db) contiene le seguenti tabelle:"),
    spacer(),
    h3("Tabella: camion"),
    mkTable(
      ["Campo", "Tipo", "Note"],
      [
        ["id",          "INTEGER PK",  "Autoincrement"],
        ["targa",       "TEXT",        "Targa del veicolo"],
        ["qr_code",     "TEXT UNIQUE", "JSON: {tipo, id, targa, fornitore}"],
        ["fornitore",   "TEXT",        "Nome fornitore"],
        ["data_arrivo", "TEXT",        "Data ISO (YYYY-MM-DD)"],
        ["stato",       "TEXT",        "in_arrivo (default)"],
        ["created_at",  "DATETIME",    "Timestamp creazione"],
      ],
      [2400, 2000, 5126]
    ),
    spacer(),
    h3("Tabella: pacchi"),
    mkTable(
      ["Campo", "Tipo", "Note"],
      [
        ["id",             "INTEGER PK",  "Autoincrement"],
        ["qr_code",        "TEXT UNIQUE", "JSON: {tipo, id, destinazione_id, destinazione_nome}"],
        ["camion_id",      "INTEGER FK",  "Riferimento a camion.id"],
        ["destinazione_id","INTEGER FK",  "Riferimento a destinazioni.id"],
        ["descrizione",    "TEXT",        "Descrizione del contenuto"],
        ["peso_kg",        "NUMERIC 6,2", "Peso in chilogrammi"],
        ["stato",          "TEXT",        "atteso | in_transito | consegnato"],
        ["updated_at",     "DATETIME",    "Timestamp ultimo aggiornamento"],
        ["created_at",     "DATETIME",    "Timestamp creazione"],
      ],
      [2400, 2000, 5126]
    ),
    spacer(),
    h3("Tabella: operatori"),
    mkTable(
      ["Campo", "Tipo", "Note"],
      [
        ["id",              "INTEGER PK", "Autoincrement"],
        ["nome",            "TEXT",       "Nome operatore"],
        ["cognome",         "TEXT",       "Cognome operatore"],
        ["ruolo",           "TEXT",       "operatore | supervisore"],
        ["destinazione_id", "INTEGER FK", "Zona assegnata"],
        ["face_encoding",   "TEXT",       "Array JSON di 128 float (face descriptor)"],
        ["attivo",          "BOOLEAN",    "TRUE di default"],
        ["created_at",      "DATETIME",   "Timestamp creazione"],
      ],
      [2400, 2000, 5126]
    ),
    spacer(),
    h3("Tabella: missioni_mir"),
    mkTable(
      ["Campo", "Tipo", "Note"],
      [
        ["id",              "INTEGER PK", "Autoincrement"],
        ["mir_id",          "TEXT",       "ID missione sul robot (guid)"],
        ["tipo",            "TEXT",       "consegna | ritorno"],
        ["destinazione_id", "INTEGER FK", "Zona destinazione"],
        ["stato",           "TEXT",       "pianificata | completata | annullata | errore"],
        ["inizio_ts",       "DATETIME",   "Timestamp avvio"],
        ["fine_ts",         "DATETIME",   "Timestamp completamento"],
      ],
      [2400, 2000, 5126]
    ),
    spacer(),
    h3("Tabella: config"),
    mkTable(
      ["Chiave", "Valore default", "Descrizione"],
      [
        ["raspberry_ip",               "192.168.12.242", "IP Raspberry Pi"],
        ["raspberry_port",             "8081",           "Porta stream MJPEG"],
        ["mir_ip",                     "192.168.12.20",  "IP robot MiR100"],
        ["mir_auth_user",              "itisdelpozzo",   "Username API MiR"],
        ["mir_auth_pass",              "itisdelpozzo",   "Password API MiR (hash SHA-256 in transit)"],
        ["mir_posizione_partenza",     "area B1",        "Nome posizione di partenza"],
        ["mir_posizione_destinazione", "area A1",        "Nome posizione di destinazione"],
        ["face_soglia",                "0.55",           "Soglia distanza face recognition"],
        ["qr_timeout_sec",             "300",            "Timeout scansione QR in secondi"],
      ],
      [3000, 2200, 4326]
    ),
    spacer(),
    h2("5.2 Ciclo di Vita di un Pacco", "s52"),
    spacer(60),
    codeBlock([
      "  atteso",
      "    ↓  (scansione QR durante fase carico)",
      "  in_transito",
      "    ↓  (conferma ricezione sul tablet scarico)",
      "  consegnato",
    ]),
    divider(),
  ];
}

// ─── SECTION 6 ───────────────────────────────────────────────────────────────
function section6() {
  return [
    h1("6. Moduli Software", "s6"),
    h2("6.1 MJPEGBroadcaster e CameraManager", "s61"),
    h3("Problema risolto"),
    body("Con l'implementazione iniziale basata su cv2.VideoCapture, ogni consumer (QR scanner, face recognition, camera feed browser) apriva una connessione HTTP separata verso il Raspberry Pi. Questo saturava la banda WiFi e causava latenze molto elevate: erano necessari 10 o più refresh di pagina per visualizzare il feed video."),
    spacer(),
    h3("Soluzione: architettura Broadcaster-Consumer"),
    body("Il modulo camera_manager.py implementa due classi:"),
    spacer(60),
    codeBlock([
      "class MJPEGBroadcaster:",
      "  # Apre UNA SOLA connessione HTTP persistente verso il Raspberry Pi",
      "  # Legge lo stream MJPEG in chunk da 65.536 byte",
      "  # Estrae frame JPEG completi: SOI marker (0xFF 0xD8) → EOI (0xFF 0xD9)",
      "  # Distribuisce ogni frame via threading.Condition.notify_all()",
      "  # Riconnessione automatica in caso di errore (retry ogni 2s)",
      "",
      "class CameraManager:",
      "  # Consumer del broadcaster",
      "  # Thread dedicato: JPEG → numpy array BGR via cv2.imdecode()",
      "  # leggi_frame() → ritorna sempre l'ultimo frame, non blocca mai",
    ]),
    spacer(),
    h2("6.2 QR Scanner", "s62"),
    body("Il modulo qr_scanner.py utilizza pyzbar per la decodifica dei QR code dai frame numpy acquisiti da CameraManager. Caratteristiche:"),
    bullet("Filtro per tipo atteso (camion / pacco) — scarta QR non pertinenti"),
    bullet("Timeout configurabile dal database (default 300s)"),
    bullet("Modalità scan singolo (scansiona_uno) — ritorna al primo QR valido"),
    bullet("Parsing automatico del payload JSON codificato nel QR"),
    spacer(),
    h2("6.3 Riconoscimento Facciale", "s63"),
    body("Il modulo face_module.py utilizza face_recognition (wrapper di dlib) per l'autenticazione biometrica:"),
    bullet("Rilevamento posizione volti nel frame (modello HOG)"),
    bullet("Estrazione face encoding: vettore di 128 float che identifica univocamente un volto"),
    bullet("Confronto con encoding nel DB tramite distanza euclidea"),
    bullet("Soglia configurabile (default 0.55): distanza < soglia = match"),
    bullet("Timeout configurabile (default 25 secondi)"),
    spacer(),
    h2("6.4 Database (database.py)", "s64"),
    body("Implementa l'accesso al database SQLite tramite SQLAlchemy Core:"),
    bullet("StaticPool: garantisce thread safety con una singola connessione riutilizzata"),
    bullet("check_same_thread=False: permette accesso da thread multipli"),
    bullet("Seed automatico: dati default (admin, config, zone, operatori) alla prima esecuzione"),
    bullet("Tutte le funzioni di scrittura sono transazionali (engine.begin())"),
    divider(),
  ];
}

// ─── SECTION 7 ───────────────────────────────────────────────────────────────
function section7() {
  return [
    h1("7. Flusso Operativo Dettagliato", "s7"),
    h2("7.1 Fase 1 — Scansione Camion", "s71"),
    body("L'operatore avvicina il QR code del camion alla telecamera sul MiR. Il QRScanner decodifica il payload JSON:"),
    spacer(60),
    codeBlock(['{ "tipo": "camion", "id": "CAM-XXXXXX", "targa": "AB123CD", "fornitore": "Fornitore S.r.l." }']),
    spacer(80),
    body("Il sistema cerca il camion nel database tramite l'ID logico e recupera tutti i pacchi associati con stato \"atteso\". Se il camion non esiste o non ha pacchi, viene mostrato un errore all'operatore."),
    spacer(),
    h2("7.2 Fase 2 — Prima Autenticazione Facciale", "s72"),
    body("Il modulo RiconoscimentoFacciale acquisisce frame dalla webcam e rileva volti. Per ogni volto rilevato estrae il face encoding (128 float) e lo confronta con gli encoding salvati nel DB. Se la distanza euclidea è inferiore alla soglia configurata (0.55), l'operatore viene identificato e il sistema avanza alla fase 3."),
    spacer(),
    h2("7.3 Fase 3 — Scansione Pacchi", "s73"),
    body("Il sistema entra in un loop con timeout configurabile (default 300s). Per ogni QR pacco rilevato:"),
    bullet("Verifica che il pacco appartenga al camion scansionato in fase 1"),
    bullet("Aggiorna lo stato nel DB da \"atteso\" a \"in_transito\""),
    bullet("Invia un evento SSE \"pacco_letto\" con i dettagli del pacco"),
    bullet("Il Tablet Scarico riceve l'evento e mostra la notifica \"Pacco in arrivo\""),
    body("Il loop termina quando tutti i pacchi del camion sono stati scansionati, oppure allo scadere del timeout."),
    spacer(),
    h2("7.4 Fase 4 — Seconda Autenticazione e Invio MiR", "s74"),
    body("Seconda autenticazione facciale: verifica che sia lo stesso operatore della fase 2. Poi l'operatore clicca \"Avvia MiR\". Il server esegue in background:"),
    bullet("Verifica stato robot (GET /status) — blocca se stato critico 10 o 12"),
    bullet("Recupera lista missioni (GET /missions) — cerca missione contenente \"area A1\""),
    bullet("Accoda la missione (POST /mission_queue)"),
    bullet("Dopo 2 secondi: reset stato workflow — la balia torna libera per nuove operazioni"),
    bullet("Polling stato missione ogni 5 secondi (max 10 minuti)"),
    bullet("Alla rilevazione stato \"Done\": push SSE \"mir_arrivato\", salva _mir_sessione"),
    spacer(),
    h2("7.5 Fase 5 — Conferma Ricezione (Tablet Scarico)", "s75"),
    body("All'arrivo dell'evento SSE \"mir_arrivato\":"),
    bullet("La griglia pacchi viene aggiornata con tutti i pacchi \"in_transito\" per la zona"),
    bullet("Il pulsante \"Conferma ricezione\" si abilita (era disabilitato come \"Attendi...\""),
    body("L'operatore scarico conferma ogni pacco tramite modal. All'ultimo pacco confermato:"),
    bullet("Il server rileva lista vuota da get_pacchi_by_destinazione() — tutti \"consegnato\""),
    bullet("Invia SSE \"tutti_consegnati\" al Tablet Scarico"),
    bullet("Accoda automaticamente la missione di ritorno verso \"area B1\""),
    bullet("Invia SSE \"mir_ritorno\""),
    divider(),
  ];
}

// ─── SECTION 8 ───────────────────────────────────────────────────────────────
function section8() {
  return [
    h1("8. Comunicazione con MiR100", "s8"),
    h2("8.1 Autenticazione API", "s81"),
    body("Il MiR100 usa HTTP Basic Authentication con un meccanismo di hashing specifico:"),
    bullet("La password non viene inviata in chiaro"),
    bullet("Viene calcolato l'hash SHA-256 della password"),
    bullet("Viene costruita la stringa \"username:sha256_hash\""),
    bullet("Viene applicata la codifica Base64"),
    bullet("Viene inserita nell'header: Authorization: Basic <base64>"),
    bullet("Header aggiuntivo obbligatorio: Accept-Language: en_US"),
    spacer(60),
    codeBlock([
      "# Implementazione Python (da _mir_headers() in app.py)",
      "import hashlib, base64",
      "",
      "pw_hash  = hashlib.sha256(password.encode()).hexdigest()",
      "auth_b64 = base64.b64encode(f'{user}:{pw_hash}'.encode()).decode()",
      "headers  = {",
      '    "Authorization":   f"Basic {auth_b64}",',
      '    "Content-Type":    "application/json",',
      '    "Accept-Language": "en_US",',
      "}",
    ]),
    spacer(),
    h2("8.2 Endpoint Utilizzati", "s82"),
    spacer(60),
    mkTable(
      ["Metodo", "Endpoint", "Descrizione"],
      [
        ["GET",  "/api/v2.0.0/status",             "Stato corrente del robot (state_id, batteria, posizione)"],
        ["GET",  "/api/v2.0.0/missions",            "Lista missioni configurate sul robot"],
        ["POST", "/api/v2.0.0/mission_queue",       "Accoda una missione: body {mission_id: guid}"],
        ["GET",  "/api/v2.0.0/mission_queue/{id}",  "Stato di una specifica missione in coda"],
      ],
      [1000, 3200, 5326]
    ),
    spacer(),
    h2("8.3 State ID del Robot", "s83"),
    spacer(60),
    mkTable(
      ["State ID", "Stato", "Effetto sul sistema"],
      [
        ["3",  "Ready",          "Pronto — missioni accettate"],
        ["4",  "Pause",          "In pausa — missioni accettate"],
        ["5",  "Executing",      "Esegue missione — missioni in coda"],
        ["10", "Emergency Stop", "BLOCCO — il sistema rifiuta l'invio missioni"],
        ["11", "Manual Control", "Controllo manuale"],
        ["12", "Error",          "BLOCCO — il sistema rifiuta l'invio missioni"],
      ],
      [1600, 2400, 5526]
    ),
    spacer(),
    h2("8.4 Gestione Missioni", "s84"),
    body("Il sistema cerca le missioni per nome tramite partial match case-insensitive. Le missioni devono essere pre-configurate nel software MiR con nomi che contengano le stringhe configurate nel DB:"),
    bullet("Missione andata: nome deve contenere \"area A1\""),
    bullet("Missione ritorno: nome deve contenere \"area B1\""),
    body("Il polling della mission_queue usa un intervallo di 5 secondi con timeout massimo di 10 minuti. Alla fine del polling (timeout), la missione viene marcata come \"annullata\" nel database locale."),
    spacer(),
    h2("8.5 Sessione MiR (_mir_sessione)", "s85"),
    body("Quando il robot arriva a destinazione (stato \"Done\" nella mission_queue), il sistema salva in _mir_sessione le informazioni per il viaggio di ritorno:"),
    spacer(60),
    codeBlock([
      "_mir_sessione = {",
      '    "mir_url":  "http://192.168.12.20/api/v2.0.0",',
      '    "headers":  { ... },          # headers con autenticazione',
      '    "missions": [ ... ],          # lista missioni dal robot',
      '    "pos_part": "area B1",        # posizione di partenza',
      '    "pos_dest": "area A1",        # posizione di destinazione',
      "}",
    ]),
    spacer(60),
    body("Questa sessione viene consumata (e il dizionario svuotato) quando tutti i pacchi vengono confermati o tramite il pulsante manuale \"Rimanda MiR\"."),
    divider(),
  ];
}

// ─── SECTION 9 ───────────────────────────────────────────────────────────────
function section9() {
  return [
    h1("9. Server-Sent Events (SSE)", "s9"),
    h2("9.1 Architettura SSE", "s91"),
    body("Il sistema utilizza Server-Sent Events per aggiornamenti real-time verso i browser. L'endpoint GET /events mantiene una connessione HTTP persistente aperta. I thread in background inseriscono eventi in una queue.Queue thread-safe (maxsize=200). Il generatore SSE legge dalla queue con timeout di 5 secondi: se la queue è vuota invia un heartbeat per mantenere la connessione aperta."),
    spacer(60),
    codeBlock([
      "# Lato server (Flask) — app.py",
      "@app.route('/events')",
      "def events():",
      "    def generate():",
      "        yield f'event: stato\\ndata: {json.dumps(stato)}\\n\\n'",
      "        while True:",
      "            try:",
      "                yield _sse_queue.get(timeout=5)",
      "            except queue.Empty:",
      "                yield ': heartbeat\\n\\n'",
      "    return Response(generate(), mimetype='text/event-stream')",
    ]),
    spacer(),
    h2("9.2 Tipi di Evento", "s92"),
    spacer(60),
    mkTable(
      ["Evento", "Payload", "Descrizione"],
      [
        ["stato",            "fase, operatore, camion, pacchi...", "Snapshot completo workflow (fase 0-4)"],
        ["pacco_letto",      "qr_code, id, destinazione_id, peso", "Pacco scansionato durante la fase di carico"],
        ["mir_arrivato",     "destinazione",                       "Robot arrivato — abilita conferme scarico"],
        ["mir_ritorno",      "destinazione",                       "Robot in rientro verso la base"],
        ["tutti_consegnati", "destinazione_id",                    "Tutti i pacchi della zona confermati"],
        ["notifica",         "tipo, msg, ts",                      "Messaggi info / ok / errore per l'operatore"],
      ],
      [2400, 3000, 4126]
    ),
    spacer(),
    h2("9.3 Gestione Lato Client", "s93"),
    body("Il frontend usa l'API nativa EventSource del browser. Ogni tipo di evento ha un handler specifico registrato con addEventListener:"),
    spacer(60),
    codeBlock([
      "// JavaScript — tablet_scarico.html",
      "const es = new EventSource('/events');",
      "",
      "es.addEventListener('mir_arrivato', e => {",
      "    const d = JSON.parse(e.data);",
      "    mirArrivatoActive = true;",
      "    _aggiornaBtnConfermaAlert();   // abilita pulsante conferma",
      "    aggiornaPacchi();              // refresh griglia pacchi",
      "});",
      "",
      "es.addEventListener('tutti_consegnati', e => {",
      "    const d = JSON.parse(e.data);",
      "    if (!destinazioneId || d.destinazione_id === destinazioneId)",
      "        mostraTuttiConsegnati();",
      "});",
    ]),
    body("Il Tablet Scarico filtra gli eventi per destinazione_id, mostrando solo i pacchi della propria zona. La coda di alert (paccoAlertQueue) gestisce le notifiche multiple con badge contatore."),
    divider(),
  ];
}

// ─── SECTION 10 ──────────────────────────────────────────────────────────────
function section10() {
  return [
    h1("10. Sicurezza", "s10"),
    h2("10.1 Autenticazione Amministratore", "s101"),
    bullet("Password hashata con bcrypt (algoritmo Blowfish con salt casuale)"),
    bullet("Rate limiting: max 5 tentativi di login per indirizzo IP in 60 secondi"),
    bullet("Risposta HTTP 429 al superamento del limite"),
    bullet("Sessioni Flask con durata 8 ore"),
    bullet("SECRET_KEY configurabile via variabile d'ambiente (warning a console se assente)"),
    spacer(),
    h2("10.2 Autenticazione Operatori", "s102"),
    bullet("Riconoscimento facciale con soglia di distanza configurabile (default 0.55)"),
    bullet("Doppia autenticazione: richiesta all'inizio dell'operazione e alla conferma finale"),
    bullet("Il secondo riconoscimento verifica che sia lo stesso operatore del primo"),
    bullet("Face encoding: vettore di 128 float salvato come JSON nel database"),
    spacer(),
    h2("10.3 Thread Safety", "s103"),
    bullet("threading.RLock (_stato_lock): protegge il dizionario di stato globale da race condition"),
    bullet("StaticPool SQLAlchemy: garantisce thread safety con singola connessione SQLite"),
    bullet("queue.Queue per eventi SSE: thread-safe per natura"),
    bullet("_update_stato(**kwargs): helper per aggiornamenti atomici dello stato"),
    spacer(),
    h2("10.4 Credenziali MiR100", "s104"),
    bullet("Credenziali configurabili dal pannello admin (salvate in chiaro nel DB — ambienti controllati)"),
    bullet("Override tramite variabile d'ambiente MIR_AUTH_PASS per ambienti di produzione"),
    bullet("Password mai inviata in chiaro: solo hash SHA-256 nell'header Authorization"),
    divider(),
  ];
}

// ─── SECTION 11 ──────────────────────────────────────────────────────────────
function section11() {
  return [
    h1("11. Pannello Amministratore", "s11h"),
    h2("11.1 Accesso", "s111"),
    body("Il pannello admin è accessibile all'URL /admin. Richiede login con username e password. Le credenziali di default sono admin / admin1234 — da modificare prima del deploy in produzione."),
    spacer(),
    h2("11.2 Funzionalità", "s112"),
    h3("Gestione Operatori"),
    body("Creazione, modifica dati anagrafici, attivazione/disattivazione. Registrazione face encoding tramite webcam del browser (richiede HTTPS o localhost per l'API navigator.mediaDevices). Reset encoding individuale."),
    h3("Gestione Destinazioni e Zone"),
    body("Creazione e modifica delle zone del magazzino (nome, tipo smistamento/scarico, stato attivo). Le zone vengono mostrate nel selettore del Tablet Scarico."),
    h3("Gestione Camion e Pacchi"),
    body("Registrazione camion con generazione automatica ID univoco (CAM-XXXXXX) e QR code PNG. Aggiunta pacchi con destinazione, descrizione e peso. Eliminazione pacchi. Visualizzazione QR code generato inline."),
    h3("Configurazione Sistema"),
    body("Modifica parametri operativi in tempo reale senza riavviare l'applicazione: IP Raspberry Pi, porta stream, IP MiR100, credenziali MiR, posizioni missioni (partenza/destinazione), soglia face recognition, timeout scansione QR."),
    h3("Storico e Dashboard"),
    body("Log delle ultime 100 operazioni di scarico e missioni MiR con timestamp e stato. Dashboard con statistiche in tempo reale: pacchi processati oggi, camion attivi, missioni eseguite, pacchi in transito."),
    divider(),
  ];
}

// ─── SECTION 12 ──────────────────────────────────────────────────────────────
function section12() {
  return [
    h1("12. Interfacce Utente", "s12h"),
    h2("12.1 Tablet Balia (/)", "s121"),
    body("L'interfaccia principale gestisce il workflow in 5 fasi progressive:"),
    spacer(60),
    mkTable(
      ["Fase", "Stato", "Azione disponibile"],
      [
        ["0", "Idle / standby",           "Pulsante \"Scansiona Camion\""],
        ["1", "Camion scansionato",        "Pulsante \"Autentica Operatore\" (face)"],
        ["2", "Scansione pacchi in corso", "Automatica — contatore live pacchi rimanenti"],
        ["3", "Pacchi completati",         "Pulsante \"Conferma Identità\" (face)"],
        ["4", "Operatore confermato",      "Pulsante \"Avvia MiR\""],
      ],
      [1000, 3000, 5526]
    ),
    body("Le transizioni tra fasi avvengono tramite l'evento SSE \"stato\". La camera feed dal MiR è visibile durante le fasi di scansione. Dopo l'avvio del MiR la pagina torna automaticamente alla fase 0."),
    spacer(),
    h2("12.2 Tablet Scarico (/scarico)", "s122"),
    body("Layout a due colonne:"),
    bullet("Colonna sinistra: selezione zona, stato MiR, pannello alert pacchi, pulsante \"Rimanda MiR\", storico locale"),
    bullet("Colonna destra: griglia card dei pacchi in_transito verso la zona selezionata"),
    spacer(60),
    infoBox("Comportamento pulsante Conferma ricezione", "Il pulsante negli alert è DISABILITATO (⋯ In arrivo…) durante la fase di scansione. Viene ABILITATO solo quando arriva l’evento SSE \"mir_arrivato\", garantendo che la conferma avvenga solo quando il robot è fisicamente presente."),
    spacer(80),
    body("Funzionalità aggiuntive: coda di alert multipli con badge \"+N in coda\", pulsante manuale \"Rimanda MiR\" (visibile dopo mir_arrivato), banner verde \"Tutti i pacchi ricevuti!\" dopo l'ultimo conferma."),
    divider(),
  ];
}

// ─── SECTION 13 ──────────────────────────────────────────────────────────────
function section13() {
  return [
    h1("13. Avvio e Configurazione", "s13h"),
    h2("13.1 Requisiti", "s131"),
    body("Python 3.11+ con le seguenti librerie:"),
    spacer(60),
    codeBlock([
      "pip install flask sqlalchemy opencv-python face_recognition",
      "pip install dlib pyzbar requests bcrypt qrcode Pillow numpy",
    ]),
    spacer(),
    h2("13.2 Prima Esecuzione", "s132"),
    spacer(60),
    codeBlock(["python app.py"]),
    spacer(80),
    body("Al primo avvio vengono creati automaticamente:"),
    bullet("Database magazzino.db con schema completo"),
    bullet("Admin di default: admin / admin1234"),
    bullet("Configurazione di default con IP e credenziali"),
    bullet("4 zone di esempio (Zona A Scaffali, Zona B Frigo, Zona C Spedizioni, Zona scarico)"),
    bullet("4 operatori di esempio senza face encoding"),
    spacer(),
    h2("13.3 Configurazione Manuale", "s133"),
    body("Per aggiornare la configurazione in un database esistente, modificare e eseguire reset_config.py:"),
    spacer(60),
    codeBlock(["python reset_config.py"]),
    spacer(),
    h2("13.4 URL di Accesso", "s134"),
    spacer(60),
    mkTable(
      ["Interfaccia", "URL", "Descrizione"],
      [
        ["Pannello Balia",   "http://<ip>:5000/",          "Interfaccia baia di scarico"],
        ["Pannello Scarico", "http://<ip>:5000/scarico",   "Interfaccia zona destinazione"],
        ["Pannello Admin",   "http://<ip>:5000/admin",     "Gestione e configurazione"],
        ["Camera feed",      "http://<ip>:5000/camera",    "Proxy stream MJPEG"],
        ["SSE events",       "http://<ip>:5000/events",    "Stream eventi real-time"],
      ],
      [2400, 3000, 4126]
    ),
    divider(),
  ];
}

// ─── SECTION 14 ──────────────────────────────────────────────────────────────
function section14() {
  return [
    h1("14. Problemi Tecnici e Soluzioni", "s14"),
    h2("14.1 Webcam Lenta", "s141"),
    h3("Problema"),
    body("Con l'implementazione iniziale basata su cv2.VideoCapture, ogni consumer (QR scanner, face recognition, camera feed browser) apriva una connessione HTTP separata verso il Raspberry Pi. Questo saturava la banda WiFi e causava latenze molto elevate: erano necessari 10 o più refresh di pagina per visualizzare il feed."),
    h3("Soluzione"),
    body("Architettura MJPEGBroadcaster con singola connessione HTTP persistente e fan-out tramite threading.Condition.notify_all(). Thread decode dedicato per la conversione JPEG→numpy separato dal thread di lettura stream."),
    spacer(),
    h2("14.2 QR Code con Caratteri Speciali", "s142"),
    h3("Problema"),
    body("I payload JSON dei QR code contengono virgolette doppie che rompevano il codice JavaScript inline nell'attributo onclick: onclick=\"openModal('...{\\\"id\\\":\\\"PKG\\\"}...')\"."),
    h3("Soluzione"),
    body("Sostituzione degli onclick inline con attributi data-* (data-qr, data-pid, data-desc, ecc.) e funzione escAttr() lato JavaScript per l'escape corretto dei caratteri speciali (&amp; &quot; &lt; &gt;)."),
    spacer(),
    h2("14.3 MiR Non Tornava Dopo la Consegna", "s143"),
    h3("Problema"),
    body("Il ritorno era accodato immediatamente all'arrivo del MiR, senza attendere la conferma dei pacchi. In seguito, un bug di timing rendeva _mir_sessione vuota al momento della conferma, impedendo l'invio della missione di ritorno."),
    h3("Soluzione"),
    body("Salvataggio della sessione in _mir_sessione all'arrivo. Il ritorno viene triggerato da conferma_ricezione() quando get_pacchi_by_destinazione() restituisce lista vuota (tutti \"consegnato\"). Pulsante manuale \"Rimanda MiR\" come fallback."),
    spacer(),
    h2("14.4 Face Recognition Non Funzionava dall'Admin", "s144"),
    h3("Problema"),
    body("navigator.mediaDevices risulta undefined su pagine HTTP non-localhost. Il browser blocca l'accesso alla webcam per motivi di sicurezza (accesso alla camera richiede HTTPS o localhost)."),
    h3("Soluzione"),
    body("Guard esplicito con controllo navigator.mediaDevices prima di chiamare getUserMedia, con messaggio di errore chiaro all'utente. Verifica videoWidth > 0 prima di acquisire il frame per evitare la cattura di frame neri."),
    spacer(),
    h2("14.5 Thread Safety", "s145"),
    h3("Problema"),
    body("Il dizionario di stato globale veniva modificato da thread multipli (worker scansione, worker MiR, richieste HTTP Flask) causando potenziali race condition con comportamenti imprevedibili."),
    h3("Soluzione"),
    body("threading.RLock (_stato_lock) attorno a tutte le operazioni di lettura-modifica-scrittura. Helper _update_stato(**kwargs) per aggiornamenti atomici. StaticPool SQLAlchemy per thread safety del database."),
    spacer(),
    h2("14.6 Raspberry Pi Senza Accesso SSH", "s146"),
    h3("Problema"),
    body("Il Raspberry Pi montato sul robot non era raggiungibile via SSH dall'esterno della rete MiR, rendendo impossibile la configurazione e il deploy tradizionale."),
    h3("Soluzione"),
    body("Alimentazione via porta USB del MiR (nessun cavo aggiuntivo necessario) e script Flask con autostart tramite systemd o rc.local. Zero configurazione necessaria dopo il deploy iniziale: il sistema si avvia automaticamente con il robot."),
    divider(),
  ];
}

// ─── SECTION 15 — GLOSSARIO ──────────────────────────────────────────────────
function section15() {
  return [
    h1("15. Glossario", "s15"),
    spacer(60),
    mkTable(
      ["Termine", "Definizione"],
      [
        ["MiR100",         "Mobile Industrial Robot 100 — robot mobile autonomo collaborativo di MiR Robots"],
        ["MJPEG",          "Motion JPEG — formato stream video come sequenza continua di frame JPEG"],
        ["SSE",            "Server-Sent Events — protocollo HTTP per stream unidirezionale server→client"],
        ["Face Encoding",  "Vettore di 128 float che rappresenta le caratteristiche biometriche uniche di un volto"],
        ["QR Code",        "Quick Response Code — codice a matrice 2D per codifica di dati (JSON nel nostro caso)"],
        ["REST API",       "Representational State Transfer — architettura API basata su HTTP con risorse stateless"],
        ["BCrypt",         "Algoritmo crittografico per l'hashing sicuro delle password con salt incorporato"],
        ["SQLAlchemy Core","Livello di astrazione SQL per Python senza ORM completo — query tipizzate"],
        ["StaticPool",     "Strategia di pool SQLAlchemy che mantiene e riutilizza una singola connessione DB"],
        ["RLock",          "Reentrant Lock — mutex che può essere acquisito più volte dallo stesso thread"],
        ["HOG",            "Histogram of Oriented Gradients — modello usato da dlib per il rilevamento volti"],
        ["SHA-256",        "Secure Hash Algorithm 256-bit — usato per hashare la password MiR nell'autenticazione"],
        ["DHCP",           "Dynamic Host Configuration Protocol — assegnazione automatica IP sulla rete"],
        ["DXA",            "Twentieths of a point — unità di misura nei file OOXML (1440 DXA = 1 pollice)"],
      ],
      [2800, 6726]
    ),
  ];
}

// ─── ASSEMBLE DOCUMENT ───────────────────────────────────────────────────────
const allContent = [
  ...coverPage(),
  ...tocPage(),
  ...section1(),
  ...section2(),
  ...section3(),
  ...section4(),
  ...section5(),
  ...section6(),
  ...section7(),
  ...section8(),
  ...section9(),
  ...section10(),
  ...section11(),
  ...section12(),
  ...section13(),
  ...section14(),
  ...section15(),
];

const doc = new Document({
  numbering: { config: numberingConfig },
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22, color: DARK } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 32, bold: true, font: "Calibri", color: BLUE1 },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 26, bold: true, font: "Calibri", color: BLUE2 },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run:       { size: 22, bold: true, font: "Calibri", color: BLUE3 },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size:   { width: PG_W, height: PG_H },
          margin: { top: MAR, right: MAR, bottom: MAR, left: MAR },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: "DBEAFE", space: 4 } },
              spacing: sp(0, 80),
              tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
              children: [
                new TextRun({ text: "MiR100 — Sistema di Gestione Magazzino", size: 16, color: MUTED, font: "Calibri" }),
                new TextRun({ text: "\t", size: 16 }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED }),
              ],
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 2, color: "DBEAFE", space: 4 } },
              spacing: sp(80, 0),
              children: [
                new TextRun({ text: "ITIS · Anno scolastico 2025/2026  ·  Pagina ", size: 16, color: MUTED, font: "Calibri" }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: MUTED }),
                new TextRun({ text: " di ", size: 16, color: MUTED }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: MUTED }),
              ],
            }),
          ],
        }),
      },
      children: allContent,
    },
  ],
});

Packer.toBuffer(doc)
  .then(buf => {
    fs.writeFileSync("mir100_documentazione.docx", buf);
    console.log("✅  mir100_documentazione.docx generato");
  })
  .catch(e => { console.error("❌", e); process.exit(1); });
