// make_docx_semplice.js — genera mir100_doc_semplice.docx
"use strict";
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat,
} = require("docx");

const BLUE  = "1D4ED8";
const DARK  = "1E293B";
const MUTED = "64748B";
const TH_BG = "DBEAFE";

const CM  = 567;
const PG_W  = 11906;
const PG_H  = 16838;
const MAR   = Math.round(2.5 * CM);
const TW    = PG_W - MAR * 2;

const numberingConfig = [{
  reference: "bullets",
  levels: [{
    level: 0, format: LevelFormat.BULLET, text: "•",
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 440, hanging: 220 } } },
  }],
}];

// helpers
const sp = (before = 0, after = 0) => ({ spacing: { before, after } });

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    ...sp(320, 120),
    children: [new TextRun({ text, bold: true, size: 32, color: BLUE, font: "Calibri" })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    ...sp(240, 80),
    children: [new TextRun({ text, bold: true, size: 26, color: "2563EB", font: "Calibri" })],
  });
}
function body(text) {
  return new Paragraph({
    ...sp(0, 100),
    children: [new TextRun({ text, size: 22, color: DARK, font: "Calibri" })],
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    ...sp(0, 60),
    children: [new TextRun({ text, size: 22, color: DARK, font: "Calibri" })],
  });
}
function space() {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 100, after: 100 } });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// semplice tabella a 2 colonne chiave/valore
function kvTable(rows) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [Math.round(TW * 0.35), Math.round(TW * 0.65)],
    rows: rows.map((r, i) => new TableRow({
      children: [
        new TableCell({
          borders,
          width: { size: Math.round(TW * 0.35), type: WidthType.DXA },
          shading: i === 0 ? { fill: TH_BG, type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: r[0], bold: i === 0, size: 20, font: "Calibri", color: DARK })] })],
        }),
        new TableCell({
          borders,
          width: { size: Math.round(TW * 0.65), type: WidthType.DXA },
          shading: i === 0 ? { fill: TH_BG, type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: r[1], bold: i === 0, size: 20, font: "Calibri", color: DARK })] })],
        }),
      ],
    })),
  });
}

// ─── CONTENUTO ────────────────────────────────────────────────────────────────

const content = [

  // COPERTINA
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 2000, after: 200 },
    children: [new TextRun({ text: "MiR100", bold: true, size: 72, color: BLUE, font: "Calibri" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: "Sistema di Gestione Magazzino", bold: true, size: 40, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 800 },
    children: [new TextRun({ text: "Automazione della logistica interna con robot mobile autonomo", size: 26, color: MUTED, font: "Calibri" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: "Documentazione del Progetto", size: 24, color: MUTED, font: "Calibri" })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: "ITIS  ·  Anno scolastico 2025/2026", size: 22, color: MUTED, font: "Calibri" })],
  }),

  pageBreak(),

  // 1. INTRODUZIONE
  h1("1. Introduzione"),
  body("Per il progetto di quest'anno abbiamo realizzato un sistema di gestione del magazzino che sfrutta il robot mobile MiR100. L'idea di base era quella di automatizzare il trasporto dei pacchi all'interno del magazzino, eliminando la necessità di spostamenti manuali da parte degli operatori."),
  space(),
  body("Il sistema funziona così: quando arriva un camion, l'operatore alla baia di scarico registra il camion e scansiona tutti i pacchi tramite tablet. Poi il robot parte in automatico verso la zona di destinazione. Lì un secondo operatore conferma la ricezione di ogni pacco, e il robot torna da solo alla base."),
  space(),
  body("Tutto gira su una rete locale dedicata, senza bisogno di connessione internet. L'interfaccia è una web app Flask accessibile da browser."),

  space(),
  h2("Cosa abbiamo voluto fare"),
  bullet("Automatizzare il trasporto dei pacchi dal camion alla zona di destinazione"),
  bullet("Autenticare gli operatori tramite riconoscimento facciale (niente badge o PIN)"),
  bullet("Tracciare ogni pacco con QR code, dall'arrivo alla consegna"),
  bullet("Controllare il MiR100 tramite la sua API REST"),
  bullet("Avere un'interfaccia web semplice e usabile da tablet"),

  pageBreak(),

  // 2. HARDWARE
  h1("2. Hardware utilizzato"),
  body("Il sistema utilizza diversi componenti fisici che comunicano sulla rete interna del MiR (range 192.168.12.x)."),
  space(),

  kvTable([
    ["Componente", "Descrizione"],
    ["MiR100", "Il robot mobile. Naviga in autonomia nel magazzino e riceve le missioni via API REST. IP fisso: 192.168.12.20"],
    ["Raspberry Pi", "Montato fisicamente sul robot, alimentato dalla porta USB del MiR. Gira in automatico all'avvio e trasmette il video della webcam sulla rete (porta 8081)"],
    ["Webcam USB", "Collegata al Raspberry Pi, serve per scansionare i QR code e per il riconoscimento facciale degli operatori"],
    ["Tablet Balia", "Tablet posizionato alla baia di scarico. Accede alla pagina principale della web app"],
    ["Tablet Scarico", "Tablet nella zona di destinazione dei pacchi. Accede alla pagina /scarico"],
    ["Server/PC", "Esegue la web app Flask. Raggiungibile da tutti i dispositivi sulla rete locale"],
  ]),

  space(),
  h2("Raspberry Pi + Webcam sul robot"),
  body("Uno dei problemi che abbiamo affrontato era: come collegare una telecamera al robot senza cavi lunghi e senza poter configurare il MiR dall'esterno?"),
  space(),
  body("La soluzione è stata montare un Raspberry Pi direttamente sopra il robot, alimentandolo dalla porta USB del MiR. Il Pi si avvia da solo quando si accende il robot e fa partire un piccolo server Flask che trasmette il video in formato MJPEG sulla rete interna. La web app si connette a quell'indirizzo e usa il video sia per scansionare i QR code che per il riconoscimento facciale."),
  space(),
  body("In questo modo la telecamera si sposta fisicamente insieme al robot, senza dover gestire nessun cavo aggiuntivo e senza aver bisogno di accedere via SSH al Pi ogni volta."),

  pageBreak(),

  // 3. SOFTWARE
  h1("3. Software e tecnologie"),
  body("Il backend è scritto in Python con Flask. Il frontend è HTML/CSS/JavaScript puro, senza framework esterni. Abbiamo scelto di tenerlo semplice per non aggiungere complessità inutile."),
  space(),

  h2("Backend (Python)"),
  bullet("Flask — gestisce tutte le route HTTP e la logica applicativa"),
  bullet("SQLAlchemy + SQLite — database locale per pacchi, camion, operatori e missioni"),
  bullet("OpenCV + face_recognition — acquisizione video e riconoscimento facciale"),
  bullet("pyzbar — lettura dei QR code dai frame video"),
  bullet("requests — chiamate all'API REST del MiR100"),
  bullet("Server-Sent Events (SSE) — aggiornamenti in tempo reale verso il browser"),
  bullet("threading + RLock — gestione operazioni parallele in modo thread-safe"),

  space(),
  h2("Frontend (browser)"),
  bullet("HTML5, CSS3, JavaScript vanilla"),
  bullet("EventSource API per ricevere gli eventi SSE dal server"),
  bullet("Design responsive ottimizzato per tablet (touch, font grandi, pulsanti ampi)"),
  bullet("Font Barlow e Barlow Condensed"),

  pageBreak(),

  // 4. FLUSSO
  h1("4. Come funziona — flusso operativo"),
  body("Il flusso di lavoro è diviso tra due interfacce: la pagina principale (Tablet Balia) e la pagina /scarico (Tablet Scarico)."),
  space(),

  h2("Tablet Balia — baia di scarico"),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "1.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "La webcam sul robot inquadra il QR code del camion per identificarlo", size: 22, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "2.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "L'operatore guarda la telecamera per il riconoscimento facciale (prima autenticazione)", size: 22, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "3.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "Scansione del QR code di ogni pacco del camion, uno per uno", size: 22, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "4.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "Secondo riconoscimento facciale per confermare l'identità", size: 22, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "5.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "Click su \"Avvia MiR\" → la missione viene accodata sul robot → la pagina torna subito allo stato iniziale", size: 22, color: DARK, font: "Calibri" })],
  }),

  space(),
  h2("Tablet Scarico — zona di destinazione"),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "6.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "Quando il MiR arriva, la pagina /scarico si aggiorna automaticamente (via SSE) e mostra i pacchi in arrivo", size: 22, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "7.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "L'operatore clicca \"Ricevuto\" su ogni pacco", size: 22, color: DARK, font: "Calibri" })],
  }),
  new Paragraph({
    ...sp(0, 80),
    children: [new TextRun({ text: "8.  ", bold: true, size: 22, font: "Calibri" }), new TextRun({ text: "Quando viene confermato l'ultimo pacco, il MiR parte automaticamente in ritorno alla base", size: 22, color: DARK, font: "Calibri" })],
  }),

  pageBreak(),

  // 5. DATABASE
  h1("5. Database"),
  body("Usiamo SQLite come database locale. È semplice, non richiede un server separato e va benissimo per un'applicazione che gira su una singola macchina. Le tabelle principali sono:"),
  space(),

  kvTable([
    ["Tabella", "Cosa contiene"],
    ["camion", "Targa, fornitore, data di arrivo e QR code identificativo di ogni camion registrato"],
    ["pacchi", "QR code del pacco, camion di provenienza, zona di destinazione, peso e stato attuale"],
    ["operatori", "Nome, cognome, ruolo e il face encoding (128 valori numerici) per il riconoscimento facciale"],
    ["destinazioni", "Le zone del magazzino (nome e tipo)"],
    ["missioni_mir", "Log di ogni missione: tipo (consegna o ritorno), stato e timestamp"],
    ["config", "Parametri configurabili: IP del MiR, IP del Raspberry, credenziali, soglie"],
  ]),

  space(),
  body("Lo stato di ogni pacco cambia così nel corso dell'operazione:"),
  space(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    ...sp(80, 80),
    children: [
      new TextRun({ text: "atteso", bold: true, size: 22, font: "Calibri", color: MUTED }),
      new TextRun({ text: "   →   ", size: 22, font: "Calibri", color: MUTED }),
      new TextRun({ text: "in_transito", bold: true, size: 22, font: "Calibri", color: "2563EB" }),
      new TextRun({ text: "   →   ", size: 22, font: "Calibri", color: MUTED }),
      new TextRun({ text: "consegnato", bold: true, size: 22, font: "Calibri", color: "16A34A" }),
    ],
  }),

  pageBreak(),

  // 6. API MIR
  h1("6. Integrazione con il MiR100"),
  body("Il MiR100 espone una REST API (versione 2.0.0) sulla porta 80. Per autenticarsi bisogna mandare le credenziali nel formato Basic Auth, ma con la password hashata in SHA-256 (non in chiaro)."),
  space(),
  body("La sequenza che seguiamo ogni volta che vogliamo far partire il robot è:"),
  space(),
  bullet("Verificare che il robot non sia in uno stato critico (GET /api/v2.0.0/status)"),
  bullet("Cercare la missione giusta per nome tra quelle configurate nel MiR (GET /api/v2.0.0/missions)"),
  bullet("Accodare la missione (POST /api/v2.0.0/mission_queue) e salvare l'ID restituito"),
  bullet("Aspettare che lo stato della missione diventi \"Done\" controllando ogni 5 secondi (max 10 minuti)"),
  bullet("Quando arriva il \"Done\": inviare l'evento SSE \"mir_arrivato\" alla pagina /scarico"),
  bullet("Dopo che tutti i pacchi sono stati confermati: accodare la missione di ritorno"),
  space(),
  body("Le posizioni del magazzino (partenza e destinazione) sono configurabili dal pannello admin, così non bisogna toccare il codice se cambia il layout del magazzino."),

  pageBreak(),

  // 7. AGGIORNAMENTI IN TEMPO REALE
  h1("7. Aggiornamenti in tempo reale (SSE)"),
  body("Per far sì che la pagina /scarico si aggiorni automaticamente senza dover ricaricare il browser, abbiamo usato i Server-Sent Events (SSE). È una connessione HTTP persistente in cui il server può mandare messaggi al browser in qualsiasi momento."),
  space(),
  body("Lo abbiamo preferito ai WebSocket perché è più semplice da gestire lato server con Flask e va benissimo per il nostro caso (comunicazione solo server → client)."),
  space(),

  kvTable([
    ["Evento SSE", "Quando scatta"],
    ["stato", "Aggiornamento della fase del workflow (0-4)"],
    ["pacco_letto", "Un pacco è stato scansionato (con QR code, destinazione e peso)"],
    ["mir_arrivato", "Il robot è arrivato alla destinazione — sblocca i pulsanti nella /scarico"],
    ["mir_ritorno", "Il robot è partito in rientro verso la base"],
    ["tutti_consegnati", "Tutti i pacchi della spedizione sono stati confermati"],
    ["notifica", "Messaggi informativi o di errore per l'operatore"],
  ]),

  pageBreak(),

  // 8. PROBLEMI AFFRONTATI
  h1("8. Problemi che abbiamo incontrato"),
  body("Durante lo sviluppo abbiamo dovuto risolvere vari problemi, alcuni abbastanza insidiosi."),
  space(),

  h2("Webcam lenta all'avvio"),
  body("La webcam sul Raspberry Pi non era subito disponibile quando la web app si collegava allo stream. Ci volevano anche 10+ tentativi prima che il feed apparisse. Abbiamo risolto con un'architettura a broadcaster: un unico thread si collega al Pi e distribuisce i frame a tutti i consumer (QR scanner, face recognition, camera feed) tramite un Condition di threading. In questo modo c'è sempre una sola connessione HTTP attiva verso il Raspberry Pi e i frame vengono condivisi."),
  space(),

  h2("QR code con caratteri speciali"),
  body("I QR code contenevano caratteri JSON con virgolette e parentesi graffe che rompevano gli attributi onclick nell'HTML. Abbiamo risolto mettendo i dati in attributi data-* e leggendoli via JavaScript, con una funzione di escape per gestire i caratteri speciali."),
  space(),

  h2("Il MiR non tornava indietro"),
  body("Inizialmente il robot tornava alla base subito dopo aver consegnato, senza aspettare che l'operatore confermasse la ricezione dei pacchi. Abbiamo aggiunto una variabile di sessione (_mir_sessione) che salva i dati della missione in corso. Il ritorno viene triggerato solo quando l'operatore conferma l'ultimo pacco nella pagina /scarico."),
  space(),

  h2("Race condition tra robot e operatore"),
  body("C'era un bug per cui l'operatore poteva cliccare \"Ricevuto\" prima che il robot fosse arrivato, mandando in errore la chiamata di ritorno. Abbiamo disabilitato il pulsante di conferma per default e lo abilitiamo solo quando arriva l'evento SSE \"mir_arrivato\", garantendo che la sessione MiR sia sempre popolata in quel momento."),
  space(),

  h2("Thread safety"),
  body("Più thread accedono contemporaneamente allo stato dell'applicazione (il thread SSE, il thread del polling MiR, le richieste HTTP). Abbiamo protetto tutte le operazioni read-modify-write con un threading.RLock per evitare race condition."),

  pageBreak(),

  // 9. PANNELLO ADMIN
  h1("9. Pannello di amministrazione"),
  body("Il pannello admin è accessibile solo con credenziali (password hashata con bcrypt). Da lì si gestisce tutto il sistema:"),
  space(),
  bullet("Creazione e modifica degli operatori (nome, ruolo, destinazione assegnata)"),
  bullet("Registrazione biometrica: si acquisisce il face encoding dell'operatore dalla webcam del browser"),
  bullet("Gestione delle zone del magazzino"),
  bullet("Registrazione dei camion e generazione delle etichette QR code (PNG)"),
  bullet("Configurazione dell'IP del MiR, dell'IP del Raspberry Pi e delle credenziali"),
  bullet("Storico delle operazioni e delle missioni"),
  space(),
  body("Il login è protetto da rate limiting: dopo 5 tentativi falliti dallo stesso IP si deve aspettare 60 secondi."),

  pageBreak(),

  // 10. CONCLUSIONI
  h1("10. Conclusioni"),
  body("Il sistema funziona end-to-end con il robot fisico. Siamo riusciti a integrare tutti i componenti (robot, Raspberry Pi, webcam, tablet, server) su una rete locale senza bisogno di connessione internet o di servizi cloud."),
  space(),
  body("La parte che ci ha impegnato di più è stata sicuramente la sincronizzazione tra i vari thread e la gestione degli stati del robot. Avere due pagine indipendenti (balia e scarico) che devono aggiornarsi in modo coordinato richiede una certa attenzione ai casi limite."),
  space(),
  body("Il risultato è un sistema che, anche se sviluppato in ambito scolastico, ha tutti i requisiti di una soluzione reale: autenticazione, tracciabilità completa, interfaccia operativa e integrazione hardware vera con il MiR100."),
  space(),

  h2("Tecnologie utilizzate"),
  bullet("Python 3.11 + Flask"),
  bullet("SQLite + SQLAlchemy"),
  bullet("OpenCV + face_recognition"),
  bullet("MiR100 REST API v2.0.0"),
  bullet("Raspberry Pi + Flask MJPEG stream"),
  bullet("Server-Sent Events (SSE)"),
  bullet("HTML5 + CSS3 + JavaScript"),
];

// ─── DOCUMENTO ────────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: { config: numberingConfig },
  styles: {
    default: {
      document: { run: { font: "Calibri", size: 22, color: DARK } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Calibri", color: BLUE },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Calibri", color: "2563EB" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PG_W, height: PG_H },
        margin: { top: MAR, right: MAR, bottom: MAR, left: MAR },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1", space: 8 } },
          spacing: { before: 0, after: 120 },
          children: [
            new TextRun({ text: "MiR100 — Documentazione Progetto", size: 18, color: MUTED, font: "Calibri" }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CBD5E1", space: 8 } },
          spacing: { before: 120, after: 0 },
          children: [
            new TextRun({ text: "ITIS · Anno scolastico 2025/2026  ·  Pagina ", size: 18, color: MUTED, font: "Calibri" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, color: MUTED, font: "Calibri" }),
            new TextRun({ text: " di ", size: 18, color: MUTED, font: "Calibri" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: MUTED, font: "Calibri" }),
          ],
        })],
      }),
    },
    children: content,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("mir100_doc_semplice.docx", buf);
  console.log("✅  mir100_doc_semplice.docx generato");
});
