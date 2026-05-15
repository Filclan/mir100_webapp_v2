# MiR100 — Sistema Gestione Magazzino v2

## Struttura

```
mir100_webapp_v2/
├── app.py                    ← Server Flask principale
├── magazzino.db              ← Database SQLite (auto-generato al primo avvio)
├── requirements.txt
├── moduli/
│   ├── database.py           ← SQLite locale (sostituisce Supabase)
│   ├── camera_manager.py     ← Stream Raspberry Pi (copiare dal vecchio progetto)
│   ├── qr_scanner.py         ← Scanner QR (copiare dal vecchio progetto)
│   └── face_module.py        ← Face recognition (copiare dal vecchio progetto)
└── templates/
    ├── admin_login.html       ← Login admin
    ├── admin_dashboard.html   ← Pannello admin completo
    ├── tablet_balia.html      ← Pannello operatore baia camion
    └── tablet_scarico.html    ← Pannello operatore scarico
```

## Installazione

```bash
pip install -r requirements.txt
```

## Avvio

```bash
python app.py
```

## URL

| Pannello              | URL                          |
|-----------------------|------------------------------|
| Operatore Balia       | http://[IP]:5000/            |
| Operatore Scarico     | http://[IP]:5000/scarico     |
| Admin Login           | http://[IP]:5000/admin/login |
| Admin Dashboard       | http://[IP]:5000/admin       |

## Credenziali default admin

- **Username:** `admin`
- **Password:** `admin1234`

⚠️ Cambiare la password dal pannello Config → Cambia password dopo il primo accesso.

## Database

Il file `magazzino.db` viene creato automaticamente nella cartella del progetto.
Al primo avvio vengono inseriti:
- Admin di default
- 4 destinazioni di esempio
- 4 operatori di esempio (senza face encoding)
- Configurazione IP di default

## Configurazione IP

Dal pannello Admin → Configurazione:
- **IP Raspberry**: IP del Raspberry Pi con la telecamera
- **Porta stream**: Porta dello stream MJPEG (default 8080)
- **IP MiR100**: IP del robot
- **Soglia face recognition**: 0.55 (abbassare per essere più restrittivi)

## Migrazione da Supabase

Il vecchio progetto usava Supabase (PostgreSQL remoto).
La nuova versione usa SQLite locale — stesso schema, zero dipendenze cloud.

Per importare dati esistenti dal vecchio DB, esportare le tabelle come CSV
e importarle con SQLite Browser o uno script Python ad hoc.

## Registrazione face encoding

1. Admin Dashboard → Operatori
2. Clicca **👤 Face** sull'operatore
3. Il browser chiede accesso alla webcam
4. Posiziona l'operatore davanti alla telecamera del dispositivo admin
5. Clicca **Acquisisci volto**
6. Il sistema estrae e salva l'encoding automaticamente
