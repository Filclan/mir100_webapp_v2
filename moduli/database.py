"""
moduli/database.py
Database SQLite locale — sostituisce Supabase.
Usa SQLAlchemy Core per query tipizzate e thread-safe.
"""

import os
import json
import bcrypt
from datetime import datetime
from sqlalchemy import (
    create_engine, MetaData, Table, Column, Integer, Text, Numeric,
    Boolean, DateTime, Float, ForeignKey, select, update, insert, delete
)
from sqlalchemy.pool import StaticPool

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "magazzino.db")
_engine = None
meta = MetaData()

# ─── SCHEMA ──────────────────────────────────────────────────────────────────

t_destinazioni = Table("destinazioni", meta,
    Column("id",          Integer, primary_key=True, autoincrement=True),
    Column("nome",        Text, nullable=False),
    Column("descrizione", Text),
    Column("tipo",        Text, nullable=False, default="smistamento"),
    Column("attiva",      Boolean, nullable=False, default=True),
    Column("created_at",  DateTime, default=datetime.now),
)

t_operatori = Table("operatori", meta,
    Column("id",              Integer, primary_key=True, autoincrement=True),
    Column("nome",            Text, nullable=False),
    Column("cognome",         Text, nullable=False),
    Column("ruolo",           Text, nullable=False, default="operatore"),
    Column("destinazione_id", Integer, ForeignKey("destinazioni.id")),
    Column("face_encoding",   Text),   # JSON array di 128 float
    Column("attivo",          Boolean, nullable=False, default=True),
    Column("created_at",      DateTime, default=datetime.now),
)

t_admins = Table("admins", meta,
    Column("id",           Integer, primary_key=True, autoincrement=True),
    Column("username",     Text, nullable=False, unique=True),
    Column("password_hash",Text, nullable=False),
    Column("nome",         Text),
    Column("created_at",   DateTime, default=datetime.now),
)

t_camion = Table("camion", meta,
    Column("id",          Integer, primary_key=True, autoincrement=True),
    Column("targa",       Text, nullable=False),
    Column("qr_code",     Text, nullable=False, unique=True),
    Column("fornitore",   Text),
    Column("data_arrivo", Text, default=lambda: datetime.now().date().isoformat()),
    Column("stato",       Text, nullable=False, default="in_arrivo"),
    Column("created_at",  DateTime, default=datetime.now),
)

t_pacchi = Table("pacchi", meta,
    Column("id",              Integer, primary_key=True, autoincrement=True),
    Column("qr_code",         Text, nullable=False, unique=True),
    Column("camion_id",       Integer, ForeignKey("camion.id")),
    Column("destinazione_id", Integer, ForeignKey("destinazioni.id")),
    Column("descrizione",     Text),
    Column("peso_kg",         Numeric(6, 2)),
    Column("stato",           Text, nullable=False, default="atteso"),
    Column("updated_at",      DateTime, default=datetime.now),
    Column("created_at",      DateTime, default=datetime.now),
)

t_operazioni_scarico = Table("operazioni_scarico", meta,
    Column("id",           Integer, primary_key=True, autoincrement=True),
    Column("operatore_id", Integer, ForeignKey("operatori.id")),
    Column("camion_id",    Integer, ForeignKey("camion.id")),
    Column("stato",        Text, nullable=False, default="in_corso"),
    Column("inizio_ts",    DateTime, default=datetime.now),
    Column("fine_ts",      DateTime),
)

t_operazioni_consegna = Table("operazioni_consegna", meta,
    Column("id",              Integer, primary_key=True, autoincrement=True),
    Column("operatore_id",    Integer, ForeignKey("operatori.id")),
    Column("pacco_id",        Integer, ForeignKey("pacchi.id")),
    Column("destinazione_id", Integer, ForeignKey("destinazioni.id")),
    Column("timestamp",       DateTime, default=datetime.now),
    Column("esito",           Text, nullable=False, default="consegnato"),
)

t_missioni_mir = Table("missioni_mir", meta,
    Column("id",              Integer, primary_key=True, autoincrement=True),
    Column("mir_id",          Text),
    Column("tipo",            Text, nullable=False, default="consegna"),
    Column("destinazione_id", Integer, ForeignKey("destinazioni.id")),
    Column("stato",           Text, nullable=False, default="pianificata"),
    Column("inizio_ts",       DateTime, default=datetime.now),
    Column("fine_ts",         DateTime),
)

t_config = Table("config", meta,
    Column("chiave", Text, primary_key=True),
    Column("valore", Text),
)


# ─── ENGINE ───────────────────────────────────────────────────────────────────

def get_engine():
    global _engine
    if _engine is None:
        _engine = create_engine(
            f"sqlite:///{DB_PATH}",
            connect_args={"check_same_thread": False},
            poolclass=StaticPool,
        )
        meta.create_all(_engine)
        _seed_default_data()
    return _engine


def conn():
    return get_engine().connect()


# ─── SEED ─────────────────────────────────────────────────────────────────────

def _seed_default_data():
    with get_engine().begin() as c:
        # Admin di default
        existing = c.execute(select(t_admins).where(t_admins.c.username == "admin")).fetchone()
        if not existing:
            pw_hash = bcrypt.hashpw(b"admin1234", bcrypt.gensalt()).decode()
            c.execute(insert(t_admins).values(username="admin", password_hash=pw_hash, nome="Amministratore"))

        # Config di default
        defaults = {
            "raspberry_ip":              "192.168.12.242",
            "raspberry_port":            "8081",
            "mir_ip":                    "192.168.12.20",
            "mir_auth_user":             "itisdelpozzo",
            "mir_auth_pass":             "itisdelpozzo",
            "mir_posizione_partenza":    "area B1",
            "mir_posizione_destinazione":"area A1",
            "face_soglia":               "0.55",
            "qr_timeout_sec":            "60",
        }
        for k, v in defaults.items():
            ex = c.execute(select(t_config).where(t_config.c.chiave == k)).fetchone()
            if not ex:
                c.execute(insert(t_config).values(chiave=k, valore=v))

        # Destinazioni di esempio
        ex = c.execute(select(t_destinazioni)).fetchone()
        if not ex:
            c.execute(insert(t_destinazioni).values([
                {"nome": "Zona A — Scaffali",   "descrizione": "Area scaffalatura prodotti secchi", "tipo": "smistamento"},
                {"nome": "Zona B — Frigo",       "descrizione": "Area refrigerata",                 "tipo": "smistamento"},
                {"nome": "Zona C — Spedizioni",  "descrizione": "Area imballaggio e spedizione",     "tipo": "smistamento"},
                {"nome": "Zona scarico camion",  "descrizione": "Baia di arrivo camion",             "tipo": "scarico"},
            ]))

        # Operatori di esempio (senza face_encoding)
        ex = c.execute(select(t_operatori)).fetchone()
        if not ex:
            c.execute(insert(t_operatori).values([
                {"nome": "Mario",    "cognome": "Rossi",   "ruolo": "operatore",   "destinazione_id": 4},
                {"nome": "Luca",     "cognome": "Bianchi", "ruolo": "operatore",   "destinazione_id": 1},
                {"nome": "Anna",     "cognome": "Verdi",   "ruolo": "supervisore", "destinazione_id": None},
                {"nome": "Giuseppe", "cognome": "Marino",  "ruolo": "operatore",   "destinazione_id": 2},
            ]))


# ─── CONFIG ───────────────────────────────────────────────────────────────────

def get_config() -> dict:
    with conn() as c:
        rows = c.execute(select(t_config)).fetchall()
        return {r.chiave: r.valore for r in rows}

def set_config(chiave: str, valore: str):
    with get_engine().begin() as c:
        ex = c.execute(select(t_config).where(t_config.c.chiave == chiave)).fetchone()
        if ex:
            c.execute(update(t_config).where(t_config.c.chiave == chiave).values(valore=valore))
        else:
            c.execute(insert(t_config).values(chiave=chiave, valore=valore))


# ─── ADMIN AUTH ───────────────────────────────────────────────────────────────

def verifica_admin(username: str, password: str) -> dict | None:
    with conn() as c:
        row = c.execute(select(t_admins).where(t_admins.c.username == username)).fetchone()
        if not row:
            return None
        if bcrypt.checkpw(password.encode(), row.password_hash.encode()):
            return {"id": row.id, "username": row.username, "nome": row.nome}
        return None

def cambia_password_admin(admin_id: int, nuova_password: str):
    pw_hash = bcrypt.hashpw(nuova_password.encode(), bcrypt.gensalt()).decode()
    with get_engine().begin() as c:
        c.execute(update(t_admins).where(t_admins.c.id == admin_id).values(password_hash=pw_hash))


# ─── OPERATORI ────────────────────────────────────────────────────────────────

def get_tutti_operatori(solo_attivi: bool = True) -> list:
    with conn() as c:
        q = select(t_operatori, t_destinazioni.c.nome.label("dest_nome")).outerjoin(
            t_destinazioni, t_operatori.c.destinazione_id == t_destinazioni.c.id
        )
        if solo_attivi:
            q = q.where(t_operatori.c.attivo == True)
        rows = c.execute(q).fetchall()
        return [dict(r._mapping) for r in rows]

def get_operatore_by_id(op_id: int) -> dict | None:
    with conn() as c:
        row = c.execute(select(t_operatori).where(t_operatori.c.id == op_id)).fetchone()
        return dict(row._mapping) if row else None

def crea_operatore(nome: str, cognome: str, ruolo: str, destinazione_id: int | None) -> int:
    with get_engine().begin() as c:
        res = c.execute(insert(t_operatori).values(
            nome=nome, cognome=cognome, ruolo=ruolo,
            destinazione_id=destinazione_id, attivo=True,
            created_at=datetime.now()
        ))
        return res.inserted_primary_key[0]

def aggiorna_operatore(op_id: int, **kwargs):
    with get_engine().begin() as c:
        c.execute(update(t_operatori).where(t_operatori.c.id == op_id).values(**kwargs))

def salva_face_encoding(op_id: int, encoding: list):
    with get_engine().begin() as c:
        c.execute(update(t_operatori).where(t_operatori.c.id == op_id).values(
            face_encoding=json.dumps(encoding)
        ))

def get_operatore_by_encoding(encoding_list: list, soglia: float = 0.55) -> dict | None:
    import face_recognition
    import numpy as np
    operatori = get_tutti_operatori()
    enc_input = np.array(encoding_list)
    for op in operatori:
        enc_db = op.get("face_encoding")
        if not enc_db:
            continue
        enc_db_arr = np.array(json.loads(enc_db))
        dist = face_recognition.face_distance([enc_db_arr], enc_input)[0]
        if dist < soglia:
            return op
    return None


# ─── DESTINAZIONI ─────────────────────────────────────────────────────────────

def get_destinazioni(solo_attive: bool = True) -> list:
    with conn() as c:
        q = select(t_destinazioni)
        if solo_attive:
            q = q.where(t_destinazioni.c.attiva == True)
        return [dict(r._mapping) for r in c.execute(q).fetchall()]

def crea_destinazione(nome: str, descrizione: str, tipo: str) -> int:
    with get_engine().begin() as c:
        res = c.execute(insert(t_destinazioni).values(
            nome=nome, descrizione=descrizione, tipo=tipo, attiva=True, created_at=datetime.now()
        ))
        return res.inserted_primary_key[0]

def aggiorna_destinazione(dest_id: int, **kwargs):
    with get_engine().begin() as c:
        c.execute(update(t_destinazioni).where(t_destinazioni.c.id == dest_id).values(**kwargs))


# ─── CAMION ───────────────────────────────────────────────────────────────────

def get_tutti_camion() -> list:
    with conn() as c:
        rows = c.execute(select(t_camion).order_by(t_camion.c.created_at.desc())).fetchall()
        return [dict(r._mapping) for r in rows]

def crea_camion(targa: str, fornitore: str, data_arrivo: str) -> dict:
    import uuid
    camion_id = f"CAM-{uuid.uuid4().hex[:6].upper()}"
    qr_payload = json.dumps({"tipo": "camion", "id": camion_id, "targa": targa, "fornitore": fornitore})
    with get_engine().begin() as c:
        res = c.execute(insert(t_camion).values(
            targa=targa, qr_code=qr_payload, fornitore=fornitore,
            data_arrivo=data_arrivo, stato="in_arrivo", created_at=datetime.now()
        ))
        db_id = res.inserted_primary_key[0]
    return {"db_id": db_id, "camion_id": camion_id, "qr_code": qr_payload}

def aggiorna_stato_camion(camion_db_id: int, stato: str):
    with get_engine().begin() as c:
        c.execute(update(t_camion).where(t_camion.c.id == camion_db_id).values(stato=stato))


# ─── PACCHI ───────────────────────────────────────────────────────────────────

def get_pacchi_by_camion(camion_qr_id: str) -> list:
    """Dato l'id logico del camion (es. 'CAM-XXXX'), restituisce i pacchi attesi."""
    with conn() as c:
        camion_rows = c.execute(select(t_camion)).fetchall()
        camion_record = None
        for row in camion_rows:
            try:
                parsed = json.loads(row.qr_code)
                if parsed.get("id") == camion_qr_id:
                    camion_record = row
                    break
            except Exception:
                if row.qr_code == camion_qr_id:
                    camion_record = row
                    break
        if not camion_record:
            return []
        rows = c.execute(
            select(t_pacchi, t_destinazioni.c.nome.label("dest_nome"))
            .outerjoin(t_destinazioni, t_pacchi.c.destinazione_id == t_destinazioni.c.id)
            .where(t_pacchi.c.camion_id == camion_record.id)
            .where(t_pacchi.c.stato == "atteso")
        ).fetchall()
        result = []
        for r in rows:
            d = dict(r._mapping)
            d["destinazioni"] = {"nome": d.pop("dest_nome", "—")}
            result.append(d)
        return result

def get_pacco_by_qr(qr_code: str) -> dict | None:
    with conn() as c:
        row = c.execute(
            select(t_pacchi, t_destinazioni.c.nome.label("dest_nome"))
            .outerjoin(t_destinazioni, t_pacchi.c.destinazione_id == t_destinazioni.c.id)
            .where(t_pacchi.c.qr_code == qr_code)
        ).fetchone()
        if not row:
            return None
        d = dict(row._mapping)
        d["destinazioni"] = {"nome": d.pop("dest_nome", "—")}
        return d

def get_pacchi_by_destinazione(destinazione_id: int) -> list:
    with conn() as c:
        rows = c.execute(
            select(t_pacchi, t_camion.c.targa, t_camion.c.fornitore)
            .outerjoin(t_camion, t_pacchi.c.camion_id == t_camion.c.id)
            .where(t_pacchi.c.destinazione_id == destinazione_id)
            .where(t_pacchi.c.stato == "in_transito")
        ).fetchall()
        return [dict(r._mapping) for r in rows]

def crea_pacco(camion_db_id: int, destinazione_id: int, descrizione: str, peso_kg: float) -> dict:
    import uuid
    pacco_id = f"PKG-{uuid.uuid4().hex[:6].upper()}"
    with conn() as c:
        dest = c.execute(select(t_destinazioni).where(t_destinazioni.c.id == destinazione_id)).fetchone()
        dest_nome = dest.nome if dest else "—"
    qr_payload = json.dumps({
        "tipo": "pacco", "id": pacco_id,
        "destinazione_id": destinazione_id, "destinazione_nome": dest_nome
    })
    with get_engine().begin() as c:
        res = c.execute(insert(t_pacchi).values(
            qr_code=qr_payload, camion_id=camion_db_id, destinazione_id=destinazione_id,
            descrizione=descrizione, peso_kg=peso_kg, stato="atteso",
            created_at=datetime.now(), updated_at=datetime.now()
        ))
        db_id = res.inserted_primary_key[0]
    return {"db_id": db_id, "pacco_id": pacco_id, "qr_code": qr_payload}

def aggiorna_stato_pacco(pacco_qr: str, nuovo_stato: str) -> bool:
    with get_engine().begin() as c:
        res = c.execute(
            update(t_pacchi)
            .where(t_pacchi.c.qr_code == pacco_qr)
            .values(stato=nuovo_stato, updated_at=datetime.now())
        )
        return res.rowcount > 0

def elimina_pacco(pacco_id: int):
    with get_engine().begin() as c:
        c.execute(delete(t_pacchi).where(t_pacchi.c.id == pacco_id))


# ─── OPERAZIONI ───────────────────────────────────────────────────────────────

def crea_operazione_scarico(operatore_id: int, camion_id: int) -> int:
    with get_engine().begin() as c:
        res = c.execute(insert(t_operazioni_scarico).values(
            operatore_id=operatore_id, camion_id=camion_id,
            stato="in_corso", inizio_ts=datetime.now()
        ))
        return res.inserted_primary_key[0]

def chiudi_operazione_scarico(op_id: int):
    with get_engine().begin() as c:
        c.execute(update(t_operazioni_scarico).where(t_operazioni_scarico.c.id == op_id).values(
            stato="completata", fine_ts=datetime.now()
        ))

def get_storico_scarichi(limit: int = 50) -> list:
    with conn() as c:
        rows = c.execute(
            select(
                t_operazioni_scarico,
                t_operatori.c.nome.label("op_nome"),
                t_operatori.c.cognome.label("op_cognome"),
                t_camion.c.targa,
                t_camion.c.fornitore,
            )
            .outerjoin(t_operatori, t_operazioni_scarico.c.operatore_id == t_operatori.c.id)
            .outerjoin(t_camion, t_operazioni_scarico.c.camion_id == t_camion.c.id)
            .order_by(t_operazioni_scarico.c.inizio_ts.desc())
            .limit(limit)
        ).fetchall()
        return [dict(r._mapping) for r in rows]

def get_storico_missioni(limit: int = 50) -> list:
    with conn() as c:
        rows = c.execute(
            select(t_missioni_mir, t_destinazioni.c.nome.label("dest_nome"))
            .outerjoin(t_destinazioni, t_missioni_mir.c.destinazione_id == t_destinazioni.c.id)
            .order_by(t_missioni_mir.c.inizio_ts.desc())
            .limit(limit)
        ).fetchall()
        return [dict(r._mapping) for r in rows]

def crea_missione_mir(destinazione_id: int | None = None, tipo: str = "consegna") -> int:
    with get_engine().begin() as c:
        res = c.execute(insert(t_missioni_mir).values(
            tipo=tipo, destinazione_id=destinazione_id,
            stato="pianificata", inizio_ts=datetime.now()
        ))
        return res.inserted_primary_key[0]

def aggiorna_stato_missione(missione_id: int, stato: str):
    with get_engine().begin() as c:
        vals = {"stato": stato}
        if stato in ("completata", "annullata", "errore"):
            vals["fine_ts"] = datetime.now()
        c.execute(update(t_missioni_mir).where(t_missioni_mir.c.id == missione_id).values(**vals))

def get_dashboard_stats() -> dict:
    with conn() as c:
        oggi = datetime.now().date().isoformat()
        pacchi_oggi = c.execute(
            select(t_pacchi).where(t_pacchi.c.updated_at >= oggi)
        ).fetchall()
        scarichi_oggi = c.execute(
            select(t_operazioni_scarico).where(t_operazioni_scarico.c.inizio_ts >= oggi)
        ).fetchall()
        missioni_oggi = c.execute(
            select(t_missioni_mir).where(t_missioni_mir.c.inizio_ts >= oggi)
        ).fetchall()
        camion_attivi = c.execute(
            select(t_camion).where(t_camion.c.stato == "in_arrivo")
        ).fetchall()
        pacchi_transito = c.execute(
            select(t_pacchi).where(t_pacchi.c.stato == "in_transito")
        ).fetchall()
        return {
            "pacchi_oggi":      len(pacchi_oggi),
            "scarichi_oggi":    len(scarichi_oggi),
            "missioni_oggi":    len(missioni_oggi),
            "camion_attivi":    len(camion_attivi),
            "pacchi_transito":  len(pacchi_transito),
        }
