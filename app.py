"""
app.py — Server Flask principale MiR100 v2
"""

import os
import sys
import json
import threading
import queue
import base64
import hashlib
import time
import io
from collections import defaultdict
from datetime import datetime, timedelta
from decimal import Decimal
from functools import wraps

import requests as _req

from flask import (
    Flask, render_template, Response, jsonify, request,
    session, redirect, url_for, stream_with_context
)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from moduli.database import (
    get_engine, get_config, set_config,
    verifica_admin, cambia_password_admin,
    get_tutti_operatori, get_operatore_by_id, crea_operatore,
    aggiorna_operatore, salva_face_encoding, get_operatore_by_encoding,
    get_destinazioni, crea_destinazione, aggiorna_destinazione,
    get_tutti_camion, crea_camion, aggiorna_stato_camion,
    get_pacchi_by_camion, get_pacco_by_qr, crea_pacco,
    aggiorna_stato_pacco, elimina_pacco,
    crea_operazione_scarico, chiudi_operazione_scarico,
    get_storico_scarichi, get_storico_missioni,
    crea_missione_mir, aggiorna_stato_missione, get_dashboard_stats,
    get_pacchi_by_destinazione,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or "mir100-magazzino-2026-secret"
if not os.environ.get("SECRET_KEY"):
    print("[WARN] SECRET_KEY non impostata — usare variabile d'ambiente in produzione.")
app.permanent_session_lifetime = timedelta(hours=8)


from flask.json.provider import DefaultJSONProvider

class _JsonProvider(DefaultJSONProvider):
    def default(self, obj):
        if isinstance(obj, Decimal):
            return float(obj)
        if isinstance(obj, datetime):
            return obj.isoformat()
        return super().default(obj)

app.json_provider_class = _JsonProvider
app.json = _JsonProvider(app)

# ─── STATO GLOBALE ────────────────────────────────────────────────────────────

stato = {
    "fase": 0,
    "operatore": None,
    "camion": None,
    "pacchi_lista": [],
    "pacchi_scansionati": [],
    "busy": False,
    "camera_attiva": False,
}

_sse_queue  = queue.Queue(maxsize=200)
_stato_lock = threading.RLock()
cam = scanner = rf = None


_cam_broadcaster = None   # MJPEGBroadcaster, inizializzato in _init_hardware()

# Sessione MiR attiva: viene popolata quando il robot arriva a destinazione
# e consumata quando tutti i pacchi sono confermati (o via endpoint manuale).
_mir_sessione: dict = {}   # keys: mir_url, headers, missions, pos_part, pos_dest


def _init_hardware():
    global cam, scanner, rf, _cam_broadcaster
    try:
        cfg = get_config()
        rpi_ip   = cfg.get("raspberry_ip",  "192.168.12.242")
        rpi_port = cfg.get("raspberry_port", "8081")
        stream_url = f"http://{rpi_ip}:{rpi_port}/stream"
        from moduli.camera_manager import MJPEGBroadcaster, CameraManager
        from moduli.qr_scanner import QRScanner
        from moduli.face_module import RiconoscimentoFacciale
        _cam_broadcaster = MJPEGBroadcaster(stream_url)   # unica connessione al Pi
        cam     = CameraManager(_cam_broadcaster)          # decoder JPEG→numpy
        scanner = QRScanner(camera_manager=cam)
        rf      = RiconoscimentoFacciale(camera_manager=cam)
        print(f"[HW] Camera: {stream_url}")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"[HW] Errore init hardware: {e}")


# ─── SSE ──────────────────────────────────────────────────────────────────────

def push_event(tipo, payload):
    msg = f"event: {tipo}\ndata: {json.dumps(payload, default=str)}\n\n"
    try:
        _sse_queue.put_nowait(msg)
    except queue.Full:
        print(f"[SSE] Queue piena — evento '{tipo}' perso")


def push_stato():
    push_event("stato", {k: stato[k] for k in stato})


def push_notifica(tipo, msg):
    push_event("notifica", {
        "tipo": tipo,
        "msg": msg,
        "ts": datetime.now().strftime("%H:%M:%S")
    })


# ─── HELPER FACE ENCODING ─────────────────────────────────────────────────────

def _prepara_operatori_per_face(operatori_raw):
    """
    Converte face_encoding da stringa JSON salvata nel DB
    a lista di float come si aspetta face_recognition.
    """
    result = []
    for o in operatori_raw:
        o = dict(o)
        fe = o.get("face_encoding")
        if fe and isinstance(fe, str):
            try:
                o["face_encoding"] = json.loads(fe)
            except Exception:
                o["face_encoding"] = None
        result.append(o)
    return result


# ─── HELPER SERIALIZZAZIONE PACCHI ────────────────────────────────────────────

def _ser_pacchi(pacchi):
    result = []
    for p in pacchi:
        try:
            qr  = json.loads(str(p.get("qr_code", "")))
            pid = qr.get("id", "—")
        except Exception:
            pid = p.get("qr_code", "—")
            print(f"[WARN] QR non parsabile come JSON: {p.get('qr_code')!r}")
        dest = p.get("destinazioni") or {}
        peso = p.get("peso_kg")
        result.append({
            "qr_code":      p["qr_code"],
            "id":           pid,
            "destinazione": dest.get("nome", "—"),
            "descrizione":  p.get("descrizione", ""),
            "peso_kg":      float(peso) if peso is not None else None,
        })
    return result


# ─── RATE LIMITING LOGIN ──────────────────────────────────────────────────────

_login_attempts: dict = defaultdict(list)
_LOGIN_MAX_ATTEMPTS   = 5
_LOGIN_WINDOW_SEC     = 60

def _check_rate_limit(ip: str) -> bool:
    now = time.time()
    _login_attempts[ip] = [t for t in _login_attempts[ip] if now - t < _LOGIN_WINDOW_SEC]
    if len(_login_attempts[ip]) >= _LOGIN_MAX_ATTEMPTS:
        return False
    _login_attempts[ip].append(now)
    return True


# ─── HELPER AGGIORNAMENTO STATO THREAD-SAFE ───────────────────────────────────

def _update_stato(**kwargs):
    with _stato_lock:
        stato.update(kwargs)


# ─── DECORATORE ADMIN ─────────────────────────────────────────────────────────

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("admin_id"):
            if request.is_json:
                return jsonify({"ok": False, "msg": "Non autorizzato"}), 401
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated


# ─── AUTH ADMIN ───────────────────────────────────────────────────────────────

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        if not _check_rate_limit(request.remote_addr):
            msg = "Troppi tentativi. Riprova tra un minuto."
            if request.is_json:
                return jsonify({"ok": False, "msg": msg}), 429
            return render_template("admin_login.html", errore=msg), 429
        data = request.get_json() or request.form
        admin = verifica_admin(data.get("username", ""), data.get("password", ""))
        if admin:
            session.permanent = True
            session["admin_id"]   = admin["id"]
            session["admin_nome"] = admin["nome"]
            if request.is_json:
                return jsonify({"ok": True})
            return redirect(url_for("admin_dashboard"))
        if request.is_json:
            return jsonify({"ok": False, "msg": "Credenziali errate"}), 401
        return render_template("admin_login.html", errore="Credenziali errate")
    return render_template("admin_login.html", errore=None)


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


# ─── ADMIN DASHBOARD ──────────────────────────────────────────────────────────

@app.route("/admin")
@app.route("/admin/")
@admin_required
def admin_dashboard():
    return render_template("admin_dashboard.html", admin_nome=session.get("admin_nome"))


@app.route("/admin/api/stats")
@admin_required
def api_stats():
    return jsonify(get_dashboard_stats())


# ─── ADMIN — QR GENERATOR (lato server, niente CDN) ──────────────────────────

@app.route("/admin/api/qr")
@admin_required
def api_qr_generate():
    """Genera un QR code come PNG base64 dal testo passato come ?data=..."""
    import qrcode
    data = request.args.get("data", "")
    if not data:
        return "missing data", 400
    qr = qrcode.QRCode(
        version=None,
        box_size=8,
        border=2,
        error_correction=qrcode.constants.ERROR_CORRECT_M
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return f"data:image/png;base64,{b64}", 200, {"Content-Type": "text/plain"}


# ─── ADMIN — OPERATORI ────────────────────────────────────────────────────────

@app.route("/admin/api/operatori")
@admin_required
def api_operatori_list():
    return jsonify(get_tutti_operatori(solo_attivi=False))


@app.route("/admin/api/operatori", methods=["POST"])
@admin_required
def api_operatori_crea():
    d = request.get_json()
    op_id = crea_operatore(
        nome=d["nome"],
        cognome=d["cognome"],
        ruolo=d.get("ruolo", "operatore"),
        destinazione_id=d.get("destinazione_id") or None
    )
    return jsonify({"ok": True, "id": op_id})


@app.route("/admin/api/operatori/<int:op_id>", methods=["PATCH"])
@admin_required
def api_operatori_aggiorna(op_id):
    d = request.get_json()
    allowed = {"nome", "cognome", "ruolo", "destinazione_id", "attivo"}
    aggiorna_operatore(op_id, **{k: v for k, v in d.items() if k in allowed})
    return jsonify({"ok": True})


@app.route("/admin/api/operatori/<int:op_id>/reset_face", methods=["POST"])
@admin_required
def api_reset_face(op_id):
    aggiorna_operatore(op_id, face_encoding=None)
    return jsonify({"ok": True})


@app.route("/admin/api/operatori/<int:op_id>/registra_face", methods=["POST"])
@admin_required
def api_registra_face(op_id):
    """
    Riceve un frame JPEG in base64 dal browser (webcam del dispositivo admin),
    estrae il face encoding e lo salva nel DB come stringa JSON.
    """
    import face_recognition
    import numpy as np
    from PIL import Image

    data = request.get_json()
    img_b64 = data.get("frame_b64", "")
    if not img_b64:
        return jsonify({"ok": False, "msg": "Nessun frame ricevuto"})
    try:
        img_bytes = base64.b64decode(img_b64.split(",")[-1])
        img   = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        frame = np.array(img)
        locs  = face_recognition.face_locations(frame, model="hog")
        if not locs:
            return jsonify({"ok": False, "msg": "Nessun volto rilevato nel frame"})
        encs = face_recognition.face_encodings(frame, locs)
        if not encs:
            return jsonify({"ok": False, "msg": "Impossibile estrarre encoding"})
        salva_face_encoding(op_id, encs[0].tolist())
        return jsonify({"ok": True, "msg": "Encoding salvato"})
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"ok": False, "msg": str(e)})


# ─── ADMIN — DESTINAZIONI ─────────────────────────────────────────────────────

@app.route("/admin/api/destinazioni")
@admin_required
def api_destinazioni_list():
    return jsonify(get_destinazioni(solo_attive=False))


@app.route("/admin/api/destinazioni", methods=["POST"])
@admin_required
def api_destinazioni_crea():
    d = request.get_json()
    did = crea_destinazione(d["nome"], d.get("descrizione", ""), d.get("tipo", "smistamento"))
    return jsonify({"ok": True, "id": did})


@app.route("/admin/api/destinazioni/<int:dest_id>", methods=["PATCH"])
@admin_required
def api_destinazioni_aggiorna(dest_id):
    d = request.get_json()
    allowed = {"nome", "descrizione", "tipo", "attiva"}
    aggiorna_destinazione(dest_id, **{k: v for k, v in d.items() if k in allowed})
    return jsonify({"ok": True})


# ─── ADMIN — CAMION & PACCHI ──────────────────────────────────────────────────

@app.route("/admin/api/camion")
@admin_required
def api_camion_list():
    return jsonify(get_tutti_camion())


@app.route("/admin/api/camion", methods=["POST"])
@admin_required
def api_camion_crea():
    d = request.get_json()
    result = crea_camion(
        d["targa"],
        d.get("fornitore", ""),
        d.get("data_arrivo", datetime.now().date().isoformat())
    )
    return jsonify({"ok": True, **result})


@app.route("/admin/api/camion/<int:camion_id>/pacchi")
@admin_required
def api_pacchi_by_camion(camion_id):
    from moduli.database import conn, t_pacchi, t_destinazioni
    from sqlalchemy import select
    with conn() as c:
        rows = c.execute(
            select(t_pacchi, t_destinazioni.c.nome.label("dest_nome"))
            .outerjoin(t_destinazioni, t_pacchi.c.destinazione_id == t_destinazioni.c.id)
            .where(t_pacchi.c.camion_id == camion_id)
        ).fetchall()
    return jsonify([dict(r._mapping) for r in rows])


@app.route("/admin/api/pacchi", methods=["POST"])
@admin_required
def api_pacco_crea():
    d = request.get_json()
    result = crea_pacco(
        camion_db_id=d["camion_id"],
        destinazione_id=d["destinazione_id"],
        descrizione=d.get("descrizione", ""),
        peso_kg=d.get("peso_kg", 0)
    )
    return jsonify({"ok": True, **result})


@app.route("/admin/api/pacchi/<int:pacco_id>", methods=["DELETE"])
@admin_required
def api_pacco_elimina(pacco_id):
    elimina_pacco(pacco_id)
    return jsonify({"ok": True})


@app.route("/admin/api/pacchi/qr/<path:qr_code>", methods=["PATCH"])
@admin_required
def api_pacco_stato(qr_code):
    d = request.get_json()
    aggiorna_stato_pacco(qr_code, d["stato"])
    return jsonify({"ok": True})


@app.route("/admin/qr/camion/<int:camion_id>")
@admin_required
def qr_camion_stampa(camion_id):
    from moduli.database import conn, t_camion, t_pacchi, t_destinazioni
    from sqlalchemy import select
    with conn() as c:
        camion_row = c.execute(
            select(t_camion).where(t_camion.c.id == camion_id)
        ).fetchone()
        if not camion_row:
            return "Camion non trovato", 404
        pacchi = c.execute(
            select(t_pacchi, t_destinazioni.c.nome.label("dest_nome"))
            .outerjoin(t_destinazioni, t_pacchi.c.destinazione_id == t_destinazioni.c.id)
            .where(t_pacchi.c.camion_id == camion_id)
        ).fetchall()
    return render_template(
        "qr_stampa.html",
        camion=dict(camion_row._mapping),
        pacchi=[dict(p._mapping) for p in pacchi]
    )


# ─── ADMIN — CONFIG ───────────────────────────────────────────────────────────

@app.route("/admin/api/config")
@admin_required
def api_config_get():
    return jsonify(get_config())


@app.route("/admin/api/config", methods=["POST"])
@admin_required
def api_config_set():
    d = request.get_json()
    for k, v in d.items():
        set_config(k, str(v))
    return jsonify({"ok": True})


@app.route("/admin/api/config/password", methods=["POST"])
@admin_required
def api_cambia_password():
    import bcrypt
    from moduli.database import conn, t_admins
    from sqlalchemy import select
    d = request.get_json()
    with conn() as c:
        row = c.execute(
            select(t_admins).where(t_admins.c.id == session["admin_id"])
        ).fetchone()
    if not row:
        return jsonify({"ok": False, "msg": "Sessione non valida"})
    if not bcrypt.checkpw(d.get("vecchia", "").encode(), row.password_hash.encode()):
        return jsonify({"ok": False, "msg": "Password attuale errata"})
    cambia_password_admin(session["admin_id"], d["nuova"])
    return jsonify({"ok": True})


# ─── ADMIN — STORICO ──────────────────────────────────────────────────────────

@app.route("/admin/api/storico/scarichi")
@admin_required
def api_storico_scarichi():
    return jsonify(get_storico_scarichi(100))


@app.route("/admin/api/storico/missioni")
@admin_required
def api_storico_missioni():
    return jsonify(get_storico_missioni(100))


# ─── SSE ──────────────────────────────────────────────────────────────────────

@app.route("/events")
def events():
    def generate():
        yield f"event: stato\ndata: {json.dumps(stato, default=str)}\n\n"
        while True:
            try:
                yield _sse_queue.get(timeout=5)
            except queue.Empty:
                yield ": heartbeat\n\n"
    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    )


# ─── CAMERA PROXY ─────────────────────────────────────────────────────────────

@app.route("/camera")
def camera_feed():
    import cv2
    import numpy as np

    _frame_nero = cv2.imencode(".jpg", np.zeros((240, 320, 3), dtype=np.uint8))[1].tobytes()

    def gen():
        while True:
            broadcaster = _cam_broadcaster
            if broadcaster is not None:
                jpg = broadcaster.next_frame(timeout=1.0) or _frame_nero
            else:
                jpg = _frame_nero
                time.sleep(1.0)
            yield (
                b"--frame\r\nContent-Type: image/jpeg\r\n\r\n" +
                jpg +
                b"\r\n"
            )

    return Response(
        gen(),
        mimetype="multipart/x-mixed-replace; boundary=frame",
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Pragma": "no-cache",
            "X-Accel-Buffering": "no",
            "Access-Control-Allow-Origin": "*",
        }
    )


# ─── PANNELLI OPERATORI ───────────────────────────────────────────────────────

@app.route("/")
@app.route("/balia")
def pannello_balia():
    return render_template("tablet_balia.html")


@app.route("/scarico")
def pannello_scarico():
    return render_template("tablet_scarico.html", destinazioni=get_destinazioni())


# ─── API SCARICO — SCAN CAMION ────────────────────────────────────────────────

@app.route("/api/scan_camion", methods=["POST"])
def scan_camion():
    with _stato_lock:
        if stato["busy"]:
            return jsonify({"ok": False, "msg": "Operazione già in corso."})
        if stato["fase"] != 0:
            return jsonify({"ok": False, "msg": "Fase non corretta."})
        if scanner is None:
            return jsonify({"ok": False, "msg": "Scanner non inizializzato."})
        stato["busy"] = True
        stato["camera_attiva"] = True
    push_stato()
    push_notifica("info", "Scanner attivo — avvicinare il QR del camion alla telecamera...")

    def _lavora():
        try:
            qr = scanner.scansiona_uno(tipo_atteso="camion", timeout_sec=60)
            if not qr:
                push_notifica("errore", "Nessun QR camion rilevato. Riprovare.")
                _update_stato(busy=False, camera_attiva=False)
                push_stato()
                return
            camion_id = qr["dati"].get("id", "")
            pacchi = get_pacchi_by_camion(camion_id)
            if not pacchi:
                push_notifica("errore", f"Nessun pacco per il camion '{camion_id}'.")
                _update_stato(busy=False, camera_attiva=False)
                push_stato()
                return
            _update_stato(
                camion=qr["dati"],
                pacchi_lista=_ser_pacchi(pacchi),
                pacchi_scansionati=[],
                busy=False,
                fase=1,
                camera_attiva=False,
            )
            push_notifica("ok", f"Camion {camion_id} — {len(pacchi)} pacchi registrati")
            push_stato()
        except Exception as e:
            import traceback
            traceback.print_exc()
            push_notifica("errore", str(e))
            _update_stato(busy=False)
            push_stato()

    threading.Thread(target=_lavora, daemon=True).start()
    return jsonify({"ok": True})


# ─── API SCARICO — AUTH FACCIALE ──────────────────────────────────────────────

@app.route("/api/auth_facciale", methods=["POST"])
def auth_facciale():
    fase_richiesta = request.json.get("fase", 1)
    with _stato_lock:
        if stato["busy"]:
            return jsonify({"ok": False, "msg": "Operazione già in corso."})
        if stato["fase"] != fase_richiesta:
            return jsonify({"ok": False, "msg": "Fase non corretta."})
        if rf is None:
            return jsonify({"ok": False, "msg": "Face recognition non inizializzato."})
        stato.update({"busy": True, "camera_attiva": True})
    push_stato()
    push_notifica("info", "Riconoscimento in corso — posizionarsi davanti alla telecamera...")

    def _lavora():
        try:
            operatori = _prepara_operatori_per_face(get_tutti_operatori())
            op = rf.riconosci(operatori, timeout_sec=25, mostra_preview=False)
            if not op:
                push_notifica("errore", "Operatore non riconosciuto.")
                _update_stato(busy=False, camera_attiva=False)
                push_stato()
                return

            if fase_richiesta == 3:
                if op["id"] != stato["operatore"]["id"]:
                    push_notifica("errore", "Accesso negato — deve autenticarsi lo stesso operatore.")
                    _update_stato(busy=False)
                    push_stato()
                    return
                _update_stato(fase=4, busy=False, camera_attiva=False)
                push_notifica("ok", f"Identità confermata — {op['nome']} {op['cognome']}")
                push_stato()
            else:
                _update_stato(
                    operatore={
                        "id":      op["id"],
                        "nome":    op["nome"],
                        "cognome": op["cognome"],
                        "ruolo":   op.get("ruolo", ""),
                    },
                    fase=2,
                    busy=False,
                    camera_attiva=False,
                )
                push_notifica("ok", f"Accesso autorizzato — {op['nome']} {op['cognome']}")
                push_stato()
                time.sleep(0.8)
                _avvia_scan_pacchi()
        except Exception as e:
            import traceback
            traceback.print_exc()
            push_notifica("errore", str(e))
            _update_stato(busy=False)
            push_stato()

    threading.Thread(target=_lavora, daemon=True).start()
    return jsonify({"ok": True})


def _avvia_scan_pacchi():
    _update_stato(busy=True, camera_attiva=True)
    push_notifica("info", "Scansione pacchi — avvicinare i QR code alla telecamera...")
    push_stato()

    def _lavora():
        def _norm(val):
            if isinstance(val, dict):
                return val.get("id", str(val))
            try:
                return json.loads(str(val)).get("id", val)
            except Exception:
                return val

        timeout_sec = int(get_config().get("qr_timeout_sec", "300"))
        id_pacchi   = [p["qr_code"] for p in stato["pacchi_lista"]]
        mappa       = {_norm(q): q for q in id_pacchi}
        rimanenti   = set(mappa.keys())
        t_start     = time.time()

        while rimanenti:
            if time.time() - t_start > timeout_sec:
                push_notifica("errore", f"Timeout scansione — {len(rimanenti)} pacco/i non rilevati.")
                break
            _, trovati = scanner.leggi_frame()
            for qr in trovati:
                pid = _norm(qr["id"])
                if pid in rimanenti:
                    rimanenti.discard(pid)
                    qr["qr_originale"] = mappa[pid]
                    pacco = get_pacco_by_qr(qr["qr_originale"])
                    if pacco:
                        aggiorna_stato_pacco(qr["qr_originale"], "in_transito")
                        with _stato_lock:
                            stato["pacchi_scansionati"].append(qr["qr_originale"])
                        peso = pacco.get("peso_kg")
                        push_event("pacco_letto", {
                            "qr_code":        pacco["qr_code"],
                            "id":             pid,
                            "destinazione":   (pacco.get("destinazioni") or {}).get("nome", "—"),
                            "destinazione_id": pacco.get("destinazione_id"),
                            "descrizione":    pacco.get("descrizione", ""),
                            "peso_kg":        float(peso) if peso is not None else None,
                            "rimanenti":      len(rimanenti),
                            "totale":         len(id_pacchi),
                        })
                        push_stato()

        _update_stato(fase=3, busy=False, camera_attiva=False)
        push_notifica("ok", "Tutti i pacchi sono stati registrati.")
        push_stato()

    threading.Thread(target=_lavora, daemon=True).start()


def _mir_headers(auth_user: str, auth_pass: str) -> dict:
    pw_hash  = hashlib.sha256(auth_pass.encode()).hexdigest()
    auth_b64 = base64.b64encode(f"{auth_user}:{pw_hash}".encode()).decode()
    return {
        "Authorization":  f"Basic {auth_b64}",
        "Content-Type":   "application/json",
        "Accept-Language": "en_US",
    }


def _avvia_ritorno_mir():
    """
    Accoda sul MiR la missione di ritorno verso pos_part usando la sessione
    salvata in _mir_sessione.  Chiamato automaticamente da conferma_ricezione
    quando tutti i pacchi della destinazione risultano consegnati, oppure
    manualmente tramite POST /api/ritorno_mir.
    """
    sess = dict(_mir_sessione)          # copia snapshot thread-safe
    if not sess:
        push_notifica("errore", "MiR100 — nessuna sessione attiva per il ritorno")
        return

    def _esegui_ritorno():
        mir_url  = sess["mir_url"]
        headers  = sess["headers"]
        missions = sess["missions"]
        pos_part = sess["pos_part"]

        return_mission = next(
            (m for m in missions if pos_part.lower() in m.get("name", "").lower()),
            None
        )
        if return_mission:
            try:
                _req.post(
                    f"{mir_url}/mission_queue",
                    headers=headers,
                    json={"mission_id": return_mission["guid"]},
                    timeout=5,
                )
                crea_missione_mir(tipo="ritorno")
                push_notifica("info", f"MiR100 — ritorno verso {pos_part} accodato")
                push_event("mir_ritorno", {"destinazione": pos_part})
                _mir_sessione.clear()
            except Exception as e_ret:
                print(f"[MiR ritorno] {e_ret}")
                push_notifica("errore", f"MiR100 — errore invio ritorno: {e_ret}")
        else:
            push_notifica("errore",
                f"MiR100 — missione di ritorno '{pos_part}' non trovata sul robot")

    threading.Thread(target=_esegui_ritorno, daemon=True).start()


@app.route("/api/invia_mir", methods=["POST"])
def invia_mir():
    if stato["fase"] != 4:
        return jsonify({"ok": False, "msg": "Autenticazione finale non completata."})

    cfg       = get_config()
    mir_ip    = cfg.get("mir_ip",                    "192.168.1.200")
    auth_user = cfg.get("mir_auth_user",             "Distributor")
    auth_pass = os.environ.get("MIR_AUTH_PASS") or cfg.get("mir_auth_pass", "distributor")
    pos_part  = cfg.get("mir_posizione_partenza",    "area B1")
    pos_dest  = cfg.get("mir_posizione_destinazione","area A1")

    push_notifica("info", f"MiR100 — partenza: {pos_part} → destinazione: {pos_dest}...")

    def _esegui():
        def _reset_stato():
            _update_stato(
                fase=0, operatore=None, camion=None,
                pacchi_lista=[], pacchi_scansionati=[],
                busy=False, camera_attiva=False,
            )
            push_stato()

        missione_db_id = None
        try:
            mir_url = f"http://{mir_ip}/api/v2.0.0"
            headers = _mir_headers(auth_user, auth_pass)

            # Verifica stato robot prima di procedere
            try:
                rs = _req.get(f"{mir_url}/status", headers=headers, timeout=5)
                rs.raise_for_status()
                robot_state_id = rs.json().get("state_id")
                if robot_state_id in (10, 12):  # 10=Emergency stop, 12=Error
                    push_notifica("errore", f"MiR100 in stato critico (state_id={robot_state_id}) — missione annullata.")
                    return
            except Exception as e_status:
                print(f"[MiR] Impossibile verificare stato robot: {e_status}")

            # Trova la missione configurata sul robot per la destinazione
            r = _req.get(f"{mir_url}/missions", headers=headers, timeout=5)
            r.raise_for_status()
            missions = r.json()

            mission = next(
                (m for m in missions if pos_dest.lower() in m.get("name", "").lower()),
                None
            )
            if not mission:
                push_notifica("errore", f"Nessuna missione trovata per '{pos_dest}' sul MiR.")
                return

            # Accoda la missione
            r2 = _req.post(
                f"{mir_url}/mission_queue",
                headers=headers,
                json={"mission_id": mission["guid"]},
                timeout=5,
            )
            r2.raise_for_status()
            queue_entry_id = r2.json().get("id")

            missione_db_id = crea_missione_mir(tipo="consegna")
            push_notifica("ok", f"MiR100 avviato — missione '{mission['name']}' accodata")

            # La balia può ricominciare subito: il MiR è già in viaggio
            time.sleep(2)
            _reset_stato()

            # Polling fino a completamento (max 10 min)
            if queue_entry_id:
                t_poll = time.time()
                while time.time() - t_poll < 600:
                    try:
                        rp = _req.get(f"{mir_url}/mission_queue/{queue_entry_id}", headers=headers, timeout=5)
                        rp.raise_for_status()
                        entry_state = rp.json().get("state", "")
                        if entry_state == "Done":
                            aggiorna_stato_missione(missione_db_id, "completata")
                            push_notifica("ok", f"MiR100 — robot arrivato in {pos_dest}")
                            push_event("mir_arrivato", {"destinazione": pos_dest})
                            # Salva la sessione: il ritorno verrà avviato da conferma_ricezione
                            # (o manualmente via /api/ritorno_mir) quando tutti i pacchi sono confermati.
                            _mir_sessione.update({
                                "mir_url":  mir_url,
                                "headers":  headers,
                                "missions": missions,
                                "pos_part": pos_part,
                                "pos_dest": pos_dest,
                            })
                            break
                        elif entry_state in ("Aborted", "Error"):
                            aggiorna_stato_missione(missione_db_id, "annullata")
                            push_notifica("errore", f"MiR100 — missione {entry_state.lower()}")
                            break
                    except Exception as poll_err:
                        print(f"[MiR poll] {poll_err}")
                    time.sleep(5)
                else:
                    push_notifica("errore", "MiR100 — timeout attesa completamento missione")
                    if missione_db_id:
                        aggiorna_stato_missione(missione_db_id, "annullata")

        except Exception as e:
            import traceback
            traceback.print_exc()
            push_notifica("errore", f"Errore comunicazione MiR100: {e}")
            _reset_stato()   # Reset balia anche in caso di errore

    threading.Thread(target=_esegui, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/stato")
def get_stato():
    return jsonify(stato)


@app.route("/api/pacchi_destinazione/<int:dest_id>")
def api_pacchi_destinazione(dest_id):
    return jsonify(get_pacchi_by_destinazione(dest_id))


@app.route("/api/conferma_ricezione", methods=["POST"])
def conferma_ricezione():
    d       = request.get_json()
    qr_code = d["qr_code"]
    aggiorna_stato_pacco(qr_code, "consegnato")

    # Dopo aver segnato il pacco come consegnato, controlla se ne rimangono
    # altri "in_transito" per la stessa destinazione.
    pacco = get_pacco_by_qr(qr_code)
    if pacco and pacco.get("destinazione_id"):
        dest_id    = pacco["destinazione_id"]
        rimanenti  = get_pacchi_by_destinazione(dest_id)   # solo in_transito
        if not rimanenti:
            # Tutti i pacchi di questa destinazione confermati → push SSE + ritorno MiR
            push_event("tutti_consegnati", {"destinazione_id": dest_id})
            if _mir_sessione:
                _avvia_ritorno_mir()

    return jsonify({"ok": True})


@app.route("/api/ritorno_mir", methods=["POST"])
def api_ritorno_mir():
    """Endpoint manuale: forza il ritorno del MiR anche prima che tutti i pacchi
    siano stati confermati (es. operatore clicca 'Rimanda MiR')."""
    if not _mir_sessione:
        return jsonify({"ok": False, "msg": "Nessuna sessione MiR attiva per il ritorno"})
    _avvia_ritorno_mir()
    return jsonify({"ok": True})


# ─── AVVIO ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    get_engine()
    _init_hardware()

    import socket
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        ip = "localhost"

    print("=" * 55)
    print("  MiR100 — Sistema Gestione Magazzino v2")
    print("=" * 55)
    print(f"  Pannello Balia:   http://{ip}:5000/")
    print(f"  Pannello Scarico: http://{ip}:5000/scarico")
    print(f"  Admin:            http://{ip}:5000/admin")
    print(f"  Admin login:      admin / admin1234")
    print("=" * 55)

    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True, use_reloader=False)