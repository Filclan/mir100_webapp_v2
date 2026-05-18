"""
moduli/qr_scanner.py
Gestisce la lettura dei QR code dalla webcam.
Funziona sia per il QR del camion (lista pacchi) che per i QR dei singoli pacchi.
"""

import cv2
import json
import numpy as np
from pyzbar import pyzbar
from datetime import datetime


class QRScanner:
    """
    Scanner QR che usa un CameraManager condiviso.
    Uso:
        cam = CameraManager()
        scanner = QRScanner(cam)
        dati = scanner.scansiona_uno()
    """

    def __init__(self, camera_manager=None, camera_index: int = 0):
        from moduli.camera_manager import CameraManager
        if camera_manager is not None:
            self._cam = camera_manager
            self._owns_camera = False  # non siamo noi ad averla aperta
        else:
            self._cam = CameraManager(camera_index)
            self._owns_camera = True
        self._ultimo_qr_id = None

    def _decodifica(self, dati_raw: str) -> dict:
        """Prova a parsare il contenuto del QR come JSON, altrimenti lo tratta come stringa."""
        try:
            return json.loads(dati_raw)
        except (json.JSONDecodeError, ValueError):
            return {"id": dati_raw, "tipo": "sconosciuto", "raw": dati_raw}

    def leggi_frame(self) -> tuple[np.ndarray, list[dict]]:
        """
        Legge un frame dalla camera e restituisce (frame_annotato, lista_qr_trovati).
        Ogni QR trovato è un dict con: tipo, id, dati_completi, timestamp.
        """
        frame = self._cam.leggi_frame()
        if frame is None:
            return None, []

        qr_trovati = []
        codici = pyzbar.decode(frame)

        for codice in codici:
            dati_raw = codice.data.decode("utf-8")
            dati = self._decodifica(dati_raw)

            # Disegna rettangolo verde attorno al QR
            (x, y, w, h) = codice.rect
            cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 200, 0), 2)
            label = dati.get("id", dati.get("raw", ""))[:30]
            cv2.putText(frame, label, (x, y - 8),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 200, 0), 2)

            qr_trovati.append({
                "tipo": dati.get("tipo", "sconosciuto"),
                "id": dati.get("id", dati.get("raw", "")),
                "dati": dati,
                "timestamp": datetime.now().isoformat()
            })

        return frame, qr_trovati

    def scansiona_uno(self, tipo_atteso: str = None, timeout_sec: float = 30.0) -> dict | None:
        """
        Blocca finché non legge UN QR valido (diverso dall'ultimo letto).
        - tipo_atteso: se specificato, accetta solo QR con quel tipo ("camion" o "pacco")
        - timeout_sec: secondi prima di restituire None
        Restituisce il dict del QR letto, o None se timeout.
        """
        import time
        inizio = time.time()

        while True:
            if time.time() - inizio > timeout_sec:
                return None

            frame, trovati = self.leggi_frame()

            for qr in trovati:
                # Ignora ripetizioni dello stesso QR
                if qr["id"] == self._ultimo_qr_id:
                    continue
                # Filtra per tipo se richiesto
                if tipo_atteso and qr["tipo"] != tipo_atteso:
                    continue
                self._ultimo_qr_id = qr["id"]
                return qr

            time.sleep(2)

    def scansiona_lista(self, id_pacchi_attesi: list[str],
                         callback_pacco: callable = None,
                         callback_completato: callable = None) -> list[str]:
        """
        Scansiona in loop finché tutti i pacchi nella lista sono stati letti.
        - callback_pacco(qr_dict): chiamato ogni volta che un pacco viene letto
        - callback_completato(): chiamato quando la lista è vuota
        Restituisce la lista degli ID letti.
        """
        import json as _json

        # Normalizza gli id attesi: estrai sempre l'id interno se è JSON
        def _normalizza(val):
            if isinstance(val, dict):
                return val.get("id", str(val))
            try:
                parsed = _json.loads(str(val))
                return parsed.get("id", val)
            except (_json.JSONDecodeError, TypeError):
                return val

        # Mappa id_normalizzato -> qr_code originale (per aggiornare il DB)
        mappa_id = {_normalizza(qr): qr for qr in id_pacchi_attesi}
        rimanenti = set(mappa_id.keys())
        letti = []


        import time as _time
        while rimanenti:
            frame, trovati = self.leggi_frame()

            for qr in trovati:
                pacco_id = _normalizza(qr["id"])
                if pacco_id in rimanenti and pacco_id != _normalizza(self._ultimo_qr_id or ""):
                    self._ultimo_qr_id = qr["id"]
                    # Passa il qr_code originale al callback (serve per aggiornare il DB)
                    qr["qr_originale"] = mappa_id[pacco_id]
                    rimanenti.discard(pacco_id)
                    letti.append(pacco_id)
                    if callback_pacco:
                        callback_pacco(qr)

            _time.sleep(2)

        if not rimanenti and callback_completato:
            callback_completato()

        return letti

    def rilascia(self) -> None:
        """Chiude la camera solo se è stata aperta da questo scanner."""
        if self._owns_camera:
            self._cam.rilascia()
        cv2.destroyAllWindows()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.rilascia()


# ─── GENERATORE QR DI TEST ────────────────────────────────────────────────────

def genera_qr_test(output_dir: str = "assets") -> None:
    """
    Genera QR code di esempio per testare il sistema senza dati reali.
    Richiede: pip install qrcode[pil]
    """
    try:
        import qrcode
        import os
    except ImportError:
        print("Installa qrcode: pip install qrcode[pil]")
        return

    os.makedirs(output_dir, exist_ok=True)

    esempi = [
        # QR camion
        {"tipo": "camion", "id": "CAM-001", "targa": "AB123CD",
         "fornitore": "LogiTrasporti SRL"},
        # QR pacchi con destinazioni diverse
        {"tipo": "pacco", "id": "PKG-001", "destinazione_id": 1,
         "destinazione_nome": "Zona A - Scaffali"},
        {"tipo": "pacco", "id": "PKG-002", "destinazione_id": 2,
         "destinazione_nome": "Zona B - Frigo"},
        {"tipo": "pacco", "id": "PKG-003", "destinazione_id": 1,
         "destinazione_nome": "Zona A - Scaffali"},
        {"tipo": "pacco", "id": "PKG-004", "destinazione_id": 3,
         "destinazione_nome": "Zona C - Spedizioni"},
    ]

    for dato in esempi:
        qr = qrcode.QRCode(version=1, box_size=10, border=4)
        qr.add_data(json.dumps(dato))
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        nome_file = f"{output_dir}/qr_{dato['id'].lower().replace('-', '_')}.png"
        img.save(nome_file)
        print(f"[QR] Generato: {nome_file}")

    print(f"\n[QR] {len(esempi)} QR code di test salvati in '{output_dir}/'")
    print("Aprili a schermo intero e mostrali alla webcam per testare lo scanner.")


if __name__ == "__main__":
    print("Generazione QR code di test...")
    genera_qr_test()
