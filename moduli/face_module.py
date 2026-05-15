"""
moduli/face_recognition_module.py
Gestisce il riconoscimento facciale degli operatori.
Confronta il volto rilevato dalla webcam con gli encoding salvati su Supabase.
"""

import cv2
import numpy as np
import face_recognition
from datetime import datetime


class RiconoscimentoFacciale:
    """
    Gestisce la cattura e il riconoscimento dei volti degli operatori.

    Uso tipico:
        rf = RiconoscimentoFacciale()
        operatore = rf.riconosci(operatori_db)  # blocca finché non trova un volto
        rf.rilascia()
    """

    def __init__(self, camera_manager=None, camera_index: int = 0, soglia: float = 0.55):
        """
        camera_manager: istanza CameraManager condivisa (preferita)
        camera_index:   usato solo se camera_manager è None
        soglia:         distanza massima per riconoscimento (0.0–1.0)
        """
        from moduli.camera_manager import CameraManager
        if camera_manager is not None:
            self._cam = camera_manager
            self._owns_camera = False
        else:
            self._cam = CameraManager(camera_index)
            self._owns_camera = True
        self.soglia = soglia

    def _encoding_da_frame(self, frame: np.ndarray) -> list:
        """
        Estrae gli encoding facciali da un frame BGR (formato OpenCV).
        Restituisce lista di encoding (uno per ogni volto trovato nel frame).
        """
        # face_recognition lavora in RGB, OpenCV usa BGR
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        posizioni = face_recognition.face_locations(rgb, model="hog")
        encodings = face_recognition.face_encodings(rgb, posizioni)
        return encodings, posizioni

    def riconosci(self, operatori_db: list,
                  timeout_sec: float = 20.0,
                  mostra_preview: bool = True) -> dict | None:
        """
        Apre la preview e cerca un volto riconosciuto tra gli operatori del DB.

        operatori_db: lista di dict con chiavi 'id', 'nome', 'cognome', 'face_encoding'
        timeout_sec:  secondi massimi di attesa
        mostra_preview: se True mostra la finestra della camera

        Restituisce il dict operatore se trovato, None se timeout o nessun match.
        """
        import time

        # Prepara gli encoding del DB in array numpy
        enc_db = []
        for op in operatori_db:
            fe = op.get("face_encoding")
            if fe is not None:
                enc_db.append((op, np.array(fe)))

        if not enc_db:
            print("[Face] Nessun operatore con encoding nel DB.")
            return None

        inizio = time.time()
        print(f"[Face] Riconoscimento avviato. Attendo volto ({timeout_sec}s)...")

        while True:
            elapsed = time.time() - inizio
            if elapsed > timeout_sec:
                print("[Face] Timeout — nessun operatore riconosciuto.")
                return None

            frame = self._cam.leggi_frame()
            if frame is None:
                continue

            encodings, posizioni = self._encoding_da_frame(frame)

            operatore_trovato = None
            for enc_input, pos in zip(encodings, posizioni):
                for op, enc_op in enc_db:
                    distanza = face_recognition.face_distance([enc_op], enc_input)[0]
                    if distanza < self.soglia:
                        operatore_trovato = op
                        # Disegna rettangolo verde
                        top, right, bottom, left = pos
                        cv2.rectangle(frame, (left, top), (right, bottom), (0, 200, 0), 2)
                        nome = f"{op['nome']} {op['cognome']}"
                        cv2.putText(frame, f"OK: {nome}",
                                    (left, top - 10),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 200, 0), 2)
                        break
                    else:
                        # Volto non riconosciuto — rettangolo rosso
                        top, right, bottom, left = pos
                        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 200), 2)
                        cv2.putText(frame, "Non autorizzato",
                                    (left, top - 10),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 200), 2)

            if mostra_preview:
                # Barra info in basso
                rimasti = max(0, timeout_sec - elapsed)
                cv2.putText(frame,
                            f"Avvicina il volto alla camera  |  {rimasti:.0f}s  |  Q = annulla",
                            (10, frame.shape[0] - 12),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.5, (200, 200, 200), 1)
                cv2.imshow("MiR100 — Riconoscimento operatore", frame)

            if operatore_trovato:
                print(f"[Face] Operatore riconosciuto: {operatore_trovato['nome']} {operatore_trovato['cognome']}")
                if mostra_preview:
                    cv2.waitKey(1000)  # mostra il frame di conferma 1 secondo
                return operatore_trovato

            if cv2.waitKey(1) & 0xFF == ord('q'):
                return None

    def registra_operatore(self, nome: str, cognome: str,
                           n_campioni: int = 5) -> list | None:
        """
        Cattura il volto di un nuovo operatore e calcola il suo encoding medio.
        Usato in fase di setup per popolare il DB.

        n_campioni: quanti frame usare per calcolare l'encoding medio (più = più preciso)
        Restituisce l'encoding come lista Python (pronta per Supabase JSON).
        """
        import time

        print(f"[Face] Registrazione operatore: {nome} {cognome}")
        print(f"[Face] Verranno catturati {n_campioni} campioni. Guarda in camera...")

        campioni = []
        while len(campioni) < n_campioni:
            frame = self._cam.leggi_frame()
            if frame is None:
                continue

            encodings, posizioni = self._encoding_da_frame(frame)

            for enc, pos in zip(encodings, posizioni):
                campioni.append(enc)
                top, right, bottom, left = pos
                cv2.rectangle(frame, (left, top), (right, bottom), (255, 140, 0), 2)
                cv2.putText(frame, f"Campione {len(campioni)}/{n_campioni}",
                            (left, top - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.7, (255, 140, 0), 2)

            cv2.imshow("MiR100 — Registrazione operatore", frame)
            cv2.waitKey(300)  # pausa tra campioni

        if not campioni:
            print("[Face] Nessun volto rilevato durante la registrazione.")
            return None

        # Calcola encoding medio per maggiore robustezza
        encoding_medio = np.mean(campioni, axis=0).tolist()
        print(f"[Face] Registrazione completata per {nome} {cognome}.")
        return encoding_medio

    def rilascia(self) -> None:
        if self._owns_camera:
            self._cam.rilascia()
        cv2.destroyAllWindows()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.rilascia()
