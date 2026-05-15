"""
moduli/camera_manager.py
Gestisce la connessione allo stream MJPEG del Raspberry Pi.

Architettura:
  MJPEGBroadcaster  — apre UNA SOLA connessione HTTP al Pi e distribuisce
                      i frame JPEG a tutti i consumer via threading.Condition.
  CameraManager     — consumer del broadcaster: decodifica JPEG → numpy BGR
                      in un thread dedicato, così leggi_frame() è istantaneo
                      e restituisce sempre il frame più recente senza ritardi
                      di buffering OpenCV.
"""

import cv2
import threading
import numpy as np
import requests


class MJPEGBroadcaster:
    """
    Apre UNA SOLA connessione HTTP verso lo stream MJPEG (Raspberry Pi)
    e distribuisce i frame JPEG a tutti i consumer tramite notify_all().

    Più client possono chiamare next_frame() contemporaneamente:
    vengono tutti svegliati in contemporanea ad ogni nuovo frame.
    """

    def __init__(self, url: str):
        self._url   = url
        self._frame: bytes = b""
        self._cond  = threading.Condition()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        print(f"[MJPEGBroadcaster] Avviato → {url}")

    def _loop(self):
        import time
        while True:
            try:
                with requests.get(self._url, stream=True, timeout=(5, 60)) as r:
                    buf = b""
                    for chunk in r.iter_content(chunk_size=65536):
                        buf += chunk
                        # Estrai tutti i JPEG completi presenti nel buffer
                        while True:
                            s = buf.find(b"\xff\xd8")
                            e = buf.find(b"\xff\xd9", s + 2) if s != -1 else -1
                            if s == -1 or e == -1:
                                break
                            jpg = buf[s:e + 2]
                            buf = buf[e + 2:]
                            with self._cond:
                                self._frame = jpg
                                self._cond.notify_all()
            except Exception as err:
                print(f"[MJPEGBroadcaster] Connessione persa: {err} — riconnessione in 2s")
                time.sleep(2)

    def next_frame(self, timeout: float = 1.0) -> bytes:
        """Attende il prossimo frame e lo restituisce come JPEG bytes."""
        with self._cond:
            self._cond.wait(timeout)
            return self._frame

    @property
    def latest_frame(self) -> bytes:
        """Restituisce l'ultimo frame disponibile senza attendere."""
        with self._cond:
            return self._frame


class CameraManager:
    """
    Legge frame dal MJPEGBroadcaster e li decodifica in numpy array BGR
    per QR scanner e face recognition.

    Un thread dedicato decodifica continuamente JPEG → numpy, quindi
    leggi_frame() restituisce il frame più recente senza mai bloccarsi
    né accumulare frame in ritardo (problema del buffer interno OpenCV).
    """

    def __init__(self, broadcaster: MJPEGBroadcaster):
        self._broadcaster = broadcaster
        self._lock  = threading.Lock()
        self._frame: np.ndarray | None = None
        self._thread = threading.Thread(target=self._decode_loop, daemon=True)
        self._thread.start()
        print("[CameraManager] Thread decoder avviato")

    def _decode_loop(self):
        while True:
            jpg = self._broadcaster.next_frame(timeout=1.0)
            if not jpg:
                continue
            try:
                arr   = np.frombuffer(jpg, dtype=np.uint8)
                frame = cv2.imdecode(arr, cv2.IMREAD_COLOR)
                if frame is not None:
                    with self._lock:
                        self._frame = frame
            except Exception as e:
                print(f"[CameraManager] Errore decodifica frame: {e}")

    def leggi_frame(self) -> np.ndarray | None:
        """Restituisce l'ultimo frame decodificato (numpy BGR). Non blocca mai."""
        with self._lock:
            return self._frame.copy() if self._frame is not None else None

    def rilascia(self):
        pass  # il broadcaster è gestito esternamente

    def __del__(self):
        pass
