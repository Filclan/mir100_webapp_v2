"""
reset_config.py — Utilità per aggiornare la configurazione nel DB
Esegui con: python reset_config.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from moduli.database import get_engine, set_config, get_config

get_engine()  # inizializza il DB se non esiste

print("=== Configurazione attuale ===")
cfg = get_config()
for k, v in cfg.items():
    print(f"  {k} = {v}")

print("\n=== Aggiornamento config ===")
# Modifica questi valori se necessario
set_config("raspberry_ip",              "192.168.12.242")
set_config("raspberry_port",            "8081")
set_config("mir_ip",                    "192.168.12.20")
set_config("mir_auth_user",             "itisdelpozzo")
set_config("mir_auth_pass",             "itisdelpozzo")
set_config("mir_posizione_partenza",    "area B1")
set_config("mir_posizione_destinazione","area A1")
set_config("face_soglia",               "0.55")

print("Configurazione aggiornata:")
cfg = get_config()
for k, v in cfg.items():
    print(f"  {k} = {v}")

print("\n[OK] Done.")
