"""
Script para GitHub Actions. Actualiza trm_data.json sin depender de Excel ni openpyxl.
"""

import requests
import json
from datetime import date, timedelta
from pathlib import Path

API_URL   = "https://www.datos.gov.co/resource/ceyp-9c7c.json"
JSON_PATH = Path(__file__).parent.parent / "trm_data.json"

def fetch_latest():
    params = {"$limit": 5, "$order": "vigenciadesde DESC"}
    r = requests.get(API_URL, params=params, timeout=30)
    r.raise_for_status()
    return r.json()

def expand(raw):
    hoy = date.today()
    expanded = {}
    raw_sorted = sorted(raw, key=lambda r: r["vigenciadesde"], reverse=True)
    for rec in raw_sorted:
        trm   = round(float(rec["valor"]), 2)
        desde = date.fromisoformat(rec["vigenciadesde"][:10])
        hasta = date.fromisoformat(rec["vigenciahasta"][:10])
        d = desde
        while d <= hasta and d <= hoy:
            k = d.isoformat()
            if k not in expanded:
                expanded[k] = trm
            d += timedelta(days=1)
    # Rellenar huecos hasta hoy con el último TRM conocido
    if expanded:
        ultima = date.fromisoformat(max(expanded.keys()))
        ultimo_trm = expanded[ultima.isoformat()]
        d = ultima + timedelta(days=1)
        while d <= hoy:
            expanded[d.isoformat()] = ultimo_trm
            d += timedelta(days=1)
    return expanded

def main():
    raw = fetch_latest()
    new = expand(raw)

    data = []
    if JSON_PATH.exists():
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

    existing = {r["fecha"] for r in data}
    added = 0
    for fecha, trm in new.items():
        if fecha not in existing:
            data.append({"fecha": fecha, "trm": trm})
            added += 1

    data.sort(key=lambda x: x["fecha"])
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)

    latest = data[-1]
    print(f"TRM {latest['fecha']}: ${latest['trm']:,.2f} COP — {added} registro(s) nuevo(s)")

if __name__ == "__main__":
    main()
