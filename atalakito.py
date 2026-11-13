#!/usr/bin/env python3
"""Nyers gyűjtő-CSV -> tiszta adatok (készletszűrés).
Használat: python atalakito.py adatok/nyers.csv adatok
Kimenet: adatok/adatok.json, adatok/adatok.csv, adatok/jelentes.txt, adatok/arvaltozas.json"""
import csv, json, sys, os
from datetime import date

KIZART_KESZLET = ("nincs raktáron", "elfogyott", "nem rendelhető", "megszűnt")

def szam(v):
    try:
        return float(v) if v not in (None, "") else None
    except ValueError:
        return None

def main(src, out):
    os.makedirs(out, exist_ok=True)
    with open(src, encoding="utf-8-sig", newline="") as fh:
        sorok = list(csv.DictReader(fh, delimiter=";"))
    jo, ki = [], {"nincs készleten": 0, "nem ok vagy nincs ár": 0}
    for r in sorok:
        if r.get("status") != "ok" or szam(r.get("price")) is None:
            ki["nem ok vagy nincs ár"] += 1
            continue
        if (r.get("availability") or "").lower() in KIZART_KESZLET:
            ki["nincs készleten"] += 1
            continue
        jo.append({
            "bolt": r["shop_name"], "kategoria": r.get("category"), "url": r["url"],
            "nev": r.get("name"), "valtozat": r.get("variant") or None,
            "cikkszam": r.get("sku") or None, "marka": r.get("brand") or None,
            "ar": szam(r["price"]), "eredeti_ar": szam(r.get("original_price")),
            "keszlet": r.get("availability") or "nincs adat",
            "hossz_m": szam(r.get("length_m")), "dobosuly_g": r.get("casting_weight_g") or None,
            "ido": r.get("scraped_at"),
        })
    # árváltozás az előző futáshoz képest (kulcs: url + változat)
    elozo_ut = os.path.join(out, "adatok.json")
    elozo = {}
    if os.path.exists(elozo_ut):
        with open(elozo_ut, encoding="utf-8") as fh:
            elozo = {(p["url"], p["valtozat"]): p["ar"] for p in json.load(fh)}
    valt = [{"nev": p["nev"], "bolt": p["bolt"], "valtozat": p["valtozat"], "url": p["url"],
             "regi": elozo[(p["url"], p["valtozat"])], "uj": p["ar"]}
            for p in jo if (p["url"], p["valtozat"]) in elozo and elozo[(p["url"], p["valtozat"])] != p["ar"]]
    with open(os.path.join(out, "arvaltozas.json"), "w", encoding="utf-8") as fh:
        json.dump({"datum": str(date.today()), "valtozasok": valt}, fh, ensure_ascii=False, indent=1)
    with open(elozo_ut, "w", encoding="utf-8") as fh:
        json.dump(jo, fh, ensure_ascii=False, indent=1)
    with open(os.path.join(out, "adatok.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(jo[0]) if jo else ["bolt"], delimiter=";")
        w.writeheader(); w.writerows(jo)
    boltok = {}
    for p in jo: boltok[p["bolt"]] = boltok.get(p["bolt"], 0) + 1
    with open(os.path.join(out, "jelentes.txt"), "w", encoding="utf-8") as fh:
        fh.write(f"Dátum: {date.today()}\nBemenet: {len(sorok)} sor\nBeépíthető: {len(jo)}\n"
                 f"Kiszűrve – nincs készleten/nem rendelhető: {ki['nincs készleten']}\n"
                 f"Kiszűrve – hibás vagy ár nélküli: {ki['nem ok vagy nincs ár']}\n"
                 f"Árváltozás az előző futáshoz képest: {len(valt)} termék\nBoltonként: {boltok}\n")
    print(open(os.path.join(out, "jelentes.txt"), encoding="utf-8").read())

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
