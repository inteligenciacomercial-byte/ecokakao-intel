"""
Robot de ECO-KAKAO: descarga de UN Comtrade las importaciones de cacao
(HS 1801, 1803, 1804, 1805, 1806) de los 40 mercados y las guarda en
imports.json para que la web las muestre.

La clave de API se lee de la variable de entorno COMTRADE_KEY
(guardada como "secret" en GitHub). Nunca se escribe en el código.
"""
import datetime
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

KEY = os.environ.get("COMTRADE_KEY", "").strip()
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imports.json")
API = "https://comtradeapi.un.org/data/v1/get/C/A/HS"
REF = "https://comtradeapi.un.org/files/v1/app/reference/Reporters.json"
HS = ["1801", "1803", "1804", "1805", "1806"]

# Los 40 mercados (ISO3). Turquía primero: es el piloto.
MARKETS = [
    "TUR", "NLD", "DEU", "BEL", "FRA", "ITA", "ESP", "GBR", "CHE", "POL",
    "EST", "AUT", "CZE", "SWE", "DNK", "FIN", "IRL", "PRT", "GRC", "UKR",
    "RUS", "USA", "CAN", "MEX", "BRA", "ARG", "COL", "CHL", "JPN", "CHN",
    "KOR", "MYS", "IDN", "SGP", "IND", "AUS", "SAU", "ARE", "EGY", "ZAF",
]
# Códigos de país que usa UN Comtrade (respaldo si la lista oficial no carga).
FALLBACK_CODES = {
    "TUR": 792, "NLD": 528, "DEU": 276, "BEL": 56, "FRA": 251, "ITA": 380,
    "ESP": 724, "GBR": 826, "CHE": 757, "POL": 616, "EST": 233, "AUT": 40,
    "CZE": 203, "SWE": 752, "DNK": 208, "FIN": 246, "IRL": 372, "PRT": 620,
    "GRC": 300, "UKR": 804, "RUS": 643, "USA": 842, "CAN": 124, "MEX": 484,
    "BRA": 76, "ARG": 32, "COL": 170, "CHL": 152, "JPN": 392, "CHN": 156,
    "KOR": 410, "MYS": 458, "IDN": 360, "SGP": 702, "IND": 699, "AUS": 36,
    "SAU": 682, "ARE": 784, "EGY": 818, "ZAF": 710,
}


def get_json(url, params=None, auth=True, tries=3):
    if params:
        url = url + "?" + urllib.parse.urlencode(params)
    headers = {"User-Agent": "ecokakao-intel/1.0"}
    if auth:
        headers["Ocp-Apim-Subscription-Key"] = KEY
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            # 429 = demasiadas consultas: esperar y reintentar
            if e.code in (429, 500, 502, 503, 504) and attempt < tries - 1:
                time.sleep(10 * (attempt + 1))
                continue
            raise
        except urllib.error.URLError:
            if attempt < tries - 1:
                time.sleep(10)
                continue
            raise


def reporter_codes():
    """Códigos de UN Comtrade de todos los países (no grupos) + ISO2 y nombre en inglés."""
    codes = dict(FALLBACK_CODES)
    info = {}
    try:
        ref = get_json(REF, auth=False)
        for r in ref.get("results", []):
            iso3 = r.get("reporterCodeIsoAlpha3")
            if not iso3 or r.get("isGroup") or r.get("entryExpiredDate"):
                continue
            codes[iso3] = r.get("reporterCode", codes.get(iso3))
            info[iso3] = {"iso2": r.get("reporterCodeIsoAlpha2"), "name": r.get("reporterDesc") or r.get("text")}
    except Exception as e:
        print("Aviso: no se pudo leer la lista oficial de países:", e)
    return codes, info


ECUADOR = 218
PILOT = "TUR"   # mercado piloto: siempre incluido
N_MARKETS = 40


def select_markets(years, info):
    """Los 40 principales destinos del cacao de Ecuador (HS 1801-1806),
    por valor FOB promedio de los últimos años, según lo que declara Ecuador."""
    params = {"reporterCode": ECUADOR, "period": ",".join(str(y) for y in years), "flowCode": "X",
              "cmdCode": ",".join(HS), "partner2Code": 0, "customsCode": "C00", "motCode": 0, "includeDesc": "true"}
    data = get_json(API, params)
    tot, seen_years = {}, set()
    for r in (data.get("data") or []):
        if not is_total_row(r) or str(r.get("partnerCode")) == "0":
            continue
        iso3 = r.get("partnerISO")
        if not iso3 or iso3 not in info or len(iso3) != 3:   # descarta agregados y zonas sin país
            continue
        seen_years.add(str(r.get("period")))
        tot[iso3] = tot.get(iso3, 0) + float(r.get("primaryValue") or 0)
    if not tot:
        return None
    n_years = max(1, len(seen_years))
    ranked = sorted(tot.items(), key=lambda kv: -kv[1])
    chosen = [iso for iso, _ in ranked[:N_MARKETS]]
    if PILOT not in chosen:
        chosen = chosen[:N_MARKETS - 1] + [PILOT]
    out = []
    for i, iso in enumerate(chosen):
        out.append({"iso3": iso, "iso2": info.get(iso, {}).get("iso2"), "nameEn": info.get(iso, {}).get("name"),
                    "avgUsd": round(tot.get(iso, 0) / n_years), "rank": i + 1, "pilot": iso == PILOT})
    return {"criterion": "Top %d destinos de las exportaciones de cacao de Ecuador (HS 1801-1806), valor FOB promedio %s, declarado por Ecuador a UN Comtrade. Turquía se mantiene como piloto." % (N_MARKETS, "-".join(sorted(seen_years)[::max(1, len(seen_years)-1)])),
            "years": sorted(int(y) for y in seen_years), "markets": out}


def is_total_row(row):
    """Evita duplicados: solo filas totales (sin desglose por aduana, transporte o 2º socio)."""
    if str(row.get("partner2Code", 0)) not in ("0", "None"):
        return False
    if row.get("customsCode") not in (None, "C00"):
        return False
    if str(row.get("motCode", 0)) not in ("0", "None"):
        return False
    return True


def build_market(iso3, code, years):
    params = {
        "reporterCode": code,
        "period": ",".join(str(y) for y in years),
        "flowCode": "M",
        "cmdCode": ",".join(HS),
        "partner2Code": 0,
        "customsCode": "C00",
        "motCode": 0,
        "includeDesc": "true",
    }
    data = get_json(API, params)
    rows = [r for r in (data.get("data") or []) if is_total_row(r)]
    if not rows:
        return None

    by_year = {}
    reporter_desc = rows[0].get("reporterDesc") or iso3
    for y in years:
        yr = [r for r in rows if str(r.get("period")) == str(y)]
        if not yr:
            continue
        hs = {}
        for c in HS:
            world = [r for r in yr if str(r.get("cmdCode")) == c and str(r.get("partnerCode")) == "0"]
            if world:
                usd = sum(float(r.get("primaryValue") or 0) for r in world)
                kg = sum(float(r.get("netWgt") or 0) for r in world)
            else:  # sin fila "World": sumar socios
                parts = [r for r in yr if str(r.get("cmdCode")) == c and str(r.get("partnerCode")) != "0"]
                usd = sum(float(r.get("primaryValue") or 0) for r in parts)
                kg = sum(float(r.get("netWgt") or 0) for r in parts)
            hs[c] = {"usd": round(usd), "kg": round(kg)}

        # Ecuador por código HS
        ecuador = {}
        for c in HS:
            ec = [r for r in yr if str(r.get("cmdCode")) == c and r.get("partnerISO") == "ECU"]
            usd = sum(float(r.get("primaryValue") or 0) for r in ec)
            if usd > 0 and hs[c]["usd"] > 0:
                ecuador[c] = {"usd": round(usd), "share": round(usd / hs[c]["usd"] * 100, 2)}

        # Origen del grano (1801)
        origins = []
        beans = [r for r in yr if str(r.get("cmdCode")) == "1801" and str(r.get("partnerCode")) != "0"]
        tot = hs["1801"]["usd"]
        if tot > 0:
            agg = {}
            for r in beans:
                k = r.get("partnerISO") or str(r.get("partnerCode"))
                a = agg.setdefault(k, {"iso3": k, "name": r.get("partnerDesc") or k, "usd": 0})
                a["usd"] += float(r.get("primaryValue") or 0)
            ranked = sorted(agg.values(), key=lambda a: -a["usd"])
            top = ranked[:6]
            if "ECU" in agg and all(t["iso3"] != "ECU" for t in top):
                top.append(agg["ECU"])
            for t in top:
                origins.append({"iso3": t["iso3"], "name": t["name"], "usd": round(t["usd"]),
                                "share": round(t["usd"] / tot * 100, 2)})

        if sum(v["usd"] for v in hs.values()) > 0:
            by_year[str(y)] = {"hs": hs, "ecuador": ecuador, "origins1801": origins}

    if not by_year:
        return None
    ys = sorted(int(y) for y in by_year)
    return {
        "iso3": iso3,
        "reporterDesc": reporter_desc,
        "years": ys,
        "latestYear": ys[-1],
        "byYear": by_year,
        "retrieved": datetime.date.today().isoformat(),
        "sourceUrl": "https://comtradeplus.un.org/",
        "isDemo": False,
    }


def build_mirror(iso3, code, years):
    """Datos espejo: lo que el resto de países declara haber EXPORTADO a este mercado.
    Se usa cuando el país no publica sus propias importaciones (p. ej. Rusia)."""
    params = {
        "partnerCode": code,
        "period": ",".join(str(y) for y in years),
        "flowCode": "X",
        "cmdCode": ",".join(HS),
        "partner2Code": 0,
        "customsCode": "C00",
        "motCode": 0,
        "includeDesc": "true",
    }
    data = get_json(API, params)
    rows = [r for r in (data.get("data") or [])
            if is_total_row(r) and str(r.get("reporterCode")) not in ("0", "97")]
    by_year = {}
    for y in years:
        yr = [r for r in rows if str(r.get("period")) == str(y)]
        if not yr:
            continue
        hs = {}
        for c in HS:
            cr = [r for r in yr if str(r.get("cmdCode")) == c]
            hs[c] = {"usd": round(sum(float(r.get("primaryValue") or 0) for r in cr)),
                     "kg": round(sum(float(r.get("netWgt") or 0) for r in cr))}
        ecuador = {}
        for c in HS:
            usd = sum(float(r.get("primaryValue") or 0) for r in yr
                      if str(r.get("cmdCode")) == c and r.get("reporterISO") == "ECU")
            if usd > 0 and hs[c]["usd"] > 0:
                ecuador[c] = {"usd": round(usd), "share": round(usd / hs[c]["usd"] * 100, 2)}
        origins = []
        tot = hs["1801"]["usd"]
        if tot > 0:
            agg = {}
            for r in yr:
                if str(r.get("cmdCode")) != "1801":
                    continue
                k = r.get("reporterISO") or str(r.get("reporterCode"))
                a = agg.setdefault(k, {"iso3": k, "name": r.get("reporterDesc") or k, "usd": 0})
                a["usd"] += float(r.get("primaryValue") or 0)
            ranked = sorted(agg.values(), key=lambda a: -a["usd"])
            top = ranked[:6]
            if "ECU" in agg and all(t["iso3"] != "ECU" for t in top):
                top.append(agg["ECU"])
            for t in top:
                origins.append({"iso3": t["iso3"], "name": t["name"], "usd": round(t["usd"]),
                                "share": round(t["usd"] / tot * 100, 2)})
        if sum(v["usd"] for v in hs.values()) > 0:
            by_year[str(y)] = {"hs": hs, "ecuador": ecuador, "origins1801": origins}
    if not by_year:
        return None
    ys = sorted(int(y) for y in by_year)
    return {
        "iso3": iso3,
        "reporterDesc": "los países exportadores",
        "years": ys,
        "latestYear": ys[-1],
        "byYear": by_year,
        "retrieved": datetime.date.today().isoformat(),
        "sourceUrl": "https://comtradeplus.un.org/",
        "mirror": True,
        "isDemo": False,
    }


def main():
    if not KEY:
        sys.exit("Falta la clave COMTRADE_KEY (añádela en GitHub > Settings > Secrets).")

    try:
        with open(OUT, encoding="utf-8") as f:
            prev_file = json.load(f)
        previous = prev_file.get("markets", {})
        previous_selection = prev_file.get("selection")
    except Exception:
        previous, previous_selection = {}, None

    this_year = datetime.date.today().year
    years = [this_year - 3, this_year - 2, this_year - 1]
    codes, info = reporter_codes()

    # 1) Elegir los 40 mercados según dónde vende Ecuador (se guarda dentro de imports.json)
    try:
        sel = select_markets(years, info)
    except Exception as e:
        print("Aviso: no se pudo recalcular la lista de mercados:", e)
        sel = None
    if sel:
        sel["generated"] = datetime.date.today().isoformat()
    else:
        sel = previous_selection
        print("Se conserva la lista de mercados anterior." if sel else "Se usa la lista inicial de mercados.")
    market_list = [m["iso3"] for m in sel["markets"]] if sel else MARKETS
    print("Mercados:", ", ".join(market_list))
    time.sleep(2)

    # 2) Importaciones de cada mercado
    markets, log = {}, []
    for iso3 in market_list:
        if iso3 not in codes:
            log.append(f"{iso3}: sin código en UN Comtrade")
            continue
        try:
            m = build_market(iso3, codes[iso3], years)
            if not m:  # el país no publica sus importaciones: usar datos espejo
                time.sleep(2)
                m = build_mirror(iso3, codes[iso3], years)
            if m:
                markets[iso3] = m
                log.append(f"{iso3}: ok{' (datos espejo)' if m.get('mirror') else ''} ({', '.join(map(str, m['years']))})")
            elif iso3 in previous:
                markets[iso3] = previous[iso3]  # conservar lo anterior
                log.append(f"{iso3}: sin datos nuevos, se conserva lo anterior")
            else:
                log.append(f"{iso3}: sin datos en UN Comtrade")
        except Exception as e:
            if iso3 in previous:
                markets[iso3] = previous[iso3]
            log.append(f"{iso3}: error ({e}); se conserva lo anterior")
        time.sleep(2)  # respetar el límite de la API gratuita

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump({
            "source": "UN Comtrade (https://comtradeplus.un.org/)",
            "flow": "Importaciones",
            "hs": HS,
            "generated": datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
            "markets": markets,
            "selection": sel,
        }, f, ensure_ascii=False, indent=1)

    print("\n".join(log))
    print(f"\nMercados con datos: {len(markets)} de {len(market_list)}")


if __name__ == "__main__":
    main()
