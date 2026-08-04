"""
Regenera index.html a partir de:
  - el .xlsm de circuitos/productividad (descargado de OneDrive/SharePoint)
  - el CSV de lotes publicado (se lee directo en el navegador, no hace falta acá)

Variables de entorno esperadas (se configuran como Secrets en GitHub Actions):
  XLSM_URL   -> link de descarga directa del .xlsm (con &download=1 al final)
"""
import os
import sys
from datetime import datetime, timezone

import requests
import pandas as pd
from openpyxl import load_workbook

XLSM_URL = os.environ["XLSM_URL"]
XLSM_LOCAL_PATH = "planilla.xlsm"
TEMPLATE_PATH = "template.html"
OUTPUT_PATH = "index.html"

LOTES_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTx_R6vrK_i25nuoquIzBLw2qEBtL-LQaCeu_VZ7eA1OrbI1HdkqEG-ESOcEPnZLPZeNGFc0F1icxVg/pub?gid=0&single=true&output=csv"
TAREAS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTx_R6vrK_i25nuoquIzBLw2qEBtL-LQaCeu_VZ7eA1OrbI1HdkqEG-ESOcEPnZLPZeNGFc0F1icxVg/pub?output=csv&gid=442614997"


def descargar_xlsm():
    print("Descargando planilla desde SharePoint...")
    resp = requests.get(XLSM_URL, timeout=60)
    resp.raise_for_status()
    if len(resp.content) < 10_000:
        # una respuesta sospechosamente chica probablemente es una pagina de login/error, no el xlsm real
        raise RuntimeError(
            f"La descarga vino demasiado chica ({len(resp.content)} bytes) - "
            "probablemente el link dejo de ser publico o cambio. Revisar XLSM_URL."
        )
    with open(XLSM_LOCAL_PATH, "wb") as f:
        f.write(resp.content)
    print(f"OK: {len(resp.content)} bytes descargados.")


def cargar_avances():
    wb = load_workbook(XLSM_LOCAL_PATH, read_only=True, data_only=True)
    ws = wb["2. AVANCES"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c) for c in rows[1]]
    data = [r for r in rows[2:] if r[3] is not None]
    df = pd.DataFrame(data, columns=header)
    return df


def calcular_resumen(df):
    total = len(df)
    en_proceso = (df["Estado General"] == "En proceso").sum()
    no_iniciado = (df["Estado General"] == "No iniciado").sum()
    finalizados = df["Estado General"].astype(str).str.contains("100%").sum()

    regional = (
        df.groupby("Regional")
        .agg(
            circuitos=("Alimentador / Circuito", "count"),
            km_mt_plan=("km MT Plan", "sum"),
            km_mt_real=("km MT Real", "sum"),
            pd_ande_plan=("PD ANDE Plan", "sum"),
            pd_ande_reportado=("PD ANDE Reportado", "sum"),
            pd_tercero_plan=("PD Tercero Plan", "sum"),
        )
        .round(1)
        .sort_values("circuitos", ascending=False)
    )
    return {
        "total": int(total),
        "en_proceso": int(en_proceso),
        "no_iniciado": int(no_iniciado),
        "finalizados": int(finalizados),
        "regional": regional,
    }


def regional_tabla_html(regional_df):
    filas = []
    for reg, r in regional_df.iterrows():
        filas.append(
            f"<tr><td>{reg.title()}</td><td class='num'>{int(r.circuitos)}</td>"
            f"<td class='num'>{r.km_mt_plan:,.0f}</td><td class='num'>{r.km_mt_real:,.0f}</td>"
            f"<td class='num'>{r.pd_ande_plan:,.0f}</td><td class='num'>{r.pd_ande_reportado:,.0f}</td>"
            f"<td class='num'>{r.pd_tercero_plan:,.0f}</td></tr>"
        )
    return "\n".join(filas)


def este_data_js(df):
    este = df[df["Regional"] == "ESTE"][
        ["SSEE", "Alimentador / Circuito", "km MT Plan", "PD ANDE Plan", "PD Tercero Plan", "Estado General"]
    ].fillna(0)
    filas = []
    for _, r in este.iterrows():
        filas.append(
            "[%r,%r,%s,%s,%s,%r]"
            % (r["SSEE"], r["Alimentador / Circuito"], r["km MT Plan"], int(r["PD ANDE Plan"]),
               int(r["PD Tercero Plan"]), r["Estado General"])
        )
    return ",\n".join(filas)


def calcular_productividad():
    wb = load_workbook(XLSM_LOCAL_PATH, read_only=True, data_only=True)
    ws = wb["3.REPORTE_DIARIO"]
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    data = [r for r in rows[1:] if r[0] is not None]
    df = pd.DataFrame(data, columns=header)
    # descartar filas rotas / de error
    df = df[df["Analista EO"].astype(str).str.match(r"^[A-Za-z]+$", na=False)]
    g = df.groupby("Analista EO").agg(
        pd_ande=("PD ANDE Construido", "sum"),
        pd_terceros=("PD TERCEROS Construido", "sum"),
    )
    g = g[(g.pd_ande > 0) | (g.pd_terceros > 0)].sort_values("pd_ande", ascending=False)
    return g


def prod_tabla_html(g):
    filas = []
    for analista, r in g.iterrows():
        filas.append(
            f"<tr><td>{analista}</td><td class='num'>{int(r.pd_ande)}</td>"
            f"<td class='num'>{int(r.pd_terceros)}</td></tr>"
        )
    return "\n".join(filas)


def js_str(s):
    """Escapa un string para insertarlo como literal JS entre comillas simples."""
    return str(s).replace("\\", "\\\\").replace("'", "\\'").replace("\n", " ")


def cargar_lotes():
    df = pd.read_csv(LOTES_URL, skiprows=3, header=None)
    df = df[df[0].notna()]
    df.columns = list(range(df.shape[1]))

    filas = []
    for _, r in df.iterrows():
        filas.append({
            "lote": r[2],
            "regional": str(r[1]).strip() if pd.notna(r[1]) else "",
            "estado_general": str(r[4]).strip() if pd.notna(r[4]) else "",
            "postes": int(r[22]) if pd.notna(r[22]) and str(r[22]).strip() not in ("", "-") else None,
            "estado_lote": str(r[25]).strip() if pd.notna(r[25]) else "",
            "hab8b": str(r[30]).strip() if pd.notna(r[30]) else "",
        })

    total = len(filas)
    tramite = sum(1 for f in filas if f["estado_general"] == "Realizado")
    cola = total - tramite
    suma_postes = sum(f["postes"] for f in filas if f["postes"])

    reg_counts = {}
    for f in filas:
        if f["estado_general"] == "Realizado":
            k = f["regional"] or "Sin asignar"
            reg_counts[k] = reg_counts.get(k, 0) + 1

    est_counts = {}
    for f in filas:
        k = f["estado_lote"] or "Sin iniciar"
        est_counts[k] = est_counts.get(k, 0) + 1

    return {
        "filas": filas, "total": total, "tramite": tramite, "cola": cola,
        "suma_postes": suma_postes, "reg_counts": reg_counts, "est_counts": est_counts,
    }


def lotes_data_js(lotes):
    filas = []
    for f in lotes["filas"]:
        filas.append(
            "{lote:%r,regional:%r,estado_general:%r,postes:%s,estado_lote:%r,hab8b:%r}"
            % (f["lote"], f["regional"], f["estado_general"],
               f["postes"] if f["postes"] is not None else "null",
               f["estado_lote"], f["hab8b"])
        )
    return ",\n".join(filas)


def bars_html(counts, color_map=None, default_color="#4d9fef"):
    color_map = color_map or {}
    if not counts:
        return ""
    maximo = max(counts.values())
    filas = []
    for k, v in sorted(counts.items(), key=lambda kv: -kv[1]):
        color = color_map.get(k, default_color)
        pct = round(v / maximo * 100) if maximo else 0
        filas.append(
            f'<div class="bar-row"><div class="lbl">{k}</div>'
            f'<div class="bar-track"><div class="bar-fill" style="width:{pct}%; background:{color}"></div></div>'
            f'<div class="bar-val">{v}</div></div>'
        )
    return "\n".join(filas)


def cargar_tareas():
    df = pd.read_csv(TAREAS_URL, header=0)
    df = df[df.iloc[:, 0].notna()]
    df.columns = list(range(df.shape[1]))

    filas = []
    for _, r in df.iterrows():
        filas.append({
            "item": r[0],
            "area": str(r[1]).strip() if pd.notna(r[1]) else "",
            "tarea": str(r[2]).strip() if pd.notna(r[2]) else "",
            "responsable": str(r[4]).strip() if pd.notna(r[4]) else "",
            "regional": str(r[5]).strip() if pd.notna(r[5]) else "",
            "estado": str(r[6]).strip() if pd.notna(r[6]) else "",
            "reporte": str(r[7]).strip() if pd.notna(r[7]) else "",
        })

    def es_ok(estado):
        return estado in ("Realizado", "Completado")

    total = len(filas)
    ok = sum(1 for f in filas if es_ok(f["estado"]))
    pend = total - ok

    por_persona = {}
    for f in filas:
        if not f["responsable"]:
            continue
        por_persona.setdefault(f["responsable"], {"total": 0, "pend": 0})
        por_persona[f["responsable"]]["total"] += 1
        if not es_ok(f["estado"]):
            por_persona[f["responsable"]]["pend"] += 1

    personas_con_pendientes = sum(1 for p in por_persona.values() if p["pend"] > 0)
    areas = sorted({f["area"] for f in filas if f["area"]})

    return {
        "filas": filas, "total": total, "ok": ok, "pend": pend,
        "por_persona": por_persona, "personas_con_pendientes": personas_con_pendientes,
        "areas": areas,
    }


def tareas_data_js(tareas):
    filas = []
    for f in tareas["filas"]:
        filas.append(
            "{item:%r,area:%r,tarea:%r,responsable:%r,regional:%r,estado:%r,reporte:%r}"
            % (f["item"], f["area"], f["tarea"], f["responsable"], f["regional"], f["estado"], f["reporte"])
        )
    return ",\n".join(filas)


def cards_responsables_html(por_persona):
    filas = []
    for nombre, c in sorted(por_persona.items(), key=lambda kv: -kv[1]["pend"]):
        clase = "c-red" if c["pend"] > 0 else "c-green"
        filas.append(
            f'<div class="kpi {clase}"><div class="label">{nombre}</div>'
            f'<div class="value">{c["pend"]}</div>'
            f'<div class="sub">pendientes de {c["total"]} tareas</div></div>'
        )
    return "\n".join(filas)


def filtros_area_html(areas):
    botones = ['<button data-area="Todas" onclick="setFiltroArea(\'Todas\')" class="filtro-btn activo">Todas</button>']
    for a in areas:
        botones.append(f'<button data-area="{a}" onclick="setFiltroArea(\'{a}\')" class="filtro-btn">{a}</button>')
    return "\n".join(botones)


def main():
    descargar_xlsm()
    avances = cargar_avances()
    resumen = calcular_resumen(avances)
    prod = calcular_productividad()

    print("Descargando control de lotes...")
    lotes = cargar_lotes()
    print(f"OK: {lotes['total']} lotes.")

    print("Descargando pendientes por área...")
    tareas = cargar_tareas()
    print(f"OK: {tareas['total']} tareas.")

    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        html = f.read()

    reemplazos = {
        "{{LOTES_TOTAL}}": str(lotes["total"]),
        "{{LOTES_TRAMITE}}": str(lotes["tramite"]),
        "{{LOTES_COLA}}": str(lotes["cola"]),
        "{{LOTES_POSTES}}": f"{lotes['suma_postes']:,}".replace(",", "."),
        "{{LOTES_BARS_REGIONAL}}": bars_html(lotes["reg_counts"]),
        "{{LOTES_BARS_ESTADO}}": bars_html(lotes["est_counts"], {"Aceptado": "#3ecf8e", "Rechazado": "#f0645f"}, "#f0a94d"),
        "{{LOTES_DATA_JS}}": lotes_data_js(lotes),
        "{{TAREAS_TOTAL}}": str(tareas["total"]),
        "{{TAREAS_OK}}": str(tareas["ok"]),
        "{{TAREAS_PEND}}": str(tareas["pend"]),
        "{{TAREAS_PERSONAS}}": str(tareas["personas_con_pendientes"]),
        "{{TAREAS_CARDS_RESPONSABLES}}": cards_responsables_html(tareas["por_persona"]),
        "{{TAREAS_FILTROS_AREA}}": filtros_area_html(tareas["areas"]),
        "{{TAREAS_DATA_JS}}": tareas_data_js(tareas),
        "{{TOTAL_ALIM}}": str(resumen["total"]),
        "{{EN_PROCESO}}": str(resumen["en_proceso"]),
        "{{NO_INICIADO}}": str(resumen["no_iniciado"]),
        "{{FINALIZADOS}}": str(resumen["finalizados"]),
        "{{REGIONAL_TABLE_ROWS}}": regional_tabla_html(resumen["regional"]),
        "{{REGIONAL_CHART_LABELS}}": str([r.title() for r in resumen["regional"].index]),
        "{{REGIONAL_CHART_DATA}}": str(list(resumen["regional"].circuitos.astype(int))),
        "{{ESTE_DATA_JS}}": este_data_js(avances),
        "{{PROD_TABLE_ROWS}}": prod_tabla_html(prod),
        "{{PROD_CHART_LABELS}}": str(list(prod.index)),
        "{{PROD_CHART_ANDE}}": str(list(prod.pd_ande.astype(int))),
        "{{PROD_CHART_TERC}}": str(list(prod.pd_terceros.astype(int))),
        "{{BUILD_TIMESTAMP}}": datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M UTC"),
    }
    for k, v in reemplazos.items():
        html = html.replace(k, v)

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"OK: {OUTPUT_PATH} regenerado con datos frescos.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR en build.py: {e}", file=sys.stderr)
        sys.exit(1)
