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


def main():
    descargar_xlsm()
    avances = cargar_avances()
    resumen = calcular_resumen(avances)
    prod = calcular_productividad()

    with open(TEMPLATE_PATH, encoding="utf-8") as f:
        html = f.read()

    reemplazos = {
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
