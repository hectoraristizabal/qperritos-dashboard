"""
Qperritos - Extractor de ventas Siigo API
Extrae todas las facturas, limpia y transforma los datos,
y genera el archivo Excel listo para Power BI.

Uso: python siigo_extractor.py
Requisitos: pip install requests pandas openpyxl
"""
import os
import requests
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import pytz

# ─────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────
# Por esto:
USERNAME   = os.environ.get("SIIGO_USERNAME")
ACCESS_KEY = os.environ.get("SIIGO_ACCESS_KEY")
PARTNER_ID = os.environ.get("PARTNER_ID", "PowerBI")

colombia = pytz.timezone("America/Bogota")
hoy = datetime.now(colombia)

FECHA_INICIO = "2026-05-01"
FECHA_FIN    = (hoy + timedelta(days=1)).strftime("%Y-%m-%d")

print(f"📅 Extrayendo desde {FECHA_INICIO} hasta {FECHA_FIN}")


# ─────────────────────────────────────────
# 1. AUTENTICACIÓN
# ─────────────────────────────────────────
def get_token():
    r = requests.post(
        "https://api.siigo.com/auth",
        headers={"Content-Type": "application/json", "Partner-Id": PARTNER_ID},
        json={"username": USERNAME, "access_key": ACCESS_KEY},
        timeout=15
    )
    r.raise_for_status()
    print("✅ Token obtenido")
    return r.json()["access_token"]


# ─────────────────────────────────────────
# 2. EXTRACCIÓN — todas las facturas
# ─────────────────────────────────────────
def extraer_facturas(token):
    headers = {"Authorization": f"Bearer {token}", "Partner-Id": PARTNER_ID}
    todas = []
    page  = 1

    while True:
        r = requests.get(
            "https://api.siigo.com/v1/invoices",
            headers=headers,
            params={
                "date_start": FECHA_INICIO,
                "date_end":   FECHA_FIN,
                "page":          page,
                "page_size":     100
            },
            timeout=15
        )
        r.raise_for_status()
        data = r.json()
        resultados = data.get("results", [])

        if not resultados:
            break

        for f in resultados:
            for item in f.get("items", []):
                todas.append({
                    "factura_id":  f["name"],
                    "fecha":       f["date"],
                    "hora":        f["metadata"]["created"][11:16],
                    "hora_Militar": f["metadata"]["created"][11:13],  # Solo HH en formato 24h (00-23)
                    "forma_pago":  f["payments"][0]["name"] if f.get("payments") else "",
                    "producto_cod":  item["code"],
                    "producto":      item["description"],
                    "cantidad":      item["quantity"],
                    "precio_unit":   item["price"],
                    "total_item":    item["total"],
                    "bodega_id":     item.get("warehouse", {}).get("id", ""),
                    "bodega_nombre": item.get("warehouse", {}).get("name", ""),
                })

        print(f"  Página {page} — {len(todas)} registros acumulados")

        if not data.get("_links", {}).get("next"):
            break
        page += 1

    print(f"✅ Extracción completa: {len(todas)} registros")
    return pd.DataFrame(todas)


# ─────────────────────────────────────────
# 3. LIMPIEZA Y TRANSFORMACIÓN
# ─────────────────────────────────────────
def transformar(df):

    # Acortar nombres de productos
    df["producto"] = df["producto"].replace({
        "Coca Cola Original": "Coca Original",
        "Coca Cola Zero":     "Coca Zero",
    })

    # Todas las saborizadas — detecta cualquier nombre que contenga "Saborizada"
    df["producto"] = df["producto"].apply(
        lambda x: x.replace("Saborizada ", "Sab. ") if "Saborizada" in str(x) else x
    )
    # Factura_E: True si es factura electrónica (empieza con FV-3)
    df["factura_E"] = df["factura_id"].str.startswith("FV-3")

    # Acortar nombres de formas de pago
    df["forma_pago"] = df["forma_pago"].replace({
        "Qr Banco Bogotá":        "QR Bogotá",
        "Tarjeta Débito":         "T. Débito",
        "Tarjeta Crédito":        "T. Crédito",
    })


    # precio_base: precio unitario sin impuesto
    # Si es FV-3 (electrónica): total_item / cantidad / 1.08
    # Si es FV-2 (no electrónica): total_item / cantidad
    df["precio_base"] = df["total_item"] / df["cantidad"]
    df["precio_base"] = df.apply(
        lambda r: round(r["precio_base"] / 1.08, 2) if r["factura_E"]
                  else round(r["precio_base"], 2),
        axis=1
    )

    # impuesto: 8% del precio_base solo en facturas electrónicas
    df["impuesto"] = df.apply(
        lambda r: round(r["precio_base"] * 0.08, 2) if r["factura_E"] else 0.0,
        axis=1
    )

    # total_venta: venta sin impuesto
    df["total_venta"] = (df["cantidad"] * df["precio_base"]).round(2)

    # total_impuestos: impuesto total del ítem
    df["total_impuestos"] = (df["cantidad"] * df["impuesto"]).round(2)

    # Eliminar columnas intermedias ya no necesarias
    df = df.drop(columns=["precio_unit", "total_item"])

    # Columnas de fecha
    fecha = pd.to_datetime(df["fecha"])
    df["año"]            = fecha.dt.year
    df["mes"]            = fecha.dt.month
    df["dia"]            = fecha.dt.day
    dias_es = {
        "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
        "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
    }
    df["dia_semana"] = fecha.dt.day_name().map(dias_es)

    # Reordenar columnas
    df = df[[
        "factura_id", "año", "mes", "dia", "dia_semana",
        "fecha", "hora","hora_Militar", "bodega_id", "bodega_nombre",
        "forma_pago", "producto_cod", "producto", "cantidad",
        "factura_E", "precio_base", "impuesto", "total_venta", "total_impuestos"
    ]]

    print(f"✅ Transformación completa")
    print(f"   Facturas electrónicas (FV-3): {df[df['factura_E']]['factura_id'].nunique()}")
    print(f"   Facturas normales     (FV-2): {df[~df['factura_E']]['factura_id'].nunique()}")
    print(f"   Total ventas (sin impuesto):  ${df['total_venta'].sum():,.0f}")
    print(f"   Total impuestos:              ${df['total_impuestos'].sum():,.0f}")

    print(f"Fecha máxima en datos: {df['fecha'].max()}")
    print(f"Fecha mínima en datos: {df['fecha'].min()}")
    print(f"Facturas de hoy: {len(df[df['fecha'] == FECHA_FIN])}")
    return df


# ─────────────────────────────────────────
# 4. GUARDAR EXCEL
# ─────────────────────────────────────────
def guardar_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ventas"

    hdr_fill = PatternFill("solid", start_color="1F4E79")
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    alt_fill = PatternFill("solid", start_color="EBF3FB")
    thin     = Side(style="thin", color="BFBFBF")
    border   = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols_moneda = ["precio_base","impuesto","total_venta","total_impuestos"]

    # Encabezados
    for col_idx, col in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col.upper().replace("_", " "))
        cell.fill      = hdr_fill
        cell.font      = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border

    # Datos
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font   = Font(name="Arial", size=10)
            cell.border = border
            if row_idx % 2 == 0:
                cell.fill = alt_fill
            col_name = df.columns[col_idx - 1]
            if col_name in cols_moneda:
                cell.number_format = "#,##0.00"
            elif col_name == "cantidad":
                cell.number_format = "#,##0"

    # Ancho columnas y filtros
    for col_idx, col in enumerate(df.columns, 1):
        max_len = max(len(col), df[col].astype(str).str.len().max())
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 28)
    ws.freeze_panes  = "A2"
    ws.auto_filter.ref = ws.dimensions

    archivo = "ventas_qperritos.xlsx"
    wb.save(archivo)
    print(f"✅ Archivo guardado: {archivo}")
    print(f"   Filas: {len(df)} | Columnas: {len(df.columns)}")
    return archivo


# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("  QPERRITOS — Extractor Siigo API")
    print(f"  Período: {FECHA_INICIO} → {FECHA_FIN}")
    print("=" * 50)

    token = get_token()
    df    = extraer_facturas(token)
    df    = transformar(df)
    guardar_excel(df)

    print("\n✅ Proceso completado exitosamente")
