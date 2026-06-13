"""
clean_tiktok.py
Limpia el extracto de pedidos de TikTok para cargarlo en el ERP.

Que hace:
1. Abre una ventana para elegir el archivo (Excel o CSV) descargado de TikTok.
2. Quita la segunda fila de cabecera (la fila de explicaciones).
3. Se queda solo con las columnas necesarias y las renombra en espanol.
4. Limpia los importes (quita "EUR", convierte a numero).
5. Calcula Precio (sin IVA, /1.21) e IVA (21%).
6. Convierte Shipped Time a Fecha_Recogida en formato DD/MM/YYYY.
7. Ordena por Fecha_Recogida y ID_Pedido.
8. Elimina filas sin ID_Pedido y avisa si falta Qty.
9. Guarda un archivo nuevo "..._LIMPIO.xlsx" junto al original.
"""

import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------
# Configuracion
# ----------------------------------------------------------------------
# nombre en TikTok -> nombre para el ERP
RENOMBRAR = {
    "Order ID": "ID_Pedido",
    "Seller SKU": "Codigo_Producto",
    "Quantity": "Qty",
    "SKU Unit Original Price": "Precio_Plataforma",
    "SKU Platform Discount": "Descuento_Plataforma",
    "SKU Seller Discount": "Descuento_Vendedor",
    "Original Shipping Fee": "Precio_Transporte",
    "Shipping Fee Seller Discount": "Descuento_Vendedor_Transporte",
    "Order Amount": "Precio_Cliente",
    "Shipped Time": "Fecha_Recogida",
    "Country": "Pais",
}

COLUMNAS_NECESARIAS = list(RENOMBRAR.keys())

COLUMNAS_DINERO = [
    "Precio_Plataforma",
    "Descuento_Plataforma",
    "Descuento_Vendedor",
    "Precio_Transporte",
    "Descuento_Vendedor_Transporte",
    "Precio_Cliente",
]

ORDEN_FINAL = [
    "Pais",
    "ID_Pedido",
    "Fecha_Recogida",
    "Codigo_Producto",
    "Qty",
    "Precio",
    "IVA",
    "Precio_Plataforma",
    "Precio_Cliente",
    "Precio_Transporte",
    "Descuento_Vendedor",
    "Descuento_Vendedor_Transporte",
    "Descuento_Plataforma",
]

IVA_TIPO = 0.21  # 21%


# ----------------------------------------------------------------------
# Funciones de limpieza
# ----------------------------------------------------------------------
def limpiar_dinero(valor):
    """Convierte 'EUR 12.34', '12,34 EUR', '€12.34' ... en el numero 12.34"""
    if pd.isna(valor):
        return None
    texto = str(valor)
    texto = re.sub(r"(?i)eur", "", texto)
    texto = texto.replace("€", "").replace("\u00a0", "").strip()
    texto = re.sub(r"[^\d,.\-]", "", texto)
    if texto in ("", "-", ".", ","):
        return None
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return None


def limpiar_order_id(valor):
    """ID_Pedido como texto limpio (evita notacion cientifica y '.0' de Excel)."""
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto if texto else None


# ----------------------------------------------------------------------
# Lectura robusta de Excel
# ----------------------------------------------------------------------
def leer_excel_robusto(ruta):
    """
    TikTok manda el Excel de dos formas distintas:
      - Una "normal" que lee bien openpyxl.
      - Una "rara" en la que openpyxl solo ve la primera columna
        y hace falta el motor calamine.
    Esta funcion prueba ambos motores y se queda con el que
    devuelva mas de una columna (es decir, el que lee bien).
    """
    candidatos = []
    for motor in ("calamine", "openpyxl"):
        try:
            df = pd.read_excel(ruta, dtype=str, engine=motor)
            candidatos.append((df.shape[1], df))
        except Exception:
            pass

    if not candidatos:
        # ultimo recurso: dejar que pandas elija el motor por defecto
        return pd.read_excel(ruta, dtype=str)

    # quedarse con el que tenga mas columnas
    candidatos.sort(key=lambda x: x[0], reverse=True)
    return candidatos[0][1]


# ----------------------------------------------------------------------
# Programa principal
# ----------------------------------------------------------------------
def main():
    raiz = tk.Tk()
    raiz.withdraw()

    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo descargado de TikTok",
        filetypes=[
            ("Excel o CSV", "*.xlsx *.xls *.csv"),
            ("Todos los archivos", "*.*"),
        ],
    )
    if not archivo:
        return

    ruta = Path(archivo)

    # ---- Leer (todo como texto para no perder los ID de pedido) ----
    try:
        if ruta.suffix.lower() == ".csv":
            df = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig")
        else:
            df = leer_excel_robusto(ruta)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el archivo:\n{e}")
        return

    # ---- Quitar la segunda fila de cabecera (explicaciones) ----
    if len(df) > 0:
        valor = str(df.iloc[0].get("Order ID", "")).strip()
        if not re.fullmatch(r"\d{6,}", valor.replace(".0", "")):
            df = df.iloc[1:].reset_index(drop=True)

    # ---- Comprobar columnas ----
    faltan = [c for c in COLUMNAS_NECESARIAS if c not in df.columns]
    if faltan:
        messagebox.showerror(
            "Faltan columnas",
            "El archivo no tiene estas columnas:\n\n- " + "\n- ".join(faltan)
            + "\n\nColumnas encontradas:\n" + ", ".join(df.columns),
        )
        return

    df = df[COLUMNAS_NECESARIAS].copy()
    df = df.rename(columns=RENOMBRAR)
    filas_iniciales = len(df)

    # ---- Limpieza ----
    df["ID_Pedido"] = df["ID_Pedido"].apply(limpiar_order_id)
    df = df[df["ID_Pedido"].notna()].reset_index(drop=True)
    filas_sin_id = filas_iniciales - len(df)

    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")
    filas_sin_qty = int(df["Qty"].isna().sum())
    df["Qty"] = df["Qty"].astype("Int64")

    for col in COLUMNAS_DINERO:
        df[col] = df[col].apply(limpiar_dinero)

    # fecha real (para ordenar bien) y luego texto DD/MM/YYYY
    fecha_real = pd.to_datetime(df["Fecha_Recogida"], errors="coerce", dayfirst=False)
    df["Fecha_Recogida"] = fecha_real.dt.strftime("%d/%m/%Y")

    df["Pais"] = df["Pais"].fillna("").astype(str).str.strip()

    # ---- Calculos: Precio sin IVA e IVA ----
    base = df["Precio_Cliente"] / (1 + IVA_TIPO)
    df["Precio"] = base.round(2)
    df["IVA"] = (base * IVA_TIPO).round(2)

    # ---- Ordenar por fecha y pedido ----
    df["_orden_fecha"] = fecha_real
    df = df.sort_values(by=["_orden_fecha", "ID_Pedido"], na_position="last")
    df = df.drop(columns=["_orden_fecha"]).reset_index(drop=True)

    # ---- Orden final de columnas ----
    df = df[ORDEN_FINAL]

    # ---- Guardar ----
    salida = ruta.with_name(ruta.stem + "_LIMPIO.xlsx")
    try:
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Pedidos")
            hoja = writer.sheets["Pedidos"]
            # forzar ID_Pedido (columna B) como texto en Excel
            for celda in hoja["B"]:
                celda.number_format = "@"
    except PermissionError:
        messagebox.showerror(
            "Error",
            f"No se pudo guardar:\n{salida}\n\n"
            "Probablemente el archivo esta abierto en Excel. Cierralo y vuelve a intentarlo.",
        )
        return
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar el archivo:\n{e}")
        return

    # ---- Resumen ----
    mensaje = (
        f"Archivo creado:\n{salida}\n\n"
        f"Pedidos procesados: {len(df)}\n"
        f"Filas eliminadas sin ID_Pedido: {filas_sin_id}"
    )
    if filas_sin_qty > 0:
        mensaje += (
            f"\n\nATENCION: {filas_sin_qty} fila(s) sin Qty. "
            "Revisalas en el archivo antes de cargarlo al ERP."
        )
    messagebox.showinfo("Listo", mensaje)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error inesperado", str(e))
        sys.exit(1)