"""
ZAUS - Analisis semanal de transacciones (Bistrosoft)
=======================================================

QUE HACE ESTE SCRIPT
---------------------
Toma el excel que se exporta desde Bistrosoft en:
    Ventas y Reportes > Reportes > Detalle de Trnx.
y arma automaticamente el resumen que hoy hacemos a mano cada semana:
    - Facturacion total y por dia
    - Facturacion por medio de pago (efectivo, QR, tarjeta, etc.)
    - Ticket promedio y cantidad de tickets
    - Top de productos mas vendidos (por cantidad y por facturacion)
    - Facturacion por hora del dia (para ver picos)
    - Descuentos aplicados y anulaciones (notas de credito)

Guarda todo en un excel nuevo con varias hojas, listo para mirar o
para pegar en la planilla de seguimiento semanal.

COMO SE USA
-----------
Desde la terminal, parado en la carpeta donde esta este archivo:

    python zaus_analisis_semanal.py "ruta/al/archivo/DetalleDeTransacciones_...xlsx"

Si no le pasas ninguna ruta, busca el ultimo archivo que empiece con
"DetalleDeTransacciones" en la carpeta actual.

El resultado se guarda como "resumen_semanal.xlsx" en la misma carpeta.

NOTA PARA MANU (si estas aprendiendo a programar con esto)
------------------------------------------------------------
Este script usa la libreria "pandas", que es LA herramienta estandar
en Python para trabajar con datos tipo tabla (como un excel). La idea
basica de pandas es el "DataFrame": una tabla en memoria que podes
filtrar, agrupar y sumar con pocas lineas, en vez de recorrer fila por
fila a mano.

Cada funcion de abajo hace UNA sola cosa (leer, limpiar, calcular tal
metrica). Si en el futuro Bistrosoft cambia el formato del excel, o
Lu pide una metrica nueva, lo mas probable es que solo tengas que
tocar o agregar una funcion, no reescribir todo el script.

Instalar dependencias (una sola vez):
    pip install pandas openpyxl
"""

import sys
import glob
import pandas as pd


# ----------------------------------------------------------------------
# 1. CARGA DEL ARCHIVO
# ----------------------------------------------------------------------

def encontrar_archivo():
    """Si no nos pasan una ruta, buscamos el ultimo export de Bistrosoft
    en la carpeta actual. glob.glob nos da una lista de archivos que
    matchean un patron (como buscar con * en la terminal)."""
    candidatos = sorted(glob.glob("DetalleDeTransacciones*.xlsx"))
    if not candidatos:
        raise FileNotFoundError(
            "No encontre ningun archivo 'DetalleDeTransacciones*.xlsx' en esta carpeta. "
            "Pasa la ruta como argumento: python zaus_analisis_semanal.py archivo.xlsx"
        )
    return candidatos[-1]


def cargar_transacciones(ruta):
    """Lee el excel de Bistrosoft y devuelve un DataFrame limpio.

    El export de Bistrosoft trae 7 filas de encabezado (titulo, nombre
    del comercio, fechas, etc.) antes de la tabla real -> por eso
    'skiprows=7'. Si Bistrosoft cambia el formato del export y esto
    deja de andar, lo primero que hay que revisar es este numero:
    abrir el excel y contar cuantas filas hay antes de la fila que
    dice 'Comercio, Fecha, Hora, Nro Ticket...'.
    """
    df = pd.read_excel(ruta, skiprows=7)

    # Bistrosoft a veces guarda los tipos de transaccion con espacios
    # de mas (ej: "- ITEM " en vez de "- ITEM"). .str.strip() saca los
    # espacios al principio/final para que los filtros de mas abajo
    # no fallen silenciosamente por un espacio invisible.
    df["Tipo Transacción"] = df["Tipo Transacción"].str.strip()

    # La columna Fecha viene como texto "20-07-2026" -> la convertimos
    # a un tipo fecha real, asi despues podemos agrupar por dia,
    # sacar el dia de la semana, ordenar cronologicamente, etc.
    df["Fecha_dt"] = pd.to_datetime(df["Fecha"], format="%d-%m-%Y")

    # La columna Hora viene como "timedelta" (segundos desde las 0:00).
    # Sacamos la hora como numero entero (0-23) para poder agrupar por hora.
    df["Hora_int"] = df["Hora"].apply(lambda t: int(t.total_seconds() // 3600))

    return df


# ----------------------------------------------------------------------
# 2. LIMPIEZA
# ----------------------------------------------------------------------
#
# OJO / importante: probamos sacar filas "duplicadas" automaticamente
# aca (con df.drop_duplicates()) y NO lo dejamos, aposta. Si alguien
# pide 2 panchos iguales en el mismo pedido, Bistrosoft a veces carga
# DOS filas identicas (mismo ticket, mismo producto, mismo precio, en
# vez de una fila con cantidad=2). Esas dos filas son ventas reales,
# no un error de exportacion -> si las borraramos, estariamos
# perdiendo ventas de verdad. Por eso este script no borra nada solo;
# si en algun momento sospechan que el archivo tiene un problema real
# de exportacion duplicada (por ejemplo, el mismo Nro Ticket aparece
# con una fila "Comanda"/"Venta" repetida dos veces con exactamente
# la misma hora), hay que revisarlo a mano antes de sacar algo.


# ----------------------------------------------------------------------
# 3. CALCULOS
# ----------------------------------------------------------------------
# Bistrosoft registra cada venta con varias filas: una fila "cabecera"
# (tipo "Comanda" o "Venta") con el total del ticket, y despues una
# fila por cada producto vendido en ese ticket (tipo "- ITEM" o
# "- COMBO"). Por eso casi todos los calculos de aca abajo primero
# separan "cabeceras" (para totales de facturacion) de "items"
# (para saber que productos se vendieron).

TIPOS_CABECERA = ["Comanda", "Venta"]
TIPOS_ITEM = ["- ITEM", "- COMBO"]


def facturacion_por_dia(df):
    cabeceras = df[df["Tipo Transacción"].isin(TIPOS_CABECERA)]
    resumen = cabeceras.groupby(cabeceras["Fecha_dt"].dt.date).agg(
        facturacion=("Precio", "sum"),
        tickets=("Nro Ticket", "nunique"),
    )
    resumen["ticket_promedio"] = (resumen["facturacion"] / resumen["tickets"]).round(0)
    return resumen


def facturacion_por_medio_pago(df):
    cabeceras = df[df["Tipo Transacción"].isin(TIPOS_CABECERA)]
    resumen = cabeceras.groupby("Medio de Pago")["Precio"].sum().sort_values(ascending=False)
    resumen = resumen.to_frame("facturacion")
    resumen["porcentaje"] = (resumen["facturacion"] / resumen["facturacion"].sum() * 100).round(1)
    return resumen


def facturacion_por_hora(df):
    cabeceras = df[df["Tipo Transacción"].isin(TIPOS_CABECERA)]
    return cabeceras.groupby("Hora_int")["Precio"].sum().to_frame("facturacion")


def top_productos(df, top_n=15):
    items = df[df["Tipo Transacción"].isin(TIPOS_ITEM)].copy()
    # El nombre del producto viene con el codigo interno al final,
    # ej: "Pancho Napolitano (101111503204)" -> lo sacamos con una
    # expresion regular para que no se separen ventas del mismo
    # producto por tener codigos ligeramente distintos.
    items["Producto"] = items["Item"].str.replace(r"\s*\(\d+\)", "", regex=True).str.strip()
    resumen = items.groupby("Producto").agg(
        cantidad=("Cantidad", "sum"),
        facturacion=("Precio", "sum"),
    )
    return resumen.sort_values("facturacion", ascending=False).head(top_n)

def detectar_posibles_duplicados(df):
    cabeceras = df[df["Tipo Transacción"].isin(TIPOS_CABECERA)]
    conteo_tickets = cabeceras.groupby("Nro Ticket").size()  
    sospechosos = conteo_tickets[conteo_tickets > 1]
    return sospechosos

def descuentos_y_anulaciones(df):
    descuentos = df[df["Tipo Transacción"].str.contains("DESCUENTO", na=False)]
    total_descuentos = descuentos["Precio"].sum()

    anulaciones = df[df["Comentarios"].astype(str).str.contains("NCB", na=False)]

    return {
        "total_descuentos": total_descuentos,
        "cantidad_descuentos": len(descuentos),
        "anulaciones": anulaciones[["Fecha", "Hora", "Nro Ticket", "Precio", "Comentarios"]],
    }


def resumen_general(df):
    cabeceras = df[df["Tipo Transacción"].isin(TIPOS_CABECERA)]
    return {
        "facturacion_total": cabeceras["Precio"].sum(),
        "tickets": cabeceras["Nro Ticket"].nunique(),
        "ticket_promedio": round(cabeceras["Precio"].sum() / cabeceras["Nro Ticket"].nunique(), 0),
        "dias_con_ventas": cabeceras["Fecha_dt"].dt.date.nunique(),
    }


# ----------------------------------------------------------------------
# 4. EXPORTAR TODO A UN EXCEL
# ----------------------------------------------------------------------

def guardar_excel(df, ruta_salida="resumen_semanal.xlsx"):
    """ExcelWriter nos deja escribir varias hojas ('sheets') en un
    mismo archivo, una por cada tabla que calculamos arriba."""
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        facturacion_por_dia(df).to_excel(writer, sheet_name="Por dia")
        facturacion_por_medio_pago(df).to_excel(writer, sheet_name="Por medio de pago")
        facturacion_por_hora(df).to_excel(writer, sheet_name="Por hora")
        top_productos(df).to_excel(writer, sheet_name="Top productos")
        descuentos_y_anulaciones(df)["anulaciones"].to_excel(writer, sheet_name="Anulaciones", index=False)
    print(f"\nListo. Resumen guardado en: {ruta_salida}")


# ----------------------------------------------------------------------
# 5. PUNTO DE ENTRADA (lo que corre cuando ejecutas el script)
# ----------------------------------------------------------------------

def main():
    # sys.argv es la lista de argumentos que le pasaste al script en
    # la terminal. sys.argv[0] es siempre el nombre del script, asi
    # que el primer argumento real (la ruta del excel) es sys.argv[1].
    if len(sys.argv) > 1:
        ruta = sys.argv[1]
    else:
        ruta = encontrar_archivo()
        print(f"No se paso ningun archivo, usando el mas reciente: {ruta}")

    df = cargar_transacciones(ruta)

    gen = resumen_general(df)
    desc = descuentos_y_anulaciones(df)

    print("\n=== RESUMEN GENERAL ===")
    print(f"Facturacion total:   ${gen['facturacion_total']:,.0f}".replace(",", "."))
    print(f"Tickets:             {gen['tickets']}")
    print(f"Ticket promedio:     ${gen['ticket_promedio']:,.0f}".replace(",", "."))
    print(f"Dias con ventas:     {gen['dias_con_ventas']}")
    print(f"Descuentos totales:  ${desc['total_descuentos']:,.0f}".replace(",", "."))
    print(f"Anulaciones:         {len(desc['anulaciones'])}")

    print("\n=== FACTURACION POR DIA ===")
    print(facturacion_por_dia(df))

    print("\n=== FACTURACION POR MEDIO DE PAGO ===")
    print(facturacion_por_medio_pago(df))

    print("\n=== TOP 10 PRODUCTOS ===")
    print(top_productos(df, top_n=10))

    duplicados = detectar_posibles_duplicados(df)
    print("\n=== POSIBLES DUPLICADOS (mismo ticket repetido) ===")
    if len(duplicados) == 0:
        print("No se encontraron tickets sospechosos.")
    else:
        print(duplicados)

    guardar_excel(df)

def detectar_posibles_duplicados(df):
    cabeceras = df[df["Tipo Transacción"].isin(TIPOS_CABECERA)]
    conteo_tickets = cabeceras.groupby("Nro Ticket").size()
    tickets_sospechosos = conteo_tickets[conteo_tickets > 1].index
    detalle = cabeceras[cabeceras["Nro Ticket"].isin(tickets_sospechosos)]
    return detalle[["Fecha", "Hora", "Nro Ticket", "Precio"]].sort_values("Nro Ticket")

if __name__ == "__main__":
    main()

