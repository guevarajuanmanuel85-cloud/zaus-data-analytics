import pandas as pd

archivo = pd.read_excel("/Users/manuel/Desktop/ANALISIS /Consumo_Bistrosoft_ZAUS.xlsx", sheet_name="VENTAS DIARIAS", header=2)

datos_limpios = archivo[pd.to_datetime(archivo["Fecha"], errors="coerce").notna()]

punto_equilibrio = 106700

total_semana = datos_limpios["Total Ventas ($)"].sum()
promedio = datos_limpios["Total Ventas ($)"].mean()
mejor_dia = datos_limpios["Total Ventas ($)"].max()
peor_dia = datos_limpios["Total Ventas ($)"].min()

print("Total de la semana:", total_semana)
print("Promedio diario:", promedio)
print("Mejor día:", mejor_dia)
print("Peor día:", peor_dia)

print("")
print("Detalle por día:")

for indice, fila in datos_limpios.iterrows():
    fecha = fila["Fecha"]
    venta = fila["Total Ventas ($)"]
    if venta >= punto_equilibrio:
        print(fecha, "-", venta, "- OK, cubrió equilibrio")
    else:
        print(fecha, "-", venta, "- no llegó al equilibrio")

import matplotlib.pyplot as plt

plt.bar(datos_limpios["Fecha"].astype(str), datos_limpios["Total Ventas ($)"])
plt.xticks(rotation=45)
plt.title("Ventas por día - Zaus")
plt.xlabel("Fecha")
plt.ylabel("Total vendido ($)")
plt.tight_layout()
plt.savefig("ventas_semana.png")

print("Gráfico guardado como ventas_semana.png")

producto = {"nombre": "Pancho Napolitano", "precio": 7200, "categoria": "Panchos Especiales"}

print(producto["nombre"])
print(producto["precio"])

menu = [
    {"nombre": "Pancho Viena", "precio": 5500, "categoria": "Panchos Simples"},
    {"nombre": "Pancho Napolitano", "precio": 7200, "categoria": "Panchos Especiales"},
    {"nombre": "Pancho Big ZAUS", "precio": 7800, "categoria": "Panchos Especiales"},
    {"nombre": "Gran Zaus", "precio": 13500, "categoria": "Sandwiches"},
    {"nombre": "Coca Cola 600ml", "precio": 3400, "categoria": "Bebidas"}
]

for producto in menu:
    print(producto["nombre"], "-", producto["categoria"], "-", "$" + str(producto["precio"]))

def calcular_total_pedido(items):
    total = 0
    for item in items:
        total = total + item["precio"]
    return total

pedido_cliente = [
    {"nombre": "Pancho Napolitano", "precio": 7200},
    {"nombre": "Coca Cola 600ml", "precio": 3400},
    {"nombre": "Papas Fritas", "precio": 2500}
]

total_pedido = calcular_total_pedido(pedido_cliente)
print("Total del pedido:", total_pedido)

import pandas as pd

menu_df = pd.DataFrame(menu)

precio_promedio_por_categoria = menu_df.groupby("categoria")["precio"].mean()

print(precio_promedio_por_categoria)

costos = pd.DataFrame([
    {"nombre": "Pancho Napolitano", "costo": 3200},
    {"nombre": "Pancho Viena", "costo": 2400},
    {"nombre": "Coca Cola 600ml", "costo": 1800}
])

menu_con_costos = pd.merge(menu_df, costos, on="nombre")

menu_con_costos["margen"] = menu_con_costos["precio"] - menu_con_costos["costo"]

print(menu_con_costos[["nombre", "precio", "costo", "margen"]])

precio_texto = "siete mil doscientos"

try:
    precio_numero = int(precio_texto)
    print("El precio es:", precio_numero)
except:
    print("No se pudo convertir ese valor a número")

ventas_crudas = ["76400", "0", "cincuenta mil", "86050", "104150"]

ventas_validas = []

for valor in ventas_crudas:
    try:
        numero = int(valor)
        ventas_validas.append(numero)
    except:
        print("Valor inválido, se ignoró:", valor)

print("Ventas válidas:", ventas_validas)
print("Total:", sum(ventas_validas))

import sqlite3

conexion = sqlite3.connect("zaus.db")

datos_limpios.to_sql("ventas", conexion, if_exists="replace", index=False)

resultado = pd.read_sql("SELECT Fecha, [Total Ventas ($)] FROM ventas ORDER BY [Total Ventas ($)] DESC", conexion)
print(resultado)

menu_df.to_sql("menu", conexion, if_exists="replace", index=False)

resultado = pd.read_sql("SELECT categoria, AVG(precio) FROM menu GROUP BY categoria", conexion)
print(resultado)

costos.to_sql("costos", conexion, if_exists="replace", index=False)

resultado = pd.read_sql("""
SELECT menu.nombre, menu.precio, costos.costo
FROM menu
JOIN costos ON menu.nombre = costos.nombre
""", conexion)

print(resultado)
