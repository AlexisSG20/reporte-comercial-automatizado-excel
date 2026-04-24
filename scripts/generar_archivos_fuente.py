from pathlib import Path
import random
from datetime import datetime, timedelta

import pandas as pd


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

BASE_DIR = Path(__file__).resolve().parent.parent
OUTPUT_DIR = BASE_DIR / "archivos_fuente"

random.seed(42)


# =========================================================
# CATÁLOGOS BASE
# =========================================================

PRODUCTOS = [
    {"codigo": "TEC001", "producto": "Laptop Lenovo", "categoria": "Tecnología", "precio_unitario": 2800},
    {"codigo": "TEC002", "producto": "Mouse Logitech", "categoria": "Tecnología", "precio_unitario": 75},
    {"codigo": "TEC003", "producto": "Audífonos JBL", "categoria": "Tecnología", "precio_unitario": 220},
    {"codigo": "HOG001", "producto": "Licuadora Oster", "categoria": "Hogar", "precio_unitario": 180},
    {"codigo": "HOG002", "producto": "Juego de Sábanas", "categoria": "Hogar", "precio_unitario": 120},
    {"codigo": "HOG003", "producto": "Sartén Antiadherente", "categoria": "Hogar", "precio_unitario": 95},
    {"codigo": "MOD001", "producto": "Casaca Denim", "categoria": "Moda", "precio_unitario": 160},
    {"codigo": "MOD002", "producto": "Zapatillas Urbanas", "categoria": "Moda", "precio_unitario": 210},
    {"codigo": "MOD003", "producto": "Polo Básico", "categoria": "Moda", "precio_unitario": 55},
    {"codigo": "DEP001", "producto": "Mancuernas 5kg", "categoria": "Deportes", "precio_unitario": 140},
    {"codigo": "DEP002", "producto": "Mat de Yoga", "categoria": "Deportes", "precio_unitario": 85},
    {"codigo": "DEP003", "producto": "Tomatodo Deportivo", "categoria": "Deportes", "precio_unitario": 35},
    {"codigo": "BEL001", "producto": "Perfume Floral", "categoria": "Belleza", "precio_unitario": 170},
    {"codigo": "BEL002", "producto": "Crema Facial", "categoria": "Belleza", "precio_unitario": 60},
    {"codigo": "BEL003", "producto": "Shampoo Nutritivo", "categoria": "Belleza", "precio_unitario": 45},
]

VENDEDORES_POR_SEDE = {
    "Lima": ["Carla Ramos", "Diego Salazar"],
    "Huancayo": ["Andrea Huamán", "Luis Gutiérrez"],
    "Arequipa": ["Valeria Quispe", "José Medina"],
}

CLIENTES = [
    "Ana Torres", "Carlos Rojas", "María López", "Jorge Castillo", "Rosa Mendoza",
    "Kevin Vargas", "Lucía Pérez", "Renato Flores", "Diana Salas", "Pedro Chávez",
    "Sofía Navarro", "Miguel Herrera", "Camila Paredes", "Fernando Ruiz", "Paola Arias",
    "Iván Soto", "Gabriela Núñez", "Julio Campos", "Elena Vera", "Marco Espinoza",
    "Daniela Rivas", "César Salcedo", "Fiorella Díaz", "Bruno Medina", "Karina Aguilar",
    "Andrés Romero", "Mónica Palomino", "Javier Núñez", "Alessandra Peña", "Omar Cárdenas",
]

MEDIOS_PAGO = ["Efectivo", "Tarjeta", "Yape", "Transferencia"]
DESCUENTOS = [0.00, 0.05, 0.10, 0.15]


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================

def generar_fecha_aleatoria(fecha_inicio: datetime, fecha_fin: datetime) -> datetime:
    dias_rango = (fecha_fin - fecha_inicio).days
    dias_aleatorios = random.randint(0, dias_rango)
    return fecha_inicio + timedelta(days=dias_aleatorios)


def elegir_producto() -> dict:
    """
    Damos un poco más de peso a Tecnología y Moda para que no quede
    todo perfectamente uniforme.
    """
    pesos = []
    for p in PRODUCTOS:
        if p["categoria"] in ["Tecnología", "Moda"]:
            pesos.append(1.4)
        elif p["categoria"] in ["Belleza", "Deportes"]:
            pesos.append(1.1)
        else:
            pesos.append(1.0)
    return random.choices(PRODUCTOS, weights=pesos, k=1)[0]


def elegir_descuento() -> float:
    """
    Más ventas sin descuento, menos con descuentos altos.
    """
    return random.choices(
        DESCUENTOS,
        weights=[0.55, 0.25, 0.15, 0.05],
        k=1
    )[0]


def elegir_cliente(prob_vacio: float = 0.0) -> str:
    if random.random() < prob_vacio:
        return ""
    return random.choice(CLIENTES)


def generar_ventas(sede: str, fecha_inicio: str, fecha_fin: str, n_registros: int = 100) -> pd.DataFrame:
    fecha_inicio_dt = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    fecha_fin_dt = datetime.strptime(fecha_fin, "%Y-%m-%d")

    registros = []

    for _ in range(n_registros):
        producto = elegir_producto()
        cantidad = random.randint(1, 5)
        descuento = elegir_descuento()
        total_venta = round(cantidad * producto["precio_unitario"] * (1 - descuento), 2)

        registro = {
            "Fecha": generar_fecha_aleatoria(fecha_inicio_dt, fecha_fin_dt),
            "Sede": sede,
            "Código Producto": producto["codigo"],
            "Producto": producto["producto"],
            "Categoría": producto["categoria"],
            "Vendedor": random.choice(VENDEDORES_POR_SEDE[sede]),
            "Cliente": elegir_cliente(prob_vacio=0.03),
            "Cantidad": cantidad,
            "Precio Unitario": producto["precio_unitario"],
            "Descuento": descuento,
            "Total Venta": total_venta,
            "Medio de Pago": random.choice(MEDIOS_PAGO),
        }
        registros.append(registro)

    df = pd.DataFrame(registros)
    return df

def aplicar_variaciones(df: pd.DataFrame, nombre_archivo: str) -> pd.DataFrame:
    df_mod = df.copy()

    if nombre_archivo == "ventas_huancayo_ene.xlsx":
        df_mod = df_mod.rename(columns={"Precio Unitario": "P. Unitario"})
        columnas = [
            "Sede", "Fecha", "Producto", "Código Producto", "Categoría",
            "Cliente", "Vendedor", "Cantidad", "P. Unitario",
            "Descuento", "Total Venta", "Medio de Pago"
        ]
        df_mod = df_mod[columnas]

    elif nombre_archivo == "ventas_arequipa_ene.xlsx":
        df_mod = df_mod.rename(columns={"Categoría": "Categoria"})
        df_mod["Fecha"] = pd.to_datetime(df_mod["Fecha"]).dt.strftime("%d/%m/%Y")

    elif nombre_archivo == "ventas_lima_feb.xlsx":
        filas_producto = df_mod.sample(frac=0.12, random_state=42).index
        filas_pago = df_mod.sample(frac=0.20, random_state=7).index

        df_mod.loc[filas_producto, "Producto"] = df_mod.loc[filas_producto, "Producto"].apply(lambda x: f" {x} ")
        df_mod.loc[filas_pago, "Medio de Pago"] = df_mod.loc[filas_pago, "Medio de Pago"].replace({
            "Tarjeta": random.choice(["tarjeta", "TARJETA", "Tarjeta "]),
            "Yape": random.choice(["yape", "YAPE", "Yape "]),
            "Transferencia": random.choice(["transferencia", "TRANSFERENCIA"]),
            "Efectivo": random.choice(["efectivo", "EFECTIVO"])
        })

    elif nombre_archivo == "ventas_huancayo_feb.xlsx":
        df_mod = df_mod.rename(columns={"Total Venta": "Total"})
        columnas = [
            "Cliente", "Sede", "Fecha", "Código Producto", "Producto",
            "Categoría", "Vendedor", "Cantidad", "Precio Unitario",
            "Descuento", "Total", "Medio de Pago"
        ]
        df_mod = df_mod[columnas]

    elif nombre_archivo == "ventas_arequipa_feb.xlsx":
        filas_descuento_vacio = df_mod.sample(frac=0.18, random_state=11).index
        filas_cliente_vacio = df_mod.sample(frac=0.05, random_state=21).index
        filas_categoria = df_mod[df_mod["Categoría"] == "Tecnología"].sample(frac=0.5, random_state=5).index
        filas_pago = df_mod.sample(frac=0.15, random_state=17).index

        df_mod.loc[filas_descuento_vacio, "Descuento"] = None
        df_mod.loc[filas_cliente_vacio, "Cliente"] = ""
        df_mod.loc[filas_categoria, "Categoría"] = "Tecnologia"
        df_mod.loc[filas_pago, "Medio de Pago"] = df_mod.loc[filas_pago, "Medio de Pago"].str.lower()

    return df_mod

# =========================================================
# GENERACIÓN DE ARCHIVOS
# =========================================================

def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    archivos = [
        ("ventas_lima_ene.xlsx", "Lima", "2026-01-01", "2026-01-31"),
        ("ventas_huancayo_ene.xlsx", "Huancayo", "2026-01-01", "2026-01-31"),
        ("ventas_arequipa_ene.xlsx", "Arequipa", "2026-01-01", "2026-01-31"),
        ("ventas_lima_feb.xlsx", "Lima", "2026-02-01", "2026-02-28"),
        ("ventas_huancayo_feb.xlsx", "Huancayo", "2026-02-01", "2026-02-28"),
        ("ventas_arequipa_feb.xlsx", "Arequipa", "2026-02-01", "2026-02-28"),
        ("ventas_lima_mar.xlsx", "Lima", "2026-03-01", "2026-03-31"),
        ("ventas_huancayo_mar.xlsx", "Huancayo", "2026-03-01", "2026-03-31"),
        ("ventas_arequipa_mar.xlsx", "Arequipa", "2026-03-01", "2026-03-31"),
    ]

    for nombre_archivo, sede, fecha_inicio, fecha_fin in archivos:
        df = generar_ventas(sede, fecha_inicio, fecha_fin, n_registros=100)
        df = aplicar_variaciones(df, nombre_archivo)

        ruta_salida = OUTPUT_DIR / nombre_archivo
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Ventas", index=False)

        print(f"Archivo generado: {ruta_salida}")


if __name__ == "__main__":
    main()