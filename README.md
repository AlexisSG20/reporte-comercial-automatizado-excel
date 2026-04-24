# Reporte Comercial Automatizado en Excel

Proyecto de análisis de datos desarrollado en Excel para simular un flujo real de consolidación, limpieza y visualización de reportes comerciales enviados por distintas sedes de una empresa retail.

## Empresa ficticia

**Comercial Andina S.A.C.**

## Caso de negocio

La empresa recibe reportes mensuales de ventas desde distintas sedes. Estos archivos llegan separados y con pequeñas diferencias de formato, por lo que se requiere consolidarlos, limpiarlos y transformarlos en un reporte ejecutivo actualizado.

El objetivo del proyecto es automatizar este proceso usando Excel y Power Query, reduciendo el trabajo manual y permitiendo actualizar el dashboard al incorporar nuevos archivos fuente.

## Herramientas utilizadas

- Microsoft Excel
- Power Query
- Tablas dinámicas
- Gráficos dinámicos
- Segmentadores de datos
- Python
- Git y GitHub

## Estructura del proyecto

```text
reporte-comercial-automatizado-excel/
│
├── archivos_fuente/
│   ├── ventas_lima_ene.xlsx
│   ├── ventas_huancayo_ene.xlsx
│   ├── ventas_arequipa_ene.xlsx
│   ├── ventas_lima_feb.xlsx
│   ├── ventas_huancayo_feb.xlsx
│   ├── ventas_arequipa_feb.xlsx
│   ├── ventas_lima_mar.xlsx
│   ├── ventas_huancayo_mar.xlsx
│   └── ventas_arequipa_mar.xlsx
│
├── maestro_excel/
│   └── reporte-comercial-automatizado.xlsx
│
├── scripts/
│   └── generar_archivos_fuente.py
│
├── docs/
│   └── capturas/
│
├── .gitignore
└── README.md
```

## Flujo del proyecto

1. Generación de archivos fuente sintéticos con Python.
2. Simulación de reportes mensuales por sede.
3. Consolidación automática desde carpeta con Power Query.
4. Homologación de columnas con nombres distintos.
5. Limpieza de textos, nulos y valores inconsistentes.
6. Creación de tabla limpia en Excel.
7. Análisis con tablas dinámicas.
8. Construcción de dashboard ejecutivo.
9. Uso de segmentadores para filtrar por sede, categoría y medio de pago.
10. Actualización automática al agregar nuevos archivos fuente.

## Datos simulados

El proyecto trabaja con datos sintéticos de ventas comerciales.

Periodo analizado:

- Enero 2026
- Febrero 2026
- Marzo 2026

Sedes:

- Lima
- Huancayo
- Arequipa

Categorías:

- Tecnología
- Hogar
- Moda
- Deportes
- Belleza

Total de registros consolidados:

- 900 ventas

## Problemas simulados en los archivos fuente

Para representar un escenario más realista, algunos archivos incluyen diferencias controladas:

- nombres de columnas distintos
- orden de columnas diferente
- fechas con formato distinto
- textos con espacios extra
- mayúsculas y minúsculas inconsistentes
- descuentos vacíos
- clientes no registrados
- variaciones en medios de pago

Estas diferencias se corrigen en Power Query.

## Dashboard

El dashboard incluye:

- ventas totales
- cantidad de ventas
- unidades vendidas
- ticket promedio
- descuento promedio
- ventas por sede
- ventas por categoría
- ventas por mes
- ventas por medio de pago
- segmentadores interactivos

## Resultado

El archivo maestro permite actualizar el análisis al agregar nuevos archivos de ventas a la carpeta `archivos_fuente` y presionar **Actualizar todo** en Excel.

Esto simula un flujo de trabajo real de reporting comercial automatizado.

## Autor

Alexis Suasnabar
