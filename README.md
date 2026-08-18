# Agente Comex — Automatización de Facturas de Comercio Exterior

Sistema en Python que extrae, valida y exporta información de **Commercial
Invoices (PDF)** de comercio exterior (Baker Hughes, USA/México → Colombia)
usando un LLM (Google Gemini) para el parseo y reglas de negocio propias para
la validación matemática y el saneamiento de datos.

## ¿Qué hace?

1. Lee todos los PDF de una carpeta de entrada (`facturas/` por defecto).
2. Extrae el texto de cada PDF con `pypdf`.
3. Envía el texto a Gemini con un prompt estructurado (rol, tarea, reglas,
   validaciones y schema de salida) para reconstruir la factura en tres
   bloques: **encabezado**, **items** y **pie**.
4. Sanea campos ausentes con valores por defecto explícitos.
5. Normaliza las descripciones de los productos (elimina metadatos como
   Delivery, Order, ECCN, Country of Origin, etc.).
6. Valida matemáticamente cada factura en dos niveles:
   - **Nivel 0/1**: cantidad × precio unitario vs. valor de línea, y suma de
     líneas vs. Net Value.
   - **Nivel 2**: Net Value + Freight + Impuestos vs. Total CPT.
7. Guarda los resultados en una base de datos **SQLite**, evitando reprocesar
   facturas ya existentes (salvo `--reprocesar`).
8. Exporta los resultados en CSV, JSON y Excel (con formato condicional y
   gráfico), tanto por factura individual como consolidados de la corrida.
9. Genera un reporte final en consola con el resumen del procesamiento.

## Requisitos

- Python 3.9+
- Una API key de Google Gemini (`GEMINI_API_KEY`)
- Dependencias: `pandas`, `pydantic`, `pypdf`, `python-dotenv`,
  `google-genai`, `openpyxl`

```bash
pip install pandas pydantic pypdf python-dotenv google-genai openpyxl
```

## Configuración rápida

1. Crea un archivo `.env` junto al script con:
   ```
   GEMINI_API_KEY=tu_clave_de_gemini_aqui
   ```
2. Coloca los PDF de las facturas en la carpeta `facturas/` (o indica otra
   con `--entrada`).
3. Ejecuta:
   ```bash
   python agente_comex.py --todas
   ```
4. Revisa los resultados en la carpeta `salida/`.

## Comandos más usados

| Comando | Qué hace |
|---|---|
| `python agente_comex.py --todas` | Procesa todas las facturas de la carpeta de entrada |
| `python agente_comex.py --cantidad 5` | Procesa como máximo 5 facturas |
| `python agente_comex.py --factura 4638446` | Procesa solo la factura que contenga ese número en el nombre del archivo |
| `python agente_comex.py --factura 4638446 --reprocesar` | Fuerza el reprocesamiento de una factura ya existente en la BD |
| `python agente_comex.py --buscar 4638446` | Consulta una factura ya procesada (no requiere PDF ni API key) |

Para más detalle sobre el uso día a día, ver **MANUAL_DE_USUARIO.md**.
Para el detalle de arquitectura, modelos de datos y flujo interno, ver
**DOCUMENTACION_TECNICA.md**.

## Estructura de salida

```
salida/
├── 4638446/
│   ├── cuerpo.csv        # items de esa factura
│   ├── encabezado.csv    # datos generales de esa factura
│   └── reporte.xlsx      # reporte legible para revisión humana
├── 4638447/
│   └── ...
├── cuerpo_items_facturas.csv     # consolidado de la corrida
├── encabezados_facturas.csv      # consolidado de la corrida
├── facturas_completas.json       # consolidado de la corrida
└── Reporte_Facturas_Comex.xlsx   # consolidado de la corrida
```

## Licencia

Uso interno / privado.
