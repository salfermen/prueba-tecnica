# Documentación Técnica — Sistema de Automatización de Facturas Comex
## Agente Comex Baker Hughes

**Versión:** 3.0  
**Fecha:** 2026-08-13  
**Autor:** Desarrollado para procesamiento masivo de Commercial Invoices Baker Hughes

---

## 1. Arquitectura General

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│   PDFs Input    │────▶│  PDFExtractor    │────▶│  LLMParser      │
│  (facturas/*.pdf)│     │  (pypdf + texto) │     │  (Gemini + JSON)│
└─────────────────┘     └──────────────────┘     └─────────────────┘
                                                          │
                           ┌──────────────────────────────┘
                           ▼
                    ┌──────────────────┐
                    │ DataSanitizer    │◄──── Fuerza valores por defecto
                    │ (campos ausentes)│      cuando el LLM devuelve vacío
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │ Description      │◄──── Reglas de normalización
                    │ Normalizer       │      (extensible)
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │ InvoiceValidator │◄──── Tolerancia configurable
                    │ (Ítem + Total)   │      Estados: OK / ADVERTENCIA / ERROR
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │ DatabaseManager  │◄──── SQLite con transacciones
                    │ (upsert + índices)│    Prevención de duplicados
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │ ExportManager    │◄──── CSV / JSON / Excel
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │ HistorialManager │◄──── JSON append-only
                    └──────────────────┘
                           │
                           ▼
                    ┌──────────────────┐
                    │ ReportGenerator  │◄──── Resumen final en consola + log
                    └──────────────────┘
```

---

## 2. Componentes Detallados

### 2.1 Config (`Config`)
Dataclass centralizada con todos los parámetros del sistema. Sin valores hardcodeados en el flujo principal.

| Parámetro | Tipo | Default | Descripción |
|-----------|------|---------|-------------|
| `gemini_api_key` | str | `GEMINI_API_KEY` env | Clave de API de Google Gemini |
| `modelo_llm` | str | `gemini-2.0-flash` | Modelo LLM para extracción |
| `directorio_entrada` | str | `facturas` | Carpeta con PDFs a procesar |
| `directorio_salida` | str | `salida` | Carpeta de exportaciones |
| `db_path` | str | `sistema_facturas.db` | Base de datos SQLite |
| `tolerancia_redondeo` | float | `0.05` | Umbral para marcar error en validación |
| `tolerancia_advertencia` | float | `0.01` | Umbral para marcar advertencia |

### 2.2 PDFExtractor
- **Responsabilidad:** Extraer texto plano de archivos PDF.
- **Librería:** `pypdf.PdfReader`
- **Salida:** String con delimitadores `--- PAGE N ---` para preservar estructura multipágina.
- **Manejo de errores:** Devuelve string vacío en caso de fallo; el pipeline continúa.

### 2.3 LLMParser
- **Responsabilidad:** Interpretar texto desordenado de PDF y estructurarlo en `FacturaInternacional`.
- **Modelo:** Google Gemini 2.0 Flash con `response_schema` (JSON estructurado).
- **Temperatura:** `0.0` (determinístico).
- **Prompt:** Incluye reglas críticas para:
  - Fusionar productos duplicados (aparición con datos + aparición con metadatos)
  - Excluir filas de empaque (`CARTON`, `PIWOODENBOX`)
  - Extraer campos específicos de Baker Hughes (Delivery, Order, Serial No., MEXICAN CUSTOMS IMPORT)
  - Usar valores por defecto explícitos cuando un campo no existe

### 2.4 DataSanitizer
- **Responsabilidad:** Forzar valores por defecto cuando el LLM devuelve strings vacíos o `null`.
- **Campos saneados:**
  - Ítems: `numero_referencia`, `numero_orden`, `numero_serie`, `codigo_exportacion`, `codigo_importacion`, `pais_origen`, `customs_import_mexico`
  - Encabezado: `numero_orden`, `pais_destino_final`, `modo_transporte`, `agente_forwarder`
  - Pie: pesos (`peso_bruto_total_lb`, etc.)

### 2.5 DescriptionNormalizer
- **Responsabilidad:** Limpiar descripciones técnicas eliminando metadatos incrustados.
- **Método:** Regex multi-patrón sobre líneas de texto.
- **Patrones eliminados:**
  - `Delivery: XXXXXX`
  - `Order: XXXXXX`
  - `Export country commodity code: XXXXXX`
  - `Import country commodity code: XXXXXX`
  - `ECCN: XXXXXX`
  - `License Type: XXXXXX`
  - `Country of Origin: XXXXXX`
  - `Serial No. XXXXXX`
  - `MEXICAN CUSTOMS IMPORT: XXXXXX`
  - `**PREPACKAGE ...**`

### 2.6 InvoiceValidator
Validación matemática en **tres niveles**:

#### Nivel 0 — Por ítem
```
total_calculado = cantidad × precio_unitario × (1 + porcentaje/100)
```
- Si `diferencia ≤ 0.01` → `CORRECTO`
- Si `0.01 < diferencia ≤ 0.05` → `ADVERTENCIA`
- Si `diferencia > 0.05` → `ERROR`

#### Nivel 1 — Suma ítems vs Net Value
```
suma_items = Σ(total_calculado de cada ítem)
neto = pie.valor_neto
```

#### Nivel 2 — Net Value + Freight + Impuestos vs Total CPT
```
total_esperado = neto + freight_inland + freight_internacional + impuestos
total_pie = pie.valor_total_factura
```

### 2.7 DatabaseManager
**Motor:** SQLite3  
**Tablas:**

#### `facturas_encabezado`
| Columna | Tipo | Descripción |
|-----------|------|-------------|
| `numero_factura` | TEXT PK | Número de Commercial Invoice |
| `fecha_factura` | TEXT | Fecha de emisión |
| `moneda` | TEXT | Moneda (USD) |
| `emisor_usa` | TEXT | Emisor en EE.UU. |
| `comprador_colombia` | TEXT | Comprador en Colombia |
| `modo_transporte` | TEXT | Air, Sea, etc. |
| `agente_forwarder` | TEXT | DHL EXPRESS, etc. |
| `numero_orden` | TEXT | Orden/PO general |
| `pais_destino_final` | TEXT | Colombia |
| `peso_bruto_lb` | TEXT | Peso bruto en libras |
| `peso_bruto_kg` | TEXT | Peso bruto en kilos |
| `peso_neto_lb` | TEXT | Peso neto en libras |
| `peso_neto_kg` | TEXT | Peso neto en kilos |
| `valor_neto` | REAL | Net Value |
| `freight_inland` | REAL | Inland Freight |
| `freight_internacional` | REAL | International Freight |
| `impuestos` | REAL | Impuestos |
| `valor_total_pie` | REAL | Total CPT |
| `suma_items_calculada` | REAL | Suma calculada por validador |
| `es_matematicamente_valida` | INTEGER | 1 = válida, 0 = inválida |
| `cantidad_errores` | INTEGER | Conteo de errores |
| `cantidad_advertencias` | INTEGER | Conteo de advertencias |
| `fecha_procesamiento` | TEXT | ISO 8601 |

#### `factura_items_cuerpo`
| Columna | Tipo | Descripción |
|-----------|------|-------------|
| `id` | INTEGER PK AI | ID autoincremental |
| `numero_factura` | TEXT FK | Relación con encabezado |
| `codigo_producto` | TEXT | Código de material |
| `descripcion_original` | TEXT | Texto crudo del PDF |
| `descripcion_limpia` | TEXT | Descripción normalizada |
| `cantidad` | REAL | Cantidad |
| `unidad_medida` | TEXT | EA, BAG, KG, etc. |
| `precio_unitario` | REAL | Precio unitario |
| `valor_total_item` | REAL | Extended Price |
| `porcentaje` | REAL | % descuento/adicional |
| `estado_validacion` | TEXT | CORRECTO/ADVERTENCIA/ERROR |
| `total_calculado` | REAL | Cant × PU |
| `diferencia` | REAL | Diff vs factura |
| `numero_referencia` | TEXT | Delivery |
| `numero_orden` | TEXT | Order/PO |
| `numero_serie` | TEXT | Serial No. |
| `codigo_exportacion` | TEXT | Export commodity code |
| `codigo_importacion` | TEXT | Import commodity code |
| `pais_origen` | TEXT | Country of Origin |
| `customs_import_mexico` | TEXT | MEXICAN CUSTOMS IMPORT |

**Índices:**
- `idx_items_factura` sobre `factura_items_cuerpo(numero_factura)`

**Prevención de duplicados:**
- `INSERT OR REPLACE` en encabezado (upsert por `numero_factura`)
- Borrado previo de ítems antes de re-insertar (idempotencia)

### 2.8 ExportManager
Genera tres formatos de salida en `salida/`:

| Archivo | Formato | Contenido |
|---------|---------|-----------|
| `cuerpo_items_facturas.csv` | CSV UTF-8-SIG | Todos los ítems de todas las facturas |
| `encabezados_facturas.csv` | CSV UTF-8-SIG | Resumen de encabezados |
| `facturas_completas.json` | JSON | Dump completo del modelo Pydantic |
| `Reporte_Facturas_Comex.xlsx` | Excel (.xlsx) | Hoja 1: Ítems con colores por estado. Hoja 2: Resumen + gráfico de barras CPT |

### 2.9 HistorialManager
- Archivo: `historial_facturas_comex.json`
- Formato: Array JSON append-only
- Función: Deduplicación por `numero_factura` + ordenamiento por fecha

### 2.10 ReportGenerator
Resumen final en consola y log con métricas:
- Total PDFs procesados
- Correctas / Advertencias / Errores / Fallidas
- Total productos extraídos
- Productos inconsistentes
- Facturas con total CPT incorrecto

---

## 3. Flujo de Ejecución (Pipeline)

```python
1. Cargar configuración (.env)
2. Inicializar logging
3. Buscar PDFs en directorio_entrada + raíz
4. Para cada PDF:
   4.1 Extraer texto con pypdf
   4.2 Parsear con Gemini LLM → FacturaInternacional
   4.3 Detectar duplicados en BD
   4.4 Sanitizar campos ausentes
   4.5 Normalizar descripciones
   4.6 Validar matemáticamente (3 niveles)
   4.7 Guardar en SQLite
   4.8 Actualizar historial JSON
5. Exportar CSV + JSON + Excel
6. Generar reporte final
```

**Característica crítica:** El pipeline **nunca se detiene** por un error en una factura. Cada excepción se captura, se loguea y se continúa con la siguiente.

---

## 4. Modelos de Datos (Pydantic)

### Jerarquía
```
FacturaInternacional
├── EncabezadoFactura
├── List[ItemFactura]
│   └── ResultadoValidacionItem
└── PieFactura

FacturaProcesada
├── FacturaInternacional
├── es_valida: bool
├── errores: List[str]
├── advertencias: List[str]
├── suma_items_calculada: float
└── fecha_procesamiento: str
```

### Estados de Validación
```python
class EstadoValidacion(str, Enum):
    CORRECTO = "CORRECTO"
    ADVERTENCIA = "ADVERTENCIA"
    ERROR = "ERROR"
```

---

## 5. Dependencias

```
Python >= 3.10
├── pypdf >= 4.0
├── google-genai >= 1.0
├── pydantic >= 2.0
├── pandas >= 2.0
├── openpyxl >= 3.1
└── python-dotenv >= 1.0
```

Instalación:
```bash
pip install pypdf google-genai pydantic pandas openpyxl python-dotenv
```

---

## 6. Configuración

Crear archivo `.env` en la raíz del proyecto:

```env
GEMINI_API_KEY=tu_clave_de_api_aqui
```

**Nota de seguridad:** El archivo `.env` debe estar en `.gitignore`. Nunca commitear la API key.

---

## 7. Logging

### Niveles
- **Consola:** `INFO` y superior
- **Archivo (`logs/facturador.log`):** `DEBUG` y superior
- **Rotación:** 5 MB por archivo, máximo 3 backups

### Formato
```
2026-08-13 12:38:17 | INFO     | facturador | Factura parseada: 4634461
2026-08-13 12:38:30 | ERROR    | facturador | NIVEL 1 - Net Value: Suma items ($462.77) != Net Value ($353.26)
```

---

## 8. Extensibilidad

### Agregar nuevas reglas de descripción
Modificar `DescriptionNormalizer.PATRONES_METADATOS` o implementar `aplicar_reglas_custom()`.

### Cambiar modelo LLM
Modificar `Config.modelo_llm` (ej. `gemini-2.5-pro` para mayor precisión).

### Agregar nuevo campo
1. Agregar al modelo Pydantic correspondiente
2. Actualizar `DataSanitizer.DEFAULTS_*` si aplica
3. Actualizar `DatabaseManager.COLUMNAS_*`
4. Actualizar prompts del LLM

### Cambiar formato de salida
`ExportManager` es independiente del resto del pipeline. Se puede agregar PDF, Parquet, etc. sin tocar extracción ni validación.

---

## 9. Manejo de Errores

| Escenario | Comportamiento |
|-----------|----------------|
| PDF corrupto | Log error, continúa con siguiente |
| Texto vacío del PDF | Marca como fallida, continúa |
| LLM timeout | Excepción capturada, marca como fallida |
| Validación fallida | Guarda en BD con `es_valida=0`, reporta en resumen |
| BD bloqueada | Reintenta en siguiente ejecución |
| Campo ausente | `DataSanitizer` fuerza valor por defecto explícito |

---

## 10. Rendimiento

| Métrica | Valor típico |
|---------|--------------|
| Tiempo por factura (LLM) | 5–15 segundos |
| Facturas procesadas por lote | Ilimitado (pipeline secuencial) |
| Tamaño máximo de PDF | Limitado por API de Gemini (~1MB texto) |
| Memoria | ~100 MB base + overhead de pandas/openpyxl |

---

## 11. Consideraciones de Seguridad

- API Key almacenada solo en `.env` (nunca en código)
- Base de datos SQLite local (sin exposición de red)
- No se envían datos a servicios externos excepto Google Gemini
- Logs rotativos para evitar consumo de disco ilimitado
