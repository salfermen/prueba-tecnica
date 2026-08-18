# Documentación Técnica — Agente Comex

Documento de referencia técnica de `agente_comex.py`: arquitectura, modelos
de datos, flujo de ejecución, reglas de negocio y esquema de la base de
datos.

## 1. Stack y dependencias

| Librería | Uso |
|---|---|
| `pypdf` | Extracción de texto de los PDF de las facturas |
| `google-genai` | Cliente del LLM (Gemini) para parsear el texto de la factura a JSON estructurado |
| `pydantic` | Definición y validación de los modelos de datos (schema de salida del LLM) |
| `pandas` | Generación de los CSV de salida |
| `openpyxl` | Generación de los reportes Excel con formato condicional y gráficos |
| `sqlite3` (stdlib) | Persistencia de facturas procesadas y prevención de duplicados |
| `python-dotenv` | Carga de `GEMINI_API_KEY` desde el archivo `.env` |
| `argparse` (stdlib) | Interfaz de línea de comandos |
| `logging` (stdlib) | Logging a consola y a archivo rotativo |

`GENAI_DISPONIBLE` y `EXCEL_DISPONIBLE` se detectan en tiempo de import: si
`google-genai` u `openpyxl` no están instalados, el programa sigue
funcionando mejor esfuerzo (por ejemplo, se puede seguir usando `--buscar`
sin `google-genai`, y el resto de exportaciones funcionan sin `openpyxl`,
solo se omiten los `.xlsx`).

## 2. Resolución de rutas (`base_dir`)

Todas las rutas relativas (`.env`, `facturas/`, `salida/`, base de datos,
logs) se resuelven respecto a `BASE_DIR`, calculado así:

- Si el programa corre como ejecutable compilado con PyInstaller
  (`sys.frozen == True`): la carpeta donde está el `.exe`.
- Si corre como script de Python: la carpeta donde está `agente_comex.py`.

Esto garantiza que el comportamiento sea el mismo sin importar desde qué
directorio de trabajo (`cwd`) se invoque el programa.

`sanitizar_nombre()` limpia nombres de factura para usarlos como nombre de
carpeta/archivo en Windows, removiendo `<>:"/\|?*` y caracteres de control.

## 3. Configuración (`Config`)

`dataclass` con todos los parámetros del sistema:

- `gemini_api_key`: leída de la variable de entorno `GEMINI_API_KEY`.
- `modelo_llm`: `"gemini-3.5-flash"` por defecto.
- `directorio_entrada` / `directorio_salida`: `facturas/` y `salida/`.
- `db_path`: `sistema_facturas.db`.
- `json_historial`: `historial_facturas_comex.json`.
- `log_path`: `logs/facturador.log`.
- `log_nivel_consola` / `log_nivel_archivo`: `INFO` / `DEBUG`.
- `tolerancia_redondeo` (0.05) y `tolerancia_advertencia` (0.01): umbrales en
  dólares usados por el validador matemático.
- `timeout_llm`: 120 segundos (definido, no aplicado explícitamente en la
  llamada actual al cliente).

En `__post_init__` todas las rutas se convierten a absolutas vía
`resolver_ruta()` y se crean las carpetas de salida y de logs si no existen.
`validar_api_key()` lanza `ValueError` si la clave no está configurada o
quedó con el valor placeholder.

## 4. Logging (`setup_logging`)

Logger `"facturador"` con dos handlers:

- `StreamHandler` a stdout, nivel configurable por `--log-level`
  (por defecto `INFO`).
- `RotatingFileHandler` sobre `logs/facturador.log`, nivel `DEBUG`,
  rotación a los 5 MB con 3 backups, codificación UTF-8.

Formato: `%(asctime)s | %(levelname)-8s | %(name)s | %(message)s`.

## 5. Modelos de datos (Pydantic)

```
FacturaInternacional
├── encabezado: EncabezadoFactura
├── items: List[ItemFactura]
└── pie: PieFactura
```

### `EncabezadoFactura`
`numero_factura`, `fecha_factura`, `moneda` (default `USD`), `emisor_usa`,
`comprador_colombia`, `modo_transporte`, `agente_forwarder`,
`numero_orden` (default `NO DISPONIBLE`), `pais_destino_final`
(default `Colombia`).

### `ItemFactura`
`codigo_producto`, `descripcion_original`, `descripcion_tecnica_limpia`,
`cantidad`, `unidad_medida` (default `EA`), `precio_unitario`,
`valor_total_item`, `porcentaje` (opcional), `numero_referencia`
(default `NO TIENE DELIVERY`), `numero_orden` (default `NO TIENE ORDEN`),
`numero_serie` (default `NO TIENE SERIAL`), `codigo_exportacion`,
`codigo_importacion`, `pais_origen` (default `N/A`),
`customs_import_mexico` (default `NO APLICA`), y `validacion`
(`ResultadoValidacionItem`, se rellena en la etapa de validación).

### `PieFactura`
Pesos brutos/netos en LB y KG (como texto), `valor_neto`, `freight_inland`,
`freight_internacional`, `impuestos`, `valor_total_factura` (Total CPT).

### `FacturaProcesada`
Envuelve una `FacturaInternacional` con el resultado de la validación:
`es_valida`, `errores`, `advertencias`, `suma_items_calculada`,
`fecha_procesamiento`.

### `EstadoValidacion` (Enum)
`CORRECTO`, `ADVERTENCIA`, `ERROR` — estado individual de cada ítem tras la
validación.

## 6. Pipeline de procesamiento (`InvoicePipeline`)

Orquestador principal, compuesto por:

```
PDFExtractor → LLMParser → DataSanitizer → DescriptionNormalizer
→ InvoiceValidator → DatabaseManager + HistorialManager + ExportManager
→ ReportGenerator
```

### 6.1 Selección de archivos (`_seleccionar`)
Lista los PDF de `directorio_entrada` ordenados alfabéticamente
(`_listar_pdfs`) y aplica el filtro pedido:
- `--factura NUMERO`: coincidencia de substring contra el nombre del
  archivo (`Path(a).stem`).
- `--cantidad N`: los primeros N archivos.
- Sin filtro: todos los archivos.

### 6.2 `PDFExtractor.extraer_texto`
Usa `pypdf.PdfReader`, concatena el texto de todas las páginas con
separadores `--- PAGE N ---`. Si falla la lectura, retorna cadena vacía
(la factura se marca como fallida).

### 6.3 `LLMParser.parsear`
Construye el prompt `PROMPT_TEMPLATE` (formato estructurado con etiquetas
`<ROLE>`, `<TASK>`, `<INPUT>`, `<RULES>`, `<VALIDATIONS>`, `<OUTPUT>`) y
llama a `genai.Client.models.generate_content` con:
- `response_mime_type="application/json"`
- `response_schema=FacturaInternacional` (salida forzada al schema)
- `temperature=0.0` (determinismo)
- `system_instruction` reforzando las reglas críticas.

Reglas clave codificadas en el prompt:
1. Cada producto aparece dos veces en el PDF (una con datos comerciales,
   otra con metadatos); se debe incluir una sola vez usando la descripción
   de la primera aparición, y extraer de la segunda: Delivery, Order,
   Serial No., commodity codes, ECCN, License Type, Country of Origin.
2. Se excluyen filas de empaque (`CARTON`, `PIWOODENBOX`), pesos parciales y
   subtotales/freight.
3. La descripción limpia no debe incluir metadatos.
4. Todo campo ausente debe usar el valor por defecto exacto del schema (no
   strings vacíos, no null).
5. `valor_total_item` = cantidad × precio unitario (ajustado por
   `porcentaje` si aplica).
6. Reglas de cálculo del pie de página (Net Value, freight, Total CPT,
   pesos en LB/KG).

La respuesta se valida con `FacturaInternacional.model_validate_json`.
Errores de `pydantic.ValidationError` o de la llamada al LLM se registran y
se propagan (la factura queda en `fallidas`).

### 6.4 `DataSanitizer.sanitizar`
Refuerza en código los valores por defecto del prompt (por si el LLM
devuelve `null` o string vacío), sobre tres diccionarios de defaults:
`DEFAULTS_ITEM`, `DEFAULTS_ENCABEZADO`, `DEFAULTS_PIE`.

### 6.5 `DescriptionNormalizer.normalizar`
Limpia `descripcion_tecnica_limpia` línea por línea, eliminando líneas que
matcheen alguno de los patrones regex de `PATRONES_METADATOS` (Delivery,
Order, commodity codes, ECCN, License Type, Country of Origin, Serial No.,
MEXICAN CUSTOMS IMPORT, bloques `**PREPACKAGE...**`), y colapsa espacios.
Guarda el valor previo en `descripcion_original`.

`aplicar_reglas_custom` existe como punto de extensión para reglas de
limpieza adicionales por cliente/negocio (no implementado, retorna la
factura sin cambios).

### 6.6 `InvoiceValidator.validar` — validación en dos niveles

**Nivel 0 (por ítem):**
```
total_calculado = cantidad * precio_unitario * (1 + porcentaje/100)  # si aplica
diferencia = |total_calculado - valor_total_item|
```
- `diferencia <= tolerancia_advertencia (0.01)` → `CORRECTO`
- `diferencia <= tolerancia_redondeo (0.05)` → `ADVERTENCIA`
- en otro caso → `ERROR`

**Nivel 1 (suma de ítems vs. Net Value):**
```
|suma(total_calculado de items) - pie.valor_neto| <= tolerancia_redondeo
```

**Nivel 2 (Net Value + Freight + Impuestos vs. Total CPT):**
```
total_esperado = valor_neto + freight_inland + freight_internacional + impuestos
|total_esperado - valor_total_factura| <= tolerancia_redondeo
```

`es_valida` es `True` solo si no hay errores acumulados (las advertencias no
invalidan la factura).

## 7. Persistencia

### 7.1 SQLite (`DatabaseManager`)
Dos tablas:

- **`facturas_encabezado`** (PK `numero_factura`): datos generales +
  resultado de validación (`es_matematicamente_valida`, `cantidad_errores`,
  `cantidad_advertencias`, `fecha_procesamiento`).
- **`factura_items_cuerpo`** (PK autoincremental, FK `numero_factura`):
  detalle de cada ítem, incluyendo el resultado de su validación individual.

`guardar()` hace `INSERT OR REPLACE` en el encabezado y borra + reinserta
todos los ítems de esa factura (permite reprocesamiento limpio).
`ya_procesada()` verifica existencia previa por `numero_factura`.
`buscar()` devuelve encabezado + ítems como diccionarios, usado por
`--buscar` desde línea de comandos.

### 7.2 Historial JSON (`HistorialManager`)
Mantiene `historial_facturas_comex.json` como lista de registros resumidos
(no el detalle de ítems), reemplazando el registro existente para el mismo
`numero_factura` y reordenando por `fecha_factura`.

## 8. Exportaciones (`ExportManager`)

### Por factura individual (`exportar_factura`)
Carpeta `salida/<numero_factura_sanitizado>/` con:
- `encabezado.csv`
- `cuerpo.csv`
- `reporte.xlsx` (`_excel_individual`): hoja "Información General" (datos,
  totales, comprobaciones matemáticas de los dos niveles, advertencias y
  errores) + hoja "Productos" (detalle de ítems con relleno de color según
  `EstadoValidacion`).

Un error al generar la salida de una factura se registra pero no detiene el
procesamiento de las demás (`try/except` aislado por factura).

### Consolidados de la corrida (`exportar`)
Sobre el conjunto de facturas procesadas en la ejecución actual:
- `encabezados_facturas.csv` / `cuerpo_items_facturas.csv`
- `facturas_completas.json` (dump completo de cada `FacturaInternacional`)
- `Reporte_Facturas_Comex.xlsx` (`_excel_consolidado`): hoja "Cuerpo de
  Items" (todas las líneas de todas las facturas) + hoja "Resumen Facturas"
  (una fila por factura con totales y estado) con un `BarChart` de Total
  CPT por factura.

Colores usados de forma consistente en ambos Excel:
`E2EFDA` (verde/correcto), `FFF2CC` (amarillo/advertencia),
`FCE4D6` (rojo/error), encabezados en `1F4E79` con texto blanco.

## 9. Reporte final (`ReportGenerator.generar`)

Calcula y loggea (nivel `INFO`) un resumen de la corrida: total de PDF
encontrados, facturas correctas / con advertencias / con errores /
fallidas (por excepción) / omitidas (ya en BD sin `--reprocesar`), total de
productos extraídos, productos inconsistentes y facturas con Total CPT
incorrecto. Además detalla el listado de facturas omitidas, fallidas, con
errores y con advertencias.

## 10. CLI (`construir_parser` / `main`)

Grupo mutuamente exclusivo de selección: `--todas` (default implícito),
`--factura NUMERO`, `--cantidad N`.

Otros parámetros: `--entrada`, `--salida`, `--reprocesar`, `--buscar
NUMERO` (modo consulta, no requiere API key ni procesa PDFs — retorna
antes de instanciar `InvoicePipeline`), `--log-level`.

`main()`:
1. Parsea argumentos.
2. Carga `.env` desde `BASE_DIR`.
3. Construye `Config` con los overrides de CLI.
4. Si `--buscar`, ejecuta `buscar_factura()` y termina (código 0).
5. Si no, instancia `InvoicePipeline` (puede fallar con `ValueError` si no
   hay API key configurada → código de salida 2).
6. Ejecuta `pipeline.run(factura, cantidad, reprocesar)`.

### `buscar_factura(numero_factura, cfg=None)`
Función pública independiente del pipeline: crea su propio `Config`,
`logger` y `DatabaseManager`, consulta la factura y la imprime en formato
legible en consola (encabezado, totales, y cada ítem con su estado de
validación).

## 11. Manejo de errores por factura

Dentro de `InvoicePipeline.run`, cada PDF se procesa en un bloque
`try/except` independiente: si un PDF falla (sin texto extraído, error del
LLM, error de validación de schema, etc.), se registra en `fallidas` y el
loop continúa con el siguiente archivo sin detener toda la corrida.

Si una factura no trae número detectable, se le asigna un identificador
`SIN_NUMERO_<nombre_de_archivo_sanitizado>` y se deja constancia en las
advertencias de esa factura.

## 12. Puntos de extensión sugeridos

- `DescriptionNormalizer.aplicar_reglas_custom`: hook ya definido para
  reglas de limpieza de descripción específicas por cliente.
- `Config.modelo_llm`: cambiar de modelo Gemini sin tocar el resto del
  código.
- `Config.tolerancia_redondeo` / `tolerancia_advertencia`: ajustar
  sensibilidad de la validación matemática sin tocar lógica.
- `EXCEL_DISPONIBLE` / `GENAI_DISPONIBLE`: el sistema degrada
  funcionalidad de forma controlada si faltan dependencias opcionales.
