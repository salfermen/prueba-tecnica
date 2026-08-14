# Manual de Usuario — Sistema de Automatización de Facturas Comex
## Agente Comex Baker Hughes v3.0

---

## 1. Instalación

### 1.1 Requisitos previos
- Python 3.10 o superior instalado
- Conexión a Internet (para API de Gemini)
- Archivo `.env` con tu clave de API

### 1.2 Crear entorno virtual (recomendado)

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 1.3 Instalar dependencias

```bash
pip install pypdf google-genai pydantic pandas openpyxl python-dotenv
```

### 1.4 Configurar API Key

Crea un archivo llamado `.env` en la misma carpeta donde está el script:

```env
GEMINI_API_KEY=tu_clave_de_api_aqui
```

> **Importante:** Obtén tu clave en https://aistudio.google.com/app/apikey. Nunca compartas este archivo.

---

## 2. Estructura de carpetas recomendada

```
Facturas_Comex/
├── .env                          ← Tu clave de API
├── agente_comex_bakerhughes_v3.py ← Script principal
├── facturas/                     ← Coloca aquí los PDFs
│   ├── 89_CI_4634461.pdf
│   ├── baker_CI_4636383.pdf
│   └── ...
├── salida/                       ← Se crea automáticamente
│   ├── cuerpo_items_facturas.csv
│   ├── encabezados_facturas.csv
│   ├── facturas_completas.json
│   └── Reporte_Facturas_Comex.xlsx
├── logs/                         ← Se crea automáticamente
│   └── facturador.log
├── sistema_facturas.db           ← Base de datos SQLite
└── historial_facturas_comex.json ← Historial acumulativo
```

---

## 3. Ejecución

### 3.1 Primera vez (obligatorio borrar BD vieja si existe)

```bash
# Windows
del sistema_facturas.db
python agente_comex_bakerhughes_v3.py

# macOS / Linux
rm -f sistema_facturas.db
python agente_comex_bakerhughes_v3.py
```

### 3.2 Ejecuciones posteriores

```bash
python agente_comex_bakerhughes_v3.py
```

> El sistema detecta facturas ya procesadas y las actualiza sin duplicar.

### 3.3 ¿Dónde busca PDFs?

El script busca en dos lugares:
1. La carpeta `facturas/` (configurable)
2. La carpeta raíz donde está el script

---

## 4. Interpretación de resultados

### 4.1 Salida en consola

```
2026-08-13 12:38:17 | INFO     | facturador | Iniciando procesamiento de 5 factura(s)...

Procesando: facturas\89_CI 4634461.pdf
2026-08-13 12:38:30 | INFO     | facturador | Factura parseada: 4634461
2026-08-13 12:38:30 | INFO     | facturador | 4634461 | Item #1: OK
2026-08-13 12:38:30 | INFO     | facturador | 4634461 | Item #2: OK
2026-08-13 12:38:30 | INFO     | facturador | 4634461: Net Value OK ($2514.63)
2026-08-13 12:38:30 | INFO     | facturador | 4634461: Total CPT OK ($2565.11)
```

**Significado de los mensajes:**

| Mensaje | Significado |
|---------|-------------|
| `Item #N: OK` | El cálculo Cantidad × Precio coincide con el total de la factura |
| `Item #N: Diferencia por redondeo` | Hay una pequeña diferencia (≤ $0.05), probablemente por redondeo |
| `Item #N: INCORRECTO` | La matemática no cuadra. Revisar manualmente |
| `Net Value OK` | La suma de todos los ítems coincide con el Net Value de la factura |
| `Total CPT OK` | Net Value + Freight + Impuestos = Total final |

### 4.2 Reporte final

Al terminar, verás algo como:

```
======================================================================
RESUMEN FINAL DE PROCESAMIENTO
======================================================================
  Total PDFs encontrados:        5
  Procesadas correctamente:      3
  Con advertencias:              0
  Con errores:                   2
  Fallidas (excepcion):          0
  ----------------------------------------
  Total productos extraidos:     23
  Productos inconsistentes:      4
  Facturas con total CPT mal:    2
======================================================================
```

**Qué hacer según el resultado:**

| Estado | Acción recomendada |
|--------|-------------------|
| **Correctamente** | Todo bien. Los datos están en BD, CSV y Excel. |
| **Con advertencias** | Revisar los logs. Probablemente sea solo redondeo. |
| **Con errores** | Revisar la factura manualmente. Puede haber descuentos no detectados o precios mal extraídos. |
| **Fallidas** | Revisar si el PDF está corrupto o si hay problema de conexión con Gemini. |

---

## 5. Archivos generados

### 5.1 Base de datos (`sistema_facturas.db`)
Puedes consultarla con cualquier herramienta SQLite (DB Browser, DBeaver, etc.).

**Consultas útiles:**

```sql
-- Ver todas las facturas
SELECT numero_factura, fecha_factura, es_matematicamente_valida, valor_total_pie
FROM facturas_encabezado;

-- Ver ítems de una factura específica
SELECT codigo_producto, descripcion_limpia, cantidad, precio_unitario, estado_validacion
FROM factura_items_cuerpo
WHERE numero_factura = '4634461';

-- Ver solo ítems con problemas
SELECT * FROM factura_items_cuerpo
WHERE estado_validacion != 'CORRECTO';
```

### 5.2 Excel (`salida/Reporte_Facturas_Comex.xlsx`)

**Hoja 1 — Cuerpo de Items:**
- Cada fila es un producto
- Columna `Estado Validacion` con colores:
  - 🟢 Verde = CORRECTO
  - 🟡 Amarillo = ADVERTENCIA
  - 🔴 Rojo = ERROR

**Hoja 2 — Resumen Facturas:**
- Una fila por factura
- Gráfico de barras con totales CPT
- Columna `Valida?` con color verde/rojo

### 5.3 CSVs
- `cuerpo_items_facturas.csv` — Todos los productos (abrir en Excel)
- `encabezados_facturas.csv` — Resumen de facturas

### 5.4 JSON
- `facturas_completas.json` — Datos completos en formato estructurado (para integraciones)

---

## 6. Consultar una factura desde Python

Abre Python en la misma carpeta del script:

```python
from agente_comex_bakerhughes_v3 import buscar_factura

# Buscar por número de factura
buscar_factura("4634461")
```

**Salida esperada:**

```
======================================================================
FACTURA: 4634461
======================================================================
  Fecha:      26-Feb-2026
  Moneda:     USD
  Emisor:     Baker Hughes Oilfield Operations LLC
  Comprador:  BAKER HUGHES DE COLOMBIA
  Transporte: Air
  Forwarder:  DHL EXPRESS
  N Orden:    NO DISPONIBLE
  Destino:    Colombia
  Peso Bruto: 3.00 LB / 1.36 KG
  Peso Neto:  2.00 LB / 0.91 KG
  Net Value:  $2514.63
  Freight:    $50.48
  Impuestos:  $0.0
  Total CPT:  $2565.11
  Valida?:    SI
  Errores:    0 | Advertencias: 0

  ITEMS (5):
  ------------------------------------------------------------
  [OK] F243858000 | WINDOW,TEFLON (1673MA)...
      3.0 EA | PU:$369.59 | Total:$1108.77
      Calc:$1108.77 | Dif:$0.0
      Ref:824273721 | Ord:NO TIENE ORDEN | Ser:NO TIENE SERIAL
      Exp:3926909989 | Imp:3926909090 | Orig:USA | MX:NO APLICA
```

---

## 7. Solución de problemas (Troubleshooting)

### 7.1 Error: `GEMINI_API_KEY no configurada`
**Causa:** Falta el archivo `.env` o la clave está vacía.  
**Solución:** Crear `.env` con `GEMINI_API_KEY=tu_clave`.

### 7.2 Error: `sqlite3.OperationalError: table has X columns but Y values`
**Causa:** El esquema de la base de datos cambió entre versiones.  
**Solución:** Borrar `sistema_facturas.db` y volver a ejecutar.

### 7.3 Error: `google-genai no esta instalado`
**Causa:** Falta instalar dependencias.  
**Solución:** `pip install google-genai`

### 7.4 La factura se procesa pero los precios no cuadran
**Causas posibles:**
- El PDF tiene descuentos por volumen no explícitos
- El LLM extrajo un precio unitario incorrecto
- Hay impuestos incluidos en el precio unitario

**Solución:** Revisar el log, comparar con el PDF original. Si persiste, enviar el PDF para ajustar el prompt.

### 7.5 Timeout o cuelgue con una factura
**Causa:** PDF muy grande o conexión lenta a Gemini.  
**Solución:** El script tiene manejo de excepciones. Si se cuelga, presiona `Ctrl+C` y vuelve a ejecutar. La factura problemática se marcará como fallida y el pipeline continuará con las demás.

### 7.6 Los campos "NO TIENE..." no aparecen
**Causa:** Versión anterior del script sin `DataSanitizer`.  
**Solución:** Usar la versión v3 o superior.

### 7.7 No se genera el Excel
**Causa:** `openpyxl` no está instalado.  
**Solución:** `pip install openpyxl`

---

## 8. Flujo de trabajo recomendado

### Diario
```
1. Colocar nuevos PDFs en carpeta facturas/
2. Ejecutar: python agente_comex_bakerhughes_v3.py
3. Revisar resumen en consola
4. Abrir salida/Reporte_Facturas_Comex.xlsx
5. Revisar facturas marcadas en rojo (errores)
6. Corregir manualmente las que tengan problemas
```

### Semanal
```
1. Revisar logs/facturador.log para patrones de error
2. Verificar que no haya duplicados en la BD
3. Hacer backup de sistema_facturas.db
4. Exportar CSVs para otros sistemas (ERP, contabilidad)
```

---

## 9. Glosario de términos

| Término | Significado |
|---------|-------------|
| **Commercial Invoice** | Factura comercial internacional |
| **CPT** | Carriage Paid To (Incoterm) — total con freight incluido |
| **Net Value** | Suma de todos los productos sin freight ni impuestos |
| **Extended Price** | Cantidad × Precio Unitario de un ítem |
| **Delivery** | Número de referencia de envío (Baker Hughes) |
| **ECCN** | Export Control Classification Number |
| **NLR** | No License Required (tipo de licencia de exportación) |
| **Country of Origin** | País donde se fabricó el producto |
| **Commodity Code** | Código arancelario (exportación / importación) |
| **MEXICAN CUSTOMS IMPORT** | Número de importación aduanera en México |

---

## 10. Contacto y soporte

Para reportar bugs o solicitar mejoras:
1. Revisar primero `logs/facturador.log`
2. Guardar el PDF problemático
3. Copiar el mensaje de error exacto
4. Enviar con contexto (qué factura, qué esperabas, qué obtuviste)
