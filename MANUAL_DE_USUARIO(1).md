# Manual de Usuario — Agente Comex

Este manual explica cómo instalar, configurar y usar el programa
`agente_comex.py` sin necesidad de conocimientos de programación.

## 1. ¿Para qué sirve?

Automatiza la lectura de las Commercial Invoices (facturas de comercio
exterior) en PDF, extrae sus datos con inteligencia artificial, revisa que
los números cuadren matemáticamente y entrega los resultados en Excel, CSV
y una base de datos consultable.

## 2. Instalación

1. Instala Python 3.9 o superior.
2. Instala las librerías necesarias abriendo una terminal en la carpeta del
   programa y ejecutando:
   ```bash
   pip install pandas pydantic pypdf python-dotenv google-genai openpyxl
   ```
3. Crea un archivo llamado `.env` en la misma carpeta que `agente_comex.py`
   con el siguiente contenido:
   ```
   GEMINI_API_KEY=tu_clave_de_gemini_aqui
   ```
   Reemplaza el valor por tu clave real de Google Gemini. Sin esta clave el
   programa no podrá leer las facturas (aunque sí podrás usar `--buscar`).

## 3. Carpetas del programa

Todas las rutas se resuelven respecto a la carpeta donde está el programa
(o el .exe si se usa una versión compilada), sin importar desde dónde lo
ejecutes:

| Carpeta / archivo | Contenido |
|---|---|
| `facturas/` | Aquí van los PDF que quieres procesar (carpeta de entrada por defecto) |
| `salida/` | Aquí se generan automáticamente los resultados |
| `sistema_facturas.db` | Base de datos con el historial de facturas procesadas |
| `historial_facturas_comex.json` | Historial resumido en formato JSON |
| `logs/facturador.log` | Registro detallado de cada ejecución |

## 4. Uso básico

### Procesar todas las facturas nuevas

Copia los PDF a la carpeta `facturas/` y ejecuta:

```bash
python agente_comex.py --todas
```

Si no indicas ningún parámetro de selección, este es el comportamiento por
defecto.

### Procesar solo algunas facturas

- Por cantidad (las primeras N en orden alfabético del nombre del archivo):
  ```bash
  python agente_comex.py --cantidad 5
  ```
- Por número de factura (busca ese texto en el nombre del archivo PDF):
  ```bash
  python agente_comex.py --factura 4638446
  ```

### Reprocesar una factura ya existente

Por defecto, si una factura ya fue procesada antes, el programa la omite
para no duplicar trabajo. Si necesitas actualizarla (por ejemplo, porque
corregiste el PDF), agrega `--reprocesar`:

```bash
python agente_comex.py --factura 4638446 --reprocesar
```

### Consultar una factura ya procesada

No requiere conexión a internet ni clave de API, porque solo lee de la base
de datos local:

```bash
python agente_comex.py --buscar 4638446
```

Esto imprime en pantalla el encabezado, los totales y el detalle de cada
producto, incluyendo si la factura quedó validada correctamente.

### Usar carpetas distintas a las de por defecto

```bash
python agente_comex.py --entrada "C:\ruta\a\mis_pdf" --salida "C:\ruta\a\resultados"
```

### Ver más detalle en pantalla mientras procesa

```bash
python agente_comex.py --todas --log-level DEBUG
```

## 5. Cómo leer los resultados

Por cada factura procesada se crea una carpeta dentro de `salida/` con el
número de la factura, que contiene:

- **`encabezado.csv`**: datos generales (emisor, comprador, transporte,
  totales).
- **`cuerpo.csv`**: el detalle de cada producto/ítem de la factura.
- **`reporte.xlsx`**: un Excel con dos hojas —"Información General" (datos,
  totales y comprobaciones matemáticas resaltadas en verde/amarillo/rojo) y
  "Productos" (detalle línea por línea con el mismo código de colores).

Además, al final de cada corrida se generan archivos **consolidados** en la
raíz de `salida/` con todas las facturas procesadas en esa ejecución:
`encabezados_facturas.csv`, `cuerpo_items_facturas.csv`,
`facturas_completas.json` y `Reporte_Facturas_Comex.xlsx` (este último
incluye un gráfico de barras con el Total CPT de cada factura).

### Código de colores en los Excel

| Color | Significado |
|---|---|
| 🟩 Verde | Todo cuadra matemáticamente |
| 🟨 Amarillo | Advertencia: diferencia pequeña, probablemente por redondeo |
| 🟥 Rojo | Error: la diferencia supera la tolerancia permitida |

## 6. Resumen en consola

Al terminar de procesar, el programa muestra un resumen con:

- Total de PDF encontrados.
- Cuántas facturas quedaron correctas, con advertencias, con errores,
  fallidas (por excepción/error inesperado) u omitidas (ya existían en la
  base de datos).
- Total de productos extraídos y cuántos quedaron inconsistentes.
- El detalle de cada factura con problemas.

## 7. Preguntas frecuentes

**¿Qué pasa si un PDF no tiene número de factura visible?**
El programa le asigna un identificador alternativo basado en el nombre del
archivo (`SIN_NUMERO_<nombre_del_archivo>`) y continúa procesándolo
normalmente, dejando una advertencia registrada.

**¿Qué pasa si proceso la misma factura dos veces sin `--reprocesar`?**
El programa detecta que ya existe en la base de datos y la omite,
mostrándola en la sección "FACTURAS OMITIDAS" del resumen final.

**¿Dónde reviso errores si algo falla?**
En el archivo `logs/facturador.log`, que guarda el detalle completo de cada
ejecución (incluye más información que la que se muestra en pantalla).

**¿Necesito internet para consultar una factura ya procesada?**
No. El comando `--buscar` solo lee la base de datos local.

Para el detalle técnico de cómo funciona internamente el programa, ver
**DOCUMENTACION_TECNICA.md**.
