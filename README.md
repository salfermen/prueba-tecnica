# Agente Comex — Automatizacion de Facturas Internacionales

Sistema de automatizacion para extraccion, validacion y almacenamiento estructurado de Commercial Invoices de Baker Hughes (USA y Mexico). Procesa facturas PDF masivamente usando inteligencia artificial (Google Gemini), valida matematicamente cada producto y genera reportes en CSV, JSON, Excel y SQLite.

---

## Caracteristicas

- **Extraccion inteligente** de encabezado, productos y pie de factura desde PDFs
- **Fusion automatica** de productos duplicados (descripcion + metadatos)
- **Validacion matematica** en 3 niveles: por item, Net Value y Total CPT
- **Normalizacion de descripciones** eliminando metadatos (Delivery, ECCN, commodity codes)
- **Prevencion de duplicados** con base de datos SQLite
- **Exportaciones multiples**: CSV, JSON, Excel con graficos
- **Logging profesional** con archivo rotativo
- **Pipeline robusto**: nunca se detiene por errores en una factura

---

## Requisitos

- Python 3.10+
- Clave de API de [Google AI Studio](https://aistudio.google.com/app/apikey)

## Instalacion

```bash
# 1. Clonar el repositorio
git clone https://github.com/tu-usuario/agente-comex.git
cd agente-comex

# 2. Crear entorno virtual
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Configurar API Key
cp .env.example .env
# Editar .env y colocar tu GEMINI_API_KEY
```

## Uso

```bash
# Colocar PDFs en la carpeta facturas/
mkdir facturas
# Copiar tus archivos PDF aqui

# Ejecutar
python agente_comex_bakerhughes_v3.py
```

Los resultados se guardan en:
- `sistema_facturas.db` — Base de datos SQLite
- `salida/Reporte_Facturas_Comex.xlsx` — Excel con grafico
- `salida/*.csv` — Archivos CSV
- `salida/*.json` — Datos en JSON
- `logs/facturador.log` — Logs de ejecucion

## Consultar una factura

```python
from agente_comex_bakerhughes_v3 import buscar_factura

buscar_factura("4634461")
```

## Estructura del proyecto

```
agente-comex/
├── agente_comex_bakerhughes_v3.py   # Script principal
├── requirements.txt                  # Dependencias
├── .env.example                      # Plantilla de configuracion
├── .gitignore                        # Archivos ignorados por Git
├── facturas/                         # PDFs de entrada (no subir a Git)
├── salida/                           # Exportaciones (no subir a Git)
├── logs/                             # Logs (no subir a Git)
├── sistema_facturas.db              # Base de datos (no subir a Git)
└── README.md                         # Este archivo
```

## Arquitectura

```
PDF → Extraer texto → Gemini LLM → Sanitizar campos
                                      ↓
                              Normalizar descripciones
                                      ↓
                              Validar matematicamente
                                      ↓
                    SQLite ← Guardar ← Validar
                      ↓
              Exportar CSV / JSON / Excel
```

## Documentacion

- [Manual de Usuario](MANUAL_DE_USUARIO.md) — Guia completa de instalacion y uso
- [Documentacion Tecnica](DOCUMENTACION_TECNICA.md) — Arquitectura, modelos de datos y API

## Seguridad

- **Nunca subas el archivo `.env`** — contiene tu API key
- **Nunca subas PDFs de facturas** — son documentos comerciales sensibles
- **Nunca subas la base de datos** — contiene datos procesados
- El archivo `.gitignore` ya esta configurado para excluir todo lo sensible

## Licencia

[MIT](LICENSE) — Libre para uso personal y comercial.
