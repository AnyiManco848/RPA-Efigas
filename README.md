# Bot de descarga de facturas Efigas

Automatización en Python que descarga facturas PDF desde el portal [portal.efigas.com.co](https://portal.efigas.com.co), partiendo de un archivo Excel con los contratos pendientes, y actualiza ese mismo Excel con el resultado de cada descarga. Corre completamente en segundo plano (headless), sin abrir ventanas visibles de navegador ni de Excel.

---

## Tecnologías utilizadas

| Tecnología | Versión recomendada | Propósito |
|---|---|---|
| Python | 3.10+ | Lenguaje principal |
| [Playwright](https://playwright.dev/python/) | latest | Automatización del navegador headless (Chromium) |
| [openpyxl](https://openpyxl.readthedocs.io/) | latest | Lectura del Excel (`data_only=True`, nunca guarda) |
| [pywin32](https://pypi.org/project/pywin32/) | latest | Escritura en el Excel vía Excel real en segundo plano (preserva Power Query y fórmulas) |
| [python-dotenv](https://pypi.org/project/python-dotenv/) | latest | Carga de credenciales desde archivo `.env` |

---

## Estructura del proyecto

```
RPA Efigas/
├── main.py              # Punto de entrada: orquesta el flujo completo
├── config.py            # Carga de variables de entorno y rutas globales
├── excel_handler.py     # Lectura de filas pendientes y escritura de comentarios en el Excel
├── browser.py           # Inicialización y cierre del navegador Playwright (Chromium headless)
├── efigas_portal.py     # Lógica de navegación: login, selección de contrato, validación y descarga
├── downloader.py        # Guarda el PDF descargado en la carpeta /Pdfs
├── logger.py            # Logging estructurado a consola y a archivo bot_efigas.log
├── requirements.txt     # Dependencias del proyecto
├── .env                 # Credenciales reales (nunca en git)
├── .env.example         # Plantilla de variables de entorno
├── .gitignore
├── Contratos Efigas.xlsx # Archivo Excel de entrada/salida
├── Pdfs/                # Carpeta destino de los PDFs descargados
└── bot_efigas.log       # Log de ejecución generado automáticamente
```

---

## Configuración del proyecto

### Requisitos previos

- Python 3.10 o superior instalado
- `pip` disponible en el PATH

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 2. Instalar el navegador de Playwright

```bash
playwright install chromium
```

### 3. Configurar las credenciales

Copiar el archivo de plantilla y rellenar con las credenciales reales:

```bash
cp .env.example .env
```

Editar `.env` con los datos de cada usuario:

```env
# Usuario 1
USER1_EMAIL=usuario1@dominio.com
USER1_PASSWORD=contraseña1

# Usuario 2
USER2_EMAIL=usuario2@dominio.com
USER2_PASSWORD=contraseña2
```

> El archivo `.env` está en `.gitignore` y nunca debe subirse al repositorio.

### 4. Verificar el archivo Excel

Asegurarse de que existe el archivo `Contratos Efigas.xlsx` en la raíz del proyecto con la hoja `Contratos` y la siguiente estructura de columnas:

| Col | Nombre | Descripción |
|---|---|---|
| A | Contrato | Número de contrato (clave de búsqueda en el portal) |
| B | ID Inmueble | Referencia interna |
| C | Cliente | Nombre del cliente |
| D | Comercializador | Siempre "Efigas" |
| E | Fecha emisión | Día del mes de emisión |
| F | Comentarios para descargar | Notas manuales previas |
| G | Descarga | **Filtro principal** — el bot procesa solo filas con `Pendiente` |
| H | Referencia | Referencia adicional |
| I | ID | Identificador interno |
| J | Usuario | Email del usuario con el que se debe iniciar sesión |
| K | Contraseña | (puede estar vacía si se usa `.env`) |
| L | Comentarios | **El bot escribe aquí el resultado** |

---

## Cómo ejecutar el bot

```bash
python main.py
```

- **PDFs descargados:** se guardan en `Pdfs/{numero_contrato}.pdf`
- **Logs de ejecución:** se escriben en consola y en el archivo `bot_efigas.log`
- Para interrumpir el bot de forma limpia: `Ctrl+C` — cierra el navegador y el Excel antes de salir.

---

## Flujo completo del bot

```
main.py
│
├── 1. Valida que el archivo Excel existe
├── 2. Crea la carpeta Pdfs/ si no existe
├── 3. Lee el Excel → filtra filas donde columna G (Descarga) == "Pendiente"
├── 4. Agrupa los contratos por usuario (columna J)
│
└── Para cada usuario:
    ├── 5.  Obtiene credenciales desde .env (nunca desde el Excel)
    ├── 6.  Lanza Chromium en modo headless
    ├── 7.  Navega a portal.efigas.com.co
    ├── 8.  Espera resolución del CAPTCHA de Cloudflare (Turnstile)
    ├── 9.  Hace login con email y contraseña
    │
    └── Para cada contrato del usuario:
        ├── 10. Abre el menú desplegable "Casa" y selecciona el contrato
        ├── 11. Verifica que el mes de la factura en el dashboard del contrato
        │       corresponda al mes y año actuales ("Junio 2026")
        ├── 12. Hace click en "Ver factura" directamente desde el dashboard
        ├── 13. Guarda el archivo como Pdfs/{numero_contrato}.pdf
        ├── 14. Abre Excel en segundo plano (invisible) con win32com y escribe
        │       solo la celda L{fila} con el resultado
        └── 15. Cierra Excel — Power Query y fórmulas quedan intactos
```

Si el navegador falla en algún punto, el bot lo reinicia automáticamente, vuelve a hacer login y reintenta el contrato que falló.

---

## Comentarios posibles en el Excel (columna L)

| Comentario | Significado |
|---|---|
| `Descargada` | El PDF se descargó y guardó correctamente |
| `Error descarga` | Falló la descarga, el login o hubo un error inesperado |
| `Contrato no existe` | El número de contrato no aparece en el menú "Casa" del portal |
| `Sin emisión` | La factura más reciente no corresponde al mes y año actuales |

---

## Tiempo estimado de ejecución

Basado en los tiempos de espera definidos en el código:

| Etapa | Tiempo aproximado |
|---|---|
| Login (una vez por usuario) | ~25–35 segundos |
| Resolución del CAPTCHA Cloudflare | incluido en el login (~5–15 s) |
| Selección de contrato en el dashboard | ~4–8 segundos |
| Validación del mes de la factura | ~2–4 segundos |
| Descarga del PDF | ~4–10 segundos |
| Escritura en Excel con win32com | ~3–5 segundos |
| **Total por contrato** | **~13–27 segundos** |

Para una ejecución típica con 10 contratos por usuario, se estima un tiempo total de **3 a 6 minutos** por usuario. El login se paga una sola vez; el tiempo dominante es la navegación y la descarga de cada contrato.
