CLAUDE.md — Bot de descarga de facturas Efigas

Descripción del proyecto

Bot de automatización en Python que descarga facturas PDF desde el portal
https://portal.efigas.com.co, partiendo de un archivo Excel con los contratos
pendientes, y actualiza ese mismo Excel con el resultado de cada descarga.
Corre completamente en segundo plano (headless), sin abrir ventanas visibles
de navegador ni de Excel.


Estructura de carpetas del proyecto

efigas-bot/
├── CLAUDE.md               ← este archivo
├── .env                    ← credenciales (NUNCA en git)
├── .env.example            ← plantilla de variables de entorno
├── .gitignore
├── requirements.txt
├── main.py                 ← punto de entrada
├── config.py               ← carga de variables de entorno
├── excel_handler.py        ← lectura con openpyxl (data_only=True) y escritura con win32com (Excel real)
├── browser.py              ← inicialización y helpers de Playwright
├── efigas_portal.py        ← lógica de navegación en el portal
├── downloader.py           ← descarga y guardado del PDF
├── logger.py               ← logging estructurado a archivo y consola
└── PDFs/                   ← carpeta destino de los Pdfs descargados


Variables de entorno (.env)

# Usuario 1
USER1_EMAIL=ljborrero@azimutenergia.co
USER1_PASSWORD=Azimutenergia2025*

# Usuario 2
USER2_EMAIL=10@yopmail.com
USER2_PASSWORD=Azimutenergia2025*

Nunca escribir las credenciales directamente en el código.
El archivo .env debe estar en .gitignore.


Dependencias (requirements.txt)

playwright          # automatización del navegador (headless Chromium)
openpyxl            # lectura del Excel (data_only=True, nunca guarda)
pywin32             # escritura en el Excel vía Excel real en segundo plano (preserva Power Query y fórmulas)
python-dotenv       # carga de variables de entorno

Instalar navegadores de Playwright después de pip install:

playwright install chromium


Estructura del Excel

El archivo Excel tiene una hoja llamada Contratos con las columnas:

Col | Nombre                   | Uso
----|--------------------------|-----------------------------------------------------
A   | Contrato                 | Número de contrato (clave de búsqueda en el portal)
B   | ID Inmueble              | Referencia interna
C   | Cliente                  | Nombre del cliente
D   | Comercializador          | Siempre "Efigas"
E   | Fecha emisión            | Día del mes de emisión
F   | Comentarios para descargar | Notas manuales previas
G   | Descarga                 | Filtro principal → el bot procesa solo "Pendiente"
H   | Referencia               | Referencia adicional
I   | ID                       | Identificador interno
J   | Usuario                  | Email del usuario con el que se debe iniciar sesión
K   | Contraseña               | (puede estar vacía si se usa .env)
L   | Comentarios              | ← AQUÍ escribe el bot el resultado

Regla crítica: el bot SOLO escribe en la columna L (Comentarios) y SOLO
en las filas que ya están filtradas como "Pendiente". No toca ninguna otra
celda, no aplica ni elimina filtros. La escritura se hace con win32com
(Excel real en segundo plano) para preservar Power Query, fórmulas y estilos.
openpyxl se usa únicamente para leer, con data_only=True y sin guardar jamás.


Flujo principal (main.py)

1. Cargar .env con python-dotenv
2. Abrir el Excel con openpyxl (data_only=True) solo para lectura
3. Leer todas las filas donde col G == "Pendiente" (valor calculado por fórmula)
4. Agrupar esas filas por el valor de col J (usuario/email)
5. Para cada grupo de usuario:
   a. Obtener credenciales desde variables de entorno (no desde el Excel)
   b. Lanzar Playwright Chromium en modo headless
   c. Navegar a https://portal.efigas.com.co
   d. Hacer login
   e. Para cada contrato del grupo:
      - Buscar el contrato en el portal
      - Validar mes de la factura
      - Descargar PDF si corresponde
      - Escribir comentario en col L de esa fila
      - Guardar el Excel inmediatamente después de cada fila
   f. Cerrar el navegador del usuario actual
6. Cuando todos los grupos terminan: terminar el proceso limpiamente


Módulo: excel_handler.py

Función: get_pending_rows(filepath)


Abre el Excel con openpyxl.load_workbook(filepath, data_only=True)
Lee la hoja "Contratos"
Itera filas (desde la 2, saltando el encabezado)
Retorna lista de dicts con:
{row_number, contrato, cliente, usuario_email, comentario_previo}
solo para las filas donde col G == "Pendiente" (valor calculado, ignorar mayúsculas/espacios)
Cierra el workbook con wb.close() — NUNCA llama save()


Función: write_comment(filepath, row_number, comment)


Usa win32com.client.Dispatch("Excel.Application") — NO openpyxl
Abre Excel en segundo plano: Visible=False, DisplayAlerts=False, ScreenUpdating=False
Abre el archivo con Workbooks.Open(ruta_absoluta)
Escribe SOLO la celda L{row_number} como string plano: ws.Cells(row_number, 12).Value = comment
Guarda con wb.Save() (no SaveAs)
Cierra el archivo y la instancia de Excel en el bloque finally
Preserva Power Query, fórmulas, estilos y conexiones intactas



Módulo: browser.py

Función: create_browser()


Usa playwright.sync_api o async_api
Lanza Chromium con headless=True
Configura downloads_path apuntando a la carpeta PDFs/ del proyecto
Retorna el objeto browser y page


Función: close_browser(browser)


Cierra el contexto y el browser limpiamente
No deja procesos de Chromium en segundo plano


Manejo de fallos del navegador:


Si page lanza TimeoutError o Error de red:

Loggear el error
Intentar close_browser() (ignorar error si ya está cerrado)
Llamar a create_browser() nuevamente
Volver a hacer login
Continuar desde el contrato que falló (no desde el inicio)






Módulo: efigas_portal.py

Función: login(page, email, password)


Ir a https://portal.efigas.com.co
Esperar y hacer click en el botón "Iniciar sesión" (nav superior)
Esperar el modal de login
Llenar campo "Correo electrónico" con email
Llenar campo "Contraseña" con password
Esperar a que el CAPTCHA de Cloudflare muestre "¡Operación exitosa!"
(selector: texto que contenga "Operación exitosa" dentro del iframe/div de CF)
Hacer click en el botón azul "Iniciar sesión"
Esperar redirección al dashboard (URL cambia o aparece el nombre del usuario)
Retorna True si exitoso, False si falla tras timeout


Función: select_contract(page, contract_number)


En el dashboard, hacer click en el menú desplegable "Casa ▾"
Esperar que aparezca la lista de contratos/casas
Buscar el item cuyo número coincida con contract_number
(el portal muestra "Casa # {número}")
Si se encuentra: hacer click en ese item, esperar que el dashboard cargue → retorna True
Si no se encuentra después de revisar toda la lista: retorna False
(el bot escribirá "Contrato no existe" en el Excel)
NO navega a "Mis facturas" ni a ninguna otra página — queda en el dashboard del contrato


Función: validate_current_month(page)


Directamente en el dashboard del contrato (sin navegar a /invoices)
Espera que aparezca un texto con patrón "Mes YYYY" (ej: "Junio 2026")
Parsear ese texto y comparar con el mes y año actuales (datetime.now())
Si coincide: retorna True
Si no coincide: retorna False
(el bot escribirá "Sin emisión" en el Excel)


Función: click_ver_factura(page)


Hacer click en el enlace/botón "Ver factura" del contrato seleccionado
Esperar a que se abra una nueva pestaña o descargue directamente el PDF
Retorna el objeto download de Playwright o la URL del PDF



Módulo: downloader.py

Función: save_pdf(download_or_url, contract_number, pdfs_folder)


Recibe el objeto Download de Playwright o la URL del PDF
Guarda el archivo como {pdfs_folder}/{contract_number}.pdf
Si ya existe un archivo con ese nombre: sobreescribir (es la factura del mes)
Retorna True si el archivo quedó guardado y tiene tamaño > 0
Retorna False si hubo error o el archivo quedó vacío



Módulo: logger.py


Usar el módulo estándar logging de Python
Escribir logs simultáneamente a:

Consola (stdout) con nivel INFO
Archivo bot_efigas.log con nivel DEBUG



Formato: [YYYY-MM-DD HH:MM:SS] [NIVEL] mensaje
Loggear cada acción importante:

Inicio del bot
Contratos cargados del Excel
Inicio de sesión por usuario
Resultado de cada contrato (comentario escrito)
Errores y reintentos
Cierre limpio del bot






Comentarios que el bot escribe en col L

Situación                              | Texto exacto en col L
---------------------------------------|----------------------
Descarga exitosa                       | Descargada
Error al descargar el PDF              | Error descarga
El contrato no aparece en el portal    | Contrato no existe
La factura no corresponde al mes actual| Sin emisión


Manejo de señales y cierre limpio

En main.py, registrar handlers para SIGINT y SIGTERM:

pythonimport signal, sys

def cleanup(sig, frame):
    logger.info("Señal de interrupción recibida. Cerrando...")
    if browser:
        close_browser(browser)
    if workbook:
        workbook.close()
    sys.exit(0)

signal.signal(signal.SIGINT, cleanup)
signal.signal(signal.SIGTERM, cleanup)

Esto garantiza que si el bot es interrumpido a mitad de la ejecución:


Se cierra el navegador Chromium
Se cierra el workbook de Excel
No quedan procesos huérfanos en segundo plano



Checklist de validaciones que debe implementar el bot


 El archivo Excel existe antes de intentar abrirlo
 La carpeta PDFs/ existe; crearla si no existe (os.makedirs)
 Las credenciales del usuario están en .env antes de intentar login
 El CAPTCHA de Cloudflare está resuelto antes de hacer click en "Iniciar sesión"
 El contrato buscado existe en la lista de contratos del portal
 La factura mostrada corresponde al mes y año actuales
 El PDF descargado tiene tamaño > 0 bytes antes de reportar éxito
 Cada comentario se escribe y guarda en el Excel inmediatamente (no al final)
 Al terminar todos los contratos de un usuario, se cierra su sesión en el portal
 Al terminar todos los usuarios, no quedan procesos de Chromium ni de Python activos



Orden de implementación sugerido para Claude Code


Crear la estructura de carpetas y archivos vacíos
Implementar config.py y .env.example
Implementar logger.py
Implementar excel_handler.py con sus dos funciones y tests manuales
Implementar browser.py (lanzar Chromium headless, verificar que abre el portal)
Implementar efigas_portal.py: primero login(), luego select_contract(),
luego validate_current_month(), luego click_ver_factura()
Implementar downloader.py
Ensamblar main.py con el flujo completo y los handlers de señales
Prueba de integración con un solo contrato real
Prueba con múltiples contratos y dos usuarios



Notas adicionales


El bot corre mes a mes. Se ejecuta manualmente o por tarea programada (cron/Task Scheduler).
No necesita interfaz gráfica. Todo el output va al log.
Si el portal cambia su estructura HTML, solo hay que actualizar los selectores en efigas_portal.py.
Usar page.wait_for_selector() con timeouts explícitos (ej: 15s) en cada acción del portal.
Para el CAPTCHA de Cloudflare: esperar con page.wait_for_selector('[text*="Operación exitosa"]', timeout=30000) o similar.
Playwright es preferible a Selenium para este caso porque maneja mejor las descargas de archivos y los popups de nueva pestaña.