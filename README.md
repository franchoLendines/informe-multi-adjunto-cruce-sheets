# Informe multi-adjunto → Google Sheet con cruce y fórmulas

Script para Google Apps Script que combina varios informes adjuntos en un único correo (clientes B2B, pólizas, un informe API repartido en dos archivos, y videoconsultas de varios proveedores) en un solo Google Sheet, con una pestaña de cruce que valida cada registro mediante fórmulas XLOOKUP/INDEX-MATCH.

## Qué hace

1. **Busca el correo**: localiza un correo sin leer con un asunto concreto que trae varios adjuntos a la vez.
2. **Lee cada adjunto**: soporta tanto CSV como Excel (los Excel se convierten temporalmente a Google Sheet para poder leerlos).
3. **Distribuye los datos según el nombre de archivo**:
   - Los dos archivos del informe **API** se combinan en una única tabla (el segundo se añade sin repetir su cabecera).
   - **CLIENTE B2B** y **POLIZAS**: cada uno se vuelca en su propia pestaña, recortando las filas/columnas de cabecera propias de su plantilla.
   - Los informes de **videoconsulta de cada proveedor** (hasta 3 proveedores distintos, cada uno con su propio formato de fecha/hora) se combinan todos en una tabla común `cruceData`, añadiendo el proveedor como columna extra.
4. **Genera la pestaña CRUCE**: escribe los datos combinados de videoconsultas y añade fórmulas fila a fila que cruzan esos datos contra las pestañas CLIENTE B2B, POLIZAS y API, para detectar coincidencias, duplicados y registros fuera de rango de póliza.
5. **Mueve el Sheet final** a una carpeta de Drive y marca el correo como leído.

## Requisitos previos

- Una carpeta de Google Drive donde se guardará el informe generado.
- El servicio avanzado **Drive API** habilitado en el proyecto de Apps Script (necesario para convertir adjuntos Excel a Google Sheets).
- Que el correo con todos los adjuntos llegue con un asunto identificable, y que cada archivo adjunto tenga un nombre que empiece por un prefijo reconocible (uno por tipo de informe/proveedor).

## Instalación

1. Crea un proyecto nuevo en [script.google.com](https://script.google.com).
2. Copia el contenido de `Codigo.gs` en el editor.
3. Habilita el servicio avanzado **Drive API**: Editor → Servicios (icono "+") → busca "Drive API" → Añadir.
4. Rellena la configuración al principio de `buscarYCrearInforme()`:
   - `emailAddress` / `emailSubject`: destinatario y asunto del correo con los adjuntos.
   - `folderId`: carpeta de Drive donde se guardará el Sheet generado.
   - `reportFileName` / `historialApiFileName`: prefijos de los dos archivos que forman el informe API.
5. Ajusta los prefijos `"VC PROVEEDOR_A"`, `"VC PROVEEDOR_B"`, `"VC PROVEEDOR_C"` (y los tres bloques `else if` correspondientes) a los nombres reales de tus proveedores de videoconsulta y al formato de columnas de cada uno — no todos los proveedores exportan la fecha/hora de la misma forma, así que revisa cada bloque.
6. Revisa los recortes de filas/columnas (`deleteRows`, `deleteColumn`) de las pestañas CLIENTE B2B y POLIZAS: dependen de cuántas filas de cabecera/metadatos traiga tu plantilla real.
7. Revisa las fórmulas de la pestaña CRUCE: están escritas para una estructura de columnas concreta en CLIENTE B2B, POLIZAS y API (por ejemplo, `'CLIENTE B2B'!D:D`, `POLIZAS!D:D`, `API!C:C`). Ajusta las columnas referenciadas a tu propia plantilla.
8. La primera vez que ejecutes la función, Google te pedirá autorizar permisos de Gmail, Sheets y Drive.

## Activador (ejecución automática)

En el editor de Apps Script, ve a **Activadores** (icono de reloj) y crea uno para `buscarYCrearInforme`, con la periodicidad con la que recibas el correo (diaria, por ejemplo). También puedes ejecutarlo manualmente con el botón ▶️ para probarlo.

## Notas

- El código incluido en este repositorio usa **valores de ejemplo** (`TU_ID_CARPETA_DRIVE`, `TU_ASUNTO_DE_CORREO`, `PROVEEDOR_A`/`B`/`C`, etc.) en lugar de los datos y nombres de proveedores reales del entorno original. Sustitúyelos por los tuyos antes de ejecutar.
- Este es el script más dependiente de la estructura exacta de los archivos de origen de todos los que forman esta serie: antes de automatizarlo, ejecútalo una vez manualmente y revisa cada pestaña generada para confirmar que las columnas de las fórmulas de CRUCE apuntan donde deben.
- Si en el futuro cambias de proveedor de videoconsulta o añades uno nuevo, solo hace falta añadir un bloque `else if` más siguiendo el patrón de los existentes.
