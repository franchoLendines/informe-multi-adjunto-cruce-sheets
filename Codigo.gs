/**
 * ============================================================
 *  INFORME MULTI-ADJUNTO → GOOGLE SHEET CON CRUCE Y FÓRMULAS
 * ============================================================
 *
 *  QUÉ HACE:
 *   1. Busca un correo sin leer con un asunto concreto que trae
 *      varios adjuntos (Excel y/o CSV) a la vez: un informe API
 *      (repartido en dos archivos que se combinan), un listado de
 *      clientes B2B, un listado de pólizas y varios informes de
 *      videoconsulta por proveedor.
 *   2. Crea un Google Sheet nuevo con una pestaña por cada archivo
 *      (CLIENTE B2B, POLIZAS, API), recortando filas/columnas de
 *      cabecera según cada plantilla.
 *   3. Combina los informes de videoconsulta de todos los
 *      proveedores en una única pestaña "CRUCE", añadiendo el
 *      proveedor como columna, y calcula ahí fórmulas de cruce
 *      (XLOOKUP/INDEX-MATCH) contra las pestañas CLIENTE B2B,
 *      POLIZAS y API para validar cada registro.
 *   4. Mueve el Sheet resultante a una carpeta de Drive y marca el
 *      correo como leído.
 *
 *  ACTIVADOR (Editor de Apps Script > Activadores):
 *   • buscarYCrearInforme → con la periodicidad con la que recibas
 *     el correo con los adjuntos (diaria, por ejemplo). También se
 *     puede ejecutar manualmente con el botón ▶️.
 *
 *  REQUISITO: el servicio avanzado "Drive API" debe estar activado
 *  en el proyecto (Editor → Servicios → "+" → Drive API), ya que se
 *  usa para convertir los adjuntos Excel a Google Sheets.
 * ============================================================
 */

/**
 * Lee datos de un archivo adjunto, convirtiendo Excel si es necesario.
 * @param {GmailApp.Attachment} attachment El archivo adjunto a procesar.
 * @return {Array<Array<String>>|null} Un array 2D con los datos o null.
 */
function getDataFromAttachment(attachment) {
  const fileName = attachment.getName();
  const mimeType = attachment.getContentType();

  try {
    if (mimeType === MimeType.CSV || mimeType.startsWith("text/")) {
      Logger.log("Procesando '" + fileName + "' como CSV.");
      return Utilities.parseCsv(attachment.getDataAsString('UTF-8'));
    } else if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === MimeType.MICROSOFT_EXCEL_LEGACY) {
      Logger.log("Convirtiendo el archivo adjunto Excel '" + fileName + "' a Google Sheet.");
      const resource = { title: fileName, mimeType: MimeType.GOOGLE_SHEETS };
      const convertedFile = Drive.Files.insert(resource, attachment.copyBlob(), { convert: true });
      const sheet = SpreadsheetApp.openById(convertedFile.id);
      const data = sheet.getSheets()[0].getDataRange().getValues();
      DriveApp.getFileById(convertedFile.id).setTrashed(true); // Mueve el archivo convertido a la papelera
      return data;
    } else {
      Logger.log("AVISO: El archivo adjunto '" + fileName + "' tiene un formato no compatible.");
      return null;
    }
  } catch (e) {
    Logger.log(`Error al leer los datos del adjunto '${fileName}': ${e.toString()}`);
    return null;
  }
}

/**
 * Función principal para procesar correos con archivos adjuntos.
 */
function buscarYCrearInforme() {
  // --- CONFIGURACIÓN — RELLENA AQUÍ TUS DATOS ---
  const emailAddress = "tu_correo@tudominio.com";
  const emailSubject = "TU_ASUNTO_DE_CORREO";
  const folderId = "TU_ID_CARPETA_DRIVE";
  const reportFileName = "TU_PREFIJO_ADJUNTO_API_1"; // Primer archivo del informe API (con cabecera)
  const historialApiFileName = "TU_PREFIJO_ADJUNTO_API_2"; // Segundo archivo del informe API (se le omite la cabecera al combinar)

  const query = `is:unread to:${emailAddress} subject:"${emailSubject}"`;
  const threads = GmailApp.search(query);
  if (threads.length === 0) {
    Logger.log("No se encontraron correos nuevos para procesar.");
    return;
  }
  const destinationFolder = DriveApp.getFolderById(folderId);
  threads.forEach(thread => {
    const message = thread.getMessages()[0];
    const attachments = message.getAttachments();
    if (attachments.length === 0) {
      Logger.log("No se encontraron archivos adjuntos en el correo.");
      return;
    }
    const timestamp = new Date().toLocaleString("es-ES", { timeZone: "Europe/Madrid" });
    const sheetName = `${emailSubject} - ${timestamp}`;
    const spreadsheet = SpreadsheetApp.create(sheetName);
    Logger.log(`Creado Google Sheet: ${sheetName}`);

    let cruceData = [];
    // Acumula los datos combinados de los dos archivos del informe API
    let apiData = [];
    let b2bSheet, polizasSheet, apiSheet;
    attachments.forEach(attachment => {
      try {
        const fileName = attachment.getName();
        Logger.log(`Archivo adjunto detectado: ${fileName}`);
        let data;
        // Combina los dos archivos del informe API en un único conjunto de datos
        if (fileName.startsWith(reportFileName) || fileName.startsWith(historialApiFileName)) {
          data = getDataFromAttachment(attachment);
          if (data) {
            if (apiData.length === 0) {
              // Si es el primer archivo, se cogen todos los datos (incluida la cabecera)
              apiData = data;
            } else {
              // Si es el segundo archivo, se añaden los datos omitiendo su cabecera (fila 1)
              apiData = apiData.concat(data.slice(1));
            }
            Logger.log(`Datos del archivo '${fileName}' agregados a los datos de la API.`);
          }
        } else if (fileName.startsWith("CLIENTE B2B")) {
          data = getDataFromAttachment(attachment);
          if (data) {
            b2bSheet = spreadsheet.insertSheet("CLIENTE B2B");
            b2bSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
            b2bSheet.deleteRows(1, 14);
            b2bSheet.deleteColumn(3);
            b2bSheet.deleteColumn(1);
          }
        } else if (fileName.startsWith("POLIZA")) {
          data = getDataFromAttachment(attachment);
          if (data) {
            polizasSheet = spreadsheet.insertSheet("POLIZAS");
            polizasSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
            polizasSheet.deleteRows(1, 9);
            polizasSheet.deleteColumn(3);
            polizasSheet.deleteColumn(1);
          }
        } else if (fileName.startsWith("VC PROVEEDOR_A")) {
            // Informe de videoconsultas del Proveedor A: fecha con hora, se separa en fecha/hora/hora
            data = getDataFromAttachment(attachment);
            if(data) for (let i = 1; i < data.length; i++) if (data[i] && data[i].join("").trim()) {
                const dateTime = new Date(data[i][0]);
                const fecha = Utilities.formatDate(dateTime, "GMT", "yyyy-MM-dd");
                cruceData.push([fecha, fecha, fecha, data[i][1], data[i][2], data[i][4], "PROVEEDOR_A"]);
            }
        } else if (fileName.startsWith("VC PROVEEDOR_B")) {
            // Informe de videoconsultas del Proveedor B: ya viene con fecha/hora inicio/hora fin separados
            data = getDataFromAttachment(attachment);
            if(data) for (let i = 1; i < data.length; i++) if (data[i] && data[i].join("").trim()) {
                cruceData.push([data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][5], "PROVEEDOR_B"]);
            }
        } else if (fileName.startsWith("VC PROVEEDOR_C")) {
            // Informe de videoconsultas del Proveedor C: fecha con hora, igual que Proveedor A
            data = getDataFromAttachment(attachment);
            if(data) for (let i = 1; i < data.length; i++) if (data[i] && data[i].join("").trim()) {
                const dateTime = new Date(data[i][0]);
                const fecha = Utilities.formatDate(dateTime, "GMT", "yyyy-MM-dd");
                cruceData.push([fecha, fecha, fecha, data[i][1], data[i][2], data[i][4], "PROVEEDOR_C"]);
            }
        }
      } catch (e) {
        Logger.log(`Error CRÍTICO al procesar archivo adjunto: ${e.toString()}`);
      }
    });
    // Crear la hoja API y escribir los datos combinados DESPUÉS del bucle
    if (apiData.length > 0) {
      apiSheet = spreadsheet.insertSheet("API");
      apiSheet.getRange(1, 1, apiData.length, apiData[0].length).setValues(apiData);
      Logger.log("Hoja 'API' creada con los datos combinados.");
    }
    if (spreadsheet.getSheets().length > 1 && spreadsheet.getSheets()[0].getName() === "Sheet1") {
        spreadsheet.deleteSheet(spreadsheet.getSheets()[0]);
    }
    if (cruceData.length > 0) {
      const cruceSheet = spreadsheet.insertSheet("CRUCE");
      const headers = ["Fecha", "Hora de inicio", "Hora de fin", "Nombre Del Acuerdo", "Mapfre_ID", "ID SERVICIO o MAPFRE ID FAMILIAR", "PROVEEDOR", "NIF", "Columna1", "EMAIL TITULAR", "EMAIL fam titular", "TITULAR DE LA CUENTA", "USUARIO ATENDIDO", "MISMO", "EDAD fam", "x POLIZA", "x API", "DIFERENCIAS", "Recuento", "NUM POLIZA", "F INICIO VIAJE", "F FIN VIAJE", "DENTRO", "DUPLICADO O FUERA", "Nº POLIZA A MANO"];

      cruceSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      cruceSheet.getRange(2, 1, cruceData.length, cruceData[0].length).setValues(cruceData);

      Logger.log("Insertando fórmulas en la hoja CRUCE.");
      const lastRow = cruceSheet.getLastRow();
      for (let i = 2; i <= lastRow; i++) {
        // NOTA: La fórmula para la columna H requiere ARRAYFORMULA para funcionar correctamente con la función MATCH en rangos concatenados.
        cruceSheet.getRange(`H${i}`).setFormula(`=IFERROR(INDEX('CLIENTE B2B'!P:P, MATCH(M${i}, ARRAYFORMULA('CLIENTE B2B'!S:S & " " & 'CLIENTE B2B'!T:T), 0)), "Not found")`);
        cruceSheet.getRange(`I${i}`).setFormula(`=XLOOKUP(E${i},'CLIENTE B2B'!D:D,'CLIENTE B2B'!C:C,,0)`);
        cruceSheet.getRange(`J${i}`).setFormula(`=XLOOKUP(E${i},'CLIENTE B2B'!D:D,'CLIENTE B2B'!F:F,,0)`);
        cruceSheet.getRange(`K${i}`).setFormula(`=IFERROR(XLOOKUP(XLOOKUP(E${i},'CLIENTE B2B'!D:D,'CLIENTE B2B'!C:C,,0),'CLIENTE B2B'!D:D,'CLIENTE B2B'!F:F,,0),"Not found")`);
        cruceSheet.getRange(`L${i}`).setFormula(`=XLOOKUP(E${i}, 'CLIENTE B2B'!D:D, 'CLIENTE B2B'!G:G, "Not found")`);
        cruceSheet.getRange(`M${i}`).setFormula(`=CONCATENATE(XLOOKUP(E${i}, 'CLIENTE B2B'!D:D, 'CLIENTE B2B'!S:S, "Not found"), " ", XLOOKUP(E${i}, 'CLIENTE B2B'!D:D, 'CLIENTE B2B'!T:T, ""))`);
        cruceSheet.getRange(`N${i}`).setFormula(`=IF(L${i}=M${i},"MISMO","FAMILIAR")`);
        cruceSheet.getRange(`O${i}`).setFormula(`=XLOOKUP(E${i},'CLIENTE B2B'!D:D,'CLIENTE B2B'!W:W,,0)`);
        cruceSheet.getRange(`P${i}`).setFormula(`=CONCATENATE(XLOOKUP(I${i}, POLIZAS!D:D, POLIZAS!F:F, "Not found"), "-", XLOOKUP(I${i}, POLIZAS!D:D, POLIZAS!G:G, ""))`);
        cruceSheet.getRange(`Q${i}`).setFormula(`=XLOOKUP(J${i},API!H:H,API!C:C,,0)`);
        cruceSheet.getRange(`R${i}`).setFormula(`=IF(P${i}=Q${i},"","DIFERENTE")`);
        cruceSheet.getRange(`S${i}`).setFormula(`=COUNTIF(POLIZAS!D:D,I${i})`);
        cruceSheet.getRange(`T${i}`).setFormula(`=P${i}`);
        cruceSheet.getRange(`U${i}`).setFormula(`=XLOOKUP(T${i},API!C:C,API!E:E,,0)`);
        cruceSheet.getRange(`V${i}`).setFormula(`=XLOOKUP(T${i},API!C:C,API!F:F,,0)`);
        cruceSheet.getRange(`W${i}`).setFormula(`=IF(AND(A${i}>=U${i},A${i}<=V${i}),"DENTRO","FUERA DE RANGO")`);
        cruceSheet.getRange(`X${i}`).setFormula(`=ARRAY_CONSTRAIN(ARRAYFORMULA(IFS(M${i}=M${i+1},"DOUBLE?",W${i-1}="DOUBLE?","DOUBLE",V${i}="FUERA DE RANGO","FUERA")), 1, 1)`);
      }
      Logger.log("Fórmulas insertadas.");
    }

    const finalFile = DriveApp.getFileById(spreadsheet.getId());
    destinationFolder.addFile(finalFile);
    DriveApp.getRootFolder().removeFile(finalFile);
    Logger.log(`Movido el archivo '${sheetName}' a la carpeta de destino.`);
    message.markRead();
  });

  Logger.log("Proceso completado.");
}
