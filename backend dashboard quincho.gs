/**
 * GOOGLE APPS SCRIPT - BACKEND PARA DASHBOARD STM
 * Copiar este contenido en un nuevo proyecto de Google Apps Script.
 */

const STM_MAIN_SPREADSHEET_ID = '16Y723omF3l38Ntq0MUSh-ZMYnRIvYmctrscidox5ktc';

/**
 * Mapeo de pestañas
 */
const SHEETS = {
  CONVENIOS: 'tabla_convenios',
  NOVEDADES: 'novedades',
  USUARIOS: 'user',
  MENSAJES: 'tabla_mensaje_user',
  MENSAJES_RECIBIDOS: 'tabla_mensaje_recibido',
  AFILIACIONES: 'tabla_afiliacion',
  FECHAS_QUINCHO: 'tabla_fechas_quincho',
  SOLICITUD_QUINCHO: 'tabla_solicitud_quincho'
};

function doGet(e) {
  try {
    // 1. Limpieza de registros pasados (Solicitado)
    limpiarRegistrosViejos();

    // Capturamos el userId si viene en la URL (?userId=...)
    let uFilter = (e && e.parameter && e.parameter.userId) ? e.parameter.userId.toString() : null;

    const result = {
      CONVENIOS: getSheetData(SHEETS.CONVENIOS),
      NOVEDADES: getSheetData(SHEETS.NOVEDADES),
      USUARIOS: getSheetData(SHEETS.USUARIOS),
      MENSAJES: getSheetData(SHEETS.MENSAJES, uFilter),
      MENSAJES_RECIBIDOS: getSheetData(SHEETS.MENSAJES_RECIBIDOS),
      AFILIACIONES: getSheetData(SHEETS.AFILIACIONES),
      FECHAS_QUINCHO: getSheetData(SHEETS.FECHAS_QUINCHO),
      SOLICITUD_QUINCHO: getSheetData(SHEETS.SOLICITUD_QUINCHO),
      status: 'SUCCESS'
    };
    
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({ 
      status: 'ERROR', 
      message: 'Error en doGet: ' + err.toString() 
    });
  }
}

function doPost(e) {
  try {
    let requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (err) {
      return createJsonResponse({ status: 'ERROR', message: 'Invalid JSON' });
    }

    const { action, sheet, id, data, userId } = requestData; 
    const ss = SpreadsheetApp.openById(STM_MAIN_SPREADSHEET_ID);
    
    // Función de búsqueda segura
    const getSheet = (name) => {
      if (!name) return null;
      let s = ss.getSheetByName(name);
      if (!s) {
        const norm = name.toString().toLowerCase().trim();
        s = ss.getSheets().find(sh => (sh.getName() || "").toLowerCase().trim() === norm);
      }
      return s;
    };

    // --- ACCIONES ESPECIALES (Atomic) ---
    if (action === 'APPROVE_QUINCHO') {
      const sSol = getSheet(SHEETS.SOLICITUD_QUINCHO);
      const sFec = getSheet(SHEETS.FECHAS_QUINCHO);
      if (!sSol || !sFec) throw new Error("Hojas de Quincho no encontradas");
      
      // 1. CAPTURAR EL ID_USER DIRECTAMENTE DEL EXCEL (Solicitado)
      const rowsSol = sSol.getDataRange().getValues();
      const headersSol = rowsSol[0];
      const colIdUser = headersSol.findIndex(h => (h||"").toString().toLowerCase().trim() === 'id_user');
      let realUserId = null;
      
      for (let i = 1; i < rowsSol.length; i++) {
        if (rowsSol[i][0].toString().trim() === requestData.bookingId.toString().trim()) {
           realUserId = rowsSol[i][colIdUser];
           break;
        }
      }

      // 2. ACTUALIZACIÓN DE TABLAS
      const resSol = updateRecord(sSol, requestData.bookingId, { pendiente: 'NO' });
      const resFec = updateRecord(sFec, requestData.dateId, { estado: 'OCUPADO' });
      SpreadsheetApp.flush();

      // 3. ENVÍO DE MENSAJE DIRECTO (Respetando columnas A, B, C, D, E, F, G)
      try {
        const sMsg = getSheet(SHEETS.MENSAJES);
        if (sMsg && realUserId && Number(realUserId) > 0) {
          const nextMsgId = getNextId(sMsg);
          const msgTexto = requestData.message || "Su solicitud de Quincho fue aprobada. Deberá acercarse al gremio para finalizar el trámite.";
          
          // Insertamos directamente según la imagen de 5 columnas: [id_mensaje, id_user, titulo, mensaje, fecha]
          sMsg.appendRow([
            nextMsgId, 
            realUserId, 
            "Quincho Aprobado", 
            msgTexto,
            new Date().toLocaleString()
          ]);
          console.log("Mensaje insertado en 5 columnas para usuario: " + realUserId);
        }
      } catch (msgErr) {
        console.error("Error al enviar mensaje: " + msgErr.toString());
      }

      return createJsonResponse({ status: 'SUCCESS', message: 'Trámite procesado con éxito', details: { resSol, resFec, realUserId } });
    }

    if (action === 'CANCEL_QUINCHO') {
      const sSolC = getSheet(SHEETS.SOLICITUD_QUINCHO);
      const sFecC = getSheet(SHEETS.FECHAS_QUINCHO);
      if (!sSolC) throw new Error("Hoja de solicitudes no encontrada");

      const resDel = deleteRecord(sSolC, requestData.bookingId, userId);
      let resFecC = { status: 'SKIP' };
      if (requestData.dateId && sFecC) {
        resFecC = updateRecord(sFecC, requestData.dateId, { estado: 'DISPONIBLE' });
      }
      return createJsonResponse({ status: 'SUCCESS', message: 'Anulado', details: { resDel, resFecC } });
    }

    // --- ACCIONES ESTÁNDAR ---
    const sheetName = (sheet && SHEETS[sheet]) ? SHEETS[sheet] : sheet;
    const targetSheet = getSheet(sheetName);

    if (!targetSheet) {
      return createJsonResponse({ status: 'ERROR', message: 'Sheet not found: ' + (sheet || 'undefined') });
    }

    switch (action) {
      case 'CREATE':
        return createRecord(targetSheet, data);
      case 'UPDATE':
        return createJsonResponse(updateRecord(targetSheet, id, data));
      case 'DELETE':
        return deleteRecord(targetSheet, id, userId);
      default:
        throw new Error('Action not supported: ' + action);
    }
  } catch (err) {
    return createJsonResponse({ status: 'ERROR', message: 'Error en doPost: ' + err.toString() });
  }
}

/**
 * Obtiene los datos de una pestaña, con filtro obligatorio por ID de usuario para mensajes
 */
function getSheetData(sheetName, userIdFilter = null) {
  const ss = SpreadsheetApp.openById(STM_MAIN_SPREADSHEET_ID);
  
  // Intento de obtener hoja de forma flexible
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet && sheetName) {
    const sNameNorm = sheetName.toString().toLowerCase().trim();
    sheet = ss.getSheets().find(s => (s.getName() || "").toLowerCase().trim() === sNameNorm);
  }
  
  if (!sheet) {
    console.error("Hoja no encontrada: " + sheetName);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const isMensajes = sheetName === SHEETS.MENSAJES;
  
  const result = [];
  
  rows.forEach(row => {
    const obj = {};
    headers.forEach((header, index) => {
      // FIX: Protección contra celdas vacías en encabezados
      const key = (header || "").toString().toLowerCase().trim();
      if (key) {
        obj[key] = row[index];
      }
    });

    if (isMensajes) {
      // Si es MENSAJES, SOLO incluimos si coincide el ID. 
      // Si no hay filtro, NO incluimos nada (seguridad)
      if (userIdFilter && obj.id_user && obj.id_user.toString() === userIdFilter.toString()) {
        result.push(obj);
      }
    } else {
      // Para otras tablas (CONVENIOS, NOVEDADES), incluimos todo normal
      if (obj.activo === undefined) obj.activo = 'SI';
      result.push(obj);
    }
  });
  
  console.log("Sheet: " + sheetName + " | Filtro: " + userIdFilter + " | Resultados: " + result.length);
  return result;
}

function createRecord(sheet, data) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nextId = getNextId(sheet);
  const sheetNameActual = (sheet.getName() || "").toLowerCase().trim();
  
  // VALIDACIÓN DE DUPLICADOS (Solo para la tabla de usuarios)
  if (sheetNameActual === SHEETS.USUARIOS.toLowerCase().trim()) {
    const usuarioNuevo = (data.usuario || "").toString().toLowerCase().trim();
    const rows = sheet.getDataRange().getValues();
    const head = rows[0].map(h => (h || "").toString().toLowerCase().trim());
    const userCol = head.indexOf("usuario");
    const nameCol = head.indexOf("nombre");
    const apeCol = head.indexOf("apellido");

    if (userCol !== -1) {
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][userCol].toString().toLowerCase().trim() === usuarioNuevo) {
          const nombreExistente = rows[i][nameCol];
          const apellidoExistente = rows[i][apeCol];
          return createJsonResponse({ 
            status: 'ERROR', 
            message: `El nombre de usuario ya existe. Pertenece a: ${nombreExistente} ${apellidoExistente}` 
          });
        }
      }
    }
  }

  const newRow = headers.map(header => {
    const key = (header || "").toString().toLowerCase().trim();
    if (key === 'id' || key === 'id_mensaje') {
      // Solo usa el valor del data si existe y NO está vacío; de lo contrario, asigna nextId
      return (data[key] !== undefined && data[key] !== '' && data[key] !== null) ? data[key] : nextId;
    }
    return data[key] !== undefined ? data[key] : '';
  });
  
  sheet.appendRow(newRow);

  // AUTOMATIZACIÓN: Lógica de bienvenida para nuevos usuarios
  const isUsuariosSheet = sheetNameActual === SHEETS.USUARIOS.toLowerCase().trim();
  let emailStatus = "N/A";

  if (isUsuariosSheet) {
    const ss = sheet.getParent(); // Obtenemos el acceso al libro
    console.log("Creando mensaje de bienvenida interno para ID: " + nextId);
    try {
      // 1. SIEMPRE crear mensaje en la campanita para el usuario
      const msgSheet = ss.getSheetByName(SHEETS.MENSAJES);
      if (msgSheet) {
        msgSheet.appendRow([
          getNextId(msgSheet),
          nextId, // ID del nuevo usuario
          "👋 ¡Bienvenido/a!",
          "Gracias por registrarte en la App de STM. Ya puedes consultar tus beneficios y novedades.",
          new Date()
        ]);
        console.log("Mensaje interno creado con éxito");
      }

      // 2. SOLO enviar email si existe la dirección de correo
      if (data.email && data.email.toString().includes("@")) {
        console.log("Intentando enviar email de bienvenida a: " + data.email);
        enviarEmailBienvenida(data);
        emailStatus = "SENT";
      } else {
        console.log("No se envió email: Dirección vacía o inválida");
      }
    } catch (e) {
      emailStatus = "ERROR: " + e.toString();
      console.error("Error en proceso de bienvenida: " + e.toString());
    }
  }

  return createJsonResponse({ status: 'SUCCESS', id: nextId, email_status: emailStatus });
}

function enviarEmailBienvenida(u) {
  const nombreApp = "STM App";
  const asunto = `¡Bienvenido/a a ${nombreApp}! - Tus credenciales de acceso`;
  
  // Logo en Base64 para incrustar directamente
  const logoBase64 = "iVBORw0KGgoAAAANSUhEUgAABDgAAAVGCAYAAAB7TwuKAAAQAElEQVR4AeydB4AkVbX+v1tV3T05bA4sLDmIoKJixmdGUDE8/z5zQEUMDzMmwKcoCogSFQOCAgqIIBJMgJJzziwLy+YwOzl0qPp/53bXbM/sbGTDhK+2Tt187rm/qu7tc7q6JoA2ERABERABERABERABERABERABERCB8U5g3K9PAY5xf4q1QBEQAREQAREQAREQAREQAREQgQ0TUI+xTkABjrF+BmW/CIiACIiACIiACIiACIiACGwLAppDBEY5AQU4RvkJknkiIAIiIAIiIAIiIAIiIAJjg4CsFAER2L4EFODYvvw1uwiIgAiIgAiIgAiIgAhMFAJapwiIgAhsVQIKcGxVvFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQAREQgc0joFFjk4ACHGPzvMlqERABERABERABERABERCBbUdAs4iACIxKAgpwjMrTIqNEQAREQAREQAREQAREQATGLgFZLgIisD0IKMCxPahrThEQAREQAREQAREQARGYyAS0dhEQARHYCgQU4NgKUKVSBERABERABERABERABJ4LAY0VAREQARHYdAIKcGw6M40QAREQAREQAREQARHYvgQ0uwiIgAiIgAisRUABjrWQqEIEREAEREAEREAExjoB2S8CIiACIiACE4+AAhwT75xrxSIgAiIgAiIgAiIgAiIgAiIgAiIw7ggowDHuTqkWJAIiIAIiIALPnYA0iIAIiIAIiIAIiMBYI6AAx1g7Y7JXBERABERgNBCQDSIgAiIgAiIgAiIgAqOMgAIco+yEyBwREAERGB8EtAoREAEREAEREAEREAER2LYEFODYtrw1mwiIgAiUCegoAiIgAiIgAiIgAiIgAiKwRQkowLFFcUqZCIjAliIgPSIgAiIgAiIgAiIgAiIgAiKwKQQU4NgUWuorAqOHgCwRAREQAREQAREQAREQAREQARGoIqAARxUMZccTAa1FBERABERABERABERABERABLYkAQU4tiRN6RIBERCB7U1A84uACIiACIiACIiACIjAsAI6AAx1vFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQARHYdgQ0kwiIgAiIgAish4ACHOuBoyYREAEREAEREAERGEsEZKsIiIAIiIAITGQCCnBM5LOvtYuACIiACIjAxCKg1YqACIiACIiACIxjAgpwjOOTq6WJgAiIgAiIwKYRUG8REAEREAEREAERGLsEFOAYu+dOlouACIiACGxrAppPBERABERABERABERg1BJQgGPUnhoZJgIiIAJjj4AsFgEREAEREAEREAEREIHtRUABju1FXvOKgAhMRAJaswiIgAiIgAiIgAiIgAiIwFYioADHVgIrtSIgAptDQGNEQAREQAREQAREQAREQAREYPMIKMCxedw0SgS2DwHNKgIiIAIiIAIiIAIiIAIiIAIiMCIBBThGxKLKsUpAdouACIiACIiACIiACIiACIiACExMAgpwTKZZrdWKgAiIgAiIgAiIgAiIgAiIgAiIwLgkoADHkNOqggiIgAiIgAiIgAiIgAiIgAiIgAiIwFglILtFYGISUIBjYp53rVoEREAEREAEREAEREAEJi4BrVwERGBcElCAY1yeVi1KBERABERABERABERABDafgEaKgAiIwFgkoADHWDxrslkEREAEREAEREAERGB7EtDcIiACIiACo5CAAhyj8KTIJBEQAREQAREQAREY2wRkvQiIgAiIgAhsewIKcGx75ppRBERABERABERgohPQ+kVABERABERABLY4AQU4tjhSKRQBERABERABEXiuBDReBERABERABERABDaVgAIcm0pM/UVABERABERg+xOQBSIgAiIgAiIgAiIgAsMIKMAxDIiKIiACIiAC44GA1iACIiACIiACIiACIjDRCCjAMdHOuNYrAiIgAkZAIgIiIAIiIAIiIAIiIALjjIACHOPshGo5IiACW4aAtIiACIiACIiACIiACIiACIwtAgpwjK2zJWtFYLQQkB0iIAIiIAIiIAIiIAIiIAIiMKoIKMAxqk6HjBk/BLQSERABERABERABERABERABERCBbUlAAY5tSVtzrSGgnAiIgAiIgAiIgAiIgAiIgAiIgAhsQQIKcGxBmFtSlXSJgAiIgAiIgAiIgAiIgAiIgAiIgAhsPIGxGuDY+BWqpwiIgAiIgAiIgAiIgAiIgAiIgAiIwFglsNF2K8Cx0ajUUQREQAREQAREQAREQAREQAREQARGGwHZkxJQgCMloVQEREAEREAEREAEREAEREAERGD8EdCKJgwBBTgmzKnWQkVABERABERABERABERABERgbQLV+mEYHwQkABjvFyJrUOERABERABERABERABERCBrUFAOkVABMYIAQU4xsiJkpkiIAIiIAIiIAIiIAIiMDoJyCoREAERGB0EFOAYHedBVoiACIiACIiACIiACIxXAlqXCIiACIjANiGgAMc2waxJREAEREAEREAEREAE1kVA9SIgAiIgAiKwJQgowLElKEqHCIiACIiACIiACGw9AtIsAiIgAiIgAiKwEQQU4NgISOoiAiIgAiIgAiIwmgnINhEQAREQAREQAREAAFCAQ1eBCIiACIiACIiACIiACIiACIjA+CegFU0YAgpwTJhTrYWKgAiIgAissZWdI0ABERCB8U5AKxQBEdggAfskvMGO6iACIiACIiACIiACo4CA5hABERABERABERCB8U5AAY7xfoa1PhEQAREQAREQAREQAREQgXUTUIsIiMB4JqAAx3g+u1qbCIiACIiACIiACIiACIiACIiACIjAIAEFOAZRKCMCIiACIiACIrBtCcgu0REBERABERCB8UxAAY7xfHa1NhEQAREQAREQAREQAREQga1BQDpFQATGCAEFOMbIiZKZIiACIiACIrBpBNRbBERABERABERABMYuAQU4xu65k+UiIAIiIALbmoDmE4EJQCA9v+vKrGuqMT5gQi1RixUBEdh+BPRJYvud/7E8u30aHMvrl+0iIAIiIAIiIAIiIAIiIAJjm4CsV4Bjm513WS8CIiACIiACIiACIiACIiACE5+A1jYhCSjAMSFPuxYtAiIgAiIgAiIgAiIgAiIgAiIgAuOLgAIc4+t8ajUiIAIiIAIiIAMEEREQAREQAREQAREYRkABjmFAVBQBERABERABERiv/7+9+w2VvKzjOP4+5+yhzYQkSbYH8sBlKAi0SnsMRQ95IBk+svxTUIJ/CipP+gBCwqKChUEWl8RERERERABERCB8U5AAY7xeV43vCr1EAEREAERmJAE9t5771P23XffP3/kIx9JDj/88ORzn/tccvTRRyff//73k1NOOSU566yzkl/96lfJv//97+TGG29Mbr/99uS+++5LHn300WT+/PnJokWLkqVLlyYb2u69996We9cjGxpv7cViMenp6UlWrFiRPPPMM94Gs+Wuu+5Kzj///OS8887ztp5xxhnJSSedlBx//PHJsccem3z7299ODjnkkOSNb3xj8qpXvSqxQMuEPNlatAiIgAiIgAhMMAIKcKzjhKta BERABERABEY7gT322OMj5sC/973vTb74xS8mJ554YnLuuecmf/rTn5J77rknefjhh31gYPXq1UmpVEpsY91RDzzwwGG//e1v8ctf/hKnnXYafvjDH+Jb3/oWjjrqKBxxxBH4xCc+gde85jV45StfiZe85CXYb7/9sOeee2Lu3LmYNWsWpk+fjm2xhWGIuro6TJkyBTvuuKO3wWx50YtehPe///340Ic+5G098sgj8eUvfxnf/OY3cdxxx+F73/ee/vrXv+Lvf/87brjhBligxdaez+eT3t7epLu7O3nooYd88Ofqq69OyCIhg4TrT6gzede73pVAmwiIgAiIgAiIwJgjsLgFjkC3UBksAiIgAiIgAmOJwH//93/7OyyOOeYYf5fCP/7xDx+wWLlyZdLX12f+evLYY4/91hz4P/7xj/jJT36Cr3zlK/jwhz8MOuh4wQtegL333tsHBlpaWhAE+i8/k8mgtrYW9fX12GeffXzw5i1veQs+8pGP4Oijj8Ypp5yC8847DwwQwQPmYWBgIFmyZEliAaMrr7wy+c1vfpP87Gc/Sz72sY8l73jHO5IDDjhAwZCx9MKSrSIgAiIgAptDYMyM0aedMXOqZKgIiIAIiMB4IbDvvvte96Uvfcn/nMJ+XnH55Zd7B9qCF0llu+iii/wdFt/97nf9XQpveMMbfMBi8uTJqKmpGS8oRv06stksZsyYAQsYvfWtbwUDG/j85z8PBjpw2WWX4c4770ShUPA/o7G7Qq699trkz3/+s//pz8c//vHksMMOS17+8pcrCAJtIiACIjCeCWhto4WAAhyj5UzIDhEQAREQgXFFwH428ulPf9o/2+LCCy/0P4dYtmyZD1888MADrz355JP9wyns5xVvf/vbvQNtwYtxBWGcLsY5N2RlURT5n9HYXSH/9V//hUMOOcT/9OfXv/41GOzwP5exZ5jYc00uvfTS5He/+13y0Y9+NHnPe96THHTQQcnee+99yhCFKoiACIjAeCOg9YjANiIQbKN5NI0IiIAIiIAIjCsC9vyLl73sZcn73ve+5LjjjksuvvjihIGLpKurywcx7GcjP//5z/2zLdjH/xxi2rRp44qBFjMyAfspTHXLEmT/PND7Lkm73znO/HBD34Q55xzDhj4AgMedifIUdddd51/aOoJJ5zgH/paPV55ERCB8U9AKxQBEdgyBBTg2DIcpUUEREAERGAcE7C/xvG///u/yZlnnplcddVVyW233eaff3HLLbd4J/XYY48Fv43Hvvvui4aGhnFMQkvbWALFYhFxHK+3u935YcEPBsvw2te+1j809etf/7p/6KuPkvGwYMGC5Jprrkl+/OMf+weg6ucu60WqxvFLQCsTAREQgY0ioADHRmFSJxEQAREQgYlA4JWvfGXyP//zP8nnP//55MEHH0zmzZuX2HMxGNTAT3/6U3zmM5/BwQcfjJe+9KUTAYfW+BwIWPBiQw92tQDIhgIhc+bMwZvf/GZ89atf9Q9Avfnmm8G4h79b6G9/+5v/0772c5eXvOQlyXMwV0PHPAEtQAREQAREQAgowGEUJCIgAiIgAhOOwFvf+lb/V0pOOeWUxB4MuWjRouTGG2/EBRdcgFNPPRXPe97zsMsuu8Cei2HO6oQDpAVvdQIWALFry9JNnczuFnrTm97k/7TvOeecA7t258+fb9dwcsUVVyQMiiQKelRRVVYEREAERGBCEFCAY0KcZi1SBERABCY2Abut/8gjj/QPd7Q7Mzo6OpIrr7zS/5WSo446CvZgyFmzZk1sSBNi9XaTQyrja8H2117mzp2LV77ylTj00ENxzTXX4Pbbb0dbW1tyxx13JOeee67/yy4vfvGLDcBai1eFCIiACIiACIwHAgpwjIezqDWIgAiIgAgMEth///1Xv+Utb0lOP/10/212d3d3Yrf1n3HGGf7hjnZnRlNT02B/ZTaWgPnFI8kI44d3G6GLr0r7+cKWPqTK15du9JxjtmNraysY1MCHP/xh/5ddGOzwP3G54YYbkrPOOss/JHfMLk6Gi4AIiIAIiHUb0UoAABAASURBVMAwAgpwDAOiogiIgAiIwNgi8La3vS357ne/m1x++eXJ448/ntx9990tV199NT72sY/5b7Pr6+vH1oJGpbUWJDDDLE3FyqmsI3WsN2Ey4m5tJiM2qnJrEnjVq16FI444wj8k157pYQ8zveyyy5Kvf/3ryWGHHWYneWtOL90iIAIiIAIisF6+N0//rV//d0v/T/m++r1/T//7SP/3pP+9J/2v+r7y31+h5ntH3u/8/ul/67+7+q/7V82XWvp/l/999H3p79XfU++r/p/6e5P+Lrn2e1n6PnPuL79ff2/9/fT36e8pde9X//r7p79/+/3rz6fPqb9X/p3m/56m/2/6/83U7pOfV/96UP1Bf0f/v/7+lP/n9PfK9f8h/67+/vXzf6++9/6Zev9+/37+f/v968+Pfi/9u6df/l81v6v5XfR/Z/p/r9yLfi/1v6f+XP3zS/3+Nfef+v8l/b2oP+v/3t+v/ky9r5V+N/ky6ZpS76v/7/5Xzf9v6/+9P6H//0T/zOnvRf+//0//T53Tv6/+/vX/3h//OvJvGv3vq9s3/1D/0N/n0n+bfj/c+zukfntnr/73/W9L/VHVd7j6fvC/r9TfI7+L1Pumbv/+/r330vfl95/rz9W/a7rvov+Y/98/8D/pd52/j1Lvi/9Z+C/qn6t5jn4P9XuljJL0518Pp75X9S9V+j2u+/vpl87/wb+XpO+Y/98Tqv+/RL8P/3tLev/8/Rf9fan/31L/PvnfVfq7IX1//Bf9Tv37Sj+L7ueSP8f/Dv0s+t//v67/L3vm/394v9Dn3O+n/vfo981/T/z3fq9Lf0d9Wk+vcz+u/vmlfv+a+079/5L+XtSf9X/v71d/pt7XSr+bfJn8DtT3xX83170/T//fgHfrT1j07330AAAAAElRU5ErkJggg==";
  
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(logoBase64), "image/png", "logo_stm.png");
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #334155; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background: #4f46e5; padding: 25px; text-align: center;">
        <img src="cid:logo_stm" alt="STM Logo" style="width: 80px; height: auto; margin-bottom: 10px;">
        <h1 style="color: white; margin: 0; font-size: 24px;">¡Bienvenido/a, ${u.nombre}!</h1>
      </div>
      <div style="padding: 30px;">
        <p>Le damos la bienvenida a la aplicación oficial del <b>Sindicato de Trabajadores Municipales</b>. Desde ahora, podrá acceder a toda nuestra información y servicios desde su dispositivo.</p>
        
        <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin: 25px 0; border-left: 4px solid #4f46e5;">
          <p style="margin-top: 0; color: #4f46e5; font-weight: bold; text-transform: uppercase; font-size: 12px;">Tus datos de acceso:</p>
          <p style="margin: 5px 0;"><b>Usuario:</b> ${u.usuario || u.user}</p>
          <p style="margin: 5px 0;"><b>Contraseña temporal:</b> ${u.password}</p>
        </div>

        <p style="font-size: 14px; color: #64748b;"><b>Recomendación de seguridad:</b> Por su tranquilidad, le recomendamos cambiar esta contraseña inicial. Puede hacerlo desde el menú <b>"Mi Cuenta"</b> (icono de usuario) ubicado en la barra inferior de la aplicación.</p>
        
        <div style="text-align: center; margin-top: 35px;">
          <a href="#" style="background: #4f46e5; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">Acceder a la App</a>
        </div>
      </div>
      <div style="background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #94a3b8;">
        © 2026 Sindicato de Trabajadores Municipales. Todos los derechos reservados.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: u.email,
    subject: asunto,
    htmlBody: htmlBody,
    inlineImages: {
      logo_stm: logoBlob
    }
  });
}

function updateRecord(sheet, id, data) {
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let rowIndex = -1;
  let rowData = null;

  for (let i = 1; i < rows.length; i++) {
    const valSheet = rows[i][0].toString().trim();
    const valInput = id.toString().trim();
    // Comparación robusta: Texto o Número
    if (valSheet === valInput || (!isNaN(valSheet) && !isNaN(valInput) && parseFloat(valSheet) === parseFloat(valInput))) {
      rowIndex = i + 1;
      rowData = rows[i];
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Record not found');

  headers.forEach((header, index) => {
    // FIX: Protección contra celdas vacías en encabezados
    const key = (header || "").toString().toLowerCase().trim();
    // No actualizamos el ID y verificamos que el campo exista en data
    if (key && key !== 'id' && key !== 'id_mensaje' && data[key] !== undefined) {
      sheet.getRange(rowIndex, index + 1).setValue(data[key]);
    }
  });

  // Forzar que los valores se vean en el Excel ya mismo
  SpreadsheetApp.flush();

  // AUTOMATIZACIÓN: Enviar email y registrar mensaje si se actualiza la contraseña en tabla de usuarios
  let emailStatus = "N/A";
  let messageStatus = "N/A";
  const sheetName = (sheet.getName() || "").toLowerCase().trim();
  if (sheetName === SHEETS.USUARIOS.toLowerCase().trim() && data.password) {
    // Extraer datos del usuario desde la fila encontrada
    const userData = {};
    headers.forEach((header, index) => {
      const key = (header || "").toString().toLowerCase().trim();
      if (key) userData[key] = rowData[index];
    });

    // 1. Enviar email de notificación
    if (userData.email) {
      try {
        enviarEmailCambioPassword({
          nombre: userData.nombre || userData.name || 'Afiliado/a',
          usuario: userData.usuario || userData.user || '',
          email: userData.email,
          newPassword: data.password
        });
        emailStatus = "SENT";
        console.log("Email de cambio de contraseña enviado a: " + userData.email);
      } catch (e) {
        emailStatus = "ERROR: " + e.toString();
        console.error("Fallo al enviar email de cambio de contraseña: " + e.toString());
      }
    }

    // 2. Registrar mensaje en tabla_mensaje_user para que aparezca la campanita
    try {
      const ss = SpreadsheetApp.openById(STM_MAIN_SPREADSHEET_ID);
      const msgSheet = ss.getSheetByName(SHEETS.MENSAJES);
      if (msgSheet) {
        const nextMsgId = getNextId(msgSheet);
        const now = new Date();
        const fechaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        
        // El orden de columnas debe coincidir con los encabezados de 'tabla_mensaje_user':
        // [id_mensaje, id_user, titulo, mensaje, fecha]
        msgSheet.appendRow([
          nextMsgId,
          id, 
          "🔐 Contraseña Actualizada", 
          "Su contraseña se ha cambiado satisfactoriamente el " + fechaStr + ".",
          now
        ]);
        messageStatus = "CREATED";
        console.log("Mensaje de cambio de contraseña creado para usuario ID: " + id);
      }
    } catch (e) {
      messageStatus = "ERROR: " + e.toString();
      console.error("Fallo al crear mensaje de cambio de contraseña: " + e.toString());
    }
  }

  return createJsonResponse({ status: 'SUCCESS', email_status: emailStatus, message_status: messageStatus });
}

function deleteRecord(sheet, id, userId = null) {
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  let rowIndex = -1;

  // FIX: Protección contra celdas vacías en encabezados
  const userColIndex = headers.findIndex(h => (h || "").toString().toLowerCase().trim() === 'id_user');

  for (let i = 1; i < rows.length; i++) {
    const valSheet = rows[i][0].toString().trim();
    const valInput = id.toString().trim();
    // Comparación robusta: Texto o Número
    if (valSheet === valInput || (!isNaN(valSheet) && !isNaN(valInput) && parseFloat(valSheet) === parseFloat(valInput))) {
      // Si se proporciona userId y la tabla tiene columna de usuario, validar
      if (userId && userColIndex !== -1) {
        if (rows[i][userColIndex].toString().trim() !== userId.toString().trim()) {
          throw new Error('No tienes permiso para borrar este registro');
        }
      }
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error('Record not found');

  sheet.deleteRow(rowIndex);
  return createJsonResponse({ status: 'SUCCESS' });
}

function getNextId(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return 1;
  const ids = rows.slice(1)
    .map(r => parseFloat(r[0]))
    .filter(id => !isNaN(id));
  
  if (ids.length === 0) return 1;
  return Math.max(...ids) + 1;
}

function enviarEmailCambioPassword(u) {
  const nombreApp = "STM App";
  const asunto = `🔐 Contraseña Modificada - ${nombreApp}`;
  
  // Mismo logo Base64 que el email de bienvenida
  const logoBase64 = "iVBORw0KGgoAAAANSUhEUgAABDgAAAVGCAYAAAB7TwuKAAAQAElEQVR4AeydB4AkVbX+v1tV3T05bA4sLDmIoKJixmdGUDE8/z5zQEUMDzMmwKcoCogSFQOCAgqIIBJMgJJzziwLy+YwOzl0qPp/53bXbM/sbGTDhK+2Tt187rm/qu7tc7q6JoA2ERABERABERABERABERABERABERCB8U5g3K9PAY5xf4q1QBEQAREQAREQAREQAREQAREQgQ0TUI+xTkABjrF+BmW/CIiACIiACIiACIiACIiACGwLAppDBEY5AQU4RvkJknkiIAIiIAIiIAIiIAIiIAJjg4CsFAER2L4EFODYvvw1uwiIgAiIgAiIgAiIgAhMFAJapwiIgAhsVQIKcGxVvFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQAREQgc0joFFjk4ACHGPzvMlqERABERABERABERABERCBbUdAs4iACIxKAgpwjMrTIqNEQAREQAREQAREQAREQATGLgFZLgIisD0IKMCxPahrThEQAREQAREQAREQARGYyAS0dhEQARHYCgQU4NgKUKVSBERABERABERABERABJ4LAY0VAREQARHYdAIKcGw6M40QAREQAREQAREQARHYvgQ0uwiIgAiIgAisRUABjrWQqEIEREAEREAEREAExjoB2S8CIiACIiACE4+AAhwT75xrxSIgAiIgAiIgAiIgAiIgAiIgAiIw7ggowDHuTqkWJAIiIAIiIALPnYA0iIAIiIAIiIAIiMBYI6AAx1g7Y7JXBERABERgNBCQDSIgAiIgAiIgAiIgAqOMgAIco+yEyBwREAERGB8EtAoREAEREAEREAEREAER2LYEFODYtrw1mwiIgAiUCegoAiIgAiIgAiIgAiIgAiKwRQkowLFFcUqZCIjAliIgPSIgAiIgAiIgAiIgAiIgAiKwKQQU4NgUWuorAqOHgCwRAREQAREQAREQAREQAREQARGoIqAARxUMZccTAa1FBERABERABERABERABERABLYkAQU4tiRN6RIBERCB7U1A84uACIiACIiACIiACIjAsAI6AAx1vFIuAiIgAiIgAiIgAiIgAhtLQP1EQAREQASeCwEFOJ4LPY0VAREQAREQAREQARHYdgQ0kwiIgAiIgAish4ACHOuBoyYREAEREAEREAERGEsEZKsIiIAIiIAITGQCCnBM5LOvtYuACIiACIjAxCKg1YqACIiACIiACIxjAgpwjOOTq6WJgAiIgAiIwKYRUG8REAEREAEREAERGLsEFOAYu+dOlouACIiACGxrAppPBERABERABERABERg1BJQgGPUnhoZJgIiIAJjj4AsFgEREAEREAEREAEREIHtRUABju1FXvOKgAhMRAJaswiIgAiIgAiIgAiIgAiIwFYioADHVgIrtSIgAptDQGNEQAREQAREQAREQAREQAREYPMIKMCxedw0SgS2DwHNKgIiIAIiIAIiIAIiIAIiIAIiMCIBBThGxKLKsUpAdouACIiACIiACIiACIiACIiACExMAgpwTKZZrdWKgAiIgAiIgAiIgAiIgAiIgAiIwLgkoADHkNOqggiIgAiIgAiIgAiIgAiIgAiIgAiIwFglILtFYGISUIBjYp53rVoEREAEREAEREAEREAEJi4BrVwERGBcElCAY1yeVi1KBERABERABERABERABDafgEaKgAiIwFgkoADHWDxrslkEREAEREAEREAERGB7EtDcIiACIiACo5CAAhyj8KTIJBEQAREQAREQAREY2wRkvQiIgAiIgAhsewIKcGx75ppRBERABERABERgohPQ+kVABERABERABLY4AQU4tjhSKRQBERABERABEXiuBDReBERABERABERABDaVgAIcm0pM/UVABERABERg+xOQBSIgAiIgAiIgAiIgAsMIKMAxDIiKIiACIiAC44GA1iACIiACIiACIiACIjDRCCjAMdHOuNYrAiIgAkZAIgIiIAIiIAIiIAIiIALjjIACHOPshGo5IiACW4aAtIiACIiACIiACIiACIiACIwtAgpwjK2zJWtFYLQQkB0iIAIiIAIiIAIiIAIiIAIiMKoIKMAxqk6HjBk/BLQSERABERABERABERABERABERCBbUlAAY5tSVtzrSGgnAiIgAiIgAiIgAiIgAiIgAiIgAhsQQIKcGxBmFtSlXSJgAiIgAiIgAiIgAiIgAiIgAiIgAhsPIGxGuDY+BWqpwiIgAiIgAiIgAiIgAiIgAiIgAiIwFglsNF2K8Cx0ajUUQREQAREQAREQAREQAREQAREQARGGwHZkxJQgCMloVQEREAEREAEREAEREAEREAERGD8EdCKJgwBBTgmzKnWQkVABERABERABERABERABERgbQLV+mEYHwQkABjvFyJrUOERABERABERABERABERCBrUFAOkVABMYIAQU4xsiJkpkiIAIiIAIiIAIiIAIiMDoJyCoREAERGB0EFOAYHedBVoiACIiACIiACIiACIxXAlqXCIiACIjANiGgAMc2waxJREAEREAEREAEREAE1kVA9SIgAiIgAiKwJQgowLElKEqHCIiACIiACIiACGw9AtIsAiIgAiIgAiKwEQQU4NgISOoiAiIgAiIgAiIwmgnINhEQAREQAREQAREAAFCAQ1eBCIiACIiACIiACIiACIiACIjA+CegFU0YAgpwTJhTrYWKgAiIgAissZWdI0ABERCB8U5AKxQBEdggAfskvMGO6iACIiACIiACIiACo4CA5hABERABERABERCB8U5AAY7xfoa1PhEQAREQAREQAREQAREQgXUTUIsIiMB4JqAAx3g+u1qbCIiACIiACIiACIiACIiACIiACIjAIAEFOAZRKCMCIiACIiACIrBtCcgu0REBERABERCB8UxAAY7xfHa1NhEQAREQAREQAREQAREQga1BQDpFQATGCAEFOMbIiZKZIiACIiACIrBpBNRbBERABERABERABMYuAQU4xu65k+UiIAIiIALbmoDmE4EJQCA9v+vKrGuqMT5gQi1RixUBEdh+BPRJYvud/7E8u30aHMvrl+0iIAIiIAIiIAIiIAIiIAJjm4CsV4Bjm513WS8CIiACIiACIiACIiACIiACE5+A1jYhCSjAMSFPuxYtAiIgAiIgAiIgAiIgAiIgAiIgAuOLgAIc4+t8ajUiIAIiIAIiIAMEEREQAREQAREQAREYRkABjmFAVBQBERABERABERiv/7+9+w2VvKzjOP4+5+yhzYQkSbYH8sBlKAi0SnsMRQ95IBk+svxTUIJ/CipP+gBCwqKChUEWl8RERERERABERCB8U5AAY7xeV43vCr1EAEREAERmJAE9t5771P23XffP3/kIx9JDj/88ORzn/tccvTRRyff//73k1NOOSU566yzkl/96lfJv//97+TGG29Mbr/99uS+++5LHn300WT+/PnJokWLkqVLlyYb2u69996We9cjGxpv7cViMenp6UlWrFiRPPPMM94Gs+Wuu+5Kzj///OS8887ztp5xxhnJSSedlBx//PHJsccem3z7299ODjnkkOSNb3xj8qpXvSqxQMuEPNlatAiIgAiIgAhMMAIKcKzjhKta BERABERABEY7gT322OMj5sC/973vTb74xS8mJ554YnLuuecmf/rTn5J77rknefjhh31gYPXq1UmpVEpsY91RDzzwwGG//e1v8ctf/hKnnXYafvjDH+Jb3/oWjjrqKBxxxBH4xCc+gde85jV45StfiZe85CXYb7/9sOeee2Lu3LmYNWsWpk+fjm2xhWGIuro6TJkyBTvuuKO3wWx50YtehPe///340Ic+5G098sgj8eUvfxnf/OY3cdxxx+F73/ee/vrXv+Lvf/87brjhBligxdaez+eT3t7epLu7O3nooYd88Ofqq69OyCIhg4TrT6gzede73pVAmwiIgAiIgAiIwJgjsLgFjkC3UBksAiIgAiIgAmOJwH//93/7OyyOOeYYf5fCP/7xDx+wWLlyZdLX12f+evLYY4/91hz4P/7xj/jJT36Cr3zlK/jwhz8MOuh4wQtegL333tsHBlpaWhAE+i8/k8mgtrYW9fX12GeffXzw5i1veQs+8pGP4Oijj8Ypp5yC8847DwwQwQPmYWBgIFmyZEliAaMrr7wy+c1vfpP87Gc/Sz72sY8l73jHO5IDDjhAwZCx9MKSrSIgAiIgAptDYMyM0aedMXOqZKgIiIAIiMB4IbDvvvte96Uvfcn/nMJ+XnH55Zd7B9qCF0llu+iii/wdFt/97nf9XQpveMMbfMBi8uTJqKmpGS8oRv06stksZsyYAQsYvfWtbwUDG/j85z8PBjpw2WWX4c4770ShUPA/o7G7Qq699trkz3/+s//pz8c//vHksMMOS17+8pcrCAJtIiACIjCeCWhto4WAAhyj5UzIDhEQAREQgXFFwH428ulPf9o/2+LCCy/0P4dYtmyZD1888MADrz355JP9wyns5xVvf/vbvQNtwYtxBWGcLsY5N2RlURT5n9HYXSH/9V//hUMOOcT/9OfXv/41GOzwP5exZ5jYc00uvfTS5He/+13y0Y9+NHnPe96THHTQQcnee+99yhCFKoiACIjAeCOg9YjANiIQbKN5NI0IiIAIiIAIjCsC9vyLl73sZcn73ve+5LjjjksuvvjihIGLpKurywcx7GcjP//5z/2zLdjH/xxi2rRp44qBFjMyAfspTHXLEmT/PND7Lkm73znO/HBD34Q55xzDhj4AgMedifIUdddd51/aOoJJ5zgH/paPV55ERCB8U9AKxQBEdgyBBTg2DIcpUUEREAERGAcE7C/xvG///u/yZlnnplcddVVyW233eaff3HLLbd4J/XYY48Fv43Hvvvui4aGhnFMQkvbWALFYhFxHK+3u935YcEPBsvw2te+1j809etf/7p/6KuPkvGwYMGC5Jprrkl+/OMf+weg6ucu60WqxvFLQCsTAREQgY0ioADHRmFSJxEQAREQgYlA4JWvfGXyP//zP8nnP//55MEHH0zmzZuX2HMxGNTAT3/6U3zmM5/BwQcfjJe+9KUTAYfW+BwIWPBiQw92tQDIhgIhc+bMwZvf/GZ89atf9Q9Avfnmm8G4h79b6G9/+5v/0772c5eXvOQlyXMwV0PHPAEtQAREQAREQAgowGEUJCIgAiIgAhOOwFvf+lb/V0pOOeWUxB4MuWjRouTGG2/EBRdcgFNPPRXPe97zsMsuu8Cei2HO6oQDpAVvdQIWALFry9JNnczuFnrTm97k/7TvOeecA7t258+fb9dwcsUVVyQMiiQKelRRVVYEREAERGBCEFCAY0KcZi1SBERABCY2Abut/8gjj/QPd7Q7Mzo6OpIrr7zS/5WSo446CvZgyFmzZk1sSBNi9XaTQyrja8H2117mzp2LV77ylTj00ENxzTXX4Pbbb0dbW1tyxx13JOeee67/yy4vfvGLDcBai1eFCIiACIiACIwHAgpwjIezqDWIgAiIgAgMEth///1Xv+Utb0lOP/10/212d3d3Yrf1n3HGGf7hjnZnRlNT02B/ZTaWgPnFI8kI44d3G6GLr0r7+cKWPqTK15du9JxjtmNraysY1MCHP/xh/5ddGOzwP3G54YYbkrPOOss/JHfMLk6Gi4AIiIAIiHUb0UoAABAASURBVMAwAgpwDAOiogiIgAiIwNgi8La3vS357ne/m1x++eXJ448/ntx9990tV199NT72sY/5b7Pr6+vH1oJGpbUWJDDDLE3FyqmsI3WsN2Ey4m5tJiM2qnJrEnjVq16FI444wj8k157pYQ8zveyyy5Kvf/3ryWGHHWYneWtOL90iIAIiIAIisF6+N0//rV//d0v/T/m++r1/T//7SP/3pP+9J/2v+r7y31+h5ntH3u/8/ul/67+7+q/7V82XWvp/l/999H3p79XfU++r/p/6e5P+Lrn2e1n6PnPuL79ff2/9/fT36e8pde9X//r7p79/+/3rz6fPqb9X/p3m/56m/2/6/83U7pOfV/96UP1Bf0f/v/7+lP/n9PfK9f8h/67+/vXzf6++9/6Zev9+/37+f/v968+Pfi/9u6df/l81v6v5XfR/Z/p/r9yLfi/1v6f+XP3zS/3+Nfef+v8l/b2oP+v/3t+v/ky9r5V+N/ky6ZpS76v/7/5Xzf9v6/+9P6H//0T/zOnvRf+//0//T53Tv6/+/vX/3h//OvJvGv3vq9s3/1D/0N/n0n+bfj/c+zukfntnr/73/W9L/VHVd7j6fvC/r9TfI7+L1Pumbv/+/r330vfl95/rz9W/a7rvov+Y/98/8D/pd52/j1Lvi/9Z+C/qn6t5jn4P9XuljJL0518Pp75X9S9V+j2u+/vpl87/wb+XpO+Y/98Tqv+/RL8P/3tLev/8/Rf9fan/31L/PvnfVfq7IX1//Bf9Tv37Sj+L7ueSP8f/Dv0s+t//v67/L3vm/394v9Dn3O+n/vfo981/T/z3fq9Lf0d9Wk+vcz+u/vmlfv+a+079/5L+XtSf9X/v71d/pt7XSr+bfJn8DtT3xX83170/T//fgHfrT1j07330AAAAAElRU5ErkJggg==";
  
  const logoBlob = Utilities.newBlob(Utilities.base64Decode(logoBase64), "image/png", "logo_stm.png");
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #334155; max-width: 600px; margin: 0 auto; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden;">
      <div style="background: #4f46e5; padding: 25px; text-align: center;">
        <img src="cid:logo_stm" alt="STM Logo" style="width: 80px; height: auto; margin-bottom: 10px;">
        <h1 style="color: white; margin: 0; font-size: 24px;">🔐 Contraseña Modificada</h1>
      </div>
      <div style="padding: 30px;">
        <p>Hola <b>${u.nombre}</b>,</p>
        <p>Le informamos que su contraseña de acceso a la aplicación del <b>Sindicato de Trabajadores Municipales</b> ha sido modificada exitosamente.</p>
        
        <div style="background: #f8fafc; border-radius: 8px; padding: 20px; margin: 25px 0; border-left: 4px solid #10b981;">
          <p style="margin-top: 0; color: #10b981; font-weight: bold; text-transform: uppercase; font-size: 12px;">Datos actualizados:</p>
          <p style="margin: 5px 0;"><b>Usuario:</b> ${u.usuario}</p>
          <p style="margin: 5px 0;"><b>Nueva Contraseña:</b> ${u.newPassword}</p>
        </div>

        <div style="background: #fef2f2; border-radius: 8px; padding: 15px; margin: 20px 0; border-left: 4px solid #ef4444;">
          <p style="margin: 0; color: #991b1b; font-size: 13px;"><b>⚠️ Importante:</b> Si usted no realizó este cambio, por favor comuníquese de forma urgente con la Secretaría de Acción Social al <b>472 9815 / 472 3756 interno 137</b>.</p>
        </div>

        <p style="font-size: 14px; color: #64748b;"><b>Recomendación de seguridad:</b> No comparta su contraseña con terceros. Mantenga sus credenciales en un lugar seguro.</p>
        
        <div style="text-align: center; margin-top: 35px;">
          <a href="#" style="background: #4f46e5; color: white; padding: 15px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block;">Acceder a la App</a>
        </div>
      </div>
      <div style="background: #f1f5f9; padding: 20px; text-align: center; font-size: 12px; color: #94a3b8;">
        © 2026 Sindicato de Trabajadores Municipales. Todos los derechos reservados.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: u.email,
    subject: asunto,
    htmlBody: htmlBody,
    inlineImages: {
      logo_stm: logoBlob
    }
  });
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * LÓGICA DE LIMPIEZA SOLICITADA: Borra fechas anteriores a HOY y sus reservas
 */
function limpiarRegistrosViejos() {
  try {
    const ss = SpreadsheetApp.openById(STM_MAIN_SPREADSHEET_ID);
    const sFec = ss.getSheetByName(SHEETS.FECHAS_QUINCHO);
    const sSol = ss.getSheetByName(SHEETS.SOLICITUD_QUINCHO);
    if (!sFec) return;

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);

    const dataFec = sFec.getDataRange().getValues();
    const headersFec = dataFec[0];
    const colFechaFec = headersFec.findIndex(h => (h||"").toString().toLowerCase().trim() === 'fecha');
    
    if (colFechaFec === -1) return;

    let fechasBorradas = [];
    // Borrar de abajo hacia arriba para no perder el índice
    for (let i = dataFec.length - 1; i >= 1; i--) {
      const fechaSheet = dataFec[i][colFechaFec];
      let fechaObj = (fechaSheet instanceof Date) ? fechaSheet : new Date(fechaSheet);
      
      if (fechaObj < hoy) {
        fechasBorradas.push(fechaSheet.toString());
        sFec.deleteRow(i + 1);
      }
    }

    // Borrar solicitudes asociadas
    if (sSol && fechasBorradas.length > 0) {
      const dataSol = sSol.getDataRange().getValues();
      const headersSol = dataSol[0];
      const colFechaSol = headersSol.findIndex(h => (h||"").toString().toLowerCase().trim().includes('fecha_solicitada'));
      
      if (colFechaSol !== -1) {
        for (let j = dataSol.length - 1; j >= 1; j--) {
          if (fechasBorradas.includes(dataSol[j][colFechaSol].toString())) {
            sSol.deleteRow(j + 1);
          }
        }
      }
    }
    SpreadsheetApp.flush();
  } catch (e) {
    console.error("Error limpieza: " + e.toString());
  }
}
