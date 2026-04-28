const SHEET_NAME = "Hoja 1"; // Asegúrate de que coincida con el nombre de tu hoja

const COLUMNS = {
  NOMBRE: 0, CÉDULA: 1, TELÉFONO: 2, MUNICIPIO: 3, DEPARTAMENTO: 4, EDAD: 5,
  ESTADO: 6, COMPROMISO: 7, PERFILES: 8, PERFILES_CONFIRMADOS: 9,
  FECHA_CONTACTO: 10, FECHA_RECORDATORIO: 11, ÚLTIMA_INTERACCION: 12,
  INGRESO_AL_GRUPO: 13, OBSERVACIÓN: 14, AGENTE_ASIGNADO: 15, FECHA_ASIGNACION: 16
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const agenteId = data.agente_id;
    let result = {};

    if (!agenteId && (action === "obtener_carga" || action === "obtener_asignados")) {
      throw new Error("ID de agente no proporcionado");
    }

    if (action === "obtener_carga") {
      result = obtenerCarga(agenteId);
    } else if (action === "obtener_asignados") {
      result = obtenerAsignados(agenteId);
    } else if (action === "actualizar_registro") {
      result = actualizarRegistro(data.registro);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function obtenerCarga(agenteId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const hoy = new Date().toISOString();
    
    let asignados = obtenerAsignados(agenteId);
    let activosCount = asignados.filter(r => r.ESTADO !== "INHABILITADO").length;
    
    if (activosCount >= 50) return asignados;

    let count = activosCount;
    for (let i = 1; i < data.length; i++) {
      if (count >= 50) break;
      
      const asignadoActual = data[i][COLUMNS.AGENTE_ASIGNADO];
      if (!asignadoActual || asignadoActual.toString().trim() === "") {
        const row = i + 1;
        sheet.getRange(row, COLUMNS.AGENTE_ASIGNADO + 1).setValue(agenteId);
        sheet.getRange(row, COLUMNS.FECHA_ASIGNACION + 1).setValue(hoy);
        
        data[i][COLUMNS.AGENTE_ASIGNADO] = agenteId;
        data[i][COLUMNS.FECHA_ASIGNACION] = hoy;
        asignados.push(formatRow(data[i], row));
        count++;
      }
    }
    return asignados;
  } finally {
    lock.releaseLock();
  }
}

function obtenerAsignados(agenteId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let asignados = [];
  const ahora = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const asignadoActual = String(data[i][COLUMNS.AGENTE_ASIGNADO] || '').trim();
    const agenteBuscado = String(agenteId || '').trim();

    if (agenteBuscado !== '' && asignadoActual === agenteBuscado) {
      let estado = data[i][COLUMNS.ESTADO];
      let fechaContactoRaw = data[i][COLUMNS.FECHA_CONTACTO];
      let row = i + 1;
      
      // REGLA 48H: Inhabilitación automática si no responde
      if (fechaContactoRaw && (estado === "CONTACTADO" || estado === "NO_RESPONDIO" || estado === "POR_CONTACTAR")) {
        let fechaContacto = new Date(fechaContactoRaw);
        if (!isNaN(fechaContacto.getTime())) {
          let diffHoras = (ahora - fechaContacto) / (1000 * 60 * 60);
          if (diffHoras > 48) {
            sheet.getRange(row, COLUMNS.ESTADO + 1).setValue("INHABILITADO");
            sheet.getRange(row, COLUMNS.OBSERVACIÓN + 1).setValue("Inhabilitación automática (48h sin respuesta)");
            data[i][COLUMNS.ESTADO] = "INHABILITADO";
          }
        }
      }
      asignados.push(formatRow(data[i], row));
    }
  }
  return asignados;
}

function actualizarRegistro(registro) {
  const sheet = getSheet();
  const row = registro.id; 
  const hoy = new Date().toISOString();
  
  sheet.getRange(row, COLUMNS.ESTADO + 1).setValue(registro.ESTADO || "");
  sheet.getRange(row, COLUMNS.COMPROMISO + 1).setValue(registro.COMPROMISO || "");
  sheet.getRange(row, COLUMNS.PERFILES + 1).setValue(registro.PERFILES || "");
  sheet.getRange(row, COLUMNS.PERFILES_CONFIRMADOS + 1).setValue(registro.PERFILES_CONFIRMADOS || "");
  sheet.getRange(row, COLUMNS.INGRESO_AL_GRUPO + 1).setValue(registro.INGRESO_AL_GRUPO || "");
  sheet.getRange(row, COLUMNS.OBSERVACIÓN + 1).setValue(registro.OBSERVACION || "");
  sheet.getRange(row, COLUMNS.FECHA_CONTACTO + 1).setValue(registro.FECHA_CONTACTO || "");
  sheet.getRange(row, COLUMNS.ÚLTIMA_INTERACCION + 1).setValue(hoy);
  
  return { success: true, id: row };
}

function formatRow(rowArray, rowIndex) {
  return {
    id: rowIndex,
    NOMBRE: rowArray[COLUMNS.NOMBRE],
    CEDULA: rowArray[COLUMNS.CÉDULA],
    TELEFONO: rowArray[COLUMNS.TELÉFONO],
    MUNICIPIO: rowArray[COLUMNS.MUNICIPIO],
    DEPARTAMENTO: rowArray[COLUMNS.DEPARTAMENTO],
    EDAD: rowArray[COLUMNS.EDAD],
    ESTADO: rowArray[COLUMNS.ESTADO],
    COMPROMISO: rowArray[COLUMNS.COMPROMISO],
    PERFILES: rowArray[COLUMNS.PERFILES],
    PERFILES_CONFIRMADOS: rowArray[COLUMNS.PERFILES_CONFIRMADOS],
    FECHA_CONTACTO: rowArray[COLUMNS.FECHA_CONTACTO],
    FECHA_RECORDATORIO: rowArray[COLUMNS.FECHA_RECORDATORIO],
    ULTIMA_INTERACCION: rowArray[COLUMNS.ÚLTIMA_INTERACCION],
    INGRESO_AL_GRUPO: rowArray[COLUMNS.INGRESO_AL_GRUPO],
    OBSERVACION: rowArray[COLUMNS.OBSERVACIÓN],
    AGENTE_ASIGNADO: rowArray[COLUMNS.AGENTE_ASIGNADO],
    FECHA_ASIGNACION: rowArray[COLUMNS.FECHA_ASIGNACION]
  };
}
