// ── CONFIGURACIÓN ─────────────────────────────────────────────────
const SHEET_ID         = "PEGA_AQUI_EL_ID_DE_TU_HOJA";
const SHEET_ASISTENCIA = "Asistencia";
const SHEET_CONFIG     = "Config";
const SHEET_EMPLEADOS  = "Empleados";
const DRIVE_FOLDER     = "Fotos Asistencia";

// ── ROUTER ────────────────────────────────────────────────────────
function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === "saveConfig")     return saveConfig(data);
  if (data.action === "saveEmployee")   return saveEmployee(data);
  if (data.action === "deleteEmployee") return deleteEmployee(data);
  return saveAttendance(data);
}

function doGet(e) {
  if (e.parameter.action === "getConfig")    return getConfig();
  if (e.parameter.action === "getEmployees") return getEmployees();
  return respond({ error: "Acción no válida" });
}

// ── GUARDAR ASISTENCIA + FOTO ─────────────────────────────────────
function saveAttendance(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ASISTENCIA) || ss.insertSheet(SHEET_ASISTENCIA);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Proyecto","Nombre","Tipo","Fecha","Hora","Latitud","Longitud","Distancia(m)","Foto"]);
  }

  let photoUrl = "";
  if (data.photo) {
    try {
      const folder   = getOrCreateFolder(DRIVE_FOLDER);
      const imgData  = data.photo.replace(/^data:image\/\w+;base64,/, "");
      const blob     = Utilities.newBlob(Utilities.base64Decode(imgData), "image/jpeg",
        data.nombre + "_" + data.fecha + "_" + data.tipo + ".jpg");
      const file     = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      photoUrl = file.getUrl();
    } catch(err) {
      photoUrl = "Error al guardar foto";
    }
  }

  sheet.appendRow([
    data.proyecto, data.nombre, data.tipo,
    data.fecha, data.hora,
    data.latitud, data.longitud, data.distancia,
    photoUrl
  ]);

  return respond({ ok: true });
}

// ── CONFIGURACIÓN ─────────────────────────────────────────────────
function saveConfig(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) sheet = ss.insertSheet(SHEET_CONFIG);

  sheet.clearContents();
  sheet.appendRow(["clave","valor"]);
  sheet.appendRow(["projectName",  data.projectName]);
  sheet.appendRow(["projectLat",   data.projectLat]);
  sheet.appendRow(["projectLng",   data.projectLng]);
  sheet.appendRow(["radiusMeters", data.radiusMeters]);
  sheet.appendRow(["entradaFrom",  data.entradaFrom]);
  sheet.appendRow(["entradaTo",    data.entradaTo]);
  sheet.appendRow(["salidaFrom",   data.salidaFrom]);
  sheet.appendRow(["salidaTo",     data.salidaTo]);
  return respond({ ok: true });
}

function getConfig() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) return respond({ error: "Sin config" });
  const rows   = sheet.getDataRange().getValues();
  const config = {};
  rows.forEach(([key, val]) => { if (key !== "clave") config[key] = val; });
  return respond(config);
}

// ── EMPLEADOS ─────────────────────────────────────────────────────
function saveEmployee(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SHEET_EMPLEADOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_EMPLEADOS);
    sheet.appendRow(["id","nombre","pin"]);
  }
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.getRange(i + 1, 1, 1, 3).setValues([[data.id, data.nombre, data.pin]]);
      return respond({ ok: true, updated: true });
    }
  }
  sheet.appendRow([data.id, data.nombre, data.pin]);
  return respond({ ok: true, created: true });
}

function deleteEmployee(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMPLEADOS);
  if (!sheet) return respond({ error: "Sin hoja de empleados" });
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      return respond({ ok: true });
    }
  }
  return respond({ error: "Empleado no encontrado" });
}

function getEmployees() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_EMPLEADOS);
  if (!sheet) return respond({ employees: [] });
  const rows      = sheet.getDataRange().getValues();
  const employees = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) employees.push({ id: String(rows[i][0]), nombre: rows[i][1], pin: String(rows[i][2]) });
  }
  return respond({ employees });
}

// ── HELPER ────────────────────────────────────────────────────────
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
