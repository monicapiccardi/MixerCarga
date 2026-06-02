// ============================================================
//  MIXER CARGA — Backend Google Apps Script (API REST)
//  Se despliega como "Aplicación web" y responde JSON
//  El frontend vive en GitHub Pages y lo llama con fetch()
// ============================================================

// Pegá acá el ID de tu Google Sheet (está en la URL de la planilla)
var SPREADSHEET_ID = "PEGAR_ID_DE_TU_GOOGLE_SHEET_AQUI";

var SHEET_REGISTROS = "Registros";
var SHEET_RESUMEN   = "Resumen Mensual";

// Ingredientes. Podés agregar más al final de la lista.
var INGREDIENTES = [
  "Silo Sorgo",
  "Silo AA",
  "Silo Mz",
  "Mz entero",
  "Expeller/harina",
  "Núcleo",
  "Sales anión.",
  "Mz molido",
  "Rollos (unidades)"
];

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET — lecturas
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : "";
  if (action === "ingredientes") return jsonResponse({ ok: true, data: INGREDIENTES });
  if (action === "historial")    return jsonResponse({ ok: true, data: getHistorial() });
  return jsonResponse({ ok: true, mensaje: "Mixer Carga API activa" });
}

// POST — escritura
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === "guardar") return jsonResponse(guardarRegistro(payload.datos));
    return jsonResponse({ ok: false, mensaje: "Acción desconocida: " + payload.action });
  } catch (err) {
    return jsonResponse({ ok: false, mensaje: "Error en servidor: " + err.message });
  }
}

function guardarRegistro(datos) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = obtenerOCrearHoja(ss, SHEET_REGISTROS);

    if (sheet.getLastRow() === 0) {
      var headers = ["Timestamp","Fecha","Turno","Categoría"].concat(INGREDIENTES).concat(["KG TOTALES"]);
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight("bold").setBackground("#2d6a4f").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    var totalKg = 0;
    var valores = INGREDIENTES.map(function(ing) {
      var val = parseFloat(datos.ingredientes[ing]) || 0;
      if (!ing.toLowerCase().includes("rollo")) totalKg += val;
      return val;
    });

    var fila = [new Date(), datos.fecha, datos.turno, datos.categoria].concat(valores).concat([totalKg]);
    sheet.appendRow(fila);

    var uf = sheet.getLastRow();
    sheet.getRange(uf,1).setNumberFormat("dd/mm/yyyy hh:mm");
    sheet.getRange(uf,2).setNumberFormat("dd/mm/yyyy");
    sheet.getRange(uf,5,1,INGREDIENTES.length+1).setNumberFormat("#,##0.00");

    actualizarResumenMensual(ss);
    return { ok: true, mensaje: "Registro guardado correctamente" };
  } catch (err) {
    return { ok: false, mensaje: err.message };
  }
}

function getHistorial() {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_REGISTROS);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var lastRow  = sheet.getLastRow();
    var startRow = Math.max(2, lastRow - 19);
    var data     = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn()).getValues();
    return data.reverse().map(function(fila) {
      var fecha = fila[1];
      if (fecha instanceof Date) fecha = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
      return { fecha: fecha, turno: fila[2], categoria: fila[3], kgTotales: fila[fila.length-1] };
    });
  } catch(err) { return []; }
}

function actualizarResumenMensual(ss) {
  var regSheet = ss.getSheetByName(SHEET_REGISTROS);
  if (!regSheet || regSheet.getLastRow() <= 1) return;
  var resSheet = obtenerOCrearHoja(ss, SHEET_RESUMEN);
  resSheet.clearContents();
  var datos = regSheet.getDataRange().getValues();
  var meses = {};
  for (var i = 1; i < datos.length; i++) {
    var fila  = datos[i];
    var fecha = (fila[1] instanceof Date) ? fila[1] : new Date(fila[1]);
    if (isNaN(fecha.getTime())) continue;
    var mesKey = fecha.getFullYear()+"-"+String(fecha.getMonth()+1).padStart(2,"0");
    var mesLabel = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][fecha.getMonth()]+" "+fecha.getFullYear();
    if (!meses[mesKey]) { meses[mesKey]={ label: mesLabel, totales:{} }; INGREDIENTES.forEach(function(ing){meses[mesKey].totales[ing]=0;}); meses[mesKey].totales["KG TOTALES"]=0; }
    INGREDIENTES.forEach(function(ing,j){ meses[mesKey].totales[ing]+=parseFloat(fila[4+j])||0; });
    meses[mesKey].totales["KG TOTALES"]+=parseFloat(fila[4+INGREDIENTES.length])||0;
  }
  var resHeaders = ["Mes"].concat(INGREDIENTES).concat(["KG TOTALES"]);
  resSheet.appendRow(resHeaders);
  resSheet.getRange(1,1,1,resHeaders.length).setFontWeight("bold").setBackground("#1b4332").setFontColor("#ffffff");
  Object.keys(meses).sort().forEach(function(key){
    var m=meses[key]; var row=[m.label];
    INGREDIENTES.forEach(function(ing){row.push(m.totales[ing]);}); row.push(m.totales["KG TOTALES"]);
    resSheet.appendRow(row);
  });
  if (resSheet.getLastRow()>1) resSheet.getRange(2,2,resSheet.getLastRow()-1,INGREDIENTES.length+1).setNumberFormat("#,##0.00");
  resSheet.autoResizeColumns(1, resHeaders.length);
}

function obtenerOCrearHoja(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}
