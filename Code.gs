/**
 * REGISTRO DE VISITAS — Google Apps Script (backend)
 * ---------------------------------------------------
 * Conserva tu hoja "Registro Visitas", tu orden de columnas y tus acciones
 * getAll / delete (panel admin), hoy e historial.
 *
 * CAMBIO IMPORTANTE (hora confiable):
 *  - La FECHA y la HORA ahora las pone el SERVIDOR, no el teléfono. Así ya no
 *    importa si el reloj de un celular está adelantado o atrasado: todos los
 *    registros quedan con la hora real de Ecuador.
 *  - Excepción: si la visita venía guardada "sin señal" en el teléfono
 *    (llega con offline=1), se respeta la hora que se capturó en el momento
 *    de la visita, porque esa es la hora verdadera de cuando ocurrió.
 *  - insert devuelve { ok, hora, fecha } para que la app muestre la hora real.
 *
 * CAMBIO (lat/lng como número):
 *  - insertRow ahora guarda Latitud/Longitud como Number real (no texto).
 *    Antes, al llegar como texto, la configuración regional de la hoja podía
 *    interpretar el "." como separador de miles y borrar el decimal
 *    (-78.474420 se guardaba como -78474420). Al mandar un Number real,
 *    Apps Script ya no reinterpreta el valor.
 *
 * CAMBIO (getAll con fecha/hora legibles):
 *  - getAll ahora formatea Fecha y Hora igual que hoy/historial (antes
 *    devolvía el Date crudo de la celda, saliendo como ISO completo en el
 *    panel admin).
 *
 * Orden de columnas de tu hoja (NO cambiar):
 *  A: id | B: asesor | C: cliente | D: fecha | E: hora
 *  F: motivo | G: km | H: lat | I: lng | J: acc
 *
 * IMPORTANTE: los encabezados (fila 1) deben decir exactamente, sin espacios
 * de más: ID, Asesor, Cliente, Fecha, Hora, Motivo, Kilometraje, Latitud,
 * Longitud, Precision_GPS — getAll arma cada registro a partir de ese texto,
 * así que un encabezado con un espacio de más rompe ese campo en el panel.
 *
 * Para publicar: Implementar > Administrar implementaciones > (editar) >
 * Versión: "Nueva versión" > Implementar (conservas tu misma URL /exec).
 *
 * Recomendado: en el editor, Configuración del proyecto > Zona horaria =
 * (GMT-05:00) Guayaquil. Aun así, este código fuerza 'America/Guayaquil'
 * para que la hora sea correcta pase lo que pase.
 */

const SHEET_NAME = 'Registro Visitas';

// Zona horaria fija de Ecuador: garantiza que la hora sea siempre correcta.
const TZ = 'America/Guayaquil';

// Índices de columna (0 = A, 1 = B, ...) según TU hoja
const COL = { id:0, asesor:1, cliente:2, fecha:3, hora:4, motivo:5, km:6, lat:7, lng:8, acc:9 };

function doGet(e) {
  const p = e.parameter || {};
  const action = p.action;

  try {
    if (action === 'insert')        return insertRow(p);
    if (action === 'getAll')        return getAll();
    if (action === 'delete')        return deleteRow(p.id);
    if (action === 'hoy')           return visitasHoy(p.asesor);
    if (action === 'historial')     return historial(p.asesor, p.dias);
    if (action === 'getParametros') return getParametros();
    if (action === 'setParametros') return setParametros(p);
    if (action === 'getConfig')     return getConfig();
    if (action === 'setConfig')     return setConfig(p);
    return respond({ error: 'Accion no valida' });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}

/* ================= TUS ACCIONES ORIGINALES ================= */

function getAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return respond({ records: [] });
  const headers = rows[0];
  const records = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    if (obj.Fecha !== undefined) obj.Fecha = fechaStr(obj.Fecha);
    if (obj.Hora  !== undefined) obj.Hora  = horaStr(obj.Hora);
    return obj;
  }).reverse();
  return respond({ records });
}

function deleteRow(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.id]) === String(id)) {
      sheet.deleteRow(i + 1);
      return respond({ ok: true });
    }
  }
  return respond({ error: 'No encontrado' });
}

/* ================= INSERT CON ANTI-DUPLICADOS ================= */
/**
 * Inserta una visita EVITANDO duplicados:
 *  - Mismo id ya existente -> no inserta (doble clic o reintento sin señal).
 *  - INICIO DE ACTIVIDADES: solo uno por asesor por día.
 *  - Mismo asesor + cliente + fecha con hora dentro de los últimos 10 min.
 * En esos casos responde ok:true, dup:true.
 *
 * La FECHA y la HORA las decide el SERVIDOR (reloj confiable), salvo que el
 * registro venga de la cola sin señal del teléfono (offline=1), en cuyo caso
 * se respeta la hora original de la visita.
 */
function insertRow(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que dos envíos simultáneos se dupliquen

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data  = sheet.getDataRange().getValues();
    const norm  = s => String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();
    const esInicio = norm(p.cliente) === 'INICIO DE ACTIVIDADES';

    // ---- Fecha/hora confiables ----
    const now = new Date();
    let fecha, hora;
    if (String(p.offline) === '1' && p.fecha && p.hora) {
      // Visita guardada sin señal: se respeta la hora real de cuando ocurrió.
      fecha = fechaStr(p.fecha);
      hora  = horaStr(p.hora);
    } else {
      // Registro en vivo: la hora la pone el servidor (ignora el reloj del celular).
      fecha = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
      hora  = Utilities.formatDate(now, TZ, 'HH:mm');
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Duplicado exacto por id
      if (String(row[COL.id]) === String(p.id)) {
        return respond({ ok: true, dup: true, fecha: fechaStr(row[COL.fecha]), hora: horaStr(row[COL.hora]) });
      }

      // INICIO DE ACTIVIDADES: solo se permite UNO por asesor por día
      if (esInicio &&
          norm(row[COL.cliente]) === 'INICIO DE ACTIVIDADES' &&
          norm(row[COL.asesor]) === norm(p.asesor) &&
          fechaStr(row[COL.fecha]) === fecha) {
        return respond({ ok: true, dup: true, fecha: fechaStr(row[COL.fecha]), hora: horaStr(row[COL.hora]) });
      }

      // Mismo asesor + cliente + fecha, con menos de 10 minutos de diferencia
      if (norm(row[COL.asesor]) === norm(p.asesor) &&
          norm(row[COL.cliente]) === norm(p.cliente) &&
          fechaStr(row[COL.fecha]) === fecha) {
        const dif = Math.abs(minutos(horaStr(row[COL.hora])) - minutos(hora));
        if (dif <= 10) return respond({ ok: true, dup: true, fecha: fechaStr(row[COL.fecha]), hora: horaStr(row[COL.hora]) });
      }
    }

    // Number real (no texto): evita que la configuración regional de la
    // hoja borre el punto decimal de las coordenadas GPS.
    const toNum = v => (v === '' || v === undefined || v === null) ? '' : Number(v);

    // TU orden de columnas original (usando fecha/hora del servidor)
    sheet.appendRow([
      p.id, p.asesor, p.cliente, fecha, hora,
      p.motivo, p.km, toNum(p.lat), toNum(p.lng), p.acc
    ]);
    return respond({ ok: true, dup: false, fecha: fecha, hora: hora });
  } finally {
    lock.releaseLock();
  }
}

/* ================= CONSULTAS ================= */

/** Visitas de HOY. Si se pasa ?asesor=..., filtra por ese asesor. */
function visitasHoy(asesor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const hoy   = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const norm  = s => String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();

  const visitas = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (fechaStr(row[COL.fecha]) !== hoy) continue;
    if (asesor && norm(row[COL.asesor]) !== norm(asesor)) continue;
    visitas.push({
      id:      String(row[COL.id]),
      asesor:  String(row[COL.asesor]),
      cliente: String(row[COL.cliente]),
      motivo:  String(row[COL.motivo]),
      hora:    horaStr(row[COL.hora])
    });
  }
  return respond({ ok: true, fecha: hoy, total: visitas.length, visitas });
}

/** Historial de un asesor: últimos N días (por defecto 15, máx. 60). */
function historial(asesor, dias) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const norm  = s => String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();

  const n = Math.min(Number(dias) || 15, 60);
  const desde = new Date();
  desde.setDate(desde.getDate() - n);
  const desdeStr = Utilities.formatDate(desde, TZ, 'yyyy-MM-dd');

  const visitas = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const f = fechaStr(row[COL.fecha]);
    if (f < desdeStr) continue;
    if (asesor && norm(row[COL.asesor]) !== norm(asesor)) continue;
    visitas.push({
      id:      String(row[COL.id]),
      cliente: String(row[COL.cliente]),
      motivo:  String(row[COL.motivo]),
      fecha:   f,
      hora:    horaStr(row[COL.hora]),
      km:      Number(row[COL.km]) || 0
    });
  }
  visitas.sort((a, b) => (b.fecha + b.hora).localeCompare(a.fecha + a.hora)); // recientes primero
  return respond({ ok: true, total: visitas.length, visitas });
}

/* ================= PARÁMETROS POR VENDEDOR ================= */
const PARAMS_SHEET = 'Parametros Vendedor';
const PCOL = { asesor:0, vehiculo:1, cilindraje:2, marca:3, placa:4, combustible:5, rendimiento:6, estandar:7, presCobranzas:8, presVentas:9, actualizado:10 };
const PARAMS_HEADERS = ['Asesor','Tipo Vehiculo','Cilindraje','Marca y Modelo','Placa','Combustible','Rendimiento (km/gal)','Estandar Visitas Dia','Presupuesto Cobranzas Mes','Presupuesto Ventas Mes','Actualizado'];

function getParamSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PARAMS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PARAMS_SHEET);
    sheet.appendRow(PARAMS_HEADERS);
  }
  return sheet;
}

function getParametros() {
  const sheet = getParamSheet_();
  const data  = sheet.getDataRange().getValues();
  const registros = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const asesor = String(row[PCOL.asesor] || '').trim();
    if (!asesor) continue;
    registros[asesor] = {
      asesor,
      vehiculo:      String(row[PCOL.vehiculo] || ''),
      cilindraje:    String(row[PCOL.cilindraje] || ''),
      marca:         String(row[PCOL.marca] || ''),
      placa:         String(row[PCOL.placa] || ''),
      combustible:   String(row[PCOL.combustible] || ''),
      rendimiento:   Number(row[PCOL.rendimiento]) || 0,
      estandar:      Number(row[PCOL.estandar]) || 0,
      presCobranzas: Number(row[PCOL.presCobranzas]) || 0,
      presVentas:    Number(row[PCOL.presVentas]) || 0,
      actualizado:   String(row[PCOL.actualizado] || '')
    };
  }
  return respond({ ok: true, registros });
}

function setParametros(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const asesor = String(p.asesor || '').trim();
    if (!asesor) return respond({ ok: false, error: 'Falta asesor' });

    const sheet = getParamSheet_();
    const data  = sheet.getDataRange().getValues();
    const ahora = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');

    const fila = [
      asesor,
      p.vehiculo || '', p.cilindraje || '', p.marca || '', p.placa || '',
      p.combustible || '', Number(p.rendimiento) || 0, Number(p.estandar) || 0,
      Number(p.presCobranzas) || 0, Number(p.presVentas) || 0, ahora
    ];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][PCOL.asesor]).trim() === asesor) {
        sheet.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
        return respond({ ok: true, actualizado: ahora });
      }
    }
    sheet.appendRow(fila);
    return respond({ ok: true, actualizado: ahora });
  } finally {
    lock.releaseLock();
  }
}

/* ================= CONFIGURACIÓN GLOBAL (precios de combustible) =================
 * Una sola fila con los valores editables desde el panel admin: precio del
 * galón por tipo de combustible y el umbral de cumplimiento de presupuesto.
 * Se usan junto con los parámetros de cada vendedor (rendimiento km/gal,
 * combustible) para calcular el viático semanal.
 */
const CONFIG_SHEET = 'Config Global';
const CONFIG_HEADERS = ['Precio Extra','Precio Ecopais','Precio Diesel','Umbral Cumplimiento %','Actualizado'];

function getConfigSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET);
    sheet.appendRow(CONFIG_HEADERS);
    sheet.appendRow([0, 0, 0, 75, '']);
  }
  return sheet;
}

function getConfig() {
  const sheet = getConfigSheet_();
  const row = sheet.getRange(2, 1, 1, CONFIG_HEADERS.length).getValues()[0];
  return respond({
    ok: true,
    config: {
      precioExtra:   Number(row[0]) || 0,
      precioEcopais: Number(row[1]) || 0,
      precioDiesel:  Number(row[2]) || 0,
      umbral:        Number(row[3]) || 75,
      actualizado:   String(row[4] || '')
    }
  });
}

function setConfig(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getConfigSheet_();
    const ahora = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    sheet.getRange(2, 1, 1, CONFIG_HEADERS.length).setValues([[
      Number(p.precioExtra) || 0, Number(p.precioEcopais) || 0, Number(p.precioDiesel) || 0,
      Number(p.umbral) || 75, ahora
    ]]);
    return respond({ ok: true, actualizado: ahora });
  } finally {
    lock.releaseLock();
  }
}

/* ================= REPARACIÓN ÚNICA: LAT/LNG SIN DECIMALES ================
 * Corre esta función UNA sola vez desde el editor (seleccionar
 * "repararLatLng" en el desplegable de funciones y Ejecutar). Corrige las
 * coordenadas históricas que se guardaron sin punto decimal (ej. -78474420
 * en vez de -78.474420) por el bug de configuración regional de la hoja.
 * Solo toca celdas con valor absoluto > 1000, imposible para una coordenada
 * real (siempre entre -180 y 180), así que las filas correctas no se tocan.
 * No hace falta redesplegar para correr esto: se ejecuta directo del editor.
 */
function repararLatLng() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const iLat = headers.indexOf('Latitud');
  const iLng = headers.indexOf('Longitud');
  if (iLat === -1 || iLng === -1) throw new Error('No se encontraron las columnas Latitud/Longitud');

  const nFilas = data.length - 1;
  const latCol = [];
  const lngCol = [];
  let reparadas = 0;

  for (let i = 1; i < data.length; i++) {
    let lat = data[i][iLat];
    let lng = data[i][iLng];
    if (typeof lat === 'number' && Math.abs(lat) > 1000) { lat = lat / 1000000; reparadas++; }
    if (typeof lng === 'number' && Math.abs(lng) > 1000) { lng = lng / 1000000; }
    latCol.push([lat]);
    lngCol.push([lng]);
  }

  sheet.getRange(2, iLat + 1, nFilas, 1).setValues(latCol);
  sheet.getRange(2, iLng + 1, nFilas, 1).setValues(lngCol);
  Logger.log('Filas reparadas: ' + reparadas);
  return reparadas;
}

/* ================= UTILIDADES ================= */

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// La celda de fecha puede llegar como Date o como texto en varios formatos:
// "2026-07-06", "06/07/2026", "6/7/2026". Siempre devuelve "yyyy-MM-dd".
function fechaStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();

  // Formato ISO: 2026-07-06 (con o sin hora después)
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);

  // Formato dd/mm/yyyy o d/m/yyyy (como lo muestra Google Sheets en Ecuador)
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
  if (m) return m[3] + '-' + pad2(m[2]) + '-' + pad2(m[1]);

  return s.slice(0, 10);
}

function pad2(n) { return ('0' + Number(n)).slice(-2); }

// La celda de hora puede llegar como Date o como texto "14:35"
function horaStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, TZ, 'HH:mm');
  }
  return String(v || '').slice(0, 5);
}

function minutos(hhmm) {
  const m = /^(\d{1,2}):(\d{2})/.exec(String(hhmm));
  return m ? (Number(m[1]) * 60 + Number(m[2])) : -9999;
}
