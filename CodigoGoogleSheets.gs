/**
 * MÁS PISTO — Backend de solicitudes en Google Sheets
 *
 * Este script hace 3 cosas:
 *   1. Recibe las solicitudes del formulario (doPost) y las guarda en la hoja "Solicitudes".
 *   2. Entrega el catálogo de rubros al formulario (doGet ?action=rubros),
 *      tomándolo de la hoja "Rubros" (columna A).
 *   3. Incluye una función de prueba (probar) para verificar que todo funciona.
 *
 * INSTALACIÓN
 *   a) En tu Google Sheet: Extensiones > Apps Script. Borra todo y pega este código. Guarda.
 *   b) Ejecuta una vez la función "prepararHojas" (menú Ejecutar) y autoriza los permisos.
 *      Esto crea las hojas "Solicitudes" y "Rubros" con datos de ejemplo.
 *   c) Implementar > Nueva implementación > Aplicación web.
 *        - Ejecutar como: Yo
 *        - Quién tiene acceso: CUALQUIERA
 *      Copia la URL que termina en /exec y pégala en index.html (SHEETS_URL).
 *   d) IMPORTANTE: cada vez que edites este código, ve a Implementar >
 *      Administrar implementaciones > editar (lápiz) > Versión: Nueva. Si no,
 *      los cambios NO se aplican y por eso "no se guarda".
 */

var HOJA_SOLICITUDES = 'Solicitudes';
var HOJA_RUBROS = 'Rubros';

var COLUMNAS = [
  'fecha_envio','nombre','cedula','telefono','nacimiento','monto',
  'empresa','rubro','dir_empresa','puesto','antiguedad','salario',
  'prestaciones','frecuencia_pago','dias_pago','tiene_otros','otros_ingresos',
  'metodo_pago','destino','refiere','refiere_tel',
  'direccion','tiempo_vivienda','tipo_vivienda','tipo_alquiler',
  'ref1_nombre','ref1_parentesco','ref1_tel',
  'ref2_nombre','ref2_parentesco','ref2_tel',
  'comprobantes','consentimiento'
];

/* ---------- GUARDAR SOLICITUD ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = ss.getSheetByName(HOJA_SOLICITUDES) || ss.insertSheet(HOJA_SOLICITUDES);

    var datos = JSON.parse(e.postData.contents);

    if (hoja.getLastRow() === 0) escribirEncabezado(hoja);

    var fila = COLUMNAS.map(function(c){ return datos[c] != null ? datos[c] : ''; });
    hoja.appendRow(fila);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- ENTREGAR RUBROS ---------- */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'rubros') {
    return json({ rubros: leerRubros() });
  }
  return json({ ok: true, msg: 'Más Pisto API activa' });
}

function leerRubros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_RUBROS);
  if (!hoja || hoja.getLastRow() < 2) return [];
  var valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
  var lista = valores.map(function(r){ return String(r[0]).trim(); }).filter(String);
  if (lista.indexOf('Otro') === -1) lista.push('Otro'); // siempre incluir "Otro"
  return lista;
}

/* ---------- UTILIDADES ---------- */
function escribirEncabezado(hoja) {
  hoja.getRange(1, 1, 1, COLUMNAS.length).setValues([COLUMNAS])
      .setFontWeight('bold').setBackground('#1877F2').setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- EJECUTAR UNA VEZ AL INSTALAR ---------- */
function prepararHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var s = ss.getSheetByName(HOJA_SOLICITUDES) || ss.insertSheet(HOJA_SOLICITUDES);
  if (s.getLastRow() === 0) escribirEncabezado(s);

  var r = ss.getSheetByName(HOJA_RUBROS) || ss.insertSheet(HOJA_RUBROS);
  if (r.getLastRow() === 0) {
    r.getRange(1, 1).setValue('Rubro').setFontWeight('bold')
     .setBackground('#1877F2').setFontColor('#FFFFFF');
    var ejemplos = ['Comercio','Servicios','Manufactura','Construcción','Transporte',
                    'Agricultura','Salud','Educación','Gobierno','Restaurante / Comida',
                    'Maquila','Financiera'];
    r.getRange(2, 1, ejemplos.length, 1).setValues(ejemplos.map(function(x){return [x];}));
    r.setFrozenRows(1);
  }
  SpreadsheetApp.getUi().alert('Listo. Hojas "Solicitudes" y "Rubros" preparadas. Ahora edita los rubros a tu gusto y publica la Web App.');
}

/* ---------- PRUEBA DE GUARDADO ---------- */
function probar() {
  var fake = { postData: { contents: JSON.stringify({
    fecha_envio: new Date().toLocaleString('es-HN'),
    nombre: 'PRUEBA Juan Pérez', cedula: '0801-1990-12345', telefono: '9999-9999',
    monto: '15000', empresa: 'Empresa Demo', rubro: 'Comercio',
    salario: '18000', dias_pago: 'Quincenal (día 15 y 30)',
    ref1_nombre: 'María Pérez', ref1_parentesco: 'Madre', ref1_tel: '8888-8888',
    comprobantes: 'Sí, comprobantes de pago',
    consentimiento: 'Aceptado'
  })}};
  var res = doPost(fake);
  Logger.log(res.getContent());  // debe mostrar {"ok":true}
}
