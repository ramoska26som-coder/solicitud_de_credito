/**
 * MÁS PISTO — Recepción de solicitudes de crédito en Google Sheets
 * Pega este código en Extensiones > Apps Script de tu hoja de cálculo.
 */

function doPost(e) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Solicitudes')
             || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Solicitudes');

    var datos = JSON.parse(e.postData.contents);

    // Orden de columnas (encabezados). Se crean automáticamente la primera vez.
    var columnas = [
      'fecha_envio','nombre','cedula','telefono','nacimiento','monto',
      'empresa','rubro','dir_empresa','puesto','antiguedad','salario','dias_pago',
      'prestaciones','otros_ingresos','metodo_pago','destino','refiere','refiere_tel',
      'direccion','tiempo_vivienda','tipo_vivienda',
      'ref1_nombre','ref1_tel','ref2_nombre','ref2_tel','consentimiento'
    ];

    if (hoja.getLastRow() === 0) {
      hoja.appendRow(columnas);
      hoja.getRange(1, 1, 1, columnas.length)
          .setFontWeight('bold')
          .setBackground('#1877F2')
          .setFontColor('#FFFFFF');
      hoja.setFrozenRows(1);
    }

    var fila = columnas.map(function(c){ return datos[c] || ''; });
    hoja.appendRow(fila);

    return ContentService
      .createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
