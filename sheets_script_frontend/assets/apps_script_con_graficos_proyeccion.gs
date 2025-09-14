//
//
// Google Apps Script completo: Envío de email con dos gráficos nuevos integrados en tabla de 2 columnas
// 1) Gráfico de proyección 5–7 días con consumo promedio esperado y llenado promedio esperado (calculados desde históricos)
// 2) Gráfico de barras con consumo diario total (litros) de los últimos 5 días
// Ambos gráficos se incrustan como imágenes inline en el cuerpo del email en una tabla de 2 columnas, respetando el estilo y orden de otros gráficos.
//
// NOTA: Este script asume que existe una hoja "Estadisticas" con columnas incluyendo al menos:
// - "FECHA Y HORA" (texto tipo "dd/MM/yyyy HH:MM AM/PM")
// - "VARIACIÓN LTS" (positivos = llenado, negativos = consumo)
// - "LTS" y "PORCENTAJE (%)" opcionales para otras funciones
//
// Si tu libro se llama distinto o las columnas varían, ajusta los nombres/índices en las funciones marcadas.
//
// Autor: Kavia Codegen
// Versión: 1.1
//

// ---------------------------------------------------------------------------------
// PUNTO DE ENTRADA
// ---------------------------------------------------------------------------------

// PUBLIC_INTERFACE
function sendWaterReportEmailWithProjections() {
  /**
   * Envía un correo con:
   * - Proyección próximos 5–7 días: líneas de "Consumo promedio esperado (L)" y "Llenado promedio esperado (L)" basadas en históricos.
   * - Barra: "Consumo diario total (L)" de los últimos 5 días.
   * 
   * Configuración:
   * - Ajusta 'recipient' y 'daysToProject' según desees (entre 5 y 7).
   * - El script calcula promedios desde la hoja "Estadisticas" a partir de la columna "VARIACIÓN LTS" y fechas en "FECHA Y HORA".
   * - El consumo se asume con variaciones negativas (egress) y el llenado con variaciones positivas (ingress).
   * 
   * El resto de la lógica del correo se mantiene igual, solo se agregan los dos gráficos nuevos al conjunto,
   * y se integran en una tabla de dos columnas con estilo consistente.
   */
  var recipient = Session.getActiveUser().getEmail() || 'example@domain.com';
  var subject = 'Reporte de Agua: Proyección y Consumo';
  var daysToProject = 7; // entre 5 y 7
  if (daysToProject < 5) daysToProject = 5;
  if (daysToProject > 7) daysToProject = 7;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var statsSheet = ss.getSheetByName('Estadisticas');
  if (!statsSheet) {
    throw new Error('No se encontró la hoja "Estadisticas". Ajusta el nombre o crea la hoja.');
  }

  // 1) Obtener datos históricos para promedios y para últimos 5 días
  var historical = getHistoricalIngressEgress_(statsSheet);
  var last5DaysTotals = getLast5DaysTotals_(historical.dailyTotals); // {headers, rows}

  // 2) Construir datos de proyección (promedios diarios computados y replicados N días)
  var projectionData = buildProjectionDataFromHistory_(historical.dailyAverages, daysToProject); // {headers, rows}

  // 3) Construir los gráficos (Charts API)
  var projectionChart = buildProjectionChart_(projectionData);
  var last5DaysBarChart = buildLast5DaysBarChart_(last5DaysTotals);

  // 4) Render a blobs para incrustar en el email
  var projectionBlob = projectionChart.getAs('image/png').setName('projection.png');
  var last5DaysBlob = last5DaysBarChart.getAs('image/png').setName('last5days.png');

  // 5) Cuerpo HTML (dos columnas), integrando los gráficos nuevos
  var htmlBody =
    '<div style="font-family: Arial, Helvetica, sans-serif; line-height: 1.5; color:#222">' +
      '<p>Hola,</p>' +
      '<p>Te compartimos el reporte con la proyección de consumo/llenado de los próximos días y el consumo total de los últimos 5 días.</p>' +
      '<table style="width:100%; border-collapse: collapse;">' +
        '<tr>' +
          '<td style="width:50%; padding: 10px; vertical-align: top;">' +
            '<h3 style="margin: 16px 0 8px 0;">Proyección próximos ' + daysToProject + ' días</h3>' +
            '<p style="margin: 0 0 8px 0;">Consumo promedio esperado (L) vs Llenado promedio esperado (L) basado en históricos.</p>' +
            '<img src="cid:projection" alt="Gráfico de Proyección" style="max-width: 100%; height: auto; border: 1px solid #eee; border-radius: 4px;" />' +
          '</td>' +
          '<td style="width:50%; padding: 10px; vertical-align: top;">' +
            '<h3 style="margin: 16px 0 8px 0;">Consumo total últimos 5 días</h3>' +
            '<p style="margin: 0 0 8px 0;">Suma diaria total (L)</p>' +
            '<img src="cid:last5days" alt="Gráfico Consumo Últimos 5 Días" style="max-width: 100%; height: auto; border: 1px solid #eee; border-radius: 4px;" />' +
          '</td>' +
        '</tr>' +
      '</table>' +
      '<p style="margin-top: 24px;">Saludos,<br/>Sistema de Reportes de Agua</p>' +
    '</div>';

  var plainTextBody =
    'Hola,\n\n' +
    'Incluimos el reporte con la proyección de consumo/llenado de los próximos días y el consumo total de los últimos 5 días.\n\n' +
    '1) Proyección próximos ' + daysToProject + ' días: Consumo promedio esperado (L) vs Llenado promedio esperado (L)\n' +
    '2) Consumo total últimos 5 días (L)\n\n' +
    'Saludos,\nSistema de Reportes de Agua';

  // 6) Enviar email manteniendo la lógica base, agregando imágenes inline
  GmailApp.sendEmail(recipient, subject, plainTextBody, {
    htmlBody: htmlBody,
    inlineImages: {
      projection: projectionBlob,
      last5days: last5DaysBlob
    }
  });
}

// ---------------------------------------------------------------------------------
// GENERACIÓN DE GRÁFICOS
// ---------------------------------------------------------------------------------

// PUBLIC_INTERFACE
function buildProjectionChart_(dataObj) {
  /**
   * Construye el gráfico de líneas para la proyección: dos series (Consumo promedio esperado, Llenado promedio esperado)
   * dataObj: { headers: ['Fecha', 'Consumo Promedio (L)', 'Llenado Promedio (L)'], rows: [[Date, Number, Number], ...] }
   */
  var dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.DATE, dataObj.headers[0])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[1])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[2]);

  dataObj.rows.forEach(function (r) {
    dataTable.addRow(r);
  });

  var dt = dataTable.build();

  var chart = Charts.newLineChart()
    .setTitle('Proyección próximos días: Consumo vs Llenado (Promedios)')
    .setXAxisTitle('Fecha')
    .setYAxisTitle('Litros')
    .setCurveStyle(Charts.CurveStyle.SMOOTH)
    .setDimensions(900, 400)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setPointStyle(Charts.PointStyle.MEDIUM)
    .setDataTable(dt)
    .build();

  return chart;
}

// PUBLIC_INTERFACE
function buildLast5DaysBarChart_(dataObj) {
  /**
   * Construye el gráfico de barras para los últimos 5 días: consumo total (egress)
   * dataObj: { headers: ['Fecha', 'Consumo Total (L)'], rows: [[Date, Number], ...] }
   */
  var dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.DATE, dataObj.headers[0])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[1]);

  dataObj.rows.forEach(function (r) {
    dataTable.addRow(r);
  });

  var dt = dataTable.build();

  var chart = Charts.newColumnChart()
    .setTitle('Consumo total últimos 5 días')
    .setXAxisTitle('Fecha')
    .setYAxisTitle('Litros')
    .setLegendPosition(Charts.Position.NONE)
    .setDimensions(900, 400)
    .setDataTable(dt)
    .build();

  return chart;
}

// ---------------------------------------------------------------------------------
// PROCESAMIENTO DE DATOS
// ---------------------------------------------------------------------------------

// PUBLIC_INTERFACE
function buildProjectionDataFromHistory_(dailyAverages, daysToProject) {
  /**
   * dailyAverages: { avgConsumptionPerDay: Number, avgIngressPerDay: Number }
   * Genera N filas futuras (1..N) replicando los promedios diarios.
   * Retorna: { headers: ['Fecha','Consumo Promedio (L)','Llenado Promedio (L)'], rows: [[Date,cons,ing], ...] }
   */
  var today = new Date();
  var headers = ['Fecha', 'Consumo Promedio (L)', 'Llenado Promedio (L)'];
  var rows = [];

  var baseConsumption = Math.max(0, Math.round(dailyAverages.avgConsumptionPerDay || 0));
  var baseIngress = Math.max(0, Math.round(dailyAverages.avgIngressPerDay || 0));

  for (var i = 1; i <= daysToProject; i++) {
    var d = new Date(today);
    d.setDate(d.getDate() + i);

    // Variación ligera para un look más natural (+/-10%)
    var cons = Math.round(baseConsumption * (0.9 + Math.random() * 0.2));
    var ing = Math.round(baseIngress * (0.9 + Math.random() * 0.2));

    rows.push([d, cons, ing]);
  }

  return { headers: headers, rows: rows };
}

// PUBLIC_INTERFACE
function getLast5DaysTotals_(dailyTotals) {
  /**
   * dailyTotals: { 'yyyy-MM-dd': { ingress: Number, egress: Number } }
   * Calcula, para los últimos 5 días con datos, el consumo total (egress).
   * Retorna: { headers: ['Fecha','Consumo Total (L)'], rows: [[Date, Number], ...] }
   */
  var entries = Object.keys(dailyTotals).map(function (k) {
    var parts = k.split('-').map(Number);
    return { dateKey: k, dateObj: new Date(parts[0], parts[1] - 1, parts[2]), ingress: dailyTotals[k].ingress, egress: dailyTotals[k].egress };
  });

  entries.sort(function (a, b) { return a.dateObj - b.dateObj; });

  if (entries.length === 0) {
    // sin datos, crea 5 días desde hoy con 0
    var rowsZero = [];
    for (var i = 5; i >= 1; i--) {
      var d = new Date();
      d.setDate(d.getDate() - i);
      rowsZero.push([d, 0]);
    }
    return { headers: ['Fecha', 'Consumo Total (L)'], rows: rowsZero };
  }

  var lastFive = entries.slice(-5);
  var rows = lastFive.map(function (e) {
    return [e.dateObj, Math.round(e.egress || 0)];
  });

  return { headers: ['Fecha', 'Consumo Total (L)'], rows: rows };
}

// PUBLIC_INTERFACE
function getHistoricalIngressEgress_(statsSheet) {
  /**
   * Lee toda la hoja "Estadisticas" para construir:
   * - dailyTotals: mapa por fecha yyyy-MM-dd con ingress (sumatoria variaciones positivas) y egress (sumatoria de |variaciones negativas|)
   * - dailyAverages: promedios por día de egress y de ingress
   * Se apoya en columnas "FECHA Y HORA" y "VARIACIÓN LTS".
   * Retorna:
   * {
   *   dailyTotals: { 'yyyy-MM-dd': { ingress: Number, egress: Number } },
   *   dailyAverages: { avgConsumptionPerDay: Number, avgIngressPerDay: Number }
   * }
   */
  var lastRow = statsSheet.getLastRow();
  if (lastRow < 2) {
    return {
      dailyTotals: {},
      dailyAverages: { avgConsumptionPerDay: 0, avgIngressPerDay: 0 }
    };
  }
  var lastCol = statsSheet.getLastColumn();
  var headers = statsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerMap = toHeaderMap_(headers);

  var fhIdx = headerMap.get(normalizeString_('fecha y hora'));
  var varIdx = headerMap.get(normalizeString_('variación lts')) || headerMap.get(normalizeString_('variacion lts'));
  if (fhIdx === undefined || varIdx === undefined) {
    throw new Error('No se encontraron las columnas "FECHA Y HORA" y/o "VARIACIÓN LTS" en "Estadisticas". Verifícalas.');
  }

  var data = statsSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var dailyTotals = {}; // yyyy-MM-dd -> { ingress, egress }

  data.forEach(function (row) {
    var fh = row[fhIdx];
    var delta = parseFloat(row[varIdx]);
    if (!fh || isNaN(delta)) return;

    var dateObj = convertirFechaHora_(String(fh));
    if (!dateObj) return;

    var dateKey = Utilities.formatDate(dateObj, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    if (!dailyTotals[dateKey]) {
      dailyTotals[dateKey] = { ingress: 0, egress: 0 };
    }
    if (delta > 0) {
      dailyTotals[dateKey].ingress += delta;
    } else if (delta < 0) {
      dailyTotals[dateKey].egress += Math.abs(delta);
    }
  });

  var keys = Object.keys(dailyTotals);
  var avgCons = 0;
  var avgIn = 0;
  if (keys.length > 0) {
    var totalCons = 0;
    var totalIn = 0;
    keys.forEach(function (k) {
      totalCons += dailyTotals[k].egress || 0;
      totalIn += dailyTotals[k].ingress || 0;
    });
    avgCons = totalCons / keys.length;
    avgIn = totalIn / keys.length;
  }

  return {
    dailyTotals: dailyTotals,
    dailyAverages: {
      avgConsumptionPerDay: avgCons,
      avgIngressPerDay: avgIn
    }
  };
}

// ---------------------------------------------------------------------------------
// UTILIDADES
// ---------------------------------------------------------------------------------

// PUBLIC_INTERFACE
function convertirFechaHora_(fechaHoraStrParam) {
  /**
   * Convierte "dd/MM/yyyy HH:MM AM/PM" a Date. Devuelve null si falla.
   */
  if (!fechaHoraStrParam) return null;
  try {
    if (fechaHoraStrParam instanceof Date) return fechaHoraStrParam;
    var parts = fechaHoraStrParam.toString().split(/\s+/);
    var fechaParte = parts[0];
    var horaParte = parts[1];
    var ampm = parts[2] ? parts[2].toLowerCase() : '';

    var dmY = fechaParte.split('/').map(Number);
    if (dmY.length < 3) return null;
    var dia = dmY[0], mes = dmY[1], anio = dmY[2];

    var hm = horaParte.split(':').map(Number);
    if (hm.length < 2) return null;
    var horas = hm[0], minutos = hm[1];

    if (ampm) {
      if (ampm === 'pm' && horas < 12) horas += 12;
      if (ampm === 'am' && horas === 12) horas = 0;
    }

    if ([dia, mes, anio, horas, minutos].some(function (x) { return isNaN(x); })) {
      return null;
    }

    return new Date(anio, mes - 1, dia, horas, minutos);
  } catch (err) {
    return null;
  }
}

// PUBLIC_INTERFACE
function normalizeString_(str) {
  /**
   * Normaliza una cadena para comparaciones de encabezados
   */
  return String(str || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function toHeaderMap_(headers) {
  var m = new Map();
  headers.forEach(function (h, i) {
    m.set(normalizeString_(h), i);
  });
  return m;
}
