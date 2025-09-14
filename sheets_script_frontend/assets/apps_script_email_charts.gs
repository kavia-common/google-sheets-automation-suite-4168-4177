//
// Google Apps Script: Send email with two charts embedded in the body
// 1) Line chart: projection for next 5–7 days (expected average water consumption and expected average water filling)
// 2) Bar chart: total daily water consumption (liters) of the last 5 days
//
// This script simulates data if source data is missing.
//
// PUBLIC_INTERFACE
function sendWaterReportEmail() {
  /**
   * Sends an email that embeds two charts:
   * - Projection (next 5–7 days) showing expected average water consumption and expected average water filling.
   * - Bar chart of total water consumption for the last 5 days.
   *
   * Data handling:
   * - If you have actual data in a Sheet, replace the simulate* functions with your data retrieval logic.
   * - Otherwise, the script simulates data with reasonable values.
   *
   * Email embedding:
   * - Charts are rendered to images via blob and embedded inline using inlineImages of GmailApp.
   *
   * Configuration:
   * - Update recipient, subject, intro text, and daysToProject as desired.
   */

  // CONFIG
  var recipient = Session.getActiveUser().getEmail() || 'example@domain.com';
  var subject = 'Reporte de Agua: Proyección y Consumo';
  var introText = 'Hola,\n\nA continuación encontrarás el reporte con la proyección de consumo/llenado de agua para los próximos días y el consumo total de los últimos 5 días.\n';
  var daysToProject = 7; // Use a number between 5 and 7
  if (daysToProject < 5) daysToProject = 5;
  if (daysToProject > 7) daysToProject = 7;

  // Generate data (replace these with actual data fetching from Sheets if available)
  var projectionData = simulateProjectionData(daysToProject);
  var last5DaysData = simulateLast5DaysConsumption();

  // Build charts
  var projectionChart = buildProjectionChart(projectionData);
  var last5DaysBarChart = buildLast5DaysBarChart(last5DaysData);

  // Render to blobs for inline embedding
  var projectionBlob = projectionChart.getAs('image/png').setName('projection.png');
  var last5DaysBlob = last5DaysBarChart.getAs('image/png').setName('last5days.png');

  // HTML body with inline images
  var htmlBody =
    '<div style="font-family: Arial, Helvetica, sans-serif; line-height: 1.5; color:#222">' +
      '<p>Hola,</p>' +
      '<p>' + sanitizeHtml_(introText).replace(/\n/g, '<br/>') + '</p>' +

      '<h3 style="margin: 16px 0 8px 0;">1) Proyección próximos ' + daysToProject + ' días</h3>' +
      '<p style="margin: 0 0 8px 0;">Consumo promedio esperado (L) y Llenado promedio esperado (L)</p>' +
      '<img src="cid:projection" alt="Gráfico de Proyección" style="max-width: 100%; height: auto; border: 1px solid #eee; border-radius: 4px;" />' +

      '<h3 style="margin: 24px 0 8px 0;">2) Consumo total últimos 5 días</h3>' +
      '<p style="margin: 0 0 8px 0;">Consumo total diario (L)</p>' +
      '<img src="cid:last5days" alt="Gráfico Consumo Últimos 5 Días" style="max-width: 100%; height: auto; border: 1px solid #eee; border-radius: 4px;" />' +

      '<p style="margin-top: 24px;">Saludos,<br/>Sistema de Reportes de Agua</p>' +
    '</div>';

  // Plain text fallback
  var plainTextBody =
    'Hola,\n\n' +
    'Adjuntamos el reporte con la proyección de consumo/llenado de agua y el consumo total de los últimos 5 días.\n\n' +
    '1) Proyección próximos ' + daysToProject + ' días: Consumo promedio esperado (L) vs Llenado promedio esperado (L)\n' +
    '2) Consumo total últimos 5 días (L)\n\n' +
    'Saludos,\nSistema de Reportes de Agua';

  // Send email with inline images
  GmailApp.sendEmail(recipient, subject, plainTextBody, {
    htmlBody: htmlBody,
    inlineImages: {
      projection: projectionBlob,
      last5days: last5DaysBlob
    }
  });
}

/**
 * PUBLIC_INTERFACE
 * Simulates projection data for next N days.
 * Returns an object:
 * {
 *   headers: ['Fecha', 'Consumo Promedio (L)', 'Llenado Promedio (L)'],
 *   rows: [ [Date, Number, Number], ... ]
 * }
 */
function simulateProjectionData(days) {
  var today = new Date();
  var headers = ['Fecha', 'Consumo Promedio (L)', 'Llenado Promedio (L)'];
  var rows = [];
  var baseConsumption = 120; // liters
  var baseFill = 150; // liters

  for (var i = 1; i <= days; i++) {
    var d = new Date(today);
    d.setDate(d.getDate() + i);
    // Add some variation
    var consumption = baseConsumption + variation_(-15, 20);
    var fill = baseFill + variation_(-20, 25);
    rows.push([d, Math.max(50, Math.round(consumption)), Math.max(50, Math.round(fill))]);
  }
  return { headers: headers, rows: rows };
}

/**
 * PUBLIC_INTERFACE
 * Simulates last 5 days of total consumption.
 * Returns an object:
 * {
 *   headers: ['Fecha', 'Consumo Total (L)'],
 *   rows: [ [Date, Number], ... ]
 * }
 */
function simulateLast5DaysConsumption() {
  var today = new Date();
  var headers = ['Fecha', 'Consumo Total (L)'];
  var rows = [];
  var baseDailyTotal = 800; // liters

  for (var i = 5; i >= 1; i--) {
    var d = new Date(today);
    d.setDate(d.getDate() - i);
    var total = baseDailyTotal + variation_(-120, 150);
    rows.push([d, Math.max(100, Math.round(total))]);
  }
  return { headers: headers, rows: rows };
}

/**
 * PUBLIC_INTERFACE
 * Builds the projection line chart showing two series:
 * - Expected average water consumption
 * - Expected average water filling
 * Returns a Charts.Chart object (image-renderable).
 */
function buildProjectionChart(dataObj) {
  var dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.DATE, dataObj.headers[0])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[1])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[2]);

  dataObj.rows.forEach(function (r) {
    dataTable.addRow(r);
  });

  var dt = dataTable.build();

  var chart = Charts.newLineChart()
    .setTitle('Proyección próximos días: Consumo vs Llenado')
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

/**
 * PUBLIC_INTERFACE
 * Builds the bar chart for last 5 days total consumption.
 * Returns a Charts.Chart object (image-renderable).
 */
function buildLast5DaysBarChart(dataObj) {
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
    .setDimensions(900, 400)
    .setLegendPosition(Charts.Position.NONE)
    .setDataTable(dt)
    .build();

  return chart;
}

// Utility: Generates a random integer variation between min and max (inclusive)
function variation_(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// Utility: Basic HTML sanitizer for user-provided strings used in HTML body
function sanitizeHtml_(s) {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
