/**
 * @OnlyCurrentDoc
 *
 * VERSION 4.0 (Integrado con Proyección 5–7 días y Consumo total últimos 5 días)
 *
 * RESUMEN:
 * - Mantiene toda la lógica original (ingesta, stats, gráficos, email con tabla 2 columnas).
 * - Agrega cálculo histórico para promedios diarios (ingreso y consumo) y totales diarios.
 * - Genera dos NUEVOS gráficos:
 *    1) Proyección próximos 5–7 días: líneas de Consumo promedio esperado (L) vs Llenado promedio esperado (L).
 *    2) Barras: Consumo diario total (L) de los últimos 5 días.
 * - Inserta ambos gráficos en la tabla de dos columnas del correo, como imágenes inline (CIDs).
 * - La función principal sigue siendo onFormSubmitTrigger.
 */

/* ============================== */
/*     CONFIGURACIÓN GLOBAL       */
/* ============================== */

// Capacidad del tanque en litros (ajustar si cambia)
const TANK_CAPACITY_LITERS = 169000;

/* ============================== */
/*    FUNCIÓN PRINCIPAL (ENTRY)   */
/* ============================== */

// PUBLIC_INTERFACE
function onFormSubmitTrigger(e) {
  /**
   * Punto de entrada: se dispara en envío de formulario.
   * - Inserta registro en "Estadisticas" (orden cronológico, previene duplicados).
   * - Actualiza fórmulas.
   * - Genera gráficos y los adjunta inline en email (tabla de 2 columnas).
   * - Integra además dos gráficos nuevos:
   *     a) Proyección próximos 5–7 días (líneas Consumo vs Llenado promedios).
   *     b) Consumo total de los últimos 5 días (barras).
   * - Los nuevos gráficos se agregan al conjunto y se incrustan con sus CIDs.
   */

  // Si no viene de trigger de formulario, permitir test manual
  if (!e || !e.range) {
    Logger.log("El script no fue invocado por envío de formulario. Ejecutando test manual si es posible.");
    try { testScriptManually(); } catch (err) { Logger.log("Test manual no disponible: " + err.message); }
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheetName = "Respuestas de formulario 1";
  const statsSheetName = "Estadisticas";
  const emailsSheetName = "Emails";
  const chartsSheetName = "Graficos";
  const timeSlotSheetName = "Grafico3";
  const myEmail = "aguatorrebela@gmail.com"; // Email principal del sistema

  const formResponsesSheet = spreadsheet.getSheetByName(formResponsesSheetName);
  let statsSheet = spreadsheet.getSheetByName(statsSheetName);
  let emailsSheet = spreadsheet.getSheetByName(emailsSheetName);
  let chartsSheet = spreadsheet.getSheetByName(chartsSheetName);
  let timeSlotSheet = spreadsheet.getSheetByName(timeSlotSheetName);

  // Crear/validar hojas
  const statsHeaders = [
    "FECHA", "HORA", "LTS", "PORCENTAJE (%)", "ALTURA (m)",
    "VARIACIÓN LTS", "VARIACIÓN PUNTOS %", "FECHA Y HORA", "VARIACIÓN TIEMPO (min) ENTRE MEDICIONES",
    "CAUDAL = CANTD. LTS/min", "LTS/Hora", "LTS QUE FALTAN POR LLENARSE",
    "TIEMPO ESTIMADO PARA LLENARSE TANQUE (min)", "HORAS", "DÍAS",
    "TIEMPO ESTIMADO PARA VACIARSE (min)", "HORAS_VACIADO", "DÍAS_VACIADO",
    "FECHA DEL REGISTRO POR FORMULARIO", "Email", "COLABORADOR",
    "INFO ADICIONAL - Agente de Usuario", "INFO ADICIONAL - IP del Cliente (Est.)",
    "INFO ADICIONAL - Latitud (Est.)", "INFO ADICIONAL - Longitud (Est.)"
  ];

  if (!statsSheet) {
    statsSheet = spreadsheet.insertSheet(statsSheetName);
    statsSheet.getRange(1, 1, 1, statsHeaders.length).setValues([statsHeaders]);
  }

  if (!emailsSheet) {
    emailsSheet = spreadsheet.insertSheet(emailsSheetName);
    emailsSheet.getRange(1, 1, 1, 4).setValues([["Email", "Nombre Colaborador", "Fecha de Registro", "Contador"]]);
  }

  if (!chartsSheet) {
    chartsSheet = spreadsheet.insertSheet(chartsSheetName);
    chartsSheet.getRange(1, 1, 1, 8).setValues([["Fecha y Hora", "Litros", "Porcentaje", "Caudal Llenado", "Caudal Consumo", "Día", "Total Lts Entrantes", "Total Lts Salientes"]]);
  }

  if (!timeSlotSheet) {
    timeSlotSheet = spreadsheet.insertSheet(timeSlotSheetName);
    timeSlotSheet.getRange(1, 1, 1, 2).setValues([["Franja Horaria", "Consumo Promedio (Litros)"]]);
  } else {
    timeSlotSheet.clearContents();
    timeSlotSheet.getRange(1, 1, 1, 2).setValues([["Franja Horaria", "Consumo Promedio (Litros)"]]);
  }

  // Tomar datos del evento
  const rowData = e.range.getValues()[0];
  const marcaTemporal = rowData[0];
  const email = rowData[1];
  const fechaMedicionRaw = rowData[2];
  const horaMedicion = rowData[3];
  const litros = rowData[4];
  const porcentaje = rowData[5];
  const altura = rowData[6];
  const colaborador = rowData[7];

  // Formateo de fecha y hora
  let fechaMedicionFormateada;
  try {
    const fecha = fechaMedicionRaw instanceof Date ? fechaMedicionRaw : new Date(fechaMedicionRaw);
    fechaMedicionFormateada = Utilities.formatDate(fecha, spreadsheet.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  } catch (err) {
    fechaMedicionFormateada = fechaMedicionRaw;
    Logger.log("Error al formatear fecha: " + err.message);
  }

  const horaMedicionFormateada = formatHour(horaMedicion);
  if (!horaMedicionFormateada) {
    Logger.log(`Hora inválida: ${horaMedicion}. No se insertará el registro.`);
    return;
  }
  const fechaHoraStr = `${fechaMedicionFormateada} ${horaMedicionFormateada}`;

  // Detección de duplicados e inserción cronológica
  let isDuplicate = false;
  let insertRow = statsSheet.getLastRow() + 1;
  const statsData = statsSheet.getDataRange().getValues();

  if (convertirFechaHora(fechaHoraStr)) {
    for (let i = 1; i < statsData.length; i++) {
      const fechaHoraExistente = statsData[i][7]; // Col H
      const fechaHoraExistenteDate = convertirFechaHora(fechaHoraExistente);
      const nuevaFechaHora = convertirFechaHora(fechaHoraStr);

      if (fechaHoraExistenteDate && nuevaFechaHora && nuevaFechaHora.getTime() === fechaHoraExistenteDate.getTime()) {
        isDuplicate = true;
        Logger.log(`Registro duplicado detectado para Fecha y Hora: ${fechaHoraStr}. No se insertará nueva fila.`);
        break;
      }

      if (!isDuplicate && nuevaFechaHora && fechaHoraExistenteDate && nuevaFechaHora.getTime() < fechaHoraExistenteDate.getTime()) {
        insertRow = i + 1;
        break;
      }
    }
  } else {
    Logger.log("La fecha y hora del nuevo formulario no pudieron ser procesadas, se omite verificación de duplicados.");
  }

  // Insertar registro si no es duplicado
  if (!isDuplicate) {
    if (insertRow <= statsSheet.getLastRow()) {
      statsSheet.insertRowBefore(insertRow);
    }
    const newRowRange = statsSheet.getRange(insertRow, 1, 1, statsHeaders.length);
    newRowRange.setValues([[
      fechaMedicionFormateada,
      horaMedicionFormateada,
      litros,
      porcentaje,
      Number(altura),
      null, // F VARIACIÓN LTS
      null, // G VARIACIÓN %
      fechaHoraStr, // H FECHA Y HORA
      null, // I VARIACIÓN TIEMPO
      null, // J CAUDAL
      null, // K LTS/Hora
      null, // L LTS QUE FALTAN
      null, // M TIEMPO LLENADO
      null, // N HORAS
      null, // O DÍAS
      null, // P TIEMPO VACIADO
      null, // Q HORAS_VACIADO
      null, // R DÍAS_VACIADO
      marcaTemporal,
      email,
      colaborador,
      "No disponible por Google Forms",
      "No disponible por Google Forms",
      "No disponible",
      "No disponible"
    ]]);

    // Fórmulas
    if (insertRow > 2) {
      const prevRow = insertRow - 1;
      statsSheet.getRange(insertRow, 6).setFormula(`=C${insertRow}-C${prevRow}`);
      statsSheet.getRange(insertRow, 7).setFormula(`=D${insertRow}-D${prevRow}`);

      const variacionTiempoFormula = `=LET(
      inicio_str; H${prevRow};
      fin_str; H${insertRow};
      PARSEAR_FECHA_HORA_GS; LAMBDA(texto_fecha_hora;
        LET(
          fecha_parte_str; LEFT(texto_fecha_hora; FIND(" "; texto_fecha_hora) - 1);
          hora_parte_str; MID(texto_fecha_hora; FIND(" "; texto_fecha_hora) + 1; LEN(texto_fecha_hora));
          fecha_num; DATEVALUE(TRIM(CLEAN(fecha_parte_str)));
          hora_val_num; VALUE(LEFT(hora_parte_str; FIND(":"; hora_parte_str) - 1));
          min_val_num; VALUE(MID(hora_parte_str; FIND(":"; hora_parte_str) + 1; 2));
          ampm_val_str; RIGHT(hora_parte_str; 2);
          hora_24h_num; IF(ampm_val_str = "pm"; hora_val_num + IF(hora_val_num = 12; 0; 12);
                           IF(ampm_val_str = "am"; hora_val_num - IF(hora_val_num = 12; 12; 0); hora_val_num));
          (fecha_num) + ((hora_24h_num * 60) + min_val_num) / 1440
        )
      );
      tiempo_inicio_completo; PARSEAR_FECHA_HORA_GS(inicio_str);
      tiempo_fin_completo; PARSEAR_FECHA_HORA_GS(fin_str);
      (tiempo_fin_completo - tiempo_inicio_completo) * 1440
    )`;
      statsSheet.getRange(insertRow, 9).setFormula(variacionTiempoFormula);

      statsSheet.getRange(insertRow, 10).setFormula(`=F${insertRow}/I${insertRow}`);
      statsSheet.getRange(insertRow, 11).setFormula(`=J${insertRow}*60`);
      statsSheet.getRange(insertRow, 12).setFormula(`=(${TANK_CAPACITY_LITERS}-C${insertRow})`);
      statsSheet.getRange(insertRow, 13).setFormula(`=IF(J${insertRow}>0;L${insertRow}/J${insertRow};"")`);
      statsSheet.getRange(insertRow, 14).setFormula(`=IF(J${insertRow}>0;M${insertRow}/60;"")`);
      statsSheet.getRange(insertRow, 15).setFormula(`=IF(J${insertRow}>0;N${insertRow}/24;"")`);
      statsSheet.getRange(insertRow, 16).setFormula(`=IF(J${insertRow}<0;-C${insertRow}/J${insertRow};"")`);
      statsSheet.getRange(insertRow, 17).setFormula(`=IF(P${insertRow}<>"";P${insertRow}/60;"")`);
      statsSheet.getRange(insertRow, 18).setFormula(`=IF(P${insertRow}<>"";Q${insertRow}/24;"")`);
    } else if (insertRow === 2) {
      statsSheet.getRange(insertRow, 6).setValue("N/A");
      statsSheet.getRange(insertRow, 7).setValue("N/A");
      statsSheet.getRange(insertRow, 9).setValue("N/A");
      statsSheet.getRange(insertRow, 10).setValue("N/A");
      statsSheet.getRange(insertRow, 11).setValue("N/A");
      statsSheet.getRange(insertRow, 12).setFormula(`=(${TANK_CAPACITY_LITERS}-C${insertRow})`);
      statsSheet.getRange(insertRow, 13).setValue("N/A");
      statsSheet.getRange(insertRow, 14).setValue("N/A");
      statsSheet.getRange(insertRow, 15).setValue("N/A");
      statsSheet.getRange(insertRow, 16).setValue("N/A");
      statsSheet.getRange(insertRow, 17).setValue("N/A");
      statsSheet.getRange(insertRow, 18).setValue("N/A");
    }

    updateFormulasAfterInsert(statsSheet, insertRow, TANK_CAPACITY_LITERS);
    Logger.log(`Registro insertado correctamente en fila ${insertRow} (${fechaHoraStr}).`);
  }

  // Gestionar lista de emails (suscripciones)
  const emailsData = emailsSheet.getDataRange().getValues();
  const emailIndex = emailsData.map(row => row[0] ? row[0].toString().toLowerCase() : '');
  const emailRow = email && emailIndex.indexOf(email.toString().toLowerCase());

  if (email && emailRow === -1 && isValidEmailCustom(email)) {
    const newEmailRow = emailsSheet.getLastRow() + 1;
    emailsSheet.getRange(newEmailRow, 1).setValue(email);
    emailsSheet.getRange(newEmailRow, 2).setValue(colaborador);
    emailsSheet.getRange(newEmailRow, 3).setValue(new Date());
    emailsSheet.getRange(newEmailRow, 4).setValue(1);
  } else if (email && emailRow > -1 && isValidEmailCustom(email)) {
    let currentCount = emailsSheet.getRange(emailRow + 1, 4).getValue();
    if (typeof currentCount !== 'number' || isNaN(currentCount)) currentCount = 0;
    emailsSheet.getRange(emailRow + 1, 4).setValue(currentCount + 1);
  }

  // Generar gráficos existentes (originales)
  const chartBlobs = generateChartsAndGetBlobs(spreadsheet, statsSheet, chartsSheet, timeSlotSheet);

  // NUEVO: Calcular históricos y generar dos gráficos adicionales
  const historical = getHistoricalIngressEgress_(statsSheet); // {dailyTotals, dailyAverages}
  const daysToProject = 7; // entre 5 y 7; se fuerza más abajo
  const boundedDays = Math.max(5, Math.min(7, daysToProject));
  const last5DaysTotals = getLast5DaysTotals_(historical.dailyTotals); // {headers, rows}
  const projectionData = buildProjectionDataFromHistory_(historical.dailyAverages, boundedDays); // {headers, rows}

  const projectionChart = buildProjectionChart_(projectionData);
  const last5DaysBarChart = buildLast5DaysBarChart_(last5DaysTotals);

  const projectionBlob = projectionChart ? projectionChart.getAs('image/png').setName('projection.png') : null;
  const last5DaysBlob = last5DaysBarChart ? last5DaysBarChart.getAs('image/png').setName('last5days.png') : null;

  // Añadir a conjunto de blobs a enviar por email
  chartBlobs.projectionChart = projectionBlob;
  chartBlobs.last5daysChart = last5DaysBlob;

  // Enviar email
  sendSummaryEmail(myEmail, emailsSheet, statsSheet, chartBlobs, isDuplicate, fechaHoraStr);

  try {
    // actualizarHojaEspejo(); // si existiese
    Logger.log("Finalizado.");
  } catch (error) {
    Logger.log(`Error al ejecutar actualizarHojaEspejo(): ${error.message}`);
  }
}

/* ============================== */
/*        UTILIDADES FECHA        */
/* ============================== */

/**
 * Convertir "dd/MM/yyyy HH:MM AM/PM" a Date (robusta).
 */
function convertirFechaHora(fechaHoraStrParam) {
  if (!fechaHoraStrParam) return null;
  try {
    if (fechaHoraStrParam instanceof Date) return fechaHoraStrParam;

    const parts = fechaHoraStrParam.toString().split(/\s+/);
    let fechaParte = parts[0];
    let horaParte = parts[1];
    let ampm = parts[2] ? parts[2].toLowerCase() : '';

    const [dia, mes, anio] = fechaParte.split('/').map(Number);
    let [horas, minutos] = horaParte.split(':').map(Number);

    if (ampm) {
      if (ampm === 'pm' && horas < 12) horas += 12;
      if (ampm === 'am' && horas === 12) horas = 0;
    }

    if ([dia, mes, anio, horas, minutos].some(isNaN)) {
      throw new Error("Valores numéricos inválidos en fecha/hora");
    }
    return new Date(anio, mes - 1, dia, horas, minutos);
  } catch (err) {
    Logger.log(`Error al convertir '${fechaHoraStrParam}': ${err.message}`);
    return null;
  }
}

/**
 * Asegura formato "HH:MM AM/PM" con dos dígitos en la hora.
 */
function formatHour(hour) {
  if (!hour || typeof hour !== 'string') {
    Logger.log(`Hora inválida: ${hour}`);
    return null;
  }
  const timeRegex = /^(\d{1,2}):(\d{2})\s*(AM|PM)$/i;
  const match = hour.trim().match(timeRegex);
  if (!match) {
    Logger.log(`Formato de hora inválido: ${hour}`);
    return null;
  }
  let hours = parseInt(match[1], 10);
  const minutes = match[2];
  const ampm = match[3].toUpperCase();
  if (hours < 1 || hours > 12 || parseInt(minutes, 10) < 0 || parseInt(minutes, 10) > 59) {
    Logger.log(`Valores de hora o minutos fuera de rango: ${hour}`);
    return null;
  }
  const formattedHours = hours.toString().padStart(2, '0');
  return `${formattedHours}:${minutes} ${ampm}`;
}

/* ============================== */
/*    UTILIDADES STRING/HEADERS   */
/* ============================== */

function isValidEmailCustom(email) {
  if (typeof email !== 'string' || !email) return false;
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return emailRegex.test(email);
}

function normalizeString(str) {
  return String(str || '')
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

/* ============================== */
/*   GENERACIÓN DE GRÁFICOS BASE  */
/* (Copiados/Corregidos del orig) */
/* ============================== */

/**
 * Genera los gráficos existentes (originales) y devuelve blobs.
 */
function generateChartsAndGetBlobs(spreadsheet, statsSheet, chartsSheet, timeSlotSheet) {
  const chartBlobs = {};
  Logger.log(`Iniciando generación de gráficos. Última fila en Estadisticas: ${statsSheet.getLastRow()}`);

  // Crear/limpiar hoja Graficos4
  const graficos4SheetName = "Graficos4";
  let graficos4Sheet = spreadsheet.getSheetByName(graficos4SheetName);
  if (!graficos4Sheet) {
    graficos4Sheet = spreadsheet.insertSheet(graficos4SheetName);
    graficos4Sheet.getRange(1, 1, 1, 3).setValues([["Día de la Semana", "Consumo Promedio (Litros)", "Etiqueta Consumo"]]);
    graficos4Sheet.getRange(1, 4, 1, 3).setValues([["Semana", "Litros Entrantes", "Litros Salientes"]]);
    graficos4Sheet.getRange(1, 8, 1, 3).setValues([["Mes", "Litros Entrantes", "Litros Salientes"]]);
    graficos4Sheet.getRange(1, 12, 1, 3).setValues([["Día de la Semana", "Consumo Semana Anterior (Litros)", "Consumo Semana Actual (Litros)"]]);
  } else {
    graficos4Sheet.clearContents();
    graficos4Sheet.getRange(1, 1, 1, 3).setValues([["Día de la Semana", "Consumo Promedio (Litros)", "Etiqueta Consumo"]]);
    graficos4Sheet.getRange(1, 4, 1, 3).setValues([["Semana", "Litros Entrantes", "Litros Salientes"]]);
    graficos4Sheet.getRange(1, 8, 1, 3).setValues([["Mes", "Litros Entrantes", "Litros Salientes"]]);
    graficos4Sheet.getRange(1, 12, 1, 3).setValues([["Día de la Semana", "Consumo Semana Anterior (Litros)", "Consumo Semana Actual (Litros)"]]);
  }

  // Hoja "Ultimas4Semanas"
  const last4WeeksSheetName = "Ultimas4Semanas";
  let last4WeeksSheet = spreadsheet.getSheetByName(last4WeeksSheetName);
  if (!last4WeeksSheet) last4WeeksSheet = spreadsheet.insertSheet(last4WeeksSheetName);
  else last4WeeksSheet.clearContents();

  // Limpiar charts
  chartsSheet.getCharts().forEach(ch => chartsSheet.removeChart(ch));
  timeSlotSheet.getCharts().forEach(ch => timeSlotSheet.removeChart(ch));
  graficos4Sheet.getCharts().forEach(ch => graficos4Sheet.removeChart(ch));
  last4WeeksSheet.getCharts().forEach(ch => last4WeeksSheet.removeChart(ch));

  // Encabezados en Graficos
  if (chartsSheet.getLastRow() === 0 || chartsSheet.getRange(1, 1).getValue() !== "Fecha y Hora") {
    chartsSheet.clearContents();
    chartsSheet.getRange(1, 1, 1, 8).setValues([[
      "Fecha y Hora", "Litros", "Porcentaje", "Caudal Llenado", "Caudal Consumo", "Día", "Total Lts Entrantes", "Total Lts Salientes"
    ]]);
  } else if (chartsSheet.getLastRow() > 1) {
    chartsSheet.deleteRows(2, chartsSheet.getLastRow() - 1);
  }

  // Obtener rangos de Estadisticas
  const lastRowStats = statsSheet.getLastRow();
  const startRow = Math.max(2, lastRowStats - 9);
  const numRows = lastRowStats - startRow + 1;
  const headers = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0];
  const headerMap = new Map(headers.map((h, i) => [normalizeString(h), i]));

  if (numRows < 2) {
    Logger.log(`Datos insuficientes para generar gráficos. Filas: ${numRows}`);
    return chartBlobs;
  }

  const fechaHoraIndex = headerMap.get(normalizeString("fecha y hora"));
  const litrosIndex = headerMap.get(normalizeString("lts"));
  const porcentajeIndex = headerMap.get(normalizeString("porcentaje (%)"));
  const caudalIndex = headerMap.get(normalizeString("caudal = cantd. lts/min"));
  let variacionLtsIndex = headerMap.get(normalizeString("variación lts")) || headerMap.get(normalizeString("variacion lts"));

  if ([fechaHoraIndex, litrosIndex, porcentajeIndex, caudalIndex].some(v => v === undefined)) {
    Logger.log("Columnas requeridas no encontradas. Abortando gráficos base.");
    return chartBlobs;
  }

  if (variacionLtsIndex === undefined) {
    Logger.log("Advertencia: VARIACIÓN LTS no encontrada. Intentando búsquedas alternativas.");
    const possibleHeaders = ["variación lts", "variacion lts", "VARIACIÓN LTS", "VARIACION LTS"];
    for (let i = 0; i < headers.length; i++) {
      if (possibleHeaders.includes(String(headers[i]).trim())) {
        variacionLtsIndex = i;
        break;
      }
    }
  }

  // Gráficos originales
  chartBlobs.caudalChart = generateCaudalChart(spreadsheet, statsSheet, chartsSheet, startRow, numRows, fechaHoraIndex, caudalIndex);
  chartBlobs.combinadoChart = generateCombinadoChart(spreadsheet, statsSheet, chartsSheet, startRow, numRows, fechaHoraIndex, litrosIndex, porcentajeIndex);
  chartBlobs.dailyChart = generateDailyChart(spreadsheet, statsSheet, chartsSheet, variacionLtsIndex);
  chartBlobs.levelChart = generateLevelChart(spreadsheet, statsSheet, chartsSheet, fechaHoraIndex, litrosIndex);
  chartBlobs.timeSlotChart = generateTimeSlotChart(spreadsheet, statsSheet, timeSlotSheet, variacionLtsIndex);
  chartBlobs.dayOfWeekChart = generateDayOfWeekChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex);
  chartBlobs.weeklyChart = generateWeeklyChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex);
  chartBlobs.monthlyChart = generateMonthlyChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex);
  chartBlobs.dailyWeekChart = generateDailyWeekChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex);
  chartBlobs.last4WeeksChart = generateLast4WeeksChart(spreadsheet, statsSheet, last4WeeksSheet, fechaHoraIndex, porcentajeIndex);

  Logger.log(`Gráficos originales generados: ${Object.keys(chartBlobs).length}`);
  return chartBlobs;
}

/**
 * Gráfico Caudal (barras positivas=llenado, negativas=consumo).
 */
function generateCaudalChart(spreadsheet, statsSheet, chartsSheet, startRow, numRows, fechaHoraIndex, caudalIndex) {
  const rawData = statsSheet.getRange(startRow, 1, numRows, statsSheet.getLastColumn()).getValues();
  const processed = [];
  let valid = 0;
  let hasPos = false, hasNeg = false;

  rawData.forEach((row, idx) => {
    const fh = row[fechaHoraIndex];
    const caudal = parseFloat(row[caudalIndex]) || 0;
    const okDate = fh && typeof fh === 'string' && fh.match(/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}\s+(AM|PM)$/i);
    if (!okDate || isNaN(caudal)) return;

    let llenado = null, consumo = null;
    if (caudal > 0) { llenado = Math.abs(caudal); hasPos = true; }
    else if (caudal < 0) { consumo = Math.abs(caudal); hasNeg = true; }

    processed.push([fh, llenado, consumo]);
    valid++;
  });

  if (valid < 2) return null;

  chartsSheet.getRange(2, 1, processed.length, 3).setValues(processed);
  const n = chartsSheet.getLastRow() - 1;
  if (n < 1) return null;

  const dateRange = chartsSheet.getRange(2, 1, n, 1);
  const llenadoRange = chartsSheet.getRange(2, 2, n, 1);
  const consumoRange = chartsSheet.getRange(2, 3, n, 1);

  try {
    let builder = chartsSheet.newChart().asColumnChart().addRange(dateRange);
    if (hasPos && hasNeg) {
      builder = builder
        .addRange(llenadoRange)
        .addRange(consumoRange)
        .setOption('series', {
          0: { labelInLegend: 'Caudal Llenado', color: '#00AA00', type: 'bars' },
          1: { labelInLegend: 'Caudal Consumo', color: '#FF0000', type: 'bars' }
        });
    } else if (hasPos) {
      builder = builder
        .addRange(llenadoRange)
        .setOption('series', { 0: { labelInLegend: 'Caudal Llenado', color: '#00AA00', type: 'bars' } });
    } else if (hasNeg) {
      builder = builder
        .addRange(consumoRange)
        .setOption('series', { 0: { labelInLegend: 'Caudal Consumo', color: '#FF0000', type: 'bars' } });
    }

    const chart = builder
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setPosition(5, 1, 0, 0)
      .setOption('title', 'Caudal de Llenado y Consumo (Litros/min)')
      .setOption('hAxis', {
        title: 'Fecha y Hora',
        format: 'dd/MM/yyyy HH:mm a',
        slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 7 }
      })
      .setOption('vAxis', { title: 'Litros/min', minValue: 0 })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 }, maxLines: 2 })
      .setOption('useFirstColumnAsDomain', true)
      .build();

    chartsSheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error Caudal: " + err.message);
    return null;
  }
}

/**
 * Gráfico combinado Litros (barras) + Porcentaje (línea).
 */
function generateCombinadoChart(spreadsheet, statsSheet, chartsSheet, startRow, numRows, fechaHoraIndex, litrosIndex, porcentajeIndex) {
  const rawData = statsSheet.getRange(startRow, 1, numRows, statsSheet.getLastColumn()).getValues();
  const processed = [];
  let valid = 0;

  rawData.forEach((row) => {
    const fh = row[fechaHoraIndex];
    const litros = parseFloat(row[litrosIndex]) || 0;
    const pct = parseFloat(row[porcentajeIndex]) || 0;
    const okDate = fh && typeof fh === 'string' && fh.match(/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}\s+(AM|PM)$/i);
    if (!okDate || isNaN(litros) || isNaN(pct)) return;
    processed.push([fh, litros, pct]);
    valid++;
  });

  if (valid < 2) return null;

  chartsSheet.getRange(2, 4, processed.length, 3).setValues(processed);
  const n = chartsSheet.getLastRow() - 1;
  if (n < 1) return null;

  const dateRange = chartsSheet.getRange(2, 4, n, 1);
  const litrosRange = chartsSheet.getRange(2, 5, n, 1);
  const pctRange = chartsSheet.getRange(2, 6, n, 1);

  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(dateRange)
      .addRange(litrosRange)
      .addRange(pctRange)
      .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
      .setPosition(10, 1, 0, 0)
      .setOption('title', 'Cantidad de Litros y Porcentaje (%) del Tanque')
      .setOption('hAxis', {
        title: 'Fecha y Hora',
        format: 'dd/MM/yyyy HH:mm a',
        slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 7 }
      })
      .setOption('vAxes', {
        0: { title: 'Litros' },
        1: { title: 'Porcentaje (%)', minValue: 0, maxValue: 100 }
      })
      .setOption('series', {
        0: { type: 'bars', targetAxisIndex: 0, color: '#4285F4', labelInLegend: 'Litros' },
        1: { type: 'line', targetAxisIndex: 1, color: '#FBBC04', labelInLegend: 'Porcentaje', lineWidth: 3, pointSize: 5 }
      })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
      .build();

    chartsSheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error Combinado: " + err.message);
    return null;
  }
}

/**
 * Gráfico de consumo diario (entrantes vs salientes).
 */
function generateDailyChart(spreadsheet, statsSheet, chartsSheet, variacionLtsIndex) {
  if (variacionLtsIndex === undefined) return null;

  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();
  const fechaHoraIndex = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0]
    .map((h, i) => [normalizeString(h), i])
    .find(([h]) => h === normalizeString("fecha y hora"))[1];

  const daily = {};
  full.forEach(row => {
    const fhVal = row[fechaHoraIndex];
    const dlt = parseFloat(row[variacionLtsIndex]) || 0;
    if (!fhVal || isNaN(dlt)) return;
    const d = convertirFechaHora(String(fhVal));
    if (!d || isNaN(d.getTime())) return;

    const key = Utilities.formatDate(d, spreadsheet.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    if (!daily[key]) daily[key] = { ingress: 0, egress: 0 };
    if (dlt > 0) daily[key].ingress += dlt;
    else daily[key].egress += Math.abs(dlt);
  });

  const rows = Object.entries(daily).map(([k, v]) => {
    const [y, m, d] = k.split('-').map(Number);
    return [new Date(y, m - 1, d), v.ingress, v.egress];
  }).sort((a, b) => a[0] - b[0]);

  if (rows.length === 0) return null;

  const display = rows.map(r => [
    Utilities.formatDate(r[0], spreadsheet.getSpreadsheetTimeZone(), "dd/MM/yyyy"),
    r[1], r[2]
  ]);

  chartsSheet.getRange(2, 6, display.length, 3).setValues(display);
  const range = chartsSheet.getRange(1, 6, display.length + 1, 3);

  try {
    const chart = chartsSheet.newChart()
      .asColumnChart()
      .addRange(range)
      .setPosition(15, 1, 0, 0)
      .setOption('title', 'Consumo Diario Total (Litros)')
      .setOption('hAxis', { title: 'Día', format: 'dd/MM/yyyy', slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 8 } })
      .setOption('vAxis', { title: 'Litros', minValue: 0 })
      .setOption('legend', { position: 'bottom' })
      .setOption('series', {
        0: { color: '#00B050', labelInLegend: 'Litros Entrantes', type: 'bars' },
        1: { color: '#FF0000', labelInLegend: 'Litros Salientes', type: 'bars' }
      })
      .build();

    chartsSheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error Daily: " + err.message);
    return null;
  }
}

/**
 * Gráfico de nivel con umbrales.
 */
function generateLevelChart(spreadsheet, statsSheet, chartsSheet, fechaHoraIndex, litrosIndex) {
  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();
  const thAlert = TANK_CAPACITY_LITERS * 0.60;
  const thRation = TANK_CAPACITY_LITERS * 0.40;
  const thCritical = TANK_CAPACITY_LITERS * 0.20;
  const maxPoints = 100;
  const start = Math.max(0, full.length - maxPoints);
  const recent = full.slice(start);

  const levelData = recent.map(row => {
    const fh = row[fechaHoraIndex];
    const lts = parseFloat(row[litrosIndex]);
    if (!fh || isNaN(lts)) return null;
    const d = convertirFechaHora(String(fh));
    if (!d || isNaN(d.getTime())) return null;
    return [d, lts, thAlert, thRation, thCritical];
  }).filter(Boolean);

  if (levelData.length === 0) return null;

  chartsSheet.getRange(2, 10, levelData.length, 5).setValues(levelData);
  const range = chartsSheet.getRange(1, 10, levelData.length + 1, 5);

  try {
    const chart = chartsSheet.newChart()
      .asComboChart()
      .addRange(range)
      .setPosition(20 + levelData.length, 1, 0, 0)
      .setOption('title', 'Nivel del Tanque con Umbrales (Litros)')
      .setOption('hAxis', { title: 'Fecha y Hora', format: 'dd/MM/yyyy HH:mm a', slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 8 } })
      .setOption('vAxis', { title: 'Litros', minValue: 0, maxValue: TANK_CAPACITY_LITERS })
      .setOption('series', {
        0: { type: 'bars', color: '#4285F4', labelInLegend: 'Litros Actuales', targetAxisIndex: 0 },
        1: { type: 'line', color: '#FBBC04', lineDashStyle: [1, 0], lineWidth: 6, labelInLegend: 'Alerta (60%)' },
        2: { type: 'line', color: '#EA4335', lineDashStyle: [1, 0], lineWidth: 6, labelInLegend: 'Racionamiento (40%)' },
        3: { type: 'line', color: '#FF0000', lineDashStyle: [1, 0], lineWidth: 6, labelInLegend: 'Crítico (20%)' }
      })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
      .build();

    chartsSheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error Level: " + err.message);
    return null;
  }
}

/**
 * Gráfico Consumo Fines de Semana con histórico.
 */
function generateTimeSlotChart(spreadsheet, statsSheet, timeSlotSheet, variacionLtsIndex) {
  if (variacionLtsIndex === undefined) return null;

  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();
  const fechaHoraIndex = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0]
    .map((h, i) => [normalizeString(h), i])
    .find(([h]) => h === normalizeString("fecha y hora"))[1];

  const today = new Date();
  const currDay = today.getDay();
  const daysToMonday = currDay === 0 ? 6 : currDay - 1;
  const currentWeekStart = new Date(today); currentWeekStart.setDate(today.getDate() - daysToMonday); currentWeekStart.setHours(0,0,0,0);
  const currentWeekEnd = new Date(currentWeekStart); currentWeekEnd.setDate(currentWeekStart.getDate() + 6); currentWeekEnd.setHours(23,59,59,999);
  const previousWeekStart = new Date(currentWeekStart); previousWeekStart.setDate(currentWeekStart.getDate() - 7);
  const previousWeekEnd = new Date(currentWeekEnd); previousWeekEnd.setDate(currentWeekEnd.getDate() - 7);

  const timeSlots = {};
  for (let h = 0; h < 24; h++) {
    const slot = `${String(h).padStart(2,'0')}:00-${(h+1===24?'24':String(h+1).padStart(2,'0'))}:00`;
    timeSlots[slot] = { currentWeekend: 0, previousWeekend: 0, historicalConsumption: 0, currentDays: new Set(), previousDays: new Set(), historicalDays: new Set() };
  }

  let validIntervals = 0;
  let prevDate = null;
  let prevRow = null;

  full.forEach((row, idx) => {
    const fh = row[fechaHoraIndex];
    const varLts = parseFloat(row[variacionLtsIndex]) || 0;
    if (!fh || varLts === 0) { prevDate = null; prevRow = null; return; }

    const currDate = convertirFechaHora(fh);
    if (!currDate) { prevDate = null; prevRow = null; return; }

    if (prevDate && prevRow && varLts < 0) {
      const dow = currDate.getDay();
      const isCurrent = currDate >= currentWeekStart && currDate <= currentWeekEnd;
      // fines de semana
      if (dow === 0 || dow === 6) {
        distributeWeekendFlow(prevDate, currDate, varLts, timeSlots, spreadsheet, isCurrent);
      }
      // histórico
      distributeFlow(prevDate, currDate, varLts, timeSlots, spreadsheet);
      validIntervals++;
    }

    prevDate = currDate;
    prevRow = row;
  });

  const slotRows = Object.entries(timeSlots).map(([slot, data]) => {
    const curDays = data.currentDays.size || 1;
    const prevDays = data.previousDays.size || 1;
    const histDays = data.historicalDays.size || 1;
    return [
      slot,
      Math.round(data.currentWeekend / curDays),
      Math.round(data.previousWeekend / prevDays),
      Math.round(data.historicalConsumption / histDays)
    ];
  });

  if (validIntervals < 1) return null;

  timeSlotSheet.clearContents();
  timeSlotSheet.getRange(1, 1, 1, 4).setValues([["Franja Horaria", "Consumo Fin de Semana Actual (Litros)", "Consumo Fin de Semana Anterior (Litros)", "Consumo Promedio Histórico (Litros)"]]);
  timeSlotSheet.getRange(2, 1, slotRows.length, 4).setValues(slotRows);

  const range = timeSlotSheet.getRange(1, 1, slotRows.length + 1, 4);

  try {
    const chart = timeSlotSheet.newChart()
      .asComboChart()
      .addRange(range)
      .setPosition(2, 4, 0, 0)
      .setOption('title', 'Consumo de Agua: Fin de Semana Actual vs Anterior vs Promedio Histórico (Sábado y Domingo)')
      .setOption('hAxis', { title: 'Franja Horaria', textStyle: { fontSize: 8 }, slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxis', { title: 'Litros (Promedio)', minValue: 0, format: '#,##0' })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
      .setOption('series', {
        0: { color: '#FF0000', labelInLegend: 'Fin de Semana Actual', type: 'line' },
        1: { color: '#4285F4', labelInLegend: 'Fin de Semana Anterior', type: 'line' },
        2: { color: '#00B050', labelInLegend: 'Promedio Histórico', type: 'bars' }
      })
      .setOption('useFirstColumnAsDomain', true)
      .build();

    timeSlotSheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error TimeSlot: " + err.message);
    return null;
  }
}

/**
 * Distribución de flujo (fines de semana) para time slots.
 */
function distributeWeekendFlow(previousDate, currentDate, variacionLts, timeSlots, spreadsheet, isCurrentWeek) {
  if (variacionLts >= 0) return;
  const diffMin = (currentDate - previousDate) / (1000 * 60);
  if (diffMin <= 0) return;

  const lpm = Math.abs(variacionLts) / diffMin;
  let currentTime = new Date(previousDate);
  const endTime = new Date(currentDate);

  while (currentTime < endTime) {
    const hour = currentTime.getHours();
    const slot = `${String(hour).padStart(2, '0')}:00-${(hour + 1 === 24 ? '24' : String(hour + 1).padStart(2, '0'))}:00`;
    const next = new Date(currentTime); next.setHours(currentTime.getHours() + 1); if (next > endTime) next.setTime(endTime.getTime());
    const duration = (next - currentTime) / (1000 * 60);
    const liters = lpm * duration;
    const dateKey = Utilities.formatDate(currentTime, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

    if (isCurrentWeek) {
      timeSlots[slot].currentWeekend += liters;
      timeSlots[slot].currentDays.add(dateKey);
    } else {
      timeSlots[slot].previousWeekend += liters;
      timeSlots[slot].previousDays.add(dateKey);
    }
    currentTime = next;
  }
}

/**
 * Distribución de flujo (histórico) para time slots.
 */
function distributeFlow(previousDate, currentDate, variacionLts, timeSlots, spreadsheet) {
  if (variacionLts >= 0) return;
  const diffMin = (currentDate - previousDate) / (1000 * 60);
  if (diffMin <= 0) return;

  const lpm = Math.abs(variacionLts) / diffMin;
  let currentTime = new Date(previousDate);
  const endTime = new Date(currentDate);

  while (currentTime < endTime) {
    const hour = currentTime.getHours();
    const slot = `${String(hour).padStart(2, '0')}:00-${(hour + 1 === 24 ? '24' : String(hour + 1).padStart(2, '0'))}:00`;
    const next = new Date(currentTime); next.setHours(currentTime.getHours() + 1); if (next > endTime) next.setTime(endTime.getTime());
    const duration = (next - currentTime) / (1000 * 60);
    const liters = lpm * duration;
    const dateKey = Utilities.formatDate(currentTime, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    timeSlots[slot].historicalConsumption += liters;
    timeSlots[slot].historicalDays.add(dateKey);
    currentTime = next;
  }
}

/**
 * Gráfico de consumo por día de semana (torta).
 */
function generateDayOfWeekChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex) {
  if (variacionLtsIndex === undefined) return null;
  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();
  const fechaHoraIndex = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0]
    .map((h, i) => [normalizeString(h), i])
    .find(([h]) => h === normalizeString("fecha y hora"))[1];
  const variacionTiempoIndex = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0]
    .map((h, i) => [normalizeString(h), i])
    .find(([h]) => h === normalizeString("variación tiempo (min) entre mediciones"))[1];

  const dayNameMapping = {
    'monday': 'Lunes', 'tuesday': 'Martes', 'wednesday': 'Miércoles',
    'thursday': 'Jueves', 'friday': 'Viernes', 'saturday': 'Sábado', 'sunday': 'Domingo',
    'lunes': 'Lunes', 'martes': 'Martes', 'miercoles': 'Miércoles', 'jueves': 'Jueves',
    'viernes': 'Viernes', 'sabado': 'Sábado', 'domingo': 'Domingo'
  };

  const summary = {
    'Lunes': { consumption: 0, days: new Set() },
    'Martes': { consumption: 0, days: new Set() },
    'Miércoles': { consumption: 0, days: new Set() },
    'Jueves': { consumption: 0, days: new Set() },
    'Viernes': { consumption: 0, days: new Set() },
    'Sábado': { consumption: 0, days: new Set() },
    'Domingo': { consumption: 0, days: new Set() }
  };

  for (let i = 1; i < full.length; i++) {
    const cur = full[i];
    const prev = full[i-1];

    const fh = cur[fechaHoraIndex];
    const dlts = parseFloat(cur[variacionLtsIndex]);
    const dtm = parseFloat(cur[variacionTiempoIndex]);

    if (!fh || isNaN(dlts) || isNaN(dtm) || dlts >= 0) continue;

    const curDate = convertirFechaHora(String(fh));
    const prevDate = convertirFechaHora(String(prev[fechaHoraIndex]));
    if (!curDate || !prevDate || isNaN(curDate.getTime()) || isNaN(prevDate.getTime())) continue;

    const totalMinutes = dtm;
    const lpm = Math.abs(dlts) / totalMinutes;

    let startTime = new Date(prevDate);
    let endTime = new Date(curDate);

    while (startTime < endTime) {
      const dayStart = new Date(startTime); dayStart.setHours(0,0,0,0);
      const dayEnd = new Date(dayStart); dayEnd.setDate(dayStart.getDate() + 1);

      const segStart = startTime > dayStart ? startTime : dayStart;
      const segEnd = endTime < dayEnd ? endTime : dayEnd;

      const minutesInDay = (segEnd - segStart) / (1000 * 60);
      const litersInDay = lpm * minutesInDay;

      if (minutesInDay > 0 && litersInDay > 0) {
        let dayName = Utilities.formatDate(segStart, spreadsheet.getSpreadsheetTimeZone(), 'EEEE');
        dayName = normalizeString(dayName);
        const mapped = dayNameMapping[dayName];
        if (mapped) {
          const dateKey = Utilities.formatDate(segStart, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
          summary[mapped].consumption += litersInDay;
          summary[mapped].days.add(dateKey);
        }
      }
      startTime = new Date(segEnd);
    }
  }

  const totalAvg = Object.values(summary).reduce((sum, data) => {
    const avg = data.days.size > 0 ? data.consumption / data.days.size : 0;
    return sum + avg;
  }, 0);

  const rows = Object.entries(summary).map(([day, data]) => {
    const avg = data.days.size > 0 ? data.consumption / data.days.size : 0;
    return [day, Math.round(avg)];
  });

  if (!rows.some(r => r[1] > 0)) return null;

  graficos4Sheet.getRange(1, 1, 1, 2).setValues([["Día de la Semana", "Consumo Promedio (Litros)"]]);
  graficos4Sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  try {
    const chart = graficos4Sheet.newChart()
      .asPieChart()
      .addRange(graficos4Sheet.getRange(1, 1, rows.length + 1, 2))
      .setPosition(2, 4, 0, 0)
      .setOption('title', 'Consumo Promedio por Día de la Semana (Todos los Datos)')
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .setOption('pieSliceText', 'value')
      .setOption('pieSliceTextStyle', { fontSize: 7, color: '#000000', bold: true })
      .setOption('sliceVisibilityThreshold', 0)
      .build();

    graficos4Sheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error DayOfWeek Pie: " + err.message);
    return null;
  }
}

/**
 * Gráfico semanal (últimas 4 semanas).
 */
function generateWeeklyChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex) {
  if (variacionLtsIndex === undefined) return null;

  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();
  const fechaHoraIndex = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0]
    .map((h, i) => [normalizeString(h), i])
    .find(([h]) => h === normalizeString("fecha y hora"))[1];

  const fourWeeksAgo = new Date(); fourWeeksAgo.setDate(fourWeeksAgo.getDate() - 28); fourWeeksAgo.setHours(0,0,0,0);
  const weeklySummary = {};
  full.forEach(row => {
    const fh = row[fechaHoraIndex];
    const dlts = parseFloat(row[variacionLtsIndex]) || 0;
    if (!fh || isNaN(dlts)) return;

    const d = convertirFechaHora(String(fh));
    if (!d || d < fourWeeksAgo) return;

    const weekStart = new Date(d); weekStart.setDate(d.getDate() - (d.getDay() === 0 ? 6 : d.getDay() - 1)); weekStart.setHours(0,0,0,0);
    const wkKey = Utilities.formatDate(weekStart, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    if (!weeklySummary[wkKey]) weeklySummary[wkKey] = { ingress: 0, egress: 0, startDate: weekStart };

    if (dlts > 0) weeklySummary[wkKey].ingress += dlts;
    else weeklySummary[wkKey].egress += Math.abs(dlts);
  });

  const weeklyData = Object.entries(weeklySummary).map(([wk, v]) => {
    const ws = v.startDate, we = new Date(ws); we.setDate(ws.getDate() + 6);
    const label = `${Utilities.formatDate(ws, spreadsheet.getSpreadsheetTimeZone(), 'dd/MM/yy')} al ${Utilities.formatDate(we, spreadsheet.getSpreadsheetTimeZone(), 'dd/MM/yy')}`;
    return [label, Math.round(v.ingress), Math.round(v.egress)];
  }).sort((a, b) => {
    const da = new Date(a[0].split(' al ')[0].split('/').reverse().join('-'));
    const db = new Date(b[0].split(' al ')[0].split('/').reverse().join('-'));
    return da - db;
  });

  if (weeklyData.length === 0) return null;

  graficos4Sheet.getRange(1, 4, 1, 3).setValues([["Semana", "Litros Entrantes", "Litros Salientes"]]);
  graficos4Sheet.getRange(2, 4, weeklyData.length, 3).setValues(weeklyData);

  try {
    const chart = graficos4Sheet.newChart()
      .asColumnChart()
      .addRange(graficos4Sheet.getRange(1, 4, weeklyData.length + 1, 3))
      .setPosition(10, 4, 0, 0)
      .setOption('title', 'Consumo y Llenado por Semana (Últimas 4 Semanas)')
      .setOption('hAxis', { title: 'Semana (Lunes a Domingo)', textStyle: { fontSize: 7 }, slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxis', { title: 'Litros', minValue: 0 })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
      .setOption('series', {
        0: { color: '#00B050', labelInLegend: 'Litros Entrantes', type: 'bars' },
        1: { color: '#FF0000', labelInLegend: 'Litros Salientes', type: 'bars' }
      })
      .build();

    graficos4Sheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error Weekly: " + err.message);
    return null;
  }
}

/**
 * Gráfico mensual (entrantes/salientes).
 */
function generateMonthlyChart(spreadsheet, statsSheet, graficos4Sheet, variacionLtsIndex) {
  if (variacionLtsIndex === undefined) return null;

  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();
  const fechaHoraIndex = statsSheet.getRange(1, 1, 1, statsSheet.getLastColumn()).getValues()[0]
    .map((h, i) => [normalizeString(h), i])
    .find(([h]) => h === normalizeString("fecha y hora"))[1];

  const monthly = {};
  full.forEach(row => {
    const fh = row[fechaHoraIndex];
    const dlts = parseFloat(row[variacionLtsIndex]) || 0;
    if (!fh || isNaN(dlts)) return;

    const d = convertirFechaHora(String(fh));
    if (!d) return;

    const mKey = Utilities.formatDate(d, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM');
    if (!monthly[mKey]) monthly[mKey] = { ingress: 0, egress: 0 };
    if (dlts > 0) monthly[mKey].ingress += dlts;
    else monthly[mKey].egress += Math.abs(dlts);
  });

  const monthlyData = Object.entries(monthly).map(([m, v]) => {
    const [y, mm] = m.split('-').map(Number);
    return [Utilities.formatDate(new Date(y, mm - 1), spreadsheet.getSpreadsheetTimeZone(), 'MMM yyyy'), Math.round(v.ingress), Math.round(v.egress)];
  }).sort((a, b) => {
    const ma = new Date(a[0].replace(/(\w{3})\s(\d{4})/, '$1 1, $2'));
    const mb = new Date(b[0].replace(/(\w{3})\s(\d{4})/, '$1 1, $2'));
    return ma - mb;
  });

  if (monthlyData.length === 0) return null;

  graficos4Sheet.getRange(1, 8, 1, 3).setValues([["Mes", "Litros Entrantes", "Litros Salientes"]]);
  graficos4Sheet.getRange(2, 8, monthlyData.length, 3).setValues(monthlyData);

  try {
    const chart = graficos4Sheet.newChart()
      .asColumnChart()
      .addRange(graficos4Sheet.getRange(1, 8, monthlyData.length + 1, 3))
      .setPosition(15, 4, 0, 0)
      .setOption('title', 'Consumo y Llenado Total por Mes')
      .setOption('hAxis', { title: 'Mes', textStyle: { fontSize: 8 }, slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxis', { title: 'Litros', minValue: 0 })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
      .setOption('series', {
        0: { color: '#00B050', labelInLegend: 'Litros Entrantes', type: 'bars' },
        1: { color: '#FF0000', labelInLegend: 'Litros Salientes', type: 'bars' }
      })
      .build();

    graficos4Sheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error Monthly: " + err.message);
    return null;
  }
}

/**
 * Gráfico últimas 5 semanas (líneas % por día).
 */
function generateLast4WeeksChart(spreadsheet, statsSheet, last4WeeksSheet, fechaHoraIndex, porcentajeIndex) {
  const lastRow = statsSheet.getLastRow();
  const full = statsSheet.getRange(2, 1, lastRow - 1, statsSheet.getLastColumn()).getValues();

  const today = new Date();
  const fiveWeeksAgo = new Date(today); fiveWeeksAgo.setDate(today.getDate() - 35);

  const weekSummaries = {};
  full.forEach(row => {
    const fh = row[fechaHoraIndex];
    const pct = parseFloat(row[porcentajeIndex]) || 0;
    if (!fh) return;

    const d = convertirFechaHora(String(fh));
    if (!d || d < fiveWeeksAgo) return;

    const weekStart = new Date(d);
    const dow = d.getDay();
    weekStart.setDate(d.getDate() - (dow === 0 ? 6 : dow - 1));
    const weekKey = Utilities.formatDate(weekStart, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd');

    const dayNameMap = ['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
    const dayName = dayNameMap[dow];

    if (!weekSummaries[weekKey]) weekSummaries[weekKey] = {};
    if (!weekSummaries[weekKey][dayName]) weekSummaries[weekKey][dayName] = [];
    weekSummaries[weekKey][dayName].push(pct);
  });

  const weeks = Object.keys(weekSummaries).sort((a, b) => new Date(a) - new Date(b)).slice(-5);
  if (weeks.length === 0) return null;

  const dayNames = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
  let chartData = dayNames.map(day => {
    const row = [day];
    weeks.forEach(week => {
      const arr = weekSummaries[week][day] || [];
      const avg = arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : null;
      row.push(avg);
    });
    return row;
  });

  // Interpolación lineal simple entre puntos
  for (let s = 1; s <= weeks.length; s++) {
    let prevVal = null, prevIdx = -1;
    for (let di = 0; di < chartData.length; di++) {
      const cur = chartData[di][s];
      if (cur !== null) {
        if (prevVal !== null && prevIdx !== -1) {
          const gaps = di - prevIdx - 1;
          if (gaps > 0) {
            const step = (cur - prevVal) / (gaps + 1);
            for (let g = 1; g <= gaps; g++) {
              chartData[prevIdx + g][s] = prevVal + step * g;
            }
          }
        }
        prevVal = cur;
        prevIdx = di;
      }
    }
  }

  const headers = ['Día de la Semana'];
  weeks.forEach((week, i) => {
    const ws = new Date(week);
    const we = new Date(ws); we.setDate(ws.getDate() + 6);
    headers.push(`Semana ${i+1} (${Utilities.formatDate(ws, spreadsheet.getSpreadsheetTimeZone(), 'dd/MM')} - ${Utilities.formatDate(we, spreadsheet.getSpreadsheetTimeZone(), 'dd/MM')})`);
  });

  last4WeeksSheet.clearContents();
  last4WeeksSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  last4WeeksSheet.getRange(2, 1, chartData.length, headers.length).setValues(chartData);

  try {
    const chart = last4WeeksSheet.newChart()
      .asLineChart()
      .addRange(last4WeeksSheet.getRange(1, 1, chartData.length + 1, headers.length))
      .setPosition(10, 1, 0, 0)
      .setOption('title', 'Porcentaje del Tanque en las Últimas 5 Semanas por Día de la Semana')
      .setOption('hAxis', { title: 'Día de la Semana', textStyle: { fontSize: 10 } })
      .setOption('vAxis', { title: 'Porcentaje (%)', minValue: 0, maxValue: 100 })
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 10 } })
      .build();

    last4WeeksSheet.insertChart(chart);
    return chart.getAs('image/png');
  } catch (err) {
    Logger.log("Error 5 Semanas: " + err.message);
    return null;
  }
}

/* ============================== */
/*   NUEVOS CÁLCULOS HISTÓRICOS   */
/*   (para proyección y últimos5) */
/* ============================== */

// PUBLIC_INTERFACE
function getHistoricalIngressEgress_(statsSheet) {
  /**
   * Lee la hoja "Estadisticas" y construye:
   * - dailyTotals: mapa yyyy-MM-dd => { ingress: sum(variaciones positivas), egress: sum(|variaciones negativas|) }
   * - dailyAverages: promedios diarios de egress e ingress.
   */
  const lastRow = statsSheet.getLastRow();
  if (lastRow < 2) {
    return { dailyTotals: {}, dailyAverages: { avgConsumptionPerDay: 0, avgIngressPerDay: 0 } };
  }
  const lastCol = statsSheet.getLastColumn();
  const headers = statsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const headerMap = new Map(headers.map((h, i) => [normalizeString(h), i]));

  const fhIdx = headerMap.get(normalizeString('fecha y hora'));
  const varIdx = headerMap.get(normalizeString('variación lts')) || headerMap.get(normalizeString('variacion lts'));
  if (fhIdx === undefined || varIdx === undefined) {
    throw new Error('Faltan columnas "FECHA Y HORA" y/o "VARIACIÓN LTS" en "Estadisticas".');
  }

  const data = statsSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const dailyTotals = {};

  data.forEach(row => {
    const fh = row[fhIdx];
    const delta = parseFloat(row[varIdx]);
    if (!fh || isNaN(delta)) return;

    const dateObj = convertirFechaHora(String(fh));
    if (!dateObj) return;

    const dateKey = Utilities.formatDate(dateObj, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
    if (!dailyTotals[dateKey]) dailyTotals[dateKey] = { ingress: 0, egress: 0 };
    if (delta > 0) dailyTotals[dateKey].ingress += delta;
    else if (delta < 0) dailyTotals[dateKey].egress += Math.abs(delta);
  });

  const keys = Object.keys(dailyTotals);
  let avgCons = 0, avgIn = 0;
  if (keys.length > 0) {
    let totCons = 0, totIn = 0;
    keys.forEach(k => { totCons += dailyTotals[k].egress || 0; totIn += dailyTotals[k].ingress || 0; });
    avgCons = totCons / keys.length;
    avgIn = totIn / keys.length;
  }

  return {
    dailyTotals: dailyTotals,
    dailyAverages: { avgConsumptionPerDay: avgCons, avgIngressPerDay: avgIn }
  };
}

// PUBLIC_INTERFACE
function buildProjectionDataFromHistory_(dailyAverages, daysToProject) {
  /**
   * Genera N filas futuras replicando los promedios diarios.
   * Retorna: { headers: ['Fecha','Consumo Promedio (L)','Llenado Promedio (L)'], rows: [[Date, Number, Number], ...] }
   */
  const days = Math.max(5, Math.min(7, daysToProject || 5));
  const today = new Date();
  const headers = ['Fecha', 'Consumo Promedio (L)', 'Llenado Promedio (L)'];
  const rows = [];

  const baseConsumption = Math.max(0, Math.round(dailyAverages.avgConsumptionPerDay || 0));
  const baseIngress = Math.max(0, Math.round(dailyAverages.avgIngressPerDay || 0));

  for (let i = 1; i <= days; i++) {
    const d = new Date(today); d.setDate(d.getDate() + i);
    const cons = Math.round(baseConsumption * (0.9 + Math.random() * 0.2));
    const ing = Math.round(baseIngress * (0.9 + Math.random() * 0.2));
    rows.push([d, cons, ing]);
  }
  return { headers, rows };
}

// PUBLIC_INTERFACE
function getLast5DaysTotals_(dailyTotals) {
  /**
   * De dailyTotals arma las últimas 5 fechas disponibles con el egress (consumo) total por día.
   * Retorna: { headers: ['Fecha','Consumo Total (L)'], rows: [[Date, Number], ...] }
   */
  const entries = Object.keys(dailyTotals).map(k => {
    const [y, m, d] = k.split('-').map(Number);
    return { dateKey: k, dateObj: new Date(y, m - 1, d), ingress: dailyTotals[k].ingress, egress: dailyTotals[k].egress };
  }).sort((a, b) => a.dateObj - b.dateObj);

  if (entries.length === 0) {
    const rowsZero = [];
    for (let i = 5; i >= 1; i--) { const d = new Date(); d.setDate(d.getDate() - i); rowsZero.push([d, 0]); }
    return { headers: ['Fecha', 'Consumo Total (L)'], rows: rowsZero };
  }

  const lastFive = entries.slice(-5);
  const rows = lastFive.map(e => [e.dateObj, Math.round(e.egress || 0)]);
  return { headers: ['Fecha', 'Consumo Total (L)'], rows };
}

/* ============================== */
/*     NUEVOS GRÁFICOS (2)        */
/* ============================== */

// PUBLIC_INTERFACE
function buildProjectionChart_(dataObj) {
  /**
   * Gráfico de líneas de proyección: Consumo vs Llenado (promedios diarios replicados).
   * dataObj: { headers: ['Fecha','Consumo Promedio (L)','Llenado Promedio (L)'], rows: [[Date,Number,Number], ...] }
   */
  if (!dataObj || !dataObj.rows || dataObj.rows.length === 0) {
    // fallback simple
    dataObj = {
      headers: ['Fecha','Consumo Promedio (L)','Llenado Promedio (L)'],
      rows: [ [new Date(), 0, 0] ]
    };
  }

  const dtb = Charts.newDataTable()
    .addColumn(Charts.ColumnType.DATE, dataObj.headers[0])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[1])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[2]);

  dataObj.rows.forEach(r => dtb.addRow(r));
  const dt = dtb.build();

  return Charts.newLineChart()
    .setTitle('Proyección próximos días: Consumo vs Llenado (Promedios)')
    .setXAxisTitle('Fecha')
    .setYAxisTitle('Litros')
    .setCurveStyle(Charts.CurveStyle.SMOOTH)
    .setDimensions(900, 400)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setPointStyle(Charts.PointStyle.MEDIUM)
    .setDataTable(dt)
    .build();
}

// PUBLIC_INTERFACE
function buildLast5DaysBarChart_(dataObj) {
  /**
   * Gráfico de barras para los últimos 5 días de consumo total (egress).
   * dataObj: { headers: ['Fecha','Consumo Total (L)'], rows: [[Date, Number], ...] }
   */
  if (!dataObj || !dataObj.rows || dataObj.rows.length === 0) {
    // fallback
    dataObj = { headers: ['Fecha','Consumo Total (L)'], rows: [ [new Date(), 0] ] };
  }

  const dtb = Charts.newDataTable()
    .addColumn(Charts.ColumnType.DATE, dataObj.headers[0])
    .addColumn(Charts.ColumnType.NUMBER, dataObj.headers[1]);

  dataObj.rows.forEach(r => dtb.addRow(r));
  const dt = dtb.build();

  return Charts.newColumnChart()
    .setTitle('Consumo total últimos 5 días')
    .setXAxisTitle('Fecha')
    .setYAxisTitle('Litros')
    .setLegendPosition(Charts.Position.NONE)
    .setDimensions(900, 400)
    .setDataTable(dt)
    .build();
}

/* ============================== */
/*         ENVÍO DE EMAIL         */
/* ============================== */

/**
 * Envía email con tabla 2 columnas de gráficos (incluye los dos nuevos gráficos).
 */
function sendSummaryEmail(toEmail, emailsSheet, statsSheet, chartBlobs, isDuplicate, fechaHoraStr) {
  const senderEmail = "aguatorrebela@gmail.com";
  const subject = "💧 Resumen de las últimas 10 mediciones de agua reportadas";

  // Construir BCC (suscripciones activas < 6)
  const bccEmails = [];
  if (emailsSheet.getLastRow() > 0) {
    const emailsData = emailsSheet.getDataRange().getValues();
    for (let i = 1; i < emailsData.length; i++) {
      const email = emailsData[i][0] ? emailsData[i][0].toString().trim() : '';
      let contador = Number(emailsData[i][3]) || 0;
      if (isValidEmailCustom(email) && contador < 6) {
        bccEmails.push(email);
        emailsSheet.getRange(i + 1, 4).setValue(contador + 1);
      }
    }
  }

  // Últimas 10 filas
  const lastRow = statsSheet.getLastRow();
  const startRow = Math.max(2, lastRow - 9);
  const numRows = lastRow - startRow + 1;

  let lastDataRow = null;
  if (lastRow >= 2) {
    lastDataRow = statsSheet.getRange(lastRow, 1, 1, statsSheet.getLastColumn()).getDisplayValues()[0];
  }

  if (numRows < 1) {
    MailApp.sendEmail({
      to: toEmail,
      bcc: bccEmails.length > 0 ? bccEmails.join(",") : undefined,
      subject: subject,
      htmlBody: `<html><body><p>Estimado/a Vecino/a,</p><p>Aún no hay suficientes datos registrados para generar un resumen detallado.</p><p>Saludos cordiales,<br>Comisión de Agua del Edificio</p></body></html>`,
      from: senderEmail,
      name: "Comisión de Agua Torrebela"
    });
    return;
  }

  const statsData = statsSheet.getRange(startRow, 1, numRows, statsSheet.getLastColumn()).getDisplayValues();

  // Construir tabla de imágenes (dos columnas)
  let chartImagesHtml = '<table style="width: 100%; border-collapse: collapse;">';
  const inlineImages = {};

  // Orden de gráficos (incluyendo los nuevos al inicio para destacarlos):
  const chartsList = [
    // NUEVOS
    { blob: chartBlobs.projectionChart, cid: 'projectionChart', alt: 'Proyección próximos días: Consumo vs Llenado' },
    { blob: chartBlobs.last5daysChart, cid: 'last5daysChart', alt: 'Consumo total últimos 5 días' },
    // EXISTENTES
    { blob: chartBlobs.caudalChart, cid: 'caudalChart', alt: 'Gráfico de Caudal' },
    { blob: chartBlobs.combinadoChart, cid: 'combinadoChart', alt: 'Gráfico Combinado' },
    { blob: chartBlobs.last4WeeksChart, cid: 'last4WeeksChart', alt: 'Porcentaje del Tanque en las Últimas 5 Semanas' },
    { blob: chartBlobs.dailyChart, cid: 'dailyChart', alt: 'Consumo Diario Total' },
    { blob: chartBlobs.levelChart, cid: 'levelChart', alt: 'Nivel del Tanque con Umbrales' },
    { blob: chartBlobs.timeSlotChart, cid: 'timeSlotChart', alt: 'Consumo por Franja Horaria (Fines de Semana/Histórico)' },
    { blob: chartBlobs.dayOfWeekChart, cid: 'dayOfWeekChart', alt: 'Consumo por Día de la Semana' },
    { blob: chartBlobs.weeklyChart, cid: 'weeklyChart', alt: 'Consumo y Llenado por Semana' },
    { blob: chartBlobs.monthlyChart, cid: 'monthlyChart', alt: 'Consumo y Llenado por Mes' },
    { blob: chartBlobs.dailyWeekChart, cid: 'dailyWeekChart', alt: 'Consumo Diario por Día de la Semana' }
  ];

  for (let i = 0; i < chartsList.length; i += 2) {
    chartImagesHtml += '<tr>';

    const c1 = chartsList[i];
    if (c1 && c1.blob) {
      chartImagesHtml += `<td style="width: 50%; padding: 10px; vertical-align: top;">
        <img src="cid:${c1.cid}" alt="${c1.alt}" style="width: 100%; height: auto; border: 1px solid #eee; border-radius: 4px;">
        <div style="font-size:12px;color:#555;margin-top:6px;">${c1.alt}</div>
      </td>`;
      inlineImages[c1.cid] = c1.blob;
    } else {
      chartImagesHtml += '<td style="width: 50%; padding: 10px;"><p><em>Gráfico no disponible</em></p></td>';
    }

    const c2 = chartsList[i+1];
    if (c2 && c2.blob) {
      chartImagesHtml += `<td style="width: 50%; padding: 10px; vertical-align: top;">
        <img src="cid:${c2.cid}" alt="${c2.alt}" style="width: 100%; height: auto; border: 1px solid #eee; border-radius: 4px;">
        <div style="font-size:12px;color:#555;margin-top:6px;">${c2.alt}</div>
      </td>`;
      inlineImages[c2.cid] = c2.blob;
    } else {
      chartImagesHtml += '<td style="width: 50%; padding: 10px;"><p><em>Gráfico no disponible</em></p></td>';
    }

    chartImagesHtml += '</tr>';
  }
  chartImagesHtml += '</table>';

  if (Object.keys(inlineImages).length === 0) {
    chartImagesHtml = `<p>No se pudieron generar los gráficos debido a datos insuficientes o un error de procesamiento.</p>`;
  }

  let duplicateMessage = '';
  if (isDuplicate) {
    duplicateMessage = `<p><strong>Nota:</strong> El registro enviado para la fecha y hora ${fechaHoraStr} ya existe en la base de datos y no se insertó.</p>`;
  }

  let lastRecordHtml = lastDataRow ? `
    <h3>*** Último Registro ***</h3>
    <p><strong>Fecha y Hora:</strong> ${lastDataRow[7] || 'N/A'}</p>
    <p><strong>💧 Litros:</strong> ${lastDataRow[2] || 'N/A'}</p>
    <p><strong>📊 Porcentaje:</strong> ${lastDataRow[3] || 'N/A'}</p>
    <p><strong>📈 Última Variación (L):</strong> ${lastDataRow[5] || 'N/A'}</p>
    <p><strong>Caudal (lts/min):</strong> ${lastDataRow[9] || 'N/A'}</p>
    <p><strong>🔄 Tiempo (E) Llenado (días):</strong> ${lastDataRow[14] || 'N/A'}</p>
    <p><strong>📉 Tiempo (E) Vaciado (días):</strong> ${lastDataRow[17] || 'N/A'}</p>
    <p><strong>👥 Reportado por:</strong> ${lastDataRow[20] || 'N/A'}</p>
  ` : `<p>No hay un último registro disponible para mostrar.</p>`;

  // Cuerpo HTML
  let emailBody = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; color: #333; margin: 0; padding: 10px; line-height: 1.6; }
          .table-container { margin: 15px 0; }
          table { width: 100%; border-collapse: collapse; font-size: 14px; white-space: nowrap; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; vertical-align: top; }
          th { background-color: #f2f2f2; font-weight: bold; position: sticky; top: 0; }
          tr:nth-child(even) { background-color: #f9f9f9; }
          .observaciones { margin-top: 20px; font-style: italic; font-size: 0.9em; }
          h2, h3 { color: #1a73e8; }
          img { max-width: 100%; height: auto; }
          @media screen and (max-width: 600px) {
            body { font-size: 14px; }
            table { font-size: 12px; }
            th, td { padding: 6px; }
          }
        </style>
      </head>
      <body>
        <h2>✨ Resumen de las Últimas Mediciones de Agua ✨</h2>
        <p>Estimado/a Vecino/a,</p>
        <p>Le presentamos el resumen más reciente del nivel de agua en nuestro tanque, basado en los datos aportados por la comunidad.</p>
        <p>Su participación es clave para mantener un control eficiente del recurso hídrico.</p>
        ${duplicateMessage}
        <br>
        ${chartImagesHtml}
        <h3>**** Detalle de las Últimas Mediciones registradas en la base de datos ****:</h3>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Fecha y Hora</th>
                <th>Litros</th>
                <th>Porcentaje</th>
                <th>Variación (L)</th>
                <th>Caudal (lts/min)</th>
                <th>Tiempo (e) Llenado (días)</th>
                <th>Tiempo (e) Vaciado (días)</th>
                <th>Reportado por</th>
              </tr>
            </thead>
            <tbody>
  `;

  statsData.forEach(row => {
    emailBody += `
              <tr>
                <td>${row[7] || ''}</td>
                <td>${row[2] || ''}</td>
                <td>${row[3] || ''}</td>
                <td>${row[5] || ''}</td>
                <td>${row[9] || ''}</td>
                <td>${row[14] || ''}</td>
                <td>${row[17] || ''}</td>
                <td>${row[20] || ''}</td>
              </tr>
    `;
  });

  emailBody += `
            </tbody>
          </table>
        </div>
        ${lastRecordHtml}
        <div class="observaciones">
          <h3>*** Observaciones ***</h3>
          <ol>
            <li>Este correo presenta las últimas 10 mediciones registradas por los vecinos a través del formulario.</li>
            <li>El <strong>caudal neto (lts/min)</strong> representa la tasa de cambio en el volumen de agua, positiva al llenarse y negativa al consumirse.</li>
            <li>El <strong>tiempo estimado de llenado</strong> se basa en caudales positivos, proyectando el tiempo necesario para alcanzar la capacidad máxima (${TANK_CAPACITY_LITERS.toLocaleString('es-ES')} L).</li>
            <li>El <strong>tiempo estimado de vaciado</strong> se calcula con caudales negativos, estimando cuánto tardaría el tanque en vaciarse completamente.</li>
            <li><strong>Proyección (nuevos):</strong> Muestra el consumo y llenado promedio esperados para los próximos 5–7 días, basados en los promedios diarios históricos.</li>
            <li><strong>Consumo total últimos 5 días (nuevo):</strong> Suma diaria de consumo (litros) para los últimos 5 días con datos.</li>
            <li><strong>Gráfico de Consumo por Día de la Semana:</strong> Consumo promedio por día (todos los datos).</li>
            <li><strong>Gráfico por Semana (4 últimas):</strong> Litros entrantes y salientes totalizados por semana.</li>
            <li><strong>Gráfico por Mes:</strong> Litros entrantes y salientes totalizados por mes.</li>
            <li><strong>Gráfico Diario Total:</strong> Litros entrantes y salientes por día.</li>
            <li><strong>Gráfico de Nivel con Umbrales:</strong> Nivel del tanque con líneas de alerta (60%), racionamiento (40%) y crítico (20%).</li>
            <li><strong>Gráfico de Fines de Semana:</strong> Comparativa de consumo por franja horaria (actual vs anterior) y promedio histórico.</li>
            <li><strong>Consumo Diario por Día (semana):</strong> Promedios por día de la semana, comparando semana actual vs anterior.</li>
            <li>
              <p><strong>IMPORTANTE: Suscripción de resúmenes por correo</strong></p>
              <ul>
                <li>✉️ Cada vez que usted registra un nuevo dato e incluye su correo, activa la recepción de los <strong>próximos 5 resúmenes</strong>.</li>
                <li>➡️ Tras recibir 5 correos, el ciclo finaliza automáticamente.</li>
                <li>➡️ Para reactivarlo, registre un nuevo dato incluyendo su correo.</li>
              </ul>
            </li>
          </ol>
        </div>
        <br>
        <p>Agradecemos su colaboración en el monitoreo del agua. ¡Cada dato registrado ayuda a una mejor gestión!</p>
        <p>Saludos cordiales,</p>
        <p><strong>Comisión de Agua del Edificio</strong></p>
      </body>
    </html>
  `;

  Logger.log(`Imágenes incluidas (CIDs): ${Object.keys(inlineImages)}`);
  try {
    MailApp.sendEmail({
      to: toEmail,
      bcc: bccEmails.length > 0 ? bccEmails.join(",") : undefined,
      subject: subject,
      htmlBody: emailBody,
      name: "Comisión de Agua Torrebela",
      replyTo: senderEmail,
      from: senderEmail,
      inlineImages: inlineImages
    });
    Logger.log(`Correo enviado a ${toEmail} con ${bccEmails.length} BCC`);
  } catch (error) {
    Logger.log(`Error al enviar correo: ${error.message}`);
    MailApp.sendEmail({
      to: "correojago@gmail.com",
      subject: "ERROR al enviar resumen de agua",
      body: `No se pudo enviar el resumen desde ${senderEmail}. Error: ${error.message}`
    });
  }
}

/* ============================== */
/*   SOPORTE: TEST / FORMULAS     */
/* ============================== */

function testScriptManually() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const formResponsesSheet = spreadsheet.getSheetByName("Respuestas de formulario 1");
  if (!formResponsesSheet || formResponsesSheet.getLastRow() < 2) {
    Logger.log("No hay datos en 'Respuestas de formulario 1' para test.");
    return;
  }
  const lastRow = formResponsesSheet.getLastRow();
  const lastRowData = formResponsesSheet.getRange(lastRow, 1, 1, formResponsesSheet.getLastColumn()).getValues()[0];
  const simulatedEvent = {
    range: formResponsesSheet.getRange(lastRow, 1, 1, formResponsesSheet.getLastColumn()),
    values: [lastRowData]
  };
  onFormSubmitTrigger(simulatedEvent);
}

/**
 * Actualiza fórmulas después de inserción.
 */
function updateFormulasAfterInsert(statsSheet, insertedRow, tankCapacity) {
  const lastRow = statsSheet.getLastRow();
  if (insertedRow >= lastRow) return;

  const formulaColumns = [
    { col: 6, formula: (row, prevRow) => `=C${row}-C${prevRow}` },
    { col: 7, formula: (row, prevRow) => `=D${row}-D${prevRow}` },
    {
      col: 9,
      formula: (row, prevRow) => `=LET(
        inicio_str; H${prevRow};
        fin_str; H${row};
        PARSEAR_FECHA_HORA_GS; LAMBDA(texto_fecha_hora;
          LET(
            fecha_parte_str; LEFT(texto_fecha_hora; FIND(" "; texto_fecha_hora) - 1);
            hora_parte_str; MID(texto_fecha_hora; FIND(" "; texto_fecha_hora) + 1; LEN(texto_fecha_hora));
            fecha_num; DATEVALUE(TRIM(CLEAN(fecha_parte_str)));
            hora_val_num; VALUE(LEFT(hora_parte_str; FIND(":"; hora_parte_str) - 1));
            min_val_num; VALUE(MID(hora_parte_str; FIND(":"; hora_parte_str) + 1; 2));
            ampm_val_str; RIGHT(hora_parte_str; 2);
            hora_24h_num; IF(ampm_val_str = "pm"; hora_val_num + IF(hora_val_num = 12; 0; 12);
                             IF(ampm_val_str = "am"; hora_val_num - IF(hora_val_num = 12; 12; 0); hora_val_num));
            (fecha_num) + ((hora_24h_num * 60) + min_val_num) / 1440
          )
        );
        tiempo_inicio_completo; PARSEAR_FECHA_HORA_GS(inicio_str);
        tiempo_fin_completo; PARSEAR_FECHA_HORA_GS(fin_str);
        (tiempo_fin_completo - tiempo_inicio_completo) * 1440
      )`
    },
    { col: 10, formula: (row) => `=F${row}/I${row}` },
    { col: 11, formula: (row) => `=J${row}*60` },
    { col: 12, formula: (row) => `=(${tankCapacity}-C${row})` },
    { col: 13, formula: (row) => `=IF(J${row}>0;L${row}/J${row};"")` },
    { col: 14, formula: (row) => `=IF(J${row}>0;M${row}/60;"")` },
    { col: 15, formula: (row) => `=IF(J${row}>0;N${row}/24;"")` },
    { col: 16, formula: (row) => `=IF(J${row}<0;-C${row}/J${row};"")` },
    { col: 17, formula: (row) => `=IF(P${row}<>"";P${row}/60;"")` },
    { col: 18, formula: (row) => `=IF(P${row}<>"";Q${row}/24;"")` }
  ];

  for (let row = insertedRow + 1; row <= lastRow; row++) {
    const prevRow = row - 1;
    formulaColumns.forEach(({ col, formula }) => {
      if (row > 2) statsSheet.getRange(row, col).setFormula(formula(row, prevRow));
      else statsSheet.getRange(row, col).setValue("N/A");
    });
  }
}
