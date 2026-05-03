// ============================================================================
// JURITECNIA · GENERACIÓN DE NOTAS — SISTEMA DINÁMICO
// ============================================================================
// CONFIGURAR GEMINI:
//   Opción A (recomendada): introduce tu API key de Google AI Studio en
//   la constante GEMINI_API_KEY (consíguela en https://aistudio.google.com)
//
//   Opción B: deja GEMINI_API_KEY = '' y usa Vertex AI con OAuth
//   (requiere proyecto GCP con Vertex AI habilitado)
// ============================================================================

const GEMINI_API_KEY  = '';   // ← pega aquí tu API key de Google AI Studio
const VERTEX_PROJECT  = 'story-generator-94840038-614b0';
const VERTEX_REGION   = 'us-central1';
const GEMINI_MODEL    = 'gemini-2.0-flash';
const HOJA_CONFIG     = 'Configuracion';
const HOJA_RES        = 'Resultados';
const MAX_FEEDBACK    = 15;

const VERTEX_ENDPOINT =
  `https://${VERTEX_REGION}-aiplatform.googleapis.com/v1/projects/${VERTEX_PROJECT}` +
  `/locations/${VERTEX_REGION}/publishers/google/models/${GEMINI_MODEL}-001:generateContent`;

const AISTUDIO_ENDPOINT =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;


// ── MENÚ ─────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚖️ Juritecnia Notas')
    .addItem('⚙️ 1. Aplicar configuración',   'aplicarConfiguracion')
    .addItem('🧮 2. Calcular notas + formato', 'calcularNotas')
    .addSeparator()
    .addItem('📊 Informe grupal con gráficos', 'generarInformeGrupal')
    .addItem('📄 PDFs individuales',           'generarPDFsIndividuales')
    .addItem('🤖 Feedback Gemini (IA)',         'generarFeedbackGemini')
    .addSeparator()
    .addItem('🔑 Probar conexión Gemini',      'probarGemini')
    .addItem('📈 Ver estadísticas',            'verEstadisticas')
    .addItem('📋 Abrir panel lateral',         'mostrarSidebar')
    .addToUi();
}

function mostrarSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Juritecnia Notas').setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}


// ── LEER CONFIGURACIÓN ───────────────────────────────────────────────────────
function leerConfig() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_CONFIG);
  if (!hoja) throw new Error(`Hoja "${HOJA_CONFIG}" no encontrada. Ejecuta "Aplicar configuración" primero.`);

  const filas = hoja.getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[0].toString().trim() !== '');
  if (filas.length === 0) throw new Error(`La hoja "${HOJA_CONFIG}" está vacía.`);

  const componentes = filas.map((r, i) => ({
    nombre:    r[0].toString().trim(),
    peso:      parseFloat(r[1]) || 0,
    preguntas: parseFloat(r[2]) || 0,
    max:       parseFloat(r[3]) || 0,
    esTeorico: i === 0
  }));

  const totalPeso  = componentes.reduce((s, c) => s + c.peso, 0);
  const teorico    = componentes[0];
  const practicas  = componentes.slice(1);
  const preguntas  = teorico.preguntas || 20;
  const colPracticas = practicas.map((_, i) => 6 + i);
  const colTotal     = 6 + practicas.length;

  return { preguntas, componentes, teorico, practicas, totalPeso,
           colNombre:1, colPreguntas:2, colAciertos:3,
           colExamenTeo:4, colNotaTeo:5, colPracticas, colTotal, totalCols:colTotal };
}


// ── 1. APLICAR CONFIGURACIÓN ─────────────────────────────────────────────────
function aplicarConfiguracion() {
  const ui  = SpreadsheetApp.getUi();
  const cfg = leerConfig();

  const resp = ui.alert('⚙️ Aplicar configuración',
    `Componentes:\n` + cfg.componentes.map(c => `  · ${c.nombre}: ${c.peso}% (máx ${c.max} pts)`).join('\n') +
    `\n\nTotal: ${cfg.totalPeso}%\n\n¿Reconstruir cabeceras de "${HOJA_RES}"?`, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let  hRes = ss.getSheetByName(HOJA_RES) || ss.insertSheet(HOJA_RES);

  const cabeceras = [
    'NOMBRE', 'Preguntas', 'Aciertos\nTest',
    'Examen Teórico\n(/10) ←calc', `Nota Teórica\n(/${cfg.teorico.max}) ←calc`,
    ...cfg.practicas.map(p => `${p.nombre}\n(máx ${p.max} pts)`),
    'TOTAL NOTA\n(/10) ←calc'
  ];
  hRes.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]);

  if (hRes.getLastColumn() > cabeceras.length)
    hRes.deleteColumns(cabeceras.length + 1, hRes.getLastColumn() - cabeceras.length);

  const colores = ['#1a73e8','#455a64','#455a64','#0f9d58','#1565c0',
                   ...cfg.practicas.map(() => '#e37400'), '#7b1fa2'];
  colores.forEach((color, i) => {
    hRes.getRange(1, i + 1).setBackground(color).setFontColor('#ffffff')
        .setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  });
  hRes.setRowHeight(1, 44); hRes.setFrozenRows(1);
  hRes.setColumnWidth(1, 200); hRes.setColumnWidth(2, 70); hRes.setColumnWidth(3, 70);
  hRes.setColumnWidth(4, 120); hRes.setColumnWidth(5, 120);
  cfg.practicas.forEach((_, i) => hRes.setColumnWidth(6 + i, 100));
  hRes.setColumnWidth(cfg.colTotal, 120);

  cfg.practicas.forEach((p, i) => {
    hRes.getRange(2, 6 + i, 500, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireNumberBetween(0, p.max)
        .setHelpText(`0 – ${p.max}`).build());
  });

  ss.setActiveSheet(hRes);
  ui.alert(`✅ Hoja "${HOJA_RES}" configurada con ${cabeceras.length} columnas.\n\nIntroduce los datos de los alumnos y ejecuta "Calcular notas".`);
}


// ── 2. CALCULAR NOTAS ────────────────────────────────────────────────────────
function calcularNotas() {
  const ui  = SpreadsheetApp.getUi();
  const cfg = leerConfig();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const hRes = ss.getSheetByName(HOJA_RES);
  if (!hRes) throw new Error(`Hoja "${HOJA_RES}" no encontrada.`);

  const filas = _leerFilas(hRes, cfg);
  if (filas.length === 0) { ui.alert('⚠️ No hay alumnos en la hoja.'); return; }

  filas.forEach(({ rowNum, data }) => {
    const c = _calcular(data, cfg);
    hRes.getRange(rowNum, cfg.colExamenTeo).setValue(+c.examenTeorico.toFixed(2));
    hRes.getRange(rowNum, cfg.colNotaTeo).setValue(+c.notaTeorica.toFixed(2));
    hRes.getRange(rowNum, cfg.colTotal).setValue(+c.totalNota.toFixed(2));
  });

  _aplicarFormato(hRes, cfg, filas.length);
  ui.alert(`✅ Notas calculadas para ${filas.length} alumno(s).`);
}

function _aplicarFormato(hRes, cfg, numFilas) {
  if (numFilas === 0) return;
  const lastRow = 1 + numFilas;

  // ── Alturas de fila (fila 2 a lastRow) ──────────────────────────────────────
  for (let row = 2; row <= lastRow; row++) hRes.setRowHeight(row, 22);

  // ── Columnas con estilo constante (un solo batch por grupo) ─────────────────
  hRes.getRange(2, cfg.colPreguntas, numFilas, 2)
    .setBackground('#f1f3f4').setHorizontalAlignment('center').setNumberFormat('0');
  hRes.getRange(2, cfg.colExamenTeo, numFilas, 1)
    .setBackground('#e6f4ea').setFontColor('#0f9d58').setFontWeight('bold')
    .setHorizontalAlignment('center').setNumberFormat('0.00');
  hRes.getRange(2, cfg.colNotaTeo, numFilas, 1)
    .setBackground('#e8f0fe').setFontColor('#1565c0').setFontWeight('bold')
    .setHorizontalAlignment('center').setNumberFormat('0.00');
  if (cfg.colPracticas.length > 0) {
    hRes.getRange(2, cfg.colPracticas[0], numFilas, cfg.colPracticas.length)
      .setBackground('#fff8e1').setFontColor('#b45309')
      .setHorizontalAlignment('center').setNumberFormat('0.00');
  }

  // ── Columnas que varían por fila (nombre alternado, total por nota) ──────────
  for (let row = 2; row <= lastRow; row++) {
    hRes.getRange(row, cfg.colNombre).setBackground(row % 2 === 0 ? '#f8f9fa' : '#ffffff').setFontWeight('normal');
    const total = parseFloat(hRes.getRange(row, cfg.colTotal).getValue()) || 0;
    const { bg, fg } = _colorNota(total);
    hRes.getRange(row, cfg.colTotal).setBackground(bg).setFontColor(fg).setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0.00');
  }

  // ── Fila de medias ───────────────────────────────────────────────────────────
  const filaMed = lastRow + 1;
  hRes.setRowHeight(filaMed, 22);
  hRes.getRange(filaMed, 1).setValue('📊 MEDIA').setBackground('#e8eaed').setFontWeight('bold');
  for (let col = 2; col <= cfg.colTotal; col++) {
    const l = _colLetra(col);
    hRes.getRange(filaMed, col).setFormula(`=IFERROR(AVERAGE(${l}2:${l}${lastRow}),"-")`).setBackground('#e8eaed').setFontWeight('bold').setHorizontalAlignment('center').setNumberFormat('0.00');
  }
}


// ── 3. INFORME GRUPAL CON GRÁFICOS Y ANALÍTICA ───────────────────────────────
function generarInformeGrupal() {
  const ui = SpreadsheetApp.getUi();
  try {
    const cfg             = leerConfig();
    const { filas, stats} = _obtenerDatos(cfg);
    const ss              = SpreadsheetApp.getActiveSpreadsheet();
    const folder          = _carpeta('Juritecnia/Informes');

    // ── Crear hoja temporal para gráficos ──
    const hTemp = ss.insertSheet('__graficos_temp__');

    const doc  = DocumentApp.create('Informe Grupal - ' + _fechaHoy());
    const body = doc.getBody();
    body.setMarginLeft(50).setMarginRight(50);

    // ── PORTADA ──
    body.appendParagraph('JURITECNIA')
      .setHeading(DocumentApp.ParagraphHeading.TITLE)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setBold(true);
    body.appendParagraph('Informe Grupal de Evaluación')
      .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(new Date().toLocaleString('es-ES'))
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true);
    body.appendParagraph('');

    // ── SECCIÓN 1: PONDERACIÓN ──
    _seccion(body, '1. Ponderación aplicada');
    const filasConf = [['Componente', 'Peso (%)', 'Puntos máx.']];
    cfg.componentes.forEach(c => filasConf.push([c.nombre, c.peso + ' %', c.max + ' / 10']));
    filasConf.push(['TOTAL', cfg.totalPeso + ' %', '10 / 10']);
    const tPond = body.appendTable(filasConf);
    _estilarTablaDoc(tPond, '#1a73e8', filasConf[0].length);
    body.appendParagraph('');

    // ── SECCIÓN 2: ESTADÍSTICAS CLAVE ──
    _seccion(body, '2. Estadísticas del grupo');
    const pct = n => Math.round((n / stats.total) * 100);
    const dist = _distribucionCalificaciones(filas);

    const tStats = body.appendTable([
      ['Indicador', 'Valor'],
      ['Total alumnos',                  String(stats.total)],
      ['Nota media final',               stats.media.toFixed(2) + ' / 10'],
      ['Desviación típica',              stats.desviacion.toFixed(2)],
      ['Mediana',                        stats.mediana.toFixed(2) + ' / 10'],
      ['Nota más alta',                  stats.max.toFixed(2) + ' / 10'],
      ['Nota más baja',                  stats.min.toFixed(2) + ' / 10'],
      ['Media aciertos test',            stats.mediaAciertos.toFixed(1) + ' / ' + cfg.preguntas],
      ['Aprobados (≥ 5)',                 stats.aprobados + ' (' + pct(stats.aprobados) + ' %)'],
      ['Suspensos (< 5)',                 stats.suspensos + ' (' + pct(stats.suspensos) + ' %)'],
    ]);
    _estilarTablaDoc(tStats, '#0f9d58', 2);
    body.appendParagraph('');

    // ── SECCIÓN 3: DISTRIBUCIÓN DE CALIFICACIONES ──
    _seccion(body, '3. Distribución de calificaciones');
    const tDist = body.appendTable([
      ['Calificación', 'Alumnos', '% del grupo', 'Rango de nota'],
      ['Sobresaliente', String(dist.sobresaliente), pct(dist.sobresaliente) + ' %', '9.00 – 10.00'],
      ['Notable',       String(dist.notable),       pct(dist.notable) + ' %',       '7.00 – 8.99'],
      ['Bien',          String(dist.bien),           pct(dist.bien) + ' %',          '6.00 – 6.99'],
      ['Aprobado',      String(dist.aprobado),       pct(dist.aprobado) + ' %',      '5.00 – 5.99'],
      ['Suspenso',      String(dist.suspenso),       pct(dist.suspenso) + ' %',      '0.00 – 4.99'],
    ]);
    _estilarTablaDoc(tDist, '#7b1fa2', 4);
    body.appendParagraph('');

    // ── SECCIÓN 4: GRÁFICO — Notas finales por alumno ──
    _seccion(body, '4. Notas finales por alumno');

    const datosBarras = [['Alumno', 'Nota Final']];
    filas.forEach(r => datosBarras.push([r.nombre.split(' ')[0], r.totalNota]));

    const grafBarras = _crearGrafico(ss, hTemp, datosBarras,
      Charts.ChartType.BAR,
      'Notas finales por alumno',
      { 'hAxis.title': 'Nota (/10)', 'hAxis.minValue': 0, 'hAxis.maxValue': 10,
        'colors': ['#1a73e8'], 'legend.position': 'none',
        'chartArea.width': '70%', 'chartArea.height': '80%' },
      480, 280
    );
    if (grafBarras) {
      body.appendImage(grafBarras).setWidth(440).setHeight(260);
      body.appendParagraph('');
    }

    // ── SECCIÓN 5: GRÁFICO — Distribución de calificaciones ──
    _seccion(body, '5. Distribución de calificaciones');

    const datosTarta = [['Calificación', 'Alumnos'],
      ['Sobresaliente', dist.sobresaliente],
      ['Notable',       dist.notable],
      ['Bien',          dist.bien],
      ['Aprobado',      dist.aprobado],
      ['Suspenso',      dist.suspenso],
    ].filter((r, i) => i === 0 || r[1] > 0);

    const grafTarta = _crearGrafico(ss, hTemp, datosTarta,
      Charts.ChartType.PIE,
      'Distribución de calificaciones',
      { 'colors': ['#137333','#1967d2','#b45309','#f9ab00','#c5221f'],
        'pieSliceText': 'percentage', 'chartArea.width': '80%', 'chartArea.height': '80%' },
      380, 280
    );
    if (grafTarta) {
      body.appendImage(grafTarta).setWidth(340).setHeight(260);
      body.appendParagraph('');
    }

    // ── SECCIÓN 6: GRÁFICO — Medias por componente ──
    _seccion(body, '6. Media por componente de evaluación');

    const mediasComp = _mediasComponentes(filas, cfg);
    const datosComp  = [['Componente', 'Media (puntos obtenidos)', 'Máximo posible']];
    mediasComp.forEach(m => datosComp.push([m.nombre, m.media, m.max]));

    const grafComp = _crearGrafico(ss, hTemp, datosComp,
      Charts.ChartType.COLUMN,
      'Media obtenida vs máximo por componente',
      { 'colors': ['#1a73e8','#e0e0e0'], 'vAxis.minValue': 0,
        'vAxis.title': 'Puntos', 'chartArea.width': '75%', 'chartArea.height': '70%',
        'bar.groupWidth': '60%' },
      480, 280
    );
    if (grafComp) {
      body.appendImage(grafComp).setWidth(440).setHeight(260);
      body.appendParagraph('');
    }

    // ── SECCIÓN 7: ANALÍTICA POR COMPONENTE ──
    _seccion(body, '7. Analítica detallada por componente');
    const cabComp = ['Componente', 'Media', 'Máx obtenido', 'Mín obtenido', '% eficacia'];
    const rowsComp = [cabComp];
    mediasComp.forEach(m => rowsComp.push([
      m.nombre,
      m.media.toFixed(2) + ' / ' + m.max,
      m.maxObt.toFixed(2),
      m.minObt.toFixed(2),
      m.eficacia.toFixed(1) + ' %'
    ]));
    const tComp = body.appendTable(rowsComp);
    _estilarTablaDoc(tComp, '#e37400', cabComp.length);
    body.appendParagraph('');

    // ── SECCIÓN 8: TABLA COMPLETA DE ALUMNOS ──
    _seccion(body, '8. Notas individuales');
    const cabNotas = ['Alumno', `Ac.\n(/${cfg.preguntas})`, 'Ex.Teo\n(/10)', `Nota Teo\n(/${cfg.teorico.max})`];
    cfg.practicas.forEach(p => cabNotas.push(`${p.nombre}\n(/${p.max})`));
    cabNotas.push('TOTAL\n(/10)', 'Calificación');
    const rowsNotas = [cabNotas];
    filas.forEach(r => {
      const fila = [r.nombre, String(r.aciertos), r.examenTeorico.toFixed(1), r.notaTeorica.toFixed(2)];
      r.practicas.forEach(p => fila.push(p.toFixed(2)));
      fila.push(r.totalNota.toFixed(2), _calificacion(r.totalNota));
      rowsNotas.push(fila);
    });
    const tNotas = body.appendTable(rowsNotas);
    _estilarTablaDoc(tNotas, '#455a64', cabNotas.length);
    for (let r = 1; r < tNotas.getNumRows(); r++) {
      if (r % 2 === 0) tNotas.getRow(r).setBackgroundColor('#f8f9fa');
      const nota = parseFloat(rowsNotas[r][rowsNotas[r].length - 2]);
      if (!isNaN(nota)) {
        const {bg} = _colorNota(nota);
        tNotas.getRow(r).getCell(tNotas.getRow(r).getNumCells() - 2).setBackgroundColor(bg);
      }
    }

    body.appendParagraph('');
    body.appendParagraph('Prof. Francisco Javier Flor González')
      .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true);

    // Limpiar hoja temporal
    ss.deleteSheet(hTemp);

    doc.saveAndClose();
    const file = _moverACarpeta(doc.getId(), folder);
    ui.alert('✅ Informe grupal generado con gráficos.\n\n' + file.getUrl());
    return { success: true, url: file.getUrl(), name: file.getName() };

  } catch (e) {
    // Intentar limpiar hoja temporal si existe
    try {
      const hT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('__graficos_temp__');
      if (hT) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(hT);
    } catch (_) {}
    Logger.log(e.stack || e);
    ui.alert('❌ Error al generar informe: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Helper: crear gráfico como imagen en hoja temporal ──────────────────────
function _crearGrafico(ss, hTemp, datos, tipo, titulo, opciones, ancho, alto) {
  try {
    // Escribir datos en la hoja temporal (a partir de una columna libre)
    const startCol = hTemp.getLastColumn() + 2;
    const startRow = 1;
    datos.forEach((row, i) => {
      hTemp.getRange(startRow + i, startCol, 1, row.length).setValues([row]);
    });

    const rango = hTemp.getRange(startRow, startCol, datos.length, datos[0].length);
    let builder = hTemp.newChart()
      .setChartType(tipo)
      .addRange(rango)
      .setPosition(1, 1, 0, 0)
      .setOption('title', titulo)
      .setOption('width', ancho || 480)
      .setOption('height', alto || 300)
      .setOption('fontName', 'Arial')
      .setOption('titleTextStyle', { fontSize: 13, bold: true, color: '#202124' })
      .setOption('backgroundColor', '#ffffff');

    Object.entries(opciones || {}).forEach(([k, v]) => builder = builder.setOption(k, v));
    const chart = hTemp.insertChart(builder.build());
    Utilities.sleep(1500); // esperar render
    const blob = chart.getAs('image/png').setName(titulo + '.png');
    hTemp.removeChart(chart);
    return blob;
  } catch (e) {
    Logger.log('Error al crear gráfico "' + titulo + '": ' + e.message);
    return null;
  }
}


// ── 4. PDFs INDIVIDUALES ─────────────────────────────────────────────────────
function generarPDFsIndividuales() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('📄 PDFs', 'Se generará un PDF por alumno. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  try {
    const cfg       = leerConfig();
    const { filas } = _obtenerDatos(cfg);
    const folder    = _carpeta('Juritecnia/PDFs');
    let ok = 0, err = 0;

    filas.forEach((r, i) => {
      try {
        Utilities.sleep(800);
        const safe = r.nombre.replace(/[/\\:*?"<>|,]/g, '');
        const doc  = DocumentApp.create('Cert_' + safe);
        const body = doc.getBody();
        body.setMarginLeft(60).setMarginRight(60);

        body.appendParagraph('JURITECNIA')
          .setHeading(DocumentApp.ParagraphHeading.TITLE)
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(24).setBold(true);
        body.appendParagraph('Certificado de Evaluación')
          .setHeading(DocumentApp.ParagraphHeading.SUBTITLE)
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph(_fechaHoy())
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true);
        body.appendParagraph('');

        body.appendTable([['Alumno', r.nombre], ['Fecha evaluación', _fechaHoy()]]).setBorderWidth(0);
        body.appendParagraph('');
        body.appendHorizontalRule();
        body.appendParagraph('Desglose de la nota').setHeading(DocumentApp.ParagraphHeading.HEADING2);

        const desglose = [['Componente', 'Resultado', 'Peso', 'Puntos obtenidos']];
        desglose.push([
          `Examen Teórico (${r.aciertos}/${cfg.preguntas} aciertos)`,
          `${r.examenTeorico.toFixed(1)}/10`, `${cfg.teorico.peso}%`,
          `${r.notaTeorica.toFixed(2)} / ${cfg.teorico.max}`
        ]);
        cfg.practicas.forEach((p, pi) => desglose.push([
          p.nombre, `${r.practicas[pi].toFixed(2)}/${p.max}`,
          `${p.peso}%`, `${r.practicas[pi].toFixed(2)} / ${p.max}`
        ]));

        const tD = body.appendTable(desglose);
        _estilarTablaDoc(tD, '#4285f4', 4);

        body.appendParagraph('');
        body.appendParagraph(`NOTA FINAL: ${r.totalNota.toFixed(2)} / 10  —  ${_calificacion(r.totalNota)}`)
          .setFontSize(16).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        body.appendHorizontalRule();
        body.appendParagraph('Prof. Francisco Javier Flor González')
          .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true);

        doc.saveAndClose();
        const docFile = DriveApp.getFileById(doc.getId());
        folder.createFile(docFile.getAs('application/pdf').setName(safe + '_Juritecnia.pdf'));
        docFile.setTrashed(true);
        ok++;
      } catch (ex) { err++; Logger.log(`Error PDF ${i}: ${ex.message}`); }
    });

    ui.alert(`✅ PDFs: ${ok} generados, ${err} errores.\nCarpeta: Juritecnia/PDFs en Drive.`);
    return { success: true, ok, err };
  } catch (e) {
    Logger.log(e); ui.alert('❌ ' + e.message);
    return { success: false, error: e.message };
  }
}


// ── 5. FEEDBACK GEMINI ───────────────────────────────────────────────────────
function generarFeedbackGemini() {
  const ui = SpreadsheetApp.getUi();
  try {
    const testOk = _testGemini();
    if (!testOk.ok) {
      ui.alert('❌ No se puede conectar con Gemini.\n\n' + testOk.error +
        '\n\n💡 Solución: introduce tu API key de Google AI Studio en la\nconstante GEMINI_API_KEY del código.');
      return { success: false, error: testOk.error };
    }

    const cfg       = leerConfig();
    const { filas } = _obtenerDatos(cfg);
    const lote      = filas.slice(0, MAX_FEEDBACK);
    const folder    = _carpeta('Juritecnia/Feedback');

    const { file, count, errores } = _generarDocFeedback(cfg, lote, folder);
    const msg = `✅ Feedback generado: ${count} correctos, ${errores} errores.\n\nURL:\n${file.getUrl()}`;
    ui.alert(msg);
    return { success: true, count, errores, url: file.getUrl(), name: file.getName() };

  } catch (e) {
    Logger.log(e.stack || e);
    ui.alert('❌ Error al generar feedback: ' + e.message);
    return { success: false, error: e.message };
  }
}

// ── Helper: construir prompt de feedback para un alumno ──────────────────────
function _buildFeedbackPrompt(r, cfg) {
  const lineasPract = cfg.practicas.map((p, pi) =>
    `- ${p.nombre}: ${r.practicas[pi].toFixed(2)}/${p.max} pts (peso ${p.peso}%)`).join('\n');
  return (
    `Eres el profesor Flor de Juritecnia. Tu alumno ${r.nombre} ha obtenido:\n` +
    `- Test teórico: ${r.aciertos}/${cfg.preguntas} aciertos → ${r.examenTeorico.toFixed(1)}/10 → ${r.notaTeorica.toFixed(2)} puntos (peso ${cfg.teorico.peso}%)\n` +
    lineasPract + '\n' +
    `- NOTA FINAL: ${r.totalNota.toFixed(2)}/10 → ${_calificacion(r.totalNota)}\n\n` +
    `Escribe un feedback personalizado de exactamente 4 párrafos breves:\n` +
    `1) Valora su rendimiento en el test teórico (menciona si los aciertos son buenos o mejorables).\n` +
    `2) Comenta las prácticas: señala cuál fue su punto más fuerte y cuál necesita mejorar.\n` +
    `3) Da un consejo académico concreto y accionable para mejorar.\n` +
    `4) Cierra con una frase motivadora personalizada.\n` +
    `Usa un tono cercano pero profesional. Firma como "Prof. Flor".`
  );
}

// ── Helper: crear documento de feedback para un lote de alumnos ──────────────
function _generarDocFeedback(cfg, lote, folder) {
  const doc  = DocumentApp.create('Feedback Gemini - ' + _fechaHoy());
  const body = doc.getBody();
  body.setMarginLeft(60).setMarginRight(60);

  body.appendParagraph('JURITECNIA · FEEDBACK GEMINI IA')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(`Feedback personalizado — ${lote.length} alumno(s) — ${_fechaHoy()}`)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true);
  body.appendParagraph('');

  let count = 0, errores = 0;
  lote.forEach((r, idx) => {
    Utilities.sleep(1500);
    let feedback = '';
    try {
      feedback = _llamarGemini(_buildFeedbackPrompt(r, cfg));
      count++;
    } catch (e) {
      feedback = `⚠️ No se pudo generar feedback automático.\nError: ${e.message}`;
      errores++;
      Logger.log(`Gemini error (${r.nombre}): ${e.message}`);
    }

    body.appendParagraph(`${idx + 1}. ${r.nombre}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1).setBold(true);
    body.appendParagraph(`Nota final: ${r.totalNota.toFixed(2)}/10 — ${_calificacion(r.totalNota)}`)
      .setItalic(true).setFontSize(11);
    body.appendParagraph(feedback).setFontSize(11);
    body.appendParagraph('');
    body.appendHorizontalRule();
    body.appendParagraph('');
  });

  body.appendParagraph('Prof. Francisco Javier Flor González')
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setItalic(true);

  doc.saveAndClose();
  const file = _moverACarpeta(doc.getId(), folder);
  return { file, count, errores };
}

// ── Probar conexión Gemini ────────────────────────────────────────────────────
function probarGemini() {
  const ui  = SpreadsheetApp.getUi();
  const res = _testGemini();
  if (res.ok) {
    ui.alert('✅ Gemini responde correctamente.\n\nRespuesta: ' + res.respuesta);
  } else {
    ui.alert('❌ No se pudo conectar con Gemini.\n\n' + res.error +
      '\n\n💡 Introduce tu API key en la constante GEMINI_API_KEY del código.');
  }
}

function _testGemini() {
  try {
    const respuesta = _llamarGemini('Responde solo con: "Juritecnia OK"', 20);
    return { ok: true, respuesta };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}


// ── 6. ESTADÍSTICAS ──────────────────────────────────────────────────────────
function verEstadisticas() {
  const ui = SpreadsheetApp.getUi();
  try {
    const cfg       = leerConfig();
    const { stats } = _obtenerDatos(cfg);
    const pct       = n => Math.round((n / stats.total) * 100);
    ui.alert('📊 Estadísticas del grupo',
      `Total alumnos:        ${stats.total}\n` +
      `Nota media final:     ${stats.media.toFixed(2)} / 10\n` +
      `Mediana:              ${stats.mediana.toFixed(2)} / 10\n` +
      `Desviación típica:    ${stats.desviacion.toFixed(2)}\n` +
      `Nota más alta:        ${stats.max.toFixed(2)} / 10\n` +
      `Nota más baja:        ${stats.min.toFixed(2)} / 10\n` +
      `Media aciertos test:  ${stats.mediaAciertos.toFixed(1)} / ${cfg.preguntas}\n` +
      `────────────────────────────────\n` +
      `Aprobados (≥ 5):      ${stats.aprobados} (${pct(stats.aprobados)} %)\n` +
      `Suspensos  (< 5):     ${stats.suspensos} (${pct(stats.suspensos)} %)`,
      ui.ButtonSet.OK);
    return { success: true, ...stats };
  } catch (e) {
    ui.alert('❌ ' + e.message);
    return { success: false, error: e.message };
  }
}


// ── LÓGICA DE CÁLCULO ────────────────────────────────────────────────────────

function _calcular(data, cfg) {
  const preguntas     = parseFloat(data[cfg.colPreguntas - 1]) || cfg.preguntas;
  const aciertos      = parseFloat(data[cfg.colAciertos  - 1]) || 0;
  const examenTeorico = (aciertos / preguntas) * 10;
  const notaTeorica   = (examenTeorico / 10) * cfg.teorico.max;
  const practicas     = cfg.colPracticas.map((col, i) =>
    Math.min(Math.max(parseFloat(data[col - 1]) || 0, 0), cfg.practicas[i].max));
  const totalNota     = notaTeorica + practicas.reduce((s, v) => s + v, 0);
  return { aciertos, preguntas, examenTeorico, notaTeorica, practicas, totalNota };
}

function _calificacion(nota) {
  if (nota >= 9) return 'Sobresaliente';
  if (nota >= 7) return 'Notable';
  if (nota >= 6) return 'Bien';
  if (nota >= 5) return 'Aprobado';
  if (nota >= 3) return 'Suspenso';
  return 'Muy deficiente';
}

function _colorNota(nota) {
  if (nota >= 9) return { bg:'#e6f4ea', fg:'#137333' };
  if (nota >= 7) return { bg:'#e8f0fe', fg:'#1967d2' };
  if (nota >= 6) return { bg:'#fef7e0', fg:'#b45309' };
  if (nota >= 5) return { bg:'#fff8e1', fg:'#b45309' };
  return             { bg:'#fce8e6', fg:'#c5221f'  };
}

function _distribucionCalificaciones(filas) {
  const d = { sobresaliente:0, notable:0, bien:0, aprobado:0, suspenso:0 };
  filas.forEach(r => {
    const n = r.totalNota;
    if (n >= 9) d.sobresaliente++;
    else if (n >= 7) d.notable++;
    else if (n >= 6) d.bien++;
    else if (n >= 5) d.aprobado++;
    else d.suspenso++;
  });
  return d;
}

function _mediasComponentes(filas, cfg) {
  const res = [];

  // Teórico (en puntos)
  const notasTeo = filas.map(r => r.notaTeorica);
  const mediaTeo = notasTeo.reduce((a, b) => a + b, 0) / notasTeo.length;
  res.push({
    nombre:   cfg.teorico.nombre,
    media:    mediaTeo,
    max:      cfg.teorico.max,
    maxObt:   Math.max(...notasTeo),
    minObt:   Math.min(...notasTeo),
    eficacia: (mediaTeo / cfg.teorico.max) * 100
  });

  // Prácticas
  cfg.practicas.forEach((p, pi) => {
    const vals  = filas.map(r => r.practicas[pi]);
    const media = vals.reduce((a, b) => a + b, 0) / vals.length;
    res.push({
      nombre:   p.nombre,
      media,
      max:      p.max,
      maxObt:   Math.max(...vals),
      minObt:   Math.min(...vals),
      eficacia: (media / p.max) * 100
    });
  });

  return res;
}


// ── HELPERS DE DATOS ─────────────────────────────────────────────────────────

function _leerFilas(sheet, cfg) {
  const all  = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < all.length; i++) {
    const nombre = (all[i][cfg.colNombre - 1] || '').toString().trim();
    if (nombre) rows.push({ rowNum: i + 1, data: all[i] });
  }
  return rows;
}

function _obtenerDatos(cfg) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hRes = ss.getSheetByName(HOJA_RES);
  if (!hRes) throw new Error(`Hoja "${HOJA_RES}" no encontrada. Ejecuta "Aplicar configuración" primero.`);

  const filas = _leerFilas(hRes, cfg);
  if (filas.length === 0) throw new Error(`No hay datos en "${HOJA_RES}".`);

  const resultado = filas.map(({ data }) => {
    const nombre  = (data[cfg.colNombre - 1] || '').toString().trim();
    const enHoja  = parseFloat(data[cfg.colTotal - 1]);
    const calc    = _calcular(data, cfg);
    return { nombre, ...calc, totalNota: isNaN(enHoja) ? calc.totalNota : enHoja };
  });

  const notas    = resultado.map(r => r.totalNota).sort((a,b) => a-b);
  const aciertos = resultado.map(r => r.aciertos);
  const media    = notas.reduce((a,b)=>a+b,0) / notas.length;
  const mediana  = notas.length % 2 === 0
    ? (notas[notas.length/2-1] + notas[notas.length/2]) / 2
    : notas[Math.floor(notas.length/2)];
  const desviacion = Math.sqrt(notas.map(n=>(n-media)**2).reduce((a,b)=>a+b,0)/notas.length);

  const stats = {
    total:         resultado.length,
    media, mediana, desviacion,
    max:           Math.max(...notas),
    min:           Math.min(...notas),
    mediaAciertos: aciertos.reduce((a,b)=>a+b,0)/aciertos.length,
    aprobados:     notas.filter(n=>n>=5).length,
    suspensos:     notas.filter(n=>n<5).length
  };
  return { filas: resultado, stats };
}


// ── GEMINI ───────────────────────────────────────────────────────────────────

function _llamarGemini(prompt, maxTokens) {
  maxTokens = maxTokens || 800;
  const payload = {
    contents: [{ role:'user', parts:[{ text: prompt }] }],
    generationConfig: { temperature: 0.7, maxOutputTokens: maxTokens }
  };

  // Opción A: Google AI Studio (API key)
  if (GEMINI_API_KEY && GEMINI_API_KEY.trim() !== '') {
    const url  = AISTUDIO_ENDPOINT + '?key=' + GEMINI_API_KEY;
    const opts = { method:'post', contentType:'application/json',
                   payload: JSON.stringify(payload), muteHttpExceptions: true };
    return _fetchConReintentos(url, opts, 'Gemini API');
  }

  // Opción B: Vertex AI (OAuth)
  const opts = {
    method:'post', contentType:'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  return _fetchConReintentos(VERTEX_ENDPOINT, opts, 'Vertex AI',
    'Comprueba que el proyecto GCP tiene Vertex AI habilitado o configura GEMINI_API_KEY.');
}

// ── Helper: ejecutar fetch con reintentos ante errores 429/5xx ───────────────
function _fetchConReintentos(url, opts, nombreServicio, sugerencia) {
  for (let i = 1; i <= 3; i++) {
    const res  = UrlFetchApp.fetch(url, opts);
    const code = res.getResponseCode();
    const body = res.getContentText();
    if ((code === 429 || code >= 500) && i < 3) { Utilities.sleep(2000 * i); continue; }
    const json  = JSON.parse(body);
    const texto = json?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (texto) return texto.trim();
    if (json.error) throw new Error(`${nombreServicio} error ${json.error.code}: ${json.error.message}`);
    const extra = sugerencia ? ' ' + sugerencia : '';
    throw new Error(`Sin respuesta válida de ${nombreServicio} (HTTP ${code}).${extra}`);
  }
  throw new Error(`${nombreServicio} no respondió tras 3 intentos.`);
}


// ── HELPERS GENERALES ────────────────────────────────────────────────────────

function _seccion(body, titulo) {
  body.appendParagraph(titulo).setHeading(DocumentApp.ParagraphHeading.HEADING1);
}

function _estilarTablaDoc(tabla, colorHex, numCols) {
  const cab = tabla.getRow(0);
  cab.editAsText().setBold(true);
  for (let c = 0; c < numCols; c++) cab.getCell(c).setBackgroundColor(colorHex);
  cab.editAsText().setForegroundColor('#ffffff');
}

function _fechaHoy() { return new Date().toLocaleDateString('es-ES'); }

function _colLetra(col) {
  let s = '';
  while (col > 0) { s = String.fromCharCode(((col-1)%26)+65)+s; col=Math.floor((col-1)/26); }
  return s;
}

function _carpeta(ruta) {
  let f = DriveApp.getRootFolder();
  for (const parte of ruta.split('/')) {
    const it = f.getFoldersByName(parte);
    f = it.hasNext() ? it.next() : f.createFolder(parte);
  }
  return f;
}

function _moverACarpeta(docId, folder) {
  const file = DriveApp.getFileById(docId);
  folder.addFile(file); DriveApp.getRootFolder().removeFile(file);
  return file;
}


// ── WRAPPERS SIDEBAR ─────────────────────────────────────────────────────────
function wb_leerConfig() {
  try {
    const cfg = leerConfig();
    return { success:true, preguntas:cfg.preguntas, totalPeso:cfg.totalPeso,
      componentes: cfg.componentes.map(c=>({nombre:c.nombre,peso:c.peso,max:c.max,esTeorico:c.esTeorico,preguntas:c.preguntas})) };
  } catch(e) { return { success:false, error:e.message }; }
}
function wb_guardarConfig(datos) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    let hConf  = ss.getSheetByName(HOJA_CONFIG) || ss.insertSheet(HOJA_CONFIG);
    hConf.getRange(1,1,1,4).setValues([['Componente','Peso (%)','Preguntas','Punt. Máxima']]);
    hConf.getRange(1,1,1,4).setBackground('#37474f').setFontColor('#fff').setFontWeight('bold');
    const filas = datos.componentes.map(c=>[c.nombre,c.peso,c.esTeorico?(c.preguntas||20):'',c.max]);
    if (hConf.getLastRow()>1) hConf.getRange(2,1,Math.max(hConf.getLastRow()-1,filas.length),4).clearContent();
    hConf.getRange(2,1,filas.length,4).setValues(filas);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}
function wb_aplicarConfiguracion() {
  try { aplicarConfiguracion(); return { success:true }; }
  catch(e) { return { success:false, error:e.message }; }
}
function wb_calcularNotas() {
  try { calcularNotas(); return { success:true }; }
  catch(e) { return { success:false, error:e.message }; }
}
function wb_informe()  { return generarInformeGrupal(); }
function wb_pdfs() {
  try { generarPDFsIndividuales(); return { success:true }; }
  catch(e) { return { success:false, error:e.message }; }
}
function wb_feedback() {
  // Wrapper sin ui.alert para el sidebar — usa los helpers compartidos
  try {
    const cfg       = leerConfig();
    const { filas } = _obtenerDatos(cfg);
    const lote      = filas.slice(0, MAX_FEEDBACK);
    const folder    = _carpeta('Juritecnia/Feedback');
    const { file, count, errores } = _generarDocFeedback(cfg, lote, folder);
    return { success: true, count, errores, url: file.getUrl(), name: file.getName() };
  } catch (e) {
    Logger.log(e.stack || e);
    return { success: false, error: e.message };
  }
}
function wb_stats()    { return verEstadisticas(); }
function wb_testGemini() {
  const res = _testGemini();
  return res.ok
    ? { success:true, msg:'✅ Gemini responde: ' + res.respuesta }
    : { success:false, error:res.error };
}
