/****************************************************
 * Ajustes de sincronización del catálogo
 * Usa nomenclaturas completas de sección: 072A, 072B, 072C.
 * Mantiene compatibilidad con códigos generales tipo 072ABC.
 ****************************************************/

function normalizarNomenclatura_(value) {
  if (value === null || value === undefined) return '';
  let text = String(value).trim().toUpperCase();
  if (!text) return '';

  // Corrige valores que Sheets pueda entregar como 72, 72.0, 72,0 o con espacios.
  text = text
    .replace(/\.0$/, '')
    .replace(/,0$/, '')
    .replace(/\s+/g, '')
    .replace(/-/g, '');

  const match = text.match(/^(\d{1,3})([A-Z]+)?$/);
  if (match) {
    const base = match[1].padStart(3, '0');
    const sufijo = match[2] || '';
    return base + sufijo;
  }

  return text;
}

function obtenerBaseNomenclatura_(value) {
  const normalizada = normalizarNomenclatura_(value);
  const match = normalizada.match(/^(\d{3})/);
  return match ? match[1] : normalizada;
}

function obtenerGeneralNomenclatura_(value) {
  const base = obtenerBaseNomenclatura_(value);
  return base ? base + 'ABC' : '';
}

function esNomenclaturaPermitida_(nomenclatura, validas) {
  const normalizada = normalizarNomenclatura_(nomenclatura);
  if (!normalizada) return false;
  if (!validas || !validas.length) return true;

  const base = obtenerBaseNomenclatura_(normalizada);
  return validas.includes(normalizada) || validas.includes(base);
}

function normalizarCourseId_(value) {
  if (value === null || value === undefined) return '';

  if (typeof value === 'number') {
    return Utilities.formatString('%.0f', value).trim();
  }

  return String(value).trim();
}

function obtenerFuentesAulas_() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.SOURCES);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colSpreadsheetId = buscarColumna_(headers, 'Spreadsheet ID');
  const colPestana = buscarColumna_(headers, 'Pestaña');
  const colCodigo = buscarColumna_(headers, 'Col código');
  const colNombre = buscarColumna_(headers, 'Col nombre');
  const colDescripcion = buscarColumna_(headers, 'Col descripción');
  const colNomenclatura = buscarColumna_(headers, 'Col nomenclatura');
  const colNomenclaturas = buscarColumna_(headers, 'Nomenclaturas válidas');
  const colActivo = buscarColumna_(headers, 'Activo');

  return data.slice(1)
    .filter(row => esSi_(row[colActivo]))
    .map(row => ({
      nivel: String(row[colNivel] || '').trim(),
      spreadsheetId: String(row[colSpreadsheetId] || '').trim(),
      pestana: String(row[colPestana] || '').trim(),
      colCodigo: columnaALetraNumero_(row[colCodigo]),
      colNombre: columnaALetraNumero_(row[colNombre]),
      colDescripcion: columnaALetraNumero_(row[colDescripcion]),
      colNomenclatura: columnaALetraNumero_(row[colNomenclatura]),
      nomenclaturasValidas: String(row[colNomenclaturas] || '')
        .split(',')
        .map(item => normalizarNomenclatura_(item))
        .filter(Boolean)
    }))
    .filter(fuente => fuente.nivel && fuente.spreadsheetId && fuente.pestana);
}

function sincronizarCatalogoAulas() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fuentes = obtenerFuentesAulas_();
  const catalog = obtenerHojaObligatoria_(ss, SHEETS.CATALOG);

  if (fuentes.length === 0) {
    SpreadsheetApp.getUi().alert('No hay fuentes activas en la hoja FUENTES.');
    return;
  }

  const rows = [];
  const ahora = new Date();
  const errores = [];
  const nomenclaturasEncontradas = new Set();
  const ejemplosSaltados = [];
  const diagnostico = {
    fuentes: fuentes.length,
    filasLeidas: 0,
    filasConCodigo: 0,
    filasSinNomenclatura: 0,
    filasSaltadasPorNomenclatura: 0
  };

  fuentes.forEach(fuente => {
    try {
      const sourceSs = SpreadsheetApp.openById(fuente.spreadsheetId);
      const sourceSheet = sourceSs.getSheetByName(fuente.pestana);
      if (!sourceSheet) throw new Error('No existe la pestaña: ' + fuente.pestana);

      const lastRow = sourceSheet.getLastRow();
      if (lastRow < 2) return;

      const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
      const validas = fuente.nomenclaturasValidas;
      diagnostico.filasLeidas += data.length;

      data.forEach((row, index) => {
        const courseId = normalizarCourseId_(row[fuente.colCodigo - 1]);
        const nombre = String(row[fuente.colNombre - 1] || '').trim();
        const descripcion = String(row[fuente.colDescripcion - 1] || '').trim();
        const nomenclatura = normalizarNomenclatura_(row[fuente.colNomenclatura - 1]);

        if (nomenclatura) nomenclaturasEncontradas.add(nomenclatura);
        if (!courseId) return;
        diagnostico.filasConCodigo++;

        if (!nomenclatura) {
          diagnostico.filasSinNomenclatura++;
          return;
        }

        if (!esNomenclaturaPermitida_(nomenclatura, validas)) {
          diagnostico.filasSaltadasPorNomenclatura++;
          if (ejemplosSaltados.length < 8) ejemplosSaltados.push(nomenclatura);
          return;
        }

        let classroom = null;
        try {
          classroom = Classroom.Courses.get(courseId);
        } catch (classroomError) {
          errores.push(fuente.nivel + ' fila ' + (index + 2) + ' | ' + courseId + ': ' + (classroomError.message || classroomError));
        }

        rows.push([
          fuente.nivel,
          nomenclatura,
          courseId,
          nombre,
          descripcion,
          classroom && classroom.name ? classroom.name : '',
          classroom && classroom.section ? classroom.section : '',
          classroom && classroom.subject ? classroom.subject : '',
          classroom && classroom.courseState ? classroom.courseState : '',
          classroom && classroom.alternateLink ? classroom.alternateLink : '',
          fuente.pestana + ' · ' + fuente.spreadsheetId,
          ahora
        ]);
      });
    } catch (error) {
      errores.push(fuente.nivel + ': ' + (error.message || error));
      registrarError_('sincronizarCatalogoAulas', '', '', error);
    }
  });

  limpiarDatosManteniendoEncabezado_(catalog);
  if (rows.length > 0) {
    catalog.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    catalog.autoResizeColumns(1, catalog.getLastColumn());
  }

  const muestraNomenclaturas = Array.from(nomenclaturasEncontradas).sort().slice(0, 20).join(', ');
  const detalle =
    'Fuentes: ' + fuentes.length +
    ' | Filas leídas: ' + diagnostico.filasLeidas +
    ' | Con código: ' + diagnostico.filasConCodigo +
    ' | Sin nomenclatura: ' + diagnostico.filasSinNomenclatura +
    ' | Saltadas por nomenclatura: ' + diagnostico.filasSaltadasPorNomenclatura +
    ' | Muestra: ' + muestraNomenclaturas;

  registrarLog_('Sincronizar catálogo de aulas', detalle, rows.length + ' aulas sincronizadas; errores: ' + errores.length);

  SpreadsheetApp.getUi().alert(
    'Sincronización finalizada.\n\n' +
    'Aulas agregadas al catálogo: ' + rows.length + '\n' +
    'Errores: ' + errores.length + '\n\n' +
    'Diagnóstico:\n' +
    'Filas leídas: ' + diagnostico.filasLeidas + '\n' +
    'Filas con código: ' + diagnostico.filasConCodigo + '\n' +
    'Sin nomenclatura: ' + diagnostico.filasSinNomenclatura + '\n' +
    'Saltadas por nomenclatura: ' + diagnostico.filasSaltadasPorNomenclatura +
    (muestraNomenclaturas ? '\nNomenclaturas detectadas: ' + muestraNomenclaturas : '') +
    (ejemplosSaltados.length ? '\nEjemplos saltados: ' + ejemplosSaltados.join(', ') : '') +
    (errores.length ? '\n\nPrimer error:\n' + errores[0] : '')
  );
}

function obtenerAulasCatalogo_(nivel, nomenclatura) {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.CATALOG);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colNomenclatura = buscarColumna_(headers, 'Nomenclatura');
  const colCourseId = buscarColumna_(headers, 'Course ID');
  const target = normalizarNomenclatura_(nomenclatura);
  const general = obtenerGeneralNomenclatura_(target);

  return data.slice(1)
    .filter(row => String(row[colNivel] || '').trim() === nivel)
    .filter(row => {
      const actual = normalizarNomenclatura_(row[colNomenclatura]);
      return actual === target || actual === general;
    })
    .map(row => normalizarCourseId_(row[colCourseId]))
    .filter(Boolean)
    .filter((courseId, index, all) => all.indexOf(courseId) === index);
}
