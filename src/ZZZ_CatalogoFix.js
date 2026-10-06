/****************************************************
 * Ajustes de sincronización del catálogo
 * Normaliza nomenclaturas con ceros a la izquierda.
 ****************************************************/

function normalizarNomenclatura_(value) {
  if (value === null || value === undefined) return '';
  let text = String(value).trim();
  if (!text) return '';

  // Corrige valores que Sheets pueda entregar como 11, 11.0 o 11,0
  text = text.replace(/\.0$/, '').replace(/,0$/, '');

  if (/^\d+$/.test(text) && text.length < 3) {
    return text.padStart(3, '0');
  }

  return text;
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

        if (!courseId) return;
        diagnostico.filasConCodigo++;

        if (!nomenclatura) {
          diagnostico.filasSinNomenclatura++;
          return;
        }

        if (validas.length && !validas.includes(nomenclatura)) {
          diagnostico.filasSaltadasPorNomenclatura++;
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

  const detalle =
    'Fuentes: ' + fuentes.length +
    ' | Filas leídas: ' + diagnostico.filasLeidas +
    ' | Con código: ' + diagnostico.filasConCodigo +
    ' | Sin nomenclatura: ' + diagnostico.filasSinNomenclatura +
    ' | Saltadas por nomenclatura: ' + diagnostico.filasSaltadasPorNomenclatura;

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

  return data.slice(1)
    .filter(row => String(row[colNivel] || '').trim() === nivel)
    .filter(row => normalizarNomenclatura_(row[colNomenclatura]) === target)
    .map(row => normalizarCourseId_(row[colCourseId]))
    .filter(Boolean);
}
