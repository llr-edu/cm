function panelAulasPorSeccion_(nivel, nomenclatura, incluirGenerales) {
  registrarHojasExtendidas_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.CATALOG);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colNomenclatura = buscarColumna_(headers, 'Nomenclatura');
  const colCourseId = buscarColumna_(headers, 'Course ID');
  const colNombreClassroom = panelBuscarColumnaOpcional_(headers, 'Nombre Classroom');
  const colNombreFuente = panelBuscarColumnaOpcional_(headers, 'Nombre fuente');
  const colMateria = panelBuscarColumnaOpcional_(headers, 'Materia Classroom');

  const target = normalizarNomenclatura_(nomenclatura);
  const general = obtenerGeneralNomenclatura_(target);
  const vistos = {};
  const aulas = [];

  data.slice(1).forEach(row => {
    if (String(row[colNivel] || '').trim() !== nivel) return;

    const actual = normalizarNomenclatura_(row[colNomenclatura]);
    if (actual !== target && !(incluirGenerales && actual === general)) return;

    const courseId = normalizarCourseId_(row[colCourseId]);
    if (!courseId || vistos[courseId]) return;

    vistos[courseId] = true;
    aulas.push({
      courseId: courseId,
      nomenclatura: actual,
      nombre: panelValorColumna_(row, colNombreClassroom) || panelValorColumna_(row, colNombreFuente) || panelValorColumna_(row, colMateria) || courseId
    });
  });

  return aulas
    .sort((a, b) => {
      const byName = panelOrdenTexto_(a.nombre).localeCompare(panelOrdenTexto_(b.nombre), 'es', { sensitivity: 'base', numeric: true });
      if (byName !== 0) return byName;
      return panelOrdenTexto_(a.nomenclatura).localeCompare(panelOrdenTexto_(b.nomenclatura), 'es', { sensitivity: 'base', numeric: true });
    })
    .map(aula => aula.courseId);
}

function panelAplicarAlumnoEnAulas_(correo, aulas, accion) {
  const courseIds = (aulas || [])
    .map(normalizarCourseId_)
    .filter(Boolean)
    .filter((courseId, index, all) => all.indexOf(courseId) === index);

  const result = { ok: 0, omitidos: 0, errores: [] };
  if (!courseIds.length) return result;

  const esEntrada = accion === 'entrada' || accion === 'agregar';

  if (esEntrada) {
    courseIds.forEach(courseId => {
      const response = panelSolicitudAlumnoAula_(correo, courseId, true);
      panelProcesarRespuestaAlumnoAula_(response, courseId, true, result);
      Utilities.sleep(120);
    });
    return result;
  }

  const requests = courseIds.map(courseId => panelCrearSolicitudAlumnoAula_(correo, courseId, false));
  let responses;

  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (error) {
    responses = requests.map(request => UrlFetchApp.fetch(request.url, request));
  }

  responses.forEach((response, index) => {
    panelProcesarRespuestaAlumnoAula_(response, courseIds[index], false, result);
  });

  return result;
}

function panelSolicitudAlumnoAula_(correo, courseId, esEntrada) {
  const request = panelCrearSolicitudAlumnoAula_(correo, courseId, esEntrada);
  return UrlFetchApp.fetch(request.url, request);
}

function panelCrearSolicitudAlumnoAula_(correo, courseId, esEntrada) {
  const token = ScriptApp.getOAuthToken();
  const headers = {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json'
  };

  if (esEntrada) {
    return {
      url: 'https://classroom.googleapis.com/v1/courses/' + encodeURIComponent(courseId) + '/students',
      method: 'post',
      headers: headers,
      payload: JSON.stringify({ userId: correo }),
      muteHttpExceptions: true
    };
  }

  return {
    url: 'https://classroom.googleapis.com/v1/courses/' + encodeURIComponent(courseId) + '/students/' + encodeURIComponent(correo),
    method: 'delete',
    headers: headers,
    muteHttpExceptions: true
  };
}

function panelProcesarRespuestaAlumnoAula_(response, courseId, esEntrada, result) {
  const code = response.getResponseCode();
  const body = response.getContentText() || '';

  if (code >= 200 && code < 300) {
    result.ok++;
    return;
  }

  if (esEntrada && (code === 409 || /already|already exists|Requested entity already exists/i.test(body))) {
    result.omitidos++;
    return;
  }

  if (!esEntrada && (code === 404 || /not found|Requested entity was not found/i.test(body))) {
    result.omitidos++;
    return;
  }

  result.errores.push(courseId + ': HTTP ' + code + ' ' + panelMensajeApiAulas_(body));
}

function panelMensajeApiAulas_(body) {
  if (!body) return '';
  try {
    const parsed = JSON.parse(body);
    if (parsed && parsed.error && parsed.error.message) return parsed.error.message;
  } catch (e) {}
  return String(body).substring(0, 240);
}

function panelBuscarColumnaOpcional_(headers, headerName) {
  const index = headers.indexOf(headerName);
  return index === -1 ? null : index;
}

function panelValorColumna_(row, colIndex) {
  if (colIndex === null || colIndex === undefined || colIndex < 0) return '';
  return String(row[colIndex] || '').trim();
}

function panelOrdenTexto_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}
