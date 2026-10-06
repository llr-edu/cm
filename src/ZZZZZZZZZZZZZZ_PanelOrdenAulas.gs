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
    aulas.push(courseId);
  });

  return aulas;
}

function panelAplicarAlumnoEnAulas_(correo, aulas, accion) {
  const courseIds = (aulas || [])
    .map(normalizarCourseId_)
    .filter(Boolean)
    .filter((courseId, index, all) => all.indexOf(courseId) === index);

  const result = { ok: 0, omitidos: 0, errores: [] };
  if (!courseIds.length) return result;

  const esEntrada = accion === 'entrada' || accion === 'agregar';
  const requests = courseIds.map(courseId => panelCrearSolicitudAlumnoAula_(correo, courseId, esEntrada));
  let responses;

  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (error) {
    responses = requests.map(request => UrlFetchApp.fetch(request.url, request));
  }

  responses.forEach((response, index) => {
    panelProcesarRespuestaAlumnoAula_(response, courseIds[index], esEntrada, result);
  });

  return result;
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
