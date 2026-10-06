/****************************************************
 * Panel operativo v3 · llamadas concurrentes a Classroom
 ****************************************************/

function panelAgregarAlumnoASeccion(email, nivel, seccion) {
  registrarHojasExtendidas_();
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const seccionText = normalizarNomenclatura_(seccion);

  if (!correo) throw new Error('Selecciona primero un alumno.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!seccionText) throw new Error('Selecciona la sección destino.');

  const aulas = panelAulasPorSeccion_(nivelText, seccionText, true);
  if (!aulas.length) {
    throw new Error('No hay aulas en CATÁLOGO_AULAS para ' + nivelText + ' / ' + seccionText + '.');
  }

  const resultado = panelAplicarAlumnoEnAulas_(correo, aulas, 'entrada');
  const segundos = Math.round((new Date() - inicio) / 1000);

  registrarLog_(
    'Agregar alumno desde panel',
    correo + ' | ' + nivelText + ' | ' + seccionText,
    'Aulas: ' + aulas.length + ' | Agregados: ' + resultado.ok + ' | Ya estaba: ' + resultado.omitidos + ' | Errores: ' + resultado.errores.length + ' | Tiempo: ' + segundos + ' s'
  );

  if (resultado.errores.length) {
    registrarError_('panelAgregarAlumnoASeccion', '', correo, new Error(resultado.errores.join(' | ')));
  }

  return {
    ok: resultado.errores.length === 0,
    total: aulas.length,
    agregados: resultado.ok,
    omitidos: resultado.omitidos,
    errores: resultado.errores,
    elapsedMs: new Date() - inicio,
    message: 'Alumno agregado. Aulas procesadas: ' + aulas.length + '. Agregados: ' + resultado.ok + '. Ya estaba inscrito: ' + resultado.omitidos + '. Errores: ' + resultado.errores.length + '. Tiempo: ' + segundos + ' s.'
  };
}

function panelAplicarAlumnoEnAulas_(correo, aulas, accion) {
  const courseIds = (aulas || []).map(normalizarCourseId_).filter(Boolean)
    .filter((courseId, index, all) => all.indexOf(courseId) === index);
  const result = { ok: 0, omitidos: 0, errores: [] };
  if (!courseIds.length) return result;

  const token = ScriptApp.getOAuthToken();
  const headers = {
    Authorization: 'Bearer ' + token,
    'Content-Type': 'application/json'
  };
  const esEntrada = accion === 'entrada' || accion === 'agregar';

  const requests = courseIds.map(courseId => {
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
  });

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (error) {
    // Fallback conservador si Google limita las llamadas concurrentes.
    responses = requests.map(request => UrlFetchApp.fetch(request.url, request));
  }

  responses.forEach((response, index) => {
    const code = response.getResponseCode();
    const courseId = courseIds[index];
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

    result.errores.push(courseId + ': HTTP ' + code + ' ' + extraerMensajeApi_(body));
  });

  return result;
}

function extraerMensajeApi_(body) {
  if (!body) return '';
  try {
    const parsed = JSON.parse(body);
    if (parsed && parsed.error && parsed.error.message) return parsed.error.message;
  } catch (e) {}
  return String(body).substring(0, 240);
}
