/****************************************************
 * Panel operativo · profesores
 ****************************************************/

function panelAgregarProfesorASeccion(email, nivel, seccion, incluirGenerales) {
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const seccionText = normalizarNomenclatura_(seccion);

  if (!correo) throw new Error('Selecciona primero un profesor.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!seccionText) throw new Error('Selecciona la sección.');

  const aulas = panelAulasPorSeccion_(nivelText, seccionText, incluirGenerales !== false);
  if (!aulas.length) throw new Error('No hay aulas para ' + nivelText + ' / ' + seccionText + '.');

  const resultado = panelAplicarProfesorEnAulas_(correo, aulas, 'entrada');
  const segundos = Math.round((new Date() - inicio) / 1000);

  registrarLog_(
    'Agregar profesor desde panel',
    correo + ' | ' + nivelText + ' | ' + seccionText,
    'Aulas: ' + aulas.length + ' | Agregados: ' + resultado.ok + ' | Ya estaba: ' + resultado.omitidos + ' | Errores: ' + resultado.errores.length + ' | Tiempo: ' + segundos + ' s'
  );

  if (resultado.errores.length) {
    registrarError_('panelAgregarProfesorASeccion', '', correo, new Error(resultado.errores.join(' | ')));
  }

  return {
    ok: resultado.errores.length === 0,
    total: aulas.length,
    agregados: resultado.ok,
    omitidos: resultado.omitidos,
    errores: resultado.errores,
    elapsedMs: new Date() - inicio,
    message: 'Profesor agregado. Aulas procesadas: ' + aulas.length + '. Agregados: ' + resultado.ok + '. Ya estaba: ' + resultado.omitidos + '. Errores: ' + resultado.errores.length + '. Tiempo: ' + segundos + ' s.'
  };
}

function panelRetirarProfesorDeSeccion(email, nivel, seccion, incluirGenerales) {
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const seccionText = normalizarNomenclatura_(seccion);

  if (!correo) throw new Error('Selecciona primero un profesor.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!seccionText) throw new Error('Selecciona la sección.');

  const aulas = panelAulasPorSeccion_(nivelText, seccionText, incluirGenerales !== false);
  if (!aulas.length) throw new Error('No hay aulas para ' + nivelText + ' / ' + seccionText + '.');

  const resultado = panelAplicarProfesorEnAulas_(correo, aulas, 'salida');
  const segundos = Math.round((new Date() - inicio) / 1000);

  registrarLog_(
    'Retirar profesor desde panel',
    correo + ' | ' + nivelText + ' | ' + seccionText,
    'Aulas: ' + aulas.length + ' | Retirados: ' + resultado.ok + ' | No estaba: ' + resultado.omitidos + ' | Errores: ' + resultado.errores.length + ' | Tiempo: ' + segundos + ' s'
  );

  if (resultado.errores.length) {
    registrarError_('panelRetirarProfesorDeSeccion', '', correo, new Error(resultado.errores.join(' | ')));
  }

  return {
    ok: resultado.errores.length === 0,
    total: aulas.length,
    eliminados: resultado.ok,
    omitidos: resultado.omitidos,
    errores: resultado.errores,
    elapsedMs: new Date() - inicio,
    message: 'Profesor retirado. Aulas procesadas: ' + aulas.length + '. Retirados: ' + resultado.ok + '. No estaba: ' + resultado.omitidos + '. Errores: ' + resultado.errores.length + '. Tiempo: ' + segundos + ' s.'
  };
}

function panelAplicarProfesorEnAulas_(correo, aulas, accion) {
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
        url: 'https://classroom.googleapis.com/v1/courses/' + encodeURIComponent(courseId) + '/teachers',
        method: 'post',
        headers: headers,
        payload: JSON.stringify({ userId: correo }),
        muteHttpExceptions: true
      };
    }

    return {
      url: 'https://classroom.googleapis.com/v1/courses/' + encodeURIComponent(courseId) + '/teachers/' + encodeURIComponent(correo),
      method: 'delete',
      headers: headers,
      muteHttpExceptions: true
    };
  });

  let responses;
  try {
    responses = UrlFetchApp.fetchAll(requests);
  } catch (error) {
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
