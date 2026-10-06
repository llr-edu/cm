/****************************************************
 * Panel operativo v2 · acciones directas desde barra
 ****************************************************/

function panelMoverAlumnoSeccion(email, nivel, origen, destino) {
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const origenText = normalizarNomenclatura_(origen);
  const destinoText = normalizarNomenclatura_(destino);
  if (!correo) throw new Error('Selecciona primero un alumno.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!origenText || !destinoText) throw new Error('Selecciona origen y destino.');
  if (origenText === destinoText) throw new Error('Origen y destino no pueden ser iguales.');
  if (origenText.substring(0, 3) !== destinoText.substring(0, 3)) throw new Error('El cambio debe ser dentro del mismo grado/año.');

  const aulasOrigen = panelAulasPorSeccion_(nivelText, origenText, false);
  const aulasDestino = panelAulasPorSeccion_(nivelText, destinoText, false);
  if (!aulasOrigen.length) throw new Error('No hay aulas exclusivas para el origen ' + origenText + '.');
  if (!aulasDestino.length) throw new Error('No hay aulas exclusivas para el destino ' + destinoText + '.');

  const salida = panelAplicarAlumnoEnAulas_(correo, aulasOrigen, 'salida');
  const entrada = panelAplicarAlumnoEnAulas_(correo, aulasDestino, 'entrada');
  const errores = salida.errores.concat(entrada.errores);
  const omitidos = salida.omitidos + entrada.omitidos;
  const segundos = Math.round((new Date() - inicio) / 1000);

  registrarLog_('Mover alumno desde panel', correo + ' | ' + nivelText + ' | ' + origenText + ' → ' + destinoText, 'Salidas: ' + salida.ok + ' | Entradas: ' + entrada.ok + ' | Omitidos: ' + omitidos + ' | Errores: ' + errores.length + ' | Tiempo: ' + segundos + ' s');
  if (errores.length) registrarError_('panelMoverAlumnoSeccion', '', correo, new Error(errores.join(' | ')));

  return { ok: errores.length === 0, eliminados: salida.ok, agregados: entrada.ok, omitidos: omitidos, errores: errores, elapsedMs: new Date() - inicio, message: 'Movimiento finalizado. Eliminados: ' + salida.ok + '. Agregados: ' + entrada.ok + '. Ya estaba/no estaba: ' + omitidos + '. Errores: ' + errores.length + '. Tiempo: ' + segundos + ' s.' };
}

function panelRetirarAlumnoDeSeccion(email, nivel, seccion, incluirGenerales) {
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const seccionText = normalizarNomenclatura_(seccion);
  if (!correo) throw new Error('Selecciona primero un alumno.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!seccionText) throw new Error('Selecciona la sección.');

  const aulas = panelAulasPorSeccion_(nivelText, seccionText, incluirGenerales !== false);
  if (!aulas.length) throw new Error('No hay aulas para ' + nivelText + ' / ' + seccionText + '.');

  const resultado = panelAplicarAlumnoEnAulas_(correo, aulas, 'salida');
  const segundos = Math.round((new Date() - inicio) / 1000);
  registrarLog_('Retirar alumno desde panel', correo + ' | ' + nivelText + ' | ' + seccionText, 'Procesadas: ' + resultado.ok + ' | Omitidos: ' + resultado.omitidos + ' | Errores: ' + resultado.errores.length + ' | Tiempo: ' + segundos + ' s');
  if (resultado.errores.length) registrarError_('panelRetirarAlumnoDeSeccion', '', correo, new Error(resultado.errores.join(' | ')));

  return { ok: resultado.errores.length === 0, eliminados: resultado.ok, omitidos: resultado.omitidos, errores: resultado.errores, elapsedMs: new Date() - inicio, message: 'Alumno retirado. Aulas procesadas: ' + resultado.ok + '. No estaba: ' + resultado.omitidos + '. Errores: ' + resultado.errores.length + '. Tiempo: ' + segundos + ' s.' };
}

function panelAulasPorSeccion_(nivel, nomenclatura, incluirGenerales) {
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
  const r = { ok: 0, omitidos: 0, errores: [] };
  aulas.forEach(courseId => {
    try {
      if (accion === 'entrada') agregarAlumno_(courseId, correo);
      else eliminarAlumno_(courseId, correo);
      r.ok++;
    } catch (error) {
      const msg = interpretarErrorClassroom_(error);
      if (/already|ya est|not found|no se encontr|Requested entity was not found/i.test(msg)) r.omitidos++;
      else r.errores.push(courseId + ': ' + msg);
    }
  });
  return r;
}
