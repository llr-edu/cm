/****************************************************
 * Panel operativo v2 · acciones directas desde barra
 ****************************************************/

function panelBuscarUsuarios(texto) {
  const query = normalizarTextoPanel_(texto);
  if (!query || query.length < 2) return [];

  const raw = String(texto || '').trim().toLowerCase();
  const exactEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(raw);
  if (exactEmail) {
    try {
      const user = AdminDirectory.Users.get(raw);
      return [panelSerializarUsuarioV2_(user, 'Admin directo')];
    } catch (error) {
      // Continúa con caché y búsqueda amplia.
    }
  }

  const locales = panelBuscarUsuariosCacheV2_(query, 20);
  if (locales.length >= 8 || query.length < 4) return locales;

  const admin = panelBuscarUsuariosAdminV2_(query, 20 - locales.length);
  const vistos = {};
  return locales.concat(admin).filter(user => {
    const email = String(user.email || '').toLowerCase();
    if (!email || vistos[email]) return false;
    vistos[email] = true;
    return true;
  }).slice(0, 20);
}

function panelListarAulasUsuario(email, tipo) {
  const correo = String(email || '').trim().toLowerCase();
  const modo = String(tipo || 'Alumno').trim().toLowerCase();
  if (!correo) throw new Error('Falta el correo del usuario.');

  const esAlumno = modo === 'alumno' || modo === 'student';
  const esProfesor = modo === 'profesor' || modo === 'teacher';
  if (!esAlumno && !esProfesor) throw new Error('Tipo inválido. Usa Alumno o Profesor.');

  const catalogo = panelMapaCatalogoPorCourseIdV2_();
  const rows = [];
  let pageToken = null;

  do {
    const params = { pageSize: 100, pageToken: pageToken, courseStates: ['ACTIVE'] };
    if (esAlumno) params.studentId = correo;
    if (esProfesor) params.teacherId = correo;
    const response = Classroom.Courses.list(params);
    (response.courses || []).forEach(course => {
      const id = normalizarCourseId_(course.id || '');
      const cat = catalogo[id] || {};
      rows.push({
        id: id,
        name: course.name || '',
        section: course.section || '',
        subject: course.subject || '',
        state: course.courseState || '',
        link: course.alternateLink || '',
        ownerId: course.ownerId || '',
        nivel: cat.nivel || '',
        nomenclatura: cat.nomenclatura || ''
      });
    });
    pageToken = response.nextPageToken;
  } while (pageToken);

  rows.sort((a, b) => (a.name + ' ' + a.section).localeCompare(b.name + ' ' + b.section, 'es', { sensitivity: 'base' }));
  return { aulas: rows, deteccion: panelDetectarSeccionDesdeAulasV2_(rows) };
}

function panelAgregarAlumnoASeccion(email, nivel, seccion) {
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const seccionText = normalizarNomenclatura_(seccion);
  if (!correo) throw new Error('Selecciona primero un alumno.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!seccionText) throw new Error('Selecciona la sección destino.');

  const aulas = panelAulasPorSeccion_(nivelText, seccionText, true);
  if (!aulas.length) throw new Error('No hay aulas para ' + nivelText + ' / ' + seccionText + '.');

  const resultado = panelAplicarAlumnoEnAulas_(correo, aulas, 'entrada');
  const segundos = Math.round((new Date() - inicio) / 1000);
  registrarLog_('Agregar alumno desde panel', correo + ' | ' + nivelText + ' | ' + seccionText, 'Procesadas: ' + resultado.ok + ' | Omitidos: ' + resultado.omitidos + ' | Errores: ' + resultado.errores.length + ' | Tiempo: ' + segundos + ' s');
  if (resultado.errores.length) registrarError_('panelAgregarAlumnoASeccion', '', correo, new Error(resultado.errores.join(' | ')));

  return { ok: resultado.errores.length === 0, agregados: resultado.ok, omitidos: resultado.omitidos, errores: resultado.errores, elapsedMs: new Date() - inicio, message: 'Alumno agregado. Aulas procesadas: ' + resultado.ok + '. Ya estaba: ' + resultado.omitidos + '. Errores: ' + resultado.errores.length + '. Tiempo: ' + segundos + ' s.' };
}

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

function panelBuscarUsuariosCacheV2_(query, limit) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.USERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colNombre = buscarColumna_(headers, 'Nombre completo');
  const colCorreo = buscarColumna_(headers, 'Correo');
  const colOu = buscarColumna_(headers, 'Unidad organizativa');
  const colSuspendido = buscarColumna_(headers, 'Suspendido');
  const colArchivado = buscarColumna_(headers, 'Archivado');
  const out = [];
  data.slice(1).forEach(row => {
    if (out.length >= limit) return;
    const fullName = String(row[colNombre] || '');
    const email = String(row[colCorreo] || '');
    const orgUnitPath = String(row[colOu] || '');
    if (!normalizarTextoPanel_(fullName + ' ' + email + ' ' + orgUnitPath).includes(query)) return;
    out.push({ fullName: fullName, givenName: '', familyName: '', email: email, orgUnitPath: orgUnitPath, suspended: row[colSuspendido] === true || String(row[colSuspendido]).toUpperCase() === 'TRUE', archived: row[colArchivado] === true || String(row[colArchivado]).toUpperCase() === 'TRUE', source: 'Caché' });
  });
  return out;
}

function panelBuscarUsuariosAdminV2_(query, limit) {
  const domain = obtenerConfig_('Dominio') || 'losroblesenlinea.com.ve';
  const out = [];
  let pageToken = null;
  do {
    const response = AdminDirectory.Users.list({ domain: domain, maxResults: 500, pageToken: pageToken, orderBy: 'email' });
    (response.users || []).forEach(user => {
      if (out.length >= limit) return;
      const serial = panelSerializarUsuarioV2_(user, 'Admin');
      const haystack = normalizarTextoPanel_(serial.fullName + ' ' + serial.givenName + ' ' + serial.familyName + ' ' + serial.email + ' ' + serial.orgUnitPath);
      if (haystack.includes(query)) out.push(serial);
    });
    pageToken = response.nextPageToken;
  } while (pageToken && out.length < limit);
  return out;
}

function panelSerializarUsuarioV2_(user, source) {
  user = user || {};
  const name = user.name || {};
  return { fullName: name.fullName || '', givenName: name.givenName || '', familyName: name.familyName || '', email: user.primaryEmail || '', orgUnitPath: user.orgUnitPath || '', suspended: user.suspended === true, archived: user.archived === true, source: source || '' };
}

function panelMapaCatalogoPorCourseIdV2_() {
  registrarHojasExtendidas_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CATALOG);
  const mapa = {};
  if (!sheet || sheet.getLastRow() < 2) return mapa;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colNomenclatura = buscarColumna_(headers, 'Nomenclatura');
  const colCourseId = buscarColumna_(headers, 'Course ID');
  data.slice(1).forEach(row => {
    const courseId = normalizarCourseId_(row[colCourseId]);
    if (!courseId) return;
    mapa[courseId] = { nivel: String(row[colNivel] || '').trim(), nomenclatura: normalizarNomenclatura_(row[colNomenclatura]) };
  });
  return mapa;
}

function panelDetectarSeccionDesdeAulasV2_(aulas) {
  const conteo = {};
  (aulas || []).forEach(aula => {
    const nom = normalizarNomenclatura_(aula.nomenclatura);
    if (!aula.nivel || !/^\d{3}[A-Z]$/.test(nom) || /ABC$/.test(nom)) return;
    const key = aula.nivel + '|' + nom;
    conteo[key] = (conteo[key] || 0) + 1;
  });
  const keys = Object.keys(conteo).sort((a, b) => conteo[b] - conteo[a]);
  if (!keys.length) return null;
  const parts = keys[0].split('|');
  return { nivel: parts[0], nomenclatura: parts[1], coincidencias: conteo[keys[0]] };
}
