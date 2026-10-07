/****************************************************
 * Panel Clases · administración de aulas
 ****************************************************/

const PANEL_CLASES_ESTADOS = ['ACTIVE', 'ARCHIVED', 'PROVISIONED', 'DECLINED', 'SUSPENDED'];

function abrirPanelClases() {
  registrarHojasExtendidas_();
  const html = HtmlService
    .createHtmlOutputFromFile('PanelClases')
    .setTitle('Clases');
  SpreadsheetApp.getUi().showSidebar(html);
}

function obtenerUrlPanelClases() {
  const base = ScriptApp.getService().getUrl();
  if (!base) {
    return {
      ok: false,
      url: '',
      message: 'Todavía no hay despliegue web app activo.'
    };
  }

  return {
    ok: true,
    url: base + (base.indexOf('?') === -1 ? '?' : '&') + 'panel=clases',
    message: 'URL del panel Clases disponible.'
  };
}

function mostrarUrlPanelClases() {
  const info = obtenerUrlPanelClases();
  const ui = SpreadsheetApp.getUi();

  if (!info.ok) {
    ui.alert(
      'Panel Clases no desplegado',
      'El código ya está preparado. Falta crear o actualizar el despliegue web desde Apps Script.',
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert('Panel Clases', info.url, ui.ButtonSet.OK);
}

function panelClasesListar(filtros) {
  filtros = filtros || {};
  const estados = panelClasesEstados_(filtros.estados || filtros.estado);
  const teacherEmail = String(filtros.profesor || '').trim().toLowerCase();
  const texto = panelClasesNormalizar_(filtros.texto || '');
  const seccion = panelClasesNormalizar_(filtros.seccion || '');
  const dueno = panelClasesNormalizar_(filtros.dueno || '');
  const ownerOu = panelClasesNormalizar_(filtros.ownerOu || '');
  const descriptionHeading = panelClasesNormalizar_(filtros.descriptionHeading || '');
  const matchMode = String(filtros.matchMode || 'all').toLowerCase() === 'any' ? 'any' : 'all';

  const paramsBase = {
    pageSize: 100,
    courseStates: estados
  };
  if (teacherEmail) paramsBase.teacherId = teacherEmail;

  const cursos = [];
  let pageToken = '';

  do {
    const params = Object.assign({}, paramsBase);
    if (pageToken) params.pageToken = pageToken;
    const response = panelClasesFetch_('/courses', 'get', null, params);
    (response.courses || []).forEach(course => cursos.push(panelClasesCompactar_(course)));
    pageToken = response.nextPageToken || '';
  } while (pageToken);

  let filtrados = cursos;
  const requiereOwner = !!(dueno || ownerOu);
  if (requiereOwner) {
    filtrados = filtrados.map(curso => {
      const owner = panelClasesObtenerOwner_(curso.ownerId);
      curso.ownerEmail = owner.email;
      curso.ownerOu = owner.orgUnitPath;
      return curso;
    });
  }

  const condiciones = [];
  if (texto) condiciones.push(curso => panelClasesNormalizar_([
    curso.id,
    curso.name,
    curso.section,
    curso.subject,
    curso.descriptionHeading,
    curso.description,
    curso.room,
    curso.ownerId,
    curso.ownerEmail,
    curso.ownerOu
  ].join(' ')).indexOf(texto) !== -1);
  if (seccion) condiciones.push(curso => panelClasesNormalizar_(curso.section).indexOf(seccion) !== -1);
  if (dueno) condiciones.push(curso => panelClasesNormalizar_((curso.ownerEmail || '') + ' ' + (curso.ownerId || '')).indexOf(dueno) !== -1);
  if (ownerOu) condiciones.push(curso => panelClasesNormalizar_(curso.ownerOu || '').indexOf(ownerOu) !== -1);
  if (descriptionHeading) condiciones.push(curso => panelClasesNormalizar_(curso.descriptionHeading).indexOf(descriptionHeading) !== -1);

  if (condiciones.length) {
    filtrados = filtrados.filter(curso => {
      const resultados = condiciones.map(fn => fn(curso));
      return matchMode === 'any' ? resultados.some(Boolean) : resultados.every(Boolean);
    });
  }

  filtrados.sort((a, b) => {
    const secciones = panelClasesNormalizar_(a.section).localeCompare(panelClasesNormalizar_(b.section), 'es', { sensitivity: 'base', numeric: true });
    if (secciones !== 0) return secciones;
    return panelClasesNormalizar_(a.name).localeCompare(panelClasesNormalizar_(b.name), 'es', { sensitivity: 'base', numeric: true });
  });

  return {
    total: filtrados.length,
    cursos: filtrados.slice(0, 500),
    message: filtrados.length + ' clase(s) encontradas.'
  };
}

function panelClasesDetalle(courseId) {
  const id = panelClasesCourseId_(courseId);
  if (!id) throw new Error('Falta el Course ID.');

  const course = panelClasesFetch_('/courses/' + encodeURIComponent(id), 'get');
  const detalle = panelClasesCompactar_(course);
  const owner = panelClasesObtenerOwner_(detalle.ownerId);
  detalle.ownerEmail = owner.email;
  detalle.ownerOu = owner.orgUnitPath;

  return {
    curso: detalle,
    profesores: panelClasesListarMiembros_(id, 'teachers'),
    alumnos: panelClasesListarMiembros_(id, 'students')
  };
}

function panelClasesCrear(datos) {
  datos = datos || {};
  const payload = {};

  if (datos.name) payload.name = String(datos.name).trim();
  if (datos.section) payload.section = String(datos.section).trim();
  if (datos.subject) payload.subject = String(datos.subject).trim();
  if (datos.descriptionHeading) payload.descriptionHeading = String(datos.descriptionHeading).trim();
  if (datos.description) payload.description = String(datos.description).trim();
  if (datos.room) payload.room = String(datos.room).trim();
  if (datos.ownerId) payload.ownerId = String(datos.ownerId).trim().toLowerCase();
  if (datos.courseState) payload.courseState = String(datos.courseState).trim();

  if (!payload.name) throw new Error('Escribe el nombre de la clase.');

  const course = panelClasesFetch_('/courses', 'post', payload);
  registrarLog_('Crear clase desde Panel Clases', course.id || '', course.name || payload.name || 'OK');

  return {
    ok: true,
    curso: panelClasesCompactar_(course),
    message: 'Clase creada: ' + (course.name || payload.name) + '.'
  };
}

function panelClasesActualizar(courseId, cambios) {
  const id = panelClasesCourseId_(courseId);
  if (!id) throw new Error('Falta el Course ID.');

  cambios = cambios || {};
  const camposPermitidos = ['name', 'section', 'subject', 'descriptionHeading', 'description', 'room', 'ownerId'];
  const payload = {};
  const updateMask = [];

  camposPermitidos.forEach(campo => {
    if (!Object.prototype.hasOwnProperty.call(cambios, campo)) return;
    const valor = String(cambios[campo] || '').trim();
    if (valor === '') return;
    payload[campo] = campo === 'ownerId' ? valor.toLowerCase() : valor;
    updateMask.push(campo);
  });

  if (!updateMask.length) throw new Error('No hay cambios para aplicar.');

  const course = panelClasesFetch_('/courses/' + encodeURIComponent(id), 'patch', payload, { updateMask: updateMask.join(',') });
  registrarLog_('Actualizar clase desde Panel Clases', id, updateMask.join(', '));

  return {
    ok: true,
    curso: panelClasesCompactar_(course),
    message: 'Clase actualizada. Campos: ' + updateMask.join(', ') + '.'
  };
}

function panelClasesCambiarEstado(courseId, estado) {
  const id = panelClasesCourseId_(courseId);
  const nuevoEstado = String(estado || '').trim().toUpperCase();
  if (!id) throw new Error('Falta el Course ID.');
  if (['ACTIVE', 'ARCHIVED'].indexOf(nuevoEstado) === -1) throw new Error('Estado inválido. Usa ACTIVE o ARCHIVED.');

  const course = panelClasesFetch_('/courses/' + encodeURIComponent(id), 'patch', { courseState: nuevoEstado }, { updateMask: 'courseState' });
  registrarLog_('Cambiar estado de clase desde Panel Clases', id, nuevoEstado);

  return {
    ok: true,
    curso: panelClasesCompactar_(course),
    message: nuevoEstado === 'ARCHIVED' ? 'Clase archivada.' : 'Clase restaurada como activa.'
  };
}

function panelClasesEliminar(courseId, confirmacion) {
  const id = panelClasesCourseId_(courseId);
  if (!id) throw new Error('Falta el Course ID.');

  const valor = String(confirmacion || '').trim();
  if (valor !== id && valor.toUpperCase() !== 'ELIMINAR') {
    throw new Error('Para eliminar, confirma escribiendo ELIMINAR o el Course ID.');
  }

  panelClasesFetch_('/courses/' + encodeURIComponent(id), 'delete');
  registrarLog_('Eliminar clase desde Panel Clases', id, 'Eliminada');

  return {
    ok: true,
    message: 'Clase eliminada.'
  };
}

function panelClasesAplicarUsuario(courseId, email, tipoUsuario, accion) {
  const id = panelClasesCourseId_(courseId);
  const correo = String(email || '').trim().toLowerCase();
  const tipo = String(tipoUsuario || '').trim();
  const esProfesor = tipo === 'Profesor';
  const esAlumno = tipo === 'Alumno';
  const esEntrada = accion === 'entrada' || accion === 'agregar';

  if (!id) throw new Error('Falta el Course ID.');
  if (!correo) throw new Error('Escribe el correo del usuario.');
  if (!esProfesor && !esAlumno) throw new Error('Tipo inválido.');

  const path = esProfesor ? 'teachers' : 'students';

  if (esEntrada) {
    panelClasesFetch_('/courses/' + encodeURIComponent(id) + '/' + path, 'post', { userId: correo });
  } else {
    panelClasesFetch_('/courses/' + encodeURIComponent(id) + '/' + path + '/' + encodeURIComponent(correo), 'delete');
  }

  registrarLog_((esEntrada ? 'Agregar ' : 'Retirar ') + tipo.toLowerCase() + ' desde Panel Clases', id + ' | ' + correo, 'OK');

  return {
    ok: true,
    message: (esEntrada ? 'Usuario agregado.' : 'Usuario retirado.')
  };
}

function panelClasesSincronizarAulas(courseIds) {
  const ids = (courseIds || []).map(panelClasesCourseId_).filter(Boolean)
    .filter((id, index, all) => all.indexOf(id) === index);

  if (!ids.length) throw new Error('Selecciona al menos una clase.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.CLASSES);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const rows = ids.map(id => panelClasesFilaAulas_(panelClasesFetch_('/courses/' + encodeURIComponent(id), 'get'), headers));

  if (rows.length) {
    const start = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(start, 1, rows.length, headers.length).setValues(rows);
  }

  registrarLog_('Sincronizar clases seleccionadas en AULAS', ids.join(', '), rows.length + ' filas agregadas');

  return {
    ok: true,
    message: rows.length + ' clase(s) agregadas a la hoja AULAS.'
  };
}

function panelClasesListarMiembros_(courseId, tipo) {
  const path = tipo === 'teachers' ? 'teachers' : 'students';
  const out = [];
  let pageToken = '';

  do {
    const params = { pageSize: 100 };
    if (pageToken) params.pageToken = pageToken;
    const response = panelClasesFetch_('/courses/' + encodeURIComponent(courseId) + '/' + path, 'get', null, params);
    const rows = response[path] || [];
    rows.forEach(item => {
      const profile = item.profile || {};
      const name = profile.name || {};
      out.push({
        id: profile.id || item.userId || '',
        email: profile.emailAddress || '',
        fullName: name.fullName || [name.givenName || '', name.familyName || ''].join(' ').trim()
      });
    });
    pageToken = response.nextPageToken || '';
  } while (pageToken);

  return out;
}

function panelClasesFilaAulas_(course, headers) {
  const compacto = panelClasesCompactar_(course);
  return headers.map(header => {
    switch (header) {
      case 'Course ID': return compacto.id;
      case 'Nombre del aula': return compacto.name;
      case 'Sección': return compacto.section;
      case 'Materia': return compacto.subject;
      case 'Salón': return compacto.room;
      case 'Dueño ID': return compacto.ownerId;
      case 'Estado': return compacto.courseState;
      case 'Código de inscripción': return compacto.enrollmentCode;
      case 'Link de Classroom': return compacto.alternateLink;
      case 'Tutores activados': return compacto.guardiansEnabled ? 'Sí' : 'No';
      case 'Fecha de creación': return formatearFecha_(compacto.creationTime);
      case 'Última actualización': return formatearFecha_(compacto.updateTime);
      default: return '';
    }
  });
}

function panelClasesCompactar_(course) {
  course = course || {};
  return {
    id: course.id || '',
    name: course.name || '',
    section: course.section || '',
    subject: course.subject || '',
    descriptionHeading: course.descriptionHeading || '',
    description: course.description || '',
    room: course.room || '',
    ownerId: course.ownerId || '',
    courseState: course.courseState || '',
    enrollmentCode: course.enrollmentCode || '',
    alternateLink: course.alternateLink || '',
    guardiansEnabled: !!course.guardiansEnabled,
    creationTime: course.creationTime || '',
    updateTime: course.updateTime || ''
  };
}

function panelClasesObtenerOwner_(ownerId) {
  const key = String(ownerId || '').trim();
  if (!key) return { email: '', orgUnitPath: '' };

  const cache = CacheService.getScriptCache();
  const cacheKey = 'cm_owner_' + key;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  let result = { email: key, orgUnitPath: '' };
  try {
    const user = AdminDirectory.Users.get(key);
    result = {
      email: user.primaryEmail || key,
      orgUnitPath: user.orgUnitPath || ''
    };
  } catch (error) {}

  cache.put(cacheKey, JSON.stringify(result), 21600);
  return result;
}

function panelClasesEstados_(value) {
  if (Array.isArray(value)) {
    const estados = value.map(v => String(v || '').trim().toUpperCase()).filter(Boolean);
    return estados.length ? estados : ['ACTIVE'];
  }

  const raw = String(value || 'ACTIVE').trim().toUpperCase();
  if (raw === 'ALL' || raw === 'TODAS') return PANEL_CLASES_ESTADOS;
  if (!raw) return ['ACTIVE'];
  return raw.split(',').map(v => v.trim()).filter(Boolean);
}

function panelClasesCourseId_(value) {
  return String(value || '').trim();
}

function panelClasesNormalizar_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function panelClasesFetch_(path, method, payload, params) {
  let url = 'https://classroom.googleapis.com/v1' + path;
  const query = panelClasesQuery_(params || {});
  if (query) url += '?' + query;

  const options = {
    method: method || 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) options.payload = JSON.stringify(payload);

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText() || '';

  if (code >= 200 && code < 300) {
    if (!body) return {};
    try { return JSON.parse(body); }
    catch (error) { return {}; }
  }

  throw new Error('Classroom HTTP ' + code + ': ' + panelClasesMensajeApi_(body));
}

function panelClasesQuery_(params) {
  const pairs = [];
  Object.keys(params || {}).forEach(key => {
    const value = params[key];
    if (value === null || value === undefined || value === '') return;
    if (Array.isArray(value)) {
      value.forEach(item => pairs.push(encodeURIComponent(key) + '=' + encodeURIComponent(item)));
    } else {
      pairs.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
    }
  });
  return pairs.join('&');
}

function panelClasesMensajeApi_(body) {
  if (!body) return '';
  try {
    const parsed = JSON.parse(body);
    if (parsed && parsed.error && parsed.error.message) return parsed.error.message;
  } catch (error) {}
  return String(body).substring(0, 300);
}
