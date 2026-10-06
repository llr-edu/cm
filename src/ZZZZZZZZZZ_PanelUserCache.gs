/****************************************************
 * Panel operativo · caché local de usuarios
 ****************************************************/

function onOpen() {
  registrarHojasExtendidas_();
  SpreadsheetApp.getUi()
    .createMenu(APP_NAME)
    .addItem('Panel operativo', 'abrirPanelOperativo')
    .addSeparator()
    .addItem('Start / Setup', 'startSetup')
    .addSeparator()
    .addItem('Sincronizar catálogo de aulas', 'sincronizarCatalogoAulas')
    .addItem('Buscar usuario en Admin', 'buscarUsuariosAdmin')
    .addItem('Procesar mover alumnos', 'procesarMoverAlumnos')
    .addSeparator()
    .addItem('Listar aulas activas', 'listarAulasActivas')
    .addItem('Listar todas las aulas', 'listarTodasLasAulas')
    .addSeparator()
    .addItem('Procesar alumnos', 'procesarAlumnos')
    .addItem('Procesar profesores', 'procesarProfesores')
    .addSeparator()
    .addItem('Consultar usuario - aulas activas', 'consultarUsuarioAulasActivas')
    .addItem('Consultar usuario - todas las aulas', 'consultarUsuarioTodasLasAulas')
    .addSeparator()
    .addItem('Actualizar lista de usuarios', 'actualizarListaUsuarios')
    .addItem('Programar actualización nocturna', 'instalarActualizacionNocturnaUsuarios')
    .addItem('Aplicar desplegables de usuarios', 'aplicarDesplegablesUsuarios')
    .addSeparator()
    .addItem('Ver LOG', 'irALog')
    .addItem('Ver ERRORES', 'irAErrores')
    .addToUi();
}

function abrirPanelOperativo() {
  registrarHojasExtendidas_();
  asegurarActualizacionNocturnaUsuarios_();
  const html = HtmlService
    .createHtmlOutputFromFile('PanelOperativo')
    .setTitle('Panel operativo');

  SpreadsheetApp.getUi().showSidebar(html);
}

function panelObtenerUsuariosCache() {
  registrarHojasExtendidas_();
  const usuarios = obtenerUsuariosDesdeHoja_();
  return {
    usuarios: usuarios,
    total: usuarios.length
  };
}

function panelBuscarUsuarios(texto) {
  const query = normalizarTextoPanel_(texto);
  if (!query || query.length < 2) return [];

  const resultados = filtrarUsuariosCache_(query, 30);
  const exacto = esCorreoPanel_(String(texto || '').trim());

  if (resultados.length || !exacto) return resultados;

  try {
    const user = AdminDirectory.Users.get(String(texto || '').trim().toLowerCase());
    return [usuarioAdminAObjeto_(user, 'Admin')];
  } catch (error) {
    return [];
  }
}

function filtrarUsuariosCache_(query, limite) {
  const terms = query.split(' ').filter(Boolean);
  const rows = obtenerUsuariosDesdeHoja_();
  const resultados = [];

  rows.forEach(user => {
    if (resultados.length >= limite) return;
    const haystack = normalizarTextoPanel_(
      (user.fullName || '') + ' ' +
      (user.givenName || '') + ' ' +
      (user.familyName || '') + ' ' +
      (user.email || '') + ' ' +
      (user.orgUnitPath || '')
    );

    if (terms.every(term => haystack.includes(term))) resultados.push(user);
  });

  return resultados;
}

function obtenerUsuariosDesdeHoja_() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const colFullName = buscarColumna_(headers, 'Nombre completo');
  const colGivenName = buscarColumna_(headers, 'Nombre');
  const colFamilyName = buscarColumna_(headers, 'Apellido');
  const colEmail = buscarColumna_(headers, 'Correo');
  const colOu = buscarColumna_(headers, 'Unidad organizativa');
  const colSuspended = buscarColumna_(headers, 'Suspendido');
  const colArchived = buscarColumna_(headers, 'Archivado');

  return data.slice(1)
    .map(row => ({
      fullName: String(row[colFullName] || '').trim(),
      givenName: String(row[colGivenName] || '').trim(),
      familyName: String(row[colFamilyName] || '').trim(),
      email: String(row[colEmail] || '').trim().toLowerCase(),
      orgUnitPath: String(row[colOu] || '').trim(),
      suspended: esSi_(row[colSuspended]),
      archived: esSi_(row[colArchived]),
      source: 'Local'
    }))
    .filter(user => user.email);
}

function usuarioAdminAObjeto_(user, source) {
  const name = user && user.name ? user.name : {};
  return {
    fullName: name.fullName || '',
    givenName: name.givenName || '',
    familyName: name.familyName || '',
    email: String(user.primaryEmail || '').trim().toLowerCase(),
    orgUnitPath: user.orgUnitPath || '',
    suspended: user.suspended === true,
    archived: user.archived === true,
    source: source || 'Admin'
  };
}

function actualizarListaUsuarios() {
  const total = actualizarListaUsuariosCache_(true);
  SpreadsheetApp.getUi().alert('Lista de usuarios actualizada.\n\nUsuarios encontrados: ' + total);
}

function actualizarListaUsuariosNocturno() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    actualizarListaUsuariosCache_(false);
  } finally {
    lock.releaseLock();
  }
}

function actualizarListaUsuariosCache_(mostrarErrores) {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.USERS);

  try {
    limpiarDatosManteniendoEncabezado_(sheet);

    const domain = obtenerConfig_('Dominio') || 'losroblesenlinea.com.ve';
    const rows = [];
    let pageToken = null;

    do {
      const response = AdminDirectory.Users.list({
        domain: domain,
        maxResults: 500,
        pageToken: pageToken,
        orderBy: 'email'
      });

      const users = response.users || [];
      users.forEach(user => {
        rows.push([
          user.name && user.name.fullName ? user.name.fullName : '',
          user.name && user.name.givenName ? user.name.givenName : '',
          user.name && user.name.familyName ? user.name.familyName : '',
          user.primaryEmail || '',
          user.orgUnitPath || '',
          user.suspended === true ? 'Sí' : 'No',
          user.archived === true ? 'Sí' : 'No',
          user.id || '',
          formatearFecha_(user.lastLoginTime),
          formatearFecha_(user.creationTime)
        ]);
      });

      pageToken = response.nextPageToken;
    } while (pageToken);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      sheet.autoResizeColumns(1, sheet.getLastColumn());
    }

    aplicarDesplegablesUsuarios();
    registrarLog_('Actualizar lista de usuarios', 'Dominio: ' + domain, rows.length + ' usuarios listados');
    return rows.length;
  } catch (error) {
    registrarError_('actualizarListaUsuariosCache_', '', '', error);
    if (mostrarErrores) SpreadsheetApp.getUi().alert('No se pudo actualizar la lista de usuarios.\n\nRevisa la hoja ERRORES.');
    throw error;
  }
}

function instalarActualizacionNocturnaUsuarios() {
  asegurarActualizacionNocturnaUsuarios_(true);
  SpreadsheetApp.getUi().alert('Actualización nocturna programada.\n\nLa hoja USUARIOS se actualizará diariamente entre 12:00 a. m. y 1:00 a. m.');
}

function asegurarActualizacionNocturnaUsuarios_(forzar) {
  const handler = 'actualizarListaUsuariosNocturno';
  const triggers = ScriptApp.getProjectTriggers();
  const existentes = triggers.filter(trigger => trigger.getHandlerFunction() === handler);

  if (existentes.length && !forzar) return;

  existentes.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();
}

function esCorreoPanel_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}
