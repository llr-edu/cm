/****************************************************
 * Panel operativo
 * Classroom Manager LL. RR. · búsqueda y acciones rápidas
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
    .addItem('Aplicar desplegables de usuarios', 'aplicarDesplegablesUsuarios')
    .addSeparator()
    .addItem('Ver LOG', 'irALog')
    .addItem('Ver ERRORES', 'irAErrores')
    .addToUi();
}

function abrirPanelOperativo() {
  registrarHojasExtendidas_();
  const html = HtmlService
    .createHtmlOutputFromFile('PanelOperativo')
    .setTitle('Panel operativo');

  SpreadsheetApp.getUi().showSidebar(html);
}

function panelBuscarUsuarios(texto) {
  const query = normalizarTextoPanel_(texto);
  if (!query || query.length < 2) return [];

  const domain = obtenerConfig_('Dominio') || 'losroblesenlinea.com.ve';
  const resultados = [];
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
      if (resultados.length >= 30) return;

      const fullName = user.name && user.name.fullName ? user.name.fullName : '';
      const givenName = user.name && user.name.givenName ? user.name.givenName : '';
      const familyName = user.name && user.name.familyName ? user.name.familyName : '';
      const email = user.primaryEmail || '';
      const orgUnitPath = user.orgUnitPath || '';
      const haystack = normalizarTextoPanel_(fullName + ' ' + givenName + ' ' + familyName + ' ' + email + ' ' + orgUnitPath);

      if (haystack.includes(query)) {
        resultados.push({
          fullName: fullName,
          givenName: givenName,
          familyName: familyName,
          email: email,
          orgUnitPath: orgUnitPath,
          suspended: user.suspended === true,
          archived: user.archived === true
        });
      }
    });

    pageToken = response.nextPageToken;
  } while (pageToken && resultados.length < 30);

  return resultados;
}

function panelListarAulasUsuario(email, tipo) {
  const correo = String(email || '').trim().toLowerCase();
  const modo = String(tipo || 'Alumno').trim().toLowerCase();

  if (!correo) throw new Error('Falta el correo del usuario.');

  const esAlumno = modo === 'alumno' || modo === 'student';
  const esProfesor = modo === 'profesor' || modo === 'teacher';
  if (!esAlumno && !esProfesor) throw new Error('Tipo inválido. Usa Alumno o Profesor.');

  const rows = [];
  let pageToken = null;

  do {
    const params = {
      pageSize: 100,
      pageToken: pageToken,
      courseStates: ['ACTIVE']
    };

    if (esAlumno) params.studentId = correo;
    if (esProfesor) params.teacherId = correo;

    const response = Classroom.Courses.list(params);
    const courses = response.courses || [];

    courses.forEach(course => {
      rows.push({
        id: course.id || '',
        name: course.name || '',
        section: course.section || '',
        subject: course.subject || '',
        state: course.courseState || '',
        link: course.alternateLink || '',
        ownerId: course.ownerId || ''
      });
    });

    pageToken = response.nextPageToken;
  } while (pageToken);

  rows.sort((a, b) => (a.name + ' ' + a.section).localeCompare(b.name + ' ' + b.section, 'es', { sensitivity: 'base' }));
  return rows;
}

function panelObtenerSecciones() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.CATALOG);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colNomenclatura = buscarColumna_(headers, 'Nomenclatura');
  const mapa = {};

  data.slice(1).forEach(row => {
    const nivel = String(row[colNivel] || '').trim();
    const nomenclatura = normalizarNomenclatura_(row[colNomenclatura]);
    if (!nivel || !/^\d{3}[A-Z]$/.test(nomenclatura)) return;
    if (/ABC$/.test(nomenclatura)) return;

    const key = nivel + '|' + nomenclatura;
    mapa[key] = { nivel: nivel, nomenclatura: nomenclatura, prefijo: nomenclatura.substring(0, 3), seccion: nomenclatura.substring(3) };
  });

  return Object.keys(mapa)
    .map(key => mapa[key])
    .sort((a, b) => (a.nivel + a.nomenclatura).localeCompare(b.nivel + b.nomenclatura));
}

function panelPrepararMoverAlumno(email, nivel, origen, destino) {
  registrarHojasExtendidas_();
  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const origenText = normalizarNomenclatura_(origen);
  const destinoText = normalizarNomenclatura_(destino);

  if (!correo) throw new Error('Selecciona primero un alumno.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!origenText || !destinoText) throw new Error('Selecciona origen y destino.');
  if (origenText === destinoText) throw new Error('Origen y destino no pueden ser iguales.');
  if (origenText.substring(0, 3) !== destinoText.substring(0, 3)) throw new Error('El cambio debe ser dentro del mismo grado/año.');

  validarUsuarioAdmin_(correo);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.MOVE_STUDENTS);
  const row = Math.max(sheet.getLastRow() + 1, 2);

  sheet.getRange(row, 1, 1, 10).setValues([[
    correo,
    nivelText,
    origenText,
    destinoText,
    'Y',
    '',
    '',
    '',
    '',
    ''
  ]]);

  ss.setActiveSheet(sheet);
  sheet.getRange(row, 1).activate();
  registrarLog_('Preparar mover alumno', correo + ' | ' + nivelText + ' | ' + origenText + ' → ' + destinoText, 'Fila ' + row);

  return {
    ok: true,
    row: row,
    message: 'Movimiento preparado en MOVER_ALUMNOS, fila ' + row + '. Ejecuta “Procesar mover alumnos” para aplicar el cambio.'
  };
}

function normalizarTextoPanel_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}
