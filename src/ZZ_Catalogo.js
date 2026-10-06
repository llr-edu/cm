/****************************************************
 * Extensiones de catálogo institucional y operaciones diarias
 * Classroom Manager LL. RR. · v2.0
 ****************************************************/

const FUENTES_AULAS_DEFAULT = [
  {
    nivel: 'E.B.P.',
    spreadsheetId: '17zoAyXtTcs-O_Jkv6tKRrfKkFNeCqXST2mvaFHAbAyA',
    hoja: 'Códs. Grales.',
    gid: '1732664850',
    nomenclaturas: '011,012,013,024,025,026'
  },
  {
    nivel: 'E.M.G.',
    spreadsheetId: '1jPImSlvXHxQSFCPcDSCHgHn_ecpwjYAxr72L018wqcs',
    hoja: 'Códs. Grales.',
    gid: '1489727497',
    nomenclaturas: '067,068,069,071,072'
  }
];

function registrarHojasExtendidas_() {
  SHEETS.SOURCES = SHEETS.SOURCES || 'FUENTES';
  SHEETS.CATALOG = SHEETS.CATALOG || 'CATÁLOGO_AULAS';
  SHEETS.MOVE_STUDENTS = SHEETS.MOVE_STUDENTS || 'MOVER_ALUMNOS';
  SHEETS.ADMIN_SEARCH = SHEETS.ADMIN_SEARCH || 'BUSCAR_USUARIO';
}

function onOpen() {
  registrarHojasExtendidas_();
  SpreadsheetApp.getUi()
    .createMenu(APP_NAME)
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

function startSetup() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  crearHojaConfiguracion_(ss);
  crearHojaAulas_(ss);
  crearHojaAlumnos_(ss);
  crearHojaProfesores_(ss);
  crearHojaConsultaUsuario_(ss);
  crearHojaUsuarios_(ss);
  crearHojaFuentes_(ss);
  crearHojaCatalogoAulas_(ss);
  crearHojaMoverAlumnos_(ss);
  crearHojaBuscarUsuario_(ss);
  crearHojaLog_(ss);
  crearHojaErrores_(ss);

  inicializarConfiguracion_(ss);
  aplicarDesplegablesBasicos_();
  aplicarDesplegablesUsuarios();
  aplicarDesplegablesOperativos_();

  registrarLog_('Start / Setup', 'Estructura inicial creada o verificada con catálogo institucional.', 'OK');

  SpreadsheetApp.getUi().alert(
    'Setup completado.\n\n' +
    'Se crearon o verificaron las hojas base, FUENTES, CATÁLOGO_AULAS, MOVER_ALUMNOS y BUSCAR_USUARIO.\n\n' +
    'Siguiente paso: ejecuta Sincronizar catálogo de aulas.'
  );
}

function crearHojaFuentes_(ss) {
  registrarHojasExtendidas_();
  const headers = [
    'Nivel',
    'Spreadsheet ID',
    'Pestaña',
    'GID',
    'Col código',
    'Col nombre',
    'Col descripción',
    'Col nomenclatura',
    'Nomenclaturas válidas',
    'Activo'
  ];

  let sheet = ss.getSheetByName(SHEETS.SOURCES);
  if (!sheet) sheet = ss.insertSheet(SHEETS.SOURCES);

  const debeInicializar = sheet.getLastRow() === 0 || String(sheet.getRange(1, 1).getValue() || '').trim() !== 'Nivel';
  if (debeInicializar) {
    resetearHojaCompleta_(sheet);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    const rows = FUENTES_AULAS_DEFAULT.map(fuente => [
      fuente.nivel,
      fuente.spreadsheetId,
      fuente.hoja,
      fuente.gid,
      'A',
      'C',
      'D',
      'G',
      fuente.nomenclaturas,
      'Sí'
    ]);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    sheet.setFrozenRows(1);
    aplicarFormatoBasico_(sheet, headers.length);
  }
}

function crearHojaCatalogoAulas_(ss) {
  registrarHojasExtendidas_();
  crearHojaConEncabezados_(ss, SHEETS.CATALOG, [
    'Nivel',
    'Nomenclatura',
    'Course ID',
    'Nombre fuente',
    'Descripción fuente',
    'Nombre Classroom',
    'Sección Classroom',
    'Materia Classroom',
    'Estado Classroom',
    'Link Classroom',
    'Fuente',
    'Última sincronización'
  ]);
}

function crearHojaMoverAlumnos_(ss) {
  registrarHojasExtendidas_();
  crearHojaConEncabezados_(ss, SHEETS.MOVE_STUDENTS, [
    'Alumno',
    'Nivel',
    'Nomenclatura origen',
    'Nomenclatura destino',
    'Ejecutar',
    'Fecha',
    'Agregados',
    'Eliminados',
    'Resultado',
    'Error'
  ]);
}

function crearHojaBuscarUsuario_(ss) {
  registrarHojasExtendidas_();
  let sheet = ss.getSheetByName(SHEETS.ADMIN_SEARCH);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ADMIN_SEARCH);
  resetearHojaCompleta_(sheet);

  sheet.getRange('A1').setValue('Búsqueda');
  sheet.getRange('B1').setValue('');
  sheet.getRange('D1').setValue('Resultados');
  sheet.getRange('E1').setValue('');
  sheet.getRange('A1:E1').setFontWeight('bold');
  sheet.getRange('A1:B1').setBackground(COLORS.LIGHT_YELLOW);
  sheet.getRange('D1:E1').setBackground(COLORS.LIGHT_GREEN);

  const headers = ['Nombre completo', 'Correo', 'Unidad organizativa', 'Suspendido', 'ID de usuario'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(COLORS.RED)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(3);
  for (let i = 1; i <= headers.length; i++) sheet.setColumnWidth(i, 180);
}

function aplicarDesplegablesOperativos_() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const move = ss.getSheetByName(SHEETS.MOVE_STUDENTS);
  if (!move) return;

  const nivelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['E.B.P.', 'E.M.G.'], true)
    .setAllowInvalid(false)
    .build();
  move.getRange('B2:B').setDataValidation(nivelRule);

  const ejecutarRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Y', 'Sí', 'No'], true)
    .setAllowInvalid(true)
    .build();
  move.getRange('E2:E').setDataValidation(ejecutarRule);
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

  fuentes.forEach(fuente => {
    try {
      const sourceSs = SpreadsheetApp.openById(fuente.spreadsheetId);
      const sourceSheet = sourceSs.getSheetByName(fuente.pestana);
      if (!sourceSheet) throw new Error('No existe la pestaña: ' + fuente.pestana);

      const lastRow = sourceSheet.getLastRow();
      if (lastRow < 2) return;

      const data = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
      const validas = fuente.nomenclaturasValidas;

      data.forEach((row, index) => {
        const courseId = String(row[fuente.colCodigo - 1] || '').trim();
        const nombre = String(row[fuente.colNombre - 1] || '').trim();
        const descripcion = String(row[fuente.colDescripcion - 1] || '').trim();
        const nomenclatura = String(row[fuente.colNomenclatura - 1] || '').trim();

        if (!courseId || !nomenclatura) return;
        if (validas.length && !validas.includes(nomenclatura)) return;

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

  registrarLog_('Sincronizar catálogo de aulas', 'Fuentes: ' + fuentes.length, rows.length + ' aulas sincronizadas; errores: ' + errores.length);

  SpreadsheetApp.getUi().alert(
    'Sincronización finalizada.\n\n' +
    'Aulas agregadas al catálogo: ' + rows.length + '\n' +
    'Errores: ' + errores.length +
    (errores.length ? '\n\nPrimer error:\n' + errores[0] : '')
  );
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
        .map(item => item.trim())
        .filter(Boolean)
    }))
    .filter(fuente => fuente.nivel && fuente.spreadsheetId && fuente.pestana);
}

function procesarMoverAlumnos() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.MOVE_STUDENTS);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No hay datos para procesar.');
    return;
  }

  const headers = data[0];
  const colAlumno = buscarColumna_(headers, 'Alumno');
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colOrigen = buscarColumna_(headers, 'Nomenclatura origen');
  const colDestino = buscarColumna_(headers, 'Nomenclatura destino');
  const colEjecutar = buscarColumna_(headers, 'Ejecutar');
  const colFecha = buscarColumna_(headers, 'Fecha');
  const colAgregados = buscarColumna_(headers, 'Agregados');
  const colEliminados = buscarColumna_(headers, 'Eliminados');
  const colResultado = buscarColumna_(headers, 'Resultado');
  const colError = buscarColumna_(headers, 'Error');

  let procesados = 0;
  let errores = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!esSi_(row[colEjecutar])) continue;

    const rowNumber = i + 1;
    const alumno = String(row[colAlumno] || '').trim().toLowerCase();
    const nivel = String(row[colNivel] || '').trim();
    const origen = String(row[colOrigen] || '').trim();
    const destino = String(row[colDestino] || '').trim();

    sheet.getRange(rowNumber, colFecha + 1).setValue(new Date());
    sheet.getRange(rowNumber, colAgregados + 1, 1, 4).clearContent();

    try {
      if (!alumno) throw new Error('Falta el correo del alumno.');
      if (!nivel) throw new Error('Falta el nivel.');
      if (!origen) throw new Error('Falta la nomenclatura de origen.');
      if (!destino) throw new Error('Falta la nomenclatura de destino.');
      if (origen === destino) throw new Error('La nomenclatura de origen y destino no pueden ser iguales.');

      validarUsuarioAdmin_(alumno);

      const aulasOrigen = obtenerAulasCatalogo_(nivel, origen);
      const aulasDestino = obtenerAulasCatalogo_(nivel, destino);

      if (aulasOrigen.length === 0) throw new Error('No hay aulas de origen en CATÁLOGO_AULAS para ' + nivel + ' / ' + origen + '.');
      if (aulasDestino.length === 0) throw new Error('No hay aulas de destino en CATÁLOGO_AULAS para ' + nivel + ' / ' + destino + '.');

      const resultadoAgregar = ejecutarCambioAlumno_(alumno, aulasDestino, 'agregar');
      const resultadoEliminar = ejecutarCambioAlumno_(alumno, aulasOrigen, 'eliminar');

      const erroresOperacion = resultadoAgregar.errores.concat(resultadoEliminar.errores);
      sheet.getRange(rowNumber, colAgregados + 1).setValue(resultadoAgregar.ok);
      sheet.getRange(rowNumber, colEliminados + 1).setValue(resultadoEliminar.ok);
      sheet.getRange(rowNumber, colEjecutar + 1).setValue('Procesado');

      if (erroresOperacion.length) {
        errores++;
        sheet.getRange(rowNumber, colResultado + 1).setValue('PARCIAL');
        sheet.getRange(rowNumber, colError + 1).setValue(erroresOperacion.join(' | '));
      } else {
        procesados++;
        sheet.getRange(rowNumber, colResultado + 1).setValue('OK');
        sheet.getRange(rowNumber, colError + 1).clearContent();
      }

      registrarLog_('Mover alumno', alumno + ' | ' + nivel + ' | ' + origen + ' → ' + destino, 'Agregados: ' + resultadoAgregar.ok + ' | Eliminados: ' + resultadoEliminar.ok + ' | Errores: ' + erroresOperacion.length);
    } catch (error) {
      errores++;
      sheet.getRange(rowNumber, colResultado + 1).setValue('ERROR');
      sheet.getRange(rowNumber, colError + 1).setValue(error.message || String(error));
      registrarError_('procesarMoverAlumnos', '', alumno, error);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Proceso finalizado.\n\n' +
    'Filas completadas: ' + procesados + '\n' +
    'Filas con error o parcial: ' + errores
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

  return data.slice(1)
    .filter(row => String(row[colNivel] || '').trim() === nivel)
    .filter(row => String(row[colNomenclatura] || '').trim() === nomenclatura)
    .map(row => String(row[colCourseId] || '').trim())
    .filter(Boolean);
}

function ejecutarCambioAlumno_(alumno, courseIds, accion) {
  const result = { ok: 0, omitidos: 0, errores: [] };

  courseIds.forEach(courseId => {
    try {
      if (accion === 'agregar') agregarAlumno_(courseId, alumno);
      if (accion === 'eliminar') eliminarAlumno_(courseId, alumno);
      result.ok++;
    } catch (error) {
      const mensaje = error && error.message ? error.message : String(error);
      const lower = mensaje.toLowerCase();

      if (accion === 'agregar' && (lower.includes('already') || lower.includes('already exists'))) {
        result.omitidos++;
        return;
      }

      if (accion === 'eliminar' && lower.includes('not found')) {
        result.omitidos++;
        return;
      }

      result.errores.push(courseId + ': ' + mensaje);
    }
  });

  return result;
}

function buscarUsuariosAdmin() {
  registrarHojasExtendidas_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.ADMIN_SEARCH);
  const query = String(sheet.getRange('B1').getValue() || '').trim().toLowerCase();

  if (!query) {
    SpreadsheetApp.getUi().alert('Escribe una búsqueda en B1. Puede ser nombre, apellido o correo.');
    return;
  }

  limpiarResultadosBusquedaAdmin_(sheet);

  try {
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
        const fullName = user.name && user.name.fullName ? user.name.fullName : '';
        const email = user.primaryEmail || '';
        const haystack = (fullName + ' ' + email + ' ' + (user.orgUnitPath || '')).toLowerCase();
        if (haystack.includes(query)) {
          rows.push([
            fullName,
            email,
            user.orgUnitPath || '',
            user.suspended === true ? 'Sí' : 'No',
            user.id || ''
          ]);
        }
      });

      pageToken = response.nextPageToken;
    } while (pageToken);

    if (rows.length) sheet.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
    sheet.getRange('E1').setValue(rows.length);
    registrarLog_('Buscar usuario en Admin', query, rows.length + ' resultados');
    SpreadsheetApp.getUi().alert('Búsqueda finalizada.\n\nResultados: ' + rows.length);
  } catch (error) {
    registrarError_('buscarUsuariosAdmin', '', query, error);
    SpreadsheetApp.getUi().alert('No se pudo buscar en Google Admin. Revisa ERRORES.');
  }
}

function limpiarResultadosBusquedaAdmin_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 3) sheet.getRange(4, 1, lastRow - 3, lastCol).clearContent();
  sheet.getRange('E1').clearContent();
}

function validarUsuarioAdmin_(email) {
  try {
    const user = AdminDirectory.Users.get(email);
    if (!user || !user.primaryEmail) throw new Error('Usuario no encontrado en Google Admin.');
    if (user.suspended === true) throw new Error('El usuario está suspendido en Google Admin.');
    return user;
  } catch (error) {
    throw new Error('No se pudo validar el usuario en Google Admin: ' + (error.message || error));
  }
}

function columnaALetraNumero_(value) {
  const text = String(value || '').trim().toUpperCase();
  if (!text) throw new Error('Columna vacía en FUENTES.');
  if (/^\d+$/.test(text)) return Number(text);

  let number = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code < 65 || code > 90) throw new Error('Columna inválida: ' + value);
    number = number * 26 + (code - 64);
  }
  return number;
}
