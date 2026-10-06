/****************************************************
 * Classroom Manager LL. RR.
 * v1.9
 * Base: Setup + aulas + alumnos + profesores + consulta usuario + usuarios del dominio
 ****************************************************/

const APP_NAME = 'Classroom LL. RR.';

const SHEETS = {
  CONFIG: 'CONFIGURACIÓN',
  CLASSES: 'AULAS',
  STUDENTS: 'ALUMNOS',
  TEACHERS: 'PROFESORES',
  USER_LOOKUP: 'CONSULTA_USUARIO',
  USERS: 'USUARIOS',
  LOG: 'LOG',
  ERRORS: 'ERRORES'
};

const COLORS = {
  RED: '#b71c1c',
  LIGHT_YELLOW: '#fff2cc',
  LIGHT_BLUE: '#d9eaf7',
  LIGHT_GREEN: '#d9ead3',
  WHITE: '#ffffff',
  GRAY: '#eeeeee'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_NAME)
    .addItem('Start / Setup', 'startSetup')
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  crearHojaConfiguracion_(ss);
  crearHojaAulas_(ss);
  crearHojaAlumnos_(ss);
  crearHojaProfesores_(ss);
  crearHojaConsultaUsuario_(ss);
  crearHojaUsuarios_(ss);
  crearHojaLog_(ss);
  crearHojaErrores_(ss);

  inicializarConfiguracion_(ss);
  aplicarDesplegablesBasicos_();
  aplicarDesplegablesUsuarios();

  registrarLog_('Start / Setup', 'Estructura inicial creada o verificada.', 'OK');

  SpreadsheetApp.getUi().alert(
    'Setup completado.\n\n' +
    'Se crearon o verificaron las hojas base, incluida la hoja USUARIOS.\n\n' +
    'Nota: Start / Setup crea la hoja USUARIOS, pero no la llena automáticamente. ' +
    'Para traer usuarios desde Google Admin, ejecuta: Actualizar lista de usuarios.'
  );
}

function crearHojaConfiguracion_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.CONFIG, ['Campo', 'Valor', 'Notas']);
}

function crearHojaAulas_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.CLASSES, [
    'Course ID',
    'Nombre del aula',
    'Sección',
    'Materia',
    'Salón',
    'Dueño ID',
    'Estado',
    'Código de inscripción',
    'Link de Classroom',
    'Tutores activados',
    'Fecha de creación',
    'Última actualización',
    'Crear',
    'Actualizar',
    'Archivar',
    'Eliminar',
    'Nuevo dueño',
    'Resultado',
    'Error'
  ]);
}

function crearHojaAlumnos_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.STUDENTS, [
    'Course ID',
    'Nombre del aula',
    'Alumno',
    'Agregar',
    'Eliminar',
    'Mover a Course ID',
    'Fecha',
    'Resultado',
    'Error'
  ]);
}

function crearHojaProfesores_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.TEACHERS, [
    'Course ID',
    'Nombre del aula',
    'Profesor',
    'Agregar',
    'Eliminar',
    'Fecha',
    'Resultado',
    'Error'
  ]);
}

function crearHojaConsultaUsuario_(ss) {
  let sheet = ss.getSheetByName(SHEETS.USER_LOOKUP);
  if (!sheet) sheet = ss.insertSheet(SHEETS.USER_LOOKUP);
  resetearHojaCompleta_(sheet);

  sheet.getRange('A1').setValue('Tipo de usuario');
  sheet.getRange('B1').setValue('Alumno');
  sheet.getRange('C1').setValue('Correo');
  sheet.getRange('D1').setValue('');
  sheet.getRange('F1').setValue('Aulas encontradas');
  sheet.getRange('G1').setValue('');

  sheet.getRange('A1:G1').setFontWeight('bold');
  sheet.getRange('A1:C1').setBackground(COLORS.LIGHT_YELLOW);
  sheet.getRange('D1').setBackground('#ffff00');
  sheet.getRange('F1:G1').setBackground(COLORS.LIGHT_GREEN);

  const headers = [
    'Course ID',
    'Fecha de creación',
    'Organización del dueño',
    'Link de Classroom',
    'Código de inscripción',
    'Tutores activados',
    'Dueño',
    'Nombre del aula',
    'Sección',
    'Estado',
    'Encabezado de descripción',
    'Descripción',
    'Salón',
    'Materia'
  ];

  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(3, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(COLORS.RED)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(3);
  sheet.autoResizeColumns(1, headers.length);
  for (let i = 1; i <= headers.length; i++) sheet.setColumnWidth(i, 150);
}

function crearHojaUsuarios_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.USERS, [
    'Nombre completo',
    'Nombre',
    'Apellido',
    'Correo',
    'Unidad organizativa',
    'Suspendido',
    'Archivado',
    'ID de usuario',
    'Último acceso',
    'Fecha de creación'
  ]);
}

function crearHojaLog_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.LOG, [
    'Fecha',
    'Usuario ejecutor',
    'Acción',
    'Detalle',
    'Resultado'
  ]);
}

function crearHojaErrores_(ss) {
  crearHojaConEncabezados_(ss, SHEETS.ERRORS, [
    'Fecha',
    'Usuario ejecutor',
    'Función',
    'Course ID',
    'Correo afectado',
    'Mensaje de error',
    'Detalle técnico'
  ]);
}

function inicializarConfiguracion_(ss) {
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) return;

  const existing = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
  if (existing.includes('Dominio')) return;

  const data = [
    ['Dominio', 'losroblesenlinea.com.ve', 'Dominio institucional usado para actualizar USUARIOS.'],
    ['Año escolar', '2026-2027', 'Editable.'],
    ['Modo seguro', 'ACTIVO', 'Las acciones delicadas deben pedir confirmación.'],
    ['Versión', '1.9', 'Setup + aulas + alumnos + profesores + consulta + usuarios.']
  ];

  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  sheet.autoResizeColumns(1, 3);
}

function crearHojaConEncabezados_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  resetearHojaCompleta_(sheet);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(COLORS.RED)
    .setFontColor(COLORS.WHITE)
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  aplicarFormatoBasico_(sheet, headers.length);
  return sheet;
}

function resetearHojaCompleta_(sheet) {
  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();
  const range = sheet.getRange(1, 1, maxRows, maxColumns);

  const filter = sheet.getFilter();
  if (filter) filter.remove();

  range.clearContent();
  range.clearFormat();
  range.clearDataValidations();
  range.clearNote();
  sheet.clearConditionalFormatRules();
}

function aplicarFormatoBasico_(sheet, columnCount) {
  const maxRows = Math.max(sheet.getMaxRows(), 100);
  sheet.getRange(1, 1, maxRows, columnCount).setVerticalAlignment('middle');

  const filters = sheet.getFilter();
  if (filters) filters.remove();
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), columnCount).createFilter();

  for (let i = 1; i <= columnCount; i++) sheet.setColumnWidth(i, 150);
}

function aplicarDesplegablesBasicos_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lookup = ss.getSheetByName(SHEETS.USER_LOOKUP);
  if (lookup) {
    lookup.getRange('A:A').clearDataValidations();
    lookup.getRange('B:B').clearDataValidations();
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Alumno', 'Profesor'], true)
      .setAllowInvalid(false)
      .build();
    lookup.getRange('B1').setDataValidation(rule);
  }
}

function aplicarDesplegablesUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (!usersSheet) return;

  const lastRow = usersSheet.getLastRow();
  if (lastRow < 2) {
    registrarLog_('Aplicar desplegables de usuarios', 'La hoja USUARIOS está vacía. Ejecuta Actualizar lista de usuarios.', 'Sin cambios');
    return;
  }

  const userRange = usersSheet.getRange(2, 4, lastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(userRange, true)
    .setAllowInvalid(true)
    .build();

  const lookup = ss.getSheetByName(SHEETS.USER_LOOKUP);
  if (lookup) lookup.getRange('D1').setDataValidation(rule);

  const students = ss.getSheetByName(SHEETS.STUDENTS);
  if (students) students.getRange('C2:C').setDataValidation(rule);

  const teachers = ss.getSheetByName(SHEETS.TEACHERS);
  if (teachers) teachers.getRange('C2:C').setDataValidation(rule);

  registrarLog_('Aplicar desplegables de usuarios', 'Desplegables aplicados desde hoja USUARIOS.', 'OK');
}

function listarAulasActivas() {
  listarAulas_(['ACTIVE']);
}

function listarTodasLasAulas() {
  listarAulas_(['ACTIVE', 'ARCHIVED', 'PROVISIONED', 'DECLINED', 'SUSPENDED']);
}

function listarAulas_(courseStates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.CLASSES);

  try {
    limpiarDatosManteniendoEncabezado_(sheet);

    const rows = [];
    let pageToken = null;

    do {
      const response = Classroom.Courses.list({
        pageSize: 100,
        pageToken: pageToken,
        courseStates: courseStates
      });

      const courses = response.courses || [];
      courses.forEach(course => {
        rows.push([
          course.id || '',
          course.name || '',
          course.section || '',
          course.subject || '',
          course.room || '',
          course.ownerId || '',
          course.courseState || '',
          course.enrollmentCode || '',
          course.alternateLink || '',
          normalizarBooleano_(course.guardiansEnabled),
          formatearFecha_(course.creationTime),
          formatearFecha_(course.updateTime),
          '', '', '', '', '', '', ''
        ]);
      });

      pageToken = response.nextPageToken;
    } while (pageToken);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      sheet.autoResizeColumns(1, sheet.getLastColumn());
    }

    registrarLog_('Listar aulas', 'Estados consultados: ' + courseStates.join(', '), rows.length + ' aulas listadas');
    SpreadsheetApp.getUi().alert('Listado completado.\n\nAulas encontradas: ' + rows.length);
  } catch (error) {
    registrarError_('listarAulas_', '', '', error);
    SpreadsheetApp.getUi().alert('No se pudieron listar las aulas.\n\nRevisa la hoja ERRORES.');
  }
}

function procesarAlumnos() {
  procesarMiembros_(SHEETS.STUDENTS, 'Alumno', true);
}

function procesarProfesores() {
  procesarMiembros_(SHEETS.TEACHERS, 'Profesor', false);
}

function procesarMiembros_(sheetName, memberHeader, isStudent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, sheetName);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('No hay datos para procesar.');
    return;
  }

  const headers = data[0];
  const colCourseId = buscarColumna_(headers, 'Course ID');
  const colMember = buscarColumna_(headers, memberHeader);
  const colAdd = buscarColumna_(headers, 'Agregar');
  const colDelete = buscarColumna_(headers, 'Eliminar');
  const colDate = buscarColumna_(headers, 'Fecha');
  const colResult = buscarColumna_(headers, 'Resultado');
  const colError = buscarColumna_(headers, 'Error');

  let procesados = 0;
  let errores = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const courseId = String(row[colCourseId] || '').trim();
    const email = String(row[colMember] || '').trim().toLowerCase();
    const addValue = String(row[colAdd] || '').trim().toUpperCase();
    const deleteValue = String(row[colDelete] || '').trim().toUpperCase();

    const wantsAdd = esSi_(addValue);
    const wantsDelete = esSi_(deleteValue);

    if (!wantsAdd && !wantsDelete) continue;

    const rowNumber = i + 1;
    sheet.getRange(rowNumber, colDate + 1).setValue(new Date());
    sheet.getRange(rowNumber, colResult + 1).clearContent();
    sheet.getRange(rowNumber, colError + 1).clearContent();

    try {
      if (!courseId) throw new Error('Falta el Course ID.');
      if (!email) throw new Error('Falta el correo.');
      if (wantsAdd && wantsDelete) throw new Error('No puedes marcar Agregar y Eliminar al mismo tiempo.');

      if (wantsAdd) {
        sheet.getRange(rowNumber, colDelete + 1).clearContent();

        if (isStudent) agregarAlumno_(courseId, email);
        else agregarProfesor_(courseId, email);

        sheet.getRange(rowNumber, colAdd + 1).setValue('Agregado');
        sheet.getRange(rowNumber, colResult + 1).setValue('OK');
        registrarLog_(isStudent ? 'Agregar alumno' : 'Agregar profesor', 'Curso: ' + courseId + ' | Correo: ' + email, 'Agregado');
        procesados++;
      }

      if (wantsDelete) {
        sheet.getRange(rowNumber, colAdd + 1).clearContent();

        if (isStudent) eliminarAlumno_(courseId, email);
        else eliminarProfesor_(courseId, email);

        sheet.getRange(rowNumber, colDelete + 1).setValue('Eliminado');
        sheet.getRange(rowNumber, colResult + 1).setValue('OK');
        registrarLog_(isStudent ? 'Eliminar alumno' : 'Eliminar profesor', 'Curso: ' + courseId + ' | Correo: ' + email, 'Eliminado');
        procesados++;
      }
    } catch (error) {
      const mensaje = interpretarErrorClassroom_(error, wantsAdd, wantsDelete, isStudent);

      if (wantsAdd) {
        sheet.getRange(rowNumber, colAdd + 1).setValue(isStudent ? 'Ya estaba inscrito' : 'Ya estaba agregado');
        sheet.getRange(rowNumber, colDelete + 1).clearContent();
      }

      if (wantsDelete) {
        sheet.getRange(rowNumber, colDelete + 1).setValue(isStudent ? 'No estaba inscrito' : 'No estaba agregado');
        sheet.getRange(rowNumber, colAdd + 1).clearContent();
      }

      errores++;
      sheet.getRange(rowNumber, colResult + 1).setValue('ERROR');
      sheet.getRange(rowNumber, colError + 1).setValue(mensaje);
      registrarError_(isStudent ? 'procesarAlumnos' : 'procesarProfesores', courseId, email, error);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Proceso finalizado.\n\n' +
    'Acciones completadas: ' + procesados + '\n' +
    'Errores: ' + errores
  );
}

function agregarAlumno_(courseId, studentEmail) {
  Classroom.Courses.Students.create({ userId: studentEmail }, courseId);
}

function eliminarAlumno_(courseId, studentEmail) {
  Classroom.Courses.Students.remove(courseId, studentEmail);
}

function agregarProfesor_(courseId, teacherEmail) {
  Classroom.Courses.Teachers.create({ userId: teacherEmail }, courseId);
}

function eliminarProfesor_(courseId, teacherEmail) {
  Classroom.Courses.Teachers.remove(courseId, teacherEmail);
}

function consultarUsuarioAulasActivas() {
  consultarUsuario_(['ACTIVE']);
}

function consultarUsuarioTodasLasAulas() {
  consultarUsuario_(['ACTIVE', 'ARCHIVED', 'PROVISIONED', 'DECLINED', 'SUSPENDED']);
}

function consultarUsuario_(courseStates) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.USER_LOOKUP);

  const tipo = String(sheet.getRange('B1').getValue() || '').trim().toLowerCase();
  const correo = String(sheet.getRange('D1').getValue() || '').trim().toLowerCase();

  if (!tipo || !correo) {
    SpreadsheetApp.getUi().alert('Selecciona el tipo de usuario en B1 y escribe o selecciona el correo en D1.');
    return;
  }

  const esAlumno = tipo === 'alumno' || tipo === 'student';
  const esProfesor = tipo === 'profesor' || tipo === 'teacher';

  if (!esAlumno && !esProfesor) {
    SpreadsheetApp.getUi().alert('Tipo de usuario inválido. Usa Alumno o Profesor.');
    return;
  }

  try {
    limpiarResultadosConsultaUsuario_(sheet);

    const rows = [];
    let pageToken = null;

    do {
      const params = {
        pageSize: 100,
        pageToken: pageToken,
        courseStates: courseStates
      };

      if (esAlumno) params.studentId = correo;
      if (esProfesor) params.teacherId = correo;

      const response = Classroom.Courses.list(params);
      const courses = response.courses || [];

      courses.forEach(course => {
        rows.push([
          course.id || '',
          formatearFecha_(course.creationTime),
          '',
          course.alternateLink || '',
          course.enrollmentCode || '',
          course.guardiansEnabled === true ? 'TRUE' : 'FALSE',
          course.ownerId || '',
          course.name || '',
          course.section || '',
          course.courseState || '',
          course.descriptionHeading || '',
          course.description || '',
          course.room || '',
          course.subject || ''
        ]);
      });

      pageToken = response.nextPageToken;
    } while (pageToken);

    if (rows.length > 0) {
      sheet.getRange(4, 1, rows.length, rows[0].length).setValues(rows);
      sheet.autoResizeColumns(1, rows[0].length);
    }

    sheet.getRange('G1').setValue(rows.length);
    registrarLog_('Consultar usuario', 'Tipo: ' + tipo + ' | Correo: ' + correo, rows.length + ' aulas encontradas');

    SpreadsheetApp.getUi().alert('Consulta finalizada.\n\nAulas encontradas: ' + rows.length);
  } catch (error) {
    registrarError_('consultarUsuario_', '', correo, error);
    SpreadsheetApp.getUi().alert('No se pudo consultar el usuario.\n\nRevisa la hoja ERRORES.');
  }
}

function limpiarResultadosConsultaUsuario_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 3) {
    const range = sheet.getRange(4, 1, lastRow - 3, lastCol);
    range.clearContent();
    range.clearDataValidations();
  }
  sheet.getRange('G1').clearContent();
}

function actualizarListaUsuarios() {
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
    SpreadsheetApp.getUi().alert('Lista de usuarios actualizada.\n\nUsuarios encontrados: ' + rows.length);
  } catch (error) {
    registrarError_('actualizarListaUsuarios', '', '', error);
    SpreadsheetApp.getUi().alert('No se pudo actualizar la lista de usuarios.\n\nRevisa la hoja ERRORES.');
  }
}

function obtenerConfig_(campo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!sheet) return '';

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === campo) return String(data[i][1] || '').trim();
  }
  return '';
}

function limpiarDatosManteniendoEncabezado_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    range.clearContent();
    range.clearDataValidations();
  }
}

function obtenerHojaObligatoria_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('No existe la hoja obligatoria: ' + sheetName + '. Ejecuta Start / Setup primero.');
  return sheet;
}

function buscarColumna_(headers, headerName) {
  const index = headers.indexOf(headerName);
  if (index === -1) throw new Error('No se encontró la columna obligatoria: ' + headerName);
  return index;
}

function esSi_(value) {
  const text = String(value || '').trim().toUpperCase();
  return text === 'Y' || text === 'SÍ' || text === 'SI';
}

function interpretarErrorClassroom_(error, wantsAdd, wantsDelete, isStudent) {
  const raw = error && error.message ? error.message : String(error);
  const text = raw.toLowerCase();

  if (text.includes('already exists') || text.includes('already')) {
    return isStudent ? 'El usuario ya está inscrito en esta aula.' : 'El profesor ya está agregado en esta aula.';
  }

  if (text.includes('not found')) {
    if (wantsDelete) return isStudent ? 'No se encontró el aula o el alumno en esa aula.' : 'No se encontró el aula o el profesor en esa aula.';
    return 'No se encontró el aula o el usuario. Revisa el Course ID y el correo.';
  }

  if (text.includes('permission') || text.includes('forbidden')) return 'No hay permisos suficientes para realizar esta acción.';
  if (text.includes('invalid')) return 'Dato inválido. Revisa el correo o el Course ID.';
  if (text.includes('failed precondition')) return 'Google Classroom rechazó la acción por una condición del curso o del usuario.';
  if (text.includes('quota')) return 'Se alcanzó un límite temporal de uso de la API. Intenta más tarde.';

  return raw;
}

function registrarLog_(accion, detalle, resultado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LOG);
  if (!sheet) return;
  sheet.appendRow([new Date(), obtenerUsuario_(), accion, detalle, resultado]);
}

function registrarError_(funcion, courseId, correoAfectado, error) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ERRORS);
  if (!sheet) return;

  sheet.appendRow([
    new Date(),
    obtenerUsuario_(),
    funcion,
    courseId || '',
    correoAfectado || '',
    error && error.message ? error.message : String(error),
    JSON.stringify(error, Object.getOwnPropertyNames(error))
  ]);
}

function obtenerUsuario_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    return '';
  }
}

function formatearFecha_(value) {
  if (!value) return '';
  try {
    return new Date(value);
  } catch (error) {
    return value;
  }
}

function normalizarBooleano_(value) {
  if (value === true) return 'Sí';
  if (value === false) return 'No';
  return '';
}

function irALog() {
  irAHoja_(SHEETS.LOG);
}

function irAErrores() {
  irAHoja_(SHEETS.ERRORS);
}

function irAHoja_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.setActiveSheet(sheet);
}
