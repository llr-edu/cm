/****************************************************
 * Panel operativo · aulas específicas
 ****************************************************/

function panelObtenerAulasGranulares(nivel, texto) {
  registrarHojasExtendidas_();

  const nivelText = String(nivel || '').trim();
  const filtro = normalizarTextoPanel_(texto || '');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = obtenerHojaObligatoria_(ss, SHEETS.CATALOG);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const colNivel = buscarColumna_(headers, 'Nivel');
  const colNomenclatura = buscarColumna_(headers, 'Nomenclatura');
  const colCourseId = buscarColumna_(headers, 'Course ID');
  const colNombreFuente = panelBuscarColumnaOpcional_(headers, 'Nombre fuente');
  const colNombreClassroom = panelBuscarColumnaOpcional_(headers, 'Nombre Classroom');
  const colMateria = panelBuscarColumnaOpcional_(headers, 'Materia Classroom');
  const colDescripcion = panelBuscarColumnaOpcional_(headers, 'Descripción fuente');
  const colEstado = panelBuscarColumnaOpcional_(headers, 'Estado Classroom');

  const vistos = {};
  const aulas = [];

  data.slice(1).forEach(row => {
    const nivelRow = String(row[colNivel] || '').trim();
    if (nivelText && nivelRow !== nivelText) return;

    const courseId = normalizarCourseId_(row[colCourseId]);
    if (!courseId || vistos[courseId]) return;

    const nombreFuente = panelValorColumna_(row, colNombreFuente);
    const nombreClassroom = panelValorColumna_(row, colNombreClassroom);
    const materia = panelValorColumna_(row, colMateria) || nombreFuente || nombreClassroom;
    const descripcion = panelValorColumna_(row, colDescripcion);
    const nomenclatura = normalizarNomenclatura_(row[colNomenclatura]);
    const estado = panelValorColumna_(row, colEstado);

    const hay = normalizarTextoPanel_([
      courseId,
      nivelRow,
      nomenclatura,
      nombreFuente,
      nombreClassroom,
      materia,
      descripcion,
      estado
    ].join(' '));

    if (filtro && hay.indexOf(filtro) === -1) return;

    vistos[courseId] = true;
    aulas.push({
      courseId: courseId,
      nivel: nivelRow,
      nomenclatura: nomenclatura,
      materia: materia || nombreClassroom || courseId,
      nombre: nombreClassroom || nombreFuente || materia || courseId,
      descripcion: descripcion,
      estado: estado
    });
  });

  aulas.sort((a, b) => {
    const byMateria = normalizarTextoPanel_(a.materia).localeCompare(normalizarTextoPanel_(b.materia), 'es', { sensitivity: 'base', numeric: true });
    if (byMateria !== 0) return byMateria;
    const byNom = normalizarTextoPanel_(a.nomenclatura).localeCompare(normalizarTextoPanel_(b.nomenclatura), 'es', { sensitivity: 'base', numeric: true });
    if (byNom !== 0) return byNom;
    return a.courseId.localeCompare(b.courseId);
  });

  return aulas.slice(0, 250);
}

function panelAplicarUsuarioEnAulasGranulares(email, tipoUsuario, courseIds, accion) {
  const inicio = new Date();
  const correo = String(email || '').trim().toLowerCase();
  const tipo = String(tipoUsuario || '').trim();
  const esProfesor = tipo === 'Profesor';
  const esEntrada = accion === 'entrada' || accion === 'agregar';
  const aulas = (courseIds || []).map(normalizarCourseId_).filter(Boolean)
    .filter((courseId, index, all) => all.indexOf(courseId) === index);

  if (!correo) throw new Error('Selecciona primero un usuario.');
  if (!aulas.length) throw new Error('Selecciona al menos un aula.');
  if (tipo !== 'Alumno' && tipo !== 'Profesor') throw new Error('Tipo de usuario inválido.');

  const resultado = esProfesor
    ? panelAplicarProfesorEnAulas_(correo, aulas, esEntrada ? 'entrada' : 'salida')
    : panelAplicarAlumnoEnAulas_(correo, aulas, esEntrada ? 'entrada' : 'salida');

  const segundos = Math.round((new Date() - inicio) / 1000);
  const accionTexto = esEntrada ? 'Agregar' : 'Retirar';
  const destinoTexto = esProfesor ? 'profesor' : 'alumno';

  registrarLog_(
    accionTexto + ' ' + destinoTexto + ' en aulas específicas',
    correo + ' | ' + aulas.join(', '),
    'Aulas: ' + aulas.length + ' | OK: ' + resultado.ok + ' | Omitidos: ' + resultado.omitidos + ' | Errores: ' + resultado.errores.length + ' | Tiempo: ' + segundos + ' s'
  );

  if (resultado.errores.length) {
    registrarError_('panelAplicarUsuarioEnAulasGranulares', '', correo, new Error(resultado.errores.join(' | ')));
  }

  return {
    ok: resultado.errores.length === 0,
    total: aulas.length,
    procesados: resultado.ok,
    omitidos: resultado.omitidos,
    errores: resultado.errores,
    elapsedMs: new Date() - inicio,
    message: accionTexto + ' en aulas específicas. Aulas: ' + aulas.length + '. Procesadas: ' + resultado.ok + '. Omitidas: ' + resultado.omitidos + '. Errores: ' + resultado.errores.length + '. Tiempo: ' + segundos + ' s.'
  };
}

function panelBuscarColumnaOpcional_(headers, headerName) {
  const index = headers.indexOf(headerName);
  return index === -1 ? null : index;
}

function panelValorColumna_(row, colIndex) {
  if (colIndex === null || colIndex === undefined || colIndex < 0) return '';
  return String(row[colIndex] || '').trim();
}
