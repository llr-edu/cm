/****************************************************
 * Panel operativo · agregar alumno a sección
 ****************************************************/

function panelAgregarAlumnoASeccion(email, nivel, seccion) {
  registrarHojasExtendidas_();

  const correo = String(email || '').trim().toLowerCase();
  const nivelText = String(nivel || '').trim();
  const seccionText = normalizarNomenclatura_(seccion);

  if (!correo) throw new Error('Selecciona primero un alumno.');
  if (!nivelText) throw new Error('Selecciona el nivel.');
  if (!seccionText) throw new Error('Selecciona la sección destino.');

  validarUsuarioAdmin_(correo);

  const aulas = obtenerAulasCatalogo_(nivelText, seccionText);
  if (!aulas.length) {
    throw new Error('No hay aulas en CATÁLOGO_AULAS para ' + nivelText + ' / ' + seccionText + '.');
  }

  const resultado = ejecutarCambioAlumno_(correo, aulas, 'agregar');
  const errores = resultado.errores || [];
  const omitidos = resultado.omitidos || 0;
  const agregados = resultado.ok || 0;
  const total = aulas.length;

  registrarLog_(
    'Agregar alumno desde panel',
    correo + ' | ' + nivelText + ' | ' + seccionText,
    'Agregados: ' + agregados + ' | Ya estaba: ' + omitidos + ' | Errores: ' + errores.length + ' | Total aulas: ' + total
  );

  if (errores.length) {
    registrarError_(
      'panelAgregarAlumnoASeccion',
      '',
      correo,
      new Error(errores.join(' | '))
    );
  }

  const mensaje = errores.length
    ? 'Proceso parcial. Agregados: ' + agregados + '. Ya estaba: ' + omitidos + '. Errores: ' + errores.length + '.'
    : 'Alumno agregado. Aulas agregadas: ' + agregados + '. Ya estaba inscrito: ' + omitidos + '.';

  return {
    ok: errores.length === 0,
    total: total,
    agregados: agregados,
    omitidos: omitidos,
    errores: errores,
    message: mensaje
  };
}
