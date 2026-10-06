/****************************************************
 * Panel operativo v4 · mover reconstruyendo inscripción
 ****************************************************/

function panelMoverAlumnoSeccion(email, nivel, origen, destino) {
  registrarHojasExtendidas_();

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

  const aulasOrigen = panelAulasPorSeccion_(nivelText, origenText, true);
  const aulasDestino = panelAulasPorSeccion_(nivelText, destinoText, true);

  if (!aulasOrigen.length) throw new Error('No hay aulas para el origen ' + origenText + '.');
  if (!aulasDestino.length) throw new Error('No hay aulas para el destino ' + destinoText + '.');

  const salida = panelAplicarAlumnoEnAulas_(correo, aulasOrigen, 'salida');
  const entrada = panelAplicarAlumnoEnAulas_(correo, aulasDestino, 'entrada');
  const errores = salida.errores.concat(entrada.errores);
  const omitidos = salida.omitidos + entrada.omitidos;
  const segundos = Math.round((new Date() - inicio) / 1000);

  registrarLog_(
    'Mover alumno desde panel',
    correo + ' | ' + nivelText + ' | ' + origenText + ' → ' + destinoText,
    'Origen: ' + aulasOrigen.length + ' | Destino: ' + aulasDestino.length + ' | Eliminados: ' + salida.ok + ' | Agregados: ' + entrada.ok + ' | Omitidos: ' + omitidos + ' | Errores: ' + errores.length + ' | Tiempo: ' + segundos + ' s'
  );

  if (errores.length) {
    registrarError_('panelMoverAlumnoSeccion', '', correo, new Error(errores.join(' | ')));
  }

  return {
    ok: errores.length === 0,
    eliminados: salida.ok,
    agregados: entrada.ok,
    omitidos: omitidos,
    errores: errores,
    elapsedMs: new Date() - inicio,
    message: 'Movimiento finalizado. Eliminados: ' + salida.ok + '. Agregados: ' + entrada.ok + '. Ya estaba/no estaba: ' + omitidos + '. Errores: ' + errores.length + '. Tiempo: ' + segundos + ' s.'
  };
}
