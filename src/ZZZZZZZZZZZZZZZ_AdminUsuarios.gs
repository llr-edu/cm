function panelCambiarPasswordUsuario(email, password, exigirCambio) {
  const correo = String(email || '').trim().toLowerCase();
  const nuevaPassword = String(password || '').trim();
  const forzarCambio = exigirCambio === true;

  if (!correo) throw new Error('Selecciona primero un usuario.');
  if (!nuevaPassword) throw new Error('Escribe la contraseña nueva.');
  if (nuevaPassword.length < 8) throw new Error('La contraseña debe tener al menos 8 caracteres.');

  AdminDirectory.Users.patch({
    password: nuevaPassword,
    changePasswordAtNextLogin: forzarCambio
  }, correo);

  registrarLog_(
    'Cambiar contraseña desde panel',
    correo,
    'Forzar cambio al entrar: ' + (forzarCambio ? 'Sí' : 'No')
  );

  return {
    ok: true,
    message: 'Contraseña actualizada. Cambio al entrar: ' + (forzarCambio ? 'Sí.' : 'No.')
  };
}

function panelObtenerConfigPassword() {
  let password = '';

  try {
    password = String(obtenerConfig_('Contraseña predeterminada usuarios') || '').trim();
  } catch (error) {
    password = '';
  }

  return { password: password };
}
