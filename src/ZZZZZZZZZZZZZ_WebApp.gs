/****************************************************
 * Web App móvil
 ****************************************************/

function doGet(e) {
  registrarHojasExtendidas_();
  asegurarActualizacionNocturnaUsuarios_();

  const panel = e && e.parameter && String(e.parameter.panel || '').toLowerCase();
  const file = panel === 'clases' ? 'PanelClases' : 'PanelOperativo';
  const title = panel === 'clases' ? 'Clases · Classroom LL. RR.' : 'Classroom LL. RR.';

  return HtmlService
    .createHtmlOutputFromFile(file)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerUrlAppMovil() {
  const url = ScriptApp.getService().getUrl();
  return {
    ok: !!url,
    url: url || '',
    message: url ? 'URL de la app móvil disponible.' : 'Todavía no hay despliegue web app activo.'
  };
}

function mostrarUrlAppMovil() {
  const info = obtenerUrlAppMovil();
  const ui = SpreadsheetApp.getUi();

  if (!info.ok) {
    ui.alert(
      'App móvil no desplegada',
      'El código ya está preparado como Web App. Falta crear el despliegue web desde Apps Script.',
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert('App móvil', info.url, ui.ButtonSet.OK);
}
