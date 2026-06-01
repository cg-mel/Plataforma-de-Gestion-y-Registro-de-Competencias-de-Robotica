// ==========================================
// CRODIT - UABC | Sistema de Competencias de Robótica
// Centro de Robótica y Diseño Industrial Tecnológico
// Facultad de Ingeniería Mexicali — UABC
// OAuth2 popup + polling (sin multi-cuenta)
// Deploy: Execute as ME + Access CUALQUIERA
// ==========================================

// ---- CONFIGURACIÓN ----
var CLIENT_ID     = '78676865323-8hb236h68r7u1ah64pr78tjk9e6lmegp.apps.googleusercontent.com';
var CLIENT_SECRET = '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';

// URL EXTERNA del callback (hospedar oauth-callback.html fuera de script.google.com)
// Ejemplos: GitHub Pages, Google Cloud Storage, Netlify, etc.
// ⚠️ IMPORTANTE: Esta MISMA URL debe registrarse como URI de redirección en Google Cloud Console
 // ← PEGAR AQUÍ la URL donde hospedas oauth-callback.html
var CALLBACK_URL = 'https://cristovea.github.io/robotech-callback12/oauth-callback.html';
// ==========================================
// SETUP: Ejecutar MANUALMENTE después de cada
// nueva implementación (deploy)
// ==========================================

// ▶ PASO 1: Vincula la hoja de cálculo
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    Logger.log('ERROR: Abre este script desde dentro de tu hoja de cálculo.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('✅ ID de hoja guardado: ' + ss.getId());
}

// ▶ PASO 2: Guarda la URL de producción (/exec)
//   Después de implementar, copia la URL /exec y pégala aquí:
//   Ejemplo: guardarUrlProduccion("https://script.google.com/macros/s/AKfycb.../exec")
function guardarUrlProduccion(url) {
  if (!url || url.indexOf('/exec') === -1) {
    Logger.log('❌ ERROR: La URL debe terminar en /exec');
    Logger.log('Ve a Implementar → Administrar implementaciones → Copia la URL');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('PRODUCTION_URL', url);
  Logger.log('========================================');
  Logger.log('✅ URL de producción guardada:');
  Logger.log(url);
  Logger.log('');
  Logger.log('⚠️ IMPORTANTE: Agrega esta MISMA URL como');
  Logger.log('URI de redirección en Google Cloud Console:');
  Logger.log('APIs y servicios → Credenciales → Tu Client ID → URIs de redirección');
  Logger.log('========================================');
}

// ▶ PASO 3: Verifica todo
function verificarConfiguracion() {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('PRODUCTION_URL');
  var sheetId = props.getProperty('SHEET_ID');
  Logger.log('========================================');
  Logger.log('📋 CONFIGURACIÓN:');
  Logger.log('CLIENT_ID: ' + CLIENT_ID);
  Logger.log('PRODUCTION_URL: ' + (url || '❌ NO CONFIGURADA - ejecuta guardarUrlProduccion()'));
  Logger.log('SHEET_ID: ' + (sheetId || '❌ NO CONFIGURADA - ejecuta setupSheet()'));
  Logger.log('');
  Logger.log('⚠️ En Google Cloud Console agrega como URI de redirección:');
  Logger.log(url || '(primero guarda la URL de producción)');
  Logger.log('========================================');
}

// ▶ PASO 4: Crear todas las tablas necesarias
// Ejecutar UNA VEZ desde el editor de Apps Script
function crearTablas() {
  var ss = getSpreadsheet();
  var style = function(sheet, headers) {
    var r = sheet.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  };

  // Coaches
  if (!ss.getSheetByName('Coaches')) {
    var s = ss.insertSheet('Coaches');
    style(s, ['idCoach','Nombre','Apellido','Correo','Telefono','Institucion','FechaRegistro']);
  }

  // Participantes
  if (!ss.getSheetByName('Participantes')) {
    var s = ss.insertSheet('Participantes');
    style(s, ['idParticipante','Nombre','Apellido','Edad','Correo','Escolaridad','CoachEmail','FechaRegistro']);
  }

  // Equipos
  if (!ss.getSheetByName('Equipos')) {
    var s = ss.insertSheet('Equipos');
    style(s, ['idEquipo','NombreEquipo','idCategoria','CoachEmail','FechaCreacion']);
  }

  // Equipo_Participante
  if (!ss.getSheetByName('Equipo_Participante')) {
    var s = ss.insertSheet('Equipo_Participante');
    style(s, ['idEquipo','idParticipante','FechaAsignacion']);
  }

  // Eventos
  if (!ss.getSheetByName('Eventos')) {
    var s = ss.insertSheet('Eventos');
    style(s, ['idEvento','NombreEvento','Descripcion','FechaInicio','FechaFin','Ubicacion','Estado']);
  }

  // Inscripciones
  if (!ss.getSheetByName('Inscripciones')) {
    var s = ss.insertSheet('Inscripciones');
    style(s, ['idInscripcion','idEquipo','idEvento','idCategoria','CoachEmail','FechaInscripcion','Estado','EvidenciaURL']);
  }

  // Categorias — CRODIT UABC (10 categorías oficiales)
  if (!ss.getSheetByName('Categorias')) {
    var s = ss.insertSheet('Categorias');
    style(s, ['idCategoria','Nombre','Disciplina','Descripcion','Restricciones','MaxIntegrantes','MinIntegrantes']);
    _seedCategoriasCRODIT_(s);
  }

  // Evento_Categoria
  if (!ss.getSheetByName('Evento_Categoria')) {
    var s = ss.insertSheet('Evento_Categoria');
    style(s, ['idEvento','idCategoria']);
  }

  Logger.log('✅ Todas las tablas han sido creadas correctamente.');
}

// ==========================================
// HOJA DE CÁLCULO
// ==========================================
function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Ejecuta setupSheet() desde el editor GAS para vincular la hoja.');
  return SpreadsheetApp.openById(id);
}

function getUsersSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Usuarios');
  if (!sheet) {
    sheet = ss.insertSheet('Usuarios');
    var h = ['Correo','Nombre','Foto','Institución','Rol','Ciudad','País','Celular','Fecha','BienvenidaEnviada'];
    var r = sheet.getRange(1, 1, 1, h.length);
    r.setValues([h]);
    r.setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 180);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 140);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 150);
    sheet.setColumnWidth(9, 160);
    sheet.setColumnWidth(10, 130);
  } else {
    // Auto-migración: añadir columna BienvenidaEnviada si la hoja existe pero no la tiene
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (headers.indexOf('BienvenidaEnviada') === -1) {
      sheet.getRange(1, lastCol + 1).setValue('BienvenidaEnviada')
        .setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#ffffff');
      sheet.setColumnWidth(lastCol + 1, 130);
    }
  }
  return sheet;
}

// Marca al usuario como "bienvenida enviada" en la hoja Usuarios
function _marcarBienvenidaEnviada_(email) {
  try {
    var sheet = getUsersSheet();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var col = headers.indexOf('BienvenidaEnviada') + 1;
    if (col === 0) {
      Logger.log('[_marcarBienvenidaEnviada_] columna BienvenidaEnviada no existe');
      return false;
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        sheet.getRange(i + 1, col).setValue('Sí (' + new Date().toISOString() + ')');
        Logger.log('[_marcarBienvenidaEnviada_] Marcado correctamente: ' + email);
        return true;
      }
    }
    Logger.log('[_marcarBienvenidaEnviada_] Usuario no encontrado: ' + email);
    return false;
  } catch(e) {
    Logger.log('[_marcarBienvenidaEnviada_] Error: ' + e);
    return false;
  }
}

function findUserByEmail(email) {
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
      obj._row = i + 1;
      return obj;
    }
  }
  return null;
}

// ==========================================
// HELPER: obtener URL de producción
// ==========================================
function getProductionUrl_() {
  return PropertiesService.getScriptProperties().getProperty('PRODUCTION_URL')
      || ScriptApp.getService().getUrl();
}

// ==========================================
// OAUTH2 MANUAL + POPUP + POLLING
// OAuth en popup → callback a /exec (anónimo)
// → main page detecta sesión via polling
// Deploy DEBE ser: Ejecutar como YO + CUALQUIERA
// ==========================================

function getAuthorizationUrl(loginId) {
  var gasUrl = getProductionUrl_();
  // Si hay callback externo, usarlo; si no, fallback al /exec (solo funciona con cuenta principal)
  var redirectUri = CALLBACK_URL || gasUrl;
  var stateId = loginId || Utilities.getUuid();

  // State codifica: loginId|gasExecUrl (la callback page necesita la URL de GAS para el fallback)
  var state = CALLBACK_URL
    ? stateId + '|' + encodeURIComponent(gasUrl)
    : stateId;

  return 'https://accounts.google.com/o/oauth2/v2/auth'
    + '?client_id='     + encodeURIComponent(CLIENT_ID)
    + '&redirect_uri='  + encodeURIComponent(redirectUri)
    + '&response_type=code'
    + '&scope='         + encodeURIComponent('openid email profile')
    + '&access_type=offline'
    + '&prompt=select_account'
    + '&state='         + encodeURIComponent(state);
}

// Llamada desde el frontend para verificar si el login en el popup terminó
function checkLoginResult(loginId) {
  if (!loginId) return null;
  var data = CacheService.getScriptCache().get('login_' + loginId);
  if (!data) return null;
  CacheService.getScriptCache().remove('login_' + loginId);
  try { return JSON.parse(data); } catch(e) { return null; }
}

// Llamada desde el frontend (postMessage flow): intercambia code por sesión
function exchangeCodeForSession(code, loginId) {
  if (!code) return { success: false, error: 'No se recibió código.' };
  var redirectUri = CALLBACK_URL || getProductionUrl_();
  return exchangeCodeInternal_(code, loginId, redirectUri);
}

// Procesa el callback OAuth (?code= en doGet) — flujo directo sin callback externo
function handleOAuthCallback_(code, state) {
  var loginId = state || '';
  if (CALLBACK_URL && state) {
    loginId = state.split('|')[0] || state;
  }
  var redirectUri = CALLBACK_URL || getProductionUrl_();
  return exchangeCodeInternal_(code, loginId, redirectUri);
}

// ---- Lógica compartida: intercambia code → tokens → sesión ----
function exchangeCodeInternal_(code, loginId, redirectUri) {
  // Intercambiar code por tokens
  var tokenResponse = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      code:          code,
      client_id:     CLIENT_ID,
      client_secret: CLIENT_SECRET,
      redirect_uri:  redirectUri,
      grant_type:    'authorization_code'
    },
    muteHttpExceptions: true
  });

  if (tokenResponse.getResponseCode() !== 200) {
    Logger.log('Token error: ' + tokenResponse.getContentText());
    return { success: false, error: 'Error al verificar con Google.' };
  }

  var tokens = JSON.parse(tokenResponse.getContentText());

  // Obtener info del usuario
  var userResponse = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { Authorization: 'Bearer ' + tokens.access_token },
    muteHttpExceptions: true
  });

  if (userResponse.getResponseCode() !== 200) {
    return { success: false, error: 'Error al obtener datos del usuario.' };
  }

  var userInfo = JSON.parse(userResponse.getContentText());
  if (!userInfo.email) {
    return { success: false, error: 'No se pudo obtener el email.' };
  }

  // Crear sesión
  var sessionToken = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + sessionToken, JSON.stringify({
    email:   userInfo.email,
    name:    userInfo.name    || '',
    picture: userInfo.picture || ''
  }), 21600);

  // Determinar destino
  var existing = findUserByEmail(userInfo.email);
  var targetPage = 'profile';
  if (existing) {
    var rol = (existing['Rol'] || '').toLowerCase();
    if (rol === 'coach') targetPage = 'coach';
    else if (rol === 'admin' || rol === 'organizador') targetPage = 'admin';
    else if (rol === 'participante') targetPage = 'participante';
    else targetPage = 'dashboard';
  }

  // Guardar resultado para que el polling lo detecte
  var appUrl = getProductionUrl_();
  var resultData = {
    success: true,
    token: sessionToken,
    targetPage: targetPage,
    userName: userInfo.name || userInfo.email,
    redirectUrl: appUrl + '?page=' + targetPage + '&token=' + sessionToken
  };

  if (loginId) {
    CacheService.getScriptCache().put('login_' + loginId, JSON.stringify(resultData), 300);
  }

  return resultData;
}

// ==========================================
// ROUTING
// ==========================================
function doGet(e) {
  var params = (e && e.parameter) || {};

  // ---- Fallback: ?processCode=1 (llega desde oauth-callback.html via <img>) ----
  if (params.processCode === '1' && params.code) {
    var loginId = params.loginId || '';
    var redirectUri = CALLBACK_URL || getProductionUrl_();
    exchangeCodeInternal_(params.code, loginId, redirectUri);
    // Responder con un pixel transparente (el request viene de un <img>)
    return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
  }

  // ---- OAuth callback: ?code= (llega desde el popup — flujo sin callback externo) ----
  if (params.code) {
    var result = handleOAuthCallback_(params.code, params.state || '');
    var appUrl = getProductionUrl_();

    if (!result.success) {
      return HtmlService.createHtmlOutput(
        '<html><body style="font-family:Space Grotesk,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#0A0E1A;color:#fff;margin:0;">' +
        '<div style="text-align:center;background:#0f1629;border:1px solid rgba(13,127,242,0.3);border-radius:20px;padding:48px 40px;max-width:400px;">' +
        '<div style="font-size:48px;margin-bottom:16px;">&#10060;</div>' +
        '<h2 style="margin:0 0 8px;">Error de autenticación</h2>' +
        '<p style="color:#f87171;font-size:14px;">' + (result.error || 'Error desconocido') + '</p>' +
        '<p style="color:#64748b;font-size:13px;margin-top:16px;">Puedes cerrar esta ventana e intentar de nuevo.</p>' +
        '</div></body></html>'
      ).setTitle('Error - CRODIT UABC')
       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Éxito — mostrar página que indica cerrar el popup
    return HtmlService.createHtmlOutput(
      '<html><head>' +
      '<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap" rel="stylesheet">' +
      '</head><body style="font-family:Space Grotesk,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#0A0E1A;color:#fff;margin:0;">' +
      '<div style="text-align:center;background:#0f1629;border:1px solid rgba(13,127,242,0.3);border-radius:20px;padding:48px 40px;max-width:400px;">' +
      '<div style="font-size:48px;margin-bottom:16px;">&#9989;</div>' +
      '<h2 style="margin:0 0 8px;">¡Sesión iniciada!</h2>' +
      '<p style="color:#94a3b8;font-size:14px;margin:0 0 8px;">Bienvenido, ' + result.userName + '</p>' +
      '<p style="color:#64748b;font-size:13px;">Esta ventana se cerrará automáticamente...</p>' +
      '</div>' +
      '<script>setTimeout(function(){ window.close(); }, 1500);</script>' +
      '</body></html>'
    ).setTitle('CRODIT UABC')
     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ---- Páginas normales ----
  var page  = params.page  || '';
  var token = params.token || '';

  if (!page || page === 'login') {
    return HtmlService.createHtmlOutputFromFile('Login')
      .setTitle('CRODIT UABC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var templates = {
    'profile':      { file: 'ProfileForm',     title: 'CRODIT UABC - Completa tu perfil' },
    'dashboard':    { file: 'Dashboard',       title: 'CRODIT UABC - Dashboard' },
    'coach':        { file: 'CoachDashboard',  title: 'CRODIT UABC - Panel Asesor' },
    'admin':        { file: 'AdminDashboard',  title: 'CRODIT UABC - Administración' },
    'participante': { file: 'ParticipantView', title: 'CRODIT UABC - Participante' }
  };

  var cfg = templates[page];
  if (cfg) {
    var t = HtmlService.createTemplateFromFile(cfg.file);
    t.token = token || '';
    return t.evaluate()
      .setTitle(cfg.title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('Login')
    .setTitle('CRODIT UABC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// SESIONES
// ==========================================
function getSessionUser_(token) {
  if (!token) return null;
  var data = CacheService.getScriptCache().get('s_' + token);
  if (!data) return null;
  try { return JSON.parse(data); } catch(e) { return null; }
}

// ==========================================
// OBTENER INFO DEL USUARIO
// ==========================================
function getUserInfo(token) {
  var session = getSessionUser_(token);
  if (!session) return { error: 'Sesión expirada. Inicia sesión de nuevo.' };

  var email = session.email;
  try {
    var saved = findUserByEmail(email);
    if (saved) {
      return {
        email: email,
        name: saved['Nombre'] || session.name,
        given_name: (saved['Nombre'] || session.name || '').split(' ')[0],
        picture: saved['Foto'] || session.picture,
        institucion: saved['Institución'] || '',
        rol: saved['Rol'] || '',
        ciudad: saved['Ciudad'] || '',
        pais: saved['País'] || '',
        celular: saved['Celular'] || ''
      };
    }
  } catch(e) {}

  var displayName = session.name || email.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, function(c) { return c.toUpperCase(); });
  return {
    email: email,
    name: displayName,
    given_name: displayName.split(' ')[0],
    picture: session.picture || null
  };
}

// ==========================================
// GUARDAR PERFIL
// ==========================================
function saveUserProfile(token, data) {
  try {
    Logger.log('[saveUserProfile] INICIO — token: ' + (token ? token.substring(0,8)+'...' : 'NULL'));
    var session = getSessionUser_(token);
    if (!session) {
      Logger.log('[saveUserProfile] ABORTADO: sesión expirada');
      return { success: false, error: 'Sesión expirada. Inicia sesión de nuevo.' };
    }

    var email = session.email;
    var nombre = (data.nombre || '').trim();
    var foto = session.picture || '';
    var institucion = (data.institucion || '').trim();
    var rol = (data.rol || '').trim();
    var ciudad = (data.ciudad || '').trim();
    var pais = (data.pais || '').trim();
    var celular = (data.celular || '').trim();

    Logger.log('[saveUserProfile] email: ' + email + ' | nombre: ' + nombre + ' | rol: ' + rol);

    var sheet = getUsersSheet();
    var existing = findUserByEmail(email);
    var esNuevo = !existing;
    Logger.log('[saveUserProfile] esNuevo: ' + esNuevo + ' | existing: ' + (existing ? 'fila '+existing._row : 'no encontrado'));

    if (existing) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var updates = {
        'Nombre': nombre, 'Foto': foto, 'Institución': institucion,
        'Rol': rol, 'Ciudad': ciudad, 'País': pais, 'Celular': celular
      };
      for (var key in updates) {
        var col = headers.indexOf(key) + 1;
        if (col > 0) sheet.getRange(existing._row, col).setValue(updates[key]);
      }
    } else {
      sheet.appendRow([email, nombre, foto, institucion, rol, ciudad, pais, celular, new Date()]);
    }

    // ---- Insertar en tabla específica según rol ----
    var rolLower = rol.toLowerCase();
    var ss = getSpreadsheet();

    if (rolLower === 'coach') {
      registrarEnTablaCoach_(ss, email, nombre, institucion, celular);
    } else if (rolLower === 'participante') {
      registrarEnTablaParticipante_(ss, email, nombre, data);
    } else if (rolLower === 'organizador') {
      registrarEnTablaOrganizador_(ss, email, nombre, institucion, celular);
    }

    // ENVÍO ROBUSTO DEL CORREO DE BIENVENIDA
    // Ya no depende solo de 'esNuevo' — usa la columna BienvenidaEnviada
    // Y va envuelto en try-catch para que un fallo no rompa el registro
    try {
      var bienvenidaActual = existing ? String(existing['BienvenidaEnviada'] || '').trim() : '';
      var yaEnviada = bienvenidaActual.toLowerCase().indexOf('sí') === 0 || bienvenidaActual.toLowerCase().indexOf('si') === 0;
      Logger.log('[saveUserProfile] Estado bienvenida — esNuevo: ' + esNuevo + ' | yaEnviada: ' + yaEnviada + ' | valorActual: "' + bienvenidaActual + '"');

      if (!yaEnviada) {
        Logger.log('[saveUserProfile] >>> Disparando correo de bienvenida a ' + email);
        var enviado = _enviarCorreoBienvenida_(email, nombre, rolLower);
        Logger.log('[saveUserProfile] <<< Resultado correo: ' + enviado);
        if (enviado) {
          _marcarBienvenidaEnviada_(email);
        }
      } else {
        Logger.log('[saveUserProfile] Bienvenida ya enviada previamente — saltando');
      }
    } catch(emailErr) {
      // No queremos que un error de correo rompa el registro
      Logger.log('[saveUserProfile] ERROR (no crítico) en bloque de correo: ' + emailErr + ' | stack: ' + (emailErr.stack || ''));
    }

    var appUrl = getProductionUrl_();
    var targetPage = 'dashboard';
    if (rolLower === 'coach') targetPage = 'coach';
    else if (rolLower === 'admin' || rolLower === 'organizador') targetPage = 'admin';
    else if (rolLower === 'participante') targetPage = 'participante';
    return { success: true, dashboardUrl: appUrl + '?page=' + targetPage + '&token=' + token };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

// ---- Insertar coach en tabla Coaches si no existe ----
function registrarEnTablaCoach_(ss, email, nombreCompleto, institucion, telefono) {
  var sheet = ss.getSheetByName('Coaches');
  if (!sheet) {
    sheet = ss.insertSheet('Coaches');
    var h = ['idCoach','Nombre','Apellido','Correo','Telefono','Institucion','FechaRegistro'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }

  // Verificar si ya existe por correo
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var correoIdx = headers.indexOf('Correo');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][correoIdx]).trim().toLowerCase() === email.toLowerCase()) {
      return; // Ya existe, no duplicar
    }
  }

  // Generar ID secuencial
  var newId = 'COA-' + String(data.length).padStart(3, '0');
  var partes = nombreCompleto.split(' ');
  var nombre = partes[0] || '';
  var apellido = partes.slice(1).join(' ') || '';

  sheet.appendRow([newId, nombre, apellido, email, telefono || '', institucion || '', new Date()]);
}

// ---- Insertar organizador en tabla Organizadores si no existe ----
function registrarEnTablaOrganizador_(ss, email, nombreCompleto, institucion, telefono) {
  var sheet = ss.getSheetByName('Organizadores');
  if (!sheet) {
    sheet = ss.insertSheet('Organizadores');
    var h = ['idOrganizador','Nombre','Apellido','Correo','Telefono','Institucion','FechaRegistro'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var correoIdx = headers.indexOf('Correo');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][correoIdx]).trim().toLowerCase() === email.toLowerCase()) return;
  }
  var newId = 'ORG-' + String(data.length).padStart(3, '0');
  var partes = nombreCompleto.split(' ');
  var nombre = partes[0] || '';
  var apellido = partes.slice(1).join(' ') || '';
  sheet.appendRow([newId, nombre, apellido, email, telefono || '', institucion || '', new Date()]);
}

// ---- Insertar participante en tabla Participantes si no existe ----
function registrarEnTablaParticipante_(ss, email, nombreCompleto, datos) {
  var sheet = ss.getSheetByName('Participantes');
  if (!sheet) {
    sheet = ss.insertSheet('Participantes');
    var h = ['idParticipante','Nombre','Apellido','Edad','Correo','Escolaridad','CoachEmail','FechaRegistro'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }

  // Verificar si ya existe por correo
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var correoIdx = headers.indexOf('Correo');
  if (correoIdx === -1) correoIdx = headers.indexOf('Email');
  for (var i = 1; i < data.length; i++) {
    if (correoIdx !== -1 && String(data[i][correoIdx]).trim().toLowerCase() === email.toLowerCase()) {
      return; // Ya existe, no duplicar
    }
  }

  // Generar ID secuencial
  var newId = 'PAR-' + String(data.length).padStart(3, '0');
  var partes = nombreCompleto.split(' ');
  var nombre = partes[0] || '';
  var apellido = partes.slice(1).join(' ') || '';

  sheet.appendRow([
    newId,
    nombre,
    apellido,
    datos.edad || '',
    email,
    datos.escolaridad || datos.institucion || '',
    '', // CoachEmail se asigna después cuando un coach lo registre
    new Date()
  ]);
}

// ==========================================
// LOGOUT
// ==========================================
function logout(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
  return getProductionUrl_();
}

function getLoginUrl() {
  return getProductionUrl_();
}

// ==========================================
// VERIFICAR SESIÓN
// ==========================================
function verificarSesion(token) {
  var session = getSessionUser_(token);
  if (!session) return { valid: false };
  var user = findUserByEmail(session.email);
  if (!user) return { valid: false };
  return {
    valid: true,
    rol: (user['Rol'] || '').toLowerCase(),
    nombre: user['Nombre'] || session.name,
    email: session.email
  };
}

function cerrarSesion(token) {
  if (token) CacheService.getScriptCache().remove('s_' + token);
}

function getAppUrl() {
  return getProductionUrl_();
}

// ---- PARTICIPANTES (CRUD para Coach) ----
// Participantes que pertenecen a equipos del coach logeado (para tab "Mis Participantes")
function obtenerParticipantes(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  var em = String(session.email || '').toLowerCase();
  return _cached_('pcoach:v1:' + em, function() { return _obtenerParticipantesImpl_(session); });
}
function _obtenerParticipantesImpl_(session) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];
    var miEmail = String(session.email || '').toLowerCase();
    var coachIdx = headers.indexOf('CoachEmail');

    // Solo los participantes cuyo CoachEmail = mi email.
    // (Antes también se incluían por estar en uno de mis equipos, lo cual
    // provocaba que un participante ajeno apareciera como propio si en algún
    // momento se le agregó por error a un equipo del coach.)
    if (coachIdx === -1) return { success: true, data: [] };

    // IDs de equipos del coach actual (para marcar "enEquipo" en cada participante)
    var eqSheet = ss.getSheetByName('Equipos');
    var misEquipos = {};
    if (eqSheet) {
      var eqData = eqSheet.getDataRange().getValues();
      var eqH = eqData[0];
      var ceIdx = eqH.indexOf('CoachEmail');
      var eIdIdx = eqH.indexOf('idEquipo');
      for (var i = 1; i < eqData.length; i++) {
        if (String(eqData[i][ceIdx] || '').toLowerCase() === miEmail) {
          misEquipos[String(eqData[i][eIdIdx])] = true;
        }
      }
    }
    var misParticipantesEnEquipo = {};
    var epSheet = ss.getSheetByName('Equipo_Participante');
    if (epSheet) {
      var epData = epSheet.getDataRange().getValues();
      var epH = epData[0];
      var epEIdx = epH.indexOf('idEquipo');
      var epPIdx = epH.indexOf('idParticipante');
      for (var j = 1; j < epData.length; j++) {
        if (misEquipos[String(epData[j][epEIdx])]) {
          misParticipantesEnEquipo[String(epData[j][epPIdx])] = true;
        }
      }
    }

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var coachP = String(data[i][coachIdx] || '').toLowerCase();
      if (coachP !== miEmail) continue;
      var pid = String(data[i][headers.indexOf('idParticipante')] || '');
      rows.push({
        idParticipante: pid,
        nombre: data[i][headers.indexOf('Nombre')] || '',
        apellido: data[i][headers.indexOf('Apellido')] || '',
        edad: data[i][headers.indexOf('Edad')] || '',
        correo: data[i][headers.indexOf('Correo')] || '',
        escolaridad: data[i][headers.indexOf('Escolaridad')] || '',
        coachEmail: data[i][coachIdx] || '',
        fechaRegistro: _dStr_(data[i][headers.indexOf('FechaRegistro')]),
        enEquipo: !!misParticipantesEnEquipo[pid],
        _row: i + 1
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- HELPER: normalizar texto (sin acentos, minúsculas, trim) ----
function _norm_(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// ---- COACH: buscar usuarios registrados que aún NO son participantes del coach ----
// query: texto libre. Búsqueda multi-token (cada palabra debe aparecer en cualquier orden),
// insensible a acentos y mayúsculas. Busca en nombre, correo e institución.
// Devuelve hasta 50 resultados.
function buscarUsuariosDisponibles(token, query) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  try {
    var q = _norm_(query);
    var tokens = q ? q.split(' ').filter(function(t){ return t.length > 0; }) : [];
    var miEmail = String(session.email || '').toLowerCase();
    var ss = getSpreadsheet();
    var uSheet = getUsersSheet();
    var uData = uSheet.getDataRange().getValues();
    if (uData.length < 2) return { success: true, data: [], total: 0 };
    var uH = uData[0];
    var emIdx = uH.indexOf('Correo');
    var noIdx = uH.indexOf('Nombre');
    var insIdx = uH.indexOf('Institución');
    var rolIdx = uH.indexOf('Rol');

    // Conjunto de correos de Participantes ya asignados a ESTE coach
    var yaAsignadosAMi = {};
    var pSheet = ss.getSheetByName('Participantes');
    if (pSheet) {
      var pData = pSheet.getDataRange().getValues();
      if (pData.length >= 2) {
        var pH = pData[0];
        var pCorIdx = pH.indexOf('Correo');
        var pCoachIdx = pH.indexOf('CoachEmail');
        for (var k = 1; k < pData.length; k++) {
          var coachP = pCoachIdx !== -1 ? String(pData[k][pCoachIdx] || '').toLowerCase() : '';
          if (coachP === miEmail) {
            yaAsignadosAMi[String(pData[k][pCorIdx] || '').toLowerCase()] = true;
          }
        }
      }
    }

    var rows = [];
    var totalElegibles = 0;  // usuarios que pasarían el filtro de rol y no-asignado-a-mí
    for (var i = 1; i < uData.length; i++) {
      var em = String(uData[i][emIdx] || '').toLowerCase();
      if (!em) continue;
      if (em === miEmail) continue;
      if (yaAsignadosAMi[em]) continue;
      var nombre = String(uData[i][noIdx] || '');
      var inst = insIdx !== -1 ? String(uData[i][insIdx] || '') : '';
      var rol = rolIdx !== -1 ? String(uData[i][rolIdx] || '') : '';
      var rolLow = rol.toLowerCase();
      if (rolLow === 'coach' || rolLow === 'admin' || rolLow === 'organizador') continue;
      totalElegibles++;

      if (tokens.length) {
        var hay = _norm_(nombre + ' ' + em + ' ' + inst);
        var match = true;
        for (var t = 0; t < tokens.length; t++) {
          if (hay.indexOf(tokens[t]) === -1) { match = false; break; }
        }
        if (!match) continue;
      }
      if (rows.length < 50) {
        rows.push({
          correo: em,
          nombre: nombre,
          institucion: inst,
          rol: rol
        });
      }
    }
    return { success: true, data: rows, total: totalElegibles };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// ---- COACH: agregar un usuario ya registrado como participante propio ----
function agregarParticipanteExistente(token, correo) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var em = String(correo || '').trim().toLowerCase();
    if (!em) return { success: false, error: 'Correo requerido' };
    if (em === String(session.email).toLowerCase()) {
      return { success: false, error: 'No puedes agregarte a ti mismo' };
    }

    // Buscar el usuario en la hoja Usuarios
    var user = findUserByEmail(em);
    if (!user) return { success: false, error: 'No existe un usuario registrado con ese correo' };

    var nombreCompleto = String(user['Nombre'] || em.split('@')[0]).trim();
    var partes = nombreCompleto.split(' ');
    var nombre = partes[0] || '';
    var apellido = partes.slice(1).join(' ') || '';
    var institucion = String(user['Institución'] || '').trim();

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) {
      sheet = ss.insertSheet('Participantes');
      var h = ['idParticipante','Nombre','Apellido','Edad','Correo','Escolaridad','CoachEmail','FechaRegistro'];
      sheet.getRange(1, 1, 1, h.length).setValues([h]);
      sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }

    // ¿Ya existe ese correo en Participantes?
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var corIdx = headers.indexOf('Correo');
    var coachIdx = headers.indexOf('CoachEmail');
    var nomIdx = headers.indexOf('Nombre');
    var apeIdx = headers.indexOf('Apellido');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][corIdx] || '').trim().toLowerCase() === em) {
        // Ya existe — solo (re)asignar a este coach si no estaba o estaba con otro
        var existing = String(data[i][coachIdx] || '').trim().toLowerCase();
        if (existing === String(session.email).toLowerCase()) {
          return { success: false, error: 'Este participante ya está asignado a ti' };
        }
        sheet.getRange(i + 1, coachIdx + 1).setValue(session.email);
        return { success: true, idParticipante: String(data[i][headers.indexOf('idParticipante')] || ''), reasignado: true };
      }
    }

    // No existe → crear nueva fila
    var newId = 'PAR-' + String(data.length).padStart(3, '0');
    sheet.appendRow([
      newId, nombre, apellido,
      '', em, institucion,
      session.email, new Date()
    ]);

    // Asegurar que su rol en Usuarios sea 'participante' si está vacío
    try {
      var rol = String(user['Rol'] || '').toLowerCase();
      if (!rol) {
        var uSheet = getUsersSheet();
        var uHeaders = uSheet.getRange(1, 1, 1, uSheet.getLastColumn()).getValues()[0];
        var rolCol = uHeaders.indexOf('Rol') + 1;
        if (rolCol > 0 && user._row) uSheet.getRange(user._row, rolCol).setValue('participante');
      }
    } catch(_) {}

    return { success: true, idParticipante: newId };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Participantes disponibles para el modal de equipos: SOLO los del coach actual.
// (El nombre conserva "Todos" por compatibilidad con código viejo; ahora está scoped al coach.)
function obtenerTodosParticipantes(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  var em = String(session.email || '').toLowerCase();
  return _cached_('tpa:v1:' + em, function() { return _obtenerTodosParticipantesImpl_(session); });
}
function _obtenerTodosParticipantesImpl_(session) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];
    var miEmail = String(session.email || '').toLowerCase();
    var idIdx = headers.indexOf('idParticipante');
    var nomIdx = headers.indexOf('Nombre');
    var apeIdx = headers.indexOf('Apellido');
    var corIdx = headers.indexOf('Correo');
    var coachIdx = headers.indexOf('CoachEmail');
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var coachP = coachIdx !== -1 ? String(data[i][coachIdx] || '').toLowerCase() : '';
      if (coachP !== miEmail) continue;
      rows.push({
        idParticipante: data[i][idIdx] || '',
        nombre: data[i][nomIdx] || '',
        apellido: data[i][apeIdx] || '',
        correo: data[i][corIdx] || ''
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function registrarParticipante(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) {
      sheet = ss.insertSheet('Participantes');
      var h = ['idParticipante','Nombre','Apellido','Edad','Correo','Escolaridad','CoachEmail','FechaRegistro'];
      sheet.getRange(1, 1, 1, h.length).setValues([h]);
      sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var correoIdx = headers.indexOf('Correo');
    var correo = String(datos.correo || '').trim().toLowerCase();
    if (!correo) return { success: false, error: 'El correo es requerido' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][correoIdx]).trim().toLowerCase() === correo) {
        return { success: false, error: 'Ya existe un participante con ese correo' };
      }
    }
    var newId = 'PAR-' + String(data.length).padStart(3, '0');
    sheet.appendRow([
      newId,
      datos.nombre || '',
      datos.apellido || '',
      datos.edad || '',
      datos.correo,
      datos.escolaridad || '',
      session.email,
      new Date()
    ]);

    // También crear/actualizar entrada en Usuarios para que pueda iniciar sesión
    try {
      var usuariosSheet = getUsersSheet();
      var existing = findUserByEmail(datos.correo);
      if (!existing) {
        usuariosSheet.appendRow([datos.correo, (datos.nombre||'') + ' ' + (datos.apellido||''), '', datos.escolaridad || '', 'participante', '', '', '', new Date()]);
      }
    } catch(e) {}

    // Correo de bienvenida al nuevo participante
    _enviarCorreoBienvenida_(datos.correo, (datos.nombre||'') + ' ' + (datos.apellido||''), 'participante');

    return { success: true, idParticipante: newId };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function editarParticipante(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var sheet = getSpreadsheet().getSheetByName('Participantes');
    if (!sheet) return { success: false, error: 'No existe la hoja Participantes' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idParticipante');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(datos.idParticipante)) {
        if (datos.nombre !== undefined) sheet.getRange(i + 1, headers.indexOf('Nombre') + 1).setValue(datos.nombre);
        if (datos.apellido !== undefined) sheet.getRange(i + 1, headers.indexOf('Apellido') + 1).setValue(datos.apellido);
        if (datos.edad !== undefined) sheet.getRange(i + 1, headers.indexOf('Edad') + 1).setValue(datos.edad);
        if (datos.escolaridad !== undefined) sheet.getRange(i + 1, headers.indexOf('Escolaridad') + 1).setValue(datos.escolaridad);
        return { success: true };
      }
    }
    return { success: false, error: 'Participante no encontrado' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function eliminarParticipante(token, idParticipante) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) return { success: false, error: 'No existe la hoja' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idParticipante');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIdx]) === String(idParticipante)) {
        sheet.deleteRow(i + 1);
        // Limpiar también en Equipo_Participante
        var epSheet = ss.getSheetByName('Equipo_Participante');
        if (epSheet) {
          var epData = epSheet.getDataRange().getValues();
          var epH = epData[0];
          var pIdx = epH.indexOf('idParticipante');
          for (var j = epData.length - 1; j >= 1; j--) {
            if (String(epData[j][pIdx]) === String(idParticipante)) {
              epSheet.deleteRow(j + 1);
            }
          }
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Participante no encontrado' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- EQUIPOS (CRUD para Coach) ----
// Helpers internos

// Convierte Date a string ISO (google.script.run NO puede serializar Date)
function _dStr_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
  return v == null ? '' : v;
}

// ==========================================
// CACHÉ DE LECTURAS PESADAS (TTL 30s)
// ==========================================
var _CACHE_TTL_ = 30;
var _CACHE_KEYS_ADMIN_ = ['rpt:v1','evt:v1','ins:v1','coa:v1','pad:v1','tpa:v1','cat:v1'];

function _cached_(key, fn) {
  var cache = CacheService.getScriptCache();
  try {
    var hit = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch(_) {} }
  } catch(_) {}
  var fresh = fn();
  try {
    if (fresh && fresh.success) {
      var s = JSON.stringify(fresh);
      if (s && s.length < 100000) cache.put(key, s, _CACHE_TTL_);
    }
  } catch(_) {}
  return fresh;
}

function _bustCache_() {
  try { CacheService.getScriptCache().removeAll(_CACHE_KEYS_ADMIN_); } catch(_) {}
}

function _bustCacheCoach_(email) {
  if (!email) return _bustCache_();
  var em = String(email).toLowerCase();
  var keys = _CACHE_KEYS_ADMIN_.concat(['eq:v1:'+em,'pcoach:v1:'+em,'insC:v1:'+em]);
  try { CacheService.getScriptCache().removeAll(keys); } catch(_) {}
}

// ==========================================
// BUNDLED ENDPOINTS — devuelven todos los datos del panel en UN viaje
// Reduce 5-6 round-trips a 1 sólo. Aprovecha caché TTL=30s en cada hoja.
// ==========================================

function cargarPanelAdmin(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  try {
    var u = findUserByEmail(session.email);
    var rol = (u && u['Rol'] || '').toLowerCase();
    if (rol !== 'admin' && rol !== 'organizador') {
      return { success: false, error: 'No autorizado' };
    }
    return {
      success: true,
      reportes:      _cached_('rpt:v1', function() { return _obtenerReportesImpl_(); }),
      eventos:       _cached_('evt:v1', function() { return _obtenerEventosImpl_(); }),
      inscripciones: _cached_('ins:v1', function() { return _obtenerTodasInscripcionesImpl_(); }),
      coaches:       _cached_('coa:v1', function() { return _obtenerCoachesImpl_(); }),
      participantes: _cached_('pad:v1', function() { return _obtenerParticipantesAdminImpl_(); }),
      categorias:    _cached_('cat:v1', function() { return _obtenerCategoriasImpl_(); })
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function cargarPanelCoach(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  try {
    var em = String(session.email || '').toLowerCase();
    return {
      success: true,
      participantes: _cached_('pcoach:v1:'+em, function() { return _obtenerParticipantesImpl_(session); }),
      equipos:       _cached_('eq:v1:'+em,    function() { return _obtenerEquiposImpl_(session); }),
      eventos:       _cached_('evt:v1',       function() { return _obtenerEventosImpl_(); }),
      inscripciones: _cached_('insC:v1:'+em,  function() { return _obtenerInscripcionesCoachImpl_(token, session); }),
      categorias:    _cached_('cat:v1',       function() { return _obtenerCategoriasImpl_(); })
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Genera próximo ID único con prefijo basado en el máximo existente
function _nextId_(sheet, prefix, colIdx) {
  var data = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var v = String(data[i][colIdx] || '');
    var m = v.match(/-(\d+)$/);
    if (m) { var n = parseInt(m[1], 10); if (n > max) max = n; }
  }
  return prefix + '-' + String(max + 1).padStart(3, '0');
}

function _ensureEquiposSheet_(ss) {
  var sheet = ss.getSheetByName('Equipos');
  if (!sheet) {
    sheet = ss.insertSheet('Equipos');
    var h = ['idEquipo','NombreEquipo','idCategoria','CoachEmail','FechaCreacion'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function _ensureEquipoParticipanteSheet_(ss) {
  var sheet = ss.getSheetByName('Equipo_Participante');
  if (!sheet) {
    sheet = ss.insertSheet('Equipo_Participante');
    var h = ['idEquipo','idParticipante','FechaAsignacion'];
    sheet.getRange(1, 1, 1, h.length).setValues([h]);
    sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function obtenerEquipos(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  var em = String(session.email || '').toLowerCase();
  return _cached_('eq:v1:' + em, function() { return _obtenerEquiposImpl_(session); });
}
function _obtenerEquiposImpl_(session) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Equipos');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];

    // Cargar mapas de categorías
    var catSheet = getCatSheet_(ss);
    var catNames = {};
    if (catSheet) {
      var catData = catSheet.getDataRange().getValues();
      var catH = catData[0];
      for (var i = 1; i < catData.length; i++) {
        catNames[String(catData[i][catH.indexOf('idCategoria')])] = catData[i][catH.indexOf('Nombre')] || '';
      }
    }

    // Cargar relación equipo-participante
    var epSheet = ss.getSheetByName('Equipo_Participante');
    var equipoMiembros = {};
    if (epSheet) {
      var epData = epSheet.getDataRange().getValues();
      var epH = epData[0];
      var eIdx = epH.indexOf('idEquipo');
      var pIdx = epH.indexOf('idParticipante');
      for (var i = 1; i < epData.length; i++) {
        var eId = String(epData[i][eIdx]);
        if (!equipoMiembros[eId]) equipoMiembros[eId] = [];
        equipoMiembros[eId].push(String(epData[i][pIdx]));
      }
    }

    // Cargar nombres de participantes
    var partSheet = ss.getSheetByName('Participantes');
    var partNames = {};
    if (partSheet) {
      var partData = partSheet.getDataRange().getValues();
      var partH = partData[0];
      var pIdIdx = partH.indexOf('idParticipante');
      var pNomIdx = partH.indexOf('Nombre');
      var pApeIdx = partH.indexOf('Apellido');
      for (var i = 1; i < partData.length; i++) {
        partNames[String(partData[i][pIdIdx])] = {
          nombre: partData[i][pNomIdx] || '',
          apellido: partData[i][pApeIdx] || ''
        };
      }
    }

    var coachIdx = headers.indexOf('CoachEmail');
    var idIdx = headers.indexOf('idEquipo');
    var nomIdx = headers.indexOf('NombreEquipo');
    var catIdx = headers.indexOf('idCategoria');
    var fchIdx = headers.indexOf('FechaCreacion');

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][coachIdx]).toLowerCase() !== session.email.toLowerCase()) continue;
      var idEq = String(data[i][idIdx]);
      var miembrosIds = equipoMiembros[idEq] || [];
      var miembros = miembrosIds.map(function(pid) {
        var p = partNames[pid] || { nombre: pid, apellido: '' };
        return { idParticipante: pid, nombre: p.nombre, apellido: p.apellido };
      });
      var catId = String(data[i][catIdx] || '');
      rows.push({
        idEquipo: idEq,
        nombreEquipo: data[i][nomIdx] || '',
        idCategoria: catId,
        categoriaNombre: catNames[catId] || catId || '',
        coachEmail: data[i][coachIdx] || '',
        fechaCreacion: _dStr_(data[i][fchIdx]),
        miembros: miembros
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ============ HELPERS DE CORREO (CRODIT UABC) ============
// Plantilla HTML enriquecida usada en todos los correos.
function _emailTemplate_(titulo, subtitulo, saludo, parrafos, secciones) {
  var parrafosHtml = (parrafos || []).map(function(p){
    return '<p style="color:rgba(255,255,255,0.75);font-size:16px;line-height:1.6;margin:0 0 16px;">' + p + '</p>';
  }).join('');

  var seccionesHtml = (secciones || []).map(function(sec){
    var filas = (sec.filas || []).map(function(f){
      return '<tr><td style="padding:10px 0;color:rgba(255,255,255,0.5);width:130px;font-weight:500;">' + f.label + ':</td>' +
             '<td style="padding:10px 0;color:#ffffff;font-weight:600;">' + f.valor + '</td></tr>';
    }).join('');
    var border = sec.destacada ? '2px solid #3b82f6' : '1px solid rgba(255,255,255,0.1)';
    var bg = sec.destacada ? 'rgba(59,130,246,0.08)' : 'rgba(255,255,255,0.03)';
    return '<div style="background:' + bg + ';border:' + border + ';border-radius:20px;padding:28px;margin:24px 0;">' +
      '<h3 style="color:#3b82f6;margin:0 0 16px;font-size:18px;font-weight:600;">' + sec.titulo + '</h3>' +
      '<table style="width:100%;border-collapse:collapse;">' + filas + '</table>' +
      '</div>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0"></head>' +
    '<body style="font-family:\'Segoe UI\',Arial,sans-serif;margin:0;padding:20px;background-color:#f5f7fa;">' +
    '<div style="max-width:600px;margin:0 auto;background:#0f172a;border-radius:24px;overflow:hidden;box-shadow:0 20px 40px rgba(0,0,0,0.4);border:1px solid rgba(59,130,246,0.3);">' +
      '<div style="background:linear-gradient(135deg,#3b82f6 0%,#1e3a8a 100%);padding:36px 30px;text-align:center;">' +
        '<img src="https://drive.google.com/thumbnail?id=1J-ZAItJLayReY_BfQsSAJwkHcv2xYmoW&sz=w800" alt="CRODT-UABC" style="height:56px;width:auto;object-fit:contain;display:block;margin:0 auto 4px;max-width:100%"/>' +
        '<h1 style="color:#fff;margin:0;font-size:18px;font-weight:600;letter-spacing:1px;opacity:.85">CRODT · UABC</h1>' +
        '<p style="color:rgba(255,255,255,0.9);margin:12px 0 0;font-size:17px;">' + subtitulo + '</p>' +
        '<div style="width:50px;height:3px;background:#06b6d4;margin:14px auto 0;"></div>' +
      '</div>' +
      '<div style="padding:36px 30px;">' +
        '<h2 style="color:#fff;margin:0 0 20px;font-size:22px;font-weight:600;">' + saludo + '</h2>' +
        parrafosHtml +
        seccionesHtml +
        '<div style="margin-top:32px;padding-top:20px;border-top:1px solid rgba(255,255,255,0.1);text-align:center;">' +
          '<p style="color:rgba(255,255,255,0.4);font-size:12px;margin:4px 0;">Este es un correo automático, por favor no respondas.</p>' +
          '<p style="color:rgba(255,255,255,0.4);font-size:12px;margin:4px 0;">© CRODIT — Centro de Robótica y Diseño Industrial Tecnológico · Facultad de Ingeniería Mexicali, UABC</p>' +
        '</div>' +
      '</div>' +
    '</div></body></html>';
}

function _enviarCorreoBienvenida_(email, nombre, rol) {
  Logger.log('[Correo Bienvenida] Iniciando envío a: ' + email + ' | nombre: ' + nombre + ' | rol: ' + rol);
  if (!email) {
    Logger.log('[Correo Bienvenida] ABORTADO: email vacío');
    return false;
  }
  try {
    var nombreCorto = String(nombre || '').split(' ')[0] || '';
    var rolMap = { coach:'Coach', participante:'Participante', organizador:'Organizador', admin:'Administrador' };
    var rolTexto = rolMap[String(rol||'').toLowerCase()] || 'Usuario';
    var saludo = nombreCorto ? 'Hola, ' + nombreCorto : 'Hola';

    var cuerpoHTML = _emailTemplate_(
      'Registro confirmado',
      'Registro confirmado',
      saludo,
      [
        'Has completado exitosamente tu registro en el sistema de competencias de <strong>CRODIT UABC</strong>.',
        'Estás registrado como <strong style="color:#06b6d4;">' + rolTexto + '</strong>.'
      ],
      [{
        titulo: 'Próximos pasos',
        filas: [
          { label: 'Estado', valor: 'Activo' },
          { label: 'Rol', valor: rolTexto },
          { label: 'Siguiente paso', valor: rol === 'coach' ? 'Crea tu equipo desde el panel' : (rol === 'participante' ? 'Espera a que tu coach te asigne a un equipo' : 'Accede al panel desde la app') }
        ]
      }]
    );

    var cuerpoTexto =
      saludo + ',\n\n' +
      'Has completado tu registro en CRODIT UABC.\n' +
      'Rol: ' + rolTexto + '\n\n' +
      'Pronto recibirás más información sobre las competencias.\n\n' +
      '— Equipo CRODIT FIM';

    GmailApp.sendEmail(
      email,
      '¡Bienvenido a CRODIT UABC!',
      cuerpoTexto,
      { htmlBody: cuerpoHTML, name: 'CRODIT UABC' }
    );
    Logger.log('[Correo Bienvenida] ENVIADO correctamente a: ' + email);
    return true;
  } catch (e) {
    Logger.log('[Correo Bienvenida] ERROR al enviar a ' + email + ': ' + e + ' | stack: ' + (e.stack || ''));
    return false;
  }
}

function _enviarCorreoEquipoAgregado_(email, nombre, nombreEquipo) {
  Logger.log('[Correo Equipo] Iniciando envío a: ' + email + ' | equipo: ' + nombreEquipo);
  if (!email) {
    Logger.log('[Correo Equipo] ABORTADO: email vacío');
    return false;
  }
  try {
    var nombreCorto = String(nombre || '').split(' ')[0] || '';
    var saludo = nombreCorto ? 'Hola, ' + nombreCorto : 'Hola';

    var cuerpoHTML = _emailTemplate_(
      '¡Estás en un equipo!',
      '¡Estás en un equipo!',
      saludo,
      [
        'Tu coach te ha agregado al equipo <strong style="color:#06b6d4;">' + nombreEquipo + '</strong>.',
        'Ya formas parte oficial del equipo en las competencias <strong>CRODIT UABC</strong>.'
      ],
      [{
        titulo: 'Detalles',
        destacada: true,
        filas: [
          { label: 'Equipo', valor: nombreEquipo },
          { label: 'Estado', valor: 'Asignado' }
        ]
      }]
    );

    var cuerpoTexto =
      saludo + ',\n\n' +
      'Tu asesor te agregó al equipo "' + nombreEquipo + '" para las competencias CRODIT UABC.\n\n' +
      '¡Mucho éxito!\n\n' +
      '— Equipo CRODIT FIM';

    GmailApp.sendEmail(
      email,
      '¡Fuiste agregado al equipo ' + nombreEquipo + ' — CRODIT UABC!',
      cuerpoTexto,
      { htmlBody: cuerpoHTML, name: 'CRODIT UABC' }
    );
    Logger.log('[Correo Equipo] ENVIADO correctamente a: ' + email);
    return true;
  } catch (e) {
    Logger.log('[Correo Equipo] ERROR al enviar a ' + email + ': ' + e + ' | stack: ' + (e.stack || ''));
    return false;
  }
}

function _enviarCorreoInscripcion_(coachEmail, nombreEquipo, nombreEvento, estado, motivo) {
  Logger.log('[Correo Inscripcion] Iniciando envío a: ' + coachEmail + ' | equipo: ' + nombreEquipo + ' | estado: ' + estado);
  if (!coachEmail) {
    Logger.log('[Correo Inscripcion] ABORTADO: email vacío');
    return false;
  }
  try {
    var aprobada = String(estado).toLowerCase() === 'aprobada';
    var titulo = aprobada ? 'Inscripción aprobada' : 'Inscripción rechazada';
    var subj = (aprobada ? 'Inscripción aprobada: ' : 'Inscripción rechazada: ') + nombreEquipo + ' en ' + nombreEvento;

    var parrafos = [];
    if (aprobada) {
      parrafos.push('La participación de tu equipo <strong style="color:#06b6d4;">' + nombreEquipo + '</strong> en el evento <strong>' + nombreEvento + '</strong> ha sido <strong style="color:#10b981;">aceptada</strong>.');
      parrafos.push('¡Felicidades! Prepara a tu equipo para competir.');
    } else {
      parrafos.push('La participación de tu equipo <strong style="color:#06b6d4;">' + nombreEquipo + '</strong> en el evento <strong>' + nombreEvento + '</strong> ha sido <strong style="color:#ef4444;">rechazada</strong>.');
      if (motivo) parrafos.push('<strong>Motivo:</strong> ' + motivo);
      parrafos.push('Si crees que es un error, contacta al organizador.');
    }

    var cuerpoHTML = _emailTemplate_(
      titulo,
      titulo,
      'Hola Coach',
      parrafos,
      [{
        titulo: 'Detalles de la inscripción',
        destacada: aprobada,
        filas: [
          { label: 'Equipo', valor: nombreEquipo },
          { label: 'Evento', valor: nombreEvento },
          { label: 'Estado', valor: aprobada ? 'Aprobada' : 'Rechazada' }
        ].concat(motivo && !aprobada ? [{ label: 'Motivo', valor: motivo }] : [])
      }]
    );

    var cuerpoTexto =
      'Hola Coach,\n\n' +
      'La inscripción de tu equipo "' + nombreEquipo + '" en "' + nombreEvento + '" ha sido ' + (aprobada ? 'APROBADA' : 'RECHAZADA') + '.\n' +
      (motivo && !aprobada ? 'Motivo: ' + motivo + '\n' : '') +
      '\n— Equipo CRODIT FIM';

    GmailApp.sendEmail(
      coachEmail,
      subj,
      cuerpoTexto,
      { htmlBody: cuerpoHTML, name: 'CRODIT UABC' }
    );
    Logger.log('[Correo Inscripcion] ENVIADO correctamente a: ' + coachEmail);
    return true;
  } catch (e) {
    Logger.log('[Correo Inscripcion] ERROR al enviar a ' + coachEmail + ': ' + e + ' | stack: ' + (e.stack || ''));
    return false;
  }
}

// Función de prueba — ejecútala desde el editor con tu propio correo para verificar.
function testEnviarCorreoBienvenida() {
  var miCorreo = Session.getActiveUser().getEmail();
  Logger.log('Probando envío a: ' + miCorreo);
  var ok = _enviarCorreoBienvenida_(miCorreo, 'Cristopher Vea', 'participante');
  Logger.log('Resultado: ' + ok);
  return ok;
}

// FUNCIÓN DE RECUPERACIÓN
// Recorre la hoja Usuarios y envía bienvenida a TODOS los que NO la tengan marcada.
// Ejecútala manualmente desde el editor para enviar correos pendientes.
function enviarBienvenidasPendientes() {
  Logger.log('=== ENVÍO DE BIENVENIDAS PENDIENTES — INICIO ===');
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('Hoja Usuarios vacía — nada que enviar');
    return { total: 0, enviados: 0, fallidos: 0 };
  }
  var headers = data[0];
  var idxCorreo = headers.indexOf('Correo');
  var idxNombre = headers.indexOf('Nombre');
  var idxRol = headers.indexOf('Rol');
  var idxBienv = headers.indexOf('BienvenidaEnviada');

  var enviados = 0, fallidos = 0, saltados = 0;
  for (var i = 1; i < data.length; i++) {
    var correo = String(data[i][idxCorreo] || '').trim();
    var nombre = String(data[i][idxNombre] || '').trim();
    var rol = String(data[i][idxRol] || '').trim().toLowerCase();
    var bienv = idxBienv >= 0 ? String(data[i][idxBienv] || '').trim().toLowerCase() : '';

    if (!correo) continue;
    if (bienv.indexOf('sí') === 0 || bienv.indexOf('si') === 0) {
      saltados++;
      continue;
    }

    Logger.log('--> Enviando a: ' + correo + ' (rol: ' + rol + ')');
    var ok = _enviarCorreoBienvenida_(correo, nombre, rol);
    if (ok) {
      enviados++;
      _marcarBienvenidaEnviada_(correo);
    } else {
      fallidos++;
    }
    Utilities.sleep(500); // pequeña pausa para no exceder cuotas
  }
  Logger.log('=== RESUMEN — Enviados: ' + enviados + ' | Fallidos: ' + fallidos + ' | Ya enviados antes: ' + saltados + ' ===');
  return { enviados: enviados, fallidos: fallidos, saltados: saltados, total: data.length - 1 };
}

// Carga {idParticipante: {email, nombre}} desde la hoja Participantes
function _cargarInfoParticipantes_(ss) {
  var partSheet = ss.getSheetByName('Participantes');
  var info = {};
  if (!partSheet) return info;
  var pd = partSheet.getDataRange().getValues();
  if (pd.length < 2) return info;
  var ph = pd[0];
  var piIdx = ph.indexOf('idParticipante');
  var pcIdx = ph.indexOf('Correo');
  var pnIdx = ph.indexOf('Nombre');
  for (var j = 1; j < pd.length; j++) {
    info[String(pd[j][piIdx])] = { email: pd[j][pcIdx]||'', nombre: pd[j][pnIdx]||'' };
  }
  return info;
}

function crearEquipo(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = _ensureEquiposSheet_(ss);
    var newId = _nextId_(sheet, 'EQ', 0);
    sheet.appendRow([newId, datos.nombreEquipo || '', datos.idCategoria || '', session.email, new Date()]);

    // Insertar miembros y enviar correo de notificación
    var miembros = datos.miembros || [];
    if (miembros.length > 0) {
      var epSheet = _ensureEquipoParticipanteSheet_(ss);
      var now = new Date();
      var partInfo = _cargarInfoParticipantes_(ss);

      miembros.forEach(function(idPart) {
        if (!idPart) return;
        epSheet.appendRow([newId, idPart, now]);
        var info = partInfo[idPart];
        if (info && info.email) {
          _enviarCorreoEquipoAgregado_(info.email, info.nombre, datos.nombreEquipo || '');
        }
      });
    }
    return { success: true, idEquipo: newId };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function editarEquipo(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Equipos');
    if (!sheet) return { success: false, error: 'No existe la hoja Equipos' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idEquipo');
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(datos.idEquipo)) {
        if (datos.nombreEquipo !== undefined) sheet.getRange(i + 1, headers.indexOf('NombreEquipo') + 1).setValue(datos.nombreEquipo);
        if (datos.idCategoria !== undefined) sheet.getRange(i + 1, headers.indexOf('idCategoria') + 1).setValue(datos.idCategoria);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Equipo no encontrado' };

    // Capturar miembros existentes ANTES de borrar (para detectar agregados nuevos)
    var existentes = {};
    var epSheet = ss.getSheetByName('Equipo_Participante');
    if (epSheet) {
      var epData = epSheet.getDataRange().getValues();
      var epH = epData[0];
      var eIdx = epH.indexOf('idEquipo');
      var pIdx = epH.indexOf('idParticipante');
      for (var j = epData.length - 1; j >= 1; j--) {
        if (String(epData[j][eIdx]) === String(datos.idEquipo)) {
          existentes[String(epData[j][pIdx])] = true;
          epSheet.deleteRow(j + 1);
        }
      }
    }
    var miembros = datos.miembros || [];
    if (miembros.length > 0) {
      epSheet = _ensureEquipoParticipanteSheet_(ss);
      var now = new Date();
      // Resolver nombre del equipo (puede venir en datos o leerse de la hoja)
      var nombreEquipo = datos.nombreEquipo;
      if (!nombreEquipo) {
        for (var k = 1; k < data.length; k++) {
          if (String(data[k][idIdx]) === String(datos.idEquipo)) {
            nombreEquipo = data[k][headers.indexOf('NombreEquipo')] || '';
            break;
          }
        }
      }
      var partInfo = _cargarInfoParticipantes_(ss);
      miembros.forEach(function(idPart) {
        if (!idPart) return;
        epSheet.appendRow([datos.idEquipo, idPart, now]);
        // Sólo notificar a los recién agregados
        if (!existentes[String(idPart)]) {
          var info = partInfo[idPart];
          if (info && info.email) {
            _enviarCorreoEquipoAgregado_(info.email, info.nombre, nombreEquipo || '');
          }
        }
      });
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function eliminarEquipo(token, idEquipo) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Equipos');
    if (!sheet) return { success: false, error: 'No existe la hoja' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idEquipo');
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIdx]) === String(idEquipo)) {
        sheet.deleteRow(i + 1);
        // Limpiar Equipo_Participante
        var epSheet = ss.getSheetByName('Equipo_Participante');
        if (epSheet) {
          var epData = epSheet.getDataRange().getValues();
          var epH = epData[0];
          var eIdx = epH.indexOf('idEquipo');
          for (var j = epData.length - 1; j >= 1; j--) {
            if (String(epData[j][eIdx]) === String(idEquipo)) {
              epSheet.deleteRow(j + 1);
            }
          }
        }
        return { success: true };
      }
    }
    return { success: false, error: 'Equipo no encontrado' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- HELPER: buscar hoja Categorias (con o sin tilde) ----
function getCatSheet_(ss) {
  return ss.getSheetByName('Categorías') || ss.getSheetByName('Categorias');
}

// ---- SEED: 10 categorías oficiales CRODIT (desde PDFs UABC FIM) ----
function _seedCategoriasCRODIT_(sheet) {
  var rows = [
    // SEGUIDOR DE LÍNEA
    ['CAT-001','Seguidor de Línea Amateur','Seguidor de Línea',
      'Robot autónomo que sigue una trayectoria de línea negra sobre plataforma blanca en el menor tiempo posible.',
      'Tamaño 150x150 mm | Peso 500 gr | Motores Amarillos 48-1 (207 RPM @6V) | Batería 11.1V o 3s Max | Sin turbina',
      4,1],
    ['CAT-002','Seguidor de Línea Profesional','Seguidor de Línea',
      'Robot autónomo seguidor de línea de categoría profesional, sin uso de turbina.',
      'Tamaño 250x250 mm | Peso, motores, llantas, batería y sensores sin restricción | Sin turbina',
      4,1],
    ['CAT-003','Seguidor de Línea Profesional (Turbina)','Seguidor de Línea',
      'Robot autónomo seguidor de línea profesional con sistema de tracción aumentada (turbina).',
      'Tamaño 250x250 mm | Sin restricción en peso, motores, llantas, batería y sensores | Permite turbina',
      4,1],
    // SUMO
    ['CAT-004','Minisumo Amateur','Sumo',
      'Robot autónomo de combate sumo categoría amateur. Empuja al oponente fuera del ring circular.',
      'Tamaño 150x150 mm | Peso 500 gr | Motores Amarillos 48-1 (207 RPM @6V) | Batería 11.1V o 3s Max | Sin succión',
      4,1],
    ['CAT-005','Minisumo Profesional','Sumo',
      'Robot autónomo de combate sumo categoría profesional.',
      'Tamaño 100x100 mm | Peso 500 gr | Motores, llantas y batería sin restricción | Sin succión',
      4,1],
    ['CAT-006','MinisumoRC','Sumo',
      'Robot de combate sumo miniatura controlado por radiocontrol.',
      'Tamaño 100x100 mm | Peso 500 gr | Sin restricción en motores, llantas y batería | Sin succión | Operación RC',
      4,1],
    ['CAT-007','Sumo RC 20kg','Sumo',
      'Robot de combate sumo grande RC. Permite sistema de succión.',
      'Tamaño 50x50 cm | Peso 20 kg | Sin restricción en motores, llantas y batería | Permite sistema de succión | Operación RC',
      4,1],
    // LABERINTO
    ['CAT-008','Laberinto Amateur','Laberinto',
      'Robot autónomo que recorre un laberinto de inicio a fin, en el menor tiempo posible.',
      'Tamaño 150x150 mm | Peso 500 gr | Motores Amarillos 48-1 (207 RPM @6V) | Batería 11.1V o 3s Max | Sensores ultrasónicos u ópticos',
      4,1],
    ['CAT-009','Laberinto Profesional','Laberinto',
      'Robot autónomo solucionador de laberintos categoría profesional.',
      'Tamaño 100x100 mm | Peso, motores, llantas, batería y sensores sin restricción',
      4,1],
    // ROBOFUT
    ['CAT-010','Robofut','Fútbol Robótico',
      'Partido de fútbol con robots RC. Hasta 2 robots por equipo.',
      'Tamaño 200x200 mm | Peso 1500 gr | Sin restricción en motores, llantas y batería | Operación RC | Máx 2 robots por equipo',
      4,2]
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// ---- ADMIN: Reset categorías a las 10 oficiales CRODIT ----
// Borra TODAS las categorías existentes y recarga las 10 del PDF.
// Cascada: borra equipos, integrantes asignados, inscripciones y vínculos evento-categoría.
function resetCategoriasCRODIT(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  // Solo admin/organizador puede ejecutar
  var u = findUserByEmail(session.email);
  var rol = (u && u['Rol'] || '').toLowerCase();
  if (rol !== 'admin' && rol !== 'organizador') {
    return { success: false, error: 'No autorizado. Solo admin/organizador puede recargar categorías.' };
  }
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = getCatSheet_(ss);

    // Si no existe, crear con encabezados nuevos
    if (!sheet) {
      sheet = ss.insertSheet('Categorias');
      var h = ['idCategoria','Nombre','Disciplina','Descripcion','Restricciones','MaxIntegrantes','MinIntegrantes'];
      sheet.getRange(1, 1, 1, h.length).setValues([h]);
      sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      sheet.setFrozenRows(1);
    } else {
      // Verificar/actualizar esquema (agregar columnas Disciplina y Restricciones si faltan)
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var nuevos = ['idCategoria','Nombre','Disciplina','Descripcion','Restricciones','MaxIntegrantes','MinIntegrantes'];
      var hayQueRecrear = false;
      for (var k = 0; k < nuevos.length; k++) {
        if (headers.indexOf(nuevos[k]) === -1) { hayQueRecrear = true; break; }
      }
      if (hayQueRecrear) {
        // Reescribir encabezados
        ss.deleteSheet(sheet);
        sheet = ss.insertSheet('Categorias');
        var h2 = ['idCategoria','Nombre','Disciplina','Descripcion','Restricciones','MaxIntegrantes','MinIntegrantes'];
        sheet.getRange(1, 1, 1, h2.length).setValues([h2]);
        sheet.getRange(1, 1, 1, h2.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
        sheet.setFrozenRows(1);
      } else if (sheet.getLastRow() > 1) {
        // Borrar todas las filas de datos (conserva encabezado)
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    }

    // Cascada: borrar equipos, integrantes, inscripciones y vínculos evento-categoría
    var eqSheet = ss.getSheetByName('Equipos');
    var equiposIds = [];
    if (eqSheet && eqSheet.getLastRow() > 1) {
      var eqData = eqSheet.getDataRange().getValues();
      var eqH = eqData[0];
      var eIdIdx = eqH.indexOf('idEquipo');
      for (var j = 1; j < eqData.length; j++) {
        equiposIds.push(String(eqData[j][eIdIdx]));
      }
      // Borrar todas las filas de equipos
      eqSheet.deleteRows(2, eqSheet.getLastRow() - 1);
    }
    _truncateSheetData_(ss, 'Equipo_Participante');
    _truncateSheetData_(ss, 'Inscripciones');
    _truncateSheetData_(ss, 'Evento_Categoria');

    // Re-cargar categorías oficiales
    _seedCategoriasCRODIT_(sheet);

    return {
      success: true,
      mensaje: 'Se cargaron las 10 categorías oficiales CRODIT. Equipos, inscripciones y vínculos relacionados fueron limpiados.',
      total: 10
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// Helper interno: limpia todos los datos de una hoja (conserva encabezado)
function _truncateSheetData_(ss, nombreHoja) {
  var sh = ss.getSheetByName(nombreHoja);
  if (!sh) return;
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}

// ---- CATEGORÍAS (CRODIT) ----
function obtenerCategorias() {
  return _cached_('cat:v1', function() { return _obtenerCategoriasImpl_(); });
}
function _obtenerCategoriasImpl_() {
  try {
    var ss = getSpreadsheet();
    var sheet = getCatSheet_(ss);
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      headers.forEach(function(h, idx) { obj[h] = data[i][idx]; });
      obj.idCategoria = obj['idCategoria'] || obj['ID'] || '';
      obj.nombre = obj['Nombre'] || '';
      obj.disciplina = obj['Disciplina'] || obj['disciplina'] || '';
      obj.descripcion = obj['Descripcion'] || obj['Descripción'] || '';
      obj.restricciones = obj['Restricciones'] || obj['restricciones'] || '';
      obj.minIntegrantes = obj['MinIntegrantes'] || obj['minIntegrantes'] || 1;
      obj.maxIntegrantes = obj['MaxIntegrantes'] || obj['maxIntegrantes'] || 4;
      obj.id = i + 1;
      rows.push(obj);
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function crearCategoria(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = getCatSheet_(ss);
    if (!sheet) {
      sheet = ss.insertSheet('Categorias');
      var h = ['idCategoria','Nombre','Disciplina','Descripcion','Restricciones','MaxIntegrantes','MinIntegrantes'];
      sheet.getRange(1, 1, 1, h.length).setValues([h]);
      sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var newId = 'CAT-' + String(data.length).padStart(3, '0');
    // Compatibilidad con esquema antiguo (sin Disciplina/Restricciones)
    if (headers.indexOf('Disciplina') === -1) {
      sheet.appendRow([newId, datos.nombre, datos.descripcion || '', datos.maxIntegrantes || 4, datos.minIntegrantes || 1]);
    } else {
      sheet.appendRow([newId, datos.nombre, datos.disciplina || '', datos.descripcion || '', datos.restricciones || '', datos.maxIntegrantes || 4, datos.minIntegrantes || 1]);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- EVENTOS (ADMIN CRUD) ----
function obtenerEventos(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  return _cached_('evt:v1', function() { return _obtenerEventosImpl_(); });
}
function _obtenerEventosImpl_() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Eventos');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];

    // Cargar Evento_Categoria para saber categorías de cada evento
    var ecSheet = ss.getSheetByName('Evento_Categoria');
    var evCats = {};
    if (ecSheet) {
      var ecData = ecSheet.getDataRange().getValues();
      for (var j = 1; j < ecData.length; j++) {
        var evId = String(ecData[j][0]);
        if (!evCats[evId]) evCats[evId] = [];
        evCats[evId].push(String(ecData[j][1]));
      }
    }

    // Cargar nombres de categorías
    var catSheet = getCatSheet_(ss);
    var catNames = {};
    if (catSheet) {
      var catData = catSheet.getDataRange().getValues();
      var catH = catData[0];
      for (var j = 1; j < catData.length; j++) {
        catNames[String(catData[j][catH.indexOf('idCategoria')])] = catData[j][catH.indexOf('Nombre')];
      }
    }

    // Contar inscripciones por evento
    var insSheet = ss.getSheetByName('Inscripciones');
    var insCounts = {};
    if (insSheet) {
      var insData = insSheet.getDataRange().getValues();
      var insH = insData[0];
      var evIdx = insH.indexOf('idEvento') !== -1 ? insH.indexOf('idEvento') : insH.indexOf('EventoId');
      for (var j = 1; j < insData.length; j++) {
        var eid = String(insData[j][evIdx]);
        insCounts[eid] = (insCounts[eid] || 0) + 1;
      }
    }

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var raw = {};
      headers.forEach(function(h, idx) { raw[h] = data[i][idx]; });
      var id = String(raw['idEvento'] || '');
      var catIds = evCats[id] || [];
      rows.push({
        idEvento: id,
        nombreEvento: raw['NombreEvento'] || '',
        descripcion: raw['Descripcion'] || '',
        fechaInicio: _dStr_(raw['FechaInicio']),
        fechaFin: _dStr_(raw['FechaFin']),
        ubicacion: raw['Ubicacion'] || '',
        estado: raw['Estado'] || 'activo',
        totalInscritos: insCounts[id] || 0,
        categorias: catIds.map(function(cid) { return { idCategoria: cid, nombre: catNames[cid] || cid }; })
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function crearEvento(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Eventos');
    if (!sheet) {
      sheet = ss.insertSheet('Eventos');
      var h = ['idEvento','NombreEvento','Descripcion','FechaInicio','FechaFin','Ubicacion','Estado'];
      sheet.getRange(1, 1, 1, h.length).setValues([h]);
      sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }
    var newId = _nextId_(sheet, 'EVT', 0);
    sheet.appendRow([newId, datos.nombreEvento, datos.descripcion || '', datos.fechaInicio, datos.fechaFin, datos.ubicacion || '', datos.estado || 'activo']);

    // Guardar categorías asociadas en Evento_Categoria
    if (datos.categorias && datos.categorias.length > 0) {
      var ecSheet = ss.getSheetByName('Evento_Categoria');
      if (!ecSheet) {
        ecSheet = ss.insertSheet('Evento_Categoria');
        ecSheet.getRange(1, 1, 1, 2).setValues([['idEvento','idCategoria']]);
        ecSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
        ecSheet.setFrozenRows(1);
      }
      datos.categorias.forEach(function(catId) {
        ecSheet.appendRow([newId, catId]);
      });
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function editarEvento(token, datos) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Eventos');
    if (!sheet) return { success: false, error: 'No existe la hoja Eventos' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][headers.indexOf('idEvento')]) === String(datos.idEvento)) {
        sheet.getRange(i + 1, headers.indexOf('NombreEvento') + 1).setValue(datos.nombreEvento);
        sheet.getRange(i + 1, headers.indexOf('Descripcion') + 1).setValue(datos.descripcion || '');
        sheet.getRange(i + 1, headers.indexOf('FechaInicio') + 1).setValue(datos.fechaInicio);
        sheet.getRange(i + 1, headers.indexOf('FechaFin') + 1).setValue(datos.fechaFin);
        sheet.getRange(i + 1, headers.indexOf('Ubicacion') + 1).setValue(datos.ubicacion || '');
        sheet.getRange(i + 1, headers.indexOf('Estado') + 1).setValue(datos.estado || 'activo');
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Evento no encontrado' };

    // Actualizar categorías: borrar las viejas y poner las nuevas
    var ecSheet = ss.getSheetByName('Evento_Categoria');
    if (ecSheet) {
      var ecData = ecSheet.getDataRange().getValues();
      for (var i = ecData.length - 1; i >= 1; i--) {
        if (String(ecData[i][0]) === String(datos.idEvento)) {
          ecSheet.deleteRow(i + 1);
        }
      }
    }
    if (datos.categorias && datos.categorias.length > 0) {
      if (!ecSheet) {
        ecSheet = ss.insertSheet('Evento_Categoria');
        ecSheet.getRange(1, 1, 1, 2).setValues([['idEvento','idCategoria']]);
        ecSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
        ecSheet.setFrozenRows(1);
      }
      datos.categorias.forEach(function(catId) {
        ecSheet.appendRow([datos.idEvento, catId]);
      });
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function obtenerEventosActivos() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Eventos');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];

    // Mapas: idEvento → [idCategoria, ...] y idCategoria → nombre
    var ecSheet = ss.getSheetByName('Evento_Categoria');
    var evCats = {};
    if (ecSheet) {
      var ecData = ecSheet.getDataRange().getValues();
      for (var j = 1; j < ecData.length; j++) {
        var evId = String(ecData[j][0]);
        if (!evCats[evId]) evCats[evId] = [];
        evCats[evId].push(String(ecData[j][1]));
      }
    }
    var catSheet = getCatSheet_(ss);
    var catNames = {};
    if (catSheet) {
      var catData = catSheet.getDataRange().getValues();
      var catH = catData[0];
      for (var j = 1; j < catData.length; j++) {
        catNames[String(catData[j][catH.indexOf('idCategoria')])] = catData[j][catH.indexOf('Nombre')] || '';
      }
    }

    var idIdx = headers.indexOf('idEvento');
    var nomIdx = headers.indexOf('NombreEvento');
    var descIdx = headers.indexOf('Descripcion');
    var fiIdx = headers.indexOf('FechaInicio');
    var ffIdx = headers.indexOf('FechaFin');
    var ubIdx = headers.indexOf('Ubicacion');
    var estIdx = headers.indexOf('Estado');

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var estado = String(data[i][estIdx] || '').toLowerCase();
      if (estado !== 'activo') continue;
      var idEv = String(data[i][idIdx]);
      var catIds = evCats[idEv] || [];
      rows.push({
        idEvento: idEv,
        nombreEvento: data[i][nomIdx] || '',
        descripcion: data[i][descIdx] || '',
        fechaInicio: _dStr_(data[i][fiIdx]),
        fechaFin: _dStr_(data[i][ffIdx]),
        ubicacion: data[i][ubIdx] || '',
        estado: estado,
        categorias: catIds.map(function(cid) { return { idCategoria: cid, nombre: catNames[cid] || cid }; })
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- INSCRIPCIONES (ADMIN + Coach) ----
function inscribirEquipo(token, idEquipo, idEvento, idCategorias) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Inscripciones');
    if (!sheet) {
      sheet = ss.insertSheet('Inscripciones');
      var h = ['idInscripcion','idEquipo','idEvento','idCategoria','CoachEmail','FechaInscripcion','Estado','EvidenciaURL'];
      sheet.getRange(1, 1, 1, h.length).setValues([h]);
      sheet.getRange(1, 1, 1, h.length).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      sheet.setFrozenRows(1);
    }
    // Migración automática: agregar columna idCategoria si falta
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('idCategoria') === -1) {
      var col = headers.length + 1;
      sheet.getRange(1, col).setValue('idCategoria');
      sheet.getRange(1, col).setFontWeight('bold').setBackground('#1e3a5f').setFontColor('#fff');
      headers.push('idCategoria');
    }

    // Normalizar lista de categorías (acepta string o array)
    var cats = [];
    if (Array.isArray(idCategorias)) cats = idCategorias;
    else if (idCategorias) cats = [idCategorias];
    cats = cats.map(function(c){ return String(c||'').trim(); }).filter(function(c){ return !!c; });
    if (!cats.length) return { success: false, error: 'Selecciona al menos una categoría.' };

    // Validar que las categorías pertenezcan al evento
    var ecSheet = ss.getSheetByName('Evento_Categoria');
    if (ecSheet && ecSheet.getLastRow() > 1) {
      var ecData = ecSheet.getDataRange().getValues();
      var ecH = ecData[0];
      var ecEv = ecH.indexOf('idEvento');
      var ecCat = ecH.indexOf('idCategoria');
      var permitidas = {};
      for (var r = 1; r < ecData.length; r++) {
        if (String(ecData[r][ecEv]) === String(idEvento)) permitidas[String(ecData[r][ecCat])] = true;
      }
      var noPermitidas = cats.filter(function(c){ return !permitidas[c]; });
      if (noPermitidas.length) return { success: false, error: 'Categoría(s) no permitidas en este evento: ' + noPermitidas.join(', ') };
    }

    var data = sheet.getDataRange().getValues();
    headers = data[0];
    var eqIdx = headers.indexOf('idEquipo');
    var evIdx = headers.indexOf('idEvento');
    var catIdx = headers.indexOf('idCategoria');
    var yaInscritas = {};
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][eqIdx]) === String(idEquipo) && String(data[i][evIdx]) === String(idEvento)) {
        yaInscritas[String(data[i][catIdx] || '')] = true;
      }
    }
    var aCrear = cats.filter(function(c){ return !yaInscritas[c]; });
    if (!aCrear.length) return { success: false, error: 'Este equipo ya está inscrito en las categorías seleccionadas.' };

    var creadas = [];
    var ahora = new Date();
    aCrear.forEach(function(c, k) {
      var newId = 'INS-' + String(data.length + k).padStart(3, '0');
      var row = new Array(headers.length);
      for (var j = 0; j < headers.length; j++) row[j] = '';
      row[headers.indexOf('idInscripcion')] = newId;
      row[eqIdx] = idEquipo;
      row[evIdx] = idEvento;
      if (catIdx !== -1) row[catIdx] = c;
      row[headers.indexOf('CoachEmail')] = session.email;
      row[headers.indexOf('FechaInscripcion')] = ahora;
      row[headers.indexOf('Estado')] = 'pendiente';
      var evidCol = headers.indexOf('EvidenciaURL');
      if (evidCol !== -1) row[evidCol] = '';
      sheet.appendRow(row);
      creadas.push(newId);
    });
    return { success: true, idInscripcion: creadas[0], idInscripciones: creadas, creadas: creadas.length };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// Inscripciones del coach actual (para CoachDashboard)
function obtenerInscripcionesCoach(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  var em = String(session.email || '').toLowerCase();
  return _cached_('insC:v1:' + em, function() { return _obtenerInscripcionesCoachImpl_(token, session); });
}
function _obtenerInscripcionesCoachImpl_(token, session) {
  try {
    var all = obtenerTodasInscripciones(token);
    if (!all.success) return all;
    var ss = getSpreadsheet();
    var insSheet = ss.getSheetByName('Inscripciones');
    if (!insSheet) return { success: true, data: [] };
    var insData = insSheet.getDataRange().getValues();
    if (insData.length < 2) return { success: true, data: [] };
    var insH = insData[0];
    var coachIdx = insH.indexOf('CoachEmail');
    var idIdx = insH.indexOf('idInscripcion');
    var permitidos = {};
    for (var i = 1; i < insData.length; i++) {
      if (String(insData[i][coachIdx]).toLowerCase() === session.email.toLowerCase()) {
        permitidos[String(insData[i][idIdx])] = true;
      }
    }
    var filtrados = all.data.filter(function(row) { return permitidos[String(row.idInscripcion)]; });
    return { success: true, data: filtrados };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// Subir URL de evidencia para una inscripción del coach
function subirEvidencia(token, idInscripcion, url) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCacheCoach_(session.email);
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Inscripciones');
    if (!sheet) return { success: false, error: 'No existe la hoja Inscripciones' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idInscripcion');
    var coachIdx = headers.indexOf('CoachEmail');
    var evidIdx = headers.indexOf('EvidenciaURL');
    if (evidIdx === -1) return { success: false, error: 'Columna EvidenciaURL no existe' };
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idIdx]) === String(idInscripcion)) {
        if (coachIdx !== -1 && String(data[i][coachIdx]).toLowerCase() !== session.email.toLowerCase()) {
          return { success: false, error: 'No tienes permisos sobre esta inscripción' };
        }
        sheet.getRange(i + 1, evidIdx + 1).setValue(url || '');
        return { success: true };
      }
    }
    return { success: false, error: 'Inscripción no encontrada' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function obtenerTodasInscripciones(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  return _cached_('ins:v1', function() { return _obtenerTodasInscripcionesImpl_(); });
}
function _obtenerTodasInscripcionesImpl_() {
  try {
    var ss = getSpreadsheet();
    var insSheet = ss.getSheetByName('Inscripciones');
    if (!insSheet) return { success: true, data: [] };
    var insData = insSheet.getDataRange().getValues();
    if (insData.length < 2) return { success: true, data: [] };
    var insH = insData[0];

    // Cargar equipos
    var eqSheet = ss.getSheetByName('Equipos');
    var equipos = {};
    if (eqSheet) {
      var eqData = eqSheet.getDataRange().getValues();
      var eqH = eqData[0];
      for (var i = 1; i < eqData.length; i++) {
        var eid = String(eqData[i][eqH.indexOf('idEquipo')]);
        equipos[eid] = {
          nombre: eqData[i][eqH.indexOf('NombreEquipo')] || eqData[i][eqH.indexOf('Nombre')] || '',
          catId: String(eqData[i][eqH.indexOf('idCategoria')] || eqData[i][eqH.indexOf('Categoría')] || '')
        };
      }
    }

    // Cargar eventos
    var evSheet = ss.getSheetByName('Eventos');
    var eventos = {};
    if (evSheet) {
      var evData = evSheet.getDataRange().getValues();
      var evH = evData[0];
      for (var i = 1; i < evData.length; i++) {
        eventos[String(evData[i][evH.indexOf('idEvento')])] = evData[i][evH.indexOf('NombreEvento')] || '';
      }
    }

    // Cargar categorías
    var catSheet = getCatSheet_(ss);
    var catNames = {};
    if (catSheet) {
      var catData = catSheet.getDataRange().getValues();
      var catH = catData[0];
      for (var i = 1; i < catData.length; i++) {
        catNames[String(catData[i][catH.indexOf('idCategoria')])] = catData[i][catH.indexOf('Nombre')];
      }
    }

    // Cargar coaches
    var usuarios = {};
    var usSheet = ss.getSheetByName('Usuarios');
    if (usSheet) {
      var usData = usSheet.getDataRange().getValues();
      var usH = usData[0];
      for (var i = 1; i < usData.length; i++) {
        usuarios[String(usData[i][0]).toLowerCase()] = usData[i][usH.indexOf('Nombre')] || '';
      }
    }

    var rows = [];
    for (var i = 1; i < insData.length; i++) {
      var obj = {};
      insH.forEach(function(h, idx) { obj[h] = insData[i][idx]; });
      var eqId = String(obj['idEquipo'] || obj['EquipoId'] || '');
      var evId = String(obj['idEvento'] || obj['EventoId'] || '');
      var coachEmail = String(obj['CoachEmail'] || '').toLowerCase();
      var eq = equipos[eqId] || {};
      // Categoría: preferir la de la inscripción (multi-categoría); fallback a la del equipo (legacy)
      var insCatId = String(obj['idCategoria'] || '');
      var catIdEff = insCatId || eq.catId || '';

      rows.push({
        idInscripcion: obj['idInscripcion'] || String(i + 1),
        idEvento: evId,
        idEquipo: eqId,
        idCategoria: catIdEff,
        nombreEquipo: eq.nombre || eqId,
        nombreCoach: usuarios[coachEmail] || coachEmail,
        nombreEvento: eventos[evId] || evId,
        categoria: catNames[catIdEff] || catIdEff || '-',
        fechaInscripcion: _dStr_(obj['FechaInscripcion'] || obj['Fecha']),
        estado: (obj['Estado'] || 'pendiente').toLowerCase(),
        evidenciaURL: obj['EvidenciaURL'] || ''
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

function aprobarInscripcion(token, idInscripcion) {
  return cambiarEstadoInscripcion_(token, idInscripcion, 'aprobada', '');
}

function rechazarInscripcion(token, idInscripcion, motivo) {
  return cambiarEstadoInscripcion_(token, idInscripcion, 'rechazada', motivo || '');
}

function cambiarEstadoInscripcion_(token, idInscripcion, nuevoEstado, motivo) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Inscripciones');
    if (!sheet) return { success: false, error: 'No existe la hoja Inscripciones' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = headers.indexOf('idInscripcion');
    var estCol = headers.indexOf('Estado');
    var coachCol = headers.indexOf('CoachEmail');
    var eqCol = headers.indexOf('idEquipo');
    var evCol = headers.indexOf('idEvento');
    if (idCol === -1 || estCol === -1) return { success: false, error: 'Estructura de hoja inválida' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(idInscripcion)) {
        sheet.getRange(i + 1, estCol + 1).setValue(nuevoEstado);

        // Resolver datos para el correo al coach
        var coachEmail = coachCol !== -1 ? String(data[i][coachCol] || '') : '';
        var idEq = eqCol !== -1 ? String(data[i][eqCol] || '') : '';
        var idEv = evCol !== -1 ? String(data[i][evCol] || '') : '';
        var nombreEquipo = idEq;
        var eqSheet = ss.getSheetByName('Equipos');
        if (eqSheet && idEq) {
          var eqData = eqSheet.getDataRange().getValues();
          var eqH = eqData[0];
          var eqIdIdx = eqH.indexOf('idEquipo');
          var eqNomIdx = eqH.indexOf('NombreEquipo');
          for (var k = 1; k < eqData.length; k++) {
            if (String(eqData[k][eqIdIdx]) === idEq) {
              nombreEquipo = eqData[k][eqNomIdx] || idEq;
              break;
            }
          }
        }
        var nombreEvento = idEv;
        var evSheet = ss.getSheetByName('Eventos');
        if (evSheet && idEv) {
          var evData = evSheet.getDataRange().getValues();
          var evH = evData[0];
          var evIdIdx = evH.indexOf('idEvento');
          var evNomIdx = evH.indexOf('NombreEvento');
          for (var m = 1; m < evData.length; m++) {
            if (String(evData[m][evIdIdx]) === idEv) {
              nombreEvento = evData[m][evNomIdx] || idEv;
              break;
            }
          }
        }
        _enviarCorreoInscripcion_(coachEmail, nombreEquipo, nombreEvento, nuevoEstado, motivo);
        return { success: true };
      }
    }
    return { success: false, error: 'Inscripción no encontrada' };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- COACHES (Admin lectura) ----
function obtenerCoaches(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  return _cached_('coa:v1', function() { return _obtenerCoachesImpl_(); });
}
function _obtenerCoachesImpl_() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Coaches');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      rows.push({
        idCoach: data[i][headers.indexOf('idCoach')] || '',
        nombre: data[i][headers.indexOf('Nombre')] || '',
        apellido: data[i][headers.indexOf('Apellido')] || '',
        institucion: data[i][headers.indexOf('Institucion')] || '',
        correo: data[i][headers.indexOf('Correo')] || '',
        telefono: data[i][headers.indexOf('Telefono')] || '',
        fechaRegistro: _dStr_(data[i][headers.indexOf('FechaRegistro')])
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- ADMIN: Listar todos los participantes con info completa ----
function obtenerParticipantesAdmin(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  return _cached_('pad:v2', function() { return _obtenerParticipantesAdminImpl_(); });
}
function _obtenerParticipantesAdminImpl_() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) return { success: true, data: [] };
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    var headers = data[0];
    var idIdx = headers.indexOf('idParticipante');
    var nomIdx = headers.indexOf('Nombre');
    var apeIdx = headers.indexOf('Apellido');
    var edadIdx = headers.indexOf('Edad');
    var corIdx = headers.indexOf('Correo');
    var escIdx = headers.indexOf('Escolaridad');
    var ceIdx = headers.indexOf('CoachEmail');
    var frIdx = headers.indexOf('FechaRegistro');

    // ---- Construir mapa idParticipante → {coachEmail, coachNombre} cruzando Equipo_Participante + Equipos + Usuarios ----
    var coachByPart = {}; // { idParticipante: { email, nombre } }
    var equipos = ss.getSheetByName('Equipos');
    var equipoPart = ss.getSheetByName('Equipo_Participante');
    if (equipos && equipoPart) {
      // Map idEquipo → CoachEmail
      var eqData = equipos.getDataRange().getValues();
      var eqH = eqData[0] || [];
      var eqIdIdx = eqH.indexOf('idEquipo');
      var eqCoachIdx = eqH.indexOf('CoachEmail');
      var equipoCoach = {};
      if (eqIdIdx !== -1 && eqCoachIdx !== -1) {
        for (var k = 1; k < eqData.length; k++) {
          equipoCoach[String(eqData[k][eqIdIdx])] = String(eqData[k][eqCoachIdx] || '').trim();
        }
      }
      // Map idParticipante → CoachEmail (vía Equipo_Participante)
      var epData = equipoPart.getDataRange().getValues();
      var epH = epData[0] || [];
      var epEqIdx = epH.indexOf('idEquipo');
      var epPaIdx = epH.indexOf('idParticipante');
      if (epEqIdx !== -1 && epPaIdx !== -1) {
        for (var m = 1; m < epData.length; m++) {
          var idPart = String(epData[m][epPaIdx]);
          var idEq = String(epData[m][epEqIdx]);
          var cEmail = equipoCoach[idEq];
          if (idPart && cEmail && !coachByPart[idPart]) {
            coachByPart[idPart] = { email: cEmail, nombre: '' };
          }
        }
      }
      // Resolver nombres del coach desde Usuarios
      var usuarios = ss.getSheetByName('Usuarios');
      if (usuarios) {
        var uData = usuarios.getDataRange().getValues();
        var uH = uData[0] || [];
        var uCorIdx = uH.indexOf('Correo');
        var uNomIdx = uH.indexOf('Nombre');
        if (uCorIdx !== -1 && uNomIdx !== -1) {
          var nombreByEmail = {};
          for (var n = 1; n < uData.length; n++) {
            nombreByEmail[String(uData[n][uCorIdx] || '').trim().toLowerCase()] = String(uData[n][uNomIdx] || '').trim();
          }
          for (var pid in coachByPart) {
            var lower = String(coachByPart[pid].email || '').toLowerCase();
            if (nombreByEmail[lower]) coachByPart[pid].nombre = nombreByEmail[lower];
          }
        }
      }
    }

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var pid2 = String(data[i][idIdx] || '');
      // Prioridad: 1) cruce con Equipo_Participante  2) campo CoachEmail directo en Participantes
      var resolved = coachByPart[pid2];
      var emailCoach = resolved ? resolved.email : (data[i][ceIdx] || '');
      var nombreCoach = resolved ? resolved.nombre : '';
      rows.push({
        idParticipante: pid2,
        nombre: data[i][nomIdx] || '',
        apellido: data[i][apeIdx] || '',
        edad: data[i][edadIdx] || '',
        correo: data[i][corIdx] || '',
        escolaridad: data[i][escIdx] || '',
        coachEmail: emailCoach,
        coachNombre: nombreCoach,
        fechaRegistro: _dStr_(data[i][frIdx])
      });
    }
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// Helper: borrar todas las filas de una hoja donde la columna `colName` coincida con cualquiera de los valores
function _deleteRowsWhere_(ss, sheetName, colName, values) {
  if (!values || !values.length) return;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx = headers.indexOf(colName);
  if (idx === -1) return;
  var setLower = {};
  values.forEach(function(v){ setLower[String(v).toLowerCase()] = true; });
  for (var i = data.length - 1; i >= 1; i--) {
    if (setLower[String(data[i][idx]).toLowerCase()]) sheet.deleteRow(i + 1);
  }
}

// ---- ADMIN: Borrar participante con cascada total ----
function eliminarParticipanteAdmin(token, idParticipante) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Participantes');
    if (!sheet) return { success: false, error: 'No existe la hoja Participantes' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idParticipante');
    var correoIdx = headers.indexOf('Correo');
    var correo = '';
    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIdx]) === String(idParticipante)) {
        correo = String(data[i][correoIdx] || '');
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Participante no encontrado' };

    _deleteRowsWhere_(ss, 'Equipo_Participante', 'idParticipante', [idParticipante]);
    if (correo) _deleteRowsWhere_(ss, 'Usuarios', 'Correo', [correo.toLowerCase()]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- ADMIN: Borrar coach con cascada total ----
function eliminarCoachAdmin(token, idCoach) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var coachSheet = ss.getSheetByName('Coaches');
    if (!coachSheet) return { success: false, error: 'No existe la hoja Coaches' };
    var data = coachSheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idCoach');
    var correoIdx = headers.indexOf('Correo');
    var correo = '';
    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIdx]) === String(idCoach)) {
        correo = String(data[i][correoIdx] || '').toLowerCase();
        coachSheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Coach no encontrado' };
    if (!correo) return { success: true };

    // Equipos del coach
    var equiposIds = [];
    var eqSheet = ss.getSheetByName('Equipos');
    if (eqSheet && eqSheet.getLastRow() > 1) {
      var eqData = eqSheet.getDataRange().getValues();
      var eqH = eqData[0];
      var ceIdx = eqH.indexOf('CoachEmail');
      var eqIdIdx = eqH.indexOf('idEquipo');
      for (var j = eqData.length - 1; j >= 1; j--) {
        if (String(eqData[j][ceIdx]).toLowerCase() === correo) {
          equiposIds.push(String(eqData[j][eqIdIdx]));
          eqSheet.deleteRow(j + 1);
        }
      }
    }

    // Participantes del coach
    var partIds = [];
    var partCorreos = [];
    var partSheet = ss.getSheetByName('Participantes');
    if (partSheet && partSheet.getLastRow() > 1) {
      var pData = partSheet.getDataRange().getValues();
      var pH = pData[0];
      var pCeIdx = pH.indexOf('CoachEmail');
      var pIdIdx = pH.indexOf('idParticipante');
      var pCorIdx = pH.indexOf('Correo');
      for (var k = pData.length - 1; k >= 1; k--) {
        if (String(pData[k][pCeIdx]).toLowerCase() === correo) {
          partIds.push(String(pData[k][pIdIdx]));
          partCorreos.push(String(pData[k][pCorIdx] || '').toLowerCase());
          partSheet.deleteRow(k + 1);
        }
      }
    }

    // Cascadas
    _deleteRowsWhere_(ss, 'Equipo_Participante', 'idEquipo', equiposIds);
    _deleteRowsWhere_(ss, 'Equipo_Participante', 'idParticipante', partIds);
    _deleteRowsWhere_(ss, 'Inscripciones', 'idEquipo', equiposIds);
    _deleteRowsWhere_(ss, 'Inscripciones', 'CoachEmail', [correo]);
    _deleteRowsWhere_(ss, 'Usuarios', 'Correo', [correo].concat(partCorreos.filter(function(c){ return c; })));

    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- ADMIN: Equipos inscritos en un evento (con coach y participantes) ----
function obtenerEquiposPorEvento(token, idEvento) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  try {
    var ss = getSpreadsheet();
    var insSheet = ss.getSheetByName('Inscripciones');
    if (!insSheet || insSheet.getLastRow() < 2) return { success: true, data: [] };
    var insData = insSheet.getDataRange().getValues();
    var iH = insData[0];
    var iEvIdx = iH.indexOf('idEvento'); if (iEvIdx === -1) iEvIdx = iH.indexOf('EventoId');
    var iEqIdx = iH.indexOf('idEquipo'); if (iEqIdx === -1) iEqIdx = iH.indexOf('EquipoId');
    var iEstIdx = iH.indexOf('Estado');
    var iCatIdx = iH.indexOf('idCategoria');
    // Una fila por inscripción (equipo+categoría) — un equipo puede aparecer varias veces si participa en >1 categoría
    var inscripciones = [];
    for (var i = 1; i < insData.length; i++) {
      if (String(insData[i][iEvIdx]) === String(idEvento)) {
        inscripciones.push({
          idEquipo: String(insData[i][iEqIdx]),
          idCategoria: iCatIdx !== -1 ? String(insData[i][iCatIdx] || '') : '',
          estado: (insData[i][iEstIdx] || 'pendiente')
        });
      }
    }
    if (!inscripciones.length) return { success: true, data: [] };

    // Equipos
    var equiposMap = {};
    var eqSheet = ss.getSheetByName('Equipos');
    if (eqSheet && eqSheet.getLastRow() > 1) {
      var eqData = eqSheet.getDataRange().getValues();
      var eH = eqData[0];
      var eqIdIdx = eH.indexOf('idEquipo');
      var eqNomIdx = eH.indexOf('NombreEquipo'); if (eqNomIdx === -1) eqNomIdx = eH.indexOf('Nombre');
      var eqCatIdx = eH.indexOf('idCategoria');
      var eqCoIdx = eH.indexOf('CoachEmail');
      for (var j = 1; j < eqData.length; j++) {
        equiposMap[String(eqData[j][eqIdIdx])] = {
          nombre: eqData[j][eqNomIdx] || '',
          idCategoria: String(eqData[j][eqCatIdx] || ''),
          coachEmail: String(eqData[j][eqCoIdx] || '')
        };
      }
    }

    // Categorías
    var catNames = {};
    var catSheet = getCatSheet_(ss);
    if (catSheet && catSheet.getLastRow() > 1) {
      var catData = catSheet.getDataRange().getValues();
      var cH = catData[0];
      for (var k = 1; k < catData.length; k++) {
        catNames[String(catData[k][cH.indexOf('idCategoria')])] = catData[k][cH.indexOf('Nombre')];
      }
    }

    // Coaches por correo
    var coachByEmail = {};
    var coachSheet = ss.getSheetByName('Coaches');
    if (coachSheet && coachSheet.getLastRow() > 1) {
      var coData = coachSheet.getDataRange().getValues();
      var coH = coData[0];
      for (var m = 1; m < coData.length; m++) {
        var em = String(coData[m][coH.indexOf('Correo')] || '').toLowerCase();
        coachByEmail[em] = ((coData[m][coH.indexOf('Nombre')] || '') + ' ' + (coData[m][coH.indexOf('Apellido')] || '')).trim();
      }
    }

    // Participantes por id
    var partById = {};
    var partSheet = ss.getSheetByName('Participantes');
    if (partSheet && partSheet.getLastRow() > 1) {
      var pData = partSheet.getDataRange().getValues();
      var pH = pData[0];
      for (var n = 1; n < pData.length; n++) {
        partById[String(pData[n][pH.indexOf('idParticipante')])] = {
          nombre: ((pData[n][pH.indexOf('Nombre')] || '') + ' ' + (pData[n][pH.indexOf('Apellido')] || '')).trim(),
          correo: pData[n][pH.indexOf('Correo')] || ''
        };
      }
    }

    // Equipo_Participante
    var miembrosByEquipo = {};
    var epSheet = ss.getSheetByName('Equipo_Participante');
    if (epSheet && epSheet.getLastRow() > 1) {
      var epData = epSheet.getDataRange().getValues();
      var epH = epData[0];
      var epEqIdx = epH.indexOf('idEquipo');
      var epPartIdx = epH.indexOf('idParticipante');
      for (var p = 1; p < epData.length; p++) {
        var key = String(epData[p][epEqIdx]);
        if (!miembrosByEquipo[key]) miembrosByEquipo[key] = [];
        var pid = String(epData[p][epPartIdx]);
        if (partById[pid]) miembrosByEquipo[key].push(partById[pid]);
      }
    }

    var rows = inscripciones.map(function(ins){
      var eid = ins.idEquipo;
      var eq = equiposMap[eid] || {};
      var ce = String(eq.coachEmail || '').toLowerCase();
      var catId = ins.idCategoria || eq.idCategoria || '';
      return {
        idEquipo: eid,
        nombreEquipo: eq.nombre || eid,
        idCategoria: catId,
        categoria: catNames[catId] || '-',
        coachNombre: coachByEmail[ce] || ce || '-',
        coachEmail: eq.coachEmail || '',
        estado: ins.estado || 'pendiente',
        participantes: miembrosByEquipo[eid] || []
      };
    });
    return { success: true, data: rows };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- ADMIN: Borrar evento con cascada ----
function eliminarEventoAdmin(token, idEvento) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Eventos');
    if (!sheet) return { success: false, error: 'No existe la hoja Eventos' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idEvento');
    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIdx]) === String(idEvento)) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Evento no encontrado' };
    _deleteRowsWhere_(ss, 'Evento_Categoria', 'idEvento', [idEvento]);
    _deleteRowsWhere_(ss, 'Inscripciones', 'idEvento', [idEvento]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- ADMIN: Borrar categoría con cascada ----
function eliminarCategoriaAdmin(token, idCategoria) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  _bustCache_();
  try {
    var ss = getSpreadsheet();
    var sheet = getCatSheet_(ss);
    if (!sheet) return { success: false, error: 'No existe la hoja Categorías' };
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idIdx = headers.indexOf('idCategoria');
    var found = false;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idIdx]) === String(idCategoria)) {
        sheet.deleteRow(i + 1);
        found = true;
        break;
      }
    }
    if (!found) return { success: false, error: 'Categoría no encontrada' };

    // Equipos con esta categoría -> borrar equipos + relaciones + inscripciones
    var equiposIds = [];
    var eqSheet = ss.getSheetByName('Equipos');
    if (eqSheet && eqSheet.getLastRow() > 1) {
      var eqData = eqSheet.getDataRange().getValues();
      var eqH = eqData[0];
      var catIdx = eqH.indexOf('idCategoria');
      var eqIdIdx = eqH.indexOf('idEquipo');
      for (var j = eqData.length - 1; j >= 1; j--) {
        if (String(eqData[j][catIdx]) === String(idCategoria)) {
          equiposIds.push(String(eqData[j][eqIdIdx]));
          eqSheet.deleteRow(j + 1);
        }
      }
    }
    _deleteRowsWhere_(ss, 'Equipo_Participante', 'idEquipo', equiposIds);
    _deleteRowsWhere_(ss, 'Inscripciones', 'idEquipo', equiposIds);
    _deleteRowsWhere_(ss, 'Evento_Categoria', 'idCategoria', [idCategoria]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- REPORTES (Admin dashboard) ----
function obtenerReportes(token) {
  var session = getSessionUser_(token);
  if (!session) return { success: false, error: 'Sesión expirada' };
  return _cached_('rpt:v1', function() { return _obtenerReportesImpl_(); });
}
function _obtenerReportesImpl_() {
  try {
    var ss = getSpreadsheet();
    var count = function(name) {
      var s = ss.getSheetByName(name);
      if (!s) return 0;
      return Math.max(0, s.getLastRow() - 1);
    };

    // Contar eventos activos
    var totalEventos = count('Eventos');
    var eventosActivos = 0;
    var evSheet = ss.getSheetByName('Eventos');
    if (evSheet && evSheet.getLastRow() > 1) {
      var evData = evSheet.getDataRange().getValues();
      var estIdx = evData[0].indexOf('Estado');
      for (var i = 1; i < evData.length; i++) {
        if (String(evData[i][estIdx]).toLowerCase() === 'activo') eventosActivos++;
      }
    }

    // Contar inscripciones por estado
    var inscripcionesPorEstado = {};
    var totalIns = 0;
    var insSheet = ss.getSheetByName('Inscripciones');
    if (insSheet && insSheet.getLastRow() > 1) {
      var insData = insSheet.getDataRange().getValues();
      var estIdx = insData[0].indexOf('Estado');
      for (var i = 1; i < insData.length; i++) {
        var est = String(insData[i][estIdx] || 'pendiente').toLowerCase();
        inscripcionesPorEstado[est] = (inscripcionesPorEstado[est] || 0) + 1;
        totalIns++;
      }
    }

    return {
      success: true,
      data: {
        totalCoaches: count('Coaches'),
        totalParticipantes: count('Participantes'),
        totalEquipos: count('Equipos'),
        totalEventos: totalEventos,
        eventosActivos: eventosActivos,
        totalInscripciones: totalIns,
        inscripcionesPorEstado: inscripcionesPorEstado
      }
    };
  } catch(e) { return { success: false, error: e.toString() }; }
}

// ---- CONSULTAR PARTICIPANTE (para ParticipantView) ----
function consultarParticipante(token) {
  try {
    var session = getSessionUser_(token);
    if (!session) return { success: false, error: 'Sesión expirada' };
    if (!session.email) return { success: false, error: 'Sesión sin correo' };
    var ss = getSpreadsheet();
    if (!ss) return { success: false, error: 'No se pudo abrir la hoja de cálculo' };
    var email = String(session.email).toLowerCase();

    // Buscar participante en hoja Participantes
    var partSheet = ss.getSheetByName('Participantes');
    if (!partSheet) return { success: false, error: 'No se encontró la hoja Participantes' };
    var partData = partSheet.getDataRange().getValues();
    if (partData.length < 2) return { success: false, error: 'No se encontró el participante' };
    var partHeaders = partData[0];
    var participante = null;
    for (var i = 1; i < partData.length; i++) {
      var correoIdx = partHeaders.indexOf('Correo') !== -1 ? partHeaders.indexOf('Correo') : partHeaders.indexOf('Email');
      if (correoIdx === -1) continue;
      if (String(partData[i][correoIdx]).trim().toLowerCase() === email) {
        participante = {};
        partHeaders.forEach(function(h, idx) { participante[h] = partData[i][idx]; });
        break;
      }
    }
    if (!participante) {
      // Intentar buscar por CoachEmail (compatibilidad con estructura anterior)
      for (var i = 1; i < partData.length; i++) {
        var ceIdx = partHeaders.indexOf('CoachEmail');
        var emIdx = partHeaders.indexOf('Email') !== -1 ? partHeaders.indexOf('Email') : partHeaders.indexOf('Correo');
        if (emIdx !== -1 && String(partData[i][emIdx]).trim().toLowerCase() === email) {
          participante = {};
          partHeaders.forEach(function(h, idx) { participante[h] = partData[i][idx]; });
          break;
        }
      }
    }
    if (!participante) {
      return { success: true, data: { nombre: session.name || '', apellido: '', correo: email, edad: '', escolaridad: '', idParticipante: '', equipos: [] } };
    }

    // Mapear campos
    var result = {
      idParticipante: participante['idParticipante'] || participante['ID'] || '',
      nombre: participante['Nombre'] || '',
      apellido: participante['Apellido'] || '',
      correo: participante['Correo'] || participante['Email'] || email,
      edad: participante['Edad'] || '',
      escolaridad: participante['Escolaridad'] || participante['Institución'] || '',
      equipos: []
    };

    // Buscar equipos del participante via Equipo_Participante
    var epSheet = ss.getSheetByName('Equipo_Participante');
    var eqSheet = ss.getSheetByName('Equipos');
    var catSheet = ss.getSheetByName('Categorías') || ss.getSheetByName('Categorias');
    var insSheet = ss.getSheetByName('Inscripciones');
    var evSheet = ss.getSheetByName('Eventos');

    var misEquipoIds = [];
    if (epSheet && result.idParticipante) {
      var epData = epSheet.getDataRange().getValues();
      var epHeaders = epData[0];
      for (var i = 1; i < epData.length; i++) {
        if (String(epData[i][epHeaders.indexOf('idParticipante')]) === String(result.idParticipante)) {
          misEquipoIds.push(String(epData[i][epHeaders.indexOf('idEquipo')]));
        }
      }
    }

    // Cargar categorías
    var categorias = {};
    if (catSheet) {
      var catData = catSheet.getDataRange().getValues();
      var catHeaders = catData[0];
      for (var i = 1; i < catData.length; i++) {
        categorias[String(catData[i][catHeaders.indexOf('idCategoria')])] = catData[i][catHeaders.indexOf('Nombre')];
      }
    }

    // Cargar eventos
    var eventos = {};
    if (evSheet) {
      var evData = evSheet.getDataRange().getValues();
      var evHeaders = evData[0];
      for (var i = 1; i < evData.length; i++) {
        var evObj = {};
        evHeaders.forEach(function(h, idx) { evObj[h] = evData[i][idx]; });
        eventos[String(evObj['idEvento'])] = evObj;
      }
    }

    // Cargar inscripciones
    var inscripcionesPorEquipo = {};
    if (insSheet) {
      var insData = insSheet.getDataRange().getValues();
      var insHeaders = insData[0];
      for (var i = 1; i < insData.length; i++) {
        var insObj = {};
        insHeaders.forEach(function(h, idx) { insObj[h] = insData[i][idx]; });
        var eqId = String(insObj['idEquipo'] || insObj['EquipoId']);
        if (!inscripcionesPorEquipo[eqId]) inscripcionesPorEquipo[eqId] = [];
        var ev = eventos[String(insObj['idEvento'] || insObj['EventoId'])] || {};
        var insCatId = String(insObj['idCategoria'] || '');
        inscripcionesPorEquipo[eqId].push({
          idInscripcion: insObj['idInscripcion'] || '',
          estado: (insObj['Estado'] || 'pendiente').toLowerCase(),
          fechaInscripcion: insObj['FechaInscripcion'] || insObj['Fecha'] || '',
          evidenciaURL: insObj['EvidenciaURL'] || '',
          idCategoria: insCatId,
          categoria: categorias[insCatId] || '',
          nombreEvento: ev['NombreEvento'] || ev['nombreEvento'] || '',
          fechaInicio: ev['FechaInicio'] || ev['fechaInicio'] || '',
          ubicacion: ev['Ubicacion'] || ev['ubicacion'] || ''
        });
      }
    }

    // Construir equipos con miembros
    if (eqSheet && misEquipoIds.length > 0) {
      var eqData = eqSheet.getDataRange().getValues();
      var eqHeaders = eqData[0];
      for (var i = 1; i < eqData.length; i++) {
        var eqId = String(eqData[i][eqHeaders.indexOf('idEquipo')]);
        if (misEquipoIds.indexOf(eqId) === -1) continue;
        var catId = String(eqData[i][eqHeaders.indexOf('idCategoria')] || eqData[i][eqHeaders.indexOf('Categoría')] || '');

        // Obtener miembros del equipo
        var miembros = [];
        if (epSheet) {
          var epData2 = epSheet.getDataRange().getValues();
          var epH2 = epData2[0];
          for (var j = 1; j < epData2.length; j++) {
            if (String(epData2[j][epH2.indexOf('idEquipo')]) === eqId) {
              var mId = String(epData2[j][epH2.indexOf('idParticipante')]);
              // Buscar nombre del participante
              for (var k = 1; k < partData.length; k++) {
                if (String(partData[k][partHeaders.indexOf('idParticipante')]) === mId) {
                  miembros.push({
                    idParticipante: mId,
                    nombre: partData[k][partHeaders.indexOf('Nombre')] || '',
                    apellido: partData[k][partHeaders.indexOf('Apellido')] || ''
                  });
                  break;
                }
              }
            }
          }
        }

        result.equipos.push({
          idEquipo: eqId,
          nombreEquipo: eqData[i][eqHeaders.indexOf('NombreEquipo')] || eqData[i][eqHeaders.indexOf('Nombre')] || '',
          categoriaNombre: categorias[catId] || catId || '',
          miembros: miembros,
          inscripciones: inscripcionesPorEquipo[eqId] || []
        });
      }
    }

    // Sanitizar todo antes de devolver (Date → ISO string, undefined/null → '', resto → tipo nativo)
    return { success: true, data: sanitizarParaCliente_(result) };
  } catch(e) { return { success: false, error: String(e && e.message || e) }; }
}

// Convierte recursivamente cualquier valor a algo seguro para google.script.run
function sanitizarParaCliente_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    var t = v.getTime();
    return isNaN(t) ? '' : v.toISOString();
  }
  var tipo = typeof v;
  if (tipo === 'string' || tipo === 'boolean') return v;
  if (tipo === 'number') return isFinite(v) ? v : '';
  if (Array.isArray(v)) return v.map(sanitizarParaCliente_);
  if (tipo === 'object') {
    var out = {};
    for (var k in v) {
      if (Object.prototype.hasOwnProperty.call(v, k)) out[k] = sanitizarParaCliente_(v[k]);
    }
    return out;
  }
  return String(v);
}

// ==========================================
// DATOS DE TABLA
// ==========================================
function getTableData() {
  return {
    headers: ['Equipo', 'Institución', 'Categoría', 'Disciplina', 'Puntuación', 'Estado'],
    rows: [
      ['CimarronBots',     'UABC FIM',         'Seguidor de Línea Profesional (Turbina)', 'Seguidor de Línea', '9850', '✅ Clasificado'],
      ['RoboCimarron',     'UABC FIAD',        'Minisumo Profesional',                    'Sumo',              '9720', '✅ Clasificado'],
      ['MexicaliBots',     'UABC FIM',         'Seguidor de Línea Amateur',               'Seguidor de Línea', '9600', '✅ Clasificado'],
      ['CircuitBreakers',  'UABC Tijuana',     'Laberinto Amateur',                       'Laberinto',         '9450', '⏳ Pendiente'],
      ['IronMinds',        'CETYS Mexicali',   'Sumo RC 20kg',                            'Sumo',              '9300', '✅ Clasificado'],
      ['NanoBot Squad',    'ITM',              'MinisumoRC',                              'Sumo',              '9150', '✅ Clasificado'],
      ['Volt Surge',       'UABC FIM',         'Seguidor de Línea Profesional',           'Seguidor de Línea', '9000', '⏳ Pendiente'],
      ['CyberPirates',     'UABC Ensenada',    'Laberinto Profesional',                   'Laberinto',         '8900', '❌ Eliminado'],
      ['ArcLab',           'CETYS Tijuana',    'Robofut',                                 'Fútbol Robótico',   '8750', '✅ Clasificado'],
      ['PhoenixBot',       'UABC FIM',         'Minisumo Amateur',                        'Sumo',              '8600', '⏳ Pendiente'],
    ]
  };
}
