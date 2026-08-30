// ═══════════════════════════════════════════════════════════════════════════
// File Protection · Remitos — Google Apps Script v6.0
// Autenticación Google Sheets · Sucursales dinámicas · Drive integration
//
// OPTIMIZACIONES V8 / Apps Script:
//   • Una sola apertura de Spreadsheet por request (openById cacheado)
//   • getDataRange().getValues() llamado UNA vez por función
//   • Limpieza de sesiones con setValues() en bloque (no deleteRow en loop)
//   • Lecturas de fila única con getRange(row,1,1,N).getValues()[0]
//   • Array methods V8: find(), findIndex(), filter(), map(), some()
//   • Escrituras en bloque: getRange(r,c,1,N).setValues([[...]])
//   • switch() en doGet en lugar de if-chain
//   • 'use strict' para optimizaciones adicionales del V8 runtime
//   • DriveApp: addFile + removeFile en lugar de moveTo (más compatible)
//   • Verificación de token con find() en O(n) sin bucles manuales
//
// INSTRUCCIONES:
//   1. Completá USERS_SS_ID con el ID del Spreadsheet de usuarios
//   2. En REMITOS_CONFIG completá el ssId de cada sucursal
//   3. En CARPETAS_PROCESADOS completá el ID de cada carpeta de Drive
//   4. Al agregar sucursal: agregá su ssId acá y republicá el script
// ═══════════════════════════════════════════════════════════════════════════

'use strict';

// ── CONFIGURACIÓN ────────────────────────────────────────────────────────────
const API_KEY      = 'claveparamostrarremitos'; // ← Cambiar por clave propia
const USERS_SS_ID  = 'ID_DEL_SPREADSHEET_USUARIOS'; // ← COMPLETAR
const TOKEN_TTL_MS = 60 * 60 * 1000; // 1 hora

// Spreadsheet de datos por sucursal
// ssId: ID en la URL entre /d/ y /edit · sheetName: nombre exacto de la hoja
const REMITOS_CONFIG = {
  rosario: { ssId: 'ID_DEL_SPREADSHEET_CONTROL_REMITOS_OCR',         sheetName: 'Datos' }, // ← COMPLETAR
  santafe: { ssId: 'ID_DEL_SPREADSHEET_CONTROL_REMITOS_OCR_SANTAFE', sheetName: 'Datos' }, // ← COMPLETAR
  // Nueva sucursal: cordoba: { ssId: 'ID_NUEVO', sheetName: 'Datos' },
};

// Carpetas de Drive donde se mueven los archivos al marcar como procesado
// ID: abrir carpeta en Drive → URL → copiar lo que está después de /folders/
const CARPETAS_PROCESADOS = {
  santafe: 'ID_CARPETA_PROCESADOS_SANTAFE', // ← COMPLETAR — "Remitos Procesados Santa Fe"
  rosario: 'ID_CARPETA_PROCESADOS_ROSARIO', // ← COMPLETAR
};

// ── MICRO-UTILIDADES ─────────────────────────────────────────────────────────
// Funciones puras, sin efecto secundario, sin llamadas a API.
// Definidas como const arrow para que V8 las inline donde sea posible.

const jsonResp    = d => ContentService.createTextOutput(JSON.stringify(d))
                           .setMimeType(ContentService.MimeType.JSON);

const openUsersSS = ()  => SpreadsheetApp.openById(USERS_SS_ID);

// Normaliza cualquier valor de celda a string sin espacios extremos
const str         = v   => String(v ?? '').trim();

// Interpreta el campo "activo" de una fila: false literal o string "false"
const isInactivo  = v   => v === false || str(v).toLowerCase() === 'false';

// Extrae el ID de Drive de una URL completa o devuelve el valor tal cual
// si ya es un ID puro (sin barras)
const driveId     = url => {
  if (!url) return '';
  const s = str(url);
  if (!s.includes('/')) return s;          // ya es un ID puro
  const m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
};

// ── ENTRY POINT ──────────────────────────────────────────────────────────────
function doGet(e) {
  const p      = e.parameter;
  const action = p.action || 'remitos';

  if (p.key !== API_KEY) return jsonResp({ error: 'No autorizado' });

  try {
    switch (action) {
      case 'login':           return handleLogin(p);
      case 'remitos':         return handleRemitos(p);
      case 'logout':          return handleLogout(p);
      case 'updateProcesado': return handleUpdateProcesado(p);
      case 'procesarRemito':  return handleProcesarRemito(p);
      case 'editarRemito':    return handleEditarRemito(p);
      case 'getSucursales':   return handleGetSucursales(p);
      case 'addSucursal':     return handleAddSucursal(p);
      case 'deleteSucursal':  return handleDeleteSucursal(p);
      default:                return jsonResp({ error: 'Accion invalida' });
    }
  } catch (err) {
    return jsonResp({ error: 'Error interno del servidor' });
  }
}

// ── LOGIN ────────────────────────────────────────────────────────────────────
// Una apertura de SS · getValues() una vez · Array.find() en lugar de for-loop
function handleLogin(p) {
  const email = str(p.email).toLowerCase();
  const hash  = str(p.passwordHash).toLowerCase();

  if (!email || !hash) return jsonResp({ ok: false, error: 'Datos incompletos' });

  const ss    = openUsersSS();
  const sheet = ss.getSheetByName('Usuarios');
  if (!sheet) return jsonResp({ ok: false, error: 'Hoja Usuarios no encontrada' });

  // Una lectura en bloque · cols: 0=email 1=hash 2=name 3=role 4=sucursal 5=activo
  const match = sheet.getDataRange().getValues().slice(1).find(r =>
    str(r[0]).toLowerCase() === email &&
    str(r[1]).toLowerCase() === hash  &&
    !isInactivo(r[5])
  );

  if (!match) return jsonResp({ ok: false, error: 'Credenciales incorrectas' });

  const [,, uName, uRole, uSuc] = match;
  const token      = _crearSesion(ss, email, str(uRole), str(uSuc));
  const sucursales = _leerSucursales(ss); // mismo SS — sin segunda apertura

  return jsonResp({
    ok: true, token,
    name:      str(uName),
    role:      str(uRole),
    sucursal:  str(uSuc) || null,
    sucursales,
  });
}

// ── SESIONES (privadas) ──────────────────────────────────────────────────────

// appendRow + limpieza en bloque con setValues()
function _crearSesion(ss, email, role, sucursal) {
  const token  = Utilities.getUuid();
  const expira = new Date(Date.now() + TOKEN_TTL_MS).toISOString();
  const sheet  = ss.getSheetByName('Sesiones');
  sheet.appendRow([token, email, role, sucursal || '', expira, new Date().toISOString()]);
  _limpiarSesionesExpiradas(sheet);
  return token;
}

// Reemplaza deleteRow-en-loop por filter + setValues en bloque
// deleteRow() = 1 llamada a API por fila → O(n) llamadas
// setValues()  = 1 llamada total → O(1) llamadas
function _limpiarSesionesExpiradas(sheet) {
  const rows  = sheet.getDataRange().getValues();
  const ahora = Date.now();
  const [header, ...data] = rows;
  const vigentes = data.filter(r => r[0] && new Date(r[4]).getTime() > ahora);
  sheet.clearContents();
  const newData = [header, ...vigentes];
  sheet.getRange(1, 1, newData.length, header.length).setValues(newData);
}

// getValues() una vez · find() en lugar de for-loop
function _verificarToken(token) {
  if (!token) return null;
  const sheet = openUsersSS().getSheetByName('Sesiones');
  if (!sheet) return null;
  const ahora = Date.now();
  const row   = sheet.getDataRange().getValues().slice(1).find(r =>
    String(r[0]) === token && new Date(r[4]).getTime() > ahora
  );
  return row ? { email: str(row[1]), role: str(row[2]), sucursal: str(row[3]) } : null;
}

function handleLogout(p) {
  if (!p.token) return jsonResp({ ok: true });
  const sheet = openUsersSS().getSheetByName('Sesiones');
  if (!sheet)  return jsonResp({ ok: true });
  const rows = sheet.getDataRange().getValues();
  const [header, ...data] = rows;
  const filtradas = data.filter(r => String(r[0]) !== p.token);
  sheet.clearContents();
  const newData = [header, ...filtradas];
  sheet.getRange(1, 1, newData.length, header.length).setValues(newData);
  return jsonResp({ ok: true });
}

// ── SUCURSALES ───────────────────────────────────────────────────────────────
// cols: 0=slug 1=label 2=icon 3=color 4=activo 5=creado

// filter + map en bloque V8 · sin for-loop manual
function _leerSucursales(ss) {
  const sheet = ss.getSheetByName('Sucursales');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] && !isInactivo(r[4]))
    .map(r => ({
      slug:  str(r[0]),
      label: str(r[1]) || str(r[0]),
      icon:  str(r[2]) || '🏢',
      color: str(r[3]) || '#2756d6',
    }));
}

function handleGetSucursales(p) {
  if (!_verificarToken(p.token)) return jsonResp({ error: 'Sesion invalida' });
  return jsonResp({ sucursales: _leerSucursales(openUsersSS()) });
}

function handleAddSucursal(p) {
  const perfil = _verificarToken(p.token);
  if (!perfil || perfil.role !== 'admin') return jsonResp({ error: 'Sin permisos' });

  const slug  = str(p.slug || p.key).toLowerCase().replace(/[^a-z0-9]/g, '');
  const label = str(p.label);
  const icon  = str(p.icon)  || '🏢';
  const color = str(p.color) || '#2756d6';

  if (!slug || slug.length < 2)  return jsonResp({ error: 'Slug inválido (mínimo 2 caracteres)' });
  if (slug.length > 20)          return jsonResp({ error: 'Slug demasiado largo (máximo 20)' });
  if (!label)                    return jsonResp({ error: 'Nombre requerido' });

  const ss  = openUsersSS();
  let sheet = ss.getSheetByName('Sucursales');
  if (!sheet) {
    sheet = ss.insertSheet('Sucursales');
    sheet.appendRow(['slug', 'label', 'icon', 'color', 'activo', 'creado']);
  }

  // Array.some() para chequeo de duplicado · una lectura en bloque
  const rows = sheet.getDataRange().getValues();
  if (rows.slice(1).some(r => str(r[0]).toLowerCase() === slug)) {
    return jsonResp({ error: "Ya existe una sucursal con slug '" + slug + "'" });
  }

  sheet.appendRow([slug, label, icon, color, true, new Date().toISOString()]);
  return jsonResp({
    ok: true, slug,
    aviso: "Sucursal '" + slug + "' creada. Agregá su ssId en REMITOS_CONFIG y republicá.",
  });
}

// setValues() de una celda para marcar inactivo (preserva historial)
function handleDeleteSucursal(p) {
  const perfil = _verificarToken(p.token);
  if (!perfil || perfil.role !== 'admin') return jsonResp({ error: 'Sin permisos' });

  const slug = str(p.slug || p.key).toLowerCase();
  if (!slug) return jsonResp({ error: 'Slug requerido' });

  const sheet = openUsersSS().getSheetByName('Sucursales');
  if (!sheet)  return jsonResp({ error: 'Hoja Sucursales no encontrada' });

  // findIndex() en lugar de for-loop
  const rows = sheet.getDataRange().getValues();
  const idx  = rows.slice(1).findIndex(r => str(r[0]).toLowerCase() === slug);
  if (idx === -1) return jsonResp({ error: "Sucursal no encontrada: '" + slug + "'" });

  sheet.getRange(idx + 2, 5).setValue(false); // 1 llamada, sin deleteRow
  return jsonResp({ ok: true, slug });
}

// ── REMITOS ──────────────────────────────────────────────────────────────────
// getValues() una vez · filter + map V8 · sin for-loop
// Columnas Sheet: A(1)=DriveId B(2)=Archivo C(3)=Remito D(4)=Fecha
//                 E(5)=Cliente F(6)=Link   G(7)=Procesado H(8)=Timestamp
function handleRemitos(p) {
  const perfil   = _verificarToken(p.token);
  const sucursal = str(p.sucursal);

  if (!perfil)   return jsonResp({ error: 'Sesion invalida o expirada' });
  if (!sucursal) return jsonResp({ error: 'Sucursal requerida' });

  const config = REMITOS_CONFIG[sucursal];
  if (!config)   return jsonResp({ error: "Sucursal '" + sucursal + "' no configurada. Agregala en REMITOS_CONFIG y republicá." });
  if (perfil.role !== 'admin' && perfil.sucursal !== sucursal) {
    return jsonResp({ error: 'Sin acceso a esta sucursal' });
  }

  const sheet = SpreadsheetApp.openById(config.ssId).getSheetByName(config.sheetName);
  if (!sheet) return jsonResp({ error: "Hoja '" + config.sheetName + "' no encontrada" });

  const isAdmin = perfil.role === 'admin';

  // Intentar leer del cache primero (válido 60 segundos)
  const cacheKey = 'remitos_' + sucursal + '_' + perfil.role;
  try {
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) return jsonResp({ ok: true, remitos: JSON.parse(cached), fromCache: true });
  } catch(e) {}

  // Leer solo columnas necesarias — excluir col I (texto OCR hasta 5000 chars)
  // Columnas: A(0)=driveId B(1)=archivo C(2)=remito D(3)=fecha E(4)=cliente F(5)=link G(6)=estado/procesado
  const lastRow = sheet.getLastRow();
  const allRows = lastRow > 1
    ? sheet.getRange(1, 1, lastRow, 7).getValues()  // solo primeras 7 cols, sin OCR
    : [[]];

  const remitos = allRows.slice(1)
    .map((row, i) => ({ row, rowIndex: i + 2 }))
    .filter(({ row }) => str(row[1]) || str(row[2]))
    .map(({ row, rowIndex }) => ({
      driveId:   str(row[0]),
      archivo:   str(row[1]),
      remito:    str(row[2]),
      fecha:     str(row[3]),
      cliente:   str(row[4]),
      link:      str(row[5]),
      procesado: str(row[6]),  // siempre incluir — app lo usa para badge
      rowIndex,
    }));

  // Guardar en cache 60s para acelerar la siguiente carga
  try { CacheService.getScriptCache().put(cacheKey, JSON.stringify(remitos), 60); } catch(e) {}
  return jsonResp({ ok: true, remitos });
}

// ── ACTUALIZAR PROCESADO (toggle simple) ─────────────────────────────────────
// setValues() de 1 celda · sin leer toda la hoja
function handleUpdateProcesado(p) {
  const perfil   = _verificarToken(p.token);
  const sucursal = str(p.sucursal);
  const rowIndex = parseInt(p.rowIndex, 10);
  const valor    = str(p.valor);

  if (!perfil || perfil.role !== 'admin') return jsonResp({ error: 'Sin permisos' });

  const config = REMITOS_CONFIG[sucursal];
  if (!config || !rowIndex || isNaN(rowIndex)) return jsonResp({ error: 'Parámetros inválidos' });

  const sheet = SpreadsheetApp.openById(config.ssId).getSheetByName(config.sheetName);
  if (!sheet) return jsonResp({ error: 'Hoja no encontrada' });

  sheet.getRange(rowIndex, 7).setValue(valor); // Col G = Procesado
  return jsonResp({ ok: true, rowIndex, valor });
}

// ── PROCESAR REMITO OK ────────────────────────────────────────────────────────
// Lee fila completa en 1 llamada · escribe 2 celdas en 1 llamada de bloque
// Mueve archivo Drive: addFile + removeFile (compatible con todas las versiones)
function handleProcesarRemito(p) {
  const perfil   = _verificarToken(p.token);
  const sucursal = str(p.sucursal);
  const rowIndex = parseInt(p.rowIndex, 10);
  const accion   = str(p.accion).toLowerCase(); // 'ok' | 'pendiente'

  if (!perfil || perfil.role !== 'admin') return jsonResp({ error: 'Sin permisos' });

  const config = REMITOS_CONFIG[sucursal];
  if (!config || !rowIndex || isNaN(rowIndex)) return jsonResp({ error: 'Parámetros inválidos' });

  const sheet = SpreadsheetApp.openById(config.ssId).getSheetByName(config.sheetName);
  if (!sheet) return jsonResp({ error: 'Hoja no encontrada' });

  // Leer la fila entera en 1 sola llamada (cols 1-8)
  const rowData  = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
  const fileId   = driveId(str(rowData[0]) || str(rowData[5])); // Col A (ID) o Col F (link)
  const timestamp = new Date().toISOString();

  if (accion === 'ok') {
    let moveMsg = '';
    const carpetaId = CARPETAS_PROCESADOS[sucursal] || '';
    const carpetaConfigurada = carpetaId && !carpetaId.startsWith('ID_CARPETA');

    // Mover archivo en Drive solo si hay fileId y carpeta configurada
    if (fileId && carpetaConfigurada) {
      try {
        const file    = DriveApp.getFileById(fileId);
        const destino = DriveApp.getFolderById(carpetaId);
        destino.addFile(file);
        // Quitar de todas las carpetas anteriores excepto la de destino
        // hasNext/next en lugar de Array para compatibilidad con Iterator de Drive
        const parents = file.getParents();
        while (parents.hasNext()) {
          const parent = parents.next();
          if (parent.getId() !== carpetaId) parent.removeFile(file);
        }
        moveMsg = 'Archivo movido a carpeta de procesados.';
      } catch (driveErr) {
        moveMsg = 'Advertencia: no se pudo mover el archivo (' + driveErr.message + ')';
      }
    } else if (!carpetaConfigurada) {
      moveMsg = 'Carpeta de procesados no configurada en el script. Solo se actualizó el Sheet.';
    }

    // Escribir cols G y H en 1 sola llamada de bloque
    sheet.getRange(rowIndex, 7, 1, 2).setValues([['OK', timestamp]]);
    return jsonResp({ ok: true, accion: 'ok', rowIndex, moveMsg });

  } else if (accion === 'pendiente') {
    // Limpiar cols G y H en 1 sola llamada de bloque
    sheet.getRange(rowIndex, 7, 1, 2).setValues([['', '']]);
    return jsonResp({ ok: true, accion: 'pendiente', rowIndex });

  } else {
    return jsonResp({ error: "Accion inválida. Usar 'ok' o 'pendiente'" });
  }
}

// ── EDITAR REMITO ─────────────────────────────────────────────────────────────
// Verifica col G en 1 llamada · escribe cols B-E en 1 sola llamada de bloque
// Solo permite editar remitos NO procesados
function handleEditarRemito(p) {
  const perfil   = _verificarToken(p.token);
  const sucursal = str(p.sucursal);
  const rowIndex = parseInt(p.rowIndex, 10);

  if (!perfil || perfil.role !== 'admin') return jsonResp({ error: 'Sin permisos' });

  const config = REMITOS_CONFIG[sucursal];
  if (!config || !rowIndex || isNaN(rowIndex)) return jsonResp({ error: 'Parámetros inválidos' });

  const sheet = SpreadsheetApp.openById(config.ssId).getSheetByName(config.sheetName);
  if (!sheet) return jsonResp({ error: 'Hoja no encontrada' });

  // Verificar procesado: leer solo col G (1 celda, no toda la hoja)
  const procesado = str(sheet.getRange(rowIndex, 7).getValue()).toUpperCase();
  if (procesado === 'OK') return jsonResp({ error: 'No se puede editar un remito ya procesado' });

  const archivo = str(p.archivo);
  const remito  = str(p.remito);
  const fecha   = str(p.fecha);
  const cliente = str(p.cliente);

  // Escribir 4 celdas en 1 sola llamada de bloque (cols B-E = 2 a 5)
  sheet.getRange(rowIndex, 2, 1, 4).setValues([[archivo, remito, fecha, cliente]]);

  return jsonResp({ ok: true, rowIndex, archivo, remito, fecha, cliente });
}
