/**
 * Nexo CRM — backend privado sobre Google Sheets.
 *
 * INSTALACIÓN (una sola vez):
 * 1. Cambiá ADMIN_EMAIL y ADMIN_PASSWORD.
 * 2. Desde el editor de Apps Script ejecutá configurarCRM().
 * 3. Implementá como "Aplicación web":
 *      - Ejecutar como: yo
 *      - Quién tiene acceso: cualquier usuario
 *    La aplicación sigue protegida: ninguna operación entrega datos sin una
 *    sesión válida iniciada con un usuario de la hoja "Vendedores".
 * 4. Copiá la URL /exec y pegala en APPS_SCRIPT_URL dentro de index.html.
 */

const CRM_CONFIG = Object.freeze({
  APP_NAME: 'Nexo CRM',
  ADMIN_NAME: 'Administrador',
  ADMIN_EMAIL: 'aquarianos10@gmail.com',
  ADMIN_PASSWORD: 'Juliana.123',
  SESSION_SECONDS: 21600, // 6 horas
  STALE_DAYS: 7,
  TIMEZONE: Session.getScriptTimeZone() || 'America/Argentina/Buenos_Aires'
});

const CRM_SHEETS = Object.freeze({
  Contactos: [
    'id', 'createdAt', 'updatedAt', 'name', 'company', 'email', 'phone',
    'status', 'source', 'owner', 'notes', 'archived'
  ],
  Oportunidades: [
    'id', 'createdAt', 'updatedAt', 'title', 'contactId', 'value', 'stage',
    'probability', 'owner', 'expectedClose', 'lastActivity', 'notes',
    'lostReason', 'archived'
  ],
  Seguimientos: [
    'id', 'createdAt', 'updatedAt', 'opportunityId', 'contactId', 'title',
    'dueAt', 'type', 'owner', 'status', 'notes', 'completedAt'
  ],
  Conversaciones: [
    'id', 'createdAt', 'contactId', 'opportunityId', 'channel', 'direction',
    'summary', 'owner'
  ],
  Cotizaciones: [
    'id', 'createdAt', 'updatedAt', 'quoteNumber', 'contactId',
    'opportunityId', 'subtotal', 'taxRate', 'taxAmount', 'total', 'currency',
    'status', 'validUntil', 'itemsJson', 'notes', 'owner'
  ],
  Vendedores: [
    'id', 'createdAt', 'updatedAt', 'name', 'email', 'passwordHash', 'role',
    'active'
  ],
  Configuracion: ['key', 'value']
});

const CRM_ALLOWED_STAGES = Object.freeze([
  'Nuevo', 'Contactado', 'Propuesta', 'Negociación', 'Ganado', 'Perdido'
]);

/**
 * Ejecutar manualmente una vez desde el editor vinculado a Google Sheets.
 */
function configurarCRM() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('Abrí el script desde una hoja de Google Sheets y volvé a ejecutar.');
  }
  if (
    CRM_CONFIG.ADMIN_PASSWORD === 'CAMBIAR_ESTA_CLAVE' ||
    CRM_CONFIG.ADMIN_PASSWORD.length < 8
  ) {
    throw new Error('Antes de continuar, cambiá ADMIN_PASSWORD por una clave de al menos 8 caracteres.');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('CRM_SPREADSHEET_ID', active.getId());
  if (!props.getProperty('CRM_PASSWORD_SALT')) {
    props.setProperty('CRM_PASSWORD_SALT', Utilities.getUuid() + Utilities.getUuid());
  }

  Object.keys(CRM_SHEETS).forEach(function (name) {
    ensureSheet_(active, name, CRM_SHEETS[name]);
  });

  const settings = getRows_('Configuracion');
  if (!settings.some(function (item) { return item.key === 'staleDays'; })) {
    appendObject_('Configuracion', { key: 'staleDays', value: CRM_CONFIG.STALE_DAYS });
  }
  if (!settings.some(function (item) { return item.key === 'companyName'; })) {
    appendObject_('Configuracion', { key: 'companyName', value: 'Mi empresa' });
  }

  const sellers = getRows_('Vendedores');
  const adminEmail = normalizeEmail_(CRM_CONFIG.ADMIN_EMAIL);
  const currentAdmin = sellers.find(function (seller) {
    return normalizeEmail_(seller.email) === adminEmail;
  });
  const now = isoNow_();
  const admin = {
    id: currentAdmin ? currentAdmin.id : newId_('USR'),
    createdAt: currentAdmin ? currentAdmin.createdAt : now,
    updatedAt: now,
    name: CRM_CONFIG.ADMIN_NAME,
    email: adminEmail,
    passwordHash: hashPassword_(CRM_CONFIG.ADMIN_PASSWORD),
    role: 'admin',
    active: true
  };
  upsertObject_('Vendedores', admin);

  active.setSpreadsheetTimeZone(CRM_CONFIG.TIMEZONE);
  styleSheets_(active);

  return {
    ok: true,
    message: 'CRM configurado. Ya podés implementar la aplicación web.',
    spreadsheetId: active.getId(),
    adminEmail: adminEmail
  };
}

/**
 * Volvé a ejecutar esta función después de cambiar ADMIN_PASSWORD si necesitás
 * restablecer la clave del administrador.
 */
function restablecerClaveAdministrador() {
  if (
    CRM_CONFIG.ADMIN_PASSWORD === 'CAMBIAR_ESTA_CLAVE' ||
    CRM_CONFIG.ADMIN_PASSWORD.length < 8
  ) {
    throw new Error('Definí una ADMIN_PASSWORD segura antes de restablecerla.');
  }
  const sellers = getRows_('Vendedores');
  const adminEmail = normalizeEmail_(CRM_CONFIG.ADMIN_EMAIL);
  const admin = sellers.find(function (seller) {
    return normalizeEmail_(seller.email) === adminEmail;
  });
  if (!admin) throw new Error('No se encontró el administrador. Ejecutá configurarCRM().');
  admin.passwordHash = hashPassword_(CRM_CONFIG.ADMIN_PASSWORD);
  admin.updatedAt = isoNow_();
  upsertObject_('Vendedores', admin);
  return { ok: true, message: 'Clave actualizada.' };
}

function doGet(e) {
  const callback = e && e.parameter ? e.parameter.callback : '';
  const result = {
    ok: true,
    app: CRM_CONFIG.APP_NAME,
    status: 'online',
    timestamp: isoNow_()
  };
  return jsonOutput_(result, callback);
}

function doPost(e) {
  let request = {};
  try {
    request = parseRequest_(e);
    const action = String(request.action || '').trim();
    if (!action) throw apiError_('Falta indicar la acción.', 'BAD_REQUEST');

    if (action === 'login') return jsonOutput_(login_(request));
    if (action === 'ping') return jsonOutput_({ ok: true, timestamp: isoNow_() });

    const session = requireSession_(request.token);
    let payload;

    switch (action) {
      case 'bootstrap':
        payload = bootstrap_(session);
        break;
      case 'saveContact':
        payload = withLock_(function () { return saveContact_(request.data, session); });
        break;
      case 'archiveContact':
        payload = withLock_(function () { return archiveContact_(request.id, session); });
        break;
      case 'saveOpportunity':
        payload = withLock_(function () { return saveOpportunity_(request.data, session); });
        break;
      case 'archiveOpportunity':
        payload = withLock_(function () { return archiveOpportunity_(request.id, session); });
        break;
      case 'changeStage':
        payload = withLock_(function () {
          return changeStage_(request.id, request.stage, session);
        });
        break;
      case 'saveFollowUp':
        payload = withLock_(function () { return saveFollowUp_(request.data, session); });
        break;
      case 'completeFollowUp':
        payload = withLock_(function () {
          return completeFollowUp_(request.id, request.completed, session);
        });
        break;
      case 'saveConversation':
        payload = withLock_(function () { return saveConversation_(request.data, session); });
        break;
      case 'saveQuote':
        payload = withLock_(function () { return saveQuote_(request.data, session); });
        break;
      case 'saveSeller':
        assertAdmin_(session);
        payload = withLock_(function () { return saveSeller_(request.data, session); });
        break;
      case 'logout':
        payload = logout_(request.token);
        break;
      default:
        throw apiError_('Acción no reconocida.', 'NOT_FOUND');
    }

    return jsonOutput_(payload);
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    return jsonOutput_({
      ok: false,
      error: error && error.message ? error.message : 'Ocurrió un error inesperado.',
      code: error && error.code ? error.code : 'SERVER_ERROR'
    });
  }
}

function login_(request) {
  const email = normalizeEmail_(request.email);
  const password = String(request.password || '');
  if (!email || !password) throw apiError_('Ingresá tu email y contraseña.', 'VALIDATION');
  assertLoginAllowed_(email);

  const seller = getRows_('Vendedores').find(function (item) {
    return normalizeEmail_(item.email) === email && toBoolean_(item.active);
  });

  if (!seller || !constantTimeEquals_(seller.passwordHash, hashPassword_(password))) {
    recordLoginFailure_(email);
    Utilities.sleep(350);
    throw apiError_('Email o contraseña incorrectos.', 'UNAUTHORIZED');
  }
  CacheService.getScriptCache().remove('LOGIN_FAIL_' + sha256_(email));

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const session = {
    id: seller.id,
    name: seller.name,
    email: normalizeEmail_(seller.email),
    role: seller.role || 'seller',
    issuedAt: Date.now()
  };
  CacheService.getScriptCache().put(
    'SESSION_' + sha256_(token),
    JSON.stringify(session),
    CRM_CONFIG.SESSION_SECONDS
  );

  return {
    ok: true,
    token: token,
    user: publicSeller_(seller),
    expiresIn: CRM_CONFIG.SESSION_SECONDS
  };
}

function assertLoginAllowed_(email) {
  const attempts = Number(
    CacheService.getScriptCache().get('LOGIN_FAIL_' + sha256_(email)) || 0
  );
  if (attempts >= 8) {
    throw apiError_(
      'Demasiados intentos. Esperá 10 minutos antes de volver a ingresar.',
      'RATE_LIMIT'
    );
  }
}

function recordLoginFailure_(email) {
  const cache = CacheService.getScriptCache();
  const key = 'LOGIN_FAIL_' + sha256_(email);
  const attempts = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(attempts), 600);
}

function requireSession_(token) {
  const cleanToken = String(token || '');
  if (!cleanToken) throw apiError_('Tu sesión venció. Volvé a ingresar.', 'UNAUTHORIZED');
  const raw = CacheService.getScriptCache().get('SESSION_' + sha256_(cleanToken));
  if (!raw) throw apiError_('Tu sesión venció. Volvé a ingresar.', 'UNAUTHORIZED');
  return JSON.parse(raw);
}

function logout_(token) {
  CacheService.getScriptCache().remove('SESSION_' + sha256_(String(token || '')));
  return { ok: true };
}

function bootstrap_(session) {
  const settings = {};
  getRows_('Configuracion').forEach(function (row) {
    settings[row.key] = row.value;
  });

  const conversations = getRows_('Conversaciones')
    .sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); })
    .slice(0, 500);

  return {
    ok: true,
    user: session,
    settings: settings,
    contacts: activeRows_(getRows_('Contactos')),
    opportunities: activeRows_(getRows_('Oportunidades')),
    followUps: getRows_('Seguimientos'),
    conversations: conversations,
    quotes: getRows_('Cotizaciones'),
    sellers: getRows_('Vendedores').filter(function (seller) {
      return toBoolean_(seller.active);
    }).map(publicSeller_),
    serverTime: isoNow_()
  };
}

function saveContact_(data, session) {
  data = data || {};
  requireText_(data.name, 'El nombre del contacto es obligatorio.');
  const current = data.id ? findById_('Contactos', data.id) : null;
  const now = isoNow_();
  const contact = {
    id: current ? current.id : newId_('CON'),
    createdAt: current ? current.createdAt : now,
    updatedAt: now,
    name: cleanText_(data.name, 120),
    company: cleanText_(data.company, 120),
    email: normalizeEmail_(data.email),
    phone: cleanText_(data.phone, 60),
    status: cleanChoice_(data.status, ['Prospecto', 'Cliente', 'Inactivo'], 'Prospecto'),
    source: cleanText_(data.source, 80),
    owner: cleanText_(data.owner || session.name, 120),
    notes: cleanText_(data.notes, 2000),
    archived: false
  };
  upsertObject_('Contactos', contact);
  return { ok: true, item: contact };
}

function archiveContact_(id) {
  const contact = requireRecord_('Contactos', id, 'Contacto no encontrado.');
  contact.archived = true;
  contact.updatedAt = isoNow_();
  upsertObject_('Contactos', contact);
  return { ok: true, id: id };
}

function saveOpportunity_(data, session) {
  data = data || {};
  requireText_(data.title, 'El nombre de la oportunidad es obligatorio.');
  const current = data.id ? findById_('Oportunidades', data.id) : null;
  const now = isoNow_();
  const stage = cleanChoice_(data.stage, CRM_ALLOWED_STAGES, 'Nuevo');
  const opportunity = {
    id: current ? current.id : newId_('OPP'),
    createdAt: current ? current.createdAt : now,
    updatedAt: now,
    title: cleanText_(data.title, 160),
    contactId: cleanText_(data.contactId, 80),
    value: clampNumber_(data.value, 0, 999999999999, 0),
    stage: stage,
    probability: clampNumber_(
      data.probability,
      0,
      100,
      defaultProbability_(stage)
    ),
    owner: cleanText_(data.owner || session.name, 120),
    expectedClose: cleanDate_(data.expectedClose),
    lastActivity: current ? (current.lastActivity || now) : now,
    notes: cleanText_(data.notes, 3000),
    lostReason: stage === 'Perdido' ? cleanText_(data.lostReason, 500) : '',
    archived: false
  };
  upsertObject_('Oportunidades', opportunity);
  return { ok: true, item: opportunity };
}

function changeStage_(id, stage, session) {
  const nextStage = cleanChoice_(stage, CRM_ALLOWED_STAGES, '');
  if (!nextStage) throw apiError_('Etapa inválida.', 'VALIDATION');
  const opportunity = requireRecord_('Oportunidades', id, 'Oportunidad no encontrada.');
  const previous = opportunity.stage;
  opportunity.stage = nextStage;
  opportunity.probability = defaultProbability_(nextStage);
  opportunity.lastActivity = isoNow_();
  opportunity.updatedAt = isoNow_();
  upsertObject_('Oportunidades', opportunity);

  appendObject_('Conversaciones', {
    id: newId_('ACT'),
    createdAt: isoNow_(),
    contactId: opportunity.contactId,
    opportunityId: opportunity.id,
    channel: 'Sistema',
    direction: 'Nota',
    summary: 'Etapa actualizada: ' + previous + ' → ' + nextStage,
    owner: session.name
  });
  return { ok: true, item: opportunity };
}

function archiveOpportunity_(id) {
  const opportunity = requireRecord_('Oportunidades', id, 'Oportunidad no encontrada.');
  opportunity.archived = true;
  opportunity.updatedAt = isoNow_();
  upsertObject_('Oportunidades', opportunity);
  return { ok: true, id: id };
}

function saveFollowUp_(data, session) {
  data = data || {};
  requireText_(data.title, 'El asunto del seguimiento es obligatorio.');
  if (!cleanDateTime_(data.dueAt)) {
    throw apiError_('Elegí la fecha y hora del seguimiento.', 'VALIDATION');
  }
  const current = data.id ? findById_('Seguimientos', data.id) : null;
  const now = isoNow_();
  const followUp = {
    id: current ? current.id : newId_('SEG'),
    createdAt: current ? current.createdAt : now,
    updatedAt: now,
    opportunityId: cleanText_(data.opportunityId, 80),
    contactId: cleanText_(data.contactId, 80),
    title: cleanText_(data.title, 180),
    dueAt: cleanDateTime_(data.dueAt),
    type: cleanChoice_(
      data.type,
      ['Llamada', 'Email', 'Reunión', 'WhatsApp', 'Tarea'],
      'Tarea'
    ),
    owner: cleanText_(data.owner || session.name, 120),
    status: current ? current.status : 'Pendiente',
    notes: cleanText_(data.notes, 2000),
    completedAt: current ? current.completedAt : ''
  };
  upsertObject_('Seguimientos', followUp);
  return { ok: true, item: followUp };
}

function completeFollowUp_(id, completed) {
  const followUp = requireRecord_('Seguimientos', id, 'Seguimiento no encontrado.');
  const isCompleted = completed !== false;
  followUp.status = isCompleted ? 'Completado' : 'Pendiente';
  followUp.completedAt = isCompleted ? isoNow_() : '';
  followUp.updatedAt = isoNow_();
  upsertObject_('Seguimientos', followUp);
  return { ok: true, item: followUp };
}

function saveConversation_(data, session) {
  data = data || {};
  requireText_(data.summary, 'Escribí un resumen de la conversación.');
  const conversation = {
    id: newId_('MSG'),
    createdAt: isoNow_(),
    contactId: cleanText_(data.contactId, 80),
    opportunityId: cleanText_(data.opportunityId, 80),
    channel: cleanChoice_(
      data.channel,
      ['Llamada', 'Email', 'WhatsApp', 'Reunión', 'Nota', 'Sistema'],
      'Nota'
    ),
    direction: cleanChoice_(data.direction, ['Entrante', 'Saliente', 'Nota'], 'Nota'),
    summary: cleanText_(data.summary, 3000),
    owner: cleanText_(data.owner || session.name, 120)
  };
  appendObject_('Conversaciones', conversation);

  if (conversation.opportunityId) {
    const opportunity = findById_('Oportunidades', conversation.opportunityId);
    if (opportunity) {
      opportunity.lastActivity = isoNow_();
      opportunity.updatedAt = isoNow_();
      upsertObject_('Oportunidades', opportunity);
    }
  }
  return { ok: true, item: conversation };
}

function saveQuote_(data, session) {
  data = data || {};
  const items = Array.isArray(data.items) ? data.items : [];
  const normalizedItems = items
    .map(function (item) {
      const quantity = clampNumber_(item.quantity, 0, 999999, 1);
      const price = clampNumber_(item.price, 0, 999999999999, 0);
      return {
        description: cleanText_(item.description, 200),
        quantity: quantity,
        price: price,
        total: roundMoney_(quantity * price)
      };
    })
    .filter(function (item) { return item.description; });

  if (!normalizedItems.length) {
    throw apiError_('Agregá al menos un ítem a la cotización.', 'VALIDATION');
  }

  const current = data.id ? findById_('Cotizaciones', data.id) : null;
  const subtotal = roundMoney_(normalizedItems.reduce(function (sum, item) {
    return sum + item.total;
  }, 0));
  const taxRate = clampNumber_(data.taxRate, 0, 100, 0);
  const taxAmount = roundMoney_(subtotal * taxRate / 100);
  const now = isoNow_();
  const quote = {
    id: current ? current.id : newId_('COT'),
    createdAt: current ? current.createdAt : now,
    updatedAt: now,
    quoteNumber: current ? current.quoteNumber : nextQuoteNumber_(),
    contactId: cleanText_(data.contactId, 80),
    opportunityId: cleanText_(data.opportunityId, 80),
    subtotal: subtotal,
    taxRate: taxRate,
    taxAmount: taxAmount,
    total: roundMoney_(subtotal + taxAmount),
    currency: cleanChoice_(data.currency, ['ARS', 'USD', 'EUR', 'UYU', 'CLP', 'MXN'], 'ARS'),
    status: cleanChoice_(
      data.status,
      ['Borrador', 'Enviada', 'Aceptada', 'Rechazada', 'Vencida'],
      'Borrador'
    ),
    validUntil: cleanDate_(data.validUntil),
    itemsJson: JSON.stringify(normalizedItems),
    notes: cleanText_(data.notes, 2000),
    owner: cleanText_(data.owner || session.name, 120)
  };
  upsertObject_('Cotizaciones', quote);
  return { ok: true, item: quote };
}

function saveSeller_(data, session) {
  data = data || {};
  requireText_(data.name, 'El nombre del vendedor es obligatorio.');
  const email = normalizeEmail_(data.email);
  if (!email || email.indexOf('@') < 1) {
    throw apiError_('Ingresá un email válido.', 'VALIDATION');
  }
  const current = data.id ? findById_('Vendedores', data.id) : null;
  const duplicate = getRows_('Vendedores').find(function (item) {
    return normalizeEmail_(item.email) === email && item.id !== (current && current.id);
  });
  if (duplicate) throw apiError_('Ya existe un usuario con ese email.', 'VALIDATION');

  const password = String(data.password || '');
  if (!current && password.length < 8) {
    throw apiError_('La contraseña debe tener al menos 8 caracteres.', 'VALIDATION');
  }
  const now = isoNow_();
  const seller = {
    id: current ? current.id : newId_('USR'),
    createdAt: current ? current.createdAt : now,
    updatedAt: now,
    name: cleanText_(data.name, 120),
    email: email,
    passwordHash: password ? hashPassword_(password) : current.passwordHash,
    role: cleanChoice_(data.role, ['admin', 'seller'], 'seller'),
    active: data.active !== false
  };
  upsertObject_('Vendedores', seller);
  return { ok: true, item: publicSeller_(seller) };
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('CRM_SPREADSHEET_ID');
  if (!id) throw apiError_('El CRM todavía no fue configurado.', 'NOT_CONFIGURED');
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  const currentHeaders = sheet.getLastColumn()
    ? sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getLastColumn())).getValues()[0]
    : [];
  headers.forEach(function (header, index) {
    if (currentHeaders[index] && currentHeaders[index] !== header) {
      throw new Error(
        'La hoja "' + name + '" tiene una columna inesperada en la posición ' +
        (index + 1) + '. Se esperaba "' + header + '".'
      );
    }
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getRows_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw apiError_('Falta la hoja ' + sheetName + '.', 'NOT_CONFIGURED');
  const headers = CRM_SHEETS[sheetName];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map(function (row) {
    const item = {};
    headers.forEach(function (header, index) {
      item[header] = serializeCell_(row[index]);
    });
    item.__row = row;
    return item;
  });
}

function appendObject_(sheetName, item) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  const headers = CRM_SHEETS[sheetName];
  sheet.appendRow(headers.map(function (header) {
    return safeCell_(item[header]);
  }));
  return item;
}

function upsertObject_(sheetName, item) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  const headers = CRM_SHEETS[sheetName];
  const idKey = headers[0];
  const rows = getRows_(sheetName);
  const index = rows.findIndex(function (row) {
    return String(row[idKey]) === String(item[idKey]);
  });
  const values = headers.map(function (header) {
    return safeCell_(item[header]);
  });
  if (index >= 0) {
    sheet.getRange(index + 2, 1, 1, headers.length).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
  return item;
}

function findById_(sheetName, id) {
  if (!id) return null;
  return getRows_(sheetName).find(function (row) {
    return String(row.id) === String(id);
  }) || null;
}

function requireRecord_(sheetName, id, message) {
  const record = findById_(sheetName, id);
  if (!record) throw apiError_(message, 'NOT_FOUND');
  delete record.__row;
  return record;
}

function activeRows_(rows) {
  return rows.filter(function (row) {
    return !toBoolean_(row.archived);
  }).map(function (row) {
    delete row.__row;
    return row;
  });
}

function styleSheets_(spreadsheet) {
  Object.keys(CRM_SHEETS).forEach(function (name) {
    const sheet = spreadsheet.getSheetByName(name);
    const width = CRM_SHEETS[name].length;
    sheet.getRange(1, 1, 1, width)
      .setBackground('#1f1e2c')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.autoResizeColumns(1, width);
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), width)
      .setVerticalAlignment('middle');
  });
}

function parseRequest_(e) {
  const body = e && e.postData ? String(e.postData.contents || '') : '';
  if (body) {
    try {
      return JSON.parse(body);
    } catch (ignored) {}
  }
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  return e && e.parameter ? e.parameter : {};
}

function jsonOutput_(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback && /^[a-zA-Z_$][\w$.[\]]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function nextQuoteNumber_() {
  const props = PropertiesService.getScriptProperties();
  const year = Utilities.formatDate(new Date(), CRM_CONFIG.TIMEZONE, 'yyyy');
  const key = 'QUOTE_SEQUENCE_' + year;
  const next = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(next));
  return 'COT-' + year + '-' + String(next).padStart(4, '0');
}

function hashPassword_(password) {
  const salt = PropertiesService.getScriptProperties().getProperty('CRM_PASSWORD_SALT');
  if (!salt) throw apiError_('El CRM todavía no fue configurado.', 'NOT_CONFIGURED');
  return sha256_(salt + '|' + String(password));
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function (byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function constantTimeEquals_(a, b) {
  a = String(a || '');
  b = String(b || '');
  let mismatch = a.length ^ b.length;
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i += 1) {
    mismatch |= (a.charCodeAt(i % Math.max(a.length, 1)) || 0) ^
      (b.charCodeAt(i % Math.max(b.length, 1)) || 0);
  }
  return mismatch === 0;
}

function apiError_(message, code) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function assertAdmin_(session) {
  if (!session || session.role !== 'admin') {
    throw apiError_('Esta acción requiere permisos de administrador.', 'FORBIDDEN');
  }
}

function publicSeller_(seller) {
  return {
    id: seller.id,
    name: seller.name,
    email: normalizeEmail_(seller.email),
    role: seller.role || 'seller',
    active: toBoolean_(seller.active)
  };
}

function requireText_(value, message) {
  if (!String(value || '').trim()) throw apiError_(message, 'VALIDATION');
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value)
    .replace(/\u0000/g, '')
    .trim()
    .slice(0, maxLength || 500);
}

function normalizeEmail_(value) {
  return cleanText_(value, 180).toLowerCase();
}

function cleanChoice_(value, choices, fallback) {
  const clean = String(value || '');
  return choices.indexOf(clean) >= 0 ? clean : fallback;
}

function cleanDate_(value) {
  if (!value) return '';
  const date = new Date(value);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, CRM_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function cleanDateTime_(value) {
  if (!value) return '';
  const date = new Date(value);
  if (isNaN(date.getTime())) return '';
  return date.toISOString();
}

function clampNumber_(value, min, max, fallback) {
  const number = Number(value);
  if (!isFinite(number)) return fallback;
  return Math.min(max, Math.max(min, number));
}

function roundMoney_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function defaultProbability_(stage) {
  return {
    Nuevo: 10,
    Contactado: 25,
    Propuesta: 50,
    Negociación: 75,
    Ganado: 100,
    Perdido: 0
  }[stage] || 10;
}

function toBoolean_(value) {
  return value === true || String(value).toLowerCase() === 'true' || value === 1;
}

function serializeCell_(value) {
  if (value instanceof Date) return value.toISOString();
  return value;
}

function safeCell_(value) {
  if (value == null) return '';
  if (typeof value === 'boolean' || typeof value === 'number') return value;
  const text = String(value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function newId_(prefix) {
  return prefix + '-' +
    Utilities.formatDate(new Date(), CRM_CONFIG.TIMEZONE, 'yyMMddHHmmss') + '-' +
    Utilities.getUuid().slice(0, 8).toUpperCase();
}

function isoNow_() {
  return new Date().toISOString();
}
