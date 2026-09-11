/**********************************************************************
 * NUESTRA BODA · Organizador  —  Google Apps Script backend
 * La hoja de cálculo es la base de datos. Cada pestaña = una entidad.
 * Este script sirve la web (Index.html) y sincroniza con Google Calendar.
 *
 * PRIMER USO:  menú  Extensiones ▸ Apps Script,  pegá este archivo,
 * ejecutá  setup()  una vez (autoriza permisos) y luego Implementar ▸
 * Nueva implementación ▸ Aplicación web.
 **********************************************************************/

const TABS = {
  Config:       ['clave', 'valor'],
  Presupuesto:  ['id', 'categoria', 'estimado', 'notas'],
  Proveedores:  ['id', 'nombre', 'categoria', 'estado', 'contacto', 'cotizado', 'pagado', 'notas'],
  Invitados:    ['id', 'invitacion', 'grupo', 'lado', 'personas', 'mesa', 'confirmacion', 'restricciones', 'notas'],
  Cronograma:   ['id', 'bloque', 'hora', 'duracion', 'titulo', 'lugar', 'responsable', 'notas'],
  Citas:        ['id', 'proveedor', 'titulo', 'inicio', 'fin', 'estado', 'notas', 'eventId']
};

// mapea el nombre lógico usado por la web -> nombre de pestaña
const TABMAP = {
  presupuesto: 'Presupuesto',
  proveedores: 'Proveedores',
  invitados:   'Invitados',
  cronograma:  'Cronograma',
  citas:       'Citas'
};

const CONFIG_DEFAULT = {
  pareja:"Claudia & Jorge",
  fecha:"2026-12-12",
  lugar:"Lima, Perú",
  moneda:"PEN",
  presupuestoTotal:195400,
  invitadosPlan:200,
  tipoCambio:3.4
};

/* Token compartido para proteger la ESCRITURA desde la web pública (Vercel).
   Debe coincidir con API_TOKEN en Index.html. Cambialo por el que quieras.
   Ojo: al estar en el HTML público no es un secreto fuerte; solo frena abuso casual. */
const TOKEN = 'boda-cj-2026';

/* ------------------------------------------------------------------ */
/*  WEB APP ENTRY                                                       */
/*  - Sin ?action  -> sirve la app HTML (uso nativo con google.script.run)
 *  - Con ?action   -> responde JSON/JSONP (API para el sitio de Vercel)     */
/* ------------------------------------------------------------------ */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) return handleApi_(e);
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Nuestra Boda · Organizador')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* Permite también POST (por si en el futuro se usa fetch) */
function doPost(e) { return handleApi_(e); }

/* ------------------------------------------------------------------ */
/*  API  ·  puente entre el sitio de Vercel y la base (Sheets/Calendar) */
/* ------------------------------------------------------------------ */
var WRITE_ACTIONS_ = { upsert: 1, remove: 1, saveConfig: 1, syncCita: 1 };

function handleApi_(e) {
  ensureSheets_();
  var p = (e && e.parameter) || {};
  var action = p.action;
  var cb = p.callback;
  var payload = [];
  try { payload = p.payload ? JSON.parse(p.payload) : []; } catch (err) { payload = []; }

  var out;
  try {
    if (WRITE_ACTIONS_[action] && TOKEN && p.token !== TOKEN) {
      throw new Error('No autorizado (token inválido)');
    }
    var data;
    switch (action) {
      case 'getBootstrap': data = getBootstrap(); break;
      case 'saveConfig':   data = saveConfig(payload[0]); break;
      case 'upsert':       data = upsert(payload[0], payload[1]); break;
      case 'remove':       data = remove(payload[0], payload[1]); break;
      case 'syncCita':     data = syncCita(payload[0]); break;
      default: throw new Error('Acción desconocida: ' + action);
    }
    out = { ok: true, data: data };
  } catch (err) {
    out = { error: String((err && err.message) || err) };
  }

  var body = JSON.stringify(out);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/*  SETUP                                                               */
/* ------------------------------------------------------------------ */
function setup() {
  ensureSheets_(true);
  SpreadsheetApp.getActive().toast('Pestañas creadas. Ahora implementá la app web.', 'Listo', 6);
}

function ensureSheets_(seed) {
  const ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]])
        .setFontWeight('bold').setBackground('#47512f').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      if (name === 'Config') seedConfig_(sh);
      if (seed) seedSample_(name, sh);
    }
  });
  // borra la pestaña por defecto "Hoja 1"/"Sheet1" si quedó vacía
  ['Hoja 1', 'Hoja1', 'Sheet1'].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });
}

function seedConfig_(sh) {
  const rows = Object.keys(CONFIG_DEFAULT).map(function (k) { return [k, CONFIG_DEFAULT[k]]; });
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

function seedSample_(name, sh) {
  const S = SAMPLE_[name];
  if (!S || !S.length) return;
  const headers = TABS[name];
  const values = S.map(function (o) { return headers.map(function (h) { return o[h] != null ? o[h] : ''; }); });
  sh.getRange(2, 1, values.length, headers.length).setValues(values);
}

/* ------------------------------------------------------------------ */
/*  READ                                                               */
/* ------------------------------------------------------------------ */
function getBootstrap() {
  ensureSheets_();
  return {
    config:      readConfig_(),
    presupuesto: readTab_('Presupuesto'),
    proveedores: readTab_('Proveedores'),
    invitados:   readTab_('Invitados'),
    cronograma:  readTab_('Cronograma'),
    citas:       readTab_('Citas')
  };
}

function readTab_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(function (r) { return r.join('') !== ''; })
    .map(function (r) {
      const o = {};
      headers.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

function readConfig_() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const out = Object.assign({}, CONFIG_DEFAULT);
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0]) out[r[0]] = r[1];
  });
  out.presupuestoTotal = Number(out.presupuestoTotal) || 0;
  out.fecha = formatDateKey_(out.fecha);
  return out;
}

function formatDateKey_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v ? String(v).slice(0, 10) : '';
}

/* ------------------------------------------------------------------ */
/*  WRITE  (genérico para todas las entidades)                         */
/* ------------------------------------------------------------------ */
function upsert(tabKey, obj) {
  const name = TABMAP[tabKey];
  if (!name) throw new Error('Entidad desconocida: ' + tabKey);
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const headers = TABS[name];

  if (!obj.id) obj.id = Utilities.getUuid();

  const values = sh.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(obj.id)) { rowIndex = i + 1; break; }
  }
  const row = headers.map(function (h) { return obj[h] != null ? obj[h] : ''; });

  if (rowIndex > 0) {
    sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }
  return obj;
}

function remove(tabKey, id) {
  const name = TABMAP[tabKey];
  if (!name) throw new Error('Entidad desconocida: ' + tabKey);
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      // si es una cita con evento, borra el evento del calendario
      if (name === 'Citas') {
        const eventId = values[i][headers_(name).indexOf('eventId')];
        if (eventId) { try { deleteCalendarEvent_(eventId); } catch (e) {} }
      }
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
function headers_(name) { return TABS[name]; }

function saveConfig(cfg) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Config');
  const existing = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r, i) {
    if (r[0]) existing[r[0]] = i + 2; // fila real
  });
  Object.keys(cfg).forEach(function (k) {
    if (existing[k]) sh.getRange(existing[k], 2).setValue(cfg[k]);
    else sh.appendRow([k, cfg[k]]);
  });
  return readConfig_();
}

/* ------------------------------------------------------------------ */
/*  GOOGLE CALENDAR                                                     */
/* ------------------------------------------------------------------ */
function syncCita(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName('Citas');
  const headers = TABS.Citas;
  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      const cita = {};
      headers.forEach(function (h, j) { cita[h] = values[i][j]; });

      const start = parseDate_(cita.inicio);
      let end = cita.fin ? parseDate_(cita.fin) : new Date(start.getTime() + 3600000);
      if (end <= start) end = new Date(start.getTime() + 3600000);

      const cal = CalendarApp.getDefaultCalendar();
      const title = '💍 ' + (cita.proveedor || 'Proveedor') + ' — ' + (cita.titulo || 'Cita');
      const desc = 'Estado: ' + (cita.estado || '') + '\n' + (cita.notas || '') +
                   '\n\n(Creado desde el Organizador de la boda)';

      let ev = null;
      if (cita.eventId) {
        try { ev = cal.getEventById(cita.eventId); } catch (e) { ev = null; }
      }
      if (ev) {
        ev.setTitle(title);
        ev.setTime(start, end);
        ev.setDescription(desc);
      } else {
        ev = cal.createEvent(title, start, end, { description: desc });
        try { ev.addPopupReminder(60 * 24); } catch (e) {}
      }

      const eventId = ev.getId();
      sh.getRange(i + 1, headers.indexOf('eventId') + 1).setValue(eventId);
      cita.eventId = eventId;
      return cita;
    }
  }
  throw new Error('Cita no encontrada: ' + id);
}

function deleteCalendarEvent_(eventId) {
  const ev = CalendarApp.getDefaultCalendar().getEventById(eventId);
  if (ev) ev.deleteEvent();
}

function parseDate_(v) {
  if (v instanceof Date) return v;
  // acepta "yyyy-MM-ddTHH:mm"
  const s = String(v).replace(' ', 'T');
  const d = new Date(s);
  if (!isNaN(d)) return d;
  return new Date();
}

/* ------------------------------------------------------------------ */
/*  DATOS DE EJEMPLO (solo si ejecutás setup)                          */
/* ------------------------------------------------------------------ */
const SAMPLE_ = {
  Presupuesto:[
    {
      id:"p1",
      categoria:"Wedding Planner",
      estimado:0,
      notas:"No estaba en el presupuesto de la planner"
    },
    {
      id:"p2",
      categoria:"Salón / Local",
      estimado:22100,
      notas:"ANCPP · alquiler US$6,500 @ S/3.40"
    },
    {
      id:"p3",
      categoria:"Catering y bebidas",
      estimado:77000,
      notas:"Mónica Tremolada · costo x pax"
    },
    {
      id:"p4",
      categoria:"Licores y barra",
      estimado:13600,
      notas:"Licores 12,000 + insumos barra 1,600"
    },
    {
      id:"p5",
      categoria:"Decoración y flores",
      estimado:27950,
      notas:"Decoración adicional 26,950 + bouquet 1,000"
    },
    {
      id:"p6",
      categoria:"Música y DJ",
      estimado:7000,
      notas:"DJ + sonido y luces · 8h"
    },
    {
      id:"p7",
      categoria:"Fotografía",
      estimado:8840,
      notas:"US$2,600 @ S/3.40 · 12h"
    },
    {
      id:"p8",
      categoria:"Video / Cinematografía",
      estimado:8840,
      notas:"US$2,600 @ S/3.40 · 12h"
    },
    {
      id:"p9",
      categoria:"Vestuario novia",
      estimado:8500,
      notas:"Vestido · Luna Blanco · US$2,500 @ S/3.40"
    },
    {
      id:"p10",
      categoria:"Belleza (novia)",
      estimado:3300,
      notas:"Maquillaje y peinado"
    },
    {
      id:"p11",
      categoria:"Vestuario novio",
      estimado:5500,
      notas:"Traje"
    },
    {
      id:"p12",
      categoria:"Alianzas",
      estimado:0,
      notas:"A cargo de los novios"
    },
    {
      id:"p13",
      categoria:"Papelería e invitaciones",
      estimado:670,
      notas:"Diseño invitaciones 420 + señalética 250"
    },
    {
      id:"p14",
      categoria:"Torta",
      estimado:2000,
      notas:"Costo referencial"
    },
    {
      id:"p15",
      categoria:"Producción y montaje",
      estimado:1800,
      notas:"Grupo electrógeno (montaje no estaba presupuestado)"
    },
    {
      id:"p16",
      categoria:"Show / Hora loca",
      estimado:3800,
      notas:"Hora loca 2,200 + cotillón 1,600"
    },
    {
      id:"p17",
      categoria:"Ceremonia (iglesia)",
      estimado:4500,
      notas:"Flores 2,300 + coro/grupo 2,200"
    },
    {
      id:"p18",
      categoria:"Seguridad",
      estimado:0,
      notas:"Por cotizar"
    }
  ],
  Proveedores:[
    {
      id:"v1",
      nombre:"Wedding Planner",
      categoria:"Wedding Planner",
      estado:"Contratado",
      contacto:"",
      cotizado:10800,
      pagado:5400,
      notas:"Por pagar S/5,400"
    },
    {
      id:"v2",
      nombre:"ANCPP (Local)",
      categoria:"Salón / Local",
      estado:"Contratado",
      contacto:"",
      cotizado:22100,
      pagado:11050,
      notas:"US$6,500 @ S/3.40 · garantía S/1,000 · por pagar US$3,250"
    },
    {
      id:"v3",
      nombre:"Alexis (Save the date)",
      categoria:"Papelería e invitaciones",
      estado:"Contratado",
      contacto:"",
      cotizado:2040,
      pagado:1020,
      notas:"US$600 @ S/3.40"
    },
    {
      id:"v4",
      nombre:"Mata (Maquillaje y peinado)",
      categoria:"Belleza (novia)",
      estado:"Contratado",
      contacto:"",
      cotizado:5000,
      pagado:2500,
      notas:""
    },
    {
      id:"v5",
      nombre:"DJ Andrés Viches",
      categoria:"Música y DJ",
      estado:"Contratado",
      contacto:"",
      cotizado:4500,
      pagado:2250,
      notas:""
    },
    {
      id:"v6",
      nombre:"Luna Blanco (Vestido)",
      categoria:"Vestuario novia",
      estado:"Contratado",
      contacto:"",
      cotizado:9860,
      pagado:5916,
      notas:"US$2,900 @ S/3.40"
    },
    {
      id:"v7",
      nombre:"Frank Andonaire (Fotos)",
      categoria:"Fotografía",
      estado:"Contratado",
      contacto:"",
      cotizado:11424,
      pagado:3264,
      notas:"US$3,360 @ S/3.40 · 12h"
    },
    {
      id:"v8",
      nombre:"Frank Andonaire (Video)",
      categoria:"Video / Cinematografía",
      estado:"Contratado",
      contacto:"",
      cotizado:5440,
      pagado:2176,
      notas:"US$1,600 @ S/3.40 · 12h"
    },
    {
      id:"v9",
      nombre:"Lovepot (Invitaciones)",
      categoria:"Papelería e invitaciones",
      estado:"Contratado",
      contacto:"",
      cotizado:3500,
      pagado:1750,
      notas:""
    },
    {
      id:"v10",
      nombre:"Mónica Tremolada (Catering)",
      categoria:"Catering y bebidas",
      estado:"Contratado",
      contacto:"",
      cotizado:83000,
      pagado:6000,
      notas:"Por pagar S/77,000"
    },
    {
      id:"v11",
      nombre:"Estructuras (Montaje)",
      categoria:"Producción y montaje",
      estado:"Contratado",
      contacto:"",
      cotizado:4500,
      pagado:0,
      notas:"Por pagar S/4,500"
    },
    {
      id:"v12",
      nombre:"Grupo Electrógeno",
      categoria:"Producción y montaje",
      estado:"Contratado",
      contacto:"",
      cotizado:1800,
      pagado:0,
      notas:"Por pagar S/1,800"
    },
    {
      id:"v13",
      nombre:"Alianzas (por decidir)",
      categoria:"Alianzas",
      estado:"Cotizado",
      contacto:"",
      cotizado:0,
      pagado:0,
      notas:"Opciones: Casa Banchero, Aldo"
    },
    {
      id:"v14",
      nombre:"Terno novio (por decidir)",
      categoria:"Vestuario novio",
      estado:"Cotizado",
      contacto:"",
      cotizado:0,
      pagado:0,
      notas:"Opciones: Yorgo Stratouris, Casa España, Firenze"
    },
    {
      id:"v15",
      nombre:"Pancho Security",
      categoria:"Seguridad",
      estado:"Cotizado",
      contacto:"",
      cotizado:0,
      pagado:0,
      notas:"Por cotizar"
    },
    {
      id:"v16",
      nombre:"Hora Loca (por decidir)",
      categoria:"Show / Hora loca",
      estado:"Cotizado",
      contacto:"",
      cotizado:0,
      pagado:0,
      notas:"Incluye cotillón · evaluar 'Naricitas'"
    }
  ],
  Invitados:[
    {
      id:"g1",
      invitacion:"Dante",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:5,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g2",
      invitacion:"Delia",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g3",
      invitacion:"Daniela",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g4",
      invitacion:"Brenda",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g5",
      invitacion:"Hugo",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g6",
      invitacion:"Edith",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g7",
      invitacion:"Rau",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g8",
      invitacion:"Chcha",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g9",
      invitacion:"Julio",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g10",
      invitacion:"Anibal",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:3,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g11",
      invitacion:"Joao",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g12",
      invitacion:"Bett B + Guillermo",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g13",
      invitacion:"Mery",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g14",
      invitacion:"Liss",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g15",
      invitacion:"Diego",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g16",
      invitacion:"Paty",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g17",
      invitacion:"Angela",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g18",
      invitacion:"Andrea",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g19",
      invitacion:"Cinthia",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g20",
      invitacion:"MEmo",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g21",
      invitacion:"Liss",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g22",
      invitacion:"Linzay",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g23",
      invitacion:"Gianella",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g24",
      invitacion:"Amadeo",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g25",
      invitacion:"Enrique",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:3,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g26",
      invitacion:"Step",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g27",
      invitacion:"Raul Pal",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g28",
      invitacion:"Jim",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g29",
      invitacion:"Kike",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g30",
      invitacion:"Nicole",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g31",
      invitacion:"Serapio",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g32",
      invitacion:"Sandy",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g33",
      invitacion:"Mama Delia",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g34",
      invitacion:"Luis S.",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g35",
      invitacion:"Michael",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g36",
      invitacion:"Hugo Ca",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g37",
      invitacion:"Tio Fernando",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g38",
      invitacion:"Diego Cano",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g39",
      invitacion:"Gloria Ron",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g40",
      invitacion:"Miguel Angel",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g41",
      invitacion:"T'ia Zoila",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g42",
      invitacion:"Miguel P",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g43",
      invitacion:"Teresa C",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g44",
      invitacion:"Karina",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g45",
      invitacion:"Mama + Papa",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g46",
      invitacion:"Denisse",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g47",
      invitacion:"Ceecilia",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g48",
      invitacion:"Leo",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g49",
      invitacion:"Hitomi",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g50",
      invitacion:"Yazzmin",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g51",
      invitacion:"Angela Flores",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g52",
      invitacion:"Christopher V",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g53",
      invitacion:"Jean Marco",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g54",
      invitacion:"Lorena P",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g55",
      invitacion:"Rafo",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g56",
      invitacion:"Padrinos",
      grupo:"Novia · Fam. Mamá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g57",
      invitacion:"MArisol",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g58",
      invitacion:"Massi",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g59",
      invitacion:"Efrain",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g60",
      invitacion:"Alessa C",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g61",
      invitacion:"Pepe V.",
      grupo:"Novia · Fam. Papá",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g62",
      invitacion:"Pame R.",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g63",
      invitacion:"Patty H.",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g64",
      invitacion:"Mauricio",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g65",
      invitacion:"Ricardo",
      grupo:"Novia · Amigos",
      lado:"Novia",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g66",
      invitacion:"Take + Thiago",
      grupo:"Novio · Fam. Papá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g67",
      invitacion:"Mila",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:3,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g68",
      invitacion:"Alonso",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g69",
      invitacion:"DK",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g70",
      invitacion:"Ana + Germán",
      grupo:"Novio · Fam. Papá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g71",
      invitacion:"Rodrigo",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g72",
      invitacion:"GC",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g73",
      invitacion:"Juana",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g74",
      invitacion:"Chio",
      grupo:"Novio · Fam. Papá",
      lado:"Novio",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g75",
      invitacion:"Chaqui",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:4,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g76",
      invitacion:"Diego",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g77",
      invitacion:"Jack",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g78",
      invitacion:"Calin",
      grupo:"Novio · Fam. Papá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g79",
      invitacion:"Jorge",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g80",
      invitacion:"Jessica",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g81",
      invitacion:"Rafa",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g82",
      invitacion:"Hugo",
      grupo:"Novio · Fam. Papá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g83",
      invitacion:"Teresa",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g84",
      invitacion:"Chevo",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g85",
      invitacion:"Marlio",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g86",
      invitacion:"Patty",
      grupo:"Novio · Fam. Papá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g87",
      invitacion:"Alfredo",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g88",
      invitacion:"Nicole",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g89",
      invitacion:"Hiro",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g90",
      invitacion:"Camila",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g91",
      invitacion:"Miguel",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g92",
      invitacion:"Diego",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g93",
      invitacion:"Kaori",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g94",
      invitacion:"Bruce",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g95",
      invitacion:"Yoshio",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g96",
      invitacion:"Lucho",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g97",
      invitacion:"Juan Diego",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g98",
      invitacion:"Chochi",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:1,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g99",
      invitacion:"Oneto",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g100",
      invitacion:"Juan",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g101",
      invitacion:"Pedro Hig",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g102",
      invitacion:"Adrian",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g103",
      invitacion:"Nina",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g104",
      invitacion:"Emiko",
      grupo:"Novio · Fam. Mamá",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g105",
      invitacion:"Fasho",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g106",
      invitacion:"Salvador",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g107",
      invitacion:"Majo",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g108",
      invitacion:"Santos",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g109",
      invitacion:"Mirco",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g110",
      invitacion:"Vega",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g111",
      invitacion:"MAria Alejandr",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g112",
      invitacion:"Israel",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    },
    {
      id:"g113",
      invitacion:"Hide",
      grupo:"Novio · Amigos",
      lado:"Novio",
      personas:2,
      mesa:"En espera",
      confirmacion:"En espera",
      restricciones:"",
      notas:""
    }
  ],
  Cronograma:[
    {
      id:"m1",
      bloque:"Mañana",
      hora:"08:00",
      duracion:"3h",
      titulo:"Peinado y maquillaje — Novia",
      lugar:"Suite nupcial",
      responsable:"Mata",
      notas:""
    },
    {
      id:"m2",
      bloque:"Mañana",
      hora:"11:00",
      duracion:"1h 30min",
      titulo:"Peinado y maquillaje — Damas",
      lugar:"Suite nupcial",
      responsable:"Mata",
      notas:""
    },
    {
      id:"t1",
      bloque:"Tarde",
      hora:"12:30",
      duracion:"",
      titulo:"Llegada de fotógrafo y video",
      lugar:"",
      responsable:"Frank Andonaire",
      notas:""
    },
    {
      id:"t2",
      bloque:"Tarde",
      hora:"13:00",
      duracion:"1h",
      titulo:"Sesión de fotos previa",
      lugar:"Jardín",
      responsable:"Frank Andonaire",
      notas:""
    },
    {
      id:"t3",
      bloque:"Tarde",
      hora:"15:00",
      duracion:"30 min",
      titulo:"Llegada de invitados",
      lugar:"Local ANCPP",
      responsable:"",
      notas:"Música: DJ"
    },
    {
      id:"t4",
      bloque:"Tarde",
      hora:"15:30",
      duracion:"1h",
      titulo:"Ceremonia",
      lugar:"",
      responsable:"",
      notas:"Coro / grupo musical"
    },
    {
      id:"n1",
      bloque:"Noche",
      hora:"19:00",
      duracion:"",
      titulo:"Recepción y cena",
      lugar:"Local ANCPP",
      responsable:"Mónica Tremolada",
      notas:""
    },
    {
      id:"n2",
      bloque:"Noche",
      hora:"21:00",
      duracion:"",
      titulo:"Fiesta",
      lugar:"Pista",
      responsable:"DJ Andrés Viches",
      notas:""
    },
    {
      id:"n3",
      bloque:"Noche",
      hora:"22:30",
      duracion:"",
      titulo:"Hora loca",
      lugar:"Pista",
      responsable:"",
      notas:"Cotillón"
    }
  ],
  Citas:[
    {
      id:"a1",
      proveedor:"Wedding Planner",
      titulo:"Reunión de avance",
      inicio:"2026-10-20T17:00",
      fin:"2026-10-20T18:00",
      estado:"Contratado",
      notas:"(sugerida — ajustá la fecha)",
      eventId:""
    },
    {
      id:"a2",
      proveedor:"Luna Blanco (Vestido)",
      titulo:"Prueba de vestido",
      inicio:"2026-10-28T11:00",
      fin:"2026-10-28T12:30",
      estado:"Contratado",
      notas:"(sugerida — ajustá la fecha)",
      eventId:""
    },
    {
      id:"a3",
      proveedor:"Mónica Tremolada (Catering)",
      titulo:"Degustación de menú",
      inicio:"2026-11-08T13:00",
      fin:"2026-11-08T15:00",
      estado:"Contratado",
      notas:"(sugerida — ajustá la fecha)",
      eventId:""
    },
    {
      id:"a4",
      proveedor:"Mata (Maquillaje y peinado)",
      titulo:"Prueba de maquillaje",
      inicio:"2026-11-22T10:00",
      fin:"2026-11-22T11:30",
      estado:"Contratado",
      notas:"(sugerida — ajustá la fecha)",
      eventId:""
    },
    {
      id:"a5",
      proveedor:"ANCPP (Local)",
      titulo:"Visita técnica al local",
      inicio:"2026-11-29T10:00",
      fin:"2026-11-29T11:30",
      estado:"Contratado",
      notas:"(sugerida — ajustá la fecha)",
      eventId:""
    },
    {
      id:"a6",
      proveedor:"Wedding Planner",
      titulo:"Reunión final",
      inicio:"2026-12-06T18:00",
      fin:"2026-12-06T19:00",
      estado:"Contratado",
      notas:"(sugerida — ajustá la fecha)",
      eventId:""
    }
  ,
    {id:"dc1",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-09-13T12:00",fin:"2026-09-13T13:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc2",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-09-20T16:00",fin:"2026-09-20T17:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc3",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-09-27T12:00",fin:"2026-09-27T13:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc4",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-10-04T12:00",fin:"2026-10-04T13:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc5",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-10-11T12:00",fin:"2026-10-11T13:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc6",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-10-19T20:30",fin:"2026-10-19T22:00",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc7",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-10-25T12:00",fin:"2026-10-25T13:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc8",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-11-01T16:00",fin:"2026-11-01T17:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc9",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-11-08T12:00",fin:"2026-11-08T13:30",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""},
    {id:"dc10",proveedor:"Bodanza",titulo:"Clase de baile",inicio:"2026-11-16T20:30",fin:"2026-11-16T22:00",estado:"Programado",notas:"Claudia y Jorge · Bodanza",eventId:""}
  ]
};
