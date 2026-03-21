/**
 * ═══════════════════════════════════════════════════════════════════
 * PROFAR CDP — Google Apps Script · Auto-Update Autónomo
 * ═══════════════════════════════════════════════════════════════════
 *
 * SETUP (solo una vez, ~5 minutos):
 * ─────────────────────────────────
 * 1. Reemplaza Code.gs con este archivo
 * 2. Ejecutar → _init()  (autoriza permisos + crea hoja + carga datos)
 * 3. Implementar → Nueva implementación
 *       Tipo: Aplicación web
 *       Ejecutar como: Yo
 *       Quién tiene acceso: Todos (anónimo)
 *    → Copia la URL /exec
 * 4. En profar-cdp.html: const GAS_ENDPOINT = 'PEGAR_URL_AQUI'
 * 5. Ejecutar → configurarTriggers()
 *
 * LISTO. El script se actualiza solo 3x/día desde ahí en adelante.
 * ═══════════════════════════════════════════════════════════════════
 */

// ── Entry point para auto-selección ───────────────────────────────
function _init() { inicializar(); }

// ── Configuración ──────────────────────────────────────────────────
var BASE_URL = 'https://profar.cl/dataexport/download/index';
var TOKEN    = 'a40dkakmopd01mjcrf632syd67ursq5k';
var IVA      = 1.19;

var SH = {
  ORDERS : 'Raw_Orders',
  SUBS   : 'Raw_Subscriptions',
  CUSTS  : 'Raw_Customers',
  PRODS  : 'Raw_Products',
  KPI    : 'KPI_Cache'
};

var GRUPOS = {
  '0':'General','1':'NOT LOGGED IN','2':'Wholesale',
  '3':'Club Beneficios','4':'Empleados PROFAR',
  '6':'Mutual','7':'Club Polo','8':'SubsecGue',
  '9':'MED','10':'GO Integro','11':'SUBTEL'
};

var DIRECTOS_ARR = ['0','1','3','4'];
var MESES_CORTOS = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

// ── getSpreadsheet: recupera o null si no inicializado ─────────────
function getSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch(e) { return null; }
}

// ── Entry point ────────────────────────────────────────────────────
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getData';
  var result;
  try {
    if      (action === 'getData')      { result = getCachedData(); }
    else if (action === 'getHistorico') { result = getHistoricoData(); }
    else if (action === 'forceUpdate')  { result = runUpdate(); }
    else if (action === 'ping')         { result = { ok: true, ts: new Date().toISOString() }; }
    else                                { result = { error: 'Accion desconocida: ' + action }; }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── getCachedData ──────────────────────────────────────────────────
function getCachedData() {
  var ss  = getSpreadsheet();
  if (!ss) return { error: 'Sin hoja de calculo. Ejecuta _init() primero.' };
  var sh  = ss.getSheetByName(SH.KPI);
  if (!sh || sh.getLastRow() < 1) return { error: 'Sin datos. Ejecuta runUpdate() primero.' };
  var json = sh.getRange(1,1).getValue();
  if (!json) return { error: 'Cache vacio.' };
  var data = JSON.parse(json);
  data._fromCache = true;
  data._cacheAge  = Math.round((Date.now() - data._ts) / 1000);
  return data;
}

// ── runUpdate: orquesta todo ───────────────────────────────────────
function runUpdate() {
  Logger.log('=== runUpdate START ' + new Date().toISOString() + ' ===');
  var ss = getSpreadsheet();
  if (!ss) throw new Error('Sin hoja de calculo. Ejecuta _init() primero.');

  var orders = fetchCSV('orders');
  var subs   = fetchCSV('subscriptions');
  var custs  = fetchCSV('customers');
  var prods  = fetchCSV('products');

  Logger.log('Descargados — orders:'+orders.length+' subs:'+subs.length+' custs:'+custs.length+' prods:'+prods.length);

  upsertOrders(ss, orders);
  replaceSheet(ss, SH.SUBS,  subs);
  replaceSheet(ss, SH.CUSTS, custs);
  replaceSheet(ss, SH.PRODS, prods);

  var kpis = buildKPIs(ss);
  var kpiSh = getOrCreateSheet(ss, SH.KPI);
  kpiSh.clearContents();
  kpiSh.getRange(1,1).setValue(JSON.stringify(kpis));

  Logger.log('=== runUpdate DONE ===');
  return { ok: true, updatedAt: kpis._updatedAt, ordersTotal: kpis.ordersAcumulados };
}

// ── fetchCSV ───────────────────────────────────────────────────────
function fetchCSV(type) {
  var url  = BASE_URL + '?type=' + type + '&token=' + TOKEN;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode() + ' para ' + type);
  return parseCSV(resp.getContentText());
}

// ── parseCSV ───────────────────────────────────────────────────────
function parseCSV(text) {
  var lines = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
  if (lines.length < 2) return [];
  var headers = splitCSVLine(lines[0]);
  var rows = [];
  for (var i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    var vals = splitCSVLine(lines[i]);
    var obj  = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j].trim()] = vals[j] !== undefined ? vals[j].trim() : '';
    }
    rows.push(obj);
  }
  return rows;
}

function splitCSVLine(line) {
  var result = [], cur = '', inQ = false;
  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (c === '"') {
      if (inQ && line[i+1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (c === ',' && !inQ) {
      result.push(cur); cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur);
  return result;
}

// ── upsertOrders: acumula sin duplicar ────────────────────────────
function upsertOrders(ss, newOrders) {
  var sh = getOrCreateSheet(ss, SH.ORDERS);
  if (newOrders.length === 0) return;

  if (sh.getLastRow() < 2) {
    var headers0 = Object.keys(newOrders[0]);
    var data0 = newOrders.map(function(r){ return headers0.map(function(h){ return r[h]||''; }); });
    sh.getRange(1,1,1,headers0.length).setValues([headers0]);
    sh.getRange(2,1,data0.length,headers0.length).setValues(data0);
    Logger.log('Orders: escritas '+newOrders.length+' filas (primera vez)');
    return;
  }

  var lastRow  = sh.getLastRow();
  var headers  = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var idCol    = headers.indexOf('increment_id');
  if (idCol < 0) { Logger.log('No encontre increment_id'); return; }

  var existingArr = sh.getRange(2, idCol+1, lastRow-1, 1).getValues();
  var existing = {};
  for (var k = 0; k < existingArr.length; k++) { existing[String(existingArr[k][0])] = true; }

  var toAdd = newOrders.filter(function(r){ return !existing[String(r['increment_id'])]; });
  if (toAdd.length === 0) { Logger.log('Orders: sin nuevas ordenes'); return; }

  var rows = toAdd.map(function(r){ return headers.map(function(h){ return r[h]||''; }); });
  sh.getRange(lastRow+1, 1, rows.length, headers.length).setValues(rows);
  Logger.log('Orders: agregadas '+toAdd.length+' nuevas ('+Object.keys(existing).length+' ya existian)');
}

// ── replaceSheet ──────────────────────────────────────────────────
function replaceSheet(ss, name, rows) {
  var sh = getOrCreateSheet(ss, name);
  sh.clearContents();
  if (rows.length === 0) return;
  var headers = Object.keys(rows[0]);
  var data = rows.map(function(r){ return headers.map(function(h){ return r[h]||''; }); });
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (data.length > 0) sh.getRange(2,1,data.length,headers.length).setValues(data);
}

// ── buildKPIs ─────────────────────────────────────────────────────
function buildKPIs(ss) {
  var ordersSh = ss.getSheetByName(SH.ORDERS);
  var subsSh   = ss.getSheetByName(SH.SUBS);
  var prodsSh  = ss.getSheetByName(SH.PRODS);
  var now      = new Date();

  var ordersKPI = processOrders(ordersSh);
  var subsKPI   = processSubs(subsSh);
  var topProds  = getTopProducts(ordersSh);
  var stockKPI  = processStock(prodsSh);

  return {
    _updatedAt:        now.toISOString(),
    _ts:               now.getTime(),
    ordersAcumulados:  ordersKPI.totalOrders,
    resumen: {
      last60d: {
        revenue:        ordersKPI.rev60,
        revenueConIVA:  Math.round(ordersKPI.rev60 * IVA),
        orders:         ordersKPI.orders60,
        clientes:       ordersKPI.clients60,
        ticketProm:     ordersKPI.aov60,
        topDia:         ordersKPI.topDia
      },
      yoy: ordersKPI.yoy
    },
    historico:    ordersKPI.historico,
    monthly:      ordersKPI.monthly,
    segmentos:    ordersKPI.segmentos,
    suscripciones: subsKPI,
    topProductos: topProds,
    stock:        stockKPI
  };
}

// ── processOrders ─────────────────────────────────────────────────
function processOrders(sh) {
  if (!sh || sh.getLastRow() < 2) {
    return { totalOrders:0, rev60:0, orders60:0, clients60:0, aov60:0,
             topDia:{fecha:'—',rev:0}, yoy:null, historico:[], monthly:{}, segmentos:{directos:[],convenios:[]} };
  }

  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(String);
  var idxOf   = function(h){ return headers.indexOf(h); };

  var C = {
    status:  idxOf('status'),
    total:   idxOf('grand_total'),
    date:    idxOf('created_at'),
    rut:     idxOf('customer_rut'),
    group:   idxOf('customer_group_id'),
    itemsSku:idxOf('items_sku')
  };

  var now      = new Date();
  var ms60     = 60 * 24 * 60 * 60 * 1000;
  var cutoff60 = new Date(now.getTime() - ms60);
  var cutoffYOYstart = new Date(now.getTime() - 365*24*60*60*1000 - ms60);
  var cutoffYOYend   = new Date(now.getTime() - 365*24*60*60*1000);

  var monthlyMap = {}, segs60 = {}, clients60 = {}, dayMap60 = {};
  var rev60 = 0, orders60 = 0, revYOY = 0, ordersYOY = 0, totalOrders = 0;

  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var status = String(row[C.status]||'').toLowerCase();
    if (status === 'canceled' || status === 'pending_payment') continue;

    var total = parseFloat(String(row[C.total]).replace(/,/g,'')) || 0;
    if (total <= 0) continue;

    var d = new Date(String(row[C.date]||''));
    if (isNaN(d.getTime())) continue;

    totalOrders++;

    var ym  = d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2);
    if (!monthlyMap[ym]) monthlyMap[ym] = { rev:0, orders:0, clients:{} };
    monthlyMap[ym].rev    += total;
    monthlyMap[ym].orders += 1;
    var rut = String(row[C.rut]||'');
    if (rut) monthlyMap[ym].clients[rut] = true;

    if (d >= cutoff60) {
      rev60 += total; orders60++;
      if (rut) clients60[rut] = true;
      var dk = d.toISOString().slice(0,10);
      dayMap60[dk] = (dayMap60[dk]||0) + total;

      var gid = String(row[C.group]||'0');
      if (!segs60[gid]) segs60[gid] = { rev:0, orders:0, clients:{} };
      segs60[gid].rev    += total;
      segs60[gid].orders += 1;
      if (rut) segs60[gid].clients[rut] = true;
    }

    if (d >= cutoffYOYstart && d < cutoffYOYend) {
      revYOY += total; ordersYOY++;
    }
  }

  // Histórico ordenado
  var sortedYM = Object.keys(monthlyMap).sort();
  var historico = sortedYM.map(function(ym){
    var parts = ym.split('-');
    var y = parseInt(parts[0]), m = parseInt(parts[1]);
    var days = new Date(y, m, 0).getDate();
    return {
      date:    ('0'+days).slice(-2)+'/'+('0'+m).slice(-2)+'/'+String(y).slice(-2),
      ts:      new Date(y, m-1, days).getTime(),
      rev30:   parseFloat((monthlyMap[ym].rev/1e6).toFixed(2)),
      orders30: monthlyMap[ym].orders,
      subs:    Object.keys(monthlyMap[ym].clients).length
    };
  });

  // Últimos 6 meses
  var last6 = sortedYM.slice(-6);
  var monthly = {
    labels:  last6.map(function(ym){ return MESES_CORTOS[parseInt(ym.split('-')[1])-1]; }),
    revenue: last6.map(function(ym){ return Math.round(monthlyMap[ym].rev); }),
    orders:  last6.map(function(ym){ return monthlyMap[ym].orders; }),
    clients: last6.map(function(ym){ return Object.keys(monthlyMap[ym].clients).length; })
  };

  // Top día
  var topDiaEntry = ['—', 0];
  Object.keys(dayMap60).forEach(function(dk){
    if (dayMap60[dk] > topDiaEntry[1]) topDiaEntry = [dk, dayMap60[dk]];
  });

  // Segmentos
  var directosRes = [], conveniosRes = [];
  Object.keys(segs60).forEach(function(gid){
    var v = segs60[gid];
    var item = {
      grupo:   GRUPOS[gid] || 'Grupo '+gid,
      gid:     gid,
      rev:     Math.round(v.rev),
      orders:  v.orders,
      clients: Object.keys(v.clients).length
    };
    if (DIRECTOS_ARR.indexOf(gid) >= 0) directosRes.push(item);
    else conveniosRes.push(item);
  });
  directosRes.sort(function(a,b){ return b.rev-a.rev; });
  conveniosRes.sort(function(a,b){ return b.rev-a.rev; });

  var yoy = revYOY > 0 ? parseFloat(((rev60-revYOY)/revYOY*100).toFixed(1)) : null;

  return {
    totalOrders: totalOrders,
    rev60:    Math.round(rev60),
    orders60: orders60,
    clients60: Object.keys(clients60).length,
    aov60:    orders60 > 0 ? Math.round(rev60/orders60) : 0,
    topDia:   { fecha: topDiaEntry[0], rev: Math.round(topDiaEntry[1]) },
    yoy:      yoy,
    historico: historico,
    monthly:  monthly,
    segmentos: { directos: directosRes, convenios: conveniosRes }
  };
}

// ── processSubs ────────────────────────────────────────────────────
function processSubs(sh) {
  if (!sh || sh.getLastRow() < 2) return { active:0, paused:0, canceled:0, total:0, clientes:0, proximasSemana:0 };

  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(String);
  var statusC = headers.indexOf('status');
  var rutC    = headers.indexOf('customer_rut');
  var freqC   = headers.indexOf('frequency_unit');
  var nextC   = headers.indexOf('next_run');

  var counts  = { active:0, paused:0, canceled:0, total:0 };
  var freqs   = {}, ruts = {}, prox = 0;
  var now     = new Date();
  var d7      = new Date(now.getTime() + 7*24*60*60*1000);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var st  = String(row[statusC]||'').toLowerCase();
    counts.total++;
    if (st === 'active') counts.active++;
    else if (st === 'paused') counts.paused++;
    else counts.canceled++;

    var rut = String(row[rutC]||'');
    if (rut) ruts[rut] = true;

    var freq = String(row[freqC]||'');
    if (freq) freqs[freq] = (freqs[freq]||0) + 1;

    if (st === 'active' && nextC >= 0) {
      var nr = new Date(String(row[nextC]||''));
      if (!isNaN(nr.getTime()) && nr >= now && nr <= d7) prox++;
    }
  }

  return {
    active:   counts.active,
    paused:   counts.paused,
    canceled: counts.canceled,
    total:    counts.total,
    clientes: Object.keys(ruts).length,
    proximasSemana: prox,
    frecuencias: freqs
  };
}

// ── getTopProducts ─────────────────────────────────────────────────
function getTopProducts(sh) {
  if (!sh || sh.getLastRow() < 2) return [];

  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(String);
  var dateC   = headers.indexOf('created_at');
  var statusC = headers.indexOf('status');
  var totalC  = headers.indexOf('grand_total');
  var skuC    = headers.indexOf('items_sku');
  var nameC   = headers.indexOf('items_name');

  var now     = new Date();
  var cutoff  = new Date(now.getTime() - 60*24*60*60*1000);
  var prodMap = {};

  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var status = String(row[statusC]||'').toLowerCase();
    if (status === 'canceled') continue;
    var d = new Date(String(row[dateC]||''));
    if (isNaN(d.getTime()) || d < cutoff) continue;

    var skuRaw  = String(row[skuC]||'');
    var nameRaw = nameC >= 0 ? String(row[nameC]||'') : '';
    // Limpiar comillas embebidas en SKUs
    var skus  = skuRaw.split(',').map(function(s){ return s.replace(/"/g,'').trim(); }).filter(Boolean);
    var names = nameRaw.split(',').map(function(s){ return s.replace(/"/g,'').trim(); });
    var total = parseFloat(String(row[totalC]).replace(/,/g,'')) || 0;

    skus.forEach(function(sku, idx){
      var nombre = (names[idx] && names[idx].length > 2) ? names[idx] : sku;
      if (!prodMap[sku]) prodMap[sku] = { sku:sku, nombre:nombre, qty:0, rev:0 };
      else if (nombre !== sku && prodMap[sku].nombre === sku) prodMap[sku].nombre = nombre;
      prodMap[sku].qty++;
      prodMap[sku].rev += total / skus.length;
    });
  }

  return Object.values(prodMap)
    .sort(function(a,b){ return b.rev-a.rev; })
    .slice(0,20)
    .map(function(p){ return { sku:p.sku, nombre:p.nombre, qty:p.qty, rev:Math.round(p.rev) }; });
}

// ── processStock ──────────────────────────────────────────────────
function processStock(sh) {
  if (!sh || sh.getLastRow() < 2) return { conStock:0, sinStock:0, total:0, pctDisponible:0 };

  var data    = sh.getDataRange().getValues();
  var headers = data[0].map(String);
  var stockC  = headers.indexOf('stock_qty');
  var statusC = headers.indexOf('status');
  var conStock = 0, sinStock = 0, total = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // En Magento: status=2 o 'disabled' = desactivado. Todo lo demás se incluye.
    var st  = String(row[statusC]||'1').toLowerCase();
    if (st === '2' || st === 'disabled') continue;
    total++;
    if ((parseFloat(row[stockC])||0) > 0) conStock++; else sinStock++;
  }

  return {
    conStock:  conStock,
    sinStock:  sinStock,
    total:     total,
    pctDisponible: total > 0 ? parseFloat((conStock/total*100).toFixed(1)) : 0
  };
}

// ── getHistoricoData ──────────────────────────────────────────────
function getHistoricoData() {
  var data = getCachedData();
  if (data.error) return data;
  return { snaps: data.historico||[], count: (data.historico||[]).length };
}

// ── getOrCreateSheet ──────────────────────────────────────────────
function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── inicializar ───────────────────────────────────────────────────
function inicializar() {
  var props = PropertiesService.getScriptProperties();
  var existingId = props.getProperty('SHEET_ID');

  var ss;
  if (existingId) {
    try {
      ss = SpreadsheetApp.openById(existingId);
      Logger.log('Hoja existente encontrada: ' + existingId);
    } catch(e) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('Profar CDP — Data');
    props.setProperty('SHEET_ID', ss.getId());
    Logger.log('Nueva hoja creada: ' + ss.getId() + ' — URL: ' + ss.getUrl());
  }

  try { ss.rename('Profar CDP — Data'); } catch(e) {}
  Object.keys(SH).forEach(function(k){ getOrCreateSheet(ss, SH[k]); });

  Logger.log('Hojas creadas. Corriendo primera actualizacion...');
  var result = runUpdate();
  Logger.log('Inicializacion completa: ' + JSON.stringify(result));
  Logger.log('URL de la hoja: ' + ss.getUrl());
  Logger.log('ID de la hoja: ' + ss.getId());
}

// ── configurarTriggers ────────────────────────────────────────────
function configurarTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runUpdate').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('runUpdate').timeBased().atHour(13).everyDays(1).create();
  ScriptApp.newTrigger('runUpdate').timeBased().atHour(18).everyDays(1).create();
  Logger.log('Triggers configurados: 08:00 / 13:00 / 18:00');
}
