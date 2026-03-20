/**
 * ═══════════════════════════════════════════════════════════════════
 * PROFAR CDP — Google Apps Script · Endpoint de datos
 * ═══════════════════════════════════════════════════════════════════
 *
 * INSTRUCCIONES DE DEPLOY:
 * 1. Abre https://script.google.com → Nuevo proyecto
 * 2. Pega este código reemplazando el contenido de Code.gs
 * 3. Sube también las otras pestañas (Config.gs, DashboardUpdater.gs)
 * 4. Menú: Implementar → Nueva implementación
 *    - Tipo: Aplicación web
 *    - Ejecutar como: YO (tu cuenta Google)
 *    - Quién tiene acceso: Todos (anónimo)  ← para que el dashboard pueda llamarlo
 * 5. Copia la URL del endpoint (termina en /exec)
 * 6. En profar-cdp.html, actualiza: const GAS_ENDPOINT = 'TU_URL_AQUI';
 *
 * FUNCIONAMIENTO:
 * - Dashboard llama GET /exec?action=getData → recibe JSON con todos los KPIs
 * - Trigger diario actualiza la Hoja desde Power BI exports (si están en Drive)
 * - Trigger también corre al abrir el archivo o manualmente
 * ═══════════════════════════════════════════════════════════════════
 */

// ── Configuración ──────────────────────────────────────────────────
const CONFIG = {
  // ID de la Google Sheet que tiene los datos (crea una Sheet y pega el ID de la URL)
  SHEET_ID: 'REEMPLAZA_CON_ID_DE_TU_SHEET',
  // Nombres de las hojas dentro del archivo
  SHEET_RESUMEN: 'Resumen Mensual',
  SHEET_DETALLE: 'Detalle Documentos',
  SHEET_CACHE:   'Cache API',
  // IVA Chile
  IVA: 1.19,
  // Cache en segundos (4 horas)
  CACHE_TTL: 14400
};

// ── Entry point principal ───────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action ? e.parameter.action : 'getData';
  
  let result;
  try {
    if (action === 'getData') {
      result = getData();
    } else if (action === 'getHistorico') {
      result = getHistorico();
    } else if (action === 'getTopProducts') {
      result = getTopProducts();
    } else if (action === 'ping') {
      result = { ok: true, ts: new Date().toISOString(), version: '1.0.0' };
    } else {
      result = { error: 'Acción desconocida: ' + action };
    }
  } catch(err) {
    result = { error: err.message, stack: err.stack };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── getData: devuelve todos los KPIs para el dashboard ─────────────
function getData() {
  // Intentar desde cache primero
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dashboard_data');
  if (cached) {
    const data = JSON.parse(cached);
    data._fromCache = true;
    data._cacheAge = Math.round((Date.now() - data._generatedAt) / 1000);
    return data;
  }
  
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const resumenSheet = ss.getSheetByName(CONFIG.SHEET_RESUMEN);
  const detalleSheet = ss.getSheetByName(CONFIG.SHEET_DETALLE);
  
  if (!resumenSheet) throw new Error('Hoja "' + CONFIG.SHEET_RESUMEN + '" no encontrada');
  
  const data = buildDashboardData(resumenSheet, detalleSheet);
  data._generatedAt = Date.now();
  data._fromCache = false;
  
  // Guardar en cache
  cache.put('dashboard_data', JSON.stringify(data), CONFIG.CACHE_TTL);
  
  return data;
}

// ── buildDashboardData: construye el objeto de datos ──────────────
function buildDashboardData(resumenSheet, detalleSheet) {
  const resumenData = resumenSheet.getDataRange().getValues();
  const headers = resumenData[0].map(h => String(h).trim().toUpperCase());
  
  // Índices de columnas (busca por nombre, flexible ante cambios de orden)
  const colAño   = headers.indexOf('AÑO') !== -1 ? headers.indexOf('AÑO') : headers.findIndex(h => h.includes('AO') || h.includes('AÑO'));
  const colMes   = headers.indexOf('MES');
  const colVenta = headers.findIndex(h => h.includes('VENTA') && !h.includes('UNI'));
  const colMarg  = headers.findIndex(h => h.includes('MARGEN') && h.includes('%'));
  const colUni   = headers.findIndex(h => h.includes('UNI') || h.includes('UNID'));
  const colRot   = headers.findIndex(h => h.includes('ROT'));
  
  // Construir arrays por año/mes
  const byYear = {};
  for (let i = 1; i < resumenData.length; i++) {
    const row = resumenData[i];
    const año  = parseInt(row[colAño]);
    const mes  = String(row[colMes] || '').trim().toUpperCase();
    if (!año || !mes || isNaN(año)) continue;
    
    if (!byYear[año]) byYear[año] = {};
    byYear[año][mes] = {
      venta:  parseFloat(row[colVenta]) || 0,
      margen: parseFloat(row[colMarg])  || 0,
      uni:    parseInt(row[colUni])     || 0,
      rot:    parseFloat(row[colRot])   || 0
    };
  }
  
  const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  
  // Construir D.ecom equivalente
  const buildArray = (año, campo) => MESES.map(m => byYear[año] && byYear[año][m] ? byYear[año][m][campo] : null);
  
  const v24 = buildArray(2024, 'venta');
  const v25 = buildArray(2025, 'venta');
  const v26 = buildArray(2026, 'venta');
  const mg24 = buildArray(2024, 'margen');
  const mg25 = buildArray(2025, 'margen');
  const mg26 = buildArray(2026, 'margen');
  const u24 = buildArray(2024, 'uni');
  const u25 = buildArray(2025, 'uni');
  const u26 = buildArray(2026, 'uni');
  
  // Calcular KPIs YTD 2026
  const ytd26 = v26.filter(v => v !== null && v > 0);
  const ytd25same = v25.slice(0, ytd26.length).filter(v => v !== null);
  const sumaYTD26 = ytd26.reduce((a, b) => a + b, 0);
  const sumaYTD25 = ytd25same.reduce((a, b) => a + b, 0);
  const yoyYTD = sumaYTD25 > 0 ? ((sumaYTD26 - sumaYTD25) / sumaYTD25 * 100) : null;
  
  // Último mes disponible 2026
  const lastIdx26 = v26.reduce((last, v, i) => (v !== null && v > 0 ? i : last), -1);
  const lastMonth26 = lastIdx26 >= 0 ? MESES[lastIdx26] : null;
  const lastRev26 = lastIdx26 >= 0 ? v26[lastIdx26] : null;
  const prevRev26 = lastIdx26 > 0 ? v26[lastIdx26 - 1] : null;
  const momLast = (lastRev26 && prevRev26) ? ((lastRev26 - prevRev26) / prevRev26 * 100) : null;
  
  // Construir MONTHLY_ arrays (últimos 3 meses del año en curso)
  const monthlyIdxs = [];
  for (let i = 0; i < MESES.length; i++) {
    if (v26[i] !== null && v26[i] > 0) monthlyIdxs.push(i);
  }
  
  const monthly = {
    month:   monthlyIdxs.map(i => capitalize(MESES[i])),
    revenue: monthlyIdxs.map(i => v26[i]),
    uni:     monthlyIdxs.map(i => u26[i]),
    margen:  monthlyIdxs.map(i => mg26[i]),
    yoy:     monthlyIdxs.map(i => {
      const cur = v26[i], prev = v25[i];
      return (cur && prev) ? parseFloat(((cur - prev) / prev * 100).toFixed(1)) : null;
    }),
    mom:     monthlyIdxs.map((idx, arrI) => {
      if (arrI === 0) return null;
      const cur = v26[idx], prev = v26[monthlyIdxs[arrI - 1]];
      return (cur && prev) ? parseFloat(((cur - prev) / prev * 100).toFixed(1)) : null;
    })
  };
  
  // Procesar detalle si está disponible
  let detalleStats = null;
  if (detalleSheet) {
    detalleStats = processDetalle(detalleSheet);
  }
  
  return {
    // Estructura compatible con el dashboard
    decom: {
      v24, v25, v26,
      mg24, mg25, mg26,
      cli24: u24, cli25: u25, cli26: u26,  // usamos uni como proxy
    },
    monthly,
    kpis: {
      ytdRev26SinIVA: Math.round(sumaYTD26),
      ytdRev26ConIVA: Math.round(sumaYTD26 * CONFIG.IVA),
      ytdYoY: yoyYTD !== null ? parseFloat(yoyYTD.toFixed(1)) : null,
      lastMonth: lastMonth26,
      lastMonthRev: lastRev26,
      lastMonthRevConIVA: lastRev26 ? Math.round(lastRev26 * CONFIG.IVA) : null,
      lastMonthMoM: momLast !== null ? parseFloat(momLast.toFixed(1)) : null,
      mesesDisponibles: monthlyIdxs.length
    },
    topProducts: detalleStats ? detalleStats.topProducts : null,
    regionStats: detalleStats ? detalleStats.regionStats : null,
    _updatedAt: new Date().toISOString()
  };
}

// ── processDetalle: analiza el archivo de transacciones ────────────
function processDetalle(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  
  const colProd   = headers.findIndex(h => h.includes('PRODUCT') || h === 'PRODUCTO');
  const colSku    = headers.findIndex(h => h === 'SKU' || h.includes('COD'));
  const colVenta  = headers.findIndex(h => h.includes('VENTA') && !h.includes('CANAL'));
  const colRegion = headers.findIndex(h => h.includes('REGION') || h.includes('REGIÓN'));
  const colFecha  = headers.indexOf('FECHA');
  
  const prodMap = {};
  const regionMap = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const prod  = String(row[colProd] || '').trim();
    const venta = parseFloat(row[colVenta]) || 0;
    const region = String(row[colRegion] || 'Sin región').trim();
    
    if (!prod || venta <= 0) continue;
    
    // Acumular por producto
    if (!prodMap[prod]) prodMap[prod] = { nombre: prod, sku: row[colSku], venta: 0, unidades: 0 };
    prodMap[prod].venta += venta;
    prodMap[prod].unidades++;
    
    // Acumular por región
    if (!regionMap[region]) regionMap[region] = { region, venta: 0 };
    regionMap[region].venta += venta;
  }
  
  // Top 8 productos
  const topProducts = Object.values(prodMap)
    .sort((a, b) => b.venta - a.venta)
    .slice(0, 8)
    .map(p => ({
      nombre: p.nombre,
      sku: p.sku,
      venta: Math.round(p.venta),
      unidades: p.unidades
    }));
  
  // Regiones
  const regionStats = Object.values(regionMap)
    .sort((a, b) => b.venta - a.venta)
    .map(r => ({ region: r.region, venta: Math.round(r.venta) }));
  
  return { topProducts, regionStats };
}

// ── getHistorico: devuelve historial mensual para el gráfico ───────
function getHistorico() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_RESUMEN);
  if (!sheet) throw new Error('Hoja Resumen no encontrada');
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const colAño   = headers.findIndex(h => h.includes('AO') || h.includes('AÑO'));
  const colMes   = headers.indexOf('MES');
  const colVenta = headers.findIndex(h => h.includes('VENTA') && !h.includes('UNI'));
  const colMarg  = headers.findIndex(h => h.includes('MARGEN') && h.includes('%'));
  const colUni   = headers.findIndex(h => h.includes('UNI') || h.includes('UNID'));
  
  const MESES_MAP = {
    'ENERO':1,'FEBRERO':2,'MARZO':3,'ABRIL':4,'MAYO':5,'JUNIO':6,
    'JULIO':7,'AGOSTO':8,'SEPTIEMBRE':9,'OCTUBRE':10,'NOVIEMBRE':11,'DICIEMBRE':12
  };
  const DIAS_MES = [0,31,28,31,30,31,30,31,31,30,31,30,31];
  const isLeap = y => (y%4===0 && y%100!==0) || y%400===0;
  
  const snaps = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const año  = parseInt(row[colAño]);
    const mesStr = String(row[colMes] || '').trim().toUpperCase();
    const mesNum = MESES_MAP[mesStr];
    if (!año || !mesNum) continue;
    
    const diasMes = (mesNum === 2 && isLeap(año)) ? 29 : DIAS_MES[mesNum];
    const date = new Date(año, mesNum - 1, diasMes);
    const venta = parseFloat(row[colVenta]) || 0;
    if (venta <= 0) continue;
    
    const dd = String(diasMes).padStart(2,'0');
    const mm = String(mesNum).padStart(2,'0');
    snaps.push({
      date: `${dd}/${mm}/${String(año).slice(-2)}`,
      ts: date.getTime(),
      rev30: parseFloat((venta / 1e6).toFixed(1)),
      margen: parseFloat(row[colMarg]) || null,
      orders30: parseInt(row[colUni]) || null
    });
  }
  
  // Ordenar por fecha
  snaps.sort((a, b) => a.ts - b.ts);
  return { snaps, count: snaps.length };
}

// ── getTopProducts: devuelve top productos del detalle ─────────────
function getTopProducts() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_DETALLE);
  if (!sheet) throw new Error('Hoja Detalle no encontrada');
  const stats = processDetalle(sheet);
  return stats.topProducts;
}

// ── invalidateCache: fuerza reconstrucción del cache ───────────────
function invalidateCache() {
  CacheService.getScriptCache().remove('dashboard_data');
  Logger.log('Cache invalidado a las ' + new Date().toISOString());
}

// ── Triggers automáticos ───────────────────────────────────────────
// Instalar triggers: Menú Extensiones → Apps Script → Desencadenadores
// O ejecutar setupTriggers() una vez desde este editor

function setupTriggers() {
  // Eliminar triggers existentes del mismo proyecto
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  
  // Trigger diario a las 08:00, 13:00, 18:00 hora Chile (UTC-3)
  // Apps Script usa zona local, ajustar según zona configurada en el proyecto
  ScriptApp.newTrigger('invalidateCache')
    .timeBased().atHour(11).everyDays(1).create(); // 08:00 Chile = 11:00 UTC
  ScriptApp.newTrigger('invalidateCache')
    .timeBased().atHour(16).everyDays(1).create(); // 13:00 Chile = 16:00 UTC
  ScriptApp.newTrigger('invalidateCache')
    .timeBased().atHour(21).everyDays(1).create(); // 18:00 Chile = 21:00 UTC
  
  Logger.log('Triggers configurados: 08:00 / 13:00 / 18:00 hora Chile');
}

// ── Utilidades ────────────────────────────────────────────────────
function capitalize(str) {
  return str.charAt(0) + str.slice(1).toLowerCase();
}
