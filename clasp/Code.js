/**
 * Tally Ops Interface — Backend v2 (Google Apps Script)
 * ======================================================
 * v2 agrega: envío directo de correos desde juan@ (GmailApp),
 * inicialización automática del esquema (pestañas + plantillas +
 * clientes), y respuestas legibles vía JSONP.
 *
 * ACTUALIZAR DEPLOY (Juan, 1 minuto):
 * 1. Abrir el Sheet → Extensiones → Apps Script
 * 2. Reemplazar TODO el contenido de Code.gs con este archivo
 * 3. Guardar (⌘S) → Implementar → Administrar implementaciones
 *    → lápiz ✏️ → Versión: "Nueva versión" → Implementar
 *    (la URL /exec NO cambia)
 * 4. La primera vez pedirá autorizar permisos de Gmail — aceptar.
 *
 * Después de redesplegar, Talia ejecuta init_schema y el sistema
 * queda operativo.
 */

const TOKEN = 'tly-ops-2026-Jm9xQ4vKp7Rd3TzN8wHs';

// Base de datos del sistema (Tally Ops DB) — SIEMPRE por ID, sin importar
// desde qué archivo se haya creado este proyecto de Apps Script.
const DB_ID = '1A5TSql1ksUHQ8DBYwfTDrj_V3J1HGAs8cgCF9mijmnQ';

// Sheet maestro de datos de clientes (Accounting_DataModel)
const DATAMODEL_ID = '1_RrCnxuh0mg7cDTNCqIm5o3S2OmlSs7CGzU5zS3FNsI';
// Spreadsheet de reportes contables (Balance_general, balanza_comprobacion, calculo_impuestos,
// estado_resultados, reportes_extra) — conectado a AppSheet, referenciado desde Clientes_por_periodo.
const REPORTES_ID = '1AtBItd-kqNtm-QB72byTyySQ9WaY91vYqHA_tAa0tMg';
/** Localiza la pestaña de una tabla AppSheet: primero en el DataModel, luego en Reportes. */
function hojaDeTabla(tabla) {
  try { const s = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName(tabla); if (s) return s; } catch (e) {}
  try { return SpreadsheetApp.openById(REPORTES_ID).getSheetByName(tabla); } catch (e) { return null; }
}

// Sheet PRIVADO de usuarios de la interfaz (login). Crear un Google Sheet
// nuevo SIN compartir por link, y pegar aquí su ID (el de la URL).
// Luego ejecutar una vez la acción init_usuarios (o la función initUsuarios del editor).
const USUARIOS_ID = '1_QcRFNCMkZUPunfdf7b8Crr5UgRxG8gf_eMD8ACUw-w';

const SENDER_NAME = 'Juan Vélez — Tally';

/* ══════════════ ENTRADAS HTTP ══════════════ */

// GET: health, o acciones vía query params (con soporte JSONP → respuesta legible desde la interfaz)
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) {
    return out({ ok: true, service: 'tally-ops-interface', version: 2, ts: new Date().toISOString() }, p.callback);
  }
  if (p.token !== TOKEN) return out({ ok: false, error: 'token inválido' }, p.callback);
  try {
    const body = p.payload ? JSON.parse(p.payload) : p;
    body.action = p.action;
    return out(handle(body), p.callback);
  } catch (err) {
    return out({ ok: false, error: String(err) }, p.callback);
  }
}

// POST: mismas acciones vía JSON (la interfaz lo usa en modo no-cors)
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return out({ ok: false, error: 'token inválido' });
    return out(handle(body));
  } catch (err) {
    return out({ ok: false, error: String(err) });
  }
}

/* ── Cliente Anthropic con reintentos ──
   529/429/5xx son saturación temporal de Anthropic, no errores del sistema:
   3 intentos con espera creciente (0s, 2.5s, 8s) antes de rendirse. */
function claudeApi(key, payloadObj) {
  const opts = {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payloadObj)
  };
  const esperas = [0, 2500, 8000];
  let resp = null;
  for (let i = 0; i < esperas.length; i++) {
    if (esperas[i]) Utilities.sleep(esperas[i]);
    resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', opts);
    const c = resp.getResponseCode();
    if (c === 200) return resp;
    if (c !== 429 && c !== 500 && c !== 502 && c !== 503 && c !== 529) return resp; // no reintentable
  }
  return resp;
}

/* ══════════════ RUTEO DE ACCIONES ══════════════ */

function handle(body) {
  const ss = SpreadsheetApp.openById(DB_ID);
  const now = new Date().toISOString();

  switch (body.action) {

    /* ── Bandeja: flujo de respuestas ── */
    case 'save_prompt':
      return setEmailFields(ss, body.email_id, { prompt_juan: body.prompt, estado: 'Prompt recibido' }, now);
    case 'approve_draft':
      return setEmailFields(ss, body.email_id, { draft_final: body.draft_final || '', estado: 'Aprobado' }, now);
    case 'request_new_draft':
      return setEmailFields(ss, body.email_id, { prompt_juan: body.prompt, estado: 'Prompt recibido' }, now);
    case 'discard':
      return setEmailFields(ss, body.email_id, { estado: 'Descartado' }, now);

    /* ── Mantenimiento: compactar Clientes_por_periodo (elimina filas 100% vacías) ──
       Las filas fantasma infladas por getLastRow() ralentizan el proxy del dashboard. */
    case 'cxp_compactar': {
      const shC = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName('Clientes_por_periodo');
      if (!shC) return { ok: false, error: 'Clientes_por_periodo no encontrada' };
      const dataC = shC.getDataRange().getValues();
      const vaciasC = [];
      for (let i = 1; i < dataC.length; i++) {
        if (dataC[i].every(function (v) { return v === '' || v === null; })) vaciasC.push(i + 1);
      }
      // borrar de abajo hacia arriba, agrupando rangos contiguos
      let borradasC = 0;
      for (let i = vaciasC.length - 1; i >= 0;) {
        let fin = i;
        while (i > 0 && vaciasC[i - 1] === vaciasC[i] - 1) i--;
        shC.deleteRows(vaciasC[i], vaciasC[fin] - vaciasC[i] + 1);
        borradasC += vaciasC[fin] - vaciasC[i] + 1;
        i--;
      }
      return { ok: true, filas_vacias_eliminadas: borradasC, filas_restantes: shC.getLastRow() };
    }

    /* ── Dashboard: presencia documental por período (SOP Auditoría p2–p11) ──
       body.periodo = 'YYYY-MM'. Un documento EXISTE si su tabla destino tiene fila
       con el mismo PeriodID (regla del SOP 319325ede0a380f3afd0c83e2e3ed59a).
       Devuelve, por tipo de documento, los company_ids con documento en el período. */
    case 'docs_presencia': {
      // body.periodo = 'YYYY-MM'. Regla SOP p2-p11: documento cargado = fila de la tabla
      // destino en ese período (con su archivo). Esquemas REALES verificados 30-jul-2026:
      // la columna tipo PeriodID puede llamarse company_id/Company_id y trae '2026-3_AZ...'
      // o '2026-03_AZ...'; el período también viene en MesPeriodo+AñoPeríodo. El archivo
      // vive en columnas tipo Documento/UrlVentas/URLRetencion/Archivo.
      const perD = String(body.periodo || '').trim();
      const mD = perD.match(/^(\d{4})-(\d{1,2})$/);
      if (!mD) return { ok: false, error: 'periodo inválido (usa YYYY-MM)' };
      const anioD = mD[1], mesND = parseInt(mD[2], 10);
      const varD = [anioD + '-' + mesND, anioD + '-' + String(mesND).padStart(2, '0')];
      const MESES_D = ['', 'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
      const mesNombreD = MESES_D[mesND];
      const normD = function (x) {
        return String(x || '').toLowerCase().replace(/[\s_]/g, '')
          .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
          .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u').replace(/ñ/g, 'n');
      };
      const TABS_D = {
        decl: ['declaracion_periodo'],
        edo:  ['Estados_cuenta', 'Estados_cuentas'],
        diot: ['diot_periodo'],
        vtas: ['Reportes_de_venta', 'Reportes_de_ventas'],
        ret:  ['Retenciones_por_periodo', 'Retenciones_por_periodos'],
        inv:  ['Inventario_por_periodo', 'inventario_por_periodo', 'Inventario_por_periodos']
      };
      const abrirTabD = function (nombres) {
        const libros = [DATAMODEL_ID, REPORTES_ID];
        for (let b = 0; b < libros.length; b++) {
          let libro = null;
          try { libro = SpreadsheetApp.openById(libros[b]); } catch (e) { continue; }
          for (let n = 0; n < nombres.length; n++) {
            const sh = libro.getSheetByName(nombres[n]);
            if (sh) return sh;
          }
        }
        return null;
      };
      const presD = {};
      const debugD = {};
      Object.keys(TABS_D).forEach(function (k) {
        presD[k] = [];
        const shD = abrirTabD(TABS_D[k]);
        if (!shD) { debugD[k] = 'tab no encontrada en DataModel ni Reportes'; return; }
        const dataD = shD.getDataRange().getValues();
        if (dataD.length < 2) { debugD[k] = shD.getName() + ': sin filas'; return; }
        const headD = dataD[0].map(normD);
        // columnas de interés
        let iMes = headD.indexOf('mesperiodo');
        let iAnio = headD.indexOf('anoperiodo');
        let iProv = headD.indexOf('idprov');
        let iDoc = -1;
        for (let h = 0; h < headD.length; h++) {
          if (/documento|url|archivo/.test(headD[h])) { iDoc = h; break; }
        }
        const setD = {};
        for (let i = 1; i < dataD.length; i++) {
          const fila = dataD[i];
          // 1) período + company desde cualquier celda con patrón 'YYYY-M_ID' (primeras 4 columnas)
          let cid = '', enPeriodo = false;
          for (let cCol = 0; cCol < Math.min(4, fila.length); cCol++) {
            const mm = String(fila[cCol] || '').trim().match(/^(\d{4}-\d{1,2})_(.+)$/);
            if (mm) {
              cid = mm[2].trim();
              enPeriodo = (mm[1] === varD[0] || mm[1] === varD[1]);
              break;
            }
          }
          // 2) vía alterna por MesPeriodo + AñoPeríodo (unión: vale cualquiera de las dos;
          //    hay filas con PeriodID mal capturado pero Mes/Año correctos — visto 30-jul)
          if (!enPeriodo && iMes >= 0 && iAnio >= 0) {
            const mesTxt = normD(fila[iMes]);
            const anioTxt = String(fila[iAnio] || '').trim();
            if (mesTxt === mesNombreD && anioTxt === anioD) enPeriodo = true;
          }
          if (!enPeriodo) continue;
          if (!cid && iProv >= 0) cid = String(fila[iProv] || '').trim();
          if (!cid) continue;
          // 3) documento cargado: si la tabla tiene columna de archivo, exigirla no vacía
          if (iDoc >= 0 && String(fila[iDoc] || '').trim() === '') continue;
          setD[cid] = true;
        }
        presD[k] = Object.keys(setD);
        debugD[k] = shD.getName() + ': ' + presD[k].length + ' con doc en ' + perD;
      });
      // ── Universo del período: SOLO clientes con operación (DeclaracionTipo='Con datos') ──
      // Se lee directo del Sheet (no del proxy GViz, que omite columnas) e incluye owner
      // y nombre desde Clients_Load para que el dashboard no tenga que cruzar nada.
      const univD = [];
      let cerosD = 0;
      try {
        const shCpp = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName('Clientes_por_periodo');
        const shCl = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName('Clients_Load');
        const ownerMapD = {}, nombreMapD = {}, suspMapD = {};
        if (shCl) {
          const dCl = shCl.getDataRange().getValues();
          const hCl = dCl[0].map(normD);
          const iId = hCl.indexOf('companyid'), iOw = hCl.indexOf('owner');
          const iNm = hCl.indexOf('clientname'), iSu = hCl.indexOf('suspension');
          for (let i = 1; i < dCl.length; i++) {
            const idc = String(dCl[i][iId] || '').trim(); if (!idc) continue;
            ownerMapD[idc] = iOw >= 0 ? String(dCl[i][iOw] || '').trim() : '';
            nombreMapD[idc] = iNm >= 0 ? String(dCl[i][iNm] || '').trim() : '';
            suspMapD[idc] = iSu >= 0 ? String(dCl[i][iSu] || '').trim() : '';
          }
        }
        if (shCpp) {
          const dCpp = shCpp.getDataRange().getValues();
          const hCpp = dCpp[0].map(normD);
          const jPid = hCpp.indexOf('periodid'), jCid = hCpp.indexOf('companyid');
          const jTipo = hCpp.indexOf('declaraciontipo'), jEst = hCpp.indexOf('estadocliente');
          const jMes = hCpp.indexOf('mesperiodo'), jAnio = hCpp.indexOf('anoperiodo');
          const vistosD = {};
          for (let i = 1; i < dCpp.length; i++) {
            const fila = dCpp[i];
            // período: por PeriodID normalizado o por Mes+Año
            let enPer = false;
            if (jPid >= 0) {
              const mp = String(fila[jPid] || '').trim().match(/^(\d{4})-0?(\d{1,2})_/);
              if (mp && mp[1] === anioD && parseInt(mp[2], 10) === mesND) enPer = true;
            }
            if (!enPer && jMes >= 0 && jAnio >= 0) {
              if (normD(fila[jMes]) === mesNombreD && String(fila[jAnio] || '').trim() === anioD) enPer = true;
            }
            if (!enPer) continue;
            const cidU = jCid >= 0 ? String(fila[jCid] || '').trim() : '';
            if (!cidU || vistosD[cidU]) continue;
            const tipoU = jTipo >= 0 ? String(fila[jTipo] || '').trim() : '';
            if (tipoU !== 'Con datos') { cerosD++; continue; }   // sin operación → fuera
            vistosD[cidU] = true;
            univD.push({
              cid: cidU,
              estado: jEst >= 0 ? (String(fila[jEst] || '').trim() || 'Sin estado') : 'Sin estado',
              owner: ownerMapD[cidU] || '',
              nombre: nombreMapD[cidU] || '',
              suspension: suspMapD[cidU] || ''
            });
          }
        }
      } catch (e) { debugD.universo = 'ERROR: ' + e; }
      debugD.universo_con_datos = univD.length + ' clientes (excluidos ' + cerosD + ' en ceros/sin tipo)';
      return { ok: true, periodo: perD, presencia: presD, universo: univD, excluidos_ceros: cerosD, detalle: debugD };
    }

    /* ── Tareas: auto-bloqueo por vencimiento ──
       Toda tarea 'Sin iniciar' o 'En proceso' cuya fecha_entrega ya pasó se mueve
       a 'Bloqueado'. Corre dentro de sync_inbox (cron 15 min); también invocable directo. */
    case 'tareas_auto_bloqueo': {
      const shT = ss.getSheetByName('Tareas');
      if (!shT) return { ok: true, bloqueadas: 0 };
      const dataT = shT.getDataRange().getValues();
      const hoyT = Utilities.formatDate(new Date(), 'America/Mexico_City', 'yyyy-MM-dd');
      let nB = 0;
      for (let i = 1; i < dataT.length; i++) {
        const estT = String(dataT[i][8] || '').trim();          // col 9 estado
        if (estT !== 'Sin iniciar' && estT !== 'En proceso' && estT !== '') continue;
        const feRaw = dataT[i][13];                              // col 14 fecha_entrega
        if (!feRaw) continue;
        const fe = (feRaw instanceof Date)
          ? Utilities.formatDate(feRaw, 'America/Mexico_City', 'yyyy-MM-dd')
          : String(feRaw).substring(0, 10);
        if (fe && fe < hoyT) {
          shT.getRange(i + 1, 9).setValue('Bloqueado');
          shT.getRange(i + 1, 11).setValue(now);                 // ultima_actualizacion
          nB++;
        }
      }
      return { ok: true, bloqueadas: nB };
    }

    /* ── Dashboard: Pulso Semanal — ítems editables por columna ──
       body: { item_id?, columna (movio|atorado|foco), chip, chip_tipo (done|block|warn|next), titulo, texto }
       Sin item_id crea uno nuevo; con item_id edita. Pestaña Pulso_Semanal del Ops DB. */
    case 'pulso_set': {
      const shP = ss.getSheetByName('Pulso_Semanal') || ss.insertSheet('Pulso_Semanal');
      if (!String(shP.getRange(1, 1).getValue()).trim()) {
        shP.getRange(1, 1, 1, 7).setValues([['item_id', 'columna', 'chip', 'chip_tipo', 'titulo', 'texto', 'updated']]);
      }
      let iidP = String(body.item_id || '').trim();
      const nuevoP = !iidP;
      if (nuevoP) iidP = 'pw' + Date.now().toString(36) + Math.floor(Math.random() * 90 + 10);
      const valsP = [iidP, body.columna || 'movio', body.chip || '', body.chip_tipo || 'done', body.titulo || '', body.texto || '', now];
      const rowP = nuevoP ? 0 : findRow(shP, 1, iidP);
      if (rowP) shP.getRange(rowP, 1, 1, 7).setValues([valsP]); else shP.appendRow(valsP);
      return { ok: true, item_id: iidP };
    }
    case 'pulso_del': {
      const shPd = ss.getSheetByName('Pulso_Semanal');
      if (!shPd) return { ok: false, error: 'sin pestaña Pulso_Semanal' };
      const rowPd = findRow(shPd, 1, String(body.item_id || '').trim());
      if (!rowPd) return { ok: false, error: 'item no encontrado' };
      shPd.deleteRow(rowPd);
      return { ok: true };
    }

    /* ── Dashboard: Journey Top 10 — estados de tags y notas por cliente ──
       body: { company_id, cliente?, tags? (JSON string), nota? }
       Upsert por company_id en la pestaña Journey_Top10 del Ops DB. */
    case 'journey_set': {
      const shJ = ss.getSheetByName('Journey_Top10') || ss.insertSheet('Journey_Top10');
      if (!String(shJ.getRange(1, 1).getValue()).trim()) {
        shJ.getRange(1, 1, 1, 5).setValues([['company_id', 'cliente', 'tags', 'nota', 'updated']]);
      }
      const cidJ = String(body.company_id || '').trim();
      if (!cidJ) return { ok: false, error: 'company_id requerido' };
      const rowJ = findRow(shJ, 1, cidJ);
      if (rowJ) {
        if (body.cliente !== undefined) shJ.getRange(rowJ, 2).setValue(body.cliente);
        if (body.tags !== undefined) shJ.getRange(rowJ, 3).setValue(body.tags);
        if (body.nota !== undefined) shJ.getRange(rowJ, 4).setValue(body.nota);
        shJ.getRange(rowJ, 5).setValue(now);
      } else {
        shJ.appendRow([cidJ, body.cliente || '', body.tags || '', body.nota || '', now]);
      }
      return { ok: true, company_id: cidJ };
    }

    /* ── Contabilidad: registro de períodos en Clientes_por_periodo (DataModel) ──
       body.filas = [{ company_id, periodo:'2026-4', rfc?, ventas?, iva?, isr?, retencion?, nota? }]
       Clona NombreCliente/Constitutive de la última fila existente de la empresa.
       Idempotente: si el PeriodID ya existe, lo salta. */
    case 'periodo_registrar': {
      let logSh = null;
      try { logSh = ss.getSheetByName('Log_Periodos') || ss.insertSheet('Log_Periodos'); } catch (e) {}
      const plog = function (msg) { try { if (logSh) logSh.appendRow([new Date().toISOString(), msg]); } catch (e) {} };
      try {
        const MESES = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
        const sh = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName('Clientes_por_periodo');
        if (!sh) { plog('ERROR: Clientes_por_periodo no encontrada'); return { ok: false, error: 'Clientes_por_periodo no encontrada' }; }
        let filas = body.filas || [];
        if (typeof filas === 'string') { try { filas = JSON.parse(filas.replace(/\s+/g, '')); } catch (e) { plog('ERROR parse filas string: ' + e); return { ok: false, error: 'filas string inválido' }; } }
        plog('inicio: keys=' + Object.keys(body).join(',') + ' filas=' + filas.length);
        const data = sh.getDataRange().getValues();
        const existentes = {};
        const ultimaDeCompany = {};
        for (let i = 1; i < data.length; i++) {
          const pid = String(data[i][0] || '').trim();
          const cid = String(data[i][1] || '').trim();
          if (pid) existentes[pid] = true;
          if (cid) ultimaDeCompany[cid] = i;
        }
        const detalle = [];
        let registradas = 0;
        filas.forEach(function (f) {
          const cid = String(f.company_id || '').replace(/\s+/g, '');
          const per = String(f.periodo || '').replace(/\s+/g, '');   // ej. '2026-4'
          const pid = per + '_' + cid;
          if (!cid || !per) { detalle.push(pid + ':faltan datos'); return; }
          if (existentes[pid]) { detalle.push(pid + ':ya existe'); return; }
          const idx = ultimaDeCompany[cid];
          if (idx === undefined) { detalle.push(pid + ':company sin filas previas'); return; }
          const base = data[idx];
          const mesNum = parseInt(per.split('-')[1], 10) || 0;
          const anio = per.split('-')[0];
          const ventas = Number(f.ventas || 0);
          const iva = (f.iva === undefined || f.iva === '' || Number(f.iva) === 0) ? '' : Number(f.iva);
          const isr = (f.isr === undefined || f.isr === '' || Number(f.isr) === 0) ? '' : Number(f.isr);
          const ret = Number(f.retencion || 0);
          const conDatos = ventas > 0 || iva !== '' || isr !== '' || ret > 0;
          const row = new Array(25).fill('');
          row[0] = pid;                                   // A PeriodID
          row[1] = cid;                                   // B Company_id
          row[2] = base[2];                               // C NombreCliente (clonado)
          row[3] = base[3];                               // D Constitutive (clonado)
          row[4] = conDatos ? 'Con datos' : 'Ceros';      // E DeclaracionTipo
          row[5] = String(f.rfc || base[5] || '').replace(/\s+/g, ''); // F RFC
          row[6] = ventas;                                // G SalesLastPeriod
          row[7] = 'Declarado';                           // H EstadoCliente
          row[8] = iva;                                   // I IVA_pagar
          row[9] = isr;                                   // J ISR_pagar
          row[11] = ret > 0 ? ret : '';                   // L PeriodoRetencion
          row[13] = ventas;                               // N VentasList
          row[15] = MESES[mesNum] || '';                  // P MesPeriodo
          row[16] = anio;                                 // Q AñoPeríodo
          row[23] = Utilities.formatDate(new Date(), 'America/Mexico_City', 'M/d/yyyy'); // X Fecha_Declaracion
          row[24] = f.nota || 'Registro vía Tally Ops';   // Y Notas_Declaracion
          try {
            sh.appendRow(row);                            // appendRow expande la grilla si hace falta
            existentes[pid] = true;
            registradas++;
            detalle.push(pid + ':OK');
          } catch (e) {
            detalle.push(pid + ':ERROR ' + e);
          }
        });
        plog('fin: registradas=' + registradas + ' | ' + detalle.join(' ; '));
        return { ok: true, registradas: registradas, detalle: detalle };
      } catch (e) {
        plog('EXCEPCION: ' + e);
        return { ok: false, error: String(e) };
      }
    }

    /* ── Seller Central ── */
    case 'sc_decision': {
      const sh = ss.getSheetByName('SC_Seguimiento');
      const row = findRow(sh, 1, body.company_id);
      if (!row) return { ok: false, error: 'company_id no encontrado' };
      sh.getRange(row, 11).setValue(body.decision);   // K aprobacion
      sh.getRange(row, 12).setValue(now);              // L fecha_aprobacion
      sh.getRange(row, 15).setValue(now);              // O ultima_actualizacion
      return { ok: true, aprobacion: body.decision };
    }

    case 'sc_mark_sent': {
      // Sella una fila de SC como enviada con su escenario (lo usa send_direct y backfills)
      const sh = ss.getSheetByName('SC_Seguimiento');
      const row = findRow(sh, 1, body.company_id);
      if (!row) return { ok: true, skipped: 'no está en la cola SC' };
      if (body.escenario) sh.getRange(row, 5).setValue(String(body.escenario));
      sh.getRange(row, 13).setValue(now);   // M fecha_ultimo_envio
      sh.getRange(row, 15).setValue(now);   // O ultima_actualizacion
      return { ok: true, company_id: body.company_id };
    }
    case 'sc_set_escenario': {
      // Juan cambia manualmente el escenario de envío de un cliente desde la interfaz
      const sh = ss.getSheetByName('SC_Seguimiento');
      const row = findRow(sh, 1, body.company_id);
      if (!row) return { ok: false, error: 'company_id no encontrado' };
      if (['Esc.1','Esc.2','Esc.3'].indexOf(String(body.escenario)) < 0) return { ok: false, error: 'escenario inválido' };
      sh.getRange(row, 5).setValue(String(body.escenario));                 // E escenario_actual
      const notas = String(sh.getRange(row, 14).getValue() || '');
      sh.getRange(row, 14).setValue((notas ? notas + ' | ' : '') + 'Escenario fijado manualmente por Juan: ' + body.escenario);
      sh.getRange(row, 15).setValue(now);
      return { ok: true, company_id: body.company_id, escenario: body.escenario };
    }

    /* ── Redacción desde cero: ENVÍO DIRECTO desde juan@ ── */
    case 'send_direct': {
      // body: { to, cc?, subject, body_text, company_id, cliente, categoria, plantilla }
      if (!body.to || !body.subject || !body.body_text) return { ok: false, error: 'faltan campos (to/subject/body_text)' };
      // REGLA DURA: SIEMPRE en copia customersuccess@, accounting@ y el OWNER del cliente
      const ownEm = ownerEmail(ss, body.company_id);
      const senderD = resolveSender(ss, body, body.categoria);
      const ccDirect = mergeCc(String(body.cc || '') + (ownEm ? ',' + ownEm : '') + ',' + ccCliente(ss, body.company_id), String(body.to), senderD);
      // Adjuntos desde Drive vía tablas AppSheet: body.adjuntos = [{tabla, columna}]
      const docs = [];
      if (body.adjuntos && body.adjuntos.length) {
        const faltantes = [];
        body.adjuntos.forEach(function(a){
          const d = buscarDocumento(body.company_id, a.tabla, a.columna);
          if (d) docs.push(d); else faltantes.push(a.tabla + '.' + a.columna);
        });
        if (faltantes.length) return { ok: false, error: 'documento(s) no encontrados para ' + (body.cliente || body.company_id) + ': ' + faltantes.join(', ') + ' — envío bloqueado para no prometer adjuntos vacíos' };
      }
      if (senderD === 'juan@tally.legal') {
        const opts = { cc: ccDirect, name: SENDER_NAME };
        if (docs.length) opts.attachments = docs.map(function(d){ return d.blob; });
        GmailApp.sendEmail(String(body.to), String(body.subject), String(body.body_text), opts);
      } else {
        const rd = sendViaDwd(senderD, String(body.to), ccDirect, String(body.subject), String(body.body_text), null, docs);
        if (!rd.ok) return rd;
      }
      const sal = getOrCreate(ss, 'Salientes', HEADERS.Salientes);
      const id = 'SAL-' + Date.now() + '-' + (body.company_id || 'X');
      sal.appendRow([id, now, body.company_id || '', body.cliente || '', String(body.to) + (body.cc ? ' cc:' + body.cc : ''),
                     body.categoria || '', body.plantilla || '', String(body.subject), String(body.body_text), 'Enviado', senderD + ' (interfaz)']);
      // Si es un envío de Seller Central a un cliente en la cola → sellar la fila con su escenario
      if (/seller\s*central/i.test(String(body.categoria || '')) && body.company_id) {
        const MAPA_ESC = { 'TPL-SC-01': 'Esc.1', 'TPL-SC-02': 'Esc.2', 'TPL-SC-03': 'Esc.3', 'TPL-SC-04': 'Esc.1' };
        const sch = ss.getSheetByName('SC_Seguimiento');
        const srow = findRow(sch, 1, body.company_id);
        if (srow) {
          const escMap = MAPA_ESC[String(body.plantilla || '')];
          if (escMap) sch.getRange(srow, 5).setValue(escMap);
          sch.getRange(srow, 13).setValue(now);
          sch.getRange(srow, 15).setValue(now);
        }
      }
      return { ok: true, saliente_id: id, enviado_a: body.to, desde: senderD };
    }

    /* ── Escritura del AGENTE (triage, drafts, envíos, SC, log) ──
       El agente usa estas acciones vía GET mientras el MCP de Sheets
       no tenga permiso de editor sobre este archivo. */
    case 'append_email': {
      const sh = ss.getSheetByName('Emails');
      ensureCcCol(sh);
      if (findRow(sh, 1, body.email_id)) return { ok: true, skipped: 'ya existe', email_id: body.email_id };
      sh.appendRow([body.email_id||'', body.thread_id||'', body.cuenta||'', body.fecha_recibido||'', body.remitente_nombre||'',
                    body.remitente_email||'', body.company_id||'', body.cliente||'', body.asunto||'', body.resumen||'',
                    body.categoria||'', body.prioridad||'Media', 'Nuevo', '', '', '', '', '', '', body.notas_agente||'', now,
                    body.cc_originales||'', body.mensaje_original||'']);
      return { ok: true, inserted: body.email_id };
    }
    case 'update_email': {
      // Actualiza cualquier columna de una fila de Emails. Clave: email_id (el actual).
      // body.fields = { columna: valor } — puede incluir email_id/thread_id para corregir llaves.
      const sh = ss.getSheetByName('Emails');
      const row = findRow(sh, 1, body.email_id);
      if (!row) return { ok: false, error: 'email_id no encontrado: ' + body.email_id };
      const COLS = { email_id:1, thread_id:2, cuenta:3, fecha_recibido:4, remitente_nombre:5, remitente_email:6,
                     company_id:7, cliente:8, asunto:9, resumen:10, categoria:11, prioridad:12, estado:13,
                     prompt_juan:14, draft_asunto:15, draft_cuerpo:16, draft_final:17, notas_agente:20, cc_originales:22, mensaje_original:23 };
      ensureCcCol(sh);
      const f = body.fields || {};
      Object.keys(f).forEach(k => { if (COLS[k]) sh.getRange(row, COLS[k]).setValue(f[k]); });
      sh.getRange(row, 21).setValue(now);
      return { ok: true, updated: body.email_id, campos: Object.keys(f).length };
    }
    case 'set_draft': {
      const sh = ss.getSheetByName('Emails');
      const row = findRow(sh, 1, body.email_id);
      if (!row) return { ok: false, error: 'email_id no encontrado' };
      sh.getRange(row, 15).setValue(body.draft_asunto||'');  // O
      sh.getRange(row, 16).setValue(body.draft_cuerpo||'');  // P
      sh.getRange(row, 13).setValue('Draft listo');          // M
      sh.getRange(row, 21).setValue(now);                    // U
      return { ok: true, estado: 'Draft listo' };
    }
    case 'mark_sent': {
      const sh = ss.getSheetByName('Emails');
      const row = findRow(sh, 1, body.email_id);
      if (!row) return { ok: false, error: 'email_id no encontrado' };
      sh.getRange(row, 13).setValue('Enviado');
      sh.getRange(row, 18).setValue(now);                    // R fecha_envio
      sh.getRange(row, 19).setValue(body.msg_id||'');        // S
      sh.getRange(row, 21).setValue(now);
      return { ok: true, estado: 'Enviado' };
    }
    case 'upsert_sc': {
      const sh = ss.getSheetByName('SC_Seguimiento');
      const row = findRow(sh, 1, body.company_id);
      const vals = [body.company_id||'', body.cliente||'', body.owner||'', body.periodo||'', body.escenario_actual||'',
                    body.estado_sop||'', body.dias_desde_contacto||'', body.bloque||'', body.motivo_exclusion||'', body.accion_pendiente||''];
      if (row) {
        sh.getRange(row, 1, 1, 10).setValues([vals]);        // A-J (no toca K aprobacion / L fecha)
        if (body.fecha_ultimo_envio) sh.getRange(row, 13).setValue(body.fecha_ultimo_envio);
        if (body.notas !== undefined) sh.getRange(row, 14).setValue(body.notas);
        sh.getRange(row, 15).setValue(now);
        return { ok: true, updated: body.company_id };
      }
      sh.appendRow(vals.concat(['Pendiente', '', body.fecha_ultimo_envio||'', body.notas||'', now]));
      return { ok: true, inserted: body.company_id };
    }
    case 'update_cliente': {
      // La interfaz cura el registro: agrega/corrige correos de contacto de un cliente
      const sh = ss.getSheetByName('Clientes');
      const row = findRow(sh, 1, body.company_id);
      if (!row) return { ok: false, error: 'company_id no encontrado en Clientes' };
      if (body.contacto_email !== undefined) sh.getRange(row, 7).setValue(String(body.contacto_email)); // G
      if (body.cc_email !== undefined) sh.getRange(row, 8).setValue(String(body.cc_email));             // H
      if (body.contacto_nombre !== undefined) sh.getRange(row, 6).setValue(String(body.contacto_nombre)); // F
      const notas = String(sh.getRange(row, 9).getValue() || '');
      sh.getRange(row, 9).setValue((notas ? notas + ' | ' : '') + 'Correo actualizado desde la interfaz ' + now.slice(0, 10));
      return { ok: true, company_id: body.company_id };
    }
    case 'sync_clientes': {
      // Sincroniza correos de la pestaña Clientes desde Clients_Load (Accounting_DataModel):
      // extrae SOLO correos (regex), rellena vacíos, depura ruido (estados de stripe, contexto).
      // NUNCA pisa un correo válido ya registrado.
      const cl = ss.getSheetByName('Clientes');
      const dm = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName('Clients_Load');
      const dmData = dm.getDataRange().getValues();
      const emailRe = /[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/g;
      const mapa = {};
      for (let i = 1; i < dmData.length; i++) {
        const cid = String(dmData[i][0] || '').trim();
        if (!cid) continue;
        let ems = [];
        for (let j = 1; j < dmData[i].length; j++) {
          const cell = String(dmData[i][j] || '');
          if (cell.length > 120) continue; // campos largos de contexto = ruido
          const f = cell.match(emailRe);
          if (f) ems = ems.concat(f);
        }
        ems = ems.map(function(e){ return e.toLowerCase(); })
                 .filter(function(e, ix, a){ return a.indexOf(e) === ix; }).slice(0, 3);
        if (ems.length) mapa[cid] = ems.join(', ');
      }
      const data = cl.getDataRange().getValues();
      let rellenados = 0, depurados = 0;
      for (let i = 1; i < data.length; i++) {
        const cid = String(data[i][0] || '').trim();
        if (!cid) continue;
        const actual = String(data[i][6] || '').trim();
        const ext = (actual.match(emailRe) || []).filter(function(e, ix, a){ return a.indexOf(e) === ix; });
        const limpio = ext.join(', ');
        if (ext.length) {
          if (limpio !== actual) { cl.getRange(i + 1, 7).setValue(limpio); depurados++; } // quitar ruido, conservar correos
        } else if (mapa[cid]) {
          cl.getRange(i + 1, 7).setValue(mapa[cid]); rellenados++;                        // rellenar vacío
        } else if (actual) {
          cl.getRange(i + 1, 7).setValue(''); depurados++;                                // pura basura → vaciar
        }
      }
      return { ok: true, rellenados: rellenados, depurados: depurados };
    }
    case 'request_run': {
      // La interfaz solicita una corrida inmediata del agente (botón 🔄 o al guardar Config).
      // El vigilante del agente revisa esta llave cada 5 minutos.
      const sh = ss.getSheetByName('Config');
      const row = findRow(sh, 1, 'corrida_solicitada');
      if (row) sh.getRange(row, 2).setValue(now);
      else sh.appendRow(['corrida_solicitada', now]);
      return { ok: true, corrida_solicitada: now };
    }
    case 'update_config': {
      // Upsert de un par llave-valor en Config (lo usan la interfaz y el agente)
      const sh = ss.getSheetByName('Config');
      const row = findRow(sh, 1, body.key);
      if (row) sh.getRange(row, 2).setValue(body.value);
      else sh.appendRow([body.key, body.value]);
      return { ok: true, key: body.key };
    }
    case 'upsert_plantilla': {
      // Sync Notion → Sheet (cache de plantillas para la interfaz). Clave: plantilla_id
      const sh = ss.getSheetByName('Plantillas');
      const row = findRow(sh, 1, body.plantilla_id);
      const vals = [body.plantilla_id||'', body.categoria||'', body.nombre||'', body.asunto||'', body.cuerpo||''];
      if (row) { sh.getRange(row, 1, 1, 5).setValues([vals]); return { ok: true, updated: body.plantilla_id }; }
      sh.appendRow(vals);
      return { ok: true, inserted: body.plantilla_id };
    }
    /* ══════════ 🎖️ COMANDO CONTABILIDAD — micro-cascada de actividades (HTML local de Juan) ══════════ */
    case 'cmd_item_add': {
      if (!String(body.titulo || '').trim()) return { ok: false, error: 'titulo vacío' };
      const sh = getOrCreate(ss, 'Comando_Items', ['item_id', 'fase', 'titulo', 'detalle', 'orden', 'estado', 'creado', 'actualizado']);
      const id = 'CMD-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
      sh.appendRow([id, String(body.fase || 'General'), String(body.titulo).trim(), String(body.detalle || ''),
                    Number(body.orden) || 999, ['Pendiente','En curso','Hecho'].indexOf(String(body.estado)) >= 0 ? String(body.estado) : 'Pendiente', now, now]);
      return { ok: true, item_id: id };
    }
    case 'cmd_item_edit': {
      const sh = ss.getSheetByName('Comando_Items');
      if (!sh) return { ok: false, error: 'sin pestaña Comando_Items' };
      const row = findRow(sh, 1, body.item_id);
      if (!row) return { ok: false, error: 'item no encontrado' };
      if (body.titulo !== undefined && String(body.titulo).trim()) sh.getRange(row, 3).setValue(String(body.titulo).trim());
      if (body.detalle !== undefined) sh.getRange(row, 4).setValue(String(body.detalle));
      if (body.fase !== undefined) sh.getRange(row, 2).setValue(String(body.fase));
      if (body.orden !== undefined) sh.getRange(row, 5).setValue(Number(body.orden) || 999);
      if (body.estado !== undefined && ['Pendiente','En curso','Hecho'].indexOf(String(body.estado)) >= 0) sh.getRange(row, 6).setValue(String(body.estado));
      sh.getRange(row, 8).setValue(now);
      return { ok: true, item_id: body.item_id };
    }
    case 'cmd_item_del': {
      const sh = ss.getSheetByName('Comando_Items');
      if (!sh) return { ok: false, error: 'sin pestaña Comando_Items' };
      const row = findRow(sh, 1, body.item_id);
      if (!row) return { ok: false, error: 'item no encontrado' };
      sh.deleteRow(row);
      return { ok: true, deleted: body.item_id };
    }
    /* ══════════ 🗺️ ROADMAP KPIs CONTABILIDAD (dashboard.tallylegal.io/accounting) ══════════ */
    case 'kpi_item_add': {
      // body: {texto, estado: 'Next Steps'|'En proceso'|'Implementado'}
      if (!String(body.texto || '').trim()) return { ok: false, error: 'texto vacío' };
      const sh = getOrCreate(ss, 'Roadmap_KPIs', ['item_id', 'texto', 'estado', 'creado', 'actualizado']);
      const id = 'RK-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
      const est = ['Next Steps', 'En proceso', 'Implementado'].indexOf(String(body.estado)) >= 0 ? String(body.estado) : 'Next Steps';
      sh.appendRow([id, String(body.texto).trim(), est, now, now]);
      return { ok: true, item_id: id, estado: est };
    }
    case 'kpi_item_edit': {
      const sh = ss.getSheetByName('Roadmap_KPIs');
      if (!sh) return { ok: false, error: 'sin pestaña Roadmap_KPIs' };
      const row = findRow(sh, 1, body.item_id);
      if (!row) return { ok: false, error: 'item no encontrado' };
      if (body.texto !== undefined && String(body.texto).trim()) sh.getRange(row, 2).setValue(String(body.texto).trim());
      if (body.estado !== undefined && ['Next Steps', 'En proceso', 'Implementado'].indexOf(String(body.estado)) >= 0) sh.getRange(row, 3).setValue(String(body.estado));
      sh.getRange(row, 5).setValue(now);
      return { ok: true, item_id: body.item_id };
    }
    case 'kpi_item_del': {
      // Eliminación explícita desde la interfaz del roadmap (acción del usuario, con confirm previo)
      const sh = ss.getSheetByName('Roadmap_KPIs');
      if (!sh) return { ok: false, error: 'sin pestaña Roadmap_KPIs' };
      const row = findRow(sh, 1, body.item_id);
      if (!row) return { ok: false, error: 'item no encontrado' };
      sh.deleteRow(row);
      return { ok: true, deleted: body.item_id };
    }
    case 'kpi_item_estado': {
      const sh = ss.getSheetByName('Roadmap_KPIs');
      if (!sh) return { ok: false, error: 'sin pestaña Roadmap_KPIs' };
      const row = findRow(sh, 1, body.item_id);
      if (!row) return { ok: false, error: 'item no encontrado' };
      const est = ['Next Steps', 'En proceso', 'Implementado'].indexOf(String(body.estado)) >= 0 ? String(body.estado) : 'Next Steps';
      sh.getRange(row, 3).setValue(est);
      sh.getRange(row, 5).setValue(now);
      return { ok: true, item_id: body.item_id, estado: est };
    }
    /* ══════════ ✅ TAREAS (kanban Sin iniciar / Finalizado, alimentado desde Bandeja y Documentación) ══════════ */
    case 'usuarios_publicos': {
      // Lista de responsables asignables (usuarios activos de la plataforma) — SIN contraseñas.
      // ALIAS de nombre visible (27-jul-2026, pedido de Juan): contabilidad@ es Arturo Cerón,
      // y ninguna "cuenta genérica" aparece como responsable.
      const ALIAS_NOMBRE = { 'contabilidad@tally.legal': 'Arturo Cerón' };
      const NOMBRES_OCULTOS = ['cuenta genérica', 'cuenta generica', 'genérica', 'generica'];
      const EMAILS_OCULTOS = ['elizabeth@tally.legal']; // salió de Tally 29-jul-2026 — fuera del selector de responsables, hasta nueva orden
      const u = checkUser(body.auth);
      if (!u.ok) return u;
      const us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios');
      const data = us.getDataRange().getValues();
      const list = [];
      for (let i = 1; i < data.length; i++) {
        const em = String(data[i][0] || '').trim().toLowerCase();
        if (!em || String(data[i][3]).toLowerCase() === 'no') continue;
        let nom = String(data[i][2] || em);
        if (EMAILS_OCULTOS.indexOf(em) !== -1) continue;
        if (ALIAS_NOMBRE[em]) nom = ALIAS_NOMBRE[em];
        else if (NOMBRES_OCULTOS.indexOf(nom.toLowerCase().trim()) !== -1) continue;
        list.push({ email: em, nombre: nom });
      }
      return { ok: true, usuarios: list };
    }
    case 'tarea_crear': {
      // body: {tareas:[{titulo, descripcion, responsable, origen, ref_id, cliente, clientes, fecha_entrega}], auth}
      const u = checkUser(body.auth);
      if (!u.ok) return u;
      const sh = getOrCreate(ss, 'Tareas', ['tarea_id','fecha_creacion','creado_por','responsable','titulo','origen','ref_id','cliente','estado','fecha_finalizacion','ultima_actualizacion','descripcion','clientes','fecha_entrega']);
      ensureTareasCols(sh);
      const ids = [];
      (body.tareas || []).forEach(function(t){
        if (!String(t.titulo || '').trim()) return;
        const id = 'T-' + Date.now() + '-' + Math.floor(Math.random() * 10000);
        sh.appendRow([id, now, u.email, String(t.responsable || '').toLowerCase(), String(t.titulo), String(t.origen || ''), String(t.ref_id || ''), String(t.cliente || ''), 'Sin iniciar', '', now,
                      String(t.descripcion || ''), String(t.clientes || ''), String(t.fecha_entrega || '')]);
        ids.push(id);
      });
      return { ok: true, creadas: ids.length, ids: ids };
    }
    case 'tarea_edit': {
      // Edición en línea desde la lista: titulo, descripcion, responsable, fecha_entrega, estado, clientes
      const u = checkUser(body.auth);
      if (!u.ok) return u;
      const sh = ss.getSheetByName('Tareas');
      if (!sh) return { ok: false, error: 'sin pestaña Tareas' };
      ensureTareasCols(sh);
      const row = findRow(sh, 1, body.tarea_id);
      if (!row) return { ok: false, error: 'tarea no encontrada' };
      if (body.titulo !== undefined && String(body.titulo).trim()) sh.getRange(row, 5).setValue(String(body.titulo).trim());
      if (body.responsable !== undefined) sh.getRange(row, 4).setValue(String(body.responsable).toLowerCase());
      if (body.descripcion !== undefined) sh.getRange(row, 12).setValue(String(body.descripcion));
      if (body.clientes !== undefined) sh.getRange(row, 13).setValue(String(body.clientes));
      if (body.fecha_entrega !== undefined) sh.getRange(row, 14).setValue(String(body.fecha_entrega));
      if (body.estado !== undefined) {
        const est = body.estado === 'Finalizado' ? 'Finalizado' : 'Sin iniciar';
        sh.getRange(row, 9).setValue(est);
        sh.getRange(row, 10).setValue(est === 'Finalizado' ? now : '');
      }
      sh.getRange(row, 11).setValue(now);
      return { ok: true, tarea_id: body.tarea_id };
    }
    case 'tarea_del': {
      const u = checkUser(body.auth);
      if (!u.ok) return u;
      const sh = ss.getSheetByName('Tareas');
      if (!sh) return { ok: false, error: 'sin pestaña Tareas' };
      const row = findRow(sh, 1, body.tarea_id);
      if (!row) return { ok: false, error: 'tarea no encontrada' };
      sh.deleteRow(row);
      return { ok: true, tarea_id: body.tarea_id, eliminada: true };
    }
    case 'tarea_estado': {
      const u2 = checkUser(body.auth);
      if (!u2.ok) return u2;
      const sh = ss.getSheetByName('Tareas');
      if (!sh) return { ok: false, error: 'sin pestaña Tareas aún' };
      const row = findRow(sh, 1, body.tarea_id);
      if (!row) return { ok: false, error: 'tarea no encontrada' };
      const est = body.estado === 'Finalizado' ? 'Finalizado' : 'Sin iniciar';
      sh.getRange(row, 9).setValue(est);
      sh.getRange(row, 10).setValue(est === 'Finalizado' ? now : '');
      sh.getRange(row, 11).setValue(now);
      return { ok: true, tarea_id: body.tarea_id, estado: est };
    }
    case 'guardar_borrador': {
      // 📝 Guarda instrucción/ediciones SIN generar ni enviar — queda pendiente de confirmación
      // de otro miembro del equipo. El agente NO procesa filas en estado Borrador.
      const sh = ss.getSheetByName('Emails');
      const row = findRow(sh, 1, body.email_id);
      if (!row) return { ok: false, error: 'email_id no encontrado' };
      if (body.prompt_juan !== undefined) sh.getRange(row, 14).setValue(String(body.prompt_juan)); // N
      if (body.draft_final !== undefined) sh.getRange(row, 17).setValue(String(body.draft_final)); // Q
      sh.getRange(row, 13).setValue('Borrador');   // M estado
      sh.getRange(row, 21).setValue(now);          // U ultima_actualizacion
      return { ok: true, email_id: body.email_id, estado: 'Borrador' };
    }
    case 'sync_documentacion': {
      // 📦 Documentación mensual — elegibilidad y checklist por cliente.
      // Blueprint: docs/documentacion-mensual-blueprint.md (decidido con Juan 20-jul-2026).
      return syncDocumentacion(ss);
    }
    case 'verificar_recursos': {
      // Aviso preventivo: reporta qué datos/documentos esperados NO están cargados en el sistema
      // para un cliente, ANTES de generar/enviar el correo. body: {company_id, datos:[], docs:[]}
      const faltan = [];
      (body.datos || []).forEach(function(d){
        if (!buscarDato(body.company_id, d.tabla, d.columna)) faltan.push({ tipo: 'dato', tabla: d.tabla, columna: d.columna });
      });
      (body.docs || []).forEach(function(d){
        if (!existeDocumento(body.company_id, d.tabla, d.columna)) faltan.push({ tipo: 'documento', tabla: d.tabla, columna: d.columna });
      });
      return { ok: true, company_id: body.company_id, faltantes: faltan };
    }
    case 'append_log': {
      const sh = ss.getSheetByName('Log');
      sh.appendRow([now, body.corrida||'manual', body.correos_revisados||0, body.correos_filtrados_fuera||0,
                    body.nuevos_triaged||0, body.drafts_generados||0, body.enviados||0, body.sc_actualizados||0,
                    body.errores||0, body.notas||'']);
      return { ok: true, logged: now };
    }

    /* ── REDACCIÓN EN TIEMPO REAL (Claude API desde Apps Script) ──
       Requiere UNA VEZ: editor Apps Script → ⚙️ Configuración del proyecto →
       Propiedades del script → propiedad ANTHROPIC_API_KEY = tu clave de
       console.anthropic.com. La clave nunca sale del Apps Script. */
    case 'generate_draft': {
      const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!key) return { ok: false, error: 'SIN_API_KEY: configura ANTHROPIC_API_KEY en Propiedades del script; mientras tanto el draft lo hará el agente en su corrida.' };
      const sh = ss.getSheetByName('Emails');
      const row = findRow(sh, 1, body.email_id);
      if (!row) return { ok: false, error: 'email_id no encontrado' };
      const v = sh.getRange(row, 1, 1, 21).getValues()[0];
      const [ , , , , remNombre, remEmail, companyId, cliente, asunto, resumen, , , , promptJuan] = v;
      const notas = v[19];
      const cfg = ss.getSheetByName('Config');
      const firmaRow = findRow(cfg, 1, 'firma_juan');
      const firma = firmaRow ? cfg.getRange(firmaRow, 2).getValue() : 'Best regards,\nJuan Vélez\nTally';
      const modelRow = findRow(cfg, 1, 'modelo_redaccion');
      const model = modelRow ? String(cfg.getRange(modelRow, 2).getValue()) : 'claude-sonnet-5';

      const system = 'Eres el asistente de redacción de Juan Vélez, Director de Estrategia de Tally (contabilidad para empresas extranjeras en México). Redactas la RESPUESTA a un correo de cliente siguiendo EXACTAMENTE la instrucción de Juan.\n\n🔒 REGLA DURA DE IDIOMA (prioridad máxima, sin excepciones): el correo de salida se redacta SIEMPRE en el idioma en que el CLIENTE escribió su correo original — NUNCA en el idioma de la instrucción de Juan (que suele venir en español). Detecta el idioma del cliente en este orden: (1) marcador "🌐 Idioma" en el contexto si existe; (2) el asunto original del correo; (3) citas textuales del cliente dentro del contexto; (4) si nada es concluyente, inglés (default de clientes extranjeros). Aunque la instrucción de Juan esté en español, si el cliente escribió en inglés, respondes en inglés.\n\nDemás reglas: tono profesional, cálido y directo; NO inventes compromisos, montos ni fechas que Juan no haya dado; no uses corchetes ni placeholders; cierra con la firma tal cual se te da. Devuelve SOLO el cuerpo del correo, sin asunto ni comentarios.';
      const msgOriginal = String(sh.getRange(row, 23).getValue() || '');
      const user = 'CORREO A RESPONDER\nDe: ' + remNombre + ' <' + remEmail + '>\nCliente: ' + cliente + (companyId ? ' (' + companyId + ')' : '') + '\nAsunto: ' + asunto +
        (msgOriginal ? '\n\nMENSAJE ORIGINAL DEL CLIENTE (texto literal — responde a ESTO):\n' + msgOriginal : '') +
        '\n\nResumen del correo: ' + resumen + '\nContexto del caso:\n' + notas +
        '\n\nINSTRUCCIÓN DE JUAN PARA LA RESPUESTA:\n' + (body.prompt || promptJuan) +
        '\n\nFIRMA A USAR AL FINAL:\n' + firma;

      const resp = claudeApi(key, { model: model, max_tokens: 1200, system: system, messages: [{ role: 'user', content: user }] });
      const code = resp.getResponseCode();
      if (code !== 200) return { ok: false, error: 'Claude API ' + code + ((code === 529 || code === 429) ? ' — Anthropic saturado (reintenté 3 veces). Espera 1-2 min y vuelve a generar.' : '') + ': ' + resp.getContentText().slice(0, 200) };
      // La respuesta puede traer varios bloques (p.ej. razonamiento); tomar el bloque de TEXTO
      const blocks = (JSON.parse(resp.getContentText()).content) || [];
      const textBlock = blocks.filter(function(b){ return b && b.type === 'text' && b.text; })[0];
      if (!textBlock) return { ok: false, error: 'respuesta sin bloque de texto: ' + JSON.stringify(blocks).slice(0, 200) };
      const draft = String(textBlock.text).trim();
      const draftAsunto = String(asunto).startsWith('Re:') ? String(asunto) : 'Re: ' + asunto;
      if (body.prompt) sh.getRange(row, 14).setValue(body.prompt);
      sh.getRange(row, 15).setValue(draftAsunto);
      sh.getRange(row, 16).setValue(draft);
      sh.getRange(row, 13).setValue('Draft listo');
      sh.getRange(row, 21).setValue(now);
      return { ok: true, draft_asunto: draftAsunto, draft_cuerpo: draft };
    }
    case 'generate_new': {
      // Redacción libre con IA para CORREOS NUEVOS: Juan da objetivo+contexto,
      // Claude redacta asunto y cuerpo personalizados por cliente.
      const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!key) return { ok: false, error: 'SIN_API_KEY: configura ANTHROPIC_API_KEY en Propiedades del script.' };
      const cfg = ss.getSheetByName('Config');
      const firmaRow = findRow(cfg, 1, 'firma_juan');
      const firma = firmaRow ? cfg.getRange(firmaRow, 2).getValue() : 'Best regards,\nJuan Vélez\nTally';
      const modelRow = findRow(cfg, 1, 'modelo_redaccion');
      const model = modelRow ? String(cfg.getRange(modelRow, 2).getValue()) : 'claude-sonnet-5';

      const system = 'Eres el asistente de redacción de Juan Vélez, Director de Estrategia de Tally (contabilidad para empresas extranjeras en México). Redactas un CORREO NUEVO a un cliente siguiendo el objetivo y contexto que da Juan.\n\nReglas: idioma = el que Juan indique en su instrucción; si no indica ninguno, inglés (los clientes son empresas extranjeras). Tono profesional, cálido y directo. NO inventes datos, montos, fechas ni compromisos que Juan no haya dado. Sin corchetes ni placeholders. Cierra con la firma tal cual se te da.\n\nFORMATO DE SALIDA OBLIGATORIO: primera línea exactamente "ASUNTO: <asunto del correo>", luego una línea en blanco, luego el cuerpo completo. Nada más.';
      // Datos reales del cliente desde AppSheet (redacción libre con variables de columnas)
      let datosCtx = '';
      if (body.datos && body.datos.length) {
        const lineas = [];
        body.datos.forEach(function(d){
          const val = buscarDato(body.company_id, d.tabla, d.columna);
          lineas.push(d.tabla + '.' + d.columna + ' = ' + (val || '(sin dato registrado)'));
        });
        datosCtx = '\n\nDATOS REALES DEL CLIENTE (inclúyelos con naturalidad donde correspondan; si alguno dice "sin dato registrado", NO lo menciones ni lo inventes):\n' + lineas.join('\n');
      }
      const user = 'CLIENTE DESTINATARIO: ' + (body.cliente || '') + (body.company_id ? ' (' + body.company_id + ')' : '') +
        '\nNOMBRE DEL CONTACTO: ' + (body.contact_name || 'no disponible — usa un saludo genérico profesional') +
        datosCtx +
        '\n\nOBJETIVO Y CONTEXTO DEL CORREO (instrucción de Juan):\n' + (body.prompt || '') +
        '\n\nFIRMA A USAR AL FINAL:\n' + firma;

      const resp = claudeApi(key, { model: model, max_tokens: 1200, system: system, messages: [{ role: 'user', content: user }] });
      const code = resp.getResponseCode();
      if (code !== 200) return { ok: false, error: 'Claude API ' + code + ((code === 529 || code === 429) ? ' — Anthropic saturado (reintenté 3 veces). Espera 1-2 min y vuelve a generar.' : '') + ': ' + resp.getContentText().slice(0, 200) };
      const blocks = (JSON.parse(resp.getContentText()).content) || [];
      const tb = blocks.filter(function(b){ return b && b.type === 'text' && b.text; })[0];
      if (!tb) return { ok: false, error: 'respuesta sin bloque de texto' };
      const mt2 = String(tb.text).trim().match(/^ASUNTO:\s*(.+)\n+([\s\S]+)$/);
      if (!mt2) return { ok: false, error: 'formato inesperado del modelo: ' + String(tb.text).slice(0, 120) };
      return { ok: true, subject: mt2[1].trim(), body_text: mt2[2].trim() };
    }
    case 'translate_text': {
      // Traduce asunto+cuerpo al idioma pedido (para envíos con plantilla según prefijo IN/AZ)
      const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!key) return { ok: false, error: 'SIN_API_KEY' };
      const idioma = String(body.idioma || 'es') === 'es' ? 'español' : 'inglés';
      const resp = claudeApi(key, { model: 'claude-haiku-4-5-20251001', max_tokens: 1500,
          system: 'Traduce el correo al ' + idioma + ' manteniendo EXACTOS: tono profesional, formato, saltos de línea, variables {{...}}, campos [...] y direcciones de correo. FORMATO DE SALIDA: primera línea "ASUNTO: <asunto traducido>", línea en blanco, luego el cuerpo. Nada más.',
          messages: [{ role: 'user', content: 'ASUNTO: ' + (body.subject || '') + '\n\n' + (body.body_text || '') }] });
      if (resp.getResponseCode() !== 200) return { ok: false, error: 'Claude API ' + resp.getResponseCode() };
      const blocks2 = (JSON.parse(resp.getContentText()).content) || [];
      const tb2 = blocks2.filter(function(b){ return b && b.type === 'text' && b.text; })[0];
      if (!tb2) return { ok: false, error: 'sin texto' };
      const m2 = String(tb2.text).trim().match(/^ASUNTO:\s*(.+)\n+([\s\S]+)$/);
      if (!m2) return { ok: false, error: 'formato inesperado' };
      return { ok: true, subject: m2[1].trim(), body_text: m2[2].trim() };
    }
    case 'analizar_plantilla': {
      // Wizard de plantillas nuevas: la IA interpreta las variables [.] y propone el match
      const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
      if (!key) return { ok: false, error: 'SIN_API_KEY' };
      const resp = claudeApi(key, { model: 'claude-sonnet-5', max_tokens: 1200,
          system: 'Analizas variables de una plantilla de correo de Tally (contabilidad marketplaces México). Variables del SISTEMA disponibles (se llenan solas por cliente): {{contact_name}} nombre del contacto, {{company_name}} nombre de la empresa, {{period}} período fiscal (mes anterior), {{owner_name}} owner interno del cliente, {{firma}} firma del remitente. Para cada variable [entre corchetes] de la plantilla decide: tipo "sistema" (equivale a una variable del sistema — indica cuál), tipo "manual" (dato puntual que el usuario debe escribir al enviar, ej. fechas, montos acordados, temas), tipo "appsheet" (dato de texto que vive en las tablas del negocio: ventas, IVA/ISR, accesos, declaraciones — indica Tabla.Columna probable de: Clients_Load, Clientes_por_periodo, Accesos_SellerCentral, WeeklyPlan, Estados_cuenta, Reportes_de_venta), o tipo "documento" (la variable pide ADJUNTAR un archivo del cliente — indica Tabla.Columna de archivo entre: declaracion_periodo.Documento, Reportes_de_venta.Archivo, Retenciones_por_periodo.URLRetencion, Estados_cuenta.UrlVentas, Inventario_por_periodo.URLInventario, diot_periodo.Documento). Responde SOLO un JSON array: [{"var":"[nombre]","tipo":"sistema|manual|appsheet","match":"{{variable}} o Tabla.Columna o vacío","razon":"breve"}]',
          messages: [{ role: 'user', content: 'ASUNTO: ' + (body.subject || '') + '\n\nCUERPO:\n' + (body.body_text || '') }] });
      if (resp.getResponseCode() !== 200) return { ok: false, error: 'Claude API ' + resp.getResponseCode() };
      const blocks3 = (JSON.parse(resp.getContentText()).content) || [];
      const tb3 = blocks3.filter(function(b){ return b && b.type === 'text' && b.text; })[0];
      if (!tb3) return { ok: false, error: 'sin texto' };
      const jm = String(tb3.text).match(/\[[\s\S]*\]/);
      if (!jm) return { ok: false, error: 'sin JSON en respuesta' };
      try { return { ok: true, analisis: JSON.parse(jm[0]) }; }
      catch (e) { return { ok: false, error: 'JSON inválido de la IA' }; }
    }
    case 'send_reply': {
      // Envío inmediato de la respuesta aprobada, desde juan@.
      // REGLAS DURAS: (1) si el hilo existe en el buzón de juan@, responder DENTRO del hilo;
      // (2) CC = los CC originales del correo + SIEMPRE customersuccess@ y accounting@.
      const sh = ss.getSheetByName('Emails');
      ensureCcCol(sh);
      const row = findRow(sh, 1, body.email_id);
      if (!row) return { ok: false, error: 'email_id no encontrado' };
      const v = sh.getRange(row, 1, 1, 22).getValues()[0];
      const to = String(v[5]);
      const asuntoOrig = String(v[8]);
      let asunto = String(v[14] || ((asuntoOrig.toLowerCase().indexOf('re:') === 0 ? '' : 'Re: ') + asuntoOrig));
      const cuerpo = String(body.draft_final || v[16] || v[15]);
      if (to.indexOf('@') < 0) return { ok: false, error: 'la fila no tiene remitente_email válido' };
      const categoria = String(v[10] || '');
      const cuentaOrigen = String(v[2] || '').toLowerCase();      // cuenta que recibió el correo
      const threadOrigen = String(v[1] || '');                     // thread_id en ESA cuenta
      const sender = resolveSender(ss, body, categoria);
      const cc = mergeCc(String(v[21] || '') + ',' + ccCliente(ss, String(v[6] || '')), to, sender);

      let enHilo = false, resultado = null;

      /* ══ REGLA DURA DE HILO (auditoría 21-jul-2026): toda respuesta lleva In-Reply-To/References
         del mensaje original (leído vía DWD readonly en la cuenta que lo recibió) — así el hilo se
         mantiene en el buzón del CLIENTE en cualquier cliente de correo, sin importar el remitente. ══ */
      const hh = hiloHeaders(cuentaOrigen ? cuentaOrigen + '@tally.legal' : '', String(v[0] || ''));
      if (hh && hh.subject) {
        // Normaliza el asunto al del hilo real (requisito de la Gmail API cuando se usa threadId)
        asunto = 'Re: ' + hh.subject.replace(/^\s*((re|fwd|rv|fw)\s*:\s*)+/i, '').trim();
      }

      if (sender === 'juan@tally.legal') {
        // vía nativa: buscar el hilo en el buzón de juan@ (suele estar en CC)
        try {
          const subjClean = asuntoOrig.replace(/^\s*((re|fwd|rv|fw)\s*:\s*)+/i, '').replace(/"/g, ' ').trim();
          if (subjClean) {
            const ths = GmailApp.search('subject:"' + subjClean + '"', 0, 10);
            outer:
            for (let t = 0; t < ths.length; t++) {
              const msgs = ths[t].getMessages();
              for (let m = msgs.length - 1; m >= 0; m--) {
                if (String(msgs[m].getFrom()).toLowerCase().indexOf(to.toLowerCase()) >= 0) {
                  msgs[m].reply(cuerpo, { cc: cc, name: SENDER_NAME });
                  enHilo = true;
                  break outer;
                }
              }
            }
          }
        } catch (e) {}
        if (!enHilo && hh) {
          // Fallback 1: el hilo no está en el buzón de juan@ → DWD como juan@ CON headers de hilo
          const rj = sendViaDwd('juan@tally.legal', to, cc, asunto, cuerpo, null, null, hh);
          if (rj.ok) enHilo = true;
        }
        if (!enHilo) GmailApp.sendEmail(to, asunto, cuerpo, { cc: cc, name: SENDER_NAME }); // último recurso
        resultado = { ok: true };
      } else {
        // vía DWD: threadId solo agrupa NUESTRO buzón y solo es válido en la cuenta que recibió;
        // los headers de hilo (hh) son los que garantizan el hilo del lado del cliente.
        const tid = (sender === cuentaOrigen + '@tally.legal') ? threadOrigen : null;
        resultado = sendViaDwd(sender, to, cc, asunto, cuerpo, tid, null, hh);
        if (!resultado.ok && tid) {
          // Si la API rechazó el threadId (p.ej. asunto distinto), reintenta sin él pero CON headers
          resultado = sendViaDwd(sender, to, cc, asunto, cuerpo, null, null, hh);
        }
        if (!resultado.ok) return resultado;
        enHilo = !!tid || !!hh;
      }

      if (body.draft_final) sh.getRange(row, 17).setValue(body.draft_final);
      sh.getRange(row, 13).setValue('Enviado');
      sh.getRange(row, 18).setValue(now);
      sh.getRange(row, 19).setValue('desde: ' + sender);
      sh.getRange(row, 21).setValue(now);
      return { ok: true, enviado_a: to, cc: cc, en_hilo: enHilo, desde: sender };
    }

    /* ── 🌐 RESPUESTA OS — respuestas del equipo enviadas POR FUERA de la interfaz ──
       Recorre filas recientes de Emails, lee el hilo real vía DWD readonly en la cuenta
       que recibió el correo, y si encuentra un mensaje de @tally.legal que NO salió de
       la interfaz (≠ msg_id_enviado y ≠ correo original), marca estado="Respuesta OS"
       ("Out Side") y guarda quién/cuándo/snippet en la columna X respuesta_os.
       Las filas en Borrador conservan su estado (solo se anota la columna X). */
    case 'detectar_os': {
      const sh = ss.getSheetByName('Emails');
      if (!sh) return { ok: false, error: 'sin pestaña Emails' };
      ensureCcCol(sh);
      const data = sh.getDataRange().getValues();
      const max = Number(body.max || 25);
      let revisados = 0, detectados = 0; const errs = [];
      const tokens = {};
      for (let i = data.length - 1; i >= 1 && revisados < max; i--) {
        const v = data[i];
        const estado = String(v[12] || '');
        if (estado === 'Respuesta OS' || estado === 'Descartado') continue;
        if (String(v[23] || '').trim()) continue;                    // ya anotada en corridas previas
        const threadId = String(v[1] || '');
        let cuenta = String(v[2] || '').trim();
        if (!threadId || !cuenta) continue;
        if (cuenta.indexOf('@') < 0) cuenta += '@tally.legal';
        revisados++;
        try {
          if (!(cuenta in tokens)) tokens[cuenta] = dwdToken(cuenta, 'https://www.googleapis.com/auth/gmail.readonly');
          const token = tokens[cuenta];
          if (!token) { errs.push('sin token DWD para ' + cuenta); continue; }
          const r = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(cuenta)
            + '/threads/' + encodeURIComponent(threadId) + '?format=metadata&metadataHeaders=From&metadataHeaders=Message-ID',
            { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
          if (r.getResponseCode() !== 200) continue;
          const th = JSON.parse(r.getContentText());
          const emailIdOrig = String(v[0] || '');
          const msgIdEnviado = String(v[18] || '');
          let os = null;
          (th.messages || []).forEach(function(m) {
            const h = {};
            ((m.payload && m.payload.headers) || []).forEach(function(x){ h[String(x.name).toLowerCase()] = x.value; });
            const from = String(h['from'] || '').toLowerCase();
            if (from.indexOf('@tally.legal') < 0) return;            // solo mensajes del equipo
            if (m.id === emailIdOrig) return;                        // el correo original del cliente
            if (msgIdEnviado && (m.id === msgIdEnviado || String(h['message-id'] || '') === msgIdEnviado)) return; // salió de la interfaz
            if (!os || Number(m.internalDate) > Number(os.ts)) {
              os = { ts: m.internalDate, from: h['from'] || '', snippet: String(m.snippet || '').substring(0, 300) };
            }
          });
          if (os) {
            const f = Utilities.formatDate(new Date(Number(os.ts)), 'America/Mexico_City', 'yyyy-MM-dd HH:mm');
            sh.getRange(i + 1, 24).setValue(os.from + ' | ' + f + ' | ' + os.snippet);
            if (estado !== 'Borrador') sh.getRange(i + 1, 13).setValue('Respuesta OS');
            sh.getRange(i + 1, 21).setValue(now);
            detectados++;
          }
        } catch (e) { errs.push(String(e).substring(0, 80)); }
      }

      /* ── FASE 2: detección CRUZADA entre cuentas ──
         Cubre el caso en que alguien respondió al cliente desde OTRA cuenta @tally.legal
         distinta a la que recibió el correo (ese mensaje no aparece en el hilo receptor).
         Índice: enviados recientes (in:sent, 3 días) de TODAS las cuentas monitoreadas;
         match por destinatario (email del cliente) + asunto normalizado. */
      try {
        const t0 = Date.now();
        const cfg = ss.getSheetByName('Config');
        let cuentas = [];
        if (cfg) {
          const cv = cfg.getDataRange().getValues();
          for (let c = 0; c < cv.length; c++) {
            if (String(cv[c][0]).trim() === 'cuentas_monitoreadas') { cuentas = String(cv[c][1] || '').split(','); break; }
          }
        }
        cuentas = cuentas.map(function(a){ a = String(a).trim(); if (!a) return ''; return a.indexOf('@') < 0 ? a + '@tally.legal' : a; }).filter(Boolean);
        const normSubj = function(s){ return String(s || '').replace(/^(\s*(re|rv|fwd?|fw)\s*:)+/i, '').replace(/\s+/g, ' ').trim().toLowerCase(); };
        // enviados recientes de cada cuenta monitoreada
        const enviados = [];
        cuentas.forEach(function(acc) {
          if (Date.now() - t0 > 90000) return; // presupuesto de tiempo
          try {
            if (!(acc in tokens)) tokens[acc] = dwdToken(acc, 'https://www.googleapis.com/auth/gmail.readonly');
            if (!tokens[acc]) return;
            const lr = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(acc)
              + '/messages?q=' + encodeURIComponent('in:sent newer_than:3d') + '&maxResults=12',
              { headers: { Authorization: 'Bearer ' + tokens[acc] }, muteHttpExceptions: true });
            if (lr.getResponseCode() !== 200) return;
            ((JSON.parse(lr.getContentText()).messages) || []).forEach(function(mm) {
              if (Date.now() - t0 > 90000) return;
              const mr = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(acc)
                + '/messages/' + mm.id + '?format=metadata&metadataHeaders=To&metadataHeaders=Cc&metadataHeaders=Subject&metadataHeaders=Message-ID',
                { headers: { Authorization: 'Bearer ' + tokens[acc] }, muteHttpExceptions: true });
              if (mr.getResponseCode() !== 200) return;
              const md = JSON.parse(mr.getContentText());
              const hh = {};
              ((md.payload && md.payload.headers) || []).forEach(function(x){ hh[String(x.name).toLowerCase()] = x.value; });
              enviados.push({ acc: acc, id: md.id, ts: md.internalDate, snippet: String(md.snippet || '').substring(0, 300),
                              to: String((hh['to'] || '') + ',' + (hh['cc'] || '')).toLowerCase(),
                              subj: normSubj(hh['subject']), msgIdH: String(hh['message-id'] || '') });
            });
          } catch (e2) {}
        });
        // match contra filas aún sin OS
        if (enviados.length) {
          const data2 = sh.getDataRange().getValues();
          for (let i = data2.length - 1; i >= 1; i--) {
            const v = data2[i];
            if (String(v[23] || '').trim()) continue;
            const estado = String(v[12] || '');
            if (estado === 'Respuesta OS' || estado === 'Descartado') continue;
            const cliEmail = String(v[5] || '').toLowerCase().trim();
            const asunto = normSubj(v[8]);
            if (!cliEmail || !asunto) continue;
            const msgIdEnviado = String(v[18] || '');
            for (let k = 0; k < enviados.length; k++) {
              const sd = enviados[k];
              if (sd.to.indexOf(cliEmail) < 0 || sd.subj !== asunto) continue;
              if (msgIdEnviado && (sd.id === msgIdEnviado || sd.msgIdH === msgIdEnviado)) continue; // salió de la interfaz
              const f = Utilities.formatDate(new Date(Number(sd.ts)), 'America/Mexico_City', 'yyyy-MM-dd HH:mm');
              sh.getRange(i + 1, 24).setValue(sd.acc + ' (cuenta cruzada) | ' + f + ' | ' + sd.snippet);
              if (estado !== 'Borrador') sh.getRange(i + 1, 13).setValue('Respuesta OS');
              sh.getRange(i + 1, 21).setValue(now);
              detectados++;
              break;
            }
          }
        }
      } catch (e) { errs.push('fase2: ' + String(e).substring(0, 60)); }

      return { ok: true, revisados: revisados, detectados: detectados, errores: errs.slice(0, 5) };
    }

    /* ── Sincronización por correo de control ──
       El agente envía un correo interno (accounting@ → juan@) con asunto
       [TALLY-OPS-SYNC] y un bloque <<<JSON [...operaciones...] JSON>>>.
       Esta acción (GET corto) lee esos correos, aplica cada operación
       (append_email, set_draft, upsert_sc, append_log, mark_sent) y
       etiqueta el hilo como procesado. */
    case 'sync_inbox': {
      const label = GmailApp.getUserLabelByName('tally-ops-processed') || GmailApp.createLabel('tally-ops-processed');
      const threads = GmailApp.search('subject:"[TALLY-OPS-SYNC]" -label:tally-ops-processed newer_than:3d', 0, 10);
      let ops = 0, errs = [];
      threads.forEach(th => {
        try {
          const msgs = th.getMessages();
          const bodyTxt = msgs[msgs.length - 1].getPlainBody();
          const mt = bodyTxt.match(/<<<JSON([\s\S]*?)JSON>>>/);
          if (!mt) { errs.push('sin bloque JSON en: ' + th.getFirstMessageSubject()); }
          else {
            // Gmail inserta saltos de línea al transportar el correo;
            // los strings del bloque no llevan saltos intencionales → sanear.
            const clean = mt[1].replace(/[\u0000-\u001F]+/g, ' ').trim();
            const list = JSON.parse(clean);
            list.forEach(op => {
              if (op.action === 'sync_inbox' || op.action === 'init_schema') return;
              const r = handle(op);
              if (r && r.ok) ops++; else errs.push(JSON.stringify(r));
            });
          }
          th.addLabel(label);
          th.moveToArchive(); // fuera del inbox de Juan — queda etiquetado como registro
        } catch (e) { errs.push(String(e)); }
      });
      // Barrido de respuestas fuera de la interfaz (aprovecha la corrida de cron cada 15 min)
      let osres = null;
      try { osres = handle({ action: 'detectar_os', max: 20 }); } catch (e) { osres = { ok: false, error: String(e) }; }
      // Auto-bloqueo de tareas vencidas (misma corrida de 15 min)
      let tkres = null;
      try { tkres = handle({ action: 'tareas_auto_bloqueo' }); } catch (e) { tkres = { ok: false, error: String(e) }; }
      return { ok: true, threads: threads.length, operaciones_aplicadas: ops, errores: errs.slice(0, 5), respuesta_os: osres, tareas_bloqueadas: tkres };
    }

    /* ── LOGIN de la interfaz (usuarios en Sheet privado separado) ── */
    case 'login': {
      if (USUARIOS_ID.indexOf('PEGAR') === 0) return { ok: false, error: 'USUARIOS_ID sin configurar en el Apps Script' };
      let us;
      try { us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios'); }
      catch (e) { return { ok: false, error: 'no pude abrir el Sheet de usuarios: ' + e }; }
      if (!us) return { ok: false, error: 'el Sheet privado no tiene pestaña "Usuarios" — ejecuta init_usuarios' };
      const data = us.getDataRange().getValues();
      const email = String(body.email || '').trim().toLowerCase();
      const pass = String(body.password || '');
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === email) {
          if (String(data[i][3]).toLowerCase() === 'no') return { ok: false, error: 'usuario desactivado' };
          if (String(data[i][1]) === pass) {
            us.getRange(i + 1, 5).setValue(new Date().toISOString()); // último acceso
            const admin = (email === 'juan@tally.legal') || String(data[i][5] || '').toLowerCase() === 'admin';
            return { ok: true, nombre: String(data[i][2] || email), email: email, admin: admin };
          }
          return { ok: false, error: 'contraseña incorrecta' };
        }
      }
      return { ok: false, error: 'usuario no registrado' };
    }
    case 'user_list': {
      if (!checkAdmin(body.auth)) return { ok: false, error: 'solo el administrador puede gestionar usuarios' };
      const us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios');
      const data = us.getDataRange().getValues();
      const list = [];
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim()) list.push({ email: data[i][0], nombre: data[i][2], activo: data[i][3], ultimo_acceso: data[i][4], rol: data[i][5] || 'usuario' });
      }
      return { ok: true, usuarios: list };
    }
    case 'user_upsert': {
      // Alta/edición de usuario — SOLO admin. body.user = {email, password?, nombre?, activo?, rol?}
      if (!checkAdmin(body.auth)) return { ok: false, error: 'solo el administrador puede gestionar usuarios' };
      const u = body.user || {};
      const email2 = String(u.email || '').trim().toLowerCase();
      if (email2.indexOf('@') < 0) return { ok: false, error: 'email inválido' };
      const us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios');
      const row = findRow(us, 1, email2);
      if (row) {
        if (u.password) us.getRange(row, 2).setValue(String(u.password));
        if (u.nombre !== undefined) us.getRange(row, 3).setValue(String(u.nombre));
        if (u.activo !== undefined) us.getRange(row, 4).setValue(String(u.activo));
        if (u.rol !== undefined) us.getRange(row, 6).setValue(String(u.rol));
        return { ok: true, updated: email2 };
      }
      if (!u.password) return { ok: false, error: 'usuario nuevo requiere contraseña' };
      us.appendRow([email2, String(u.password), String(u.nombre || ''), String(u.activo || 'si'), '', String(u.rol || 'usuario')]);
      return { ok: true, inserted: email2 };
    }
    case 'init_usuarios': {
      if (USUARIOS_ID.indexOf('PEGAR') === 0) return { ok: false, error: 'USUARIOS_ID sin configurar' };
      const uss = SpreadsheetApp.openById(USUARIOS_ID);
      let us = uss.getSheetByName('Usuarios');
      if (!us) { us = uss.insertSheet('Usuarios'); }
      if (us.getLastRow() === 0) {
        us.getRange(1, 1, 1, 6).setValues([['email', 'password', 'nombre', 'activo', 'ultimo_acceso', 'rol']]);
        us.setFrozenRows(1);
        us.appendRow(['juan@tally.legal', 'CAMBIAME-' + Math.random().toString(36).slice(2, 8), 'Juan Vélez', 'si', '', 'admin']);
      }
      return { ok: true, usuarios_sheet: uss.getName() };
    }

    /* ── Inicialización del esquema (una vez; idempotente) ── */
    case 'init_schema':
      return initSchema(ss);

    default:
      return { ok: false, error: 'acción desconocida: ' + body.action };
  }
}

/**
 * Wrapper para TRIGGER DE TIEMPO (auto-sincronización).
 * Configurar UNA vez: editor Apps Script → ⏰ Activadores (reloj, menú izq.)
 * → + Agregar activador → función: cronSync → basado en tiempo → cada 15 minutos.
 * Con esto los correos [TALLY-OPS-SYNC] del agente se aplican solos al Sheet.
 */
function cronSync() {
  const r = handle({ action: 'sync_inbox' });
  console.log(JSON.stringify(r));
}

/**
 * DIAGNÓSTICO — ejecutar desde el editor para forzar el permiso de
 * "conectarse a un servicio externo" y probar la API de Claude.
 * Resultado esperado en el log: "✅ Claude respondió: OK"
 */
// Prueba directa del scope de Drive (adjuntos de plantillas). Correr desde el editor.
function probarDrive() {
  const raiz = DriveApp.getRootFolder().getName();
  console.log('✅ Drive OK — carpeta raíz visible: ' + raiz);
}

function probarConexionClaude() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) { console.log('❌ Falta ANTHROPIC_API_KEY en Propiedades del script'); return; }
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 10, messages: [{ role: 'user', content: 'Responde solo: OK' }] })
  });
  const code = resp.getResponseCode();
  if (code === 200) console.log('✅ Claude respondió: ' + JSON.parse(resp.getContentText()).content[0].text);
  else console.log('❌ Claude API ' + code + ': ' + resp.getContentText().slice(0, 300));
}

/* ══════════════ ESQUEMA ══════════════ */

const HEADERS = {
  Emails: ['email_id','thread_id','cuenta','fecha_recibido','remitente_nombre','remitente_email','company_id','cliente','asunto','resumen','categoria','prioridad','estado','prompt_juan','draft_asunto','draft_cuerpo','draft_final','fecha_envio','msg_id_enviado','notas_agente','ultima_actualizacion'],
  SC_Seguimiento: ['company_id','cliente','owner','periodo','escenario_actual','estado_sop','dias_desde_contacto','bloque','motivo_exclusion','accion_pendiente','aprobacion','fecha_aprobacion','fecha_ultimo_envio','notas','ultima_actualizacion'],
  Config: ['key','value'],
  Log: ['timestamp','corrida','correos_revisados','correos_filtrados_fuera','nuevos_triaged','drafts_generados','enviados','sc_actualizados','errores','notas'],
  Clientes: ['company_id','cliente','owner','suspension','subscription','contacto_nombre','contacto_email','cc_email','notas'],
  Plantillas: ['plantilla_id','categoria','nombre','asunto','cuerpo'],
  Salientes: ['saliente_id','fecha','company_id','cliente','destinatarios','categoria','plantilla','asunto','cuerpo','estado','enviado_por']
};

function initSchema(ss) {
  const created = [], seeded = [];

  Object.keys(HEADERS).forEach(name => {
    const sh = getOrCreate(ss, name, HEADERS[name]);
    if (sh.__wasCreated) created.push(name);
  });

  // Seed Config (solo si está vacía)
  const cfg = ss.getSheetByName('Config');
  if (cfg.getLastRow() < 2) {
    [['cuentas_monitoreadas','contabilidad@tally.legal, accounting@tally.legal, elizabeth@tally.legal, juan@tally.legal'],
     ['dominios_excluidos','no-reply, noreply, notifications, calendar-notification, mailer-daemon, stripe.com, google.com, appsheet.com, anthropic.com, openai.com, netlify, github.com, notion.so, slack.com, intuit, docusign, zoom.us'],
     ['asuntos_excluidos','unsubscribe, invitación:, invitation:, receipt, invoice #, payment received, renewal, license'],
     ['firma_juan','Best regards,\nJuan Vélez\nTally — Accounting & Tax\njuan@tally.legal'],
     ['version_esquema','2.0']
    ].forEach(r => cfg.appendRow(r));
    seeded.push('Config');
  }

  // Seed Plantillas (solo si está vacía)
  const pl = ss.getSheetByName('Plantillas');
  if (pl.getLastRow() < 2) {
    PLANTILLAS_SEED.forEach(r => pl.appendRow(r));
    seeded.push('Plantillas (' + PLANTILLAS_SEED.length + ')');
  }

  // Importar Clientes desde Accounting_DataModel.Clients_Load (solo si está vacía)
  const cl = ss.getSheetByName('Clientes');
  if (cl.getLastRow() < 2) {
    try {
      const dm = SpreadsheetApp.openById(DATAMODEL_ID).getSheetByName('Clients_Load');
      const data = dm.getDataRange().getValues();
      const H = data[0].map(String);
      const iId = H.indexOf('Company_Id'), iName = H.indexOf('ClientName'), iOwner = H.indexOf('Owner'),
            iSusp = H.indexOf('Suspension'), iSub = H.indexOf('SubscriptionType'), iMail = H.indexOf('stripe_email');
      const rows = [];
      for (let i = 1; i < data.length; i++) {
        const id = String(data[i][iId] || '').trim();
        if (!id) continue;
        rows.push([id, data[i][iName] || '', data[i][iOwner] || '', data[i][iSusp] || '', data[i][iSub] || '', '', data[i][iMail] || '', '', '']);
      }
      if (rows.length) cl.getRange(2, 1, rows.length, 9).setValues(rows);
      seeded.push('Clientes (' + rows.length + ')');
    } catch (err) {
      seeded.push('Clientes: ERROR ' + String(err));
    }
  }

  return { ok: true, created: created, seeded: seeded };
}

/* ══════════════ PLANTILLAS SEED ══════════════
   Variables: {{contact_name}} {{company_name}} {{period}} {{owner_name}} {{firma}} */
const PLANTILLAS_SEED = [
  ['TPL-SP-01','Seguimiento proceso','Status mensual del proceso contable',
   '{{company_name}} | Monthly Accounting Status — {{period}}',
   'Dear {{contact_name}},\n\nI hope you are doing well. I am writing to share the status of {{company_name}}\'s accounting process for {{period}} and the items we need from your side to keep everything on schedule:\n\n• [PENDING ITEM 1]\n• [PENDING ITEM 2]\n\nOnce we receive these, our team will complete the filing within the statutory deadline. Please reply to this email or let me know if you would like a quick call.\n\n{{firma}}'],
  ['TPL-SC-01','Seguimiento Seller Central','Solicitud de accesos/documentos SC',
   '{{company_name}} | Seller Central access & monthly documents — {{period}}',
   'Dear {{contact_name}},\n\nTo prepare {{company_name}}\'s tax filing for {{period}} we need the following from your Amazon Seller Central account:\n\n• View & Edit access for our user marketplaces@tally.legal (Reports: Sales, Inventory, Tax Document Library)\n• Monthly sales report ({{period}})\n• Tax withholding certificate ({{period}})\n• Bank / Payoneer statement ({{period}})\n\nStep-by-step guide: Settings → User Permissions → Add marketplaces@tally.legal → grant View & Edit on Reports.\n\nIf any item is not available yet, just reply and we will guide you.\n\n{{firma}}'],
  ['TPL-CP-01','Cambio de plan','Propuesta de cambio de plan',
   '{{company_name}} | Update to your Tally service plan',
   'Dear {{contact_name}},\n\nBased on {{company_name}}\'s current operation in Mexico, we believe an adjustment to your service plan would serve you better.\n\nWhat changes:\n• [CURRENT PLAN] → [NEW PLAN]\n• [KEY DIFFERENCE / BENEFIT]\n• Effective date: [DATE]\n\nThere is no action needed on your side beyond confirming by reply to this email. Happy to walk you through the details on a call if useful.\n\n{{firma}}'],
  ['TPL-PC-01','Profundización con cliente','Revisión de operación y oportunidades',
   '{{company_name}} | Let\'s review your Mexico operation together',
   'Dear {{contact_name}},\n\nWe have been working with {{company_name}} for some time now, and I would like to schedule a brief session to review how your Mexico operation is performing and where we see opportunities:\n\n• Sales performance and tax efficiency review\n• Marketplace growth opportunities in Mexico\n• Pending structural items (banking, compliance, imports)\n\nWould you have 30 minutes this or next week? Share a couple of time slots and I will send the invite.\n\n{{firma}}'],
  ['TPL-SA-01','Sesión de asesoría','Invitación a sesión de asesoría',
   '{{company_name}} | Advisory session with Tally',
   'Dear {{contact_name}},\n\nAs part of your service with Tally, I would like to offer {{company_name}} an advisory session with our team to address:\n\n• [TOPIC 1]\n• [TOPIC 2]\n\nThe session takes about 45 minutes over video call. Please reply with 2-3 time slots that work for you (CST) and we will confirm the invite.\n\n{{firma}}'],
  ['TPL-AS-01','Aviso de suspensión de actividades','Aviso formal de suspensión',
   '🔴 {{company_name}} | Notice of service suspension',
   'Dear {{contact_name}},\n\nWe are writing to formally notify you that, as of [DATE], Tally will suspend accounting services for {{company_name}} due to [REASON: outstanding balance / prolonged lack of required documentation / client request].\n\nWhat this means:\n• Tax filings after [LAST PERIOD] will not be prepared or submitted by Tally.\n• Statutory obligations with SAT remain the company\'s responsibility; non-filing may generate fines and surcharges.\n• Your file and working papers remain available for handover upon request.\n\nIf you wish to regularize and resume the service, reply to this email before [DEADLINE] and we will send the steps.\n\n{{firma}}']
];

/* ══════════════ HELPERS ══════════════ */

const EMAIL_COL = { estado: 13, prompt_juan: 14, draft_final: 17, ultima_actualizacion: 21 };

function setEmailFields(ss, emailId, fields, now) {
  const sh = ss.getSheetByName('Emails');
  const row = findRow(sh, 1, emailId);
  if (!row) return { ok: false, error: 'email_id no encontrado' };
  Object.keys(fields).forEach(k => { if (EMAIL_COL[k]) sh.getRange(row, EMAIL_COL[k]).setValue(fields[k]); });
  sh.getRange(row, EMAIL_COL.ultima_actualizacion).setValue(now);
  return { ok: true, estado: fields.estado || '' };
}

/** REGLA DURA DE COPIAS: conserva los CC originales y agrega SIEMPRE
 *  customersuccess@ y accounting@ si no están. Excluye al destinatario y a la CUENTA REMITENTE
 *  (para que el correo no llegue a la bandeja de entrada de quien lo envía). */
/* REGLA DURA (20-jul-2026): los correos en `cc_email` de la ficha del cliente van SIEMPRE
 * en copia en cualquier comunicación que salga de la interfaz hacia ese cliente. */
function ccCliente(ss, companyId) {
  if (!companyId) return '';
  const sh = ss.getSheetByName('Clientes');
  if (!sh) return '';
  const row = findRow(sh, 1, companyId);
  if (!row) return '';
  return String(sh.getRange(row, 8).getValue() || ''); // H = cc_email
}
function mergeCc(ccOriginal, to, sender) {
  const OBLIGATORIOS = ['customersuccess@tally.legal', 'accounting@tally.legal'];
  const EXCLUIR = [String(to || '').toLowerCase(), String(sender || 'juan@tally.legal').toLowerCase()];
  const set = [];
  const found = String(ccOriginal || '').match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/g) || [];
  found.concat(OBLIGATORIOS).forEach(function(e) {
    e = e.toLowerCase();
    if (EXCLUIR.indexOf(e) < 0 && set.indexOf(e) < 0) set.push(e);
  });
  return set.join(', ');
}

/* ══════════ ENVÍO MULTI-CUENTA (Service Account + Domain-Wide Delegation) ══════════
   Requiere UNA VEZ: Propiedades del script → GOOGLE_SA_KEY = contenido completo del
   credentials.json del Service Account (el mismo de tally-gmail-mcp). */

function dwdToken(userEmail, scope) {
  try {
    const key = JSON.parse(PropertiesService.getScriptProperties().getProperty('GOOGLE_SA_KEY') || '{}');
    if (!key.client_email || !key.private_key) return null;
    const now = Math.floor(Date.now() / 1000);
    const header = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
    const claim = Utilities.base64EncodeWebSafe(JSON.stringify({
      iss: key.client_email, sub: userEmail,
      scope: scope || 'https://www.googleapis.com/auth/gmail.send',
      aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600
    }));
    const sig = Utilities.computeRsaSha256Signature(header + '.' + claim, key.private_key);
    const jwt = header + '.' + claim + '.' + Utilities.base64EncodeWebSafe(sig);
    const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
      method: 'post', muteHttpExceptions: true,
      payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt }
    });
    if (resp.getResponseCode() !== 200) return null;
    return JSON.parse(resp.getContentText()).access_token;
  } catch (e) { return null; }
}

/** Lee Message-ID/References/Subject del mensaje original (DWD readonly en la cuenta que lo recibió).
 *  Es la llave del hilo RFC-2822: con In-Reply-To/References el correo se agrupa en el hilo
 *  del CLIENTE en cualquier cliente de correo (Gmail, Outlook, Apple Mail). */
function hiloHeaders(cuentaEmail, gmailMsgId) {
  try {
    if (!cuentaEmail || !gmailMsgId) return null;
    const token = dwdToken(cuentaEmail, 'https://www.googleapis.com/auth/gmail.readonly');
    if (!token) return null;
    const r = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(cuentaEmail)
      + '/messages/' + encodeURIComponent(gmailMsgId)
      + '?format=metadata&metadataHeaders=Message-ID&metadataHeaders=References&metadataHeaders=Subject',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return null;
    const d = JSON.parse(r.getContentText());
    const h = {};
    ((d.payload && d.payload.headers) || []).forEach(function(x){ h[String(x.name).toLowerCase()] = x.value; });
    if (!h['message-id']) return null;
    return { inReplyTo: h['message-id'], references: ((h['references'] || '') + ' ' + h['message-id']).trim(), subject: h['subject'] || '' };
  } catch (e) { return null; }
}

function sendViaDwd(from, to, cc, subject, bodyText, threadId, adjuntos, hilo) {
  const token = dwdToken(from);
  if (!token) return { ok: false, error: 'envío como ' + from + ' no disponible: falta GOOGLE_SA_KEY en Propiedades del script o DWD sin permiso' };
  let mime = 'From: ' + from + '\r\nTo: ' + to + '\r\n';
  if (cc) mime += 'Cc: ' + cc + '\r\n';
  if (hilo && hilo.inReplyTo) {
    mime += 'In-Reply-To: ' + hilo.inReplyTo + '\r\n';
    mime += 'References: ' + (hilo.references || hilo.inReplyTo) + '\r\n';
  }
  mime += 'Subject: =?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=\r\nMIME-Version: 1.0\r\n';
  if (adjuntos && adjuntos.length) {
    const b = 'tallyops' + Date.now();
    mime += 'Content-Type: multipart/mixed; boundary="' + b + '"\r\n\r\n';
    mime += '--' + b + '\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n';
    mime += Utilities.base64Encode(bodyText, Utilities.Charset.UTF_8) + '\r\n';
    adjuntos.forEach(function(a){
      mime += '--' + b + '\r\nContent-Type: application/octet-stream; name="' + a.nombre + '"\r\n';
      mime += 'Content-Disposition: attachment; filename="' + a.nombre + '"\r\nContent-Transfer-Encoding: base64\r\n\r\n';
      mime += Utilities.base64Encode(a.blob.getBytes()) + '\r\n';
    });
    mime += '--' + b + '--';
  } else {
    mime += 'Content-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n';
    mime += Utilities.base64Encode(bodyText, Utilities.Charset.UTF_8);
  }
  const payload = { raw: Utilities.base64EncodeWebSafe(mime) };
  if (threadId) payload.threadId = threadId;
  const r = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(from) + '/messages/send', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(payload)
  });
  if (r.getResponseCode() !== 200) return { ok: false, error: 'Gmail API ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 160) };
  return { ok: true };
}

/** Busca el documento MÁS RECIENTE de un cliente en una tabla de AppSheet (Accounting_DataModel)
 *  y lo trae del Drive por nombre de archivo. Columnas de archivo típicas:
 *  declaracion_periodo.Documento · Reportes_de_venta.Archivo · Retenciones_por_periodo.URLRetencion
 *  Estados_cuenta.UrlVentas · Inventario_por_periodo.URLInventario · diot_periodo.Documento */
/* ══════════ 📦 DOCUMENTACIÓN MENSUAL — elegibilidad + checklist (blueprint 20-jul-2026) ══════════
   Universo: AZ con First Shipment=Finalizado · CH/IN con RFC válido y banco activo.
   Checklist: sc (solo AZ) · edo (Payoneer Y = auto-cumplido) · fact (Clients_Load.Control_facturas=Sí).
   Escribe SC_Seguimiento en UN solo setValues (preserva escenario/aprobación/envíos). */
function syncDocumentacion(ss) {
  const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const hoy = new Date(), mesNombre = MESES[hoy.getMonth()], anio = String(hoy.getFullYear());
  const periodo = mesNombre + ' ' + anio;
  const dm = SpreadsheetApp.openById(DATAMODEL_ID);

  const cl = dm.getSheetByName('Clients_Load').getDataRange().getValues();
  const H = cl[0].map(String);
  const iId = H.indexOf('Company_Id'), iRfc = H.indexOf('RFC'), iName = H.indexOf('ClientName'),
        iOwner = H.indexOf('Owner'), iSusp = H.indexOf('Suspension'), iCtrl = H.indexOf('Control_facturas'),
        iFS = H.indexOf('First Shipment'), iFiel = H.indexOf('Cita FIEL');

  const ac = dm.getSheetByName('Accesos_SellerCentral').getDataRange().getValues();
  const hA = ac[0].map(String), aId = hA.indexOf('Company_id'), aEst = hA.indexOf('EstadoAcceso'), aPay = hA.indexOf('Acceso_Payoneer');
  const ACC = {};
  for (let i = 1; i < ac.length; i++) { const k = String(ac[i][aId] || '').trim(); if (k) ACC[k] = { estado: String(ac[i][aEst] || ''), pay: String(ac[i][aPay] || '') }; }

  const ec = dm.getSheetByName('Estados_cuenta').getDataRange().getValues();
  const hE = ec[0].map(String), eId = hE.indexOf('Company_id'), eMes = hE.indexOf('MesPeriodo'), eAnio = hE.indexOf('AñoPeríodo');
  const EDO = {};
  for (let i = 1; i < ec.length; i++) {
    if (String(ec[i][eMes] || '') === mesNombre && String(ec[i][eAnio] || '') === anio) {
      const m = String(ec[i][eId] || '').match(/([A-Z]{2,4}\d{6})/);
      if (m) EDO[m[1]] = true;
    }
  }

  const sh = ss.getSheetByName('SC_Seguimiento');
  const sc = sh.getDataRange().getValues();
  const W = 17; // A..Q (P tipo_perfil, Q checklist)
  sc.forEach(function(r){ while (r.length < W) r.push(''); });
  sc[0][15] = 'tipo_perfil'; sc[0][16] = 'checklist';
  const IDX = {};
  for (let r = 1; r < sc.length; r++) IDX[String(sc[r][0]).trim()] = r;

  let conPend = 0, completos = 0, fuera = 0;
  for (let i = 1; i < cl.length; i++) {
    const cid = String(cl[i][iId] || '').trim();
    if (!cid) continue;
    if (/^s/i.test(String(cl[i][iSusp] || ''))) continue; // suspendidos fuera
    let tipo = '';
    if (cid.indexOf('AZ') === 0) tipo = 'AZ';
    else if (cid.indexOf('CH') === 0) tipo = 'CH';
    else if (cid.indexOf('IN') === 0) tipo = 'IN';
    else { fuera++; continue; } // ML/MX fuera de alcance v1
    const rfc = String(cl[i][iRfc] || '').trim();
    const rfcOK = !!(rfc && rfc !== 'NO MATCH');
    const fielOK = /exitosa/i.test(String(cl[i][iFiel] || ''));
    // Elegibilidad: AZ solo con First Shipment=Finalizado. CH/IN entran TODOS (regla Juan 21-jul):
    // RFC y FIEL NO filtran — se muestran como indicadores preventivos en la card.
    if (tipo === 'AZ' && String(cl[i][iFS] || '') !== 'Finalizado') continue;

    const acc = ACC[cid] || { estado: '', pay: '' };
    const payOK = String(acc.pay).toUpperCase() === 'Y';
    const check = [];
    if (tipo !== 'AZ') { check.push({ k: 'rfc', ok: rfcOK, info: true }); check.push({ k: 'fiel', ok: fielOK, info: true }); }
    if (tipo === 'AZ') check.push({ k: 'sc', ok: /complet|total/i.test(acc.estado) });
    check.push(payOK ? { k: 'edo', ok: true, auto: true } : { k: 'edo', ok: !!EDO[cid] });
    if (/^s/i.test(String(cl[i][iCtrl] || ''))) check.push({ k: 'fact', ok: false });
    const pend = check.filter(function(c){ return !c.ok && !c.info; });
    const accion = pend.length
      ? 'Solicitar: ' + pend.map(function(c){ return { sc: 'acceso SC', edo: 'estado de cuenta', fact: 'facturas y gastos' }[c.k]; }).join(', ') + ' (' + periodo + ')'
      : '';
    const chk = JSON.stringify(check);
    const r = IDX[cid];
    if (r !== undefined) {
      sc[r][1] = String(cl[i][iName] || ''); sc[r][2] = String(cl[i][iOwner] || ''); sc[r][3] = periodo;
      sc[r][8] = pend.length ? '' : '✅ Documentación completa';
      sc[r][9] = accion;
      sc[r][14] = new Date().toISOString();
      sc[r][15] = tipo; sc[r][16] = chk;
      // E escenario, F estado, K aprobación, L/M envíos: NO se tocan
    } else if (pend.length || tipo !== 'AZ') {
      // CH/IN entran SIEMPRE al tablero (aunque su checklist esté completo); AZ solo con pendientes
      const fila = [cid, String(cl[i][iName] || ''), String(cl[i][iOwner] || ''), periodo, 'Esc.1', 'En espera', '', 'A', '', accion, 'Pendiente', '', '', '', new Date().toISOString(), tipo, chk];
      sc.push(fila); IDX[cid] = sc.length - 1;
    } else { completos++; continue; }
    if (pend.length) conPend++; else completos++;
  }
  sh.getRange(1, 1, sc.length, W).setValues(sc);
  return { ok: true, periodo: periodo, con_pendientes: conPend, completos: completos, fuera_alcance_v1: fuera };
}

/** Verifica existencia (barato, sin descargar el archivo) de un documento en el expediente. */
function existeDocumento(companyId, tabla, columna) {
  try {
    if (!companyId || !tabla || !columna) return false;
    const dm = hojaDeTabla(tabla);
    if (!dm) return false;
    const data = dm.getDataRange().getValues();
    const colIdx = data[0].map(String).indexOf(columna);
    if (colIdx < 0) return false;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i].join('|').indexOf(companyId) >= 0) {
        const val = String(data[i][colIdx] || '').trim();
        if (val) {
          const fname = val.split('/').pop().trim();
          return fname ? DriveApp.getFilesByName(fname).hasNext() : false;
        }
      }
    }
    return false;
  } catch (e) { return false; }
}

function buscarDocumento(companyId, tabla, columna) {
  try {
    if (!companyId || !tabla || !columna) return null;
    const dm = hojaDeTabla(tabla);
    if (!dm) return null;
    const data = dm.getDataRange().getValues();
    const H = data[0].map(String);
    const colIdx = H.indexOf(columna);
    if (colIdx < 0) return null;
    let ruta = null;
    for (let i = data.length - 1; i >= 1; i--) { // desde abajo = registro más reciente
      if (data[i].join('|').indexOf(companyId) >= 0) {
        const val = String(data[i][colIdx] || '').trim();
        if (val) { ruta = val; break; }
      }
    }
    if (!ruta) return null;
    const fname = ruta.split('/').pop().trim();
    if (!fname) return null;
    const files = DriveApp.getFilesByName(fname);
    if (files.hasNext()) { const f = files.next(); return { blob: f.getBlob(), nombre: fname }; }
    return null;
  } catch (e) { return null; }
}

/** Busca el DATO (texto) más reciente de un cliente en una tabla de AppSheet. */
function buscarDato(companyId, tabla, columna) {
  try {
    const dm = hojaDeTabla(tabla);
    if (!dm) return null;
    const data = dm.getDataRange().getValues();
    const colIdx = data[0].map(String).indexOf(columna);
    if (colIdx < 0) return null;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i].join('|').indexOf(companyId) >= 0) {
        const val = String(data[i][colIdx] || '').trim();
        if (val) return val;
      }
    }
    return null;
  } catch (e) { return null; }
}

/** Email del owner asignado a un cliente (mapa editable en Config.owners_emails). */
function ownerEmail(ss, companyId) {
  try {
    if (!companyId) return '';
    const cl = ss.getSheetByName('Clientes');
    const row = findRow(cl, 1, companyId);
    if (!row) return '';
    const owner = String(cl.getRange(row, 3).getValue() || '').trim().toLowerCase();
    if (!owner) return '';
    const DEFAULTS = { 'eduardo': 'eduardo@tally.legal', 'cristina': 'cristina@tally.legal',
                       'edgar': 'edgar.martinez@tally.legal', 'arturo': 'arturo@tally.legal' };
    const cfg = ss.getSheetByName('Config');
    const r = findRow(cfg, 1, 'owners_emails');
    if (r) {
      const mapa = String(cfg.getRange(r, 2).getValue() || '');
      const m = mapa.split(',').map(function(p){ return p.split(':'); })
        .filter(function(p){ return p.length === 2 && p[0].trim().toLowerCase() === owner; })[0];
      if (m) return m[1].trim().toLowerCase();
    }
    return DEFAULTS[owner] || '';
  } catch (e) { return ''; }
}

/** Cuentas habilitadas para envío (Config.cuentas_mcp la mantiene el agente desde list_accounts). */
function cuentasMcp(ss) {
  const cfg = ss.getSheetByName('Config');
  let lista = '';
  const row = findRow(cfg, 1, 'cuentas_mcp');
  if (row) lista = String(cfg.getRange(row, 2).getValue() || '');
  const arr = (lista.match(/[\w.+-]+@[\w.-]+/g) || []).map(function(e){ return e.toLowerCase(); });
  if (arr.indexOf('accounting@tally.legal') < 0) arr.push('accounting@tally.legal');
  if (arr.indexOf('juan@tally.legal') < 0) arr.push('juan@tally.legal');
  return arr;
}

/** Resuelve el remitente según las reglas de negocio:
 *  - Seller Central → accounting@ por default
 *  - No-admin → SIEMPRE su propia cuenta (o accounting@ si no está en el MCP)
 *  - Admin → puede elegir (from_account); si no elige, su cuenta
 *  - Cualquier remitente fuera del MCP → accounting@ */
function resolveSender(ss, body, categoria) {
  const habilitadas = cuentasMcp(ss);
  const esSC = /seller\s*central/i.test(String(categoria || ''));
  let userEmail = '', admin = false;
  try {
    if (body.auth && body.auth.email && USUARIOS_ID.indexOf('PEGAR') !== 0) {
      const us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios');
      const data = us.getDataRange().getValues();
      const em = String(body.auth.email).trim().toLowerCase();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === em && String(data[i][1]) === String(body.auth.password)) {
          userEmail = em;
          admin = (em === 'juan@tally.legal') || String(data[i][5] || '').toLowerCase() === 'admin';
          break;
        }
      }
    }
  } catch (e) {}
  let sender;
  if (admin) {
    sender = String(body.from_account || '').toLowerCase() || (esSC ? 'accounting@tally.legal' : (userEmail || 'juan@tally.legal'));
  } else {
    sender = esSC ? 'accounting@tally.legal' : (userEmail || 'accounting@tally.legal');
  }
  if (habilitadas.indexOf(sender) < 0) sender = 'accounting@tally.legal';
  return sender;
}

/** Valida credenciales de administrador (para acciones de gestión de usuarios). */
/** Garantiza las columnas extendidas de Tareas: L descripcion, M clientes, N fecha_entrega. */
function ensureTareasCols(sh) {
  if (String(sh.getRange(1, 12).getValue()) !== 'descripcion') {
    sh.getRange(1, 12, 1, 3).setValues([['descripcion', 'clientes', 'fecha_entrega']]);
  }
}

/** Valida credenciales de CUALQUIER usuario activo de la plataforma. Devuelve {ok,email,nombre} o {ok:false}. */
function checkUser(auth) {
  try {
    if (!auth || !auth.email || !auth.password) return { ok: false, error: 'sin credenciales' };
    const us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios');
    const data = us.getDataRange().getValues();
    const email = String(auth.email).trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email && String(data[i][1]) === String(auth.password)) {
        if (String(data[i][3]).toLowerCase() === 'no') return { ok: false, error: 'usuario desactivado' };
        return { ok: true, email: email, nombre: String(data[i][2] || email) };
      }
    }
  } catch (e) {}
  return { ok: false, error: 'credenciales inválidas' };
}

function checkAdmin(auth) {
  try {
    if (!auth || !auth.email || !auth.password) return false;
    const us = SpreadsheetApp.openById(USUARIOS_ID).getSheetByName('Usuarios');
    const data = us.getDataRange().getValues();
    const email = String(auth.email).trim().toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email && String(data[i][1]) === String(auth.password)) {
        return (email === 'juan@tally.legal') || String(data[i][5] || '').toLowerCase() === 'admin';
      }
    }
  } catch (e) {}
  return false;
}

/** Garantiza las columnas extendidas en Emails: V cc_originales, W mensaje_original, X respuesta_os. */
function ensureCcCol(sh) {
  if (!String(sh.getRange(1, 22).getValue()).trim()) sh.getRange(1, 22).setValue('cc_originales');
  if (!String(sh.getRange(1, 23).getValue()).trim()) sh.getRange(1, 23).setValue('mensaje_original');
  if (!String(sh.getRange(1, 24).getValue()).trim()) sh.getRange(1, 24).setValue('respuesta_os');
}

function findRow(sh, keyCol, keyValue) {
  if (!sh || sh.getLastRow() < 2) return null;
  const values = sh.getRange(2, keyCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(keyValue).trim()) return i + 2;
  }
  return null;
}

function getOrCreate(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.__wasCreated = true;
  }
  if (sh.getLastRow() === 0 || !String(sh.getRange(1,1).getValue()).trim()) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

