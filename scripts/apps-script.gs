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
      const ccDirect = mergeCc(String(body.cc || '') + (ownEm ? ',' + ownEm : ''), String(body.to), senderD);
      if (senderD === 'juan@tally.legal') {
        GmailApp.sendEmail(String(body.to), String(body.subject), String(body.body_text), { cc: ccDirect, name: SENDER_NAME });
      } else {
        const rd = sendViaDwd(senderD, String(body.to), ccDirect, String(body.subject), String(body.body_text), null);
        if (!rd.ok) return rd;
      }
      const sal = getOrCreate(ss, 'Salientes', HEADERS.Salientes);
      const id = 'SAL-' + Date.now() + '-' + (body.company_id || 'X');
      sal.appendRow([id, now, body.company_id || '', body.cliente || '', String(body.to) + (body.cc ? ' cc:' + body.cc : ''),
                     body.categoria || '', body.plantilla || '', String(body.subject), String(body.body_text), 'Enviado', senderD + ' (interfaz)']);
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

      const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: model, max_tokens: 1200, system: system, messages: [{ role: 'user', content: user }] })
      });
      const code = resp.getResponseCode();
      if (code !== 200) return { ok: false, error: 'Claude API ' + code + ': ' + resp.getContentText().slice(0, 200) };
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
      const user = 'CLIENTE DESTINATARIO: ' + (body.cliente || '') + (body.company_id ? ' (' + body.company_id + ')' : '') +
        '\nNOMBRE DEL CONTACTO: ' + (body.contact_name || 'no disponible — usa un saludo genérico profesional') +
        '\n\nOBJETIVO Y CONTEXTO DEL CORREO (instrucción de Juan):\n' + (body.prompt || '') +
        '\n\nFIRMA A USAR AL FINAL:\n' + firma;

      const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: model, max_tokens: 1200, system: system, messages: [{ role: 'user', content: user }] })
      });
      const code = resp.getResponseCode();
      if (code !== 200) return { ok: false, error: 'Claude API ' + code + ': ' + resp.getContentText().slice(0, 200) };
      const blocks = (JSON.parse(resp.getContentText()).content) || [];
      const tb = blocks.filter(function(b){ return b && b.type === 'text' && b.text; })[0];
      if (!tb) return { ok: false, error: 'respuesta sin bloque de texto' };
      const mt2 = String(tb.text).trim().match(/^ASUNTO:\s*(.+)\n+([\s\S]+)$/);
      if (!mt2) return { ok: false, error: 'formato inesperado del modelo: ' + String(tb.text).slice(0, 120) };
      return { ok: true, subject: mt2[1].trim(), body_text: mt2[2].trim() };
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
      const asunto = String(v[14] || ((asuntoOrig.toLowerCase().indexOf('re:') === 0 ? '' : 'Re: ') + asuntoOrig));
      const cuerpo = String(body.draft_final || v[16] || v[15]);
      if (to.indexOf('@') < 0) return { ok: false, error: 'la fila no tiene remitente_email válido' };
      const categoria = String(v[10] || '');
      const cuentaOrigen = String(v[2] || '').toLowerCase();      // cuenta que recibió el correo
      const threadOrigen = String(v[1] || '');                     // thread_id en ESA cuenta
      const sender = resolveSender(ss, body, categoria);
      const cc = mergeCc(String(v[21] || ''), to, sender);

      let enHilo = false, resultado = null;

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
        if (!enHilo) GmailApp.sendEmail(to, asunto, cuerpo, { cc: cc, name: SENDER_NAME });
        resultado = { ok: true };
      } else {
        // vía DWD: si enviamos desde la MISMA cuenta que recibió el correo,
        // el thread_id almacenado es válido → hilo nativo perfecto
        const tid = (sender === cuentaOrigen + '@tally.legal') ? threadOrigen : null;
        resultado = sendViaDwd(sender, to, cc, asunto, cuerpo, tid);
        if (!resultado.ok) return resultado;
        enHilo = !!tid;
      }

      if (body.draft_final) sh.getRange(row, 17).setValue(body.draft_final);
      sh.getRange(row, 13).setValue('Enviado');
      sh.getRange(row, 18).setValue(now);
      sh.getRange(row, 19).setValue('desde: ' + sender);
      sh.getRange(row, 21).setValue(now);
      return { ok: true, enviado_a: to, cc: cc, en_hilo: enHilo, desde: sender };
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
      return { ok: true, threads: threads.length, operaciones_aplicadas: ops, errores: errs.slice(0, 5) };
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

function dwdToken(userEmail) {
  try {
    const key = JSON.parse(PropertiesService.getScriptProperties().getProperty('GOOGLE_SA_KEY') || '{}');
    if (!key.client_email || !key.private_key) return null;
    const now = Math.floor(Date.now() / 1000);
    const header = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
    const claim = Utilities.base64EncodeWebSafe(JSON.stringify({
      iss: key.client_email, sub: userEmail,
      scope: 'https://www.googleapis.com/auth/gmail.send',
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

function sendViaDwd(from, to, cc, subject, bodyText, threadId) {
  const token = dwdToken(from);
  if (!token) return { ok: false, error: 'envío como ' + from + ' no disponible: falta GOOGLE_SA_KEY en Propiedades del script o DWD sin permiso' };
  let mime = 'From: ' + from + '\r\nTo: ' + to + '\r\n';
  if (cc) mime += 'Cc: ' + cc + '\r\n';
  mime += 'Subject: =?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=\r\n';
  mime += 'MIME-Version: 1.0\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n';
  mime += Utilities.base64Encode(bodyText, Utilities.Charset.UTF_8);
  const payload = { raw: Utilities.base64EncodeWebSafe(mime) };
  if (threadId) payload.threadId = threadId;
  const r = UrlFetchApp.fetch('https://gmail.googleapis.com/gmail/v1/users/' + encodeURIComponent(from) + '/messages/send', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }, payload: JSON.stringify(payload)
  });
  if (r.getResponseCode() !== 200) return { ok: false, error: 'Gmail API ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 160) };
  return { ok: true };
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

/** Garantiza las columnas extendidas en Emails: V cc_originales, W mensaje_original. */
function ensureCcCol(sh) {
  if (!String(sh.getRange(1, 22).getValue()).trim()) sh.getRange(1, 22).setValue('cc_originales');
  if (!String(sh.getRange(1, 23).getValue()).trim()) sh.getRange(1, 23).setValue('mensaje_original');
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

