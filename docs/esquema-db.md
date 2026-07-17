# Esquema de Base de Datos — Tally Ops Interface

**Google Sheet:** `1A5TSql1ksUHQ8DBYwfTDrj_V3J1HGAs8cgCF9mijmnQ`
**Estado v2:** lectura OK (link público). Escritura vía MCP aún en 403 → el esquema se inicializa con la acción `init_schema` del Apps Script (que corre como juan@, dueño del Sheet). Para que el **agente** pueda escribir en sus corridas, Juan debe agregar como editor la cuenta que aparece como editor "extraño" (service account) en Accounting_DataModel.

El Sheet es la **única fuente de verdad compartida**. La interfaz lee vía GViz y escribe/envía vía Apps Script; el agente lee/escribe vía MCP de Sheets (pendiente permiso).

**v2 — 7 pestañas:** a las 4 originales se suman `Clientes` (directorio para redacción, importado de Accounting_DataModel.Clients_Load), `Plantillas` (plantillas editables por categoría) y `Salientes` (correos enviados desde cero vía interfaz, con seguimiento en Bandeja → "📤 Enviados por mí").

### Pestaña `Clientes`
`company_id · cliente · owner · suspension · subscription · contacto_nombre · contacto_email · cc_email · notas` — contacto_email se importa de stripe_email como default; el equipo lo corrige a mano cuando difiera.

### Pestaña `Plantillas`
`plantilla_id · categoria · nombre · asunto · cuerpo` — categorías: Seguimiento proceso, Seguimiento Seller Central, Cambio de plan, Profundización con cliente, Sesión de asesoría, Aviso de suspensión de actividades. Variables: `{{contact_name}} {{company_name}} {{period}} {{owner_name}} {{firma}}`. Se editan/agregan directo en la pestaña.

### Pestaña `Salientes`
`saliente_id · fecha · company_id · cliente · destinatarios · categoria · plantilla · asunto · cuerpo · estado · enviado_por` — la escribe el Apps Script en cada envío directo.

---

## Pestaña 1: `Emails`

Una fila por correo triaged. El agente inserta; Juan actualiza estado/prompt desde la interfaz; el agente completa drafts y envíos.

| Col | Campo | Escribe | Valores / formato |
|---|---|---|---|
| A | `email_id` | Agente | ID de mensaje Gmail (llave única) |
| B | `thread_id` | Agente | ID de hilo Gmail |
| C | `cuenta` | Agente | `juan` \| `contabilidad` \| `accounting` \| `elizabeth` |
| D | `fecha_recibido` | Agente | ISO `YYYY-MM-DD HH:MM` |
| E | `remitente_nombre` | Agente | texto |
| F | `remitente_email` | Agente | email |
| G | `company_id` | Agente | `AZxxxxxx` si se detecta match en Clients_Load, vacío si no |
| H | `cliente` | Agente | nombre de la empresa cliente (o `INTERNO` si es del equipo) |
| I | `asunto` | Agente | texto |
| J | `resumen` | Agente | 1-3 líneas: qué pide/informa el correo |
| K | `categoria` | Agente | `Documentación` \| `Proceso` \| `Queja y reclamo` \| `Alerta` |
| L | `prioridad` | Agente | `Alta` (Alerta y Queja siempre alta) \| `Media` |
| M | `estado` | Ambos | `Nuevo` → `Prompt recibido` → `Draft listo` → `Aprobado` → `Enviado`; o `Descartado` |
| N | `prompt_juan` | Interfaz | instrucción de Juan para redactar la respuesta |
| O | `draft_asunto` | Agente | asunto propuesto |
| P | `draft_cuerpo` | Agente | cuerpo propuesto (redactado por Talia con base en el prompt) |
| Q | `draft_final` | Interfaz | versión editada/aprobada por Juan (si difiere del draft) |
| R | `fecha_envio` | Agente | ISO al enviarse |
| S | `msg_id_enviado` | Agente | ID Gmail del correo enviado |
| T | `notas_agente` | Agente | contexto extra: hilo previo, adjuntos, referencias AppSheet |
| U | `ultima_actualizacion` | Ambos | ISO |

### Máquina de estados

```
Nuevo ──(Juan escribe prompt)──> Prompt recibido ──(agente redacta)──> Draft listo
Draft listo ──(Juan aprueba)──> Aprobado ──(agente envía desde juan@)──> Enviado
Draft listo ──(Juan pide otra versión con nuevo prompt)──> Prompt recibido
Cualquier estado ──(Juan descarta)──> Descartado
```

**Regla dura:** el agente SOLO envía filas en estado `Aprobado`. La aprobación es explícita, por correo individual, hecha por Juan en la interfaz. Todo envío sale de `juan@tally.legal`.

---

## Pestaña 2: `SC_Seguimiento`

Espejo operativo de la cola Seller Central (SOP Seguimiento Customer Support). El agente sincroniza desde AppSheet (WeeklyPlan owner=AI + Accesos_SellerCentral) en cada corrida; Juan aprueba envíos de escenarios desde la interfaz.

| Col | Campo | Escribe | Valores |
|---|---|---|---|
| A | `company_id` | Agente | AZxxxxxx (llave) |
| B | `cliente` | Agente | nombre |
| C | `owner` | Agente | Owner_vc del cliente |
| D | `periodo` | Agente | ej. `Junio 2026` |
| E | `escenario_actual` | Agente | `Esc.1` \| `Esc.2` \| `Esc.3` |
| F | `estado_sop` | Agente | `En espera` \| `Respondió` \| `Silencio` \| `Excluido` |
| G | `dias_desde_contacto` | Agente | número |
| H | `bloque` | Agente | `A` (listo para envío) \| `B` (discrepancia) |
| I | `motivo_exclusion` | Agente | regla SOP aplicada (suspensión, FS≠Finalizado, acceso total, sin tarea AI) |
| J | `accion_pendiente` | Agente | ej. `Enviar Esc.3 ceros Junio` |
| K | `aprobacion` | Interfaz | `Pendiente` \| `Aprobado` \| `Rechazado` |
| L | `fecha_aprobacion` | Interfaz | ISO |
| M | `fecha_ultimo_envio` | Agente | ISO |
| N | `notas` | Agente | detalle de la última interacción |
| O | `ultima_actualizacion` | Agente | ISO |

**Regla dura:** los correos de escenario (Esc.2/Esc.3) salen de `accounting@tally.legal` CC `customersuccess@tally.legal` (conforme al SOP CX), solo con `aprobacion=Aprobado`.

---

## Pestaña 3: `Config`

Par llave-valor. La interfaz y el agente la leen al inicio de cada ciclo.

| key | value (ejemplo) |
|---|---|
| `cuentas_monitoreadas` | `contabilidad@tally.legal, accounting@tally.legal, elizabeth@tally.legal` (agregar `juan@tally.legal` cuando esté en el MCP) |
| `dominios_excluidos` | `no-reply, noreply, notifications, calendar-notification, mailer-daemon, stripe.com, google.com, appsheet.com, anthropic.com, openai.com, netlify.com, github.com, notion.so, slack.com, intuit.com, docusign` |
| `asuntos_excluidos` | `unsubscribe, invitación:, invitation:, receipt, invoice #, payment received, factura de suscripción, renewal, license` |
| `equipo_interno` | `@tally.legal` (todos los remitentes del dominio = correo interno → categoría según contexto, default Alerta si urgente) |
| `firma_juan` | texto de la firma con la que cierran los correos enviados |
| `version_esquema` | `1.0` |

---

## Pestaña 4: `Log`

Bitácora de corridas del agente. Una fila por corrida.

| Col | Campo |
|---|---|
| A | `timestamp` (ISO) |
| B | `corrida` (`AM` / `PM` / `manual`) |
| C | `correos_revisados` |
| D | `correos_filtrados_fuera` |
| E | `nuevos_triaged` |
| F | `drafts_generados` |
| G | `enviados` |
| H | `sc_actualizados` |
| I | `errores` |
| J | `notas` |

---

## Flujo de datos (quién escribe dónde)

```
Gmail (3 cuentas) ──lectura──> AGENTE ──escritura──> Sheet.Emails
Interfaz (Netlify) ──lectura GViz──> Sheet (todas las pestañas)
Interfaz ──POST Apps Script──> Sheet.Emails (prompt/aprobación) + Sheet.SC_Seguimiento (aprobación)
AGENTE ──lectura──> Sheet (prompts y aprobaciones pendientes)
AGENTE ──redacta draft / envía correo──> Gmail + actualiza Sheet
AppSheet (WeeklyPlan, Accesos_SellerCentral, Clients_Load) <──lectura/escritura── AGENTE ──> Sheet.SC_Seguimiento
```
