# Clientes Marketplace — Kanban por estado de declaración (13-oct-2026)

**Dueño:** Juan (CSO) · **Construyó:** Talia · **Dónde vive:** Tally Ops → «Clientes Marketplace» (interfaz) + módulo `cm_*` del Apps Script (backend) + pestañas `cm_*` del Almacén privado.

## Qué es
La lista completa de clientes del padrón del 360, acomodada en cuatro columnas según el estado de sus declaraciones, con filtro por owner contable. Cada card muestra owner contable, KAM, Friendly ID, RFC, estado de la suscripción (Al día / Con deuda / S/D), bitácora del 360, suspensión/churn y si está conectado a Syntage. El popup muestra First Shipment (estado y fecha cuando está Finalizado), ventas y costos reportados del último mes con dato, y ventas y costos acumulados de los últimos 13 meses. Cuando ninguna fuente trae el dato se muestra **S/D (Sin Datos)**, nunca $0.

## Principio de diseño (orden de Juan)
Backend estable y costo-eficiente: **cero tokens de IA en el camino crítico**. Todo se calcula con reglas y queda guardado; la interfaz solo lee la tabla calculada (una llamada). La IA entra únicamente bajo demanda (`cm_clasificar_ia`) para clasificar emisores de facturas que las reglas no reconocen, y su veredicto se cachea en `cm_emisores`.

## Fuentes y qué aporta cada una
| Fuente | Ruta | Aporta | Costo |
|---|---|---|---|
| Tally 360 | `accounting/roster/list` (páginas de 500) | padrón: ID, RFC, nombre, owner contable (+correo), perfil, First Shipment, suspensión, conexión Syntage | 1–3 llamadas |
| Tally 360 | `accounting/owner-audit` | Stripe (Al día / Con deuda), KAM, ciclo de vida (activa / churn / suspensión) | 1 llamada |
| Tally 360 | `accounting/zero-filing/candidates` (páginas de 100) | `has_bank`: cuenta de banco registrada en el 360 | ~6 llamadas |
| Tally 360 | `accounting/global/periods` (una por mes) | ventas y costos del mes (`sales_last_period`, `costs`) | 13 llamadas (2 en modo reciente) |
| Tally 360 → Syntage | `accounting/syntage/overview` (una por empresa conectada) | 12 meses de facturas emitidas/recibidas: cantidad y subtotal | 1 por conectada |
| Tally 360 → Syntage | `accounting/syntage/cfdis` (empresa × mes, solo recibidas) | emisor de cada factura recibida → fees / marketing / otros | 1 por (conectada, mes con recibidas), solo meses sin clasificar + los 2 últimos |
| AppSheet (hoja Accounting_DataModel) | `Clientes_por_periodo.SalesLastPeriod` | ventas declaradas por mes | lectura directa, 0 llamadas |
| AppSheet | `Clients_Load` | respaldo de First Shipment, suscripción, suspensión y owner cuando el 360 no los trae | lectura directa |

**Bitácora del 360:** el apartado de Tally Ops no expone notas/bitácora (contrato 29-sep-2026, §15). La card muestra «S/D (sin ruta en el 360)». Pendiente: pedir la ruta al equipo del 360.

## Reglas del Kanban (gana la primera que se cumple, de arriba hacia abajo)
4. **Declaraciones con operación** — ingresos en ≥ 2 meses **y** costos en ≥ 2 meses, cruzando AppSheet + 360 + Syntage.
3. **Declaraciones con banco** — cuenta de banco en el 360 **y** facturas recibidas visibles en Syntage que no son solo fees (hay pauta/marketing u otros proveedores).
2. **Declaraciones en ceros con banco** — cuenta de banco en el 360 **y** facturas recibidas únicamente de fees de Amazon/Tally **y** registro de ventas/costos presente (no S/D).
1. **Declaraciones en ceros sin banco** — todo lo demás (sin facturas visibles en Syntage o sin registro de ventas/costos).

Cada card guarda el **motivo** en lenguaje humano (se ve en el popup).

### Clasificación de facturas recibidas (sin IA)
- Se excluyen complementos de pago (`cfdi_type = P`) y canceladas.
- `fee`: emisor en `Config.cm_rfc_fees` (lista de RFCs separada por coma, editable desde el Sheet DB) o cuyo nombre contiene «Amazon» o «Tally».
- `marketing`: descripción con publicidad / advertising / sponsored / pauta / marketing / ads / campaña / promoción (manda sobre el emisor).
- `otro`: cualquier otro emisor. Queda en `cm_emisores` con `origen = regla`; una persona puede corregir `clasificacion` en la pestaña (`origen = manual`, nunca se pisa) o pulsar la clasificación con IA (`origen = ia`).

### Ventas y costos
- **Reportados:** último mes con dato. Ventas: AppSheet › 360 › Syntage (emitidas). Costos: 360 › Syntage (recibidas). Se muestra importe, mes y fuente.
- **Acumulados:** suma de los últimos 13 meses con la misma prioridad de fuente por mes.
- **Suscripción:** Stripe vía owner-audit (`paying`, `sub_status`); si el 360 no trae nada, `Clients_Load.SubscriptionType`; si tampoco, S/D.

## Ciclo (máquina de estados `CM_CYCLE`, avanza con `cronSync` cada 15 min)
`padron → audit → banco → periodos360 → syntage → cfdis → appsheet → calcular → listo`
- Presupuesto por tick: 3.8 min; límite del 360 respetado (≤ 55 llamadas por minuto, con pausa automática).
- Arranque automático diario ≥ 04:00 CDMX en modo **reciente** (solo los 2 últimos meses de facturas y de `global/periods`).
- Desde la interfaz: «Recalcular» (reciente) o «Recalcular todo» (13 meses). El primer tramo corre en la misma petición (≤ 60 s) y el resto sigue solo.
- Estado visible en la barra de la sección (etapa, avance, avisos).

## Pestañas del Almacén
- `cm_clientes` — una fila por empresa con todos los campos de la card/popup + `estado_num`, `estado_kanban`, `motivo`.
- `cm_periodos` — empresa × mes: ventas por fuente, costos por fuente, facturas (emitidas/recibidas) y clasificación (fees/marketing/otros, `emisores`).
- `cm_emisores` — catálogo de emisores con su clasificación y origen (config / regla / manual / ia).
- `cm_meta` — `ultimo_calculo`, `entorno_360`.

## Propiedades del script
- `TALLY_OPS_SECRET` — secreto del 360 (lo guarda una persona; el agente nunca lo pide). Mientras no exista, en staging se usa `STAGING_SCRIPT_API_KEY`.
- `TALLY_360_ENV` — `staging` (default) o `production`. Alternativa: `TALLY_360_BASE` con la URL completa del apartado.
- `ANTHROPIC_API_KEY` — solo para `cm_clasificar_ia` (bajo demanda).

## Acciones del backend
`cm_tabla` (lee la tabla), `cm_estado`, `cm_refresh` (auth; `modo: reciente|completo`), `cm_tick`, `cm_periodos` (detalle de una empresa), `cm_emisores`, `cm_clasificar_ia` (auth), `cm_probar_360` (auth). Función de editor: `cmProbarConexion360()`.

## Pendientes
- Secreto de producción (`TALLY_OPS_SECRET` + `TALLY_360_ENV=production`) para pasar del padrón de staging al real.
- Ruta de bitácora en el apartado del 360.
- Si Juan quiere costos declarados desde AppSheet: hoy `Clientes_por_periodo` no tiene columna de costos; los costos salen del 360 y de Syntage.
