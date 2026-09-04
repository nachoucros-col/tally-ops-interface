# Módulo "Cálculo de impuestos" para Tally Ops — guía de integración

Paquete construido el 3–4 sep 2026. Todo lo que está aquí funciona sin IA en los números y respeta brand-master §6.

## Contenido
| Archivo | Qué es |
|---|---|
| `formato/calculo-engine.js` | Motor determinista (IVA, ISR provisional, arrastre de saldo a favor, serie enero→mes). Navegador, Node y Apps Script. |
| `formato/calculo-render.js` | Formato del documento (7 páginas) en ES/EN/中文 + plantilla del correo. Mismo archivo para vista previa, PDF y adjunto. |
| `tally-ops/backend-calculo.gs` | Endpoints `calc_*` para el backend vivo (Apps Script): clientes, subida a Drive, OCR + parsers, Syntage por RFC, generar/guardar/listar/corregir, plantilla de correo, solicitud de conexión a Syntage con aviso a Juan. |
| `tally-ops/frontend-calculo.html` | Sección `#view-calculo`: asistente de 5 pasos, arrastrar y soltar, alerta Syntage con CIEC, tabla filtrable (cliente, RFC, mes, año, estado). |
| `tally-ops/bandeja-adjunto.js` | Gancho en "Redactar correo": elegir cliente → período → carga plantilla + adjunto. |
| `sop/SOP-calculo-de-impuestos.md` | SOP del proceso. |
| `demo/` | EuroPartes · julio 2026 generado con el motor y el formato en los tres idiomas (HTML + PDF) y los tres correos. |

## Pasos de integración (repo vivo `tally-ops-interface`, auto-push → Netlify)
1. **Backend (Apps Script del Sheet DB):** crear tres archivos `calculo-engine.gs`, `calculo-render.gs`, `calculo.gs` con el contenido de los tres archivos correspondientes. En `doPost`, antes de `default:` → `if (action.indexOf('calc_') === 0) return calcDispatch_(action, body, user);`. En `send_direct`/`send_reply` agregar `attachments` a partir de `body.adjuntos_drive_ids`. Propiedades nuevas: `SLACK_BOT_TOKEN`, `JUAN_SLACK_ID=U08095WJXTR`, `CALC_DRIVE_ROOT` (opcional), `CIEC_SECRET`, `INTERFACE_URL`. Config: `ciec_en_slack=no`.
2. **Frontend (`index.html`):** enlace en el menú `🧮 Cálculo de impuestos`; pegar la `<section>` + `<style>` + `<script>` de `frontend-calculo.html`; incluir `<script src="calculo-engine.js">` y `<script src="calculo-render.js">` (copiar ambos al repo); pegar `bandeja-adjunto.js` y el bloque `.cx-attach` en el modal de redacción; llamar `calcComposeHook()` al final de `openCompose()`.
3. **Logo:** copiar `tally-logotipo-violet.svg` y `-white.svg` del repo web a `/images/brand/` del repo de Ops y pasar su contenido en `input.logo_svg` (la portada lo usa; si no, wordmark tipográfico).
4. **Permisos:** la vista la ven todos los usuarios; la tabla de cálculos respeta la cartera (no-admin solo ve la suya). "Revelar CIEC" solo admin.
5. **Prueba de aceptación:** EuroPartes EES251001N28 · julio 2026 con los PDF de esta sesión debe dar IVA a pagar 0.00, saldo a favor 10,181.29, ISR "pendiente" (sin coeficiente), 178 órdenes / 18 con IVA. Es exactamente el `demo/`.

## Decisiones de diseño que conviene conocer
- **La serie enero→mes se completa desde los cálculos guardados** del mismo RFC (arrastre de saldo a favor y acumulados de ISR). El primer mes que se fabrique de un cliente debe ser el primero del ejercicio con ventas, o capturar `saldo_favor_inicial`.
- **Syntage se consulta en vivo por RFC** (no desde el almacén nocturno), para que un cliente conectado hace 5 minutos ya aparezca.
- **CIEC:** por defecto viaja como liga de un solo uso que solo un admin puede abrir; la fila cifrada se borra al revelarse. `ciec_en_slack=si` manda la contraseña en claro, tal como lo pidió Juan; queda registrado que es decisión suya. Mejor ruta a futuro: el backend de Tally 360 ya expone `accounting/syntage/connect` y `credentials by company`; conectar desde ahí elimina la CIEC del flujo humano por completo.
- **Excel (v5+):** el entregable es .xlsx generado en Apps Script (hoja temporal → export xlsx → se descarta), bajo demanda y cacheado por idioma en la carpeta del cliente. Descarga por base64 (JSONP), adjunto de correo y carga al sistema usan ese mismo archivo. El PDF se retiró.
- **Lectura A / B** es un selector en el paso 4 y se imprime en el documento. El motor no decide criterio; lo ejecuta.

## Diseño (v6 · 4-sep-2026)
El front porta el sistema visual de `compendio-growve-poc-v1_1.html` (Tally × Apple): tokens en `:root` (grises de sistema `--t1/--t2/--t3`, superficies `--bg/--card/--card-2`, hairlines de .5px, radios 10/16/22/28/980, sombras de doble capa, `--ease` cubic-bezier(.16,1,.3,1)) y una capa `DESIGN SYSTEM · TALLY × APPLE` al final del documento que restila cada vista. Es CSS puro: ningún id, clase o handler cambió. La base tipográfica se queda en 15px (app densa) en vez de los 17px del compendio (documento).

## Lectura de documentos (4-sep-2026, v7)
- `calcTextoNativo_` → `calcPdfTexto_`: lector de la capa de texto del PDF escrito dentro de Apps Script. Incluye `calcInflate_` (DEFLATE propio, sin dependencias), recorrido de objetos, `ToUnicode` cuando existe y, cuando no, un respaldo por desplazamiento de glifos que se elige puntuando el resultado contra un diccionario de términos esperados. Es la ruta A y no consume cuota de Drive.
- `calcTextoDeArchivo_` (OCR de Drive) queda como ruta B y solo corre si de la ruta A no salió ningún dato. `calc_extraer` devuelve `via` para que la interfaz diga por dónde se leyó.
- `calcParseTablaSaldo_`: estados de cuenta con columnas fecha · concepto · monto · saldo. El signo se deduce del salto del saldo y se cuadra contra el saldo final; si la tabla no cuadra por saldo, cae a clasificación por concepto y lo declara en `nota`.
- `calcSynEntidad_`: resolución de entidad en Syntage por `syntage_entity_id` (del front o del almacén), RFC y nombre normalizado (con tolerancia a razón social abreviada). `calc_syntage` devuelve `ligado_por`, `rfc_syntage` y `entidades_revisadas`.
- Caso de prueba verificado: estado de cuenta WorldFirst MXN de INSIGHTCONNECT, julio 2026 → 34 movimientos, abonos 6,539,241.61, cargos 31,298.79, saldo inicial 2,261,026.06, saldo final 8,768,968.88, cuadra. Certificado de retenciones de EuroPartes julio 2026 → base 116,400.19, 178 órdenes, 18 con IVA, ISR 2,910.00, IVA retenido 418.91, IVA trasladado 837.83.
