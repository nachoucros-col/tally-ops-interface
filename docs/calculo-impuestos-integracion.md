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
- **PDF:** el navegador imprime el HTML con la tipografía exacta ("Descargar"); el backend guarda además un PDF de respaldo (HtmlService) para el adjunto automático. Si la calidad del PDF de HtmlService no convence, el adjunto puede ser el HTML o generarse con una función de Netlify (Playwright) — el HTML ya es autocontenido.
- **Lectura A / B** es un selector en el paso 4 y se imprime en el documento. El motor no decide criterio; lo ejecuta.
