# SOP · Cálculo de impuestos mensual para clientes de marketplace

**Dueño del proceso:** Contabilidad (owner del cliente) · **Responsable del SOP:** Juan (CSO) · **Sistema:** Tally Ops → sección "Cálculo de impuestos" · **Versión:** 1.2 · 4-sep-2026
**Promesa que mueve:** Cálculo aprobado (3 de 5) y, al enviarse, Reporte entregado (5 de 5).

## 1. Propósito y alcance
Producir cada mes, para cada cliente con ventas, un documento de cálculo de IVA e ISR provisional que el cliente entienda y apruebe **antes** de que se presente la declaración. Aplica a personas morales que venden en Amazon México (y, con el mismo flujo, Mercado Libre / Walmart cuando el marketplace entregue certificado de retenciones). No aplica a clientes sin operación (declaración en ceros, otro SOP).

## 2. Principios (reglas duras que gobiernan el proceso)
1. **El cliente aprueba antes de declarar.** Nada se presenta con estado "borrador".
2. **Un dato = una fuente.** Cada número del documento dice de dónde salió (marketplace, SAT, banco). Si una fuente falta, el documento lo dice; no se rellena.
3. **Sin IA en los números.** Los importes los produce el motor determinista; la IA solo redacta notas y responde dudas.
4. **Saldos a favor de IVA se acreditan, nunca se solicitan en devolución.**
5. **Una sola base para CFDI y declaración.** Si el criterio de IVA (Lectura A o B) exige corregir facturas, se corrigen antes de declarar.
6. **Contabilidad no toca banca ni accesos de Seller Central.** Si el bloqueo es de ese tipo, se escala a CX/Juan.
7. **Credenciales del SAT (CIEC) no viajan en texto plano.** Se comparten por liga de un solo uso; el modo en claro solo existe si Juan lo activa en Config.

## 3. Insumos
| Insumo | Quién lo entrega | Cuándo | Sin él… |
|---|---|---|---|
| Resumen mensual de Seller Central (PDF) | Cliente / acceso Tally | día 1–3 del mes siguiente | no hay cálculo |
| Certificado de retenciones del marketplace (CFDI de retención) | Amazon lo emite ~día 6–10 | día 10 | no hay base por orden ni retenciones |
| Estado de cuenta bancario del mes (PDF/CSV) | Cliente | día 5 | no se concilia el depósito; el documento lo marca |
| CFDI emitidos/recibidos y declaraciones | SAT vía Syntage (automático) | al generar | no hay IVA acreditable ni validación; alerta de conexión |
| Coeficiente de utilidad (anual anterior) | Owner, de la anual | una vez al año | ISR queda "pendiente" |

## 4. Procedimiento (lo que hace el owner en Tally Ops)
**Paso 1 · Cliente.** Menú → Cálculo de impuestos → Nuevo cálculo. Escribe el RFC o el nombre; el sistema muestra el padrón y si el RFC está conectado a Syntage. Elige mes/año e idioma del documento (español, inglés o chino).

**Paso 2 · Documentos.** Arrastra el resumen del marketplace, el certificado de retenciones y el estado de cuenta. El sistema los guarda en Drive (carpeta `Tally · Cálculos de impuestos / RFC / AAAA-MM`), los lee y muestra los datos clave (base, órdenes, ISR retenido, ingreso neto, transferencias, abonos bancarios). La lectura tiene dos rutas y la tabla dice cuál se usó: **texto** cuando el PDF trae capa de texto (se lee tal cual, sin OCR, incluso si la fuente del PDF no declara su mapa de caracteres) y **OCR** cuando el PDF es escaneado o su texto no es aprovechable. **Verifica contra el PDF** antes de continuar; si el tipo se detectó mal, corrígelo en el selector.
- **Estados de cuenta:** se reconocen los de Payoneer y los de tabla *fecha · concepto · monto · saldo* (WorldFirst y cualquier banco con esa forma). En estos últimos el signo de cada movimiento **no se adivina por el concepto**: se deduce del salto del saldo renglón por renglón y se cuadra `saldo inicial + abonos − cargos = saldo final`. La tabla marca *cuadra* o *no cuadra*; si no cuadra, no continúes sin revisar el PDF.

**Paso 3 · SAT / Syntage.** El sistema resuelve la entidad en Syntage en este orden: el identificador que el paso 1 ya tenía guardado para ese `company_id`, luego el RFC, luego la razón social normalizada. Ese orden existe porque hay entidades conectadas cuyo catálogo en Syntage **no trae RFC** (se conectaron por CIEC): buscar solo por RFC las reportaba como no conectadas. Cuando el enlace no fue por RFC, la pastilla lo dice ("ligado por identificador guardado" o "por nombre") y avisa si Syntage no reporta RFC o reporta uno distinto al del padrón — eso es un dato a corregir, no un error del cálculo. Enseguida muestra: CFDI emitidos del mes (y su IVA), CFDI recibidos y el IVA acreditable (solo PUE; PPD pendiente aparte), si el certificado de retención ya está en el SAT y qué declaraciones mensuales del año existen.
- **Si el cliente no aparece en Syntage:** el panel dice cuántas entidades del catálogo se revisaron. Ojo con la diferencia: *no está en Syntage* no es lo mismo que *no está conectado al SAT*; un cliente puede estar declarando y no estar dado de alta en Syntage.
- **Si el cliente no está en Syntage:** aparece la alerta *"¿Deseas que Juan apruebe que este cliente se incluya en Syntage?"*. Si el sistema ya tiene la CIEC, no se pide; si no, el owner la captura. Al confirmar, Juan recibe en Slack:
  > 🔌 Solicitud de conexión de cliente a Syntage — [Usuario] ha pedido que agregues a [Empresa] en Syntage. RFC: … · CIEC: 🔒 liga de un solo uso
  Juan conecta al cliente (≈5 min) y el owner pulsa "Reintentar". Se puede continuar sin SAT, y el documento marcará esa fuente como pendiente.

**Paso 4 · Cálculo.** Revisa los insumos lado a lado (marketplace vs SAT/banco), elige el criterio de IVA (Lectura A por defecto para vendedores con ventas transfronterizas), captura el coeficiente de utilidad si lo tienes y pulsa **Calcular y generar documento**. El motor arma la serie enero→mes con los cálculos previos guardados del mismo RFC (arrastra saldo a favor y acumulados de ISR), genera el documento en los tres idiomas y lo guarda con estado **Borrador**.

**Paso 5 · Entrega.** Ver el resumen, **descargar el Excel**, **Carga en Sistema** (escribe la fila en la tabla `calculo_impuestos` de AppSheet por su API y sube el .xlsx a `calculo_impuestos_Files_`; el `PeriodID` se resuelve consultando la app, nunca se construye, y si el período del cliente no está abierto no se escribe nada), pedir corrección (queda en bitácora y avisa por Slack) o **Redactar correo al cliente**: se abre la Bandeja con el cliente elegido, la plantilla saludando a la empresa con el resumen (ventas, IVA, ISR, total, retenciones) y el documento adjunto. Se envía desde la cuenta del owner con copia a customersuccess@.

**Cierre.** Cuando el cliente responde aprobando, el owner cambia el estado a **Aprobado** (mueve la promesa 3), presenta la declaración y cambia a **Declarado**; el acuse se sube a la misma carpeta.

## 5. Criterios de cálculo (lo que hace el motor)
- **IVA:** los importes salen del bloque **Impuestos retenidos** del certificado, no de los totales del pie. En el renglón de IVA, *Base del impuesto* es el **IVA trasladado** y *Importe del impuesto* es el **IVA retenido** por el marketplace. De ahí: base 16% = IVA trasladado / 0.16 (si el residuo de la división es menor a un peso se absorbe en la base 16%); base 0% = monto de operación − base 16%; retenido = el importe del renglón de IVA (la liquidación del marketplace queda como cotejo y, si difiere, el documento lo alerta); acreditable = IVA de CFDI recibidos PUE del mes (notas de crédito restan; PPD entra al pagarse); resultado = trasladado − retenido − acreditable; posición = resultado − saldo a favor anterior; a pagar = máx(0, posición); arrastre = máx(0, −posición).
- **ISR provisional (art. 14 LISR):** ingresos nominales acumulados × coeficiente de utilidad − pérdidas = base; × 30% = ISR acumulado; − retenciones acumuladas del marketplace − pagos provisionales previos = a pagar. Sin coeficiente, "pendiente".
- **Lectura A vs B:** A toma el IVA trasladado del certificado (orden por orden); B toma el IVA de los CFDI emitidos. La elección es de criterio del owner con el contador; el documento declara cuál se usó. Si se elige A y existen CFDI que trasladan 16% sobre todo, hay que sustituirlos antes de declarar (regla 5).
- **Conciliación:** transferencias según marketplace vs abonos identificados en el estado de cuenta; la diferencia se muestra, no se explica automáticamente.

## 6. Formato del documento (Excel de 7 hojas)
Hojas: 1. Resumen ejecutivo (total a pagar, IVA, ISR, saldo a favor, retenciones, ventas, órdenes con/sin IVA, criterio, qué sigue, alertas). 3. Cómo calculamos (cinco pasos + tabla de fuentes con estado). 4. IVA e ISR renglón por renglón con la fuente de cada uno. 5. Detalle mensual de ventas/retenciones y conciliación bancaria. 6. CFDI emitidos y recibidos del mes. 7. Notas, supuestos y glosario.
Mismas cifras en ES/EN/中文. Las hojas de IVA e ISR llevan fórmulas auditables (el cliente puede rastrear cada resultado). El HTML se conserva solo como vista previa en pantalla. Paleta Tally, tipografía Arial para que abra igual en cualquier Excel.

## 7. Roles (RACI)
| Actividad | R | A | C | I |
|---|---|---|---|---|
| Recibir insumos del cliente | Owner contable | Juan | CX | — |
| Fabricar el cálculo en Tally Ops | Owner contable | Juan | contador | Talia |
| Conectar cliente a Syntage | Juan | Juan | — | owner |
| Corregir facturación si el criterio lo exige | Owner + contador | Juan | — | cliente |
| Enviar documento y obtener aprobación | Owner contable | Juan | CX | — |
| Presentar declaración y subir acuse | Owner contable | Juan | — | cliente |
| Mantener motor, formato y SOP | Talia | Juan | equipo | — |

## 8. Controles y tiempos
- Día 10: certificado disponible → cálculo generado a más tardar día 12. Envío al cliente día 12. Recordatorio si no aprueba en 48 h. Declaración antes del 17.
- Control de calidad automático: los totales del resumen del marketplace deben cuadrar con sus componentes (diferencia 0.00); el ISR retenido del certificado del mes M debe aparecer en la liquidación M+1.
- Métricas: % de cálculos generados antes del día 12; % aprobados antes del día 15; correcciones por cálculo; clientes sin Syntage.

## 9. Excepciones
- Cliente sin certificado (marketplace sin retención): se usa la liquidación como base y el documento lo indica.
- Mes sin ventas: no se fabrica cálculo; va al SOP de declaración en ceros.
- Cliente con varias cuentas/marketplaces: un cálculo por RFC y mes; los documentos se suben todos y el motor suma por tipo.
- Datos ilegibles por OCR: el owner captura manualmente los datos clave en la tabla del paso 2 (quedan marcados como "capturado a mano").

## 10. Qué queda pendiente para la versión 1.1
Lectura automática de CSV de Mercado Libre; carga masiva de meses históricos; cálculo de actualización y recargos para extemporáneas; firma de aprobación del cliente desde el correo (botón "Apruebo").

## 11. Requisitos técnicos de la carga al sistema
- `APPSHEET_ACCESS_KEY` en Propiedades del script (AppSheet → Settings → Integrations → Enable API → Application Access Key). Sin ella el botón devuelve el aviso y no escribe nada.
- La escritura es SIEMPRE por la API de AppSheet: respeta el Ref a `Clientes_por_periodo`, el Enum de `MesPeriodo`, la columna File y las automatizaciones de la app. Nunca se escribe en el Google Sheet.
- El botón es idempotente: si ya hay cálculo para ese cliente-período, edita esa fila en vez de duplicarla.
