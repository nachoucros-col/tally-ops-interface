# Tally Ops Interface — Interfaz Contable

Interfaz web para el triage y respuesta de correos del área contable + seguimiento del SOP Seller Central. Primera pieza de la plataforma operativa del equipo contable (diseñada para escalar con más módulos).

**SOP completo (funcionamiento back + front en lenguaje no técnico):** ver página Notion "SOP — Interfaz Contabilidad".

## Arquitectura (v1)

```
┌─────────────┐   lectura GViz    ┌──────────────────┐
│  Interfaz    │ ────────────────> │  Google Sheet DB  │ <── lectura/escritura ── AGENTE (Talia,
│  (Netlify)   │                   │  Emails / SC /    │                          2 corridas/día)
│              │ ──POST──────────> │  Config / Log     │                             │
└─────────────┘   Apps Script     └──────────────────┘                             │
                                                            Gmail MCP (4 cuentas) ──┤
                                                            AppSheet MCP ───────────┘
```

- **Motor**: agente Cowork programado (09:00 y 16:00 CDMX). Lee Gmail, filtra, clasifica, escribe al Sheet; procesa prompts de Juan → drafts; envía los aprobados desde juan@; sincroniza la cola Seller Central con AppSheet.
- **Interfaz**: estática (este repo), lee el Sheet vía GViz y escribe vía Apps Script.
- La estructura está preparada para migrar el motor a Netlify Functions (tiempo real) sin cambiar la interfaz ni la DB.

## Estructura del repo

```
index.html            ← la interfaz completa (SPA, un solo archivo)
netlify.toml          ← config Netlify (estático, sin build)
scripts/apps-script.gs← endpoint de escritura (se pega en el Sheet)
agente/CLAUDE.md      ← SOP operativo del agente (filtros, categorías, flujo)
docs/esquema-db.md    ← esquema de las 4 pestañas del Sheet
```

## Checklist de activación (pasos de Juan)

1. **Compartir el Sheet DB** (`1A5TSql1ksUHQ8DBYwfTDrj_V3J1HGAs8cgCF9mijmnQ`):
   - Como **editor** con la cuenta Google que autoriza el MCP de Sheets de Talia.
   - Como **"cualquiera con el link puede VER"** (necesario para que la interfaz lea vía GViz).
   - Avisar a Talia para que cree las 4 pestañas (`Emails`, `SC_Seguimiento`, `Config`, `Log`).
2. **Agregar juan@tally.legal al Gmail MCP** (tally-gmail) — sin esto no se puede leer tu bandeja ni enviar desde tu correo. Mientras tanto el agente opera con contabilidad@, accounting@ y elizabeth@.
3. **Desplegar el Apps Script**: abrir el Sheet → Extensiones → Apps Script → pegar `scripts/apps-script.gs` → cambiar el TOKEN → Implementar como Web App (ejecutar como tú, acceso: cualquiera) → copiar la URL `/exec`.
4. **Configurar `index.html`**: pegar la URL en `APPS_SCRIPT_URL` y el mismo token en `API_TOKEN`.
5. **Crear repo GitHub** `tally-ops-interface` con el contenido de esta carpeta y **conectarlo a Netlify** (mismo flujo que tally-dashboards: New site from Git → publish directory `.`).

## Seguridad

- La interfaz no contiene credenciales de Gmail ni AppSheet — solo el token del Apps Script, que únicamente permite escribir prompts/aprobaciones en el Sheet.
- El agente jamás envía un correo sin `estado=Aprobado` (aprobación explícita de Juan, correo por correo).
- Los envíos a clientes salen siempre de `juan@tally.legal`; los escenarios SOP de `accounting@` CC `customersuccess@`.
