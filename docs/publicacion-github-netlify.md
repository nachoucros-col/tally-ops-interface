# Publicar la interfaz — guía sin conocimientos de git (10 minutos)

No tengo credenciales para crear repos en tu GitHub (y por política no manejo tokens/contraseñas), así que estos clics son tuyos. No necesitas saber git: todo es arrastrar y soltar.

## Opción A — GitHub + Netlify (recomendada: queda versionado y con auto-deploy)

### Paso 1 · Crear el repo (2 min)
1. Entra a **github.com/new** (con la cuenta donde vive tally-dashboards: `nachoucros-col`).
2. Repository name: `tally-ops-interface`.
3. Visibilidad: **Private**.
4. NO marques "Add a README".
5. Click **Create repository**.

### Paso 2 · Subir los archivos (3 min)
1. En la página que aparece, click en el link **"uploading an existing file"**.
2. Abre Finder → `Documents/GitHub/co-cso-personal/tally-ops-interface/`.
3. Selecciona TODO el contenido de la carpeta (index.html, netlify.toml, README.md y las carpetas agente, docs, scripts) y **arrástralo** a la zona de upload de GitHub. *(Arrastrar las carpetas conserva su estructura.)*
4. Abajo escribe el mensaje: `v2 interfaz contable` → click **Commit changes**.

### Paso 3 · Conectar Netlify (5 min)
1. Entra a **app.netlify.com** → **Add new site → Import an existing project → GitHub**.
2. Autoriza y elige el repo `tally-ops-interface`.
3. Build settings: deja todo vacío (no hay build). Publish directory: `.`
4. Click **Deploy**. En ~1 min tienes URL tipo `https://tally-ops-interface.netlify.app`.
5. (Opcional) Site settings → Domain management para ponerle subdominio propio.

Desde entonces, cuando yo actualice archivos tú solo repites el Paso 2 (arrastrar los archivos cambiados) y Netlify redespliega solo.

## Opción B — Netlify Drop (2 minutos, sin GitHub)

Si quieres verla en internet HOY sin crear el repo:
1. Entra a **app.netlify.com/drop**.
2. Arrastra la carpeta `tally-ops-interface` completa (o el archivo `tally-ops-interface-v2.zip` que está en co-cso-personal).
3. Listo — URL pública inmediata. (Sin versionado; el repo lo haces después con la Opción A.)

## Nota sobre la carpeta .git

Dentro de `tally-ops-interface/` quedó una carpeta oculta `.git` de un intento de commit que el entorno no dejó completar. Es inofensiva; GitHub la ignora al subir por web. Si quieres borrarla: Finder → ⌘⇧. (mostrar ocultos) → borrar `.git`.
