#!/bin/bash
# Auto-push de tally-ops-interface → GitHub → Netlify + Apps Script (clasp)
# Corre cada 5 minutos vía LaunchAgent (com.tally.autopush).
# Salida detallada → /tmp/tally-autopush.out (vía launchd)

export PATH="/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"
REPO="/Users/tallylegal/Documents/GitHub/tally-ops-interface"
DEPLOYMENT_ID="AKfycbylh4xs3Ch09rUd05CnfxOE-wgERMCZIW38V-lGIU13DLIaojdynfZlQm8xqV_KLoRY"

echo "═══ $(date '+%Y-%m-%d %H:%M:%S') corrida auto-push ═══"
cd "$REPO" || { echo "❌ no pude entrar al repo"; exit 0; }

# limpiar locks y archivos temporales huérfanos del entorno de Talia
rm -f .git/index.lock .git/HEAD.lock 2>/dev/null
find .git -name "tmp_obj_*" -delete 2>/dev/null
find .git -name "*.lock" -not -path "*/refs/*" -delete 2>/dev/null

# sacar .DS_Store del tracking (una vez; luego es no-op) — Finder los modifica sin parar
git rm -r --cached --ignore-unmatch .DS_Store "*/.DS_Store" >/dev/null 2>&1

# ── 1. GitHub → Netlify ──
CAMBIOS=$(git status --porcelain | wc -l | tr -d ' ')
echo "cambios pendientes: $CAMBIOS"
if [ "$CAMBIOS" != "0" ]; then
  git add -A 2>&1
  git -c user.name="Tally AutoPush" -c user.email="juan@tally.legal" \
      commit -m "auto: actualización $(date '+%Y-%m-%d %H:%M')" 2>&1 | tail -1
fi
echo "push:"
git push origin main 2>&1 | tail -2

# ── 2. Apps Script (clasp) ──
cp scripts/apps-script.gs clasp/Code.js
cp scripts/appsscript.json clasp/appsscript.json 2>/dev/null

if ! cmp -s clasp/Code.js .last-deployed.gs 2>/dev/null; then
  if command -v clasp >/dev/null 2>&1 && ! grep -q "PEGAR_AQUI" .clasp.json; then
    echo "clasp push:"
    PUSH_OUT=$(clasp push -f 2>&1); PUSH_RC=$?
    echo "$PUSH_OUT" | tail -3
    if [ $PUSH_RC -eq 0 ] && ! echo "$PUSH_OUT" | grep -qi "error"; then
      echo "clasp deploy:"
      DEP_OUT=$(clasp deploy -i "$DEPLOYMENT_ID" -d "auto $(date '+%Y-%m-%d %H:%M')" 2>&1); DEP_RC=$?
      echo "$DEP_OUT" | tail -3
      if [ $DEP_RC -eq 0 ] && ! echo "$DEP_OUT" | grep -qi "error"; then
        cp clasp/Code.js .last-deployed.gs
        echo "$(date): Apps Script redesplegado" >> .autopush.log
      else
        echo "$(date): clasp deploy FALLÓ: $(echo "$DEP_OUT" | tail -1)" >> .autopush.log
        echo "⚠️ clasp deploy FALLÓ — se reintentará en la próxima corrida"
      fi
    else
      echo "$(date): clasp push FALLÓ: $(echo "$PUSH_OUT" | tail -1) (¿login vencido? corre: clasp login)" >> .autopush.log
      echo "⚠️ clasp push FALLÓ — se reintentará en la próxima corrida"
      # Alarma visible: notificación de macOS máx. 1 vez por hora mientras siga fallando
      LAST_NOTIF=$(cat /tmp/.tally-clasp-notif 2>/dev/null || echo 0)
      NOW_TS=$(date +%s)
      if [ $((NOW_TS - LAST_NOTIF)) -gt 3600 ]; then
        osascript -e 'display notification "clasp login vencido — el backend de Tally Ops NO se está desplegando. Corre: clasp login" with title "⚠️ Tally AutoPush" sound name "Basso"' 2>/dev/null
        echo "$NOW_TS" > /tmp/.tally-clasp-notif
      fi
    fi
  else
    echo "clasp no disponible o .clasp.json sin scriptId"
  fi
else
  echo "apps-script sin cambios — sin redeploy"
fi
echo "═══ fin corrida ═══"
