# Blueprint — Documentación mensual (unificación SC + Legacy)

**Decidido con Juan el 20-jul-2026.** Este documento es la fuente de verdad del diseño; la implementación se hace por fases sobre la sección actual de Seller Central.

## Principio
"Seller Central" se convierte en **📦 Documentación mensual**: un solo ciclo mensual (kanban existente: Sin envíos → Solicitud 01-05 → Recordatorio 05-10 → Ceros regla 17) para TODOS los clientes elegibles. Lo que cambia por cliente no es el proceso sino su **checklist de requisitos**, derivado de su perfil.

## Elegibilidad (quién entra al ciclo)
| Tipo | Regla de entrada | Fuente |
|---|---|---|
| AZ (Amazon) | First Shipment = Finalizado (ÚNICA regla) | Clients_Load."First Shipment" |
| CH (China) | TODOS entran (ajuste Juan 21-jul) | — |
| IN (Legacy/intl) | TODOS entran (ajuste Juan 21-jul) | — |

Para CH/IN la card muestra indicadores PREVENTIVOS (no filtran, no entran al correo): 🪪 RFC (Clients_Load.RFC ≠vacío/NO MATCH) y 🔑 FIEL (Clients_Load."Cita FIEL"=Exitosa) — ⚠️ = revisar a quién se envía información antes de aprobar.
Exclusiones: suspendidos, sin operación, y Payoneer ✅ sin pendiente de acceso SC (auto-cumplido: ni aparece).

## Checklist por requisito
| Requisito | Aplica a | Fuente de estado |
|---|---|---|
| 📊 Acceso SC completo (view+edit) | AZ | Accesos_SellerCentral.EstadoAcceso |
| 🏦 Estado de cuenta del mes | AZ sin Payoneer ✅ + CH + IN | Estados_cuenta (recepción del período) · Payoneer ✅ = auto-cumplido |
| 🧾 Facturas y gastos del mes | Clientes marcados | Clients_Load.Control_facturas = Sí (confirmado por Juan 20-jul) |

## Plantillas modulares
Esc.1/2/3 conservan el esqueleto legal (solicitud → recordatorio → ceros+complementaria regla 17) pero dejan de ser "de Amazon": bloque dinámico `[CHECKLIST]` con solo lo que le falta a ese cliente, en su idioma. Un cliente = un correo con todo lo pendiente. Notion sigue siendo fuente de verdad de las plantillas.

## Impacto por componente (orden de implementación)
1. Backend: cálculo de elegibilidad + checklist en el sync (nueva acción/columnas en SC_Seguimiento: tipo_perfil, checklist JSON).
2. Kanban: mini-checklist ✅/⬜ en cards y drawer; filtro por tipo (AZ/CH/IN).
3. Plantillas: bloque [CHECKLIST] en Esc.1/2/3 (Sheet + Notion) y en el envío rápido/batch.
4. Agente (CLAUDE.md FASE 4): reglas de elegibilidad nuevas, población del tablero, respetar auto-cumplidos.
5. Métricas: % completitud documental del mes, aging por owner, reincidentes en ceros (candidatos churn/suspensión → conecta con SOP legal).

## Pendientes de Juan
- Crear la columna de control de facturas/gastos donde decida y pasar `tabla.columna`.
- Registrar cualquier columna nueva en el editor de AppSheet (Regenerate structure).
