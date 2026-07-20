# Blueprint — Documentación mensual (unificación SC + Legacy)

**Decidido con Juan el 20-jul-2026.** Este documento es la fuente de verdad del diseño; la implementación se hace por fases sobre la sección actual de Seller Central.

## Principio
"Seller Central" se convierte en **📦 Documentación mensual**: un solo ciclo mensual (kanban existente: Sin envíos → Solicitud 01-05 → Recordatorio 05-10 → Ceros regla 17) para TODOS los clientes elegibles. Lo que cambia por cliente no es el proceso sino su **checklist de requisitos**, derivado de su perfil.

## Elegibilidad (quién entra al ciclo)
| Tipo | Regla de entrada | Fuente |
|---|---|---|
| AZ (Amazon) | First_Shipment = Finalizado | Folder_Clientes.First_Shipment |
| CH (China) | RFC válido (≠vacío, ≠"NO MATCH") **y** banco activo | Folder_Clientes.RFC + Bancos |
| IN (Legacy/intl) | misma regla que CH: RFC válido y banco activo | ídem |

Banco activo = fila en `Bancos` con Payoneer ✅ o banco definido en `Otro`.
Exclusiones: suspendidos, sin operación, y Payoneer ✅ sin pendiente de acceso SC (auto-cumplido: ni aparece).

## Checklist por requisito
| Requisito | Aplica a | Fuente de estado |
|---|---|---|
| 📊 Acceso SC completo (view+edit) | AZ | Accesos_SellerCentral.EstadoAcceso |
| 🏦 Estado de cuenta del mes | AZ sin Payoneer ✅ + CH + IN | Estados_cuenta (recepción del período) · Payoneer ✅ = auto-cumplido |
| 🧾 Facturas y gastos del mes | Clientes marcados | Columna que crea Juan (FUERA de Folder_Clientes — pendiente que confirme tabla.columna; el sistema la leerá vía Config `control_facturas_fuente`) |

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
