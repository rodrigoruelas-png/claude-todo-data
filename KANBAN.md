# Tablero Kanban — claude-todo-data

> Última actualización: 2026-10-08 · Responsables: Rodrigo (dueño), Claude, GitHub Copilot

## Tablero

| 📥 Backlog | 🔜 Por hacer | 🚧 En progreso | 👀 En revisión | ✅ Hecho |
|---|---|---|---|---|
| Definir formato de datos de tareas (JSON/Markdown) | Estructura inicial del repo (carpetas `data/`, `docs/`) | Tablero Kanban + historial (Claude) | | Repo creado (Initial commit) |
| Script para validar tareas | Plantilla de issue "Tarea" | Delegar issue inicial a Copilot | | |
| Automatizar recordatorio semanal (GitHub Action) | README con guía de uso | | | |

## Reparto de trabajo

| Quién | Qué se encarga |
|---|---|
| **Claude** | Planeación, tablero, historial, documentación, revisión de lo que entregue Copilot |
| **GitHub Copilot** | Implementación de issues asignados (abre PRs; Claude/Rodrigo revisan) |
| **Rodrigo** | Prioridades, aprobación y merge de PRs |

## Historial

| Fecha | Evento |
|---|---|
| 2026-10-08 | Repo inicializado (`c8f2724 Initial commit`) |
| 2026-10-08 | Se crea el tablero Kanban y este historial |
| 2026-10-08 | Se crean issues del backlog y se asigna el primero a Copilot |

## 🔔 Recordatorio de actividad

**Dónde vamos:** el repo estaba vacío; ya existe el tablero y el backlog está en issues.

**Siguiente paso inmediato:** revisar el PR que abra Copilot y mergearlo si está bien.

**Rutina sugerida:**
- Al empezar la semana: mover tarjetas de *Backlog* → *Por hacer* (máx. 3 en progreso a la vez).
- Al cerrar cada tarea: agregar una línea al Historial.
- Cada viernes: revisar PRs abiertos y actualizar "Dónde vamos".

**Regla WIP:** no más de 3 tarjetas en *En progreso*.
