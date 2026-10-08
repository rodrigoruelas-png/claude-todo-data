# Formato de tareas

Las tareas se representan como un arreglo JSON. Cada elemento es un objeto con
los siguientes campos:

| Campo | Tipo | Descripción |
|---|---|---|
| `id` | cadena | Identificador único de la tarea. |
| `titulo` | cadena | Nombre breve de la tarea. |
| `estado` | cadena | `backlog`, `por-hacer`, `en-progreso`, `en-revision` o `hecho`. |
| `creada` | cadena | Fecha y hora de creación en formato ISO 8601 UTC. |
| `actualizada` | cadena | Fecha y hora de la última actualización en formato ISO 8601 UTC. |

El archivo [`../data/tasks.example.json`](../data/tasks.example.json) contiene
ejemplos de tareas en distintos estados.
