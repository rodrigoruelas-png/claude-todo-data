# claude-todo-data

Repositorio para guardar tareas en formato JSON. El archivo
[`data/tasks.example.json`](data/tasks.example.json) muestra la estructura inicial
con tres tareas de ejemplo.

## Uso

1. Copia `data/tasks.example.json` como `data/tasks.json`.
2. Agrega, actualiza o elimina tareas en ese archivo.
3. Cada tarea debe incluir `id`, `titulo`, `estado`, `creada` y `actualizada`.
   Las fechas usan el formato ISO 8601 en UTC.
4. `estado` acepta: `backlog`, `por-hacer`, `en-progreso`, `en-revision` o
   `hecho`.

Consulta [la especificación del formato](docs/formato-tareas.md) para más
detalles.