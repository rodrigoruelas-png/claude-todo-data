# Kanban TESCOM en Power Apps — base de avance

> Supuestos (confirmar con Gaddiel): app de lienzo (Canvas), datos en una lista de SharePoint
> (o Dataverse con las mismas columnas), cuatro columnas de estado.

## 1. Modelo de datos

### Lista `Tareas`
| Columna | Tipo | Notas |
|---|---|---|
| Title | Texto | Nombre de la tarea |
| Descripcion | Varias líneas | |
| Estado | Elección | `Por hacer`, `En progreso`, `En revisión`, `Hecho` |
| Prioridad | Elección | `Alta`, `Media`, `Baja` |
| Responsable | Persona | |
| FechaLimite | Fecha | |
| Orden | Número | Posición dentro de la columna |

(Opcional) Lista `Proyectos` y columna de búsqueda `Proyecto` en `Tareas`.

## 2. Pantallas
1. **scrKanban** – tablero con 4 columnas.
2. **scrDetalle** – formulario para ver/editar una tarea.
3. **scrNueva** – formulario para crear tarea.

## 3. Fórmulas Power Fx

**App.OnStart**
```powerfx
Set(varUsuario, User().Email);
ClearCollect(colEstados, ["Por hacer", "En progreso", "En revisión", "Hecho"]);
ClearCollect(colTareas, Tareas);
```

**Galería por columna** (una por estado; ejemplo `galPorHacer`, Items):
```powerfx
SortByColumns(
    Filter(colTareas, Estado.Value = "Por hacer",
           IsBlank(txtBuscar.Text) || txtBuscar.Text in Title),
    "Orden", SortOrder.Ascending
)
```

**Mover tarea a la siguiente columna** (botón › en la plantilla de la galería):
```powerfx
Patch(Tareas, ThisItem,
    { Estado: { Value: Switch(ThisItem.Estado.Value,
        "Por hacer", "En progreso",
        "En progreso", "En revisión",
        "En revisión", "Hecho",
        "Hecho") } });
ClearCollect(colTareas, Tareas)
```

**Crear tarea** (scrNueva, botón Guardar):
```powerfx
Patch(Tareas, Defaults(Tareas), {
    Title: txtTitulo.Text,
    Descripcion: txtDesc.Text,
    Estado: { Value: "Por hacer" },
    Prioridad: ddPrioridad.Selected,
    Responsable: { Claims: "i:0#.f|membership|" & varUsuario,
                   DisplayName: User().FullName, Email: varUsuario },
    FechaLimite: dpLimite.SelectedDate
});
ClearCollect(colTareas, Tareas);
Navigate(scrKanban)
```

**Color por prioridad** (borde de la tarjeta):
```powerfx
Switch(ThisItem.Prioridad.Value, "Alta", Color.Red, "Media", Color.Orange, Color.Green)
```

**Vencida:** `ThisItem.FechaLimite < Today() && ThisItem.Estado.Value <> "Hecho"`

## 4. Buenas prácticas
- Usar colecciones locales y refrescar solo tras `Patch` (evita delegación y lentitud).
- Columnas indexadas en `Estado` para listas > 2000 elementos.
- Permisos: edición para el equipo, lectura para el resto.
- Probar en móvil: layout flexible (contenedores horizontales/verticales).

## 5. Plan de avance
- [ ] Crear la lista `Tareas` con las columnas anteriores
- [ ] Armar scrKanban con 4 galerías y buscador
- [ ] Mover / crear / editar tareas
- [ ] Filtros por responsable y prioridad
- [ ] Pruebas con el equipo y publicación

## 6. Preguntas pendientes para Gaddiel
1. ¿SharePoint o Dataverse? ¿Hay licencia Premium?
2. ¿Columnas de estado definitivas?
3. ¿Roles (admin / miembro) y notificaciones?
4. ¿Drag & drop obligatorio? (Power Apps no lo soporta nativo; se usan botones ‹ ›.)
