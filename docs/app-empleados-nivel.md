# Power Apps "Empleados – Nivel Operador / Técnico" (TESCOM cuarto limpio)

App de lienzo en SharePoint, modelo del Kanban pero para personal: nivel actual, matriz de
entrenamiento por pasos, aprobación, gráficas de efectividad y justificación de ascenso.
Datos personales: **no van en este repo**; se cargan con `scripts/crear_listas_sharepoint.ps1`.
Si ya existen listas en la app generada (EMPLEADOS, POSICIONES…), el script solo agrega lo faltante.

## 1. Listas
EMPLEADOS · COMPETENCIAS (pasos de entrenamiento por área y nivel) · ENTRENAMIENTOS (evidencia
y firmas) · EFECTIVIDAD (resultados medibles por periodo) · POSICIONES (14) · FIRMANTES.
ENTRENAMIENTOS lleva `Area` (texto, copiado de COMPETENCIAS), `EstatusAprobacion` (Pendiente / Aprobado / Rechazado,
escrito solo por el flujo de aprobación), `EvidenciaEmpleado` (lo aporta el propio empleado) y los datos
de trazabilidad de §10.
Relación por `NumEmpleado` (texto) para evitar problemas de delegación con búsquedas.

## 2. Niveles
Oracle: Assembly Operator A / C, Testing Operator, Production technician, Assembly Lead A.
Los descriptivos de RH usan nivel B (Operador de Ensamble B, Inspector B Proceso): confirmar la
regla **A/C Oracle ↔ B RH** con Gaddiel. El script mapea: Operator A→A, C→C, Testing/Production→Técnico.

## 3. Pantallas
1. **Inicio** – tarjetas: empleados activos, % cumplimiento promedio, posiciones cubiertas (x/14), pendientes de aprobación.
2. **Empleados** – galería con foto o iniciales, nombre, puesto, nivel, % cumplimiento; búsqueda y filtro por nivel/línea.
3. **Perfil** – datos, pasos de entrenamiento por área (✔ ◐ ○), aprobación y efectividad.
4. **Matriz** – empleados × áreas (HP, NPT1, NPT2, TB, Láser, Empaque…) con % por celda.
5. **Entrenamiento** – registrar paso: fechas, resultado, intentos, entrenador, comentarios.
6. **Aprobación** – Gaddiel aprueba cada paso (ver §5).
7. **Posiciones** – 14 requeridas vs cubiertas por línea y nivel.
8. **Efectividad** – gráficas por empleado/periodo.
9. **Justificación de ascenso** – resumen imprimible/PDF para Gerencia (Trin).

## 4. Fórmulas clave
Reglas: (1) columnas **Elección** siempre con `.Value`; (2) no cargar ENTRENAMIENTOS completa
(límite de 500 filas por defecto, 2000 máx.): se consulta **por empleado** con filtros delegables.
Parche adicional: Configuración › General › "Límite de filas de datos" = 2000.
```powerfx
// App.OnStart  (solo tablas pequeñas: 13 empleados y el catálogo de competencias)
Set(gblYo, LookUp(FIRMANTES, Correo = User().Email && Activo = true));
ClearCollect(colEmp, Filter(EMPLEADOS, Activo = true));
ClearCollect(colComp, COMPETENCIAS);

// % cumplimiento del empleado ThisItem (pasos aprobados por Gaddiel / pasos de su nivel objetivo)
// ENTRENAMIENTOS se filtra por NumEmpleado y EstatusAprobacion.Value (delegable: ~40 filas por empleado)
With({nivel: Coalesce(ThisItem.NivelPropuesto.Value, ThisItem.NivelActual.Value)},
  With({req: CountRows(Filter(colComp, NivelRequerido.Value = nivel)),
        ok:  CountRows(Filter(ENTRENAMIENTOS, NumEmpleado = ThisItem.NumEmpleado,
                              EstatusAprobacion.Value = "Aprobado"))},
    If(req = 0, 0, Round(ok / req * 100, 0))))

// Iniciales cuando no hay FOTO (Title = "Apellido, Nombre")
Upper(Left(Last(Split(ThisItem.Title, ", ")).Value, 1) & Left(First(Split(ThisItem.Title, ", ")).Value, 1))

// Gráfica de avance por área (Column chart, Items). Area es TEXTO en ENTRENAMIENTOS
AddColumns(
  GroupBy(Filter(ENTRENAMIENTOS, NumEmpleado = gblEmp.NumEmpleado,
                 EstatusAprobacion.Value = "Aprobado"), "Area", "g"),
  "Pasos", CountRows(g))

// Registrar un paso (entrenador o Lead): llena Area al crear
Patch(ENTRENAMIENTOS, Defaults(ENTRENAMIENTOS), {
  Title: cmbComp.Selected.Title, NumEmpleado: gblEmp.NumEmpleado,
  Competencia: cmbComp.Selected.Title, Area: cmbComp.Selected.Area.Value,
  Resultado: ddResultado.Selected, EstatusAprobacion: {Value: "Pendiente"},
  Entrenador: User().FullName, FechaInicio: dpInicio.SelectedDate })

// El propio empleado aporta evidencia (queda Pendiente hasta que Gaddiel apruebe)
Patch(ENTRENAMIENTOS, Defaults(ENTRENAMIENTOS), {
  Title: cmbComp.Selected.Title, NumEmpleado: gblEmp.NumEmpleado,
  Competencia: cmbComp.Selected.Title, Area: cmbComp.Selected.Area.Value,
  Resultado: {Value: "En proceso"}, EstatusAprobacion: {Value: "Pendiente"},
  EvidenciaEmpleado: txtEvidencia.Text })
```
Notas de prueba (en DESARROLLO): el icono amarillo de delegación no debe aparecer en esas
fórmulas; con >500 pasos de prueba el % debe seguir correcto. `gblEmp` es el empleado abierto en Perfil.

## 5. Aprobación (solo Gaddiel) — con flujo de Power Automate
La app **ya no escribe** la aprobación. Un paso cuenta al cumplimiento solo cuando
`EstatusAprobacion = Aprobado`, y ese valor solo lo escribe el flujo tras la respuesta de Gaddiel.
Flujo (propuesto, sin probar; crearlo primero en el ambiente de DESARROLLO):
1. Disparador: "Cuando se crea o modifica un elemento" en ENTRENAMIENTOS.
2. Condición: `Resultado = Aprobado` **y** `EstatusAprobacion = Pendiente` (evita ciclos).
3. "Iniciar y esperar una aprobación" asignada al correo de Gaddiel (lista FIRMANTES, rol Supervisor);
   el detalle incluye empleado, paso, área y evidencia del empleado.
4. Si aprueba: actualizar elemento → `EstatusAprobacion=Aprobado`, `AprobadoPor`=nombre del aprobador,
   `FechaAprobacion`=`utcNow()`. Si rechaza: `Rechazado` + comentario en `Comentarios`.
5. Es un registro interno con cuenta M365 y fecha reales; si RH exige firma legal, confirmarlo con ellos.
Con esto un editor de SharePoint ya no "aprueba a mano" sin dejar rastro (ver §10: historial de versiones).

## 6. Efectividad (lo que Gaddiel pidió: ascensos con resultados)
Indicadores por empleado y periodo (lista EFECTIVIDAD; se llenan desde Oracle o el Excel
de eficiencia, no se inventan):
- Eficiencia % (Oracle) · Piezas conformes % · Retrabajo % · Estaciones dominadas ·
  Pasos aprobados · Tiempo promedio por paso (FechaFin − FechaInicio) · Aprobación al primer intento.

Criterio de ascenso: lo define Gaddiel. Base propuesta: 100 % de pasos del nivel objetivo
aprobados por él y sin reprobados abiertos. Las metas numéricas de eficiencia y retrabajo son
las que Gaddiel indique; mientras no estén cargadas, la app muestra los resultados sin
compararlos contra una meta (no se inventan).
Gráficas: líneas de eficiencia por periodo, barras de pasos por área, medidor de cumplimiento,
barras de posiciones cubiertas.

## 7. Justificación de ascenso (pantalla 9)
Se arma sola: datos del empleado, nivel actual → propuesto, cumplimiento, pasos con fechas y
aprobación, gráficas de efectividad y la posición que cubriría. Casos iniciales: Francisco Uriel
(falta registrar Láser) y Erick Francisco (Operador Láser). **Solo se marca "listo" si los números
lo respaldan**; si falta evidencia, la pantalla muestra qué falta.

## 8. Orden de trabajo
1. Correr el script y cargar empleados desde tu CSV.
2. Cargar pasos en COMPETENCIAS desde la matriz de Ingeniería (Excel con los pasos de entrenamiento).
3. Migrar el avance del Excel a ENTRENAMIENTOS (con fechas y aprobación).
4. Conectar eficiencia de Oracle a EFECTIVIDAD.
5. Armar pantallas; probar con una línea.

## 9. Pendientes
- Plantilla de Empaque (6 pasos): falta; Gaddiel la mencionó. La otra sesión ya la cargó en
  SharePoint, pero yo no la tengo: pásame los pasos para incluirla en COMPETENCIAS.
- Excel de matriz de entrenamiento (pasos) para migrar el avance.
- Metas numéricas de eficiencia/retrabajo: las da Gaddiel.
- Correo de Gaddiel para FIRMANTES.
- `V:\general data` / `\\na-mex-fps-01`: sin acceso desde esta sesión.

## 10. Permisos reales y trazabilidad (quién entregó y quién fue responsable)
**Corrección a mi propuesta anterior:** una app de lienzo con el conector de SharePoint usa los
permisos de **cada usuario**, así que no lee "con cuenta de servicio" por sí sola. Opciones (decide tú/TI):
- **A. Permisos por elemento con flujo:** al crear un empleado/entrenamiento, un flujo rompe la
  herencia y deja acceso solo a ese empleado, a Gaddiel y a RH/Ingeniería. Operadores ven solo lo suyo.
- **B. Los operadores no abren la app:** la usan Lead, Gaddiel e Ingeniería; el operador solo aporta
  evidencia por un formulario/flujo. Más simple y seguro.
- **C. Dataverse con seguridad por fila** (requiere licencia Premium).
Base común: lista EMPLEADOS y ENTRENAMIENTOS sin herencia del sitio; lectura/edición solo para grupos
definidos (RH/Ingeniería = lectura, Lead = captura, Gaddiel = administra); **historial de versiones
activo** en ENTRENAMIENTOS.

**Trazabilidad** (todo queda guardado, sin depender de que alguien lo capture):
| Pregunta | Dónde queda |
|---|---|
| ¿Quién entregó/capturó el registro? | Columna del sistema *Creado por* (+ `CapturadoPor` con `User().FullName` desde la app) |
| ¿Quién fue el responsable del entrenamiento? | `Entrenador` |
| ¿Qué aportó el empleado? | `EvidenciaEmpleado` + archivos adjuntos |
| ¿Quién aprobó y cuándo? | `AprobadoPor` + `FechaAprobacion` (escritos por el flujo) |
| ¿Quién cambió qué? | *Modificado por* + historial de versiones |
| ¿Quién cargó los empleados/listas? | `CargadoPor` (script, usuario de Windows) + registro local `.log` |

## 11. Noetix / Oracle EBS: qué responsabilidad usar (según tu menú de Navigator)
No pude abrir Oracle ni Noetix (host interno; tu Edge no está conectado a esta sesión). Esto sale solo
del menú que pegaste; lo que hace cada pantalla hay que confirmarlo abriéndola tú.

| Responsabilidad | Qué muestra el menú | Uso propuesto |
|---|---|---|
| **EMR NOETIX Global, Standard** | Solo "NOETIX Reporting" | **Extracción automática/consultas** (vistas de Noetix). Primera opción. |
| **EMR INV View Only ASCO US FMC** | Transacciones, On-hand/Availability, Items, Costos, Planeación (solo consulta) + Cambiar organización | Existencias por P/N; cambiar a MC1 con *Change Organization*. Segura. |
| **EMR ONT Order Inquiry TES** | View Onhand, On Hand Quantity, View Discrete Jobs, View Material Requirements, View Material Transactions, Item Supply/Demand | Cruce de **requerimientos de WIP vs existencia** (ítems críticos) y último movimiento. Todo "View". |
| EMR CST View Only ASCO US FMC | Discrete Jobs, View Transactions, Item Costs | Costos/valor de inventario sin movimiento (si lo piden). |
| EMR PO / BOM View Only | Compras, listas de materiales | Solo si luego se cruzan compras o BOM. |
| EMR WIP Shopfloor / Supervisor ASCO MC1 | Job/Schedule Details, Material Transactions, **Move Transactions**, "Run All WIP Reports" | Consulta de jobs; contiene transacciones: **no usar para automatizar**. Ahí puede estar el reporte de WIP (confirmar si WIPREG sale de ahí). |
| **EVITAR para automatizar** (pueden registrar transacciones) | EMR INV User ASCO MC1 / AKN (Transactions, Move Orders, Counting, Kanban); EMR INV Stockroom User TES (Subinventory Transfer, Miscellaneous Transaction, Cycle Counts, Move Orders) | Solo operación manual. Un error aquí mueve inventario real. |

Orden sugerido para probar (en DESARROLLO/prueba): 1) Noetix Reporting → localizar las vistas de existencias,
transacciones y requerimientos WIP; 2) validar un P/N contra *View Onhand* de ONT Order Inquiry TES;
3) comparar el total contra `MC1 ONHAND Aging Report` del mismo corte. Las consultas de
`docs/propuestas/consultas_ebs_solo_lectura.sql` se adaptan a las vistas de Noetix cuando las tengas.
