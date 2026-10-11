# Power Apps "Empleados – Nivel Operador / Técnico" (TESCOM cuarto limpio)

App de lienzo en SharePoint, modelo del Kanban pero para personal: nivel actual, matriz de
entrenamiento por pasos, aprobación, gráficas de efectividad y justificación de ascenso.
Datos personales: **no van en este repo**; se cargan con `scripts/crear_listas_sharepoint.ps1`.
Si ya existen listas en la app generada (EMPLEADOS, POSICIONES…), el script solo agrega lo faltante.

## 1. Listas
EMPLEADOS · COMPETENCIAS (pasos de entrenamiento por área y nivel) · ENTRENAMIENTOS (evidencia
y firmas) · EFECTIVIDAD (resultados medibles por periodo) · POSICIONES (14) · FIRMANTES.
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
```powerfx
// App.OnStart
Set(gblYo, LookUp(FIRMANTES, Correo = User().Email));  // rol firmante, si aplica
ClearCollect(colEmp, Filter(EMPLEADOS, Activo = true));
ClearCollect(colEnt, ENTRENAMIENTOS);
ClearCollect(colComp, COMPETENCIAS);

// % cumplimiento del empleado ThisItem (pasos aprobados / pasos de su nivel objetivo)
With({req: CountRows(Filter(colComp, NivelRequerido = Coalesce(ThisItem.NivelPropuesto, ThisItem.NivelActual))),
      ok:  CountRows(Filter(colEnt, NumEmpleado = ThisItem.NumEmpleado, Resultado = "Aprobado",
                            !IsBlank(AprobadoPor)))},
  If(req = 0, 0, Round(ok / req * 100, 0)))

// Iniciales cuando no hay FOTO
Upper(Left(Last(Split(ThisItem.Title, ", ")).Value, 1) & Left(First(Split(ThisItem.Title, ", ")).Value, 1))

// Gráfica de avance por área (Column chart, Items)
AddColumns(GroupBy(Filter(colEnt, NumEmpleado = gblEmp.NumEmpleado, Resultado = "Aprobado"),
    "Area", "g"), "Pasos", CountRows(g))
```

## 5. Aprobación (solo Gaddiel)
Un paso cuenta al cumplimiento cuando está `Aprobado` y tiene `AprobadoPor`. Solo Gaddiel
aprueba; su correo se da de alta en la lista FIRMANTES (rol Supervisor). Si después aprueba
alguien más (otra persona o tú), se agrega ahí sin tocar la app.
```powerfx
// Botón Aprobar (visible solo si gblYo no está vacío)
Patch(ENTRENAMIENTOS, galPasos.Selected, {
  AprobadoPor: User().FullName, FechaAprobacion: Now() })
```
Es registro interno por cuenta de Microsoft 365; si RH exige firma legal, confirmarlo con ellos.

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
