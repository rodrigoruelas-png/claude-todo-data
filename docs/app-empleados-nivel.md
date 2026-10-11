# Power Apps "Empleados – Nivel Operador / Técnico" (TESCOM cuarto limpio)

App de lienzo en SharePoint, modelo del Kanban pero para personal: nivel actual, matriz de
entrenamiento por pasos, firmas, gráficas de efectividad y justificación de ascenso.
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
1. **Inicio** – tarjetas: empleados activos, % cumplimiento promedio, posiciones cubiertas (x/14), pendientes de firma.
2. **Empleados** – galería con foto o iniciales, nombre, puesto, nivel, % cumplimiento; búsqueda y filtro por nivel/línea.
3. **Perfil** – datos, pasos de entrenamiento por área (✔ ◐ ○), firmas y efectividad.
4. **Matriz** – empleados × áreas (HP, NPT1, NPT2, TB, Láser, Empaque…) con % por celda.
5. **Entrenamiento** – registrar paso: fechas, resultado, intentos, entrenador, comentarios.
6. **Firmas** – supervisor e ingeniería autorizan cada paso (ver §5).
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

// % cumplimiento del empleado ThisItem (pasos aprobados y firmados / pasos de su nivel objetivo)
With({req: CountRows(Filter(colComp, NivelRequerido = Coalesce(ThisItem.NivelPropuesto, ThisItem.NivelActual))),
      ok:  CountRows(Filter(colEnt, NumEmpleado = ThisItem.NumEmpleado, Resultado = "Aprobado",
                            !IsBlank(FirmaSupervisor), !IsBlank(FirmaIngenieria)))},
  If(req = 0, 0, Round(ok / req * 100, 0)))

// Iniciales cuando no hay FOTO
Upper(Left(Last(Split(ThisItem.Title, ", ")).Value, 1) & Left(First(Split(ThisItem.Title, ", ")).Value, 1))

// Gráfica de avance por área (Column chart, Items)
AddColumns(GroupBy(Filter(colEnt, NumEmpleado = gblEmp.NumEmpleado, Resultado = "Aprobado"),
    "Area", "g"), "Pasos", CountRows(g))
```

## 5. Firmas (Supervisor / Ingeniería / tú)
Cada paso necesita dos firmas antes de contar al cumplimiento. Firmar = botón visible solo si
`gblYo.Rol` es Supervisor o Ingenieria; guarda nombre y fecha:
```powerfx
Patch(ENTRENAMIENTOS, galPasos.Selected, {
  FirmaSupervisor: If(gblYo.Rol = "Supervisor", User().FullName, galPasos.Selected.FirmaSupervisor),
  FechaFirmaSup:   If(gblYo.Rol = "Supervisor", Now(), galPasos.Selected.FechaFirmaSup),
  FirmaIngenieria: If(gblYo.Rol = "Ingenieria", User().FullName, galPasos.Selected.FirmaIngenieria),
  FechaFirmaIng:   If(gblYo.Rol = "Ingenieria", Now(), galPasos.Selected.FechaFirmaIng) })
```
Quién firma (Jorge, Yahir y después tú) se da de alta en la lista FIRMANTES con su correo y rol,
sin tocar la app. La firma es nombre + fecha de la cuenta de Microsoft 365: sirve como registro
interno; si RH exige firma legal, confirmar con ellos.

## 6. Efectividad (lo que Gaddiel pidió: ascensos con resultados)
Indicadores por empleado y periodo (lista EFECTIVIDAD; se llenan desde Oracle o el Excel
de eficiencia, no se inventan):
- Eficiencia % (Oracle) · Piezas conformes % · Retrabajo % · Estaciones dominadas ·
  Pasos aprobados y firmados · Tiempo promedio por paso (FechaFin − FechaInicio) · Aprobación al primer intento.

Criterio sugerido para proponer ascenso (ajustar con Gaddiel): ≥100 % de pasos del nivel
objetivo aprobados y firmados, eficiencia ≥ meta del área en los últimos 3 periodos, retrabajo
≤ meta, sin reprobados abiertos.
Gráficas: líneas de eficiencia por periodo, barras de pasos por área, medidor de cumplimiento,
barras de posiciones cubiertas.

## 7. Justificación de ascenso (pantalla 9)
Se arma sola: datos del empleado, nivel actual → propuesto, cumplimiento, pasos con fechas y
firmas, gráficas de efectividad y la posición que cubriría. Casos iniciales: Francisco Uriel
(falta registrar Láser) y Erick Francisco (Operador Láser). **Solo se marca "listo" si los números
lo respaldan**; si falta evidencia, la pantalla muestra qué falta.

## 8. Orden de trabajo
1. Correr el script y cargar empleados desde tu CSV.
2. Cargar pasos en COMPETENCIAS desde la matriz de Ingeniería (Excel con pasos y firmas de Jorge y Yahir).
3. Migrar el avance del Excel a ENTRENAMIENTOS (con fechas y firmas).
4. Conectar eficiencia de Oracle a EFECTIVIDAD.
5. Armar pantallas; probar con una línea.

## 9. Pendientes
- Excel de matriz de entrenamiento (pasos + firmas) y la plantilla de Empaque de 6 pasos.
- Fuente y metas de eficiencia/retrabajo por área.
- Apellidos/correos de Jorge y Yahir para FIRMANTES.
- Confirmar si V:\general data / `\\na-mex-fps-01` se puede copiar a SharePoint (acceso bloqueado desde aquí).
