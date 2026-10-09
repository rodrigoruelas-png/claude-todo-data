# Training System TESCOM (Cuarto Limpio) — Power Apps + SharePoint

Reemplaza el Excel de Ingeniería (matriz de habilidades) con una app que registra
entrenamientos como evidencia para promociones (p. ej. Operador C → B).
Sustituye al borrador Kanban anterior.

## 1. Fuentes
- Descripciones de puesto (RH, rev. Abril 2022): Operador de Ensamble B e Inspector B Proceso
  → competencias extraídas en `data/puestos_competencias.csv`.
- Conversación con Gaddiel: 14 posiciones (HP, NPT1, NPT2, TB, Manual, Electrónicos, Kits,
  Láser, Metal Finish, Empaque), 13 personas actuales, propuestas Uriel y Eric → Operador B.
- **Pendiente:** `V:\general data` (Inf. General de operadores). No accesible desde esta sesión;
  hay que exportar a CSV/Excel y subirlo al repo o a SharePoint (ver §6).

## 2. Listas de SharePoint (sitio TESCOM)

**Operadores**: NumEmpleado (clave, texto), Nombre, Puesto (Elección: Operador C / Operador B /
Inspector B / Lead), Linea, FechaIngreso, Turno, Activo (Sí/No), Supervisor (Persona).

**Competencias**: Title, Puesto (a quién aplica), Linea/Estación, Tipo (Rutinaria / Periódica /
Esporádica / Curso), HorasEntrenamiento, RequiereCertificacion.

**Entrenamientos** (registro, evidencia): Operador (Búsqueda), Competencia (Búsqueda), Fecha,
Entrenador (Persona), Resultado (Aprobado / Reprobado / En proceso), Evidencia (adjunto),
Comentarios, AprobadoPor (Persona).

**Posiciones**: Linea, PosicionesRequeridas, NivelRequerido (B/C), Cubiertas (calculado).

**Promociones**: Operador, NivelActual, NivelPropuesto, Justificacion, Estatus
(Borrador / En revisión / Aprobada / Rechazada), FechaSolicitud.

## 3. Roles
| Rol | Permisos |
|---|---|
| Operador | Ve su propia matriz y avance |
| Lead / Entrenador | Registra entrenamientos de su línea |
| Supervisor | Aprueba entrenamientos y propone promociones |
| Ingeniería / Gerencia | Administra competencias, ve todo |

Rol en app: `LookUp(Operadores, Correo = User().Email).Puesto`; filtros por Lead/Supervisor
con grupos de SharePoint (`Office365Groups`) o una lista `Roles`.

## 4. Pantallas
1. **Inicio** – avance por línea, vencimientos.
2. **Matriz de habilidades** – operadores × competencias (galería + íconos ✔/◐/✖).
3. **Perfil del operador** – datos, historial, % cumplimiento de su puesto.
4. **Registrar entrenamiento** – formulario con adjunto de evidencia.
5. **Plantilla y posiciones** – 14 requeridas vs cubiertas por línea.
6. **Propuesta de promoción** – genera justificación con el % de matriz completada.

## 5. Fórmulas clave
```powerfx
// % cumplimiento de un operador
With({req: CountRows(Filter(Competencias, Puesto = gblOp.PuestoObjetivo)),
      ok:  CountRows(Filter(Entrenamientos, Operador.Id = gblOp.ID, Resultado.Value = "Aprobado"))},
     Round(ok / req * 100, 0))

// Registrar entrenamiento
Patch(Entrenamientos, Defaults(Entrenamientos), {
    Operador: {Id: cmbOp.Selected.ID, Value: cmbOp.Selected.Nombre},
    Competencia: {Id: cmbComp.Selected.ID, Value: cmbComp.Selected.Title},
    Fecha: dpFecha.SelectedDate,
    Resultado: ddResultado.Selected });
```
Elegible para promoción: cumplimiento ≥ 100 % de las competencias del nivel B
(ej. Uriel: falta registrar Láser).

## 6. Siguientes pasos
- [ ] Exportar Inf. General de `V:\general data` (operadores TESCOM cuarto limpio) y subir
- [ ] Cargar `puestos_competencias.csv` a la lista Competencias
- [ ] Obtener la matriz de habilidades actual (Excel de Ingeniería) para migrarla
- [ ] Crear listas y app; probar con una línea piloto
- [ ] Documento a Trin con la propuesta Uriel/Eric (usa pantalla 6)

Nota: son datos personales de empleados; limitar el sitio de SharePoint a RH/Ingeniería/Supervisores.
