# App de piso: Operador / Técnico (Power Apps, SharePoint)

Complemento del Training System (`training-system-powerapps.md`). Es la vista del empleado
en cuarto limpio: simple, táctil, usable con guantes y en tablet o pantalla de línea.

## 1. Usuarios
| Perfil | Puestos Oracle | Qué hace en la app |
|---|---|---|
| Operador | Assembly Operator A / C | Ve su matriz, sus entrenamientos y su avance |
| Técnico | Testing Operator, Production technician | Lo mismo, más pruebas/máquinas y reporte de falla |
| Lead (indirecto) | Assembly Lead A | Registra y valida entrenamientos de su línea |

Identificación: `User().Email` contra la lista Operadores (correo o NumEmpleado). Opción de
login por NumEmpleado + PIN para pantallas compartidas de línea (decidir con TI/Gaddiel).

## 2. Pantallas (máx. 5, botones grandes)
1. **Mi perfil**: nombre, puesto, línea, supervisor, % de cumplimiento de mi puesto.
2. **Mis habilidades**: lista de competencias con estado ✔ Aprobado / ◐ En proceso / ○ Pendiente
   y fecha. Filtro por línea/estación (HP, NPT1, NPT2, TB, Láser, Empaque…).
3. **Mis entrenamientos**: historial con entrenador, resultado y evidencia.
4. **Solicitar entrenamiento**: el operador pide una competencia; llega al Lead/Supervisor.
5. **Reportar** (técnico): máquina descompuesta, calibración vencida, condición insegura,
   faltante de material. Corresponde a funciones de los descriptivos de puesto.

Para el Lead: pantalla extra **Validar entrenamiento** (firma digital = aprobación).

## 3. Datos adicionales (listas)
- **Solicitudes**: Operador, Competencia, Fecha, Estatus (Pendiente / Programado / Hecho).
- **Reportes**: Tipo, Linea, Equipo, Descripcion, Foto, Reporta, Fecha, Estatus.
- Reutiliza Operadores, Competencias y Entrenamientos del Training System.

## 4. Fórmulas base
```powerfx
// App.OnStart
Set(gblOp, LookUp(Operadores, Correo = User().Email));

// Mis habilidades
AddColumns(
  Filter(Competencias, Puesto = gblOp.PuestoObjetivo),
  Estado,
  Coalesce(LookUp(Entrenamientos, Operador.Id = gblOp.ID && Competencia.Id = ID).Resultado.Value,
           "Pendiente"))

// Solicitar entrenamiento
Patch(Solicitudes, Defaults(Solicitudes), {
  Operador: {Id: gblOp.ID, Value: gblOp.Nombre},
  Competencia: {Id: galComp.Selected.ID, Value: galComp.Selected.Title},
  Estatus: {Value: "Pendiente"}, Fecha: Today() })
```
Privacidad: cada operador solo ve sus registros (`Filter(..., Operador.Id = gblOp.ID)`), y además
restringir permisos de lista en SharePoint (no basta ocultar en la app).

## 5. Diseño para piso
- Fuente grande, alto contraste, íconos y colores con texto (no solo color).
- Español; términos de planta tal cual (Work Order, Masters, NPT, HP, TB).
- Uso offline limitado: Power Apps necesita conexión; confirmar Wi-Fi en cuarto limpio.
- Tableta fija por línea o QR para abrir en teléfono personal (confirmar política de TI).

## 6. Fases
1. MVP: Mi perfil + Mis habilidades (solo lectura) con una línea piloto.
2. Registro y validación de entrenamientos por Lead.
3. Solicitudes y reportes de piso.
4. Integración con Power BI (avance por línea) y propuesta de promociones.

## 7. Decisiones pendientes
- ¿Login personal o por pantalla compartida con PIN?
- ¿El técnico es el Testing Operator o también mantenimiento?
- ¿Qué línea será el piloto?
- Mapeo de niveles A/C de Oracle contra B de RH.
