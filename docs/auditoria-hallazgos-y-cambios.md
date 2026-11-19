# Auditoría y control de cambios — App Empleados TESCOM + validación de inventario

Estado: **PROPUESTAS PARA TU REVISIÓN. No se aplicó ningún cambio a archivos existentes, ni a
SharePoint, Power Apps, Oracle o Noetix.** Todo lo de abajo es texto/código sin ejecutar.

## 0. Alcance real de esta revisión
| Pude | No pude (sin acceso desde esta sesión) |
|---|---|
| Revisar los archivos del repo (docs, script PowerShell) | Abrir tu app en Power Apps Studio / ambiente de desarrollo |
| Detectar errores por lectura de las fórmulas y del script | Ver SharePoint, Outlook (WIPREG de Yahir), el Excel `MC1 ONHAND Aging Report` |
| Redactar consultas y diseño de gráfica | Conectarme a Oracle EBS o Noetix; ejecutar PowerShell (no hay `pwsh`) |

Por eso **ninguna prueba fue ejecutada**: los hallazgos son por revisión de código, no por
reproducción. Cada propuesta indica cómo probarla en DESARROLLO. Producción: no tocada.

## 1. Hallazgos en la app y el script (ordenados por riesgo)

| # | Riesgo | Hallazgo | Solución propuesta | Estado |
|---|---|---|---|---|
| H1 | Alto | En `docs/app-empleados-nivel.md` las fórmulas comparan columnas **Elección** como texto (`Resultado = "Aprobado"`, `NivelRequerido = …`, `gblYo.Rol = "Supervisor"`). En Power Apps una columna Elección es un registro: hay que usar `.Value`. Resultado probable: cumplimiento siempre 0 % o error. | Cambiar a `Resultado.Value = "Aprobado"`, `NivelRequerido.Value = …`, `gblYo.Rol.Value = …` (y `Coalesce(…NivelPropuesto.Value, …NivelActual.Value)`). | Propuesto |
| H2 | Alto | `GroupBy(…, "Area", …)` se hace sobre ENTRENAMIENTOS, que **no tiene columna Area** (está en COMPETENCIAS). La gráfica de avance por área falla. | Opción A: agregar columna `Area` a ENTRENAMIENTOS al registrar. Opción B: `LookUp` a COMPETENCIAS para obtener el área. Recomendada: A (más simple y rápido). | Propuesto |
| H3 | Alto | `ClearCollect(colEnt, ENTRENAMIENTOS)` carga la lista completa; el límite por defecto es **500 filas** (máx. 2000). Con 13 empleados × ~40 pasos ya se llegan a ~520 registros: el cumplimiento saldría incompleto sin avisar. | Subir el límite a 2000 como parche y, mejor, filtrar por empleado con condiciones delegables (`Filter(ENTRENAMIENTOS, NumEmpleado = …)`), sin colección global. | Propuesto |
| H4 | Alto | La privacidad ("cada operador ve solo lo suyo") está solo en filtros de la app. Quien tenga acceso a la lista en SharePoint ve nombres y números de empleado de todos. | Romper la herencia de permisos del sitio/listas: lectura solo RH/Ingeniería/supervisión; operadores sin acceso directo, la app lee con una cuenta de servicio o flujo. Decisión tuya y de TI. | Propuesto |
| H5 | Alto | "Solo Gaddiel aprueba" se valida en la app. Cualquiera con permiso de edición en SharePoint puede escribir `AprobadoPor` a mano; el cumplimiento y el ascenso se apoyan en ese texto. | Activar historial de versiones en ENTRENAMIENTOS; la aprobación la escribe un flujo de Power Automate con aprobación de Gaddiel (quedan registro y fecha reales); restringir edición de esa lista. | Propuesto |
| H6 | Medio | `crear_listas_sharepoint.ps1` apunta a cualquier sitio que le pases; nada impide correrlo contra **producción** por error. | Agregar parámetro obligatorio `-Ambiente DEV` y abortar si el URL no coincide con el sitio de desarrollo (lista blanca), más `-WhatIf` para ensayo. | Propuesto |
| H7 | Medio | La carga del CSV no es idempotente: correrlo dos veces duplica empleados. | Antes de `Add-PnPListItem`, buscar por `NumEmpleado` y actualizar si existe. | Propuesto |
| H8 | Medio | El script agrega una opción vacía `""` a `NivelPropuesto`; SharePoint suele rechazar opciones vacías y detendría la creación de la lista. | Quitar `""`, dejar la columna sin valor por defecto. | Propuesto |
| H9 | Medio | El mapeo Oracle→nivel pierde información: Testing Operator y Production technician quedan ambos como "Técnico"; Operator A/C no se relaciona con el nivel B de RH. | Guardar siempre `PuestoOracle` original (ya se hace) y definir con Gaddiel la tabla de equivalencias A/C ↔ B. No inferirla. | Pendiente de decisión |
| H10 | Medio | Documentos inconsistentes: `training-system-powerapps.md` y `app-operador-tecnico.md` usan listas "Operadores/Competencias" y columnas distintas a `app-empleados-nivel.md` y a la app real (EMPLEADOS/POSICIONES). Riesgo de implementar el esquema equivocado. | Marcar los dos documentos viejos como reemplazados, o fusionarlos en uno solo con los nombres reales de las listas. | Propuesto |
| H11 | Bajo | La app se creó en formato **teléfono**; la matriz empleados × áreas no cabe. | Tarjeta por empleado con barras por área; pantallas de aprobación/justificación en tableta o PC. | Propuesto |
| H12 | Bajo | La lista de operadores salió de una foto inclinada; pueden existir errores de transcripción en números de empleado. | Contrastar contra la hoja impresa antes de cargar. | Pendiente tuyo |

Nota de seguridad de datos: el CSV con datos de empleados **no está en el repositorio ni en su
historial** (verificado con `git log`); se carga solo desde tu PC.

## 2. Inventario: ítems críticos, P/N sin movimiento y match con WIPREG / Aging

### 2.1 Definiciones que necesito que decidas (no las invento)
- **Sin movimiento:** ¿cuántos días sin transacciones? Propongo 180 como punto de partida.
- **Ítem crítico:** propuesta: (a) existencia < mínimo del almacén (min/max de Gaddiel), o
  (b) requerimiento abierto de WIP (jobs liberados) mayor que la existencia.
- Qué es exactamente WIPREG y qué columnas trae el correo de Yahir, y las columnas del Excel
  `MC1 ONHAND Aging Report`. **No las conozco**; la clave de cruce propuesta es (Organización MC1, P/N).

### 2.2 Automatización de datos (Oracle EBS / Noetix)
- Archivo propuesto: `docs/propuestas/consultas_ebs_solo_lectura.sql` — solo `SELECT`,
  para correr en **ambiente de desarrollo/prueba** con un usuario de solo lectura.
- Usa tablas base de EBS. Las vistas de Noetix dependen de cómo esté instalado en tu empresa; no
  adiviné sus nombres: si prefieres Noetix, indícame las vistas (o pásame el diccionario) y las adapto.
- **No probado** (no hay conexión). Primera prueba sugerida: ejecutar contra DEV, comparar el
  total de existencias con el Aging Report del mismo corte.

### 2.3 Gráfica de nivel de inventario (propuesta)
1. **Barras por P/N:** existencia actual con banda mínimo–máximo; rojo bajo mínimo, ámbar sobre máximo.
2. **Barras apiladas por antigüedad:** 0-30, 31-90, 91-180, >180 días (de fecha de recepción).
3. **Tabla de conciliación** Oracle vs WIPREG vs Aging (clave Org+P/N), con estatus:
   `Coincide` · `Diferencia de cantidad` · `Solo en Oracle` · `Solo en WIPREG/Excel`.
4. Dónde: Power BI si ya lo usan (Gaddiel lo mencionó) o pantalla en Power Apps; en teléfono, una gráfica por pantalla.
La gráfica es un diseño: no hay datos reales cargados, no se hizo ninguna cifra.

## 3. Decisiones que necesito de ti al regresar
Marca qué se implementa: H1–H12 (individual) · definiciones de §2.1 · si el inventario va en
Power BI o Power Apps · cuál es el sitio de DESARROLLO para pruebas y quién da acceso de solo lectura a EBS.
