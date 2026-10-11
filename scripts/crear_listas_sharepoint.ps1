<#
 Crea las listas SharePoint de la app "Empleados nivel operador/técnico" (TESCOM).
 NO PROBADO desde la sesión de Claude (no hay acceso a tu tenant). Ejecutar en tu PC:
   Install-Module PnP.PowerShell -Scope CurrentUser
   .\crear_listas_sharepoint.ps1 -SiteUrl "https://TU-TENANT.sharepoint.com/sites/TU-SITIO" `
        -CsvEmpleados "C:\ruta\operadores.csv"
 Si alguna lista ya existe (EMPLEADOS, POSICIONES...), el script la reutiliza y solo agrega columnas faltantes.
#>
param(
  [Parameter(Mandatory)] [string] $SiteUrl,
  [string] $CsvEmpleados
)

Connect-PnPOnline -Url $SiteUrl -Interactive

function Ensure-List($title) {
  if (-not (Get-PnPList -Identity $title -ErrorAction SilentlyContinue)) {
    New-PnPList -Title $title -Template GenericList -OnQuickLaunch | Out-Null
  }
}
function Ensure-Field($list, $name, $type, $choices) {
  if (Get-PnPField -List $list -Identity $name -ErrorAction SilentlyContinue) { return }
  if ($choices) { Add-PnPField -List $list -DisplayName $name -InternalName $name -Type Choice -Choices $choices -AddToDefaultView | Out-Null }
  else          { Add-PnPField -List $list -DisplayName $name -InternalName $name -Type $type -AddToDefaultView | Out-Null }
}

# --- EMPLEADOS ---
Ensure-List "EMPLEADOS"
Ensure-Field "EMPLEADOS" "NumEmpleado"   Text
Ensure-Field "EMPLEADOS" "PuestoOracle"  Choice @("Assembly Operator A","Assembly Operator C","Testing Operator","Production technician","Assembly Lead A")
Ensure-Field "EMPLEADOS" "NivelActual"   Choice @("A","B","C","Tecnico","Lead")
Ensure-Field "EMPLEADOS" "NivelPropuesto" Choice @("","A","B","C","Tecnico","Lead")
Ensure-Field "EMPLEADOS" "JobGroup"      Text
Ensure-Field "EMPLEADOS" "Supervisor"    Text
Ensure-Field "EMPLEADOS" "Grupo"         Choice @("Ensamble","Indirectos")
Ensure-Field "EMPLEADOS" "Linea"         Text
Ensure-Field "EMPLEADOS" "Correo"        Text
Ensure-Field "EMPLEADOS" "Notas"         Note
Ensure-Field "EMPLEADOS" "Activo"        Boolean

# --- COMPETENCIAS (catálogo de pasos de entrenamiento) ---
Ensure-List "COMPETENCIAS"
Ensure-Field "COMPETENCIAS" "Area"            Choice @("HP","NPT1","NPT2","TB","Manual","Electronicos","Kits","Laser","Metal Finish","Empaque","Testing")
Ensure-Field "COMPETENCIAS" "NivelRequerido"  Choice @("A","B","C","Tecnico")
Ensure-Field "COMPETENCIAS" "Paso"            Number
Ensure-Field "COMPETENCIAS" "Tipo"            Choice @("Rutinaria","Periodica","Esporadica","Curso")
Ensure-Field "COMPETENCIAS" "Criterio"        Note

# --- ENTRENAMIENTOS (evidencia por empleado y paso) ---
Ensure-List "ENTRENAMIENTOS"
Ensure-Field "ENTRENAMIENTOS" "NumEmpleado"   Text
Ensure-Field "ENTRENAMIENTOS" "Competencia"   Text
Ensure-Field "ENTRENAMIENTOS" "FechaInicio"   DateTime
Ensure-Field "ENTRENAMIENTOS" "FechaFin"      DateTime
Ensure-Field "ENTRENAMIENTOS" "Resultado"     Choice @("En proceso","Aprobado","Reprobado")
Ensure-Field "ENTRENAMIENTOS" "Intentos"      Number
Ensure-Field "ENTRENAMIENTOS" "Entrenador"    Text
Ensure-Field "ENTRENAMIENTOS" "AprobadoPor"     Text      # Gaddiel (único aprobador)
Ensure-Field "ENTRENAMIENTOS" "FechaAprobacion" DateTime
Ensure-Field "ENTRENAMIENTOS" "Comentarios"   Note

# --- EFECTIVIDAD (resultados medibles por empleado y periodo) ---
Ensure-List "EFECTIVIDAD"
Ensure-Field "EFECTIVIDAD" "NumEmpleado"        Text
Ensure-Field "EFECTIVIDAD" "Periodo"            Text      # ej. 2026-09
Ensure-Field "EFECTIVIDAD" "EficienciaPct"      Number    # desde Oracle / Excel query
Ensure-Field "EFECTIVIDAD" "PiezasConformesPct" Number
Ensure-Field "EFECTIVIDAD" "RetrabajoPct"       Number
Ensure-Field "EFECTIVIDAD" "EstacionesDominadas" Number
Ensure-Field "EFECTIVIDAD" "Fuente"             Text

# --- POSICIONES (14 requeridas) ---
Ensure-List "POSICIONES"
Ensure-Field "POSICIONES" "Linea"               Choice @("HP","NPT1","NPT2","TB","Manual","Electronicos","Kits","Laser","Metal Finish","Empaque")
Ensure-Field "POSICIONES" "NivelRequerido"      Choice @("A","B","C","Tecnico")
Ensure-Field "POSICIONES" "NumEmpleadoAsignado" Text

# --- FIRMANTES (quién aprueba; inicialmente solo Gaddiel) ---
Ensure-List "FIRMANTES"
Ensure-Field "FIRMANTES" "Correo" Text
Ensure-Field "FIRMANTES" "Rol"    Choice @("Supervisor","Gerencia")
Ensure-Field "FIRMANTES" "Activo" Boolean

# --- Carga de empleados desde CSV (opcional; el CSV queda en tu PC, no en git) ---
if ($CsvEmpleados -and (Test-Path $CsvEmpleados)) {
  Import-Csv $CsvEmpleados -Encoding UTF8 | ForEach-Object {
    $nivel = switch -Wildcard ($_.Puesto_Oracle) {
      "*Operator A" { "A" } "*Operator C" { "C" } "Testing*" { "Tecnico" }
      "Production*" { "Tecnico" } "*Lead*" { "Lead" } default { "C" } }
    Add-PnPListItem -List "EMPLEADOS" -Values @{
      Title = $_.Nombre; NumEmpleado = $_.NumEmpleado; PuestoOracle = $_.Puesto_Oracle
      NivelActual = $nivel; JobGroup = $_.JobGroup; Supervisor = $_.Supervisor
      Grupo = $_.Grupo; Notas = $_.Notas; Activo = $true } | Out-Null
  }
  Write-Host "Empleados cargados desde $CsvEmpleados"
}
Write-Host "Listo."
