<#
 Crea las listas SharePoint de la app "Empleados nivel operador/técnico" (TESCOM).
 NO PROBADO desde la sesión de Claude (sin acceso a tu tenant ni a PowerShell). Probar primero con -WhatIf.

 CANDADO DE AMBIENTE: solo corre contra el sitio de DESARROLLO escrito en $SitioDesarrolloPermitido.
 Edita esa línea una vez con el URL de tu sitio de desarrollo. Cualquier otro URL (producción) aborta.

 Uso (en tu PC):
   Install-Module PnP.PowerShell -Scope CurrentUser
   .\crear_listas_sharepoint.ps1 -SiteUrl "https://TU-TENANT.sharepoint.com/sites/TESCOM-DEV" -WhatIf
   .\crear_listas_sharepoint.ps1 -SiteUrl "https://TU-TENANT.sharepoint.com/sites/TESCOM-DEV" -CsvEmpleados "C:\ruta\operadores.csv"
 Es idempotente: reutiliza listas y columnas existentes y no duplica empleados (clave NumEmpleado).
#>
[CmdletBinding(SupportsShouldProcess)]
param(
  [Parameter(Mandatory)] [string] $SiteUrl,
  [string] $CsvEmpleados
)

# ---- Candado: solo el sitio de desarrollo ----
$SitioDesarrolloPermitido = "REEMPLAZAR-CON-URL-DEL-SITIO-DE-DESARROLLO"   # <-- editar
if ($SitioDesarrolloPermitido -like "REEMPLAZAR*") { throw "Falta configurar `$SitioDesarrolloPermitido en el script." }
if ($SiteUrl.TrimEnd('/') -ne $SitioDesarrolloPermitido.TrimEnd('/')) {
  throw "ABORTADO: '$SiteUrl' no es el sitio de desarrollo permitido. Este script no corre en producción."
}

$Quien = "$env:USERDOMAIN\$env:USERNAME"
$Log = Join-Path $PSScriptRoot ("crear_listas_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
function Log($m) { $l = "{0:s} [{1}] {2}" -f (Get-Date), $Quien, $m; Add-Content -Path $Log -Value $l; Write-Host $l }
Log "Inicio. Sitio: $SiteUrl  WhatIf=$($WhatIfPreference)"

Connect-PnPOnline -Url $SiteUrl -Interactive

function Ensure-List($title) {
  if (-not (Get-PnPList -Identity $title -ErrorAction SilentlyContinue)) {
    if ($PSCmdlet.ShouldProcess($title, "Crear lista")) {
      New-PnPList -Title $title -Template GenericList -OnQuickLaunch | Out-Null
      Log "Lista creada: $title"
    }
  }
}
function Ensure-Field($list, $name, $type, $choices) {
  if (-not (Get-PnPList -Identity $list -ErrorAction SilentlyContinue)) { return }   # (WhatIf: lista aún no existe)
  if (Get-PnPField -List $list -Identity $name -ErrorAction SilentlyContinue) { return }
  if ($PSCmdlet.ShouldProcess("$list.$name", "Agregar columna")) {
    if ($choices) { Add-PnPField -List $list -DisplayName $name -InternalName $name -Type Choice -Choices $choices -AddToDefaultView | Out-Null }
    else          { Add-PnPField -List $list -DisplayName $name -InternalName $name -Type $type -AddToDefaultView | Out-Null }
    Log "Columna agregada: $list.$name"
  }
}

# --- EMPLEADOS ---
Ensure-List "EMPLEADOS"
Ensure-Field "EMPLEADOS" "NumEmpleado"    Text
Ensure-Field "EMPLEADOS" "PuestoOracle"   Choice @("Assembly Operator A","Assembly Operator C","Testing Operator","Production technician","Assembly Lead A")
Ensure-Field "EMPLEADOS" "NivelActual"    Choice @("A","B","C","Tecnico","Lead")
Ensure-Field "EMPLEADOS" "NivelPropuesto" Choice @("A","B","C","Tecnico","Lead")   # sin opción vacía; la columna puede quedar sin valor
Ensure-Field "EMPLEADOS" "JobGroup"       Text
Ensure-Field "EMPLEADOS" "Supervisor"     Text
Ensure-Field "EMPLEADOS" "Grupo"          Choice @("Ensamble","Indirectos")
Ensure-Field "EMPLEADOS" "Linea"          Text
Ensure-Field "EMPLEADOS" "Correo"         Text
Ensure-Field "EMPLEADOS" "Notas"          Note
Ensure-Field "EMPLEADOS" "Activo"         Boolean
Ensure-Field "EMPLEADOS" "CargadoPor"     Text      # quién cargó el registro (usuario de Windows)

# --- COMPETENCIAS ---
Ensure-List "COMPETENCIAS"
Ensure-Field "COMPETENCIAS" "Area"           Choice @("HP","NPT1","NPT2","TB","Manual","Electronicos","Kits","Laser","Metal Finish","Empaque","Testing")
Ensure-Field "COMPETENCIAS" "NivelRequerido" Choice @("A","B","C","Tecnico")
Ensure-Field "COMPETENCIAS" "Paso"           Number
Ensure-Field "COMPETENCIAS" "Tipo"           Choice @("Rutinaria","Periodica","Esporadica","Curso")
Ensure-Field "COMPETENCIAS" "Criterio"       Note

# --- ENTRENAMIENTOS ---
Ensure-List "ENTRENAMIENTOS"
Ensure-Field "ENTRENAMIENTOS" "NumEmpleado"       Text
Ensure-Field "ENTRENAMIENTOS" "Competencia"       Text
Ensure-Field "ENTRENAMIENTOS" "Area"              Text      # copiada de COMPETENCIAS al registrar (permite agrupar)
Ensure-Field "ENTRENAMIENTOS" "FechaInicio"       DateTime
Ensure-Field "ENTRENAMIENTOS" "FechaFin"          DateTime
Ensure-Field "ENTRENAMIENTOS" "Resultado"         Choice @("En proceso","Aprobado","Reprobado")
Ensure-Field "ENTRENAMIENTOS" "Intentos"          Number
Ensure-Field "ENTRENAMIENTOS" "Entrenador"        Text      # responsable del entrenamiento
Ensure-Field "ENTRENAMIENTOS" "CapturadoPor"      Text      # quién entregó/capturó el registro
Ensure-Field "ENTRENAMIENTOS" "EvidenciaEmpleado" Note      # lo que aporta el propio empleado
Ensure-Field "ENTRENAMIENTOS" "EstatusAprobacion" Choice @("Pendiente","Aprobado","Rechazado")   # solo lo escribe el flujo
Ensure-Field "ENTRENAMIENTOS" "AprobadoPor"       Text
Ensure-Field "ENTRENAMIENTOS" "FechaAprobacion"   DateTime
Ensure-Field "ENTRENAMIENTOS" "Comentarios"       Note
if ((Get-PnPList -Identity "ENTRENAMIENTOS" -ErrorAction SilentlyContinue) -and
    $PSCmdlet.ShouldProcess("ENTRENAMIENTOS", "Activar historial de versiones")) {
  Set-PnPList -Identity "ENTRENAMIENTOS" -EnableVersioning $true | Out-Null
  Log "Historial de versiones activado en ENTRENAMIENTOS"
}

# --- EFECTIVIDAD ---
Ensure-List "EFECTIVIDAD"
Ensure-Field "EFECTIVIDAD" "NumEmpleado"         Text
Ensure-Field "EFECTIVIDAD" "Periodo"             Text
Ensure-Field "EFECTIVIDAD" "EficienciaPct"       Number
Ensure-Field "EFECTIVIDAD" "PiezasConformesPct"  Number
Ensure-Field "EFECTIVIDAD" "RetrabajoPct"        Number
Ensure-Field "EFECTIVIDAD" "EstacionesDominadas" Number
Ensure-Field "EFECTIVIDAD" "Fuente"              Text

# --- POSICIONES ---
Ensure-List "POSICIONES"
Ensure-Field "POSICIONES" "Linea"               Choice @("HP","NPT1","NPT2","TB","Manual","Electronicos","Kits","Laser","Metal Finish","Empaque")
Ensure-Field "POSICIONES" "NivelRequerido"      Choice @("A","B","C","Tecnico")
Ensure-Field "POSICIONES" "NumEmpleadoAsignado" Text

# --- FIRMANTES (aprobadores; inicialmente solo Gaddiel) ---
Ensure-List "FIRMANTES"
Ensure-Field "FIRMANTES" "Correo" Text
Ensure-Field "FIRMANTES" "Rol"    Choice @("Supervisor","Gerencia")
Ensure-Field "FIRMANTES" "Activo" Boolean

# --- Carga de empleados desde CSV (idempotente; el CSV queda en tu PC, no en git) ---
if ($CsvEmpleados -and (Test-Path $CsvEmpleados) -and (Get-PnPList -Identity "EMPLEADOS" -ErrorAction SilentlyContinue)) {
  foreach ($r in (Import-Csv $CsvEmpleados -Encoding UTF8)) {
    $nivel = "C"
    if ($r.Puesto_Oracle -like "*Operator A")   { $nivel = "A" }
    elseif ($r.Puesto_Oracle -like "*Operator C") { $nivel = "C" }
    elseif ($r.Puesto_Oracle -like "Testing*" -or $r.Puesto_Oracle -like "Production*") { $nivel = "Tecnico" }
    elseif ($r.Puesto_Oracle -like "*Lead*")    { $nivel = "Lead" }
    $valores = @{ Title = $r.Nombre; NumEmpleado = $r.NumEmpleado; PuestoOracle = $r.Puesto_Oracle
                  NivelActual = $nivel; JobGroup = $r.JobGroup; Supervisor = $r.Supervisor
                  Grupo = $r.Grupo; Notas = $r.Notas; Activo = $true; CargadoPor = $Quien }
    $existe = Get-PnPListItem -List "EMPLEADOS" -Query "<View><Query><Where><Eq><FieldRef Name='NumEmpleado'/><Value Type='Text'>$($r.NumEmpleado)</Value></Eq></Where></Query></View>"
    if ($existe) {
      if ($PSCmdlet.ShouldProcess($r.NumEmpleado, "Actualizar empleado")) { Set-PnPListItem -List "EMPLEADOS" -Identity $existe[0].Id -Values $valores | Out-Null; Log "Actualizado: $($r.NumEmpleado)" }
    } else {
      if ($PSCmdlet.ShouldProcess($r.NumEmpleado, "Agregar empleado")) { Add-PnPListItem -List "EMPLEADOS" -Values $valores | Out-Null; Log "Agregado: $($r.NumEmpleado)" }
    }
  }
}
Log "Fin. Registro en $Log"
