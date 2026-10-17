Attribute VB_Name = "ModImpresion"
Option Explicit
'==============================================================================
' Botones de la hoja ETIQUETA, LISTA y CONFIG.
' Los nombres de rango (In_JOB, BC_JOB, cfgFuente, ...) estan definidos en el
' libro; si se mueve una celda, el nombre la sigue y las macros no cambian.
'==============================================================================

Private Const CLAVE As String = ""          ' contraseña de proteccion de hojas
Private Const L_CAMPOS As String = "In_JOB,In_OB,In_QTY,In_PN,In_HEATCODE,In_HEATLOT"
Private Const L_PRUEBAS As String = "Scan_JOB,Scan_OB,Scan_QTY,Scan_PN"
Private Const L_BARRAS As String = "BC_JOB,BC_OB,BC_QTY,BC_PN"
Private Const L_NOMBRES As String = "W.O. / JOB,OB,QTY TOTAL,P/N (MODELO)"

' Hoja LISTA
Private Const FILA_INICIO As Long = 5
Private Const COL_JOB As Long = 2           ' B: JOB ... G: HEAT LOT
Private Const COL_COPIAS As Long = 8        ' H
Private Const COL_ESTADO As Long = 9        ' I
Private Const MAX_COPIAS As Long = 500

'------------------------------------------------------------------------------
' ETIQUETA
'------------------------------------------------------------------------------
Public Sub Imprimir_Etiqueta()
    If Not DatosListos(True) Then Exit Sub
    If ImprimirActual(NumeroCopias(Rango("In_COPIAS").Value)) Then
        On Error Resume Next
        Rango("In_JOB").Select
    End If
End Sub

Public Sub Vista_Previa()
    Application.Calculate
    shEtiqueta.PrintPreview
End Sub

Public Sub Limpiar_Etiqueta()
    Dim nm As Variant
    For Each nm In Split(L_CAMPOS & "," & L_PRUEBAS, ",")
        Rango(CStr(nm)).MergeArea.ClearContents
    Next nm
    Rango("In_COPIAS").Value = 1
    On Error Resume Next
    shEtiqueta.Activate
    Rango("In_JOB").Select
End Sub

Public Sub Elegir_Impresora()
    If Application.Dialogs(xlDialogPrinterSetup).Show Then
        MsgBox "Impresora seleccionada:" & vbCrLf & Application.ActivePrinter, vbInformation
    End If
End Sub

'------------------------------------------------------------------------------
' LISTA: imprime cada fila cuyo ESTADO no diga "IMPRESO".
'------------------------------------------------------------------------------
Public Sub Imprimir_Lista()
    Dim ws As Worksheet, fila As Long, ultima As Long, i As Long
    Dim ordenes As Long, total As Long, hechas As Long, saltadas As Long
    Dim arrCampos() As String, respaldo(0 To 6) As Variant, motivo As String

    Set ws = shLista
    ultima = ws.Cells(ws.Rows.Count, COL_JOB).End(xlUp).Row
    For fila = FILA_INICIO To ultima
        If FilaPendiente(ws, fila) Then
            ordenes = ordenes + 1
            total = total + NumeroCopias(ws.Cells(fila, COL_COPIAS).Value)
        End If
    Next fila
    If ordenes = 0 Then
        MsgBox "No hay filas pendientes en la LISTA." & vbCrLf & _
               "(Para reimprimir una fila, borra su ESTADO.)", vbInformation
        Exit Sub
    End If
    If MsgBox("Se imprimirán " & ordenes & " órdenes (" & total & " etiquetas)." & _
              vbCrLf & vbCrLf & "¿Continuar?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    ' Guardar lo que esta capturado en ETIQUETA para regresarlo al final.
    arrCampos = Split(L_CAMPOS, ",")
    For i = 0 To 5
        respaldo(i) = Rango(arrCampos(i)).Value
    Next i
    respaldo(6) = Rango("In_COPIAS").Value

    Application.ScreenUpdating = False
    For fila = FILA_INICIO To ultima
        If FilaPendiente(ws, fila) Then
            For i = 0 To 5
                Rango(arrCampos(i)).Value = ws.Cells(fila, COL_JOB + i).Value
            Next i
            motivo = ""
            If Not DatosValidos(motivo) Then
                ws.Cells(fila, COL_ESTADO).Value = "REVISAR: " & motivo
                saltadas = saltadas + 1
            ElseIf ImprimirActual(NumeroCopias(ws.Cells(fila, COL_COPIAS).Value)) Then
                ws.Cells(fila, COL_ESTADO).Value = "IMPRESO " & Format$(Now, "dd/mm/yyyy hh:nn")
                hechas = hechas + 1
            Else
                Exit For
            End If
        End If
    Next fila

    For i = 0 To 5
        Rango(arrCampos(i)).Value = respaldo(i)
    Next i
    Rango("In_COPIAS").Value = respaldo(6)
    Application.ScreenUpdating = True

    MsgBox "Órdenes impresas: " & hechas & vbCrLf & _
           "Órdenes por revisar: " & saltadas, IIf(saltadas > 0, vbExclamation, vbInformation)
End Sub

Public Sub Limpiar_Lista()
    If MsgBox("¿Borrar toda la LISTA?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    shLista.Range(shLista.Cells(FILA_INICIO, COL_JOB), _
                  shLista.Cells(shLista.Rows.Count, COL_ESTADO)).ClearContents
End Sub

'------------------------------------------------------------------------------
' CONFIG: aplica la fuente y el tamaño a las celdas de codigo de barras.
'------------------------------------------------------------------------------
Public Sub Aplicar_Fuente()
    Dim fuente As String, tam As Double, nm As Variant

    fuente = Trim$(CStr(Rango("cfgFuente").Value))
    If fuente = "" Then
        MsgBox "Escribe el nombre de la fuente en CONFIG.", vbExclamation
        Exit Sub
    End If
    If IsNumeric(Rango("cfgTamano").Value) Then tam = CDbl(Rango("cfgTamano").Value)
    If tam < 8 Or tam > 120 Then tam = 35.5

    shEtiqueta.Unprotect CLAVE
    For Each nm In Split(L_BARRAS, ",")
        With Rango(CStr(nm)).MergeArea.Font
            .Name = fuente
            .Size = tam
        End With
    Next nm
    shEtiqueta.Protect CLAVE
    With Rango("cfgPruebaBarra").MergeArea.Font
        .Name = fuente
        .Size = tam
    End With
    Application.Calculate
    MsgBox "Fuente aplicada: " & fuente & " (" & tam & " pt)", vbInformation
End Sub

'------------------------------------------------------------------------------
' Apoyo
'------------------------------------------------------------------------------
Private Function Rango(ByVal nombre As String) As Range
    Set Rango = ThisWorkbook.Names(nombre).RefersToRange
End Function

Private Function TextoCelda(ByVal r As Range) As String
    Dim v As Variant
    v = r.Cells(1, 1).Value
    If Not IsError(v) Then TextoCelda = Trim$(CStr(v))
End Function

Private Function NumeroCopias(ByVal v As Variant) As Long
    Dim x As Double
    If Not IsError(v) Then
        If IsNumeric(v) Then x = CDbl(v)
    End If
    If x > MAX_COPIAS Then x = MAX_COPIAS
    NumeroCopias = CLng(Int(x))
    If NumeroCopias < 1 Then NumeroCopias = 1
End Function

Private Function FilaPendiente(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    If TextoCelda(ws.Cells(fila, COL_JOB)) = "" Then Exit Function
    FilaPendiente = (UCase$(Left$(TextoCelda(ws.Cells(fila, COL_ESTADO)), 7)) <> "IMPRESO")
End Function

' Revisa los 4 campos con codigo de barras. Regresa False si alguno falta o
' tiene caracteres que Code 128 no puede codificar; "motivo" dice cual.
Private Function DatosValidos(ByRef motivo As String) As Boolean
    Dim arrCampos() As String, arrBarras() As String, arrNombres() As String, i As Long
    Dim faltan As String, malos As String

    arrCampos = Split(L_CAMPOS, ",")
    arrBarras = Split(L_BARRAS, ",")
    arrNombres = Split(L_NOMBRES, ",")
    Application.Calculate
    For i = 0 To 3
        If TextoCelda(Rango(arrCampos(i))) = "" Then
            faltan = faltan & IIf(faltan = "", "", ", ") & arrNombres(i)
        ElseIf TextoCelda(Rango(arrBarras(i))) = "" Then
            malos = malos & IIf(malos = "", "", ", ") & arrNombres(i)
        End If
    Next i
    If malos <> "" Then
        motivo = "carácter no válido en " & malos
    ElseIf faltan <> "" Then
        motivo = "falta " & faltan
    End If
    DatosValidos = (motivo = "")
End Function

Private Function DatosListos(ByVal preguntar As Boolean) As Boolean
    Dim motivo As String

    If DatosValidos(motivo) Then
        DatosListos = True
    ElseIf Left$(motivo, 5) = "falta" And preguntar Then
        DatosListos = (MsgBox("Falta capturar: " & Mid$(motivo, 7) & vbCrLf & vbCrLf & _
                              "¿Imprimir de todos modos?", vbYesNo + vbExclamation) = vbYes)
    Else
        MsgBox "No se puede generar el código de barras: " & motivo & "." & vbCrLf & vbCrLf & _
               "Code 128 solo acepta letras sin acento, números y símbolos (sin Ñ).", vbCritical
    End If
End Function

Private Function ImprimirActual(ByVal copias As Long) As Boolean
    On Error GoTo Falla
    Application.Calculate
    shEtiqueta.PrintOut Copies:=copias, Collate:=True
    ImprimirActual = True
    Exit Function
Falla:
    Application.ScreenUpdating = True
    MsgBox "No se pudo imprimir: " & Err.Description, vbCritical
End Function
