Attribute VB_Name = "ModCode128"
Option Explicit
'==============================================================================
' Code 128 "Auto" (subconjuntos B y C) con digito verificador modulo 103,
' para imprimir con una fuente de codigo de barras Code 128.
'
' Uso en una celda (aplicar la fuente de codigo de barras a esa celda):
'   =CODE128(A1)              texto codificado listo para la fuente
'   =CODE128(A1, cfgMapeo)    igual, eligiendo el mapeo de la fuente
'   =CODE128_CHECK(A1)        valor del digito verificador mod 103 (0-102)
'
' "Auto": usa el subconjunto C (pares de digitos) cuando hay 4 o mas digitos
' al inicio o al final, o 6 o mas en medio; el resto va en subconjunto B.
' El escaner lee exactamente el mismo texto que se capturo.
'
' Mapeo "LIBRE" (Libre Barcode 128 / IDAutomation):
'   valor 0 -> Chr 194, 1-94 -> Chr 33-126, 95-106 -> Chr 195-206
'   Start B = Chr 204, Start C = Chr 205, Stop = Chr 206
' Mapeo "GRANDZEBU" (fuente "Code 128" de grandzebu.net):
'   valor 0-94 -> Chr 32-126, 95-106 -> Chr 200-211
'   Start B = Chr 209, Start C = Chr 210, Stop = Chr 211
'
' Solo se aceptan caracteres ASCII 32-126 (sin acentos ni Ñ). Si el texto
' trae otro caracter, la funcion regresa "" para no imprimir un codigo malo.
'==============================================================================

Private Const C_CODE_C As Long = 99
Private Const C_CODE_B As Long = 100
Private Const C_START_B As Long = 104
Private Const C_START_C As Long = 105
Private Const C_STOP As Long = 106

Private Const SUB_NINGUNO As Long = 0
Private Const SUB_B As Long = 1
Private Const SUB_C As Long = 2

Public Function CODE128(ByVal Texto As Variant, Optional ByVal Mapeo As Variant) As String
    Dim v() As Long, n As Long, i As Long, zebu As Boolean, s As String

    If Not Code128Valores(ATexto(Texto), v, n) Then Exit Function
    If Not IsMissing(Mapeo) Then zebu = (InStr(1, UCase$(ATexto(Mapeo)), "ZEBU") > 0)
    For i = 0 To n - 1
        s = s & ValorAChar(v(i), zebu)
    Next i
    CODE128 = s
End Function

Public Function CODE128_CHECK(ByVal Texto As Variant) As Variant
    Dim v() As Long, n As Long

    If Code128Valores(ATexto(Texto), v, n) Then
        CODE128_CHECK = v(n - 2)
    Else
        CODE128_CHECK = ""
    End If
End Function

' Convierte el argumento (valor, celda o error) a texto sin espacios extremos.
' Una celda con error regresa "" (CStr falla y se atrapa el error).
Private Function ATexto(ByVal x As Variant) As String
    On Error GoTo SinValor
    If IsError(x) Then Exit Function
    ATexto = Trim$(CStr(x))
SinValor:
End Function

' Calcula los valores de simbolo: start, datos, verificador y stop.
Private Function Code128Valores(ByVal s As String, ByRef v() As Long, ByRef n As Long) As Boolean
    Dim L As Long, i As Long, d As Long, k As Long, c As Long
    Dim subset As Long, suma As Long

    L = Len(s)
    n = 0
    If L = 0 Then Exit Function
    For i = 1 To L
        c = AscW(Mid$(s, i, 1))
        If c < 32 Or c > 126 Then Exit Function
    Next i

    ReDim v(0 To 2 * L + 4)
    subset = SUB_NINGUNO
    i = 1
    Do While i <= L
        d = DigitosSeguidos(s, i)
        If subset <> SUB_C And UsarSubsetC(i, d, L) Then
            ' Corrida impar que llega al final: el primer digito va en B.
            If (d Mod 2 = 1) And (i + d - 1 = L) Then
                If subset = SUB_NINGUNO Then Agregar v, n, C_START_B
                subset = SUB_B
                Agregar v, n, AscW(Mid$(s, i, 1)) - 32
                i = i + 1
                d = d - 1
            End If
            If subset = SUB_NINGUNO Then
                Agregar v, n, C_START_C
            Else
                Agregar v, n, C_CODE_C
            End If
            subset = SUB_C
            For k = 1 To d \ 2
                Agregar v, n, CLng(Mid$(s, i, 2))
                i = i + 2
            Next k
        Else
            If subset = SUB_NINGUNO Then
                Agregar v, n, C_START_B
            ElseIf subset = SUB_C Then
                Agregar v, n, C_CODE_B
            End If
            subset = SUB_B
            Agregar v, n, AscW(Mid$(s, i, 1)) - 32
            i = i + 1
        End If
    Loop

    ' Digito verificador modulo 103: start + suma(posicion * valor).
    suma = v(0)
    For k = 1 To n - 1
        suma = suma + k * v(k)
    Next k
    Agregar v, n, suma Mod 103
    Agregar v, n, C_STOP
    Code128Valores = True
End Function

Private Function UsarSubsetC(ByVal i As Long, ByVal d As Long, ByVal L As Long) As Boolean
    If d < 2 Then
        UsarSubsetC = False
    ElseIf i = 1 And d = L Then
        UsarSubsetC = (d = 2 Or d >= 4)
    ElseIf i = 1 Or i + d - 1 = L Then
        UsarSubsetC = (d >= 4)
    Else
        UsarSubsetC = (d >= 6)
    End If
End Function

Private Function DigitosSeguidos(ByVal s As String, ByVal i As Long) As Long
    Dim c As Long
    Do While i <= Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then Exit Do
        DigitosSeguidos = DigitosSeguidos + 1
        i = i + 1
    Loop
End Function

Private Sub Agregar(ByRef v() As Long, ByRef n As Long, ByVal valor As Long)
    v(n) = valor
    n = n + 1
End Sub

Private Function ValorAChar(ByVal valor As Long, ByVal zebu As Boolean) As String
    If zebu Then
        If valor < 95 Then
            ValorAChar = ChrW$(valor + 32)
        Else
            ValorAChar = ChrW$(valor + 105)
        End If
    Else
        If valor = 0 Then
            ValorAChar = ChrW$(194)
        ElseIf valor < 95 Then
            ValorAChar = ChrW$(valor + 32)
        Else
            ValorAChar = ChrW$(valor + 100)
        End If
    End If
End Function
