Option Explicit

' Al abrir (o al habilitar macros) se recalculan las formulas CODE128 para
' que los codigos de barras se dibujen aunque Excel los haya calculado antes
' de cargar las macros.
Private Sub Workbook_Open()
    On Error Resume Next
    Application.CalculateFull
End Sub
