Attribute VB_Name = "vano"
'//
'// Rutina destinada a encontrar el vano correspondiente al radio
'//
Function vano(ByRef rayon As Double) As Double
Dim m As Integer
'//
'// inicializar variables
'//
m = 3
'//
'// buscar fila de la hoja 2 con el rango del radio actual y elegir el vano correspondiente
'//
If Not rayon = 0 Then
    While Abs(rayon) < Sheets(2).Cells(m, 3).Value _
    Or Abs(rayon) > Sheets(2).Cells(m, 2).Value
    m = m + 1
    Wend
End If
'//
'// Actualizar el vano en su celda correspondiente
'//
vano = Sheets(2).Cells(m, 1).Value
End Function

