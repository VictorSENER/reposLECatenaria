Attribute VB_Name = "vano"
'//
'// Rutina destinada a encontrar el vano correspondiente al radio
'//
Function vano(ByRef rayon As Double, z) As Double
Dim M As Integer
'//
'// inicializar variables
'//
M = 3
'//
'// buscar fila de la hoja 2 con el rango del radio actual y elegir el vano correspondiente
'//
If Not rayon = 0 And Abs(rayon) < Sheets("Vano").Cells(M, 2).Value Then
    While Abs(rayon) < Sheets("Vano").Cells(M, 3).Value _
    Or Abs(rayon) >= Sheets("Vano").Cells(M, 2).Value ' si quito el = se queda como antes...
    M = M + 1
    Wend
End If
'//
'// Actualizar el vano en su celda correspondiente
'//
If (z <= 18 And rayon < 450) Then
    vano = 27
Else
    vano = Sheets("Vano").Cells(M, 1).Value
End If
b = 4
While Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Punto singular").Cells(b, 2).Value + 243 Or Sheets("Punto singular").Cells(b, 1).Value <> "Aguja"
    b = b + 1
Wend
If Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Punto singular").Cells(b, 2).Value - 243 And Sheets("Replanteo").Cells(z, 33).Value < Sheets("Punto singular").Cells(b, 2).Value - 108 And Abs(Sheets("Replanteo").Cells(z, 6).Value) < 450 And Not IsEmpty(Sheets("Replanteo").Cells(z, 6).Value) _
And Sheets("Punto singular").Cells(b, 22).Value = "IN" Then
    vano = 27
ElseIf Sheets("Replanteo").Cells(z, 33).Value < Sheets("Punto singular").Cells(b, 2).Value + 243 And Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(b, 2).Value + 108 And Abs(Sheets("Replanteo").Cells(z, 6).Value) < 450 And Not IsEmpty(Sheets("Replanteo").Cells(z, 6).Value) _
And Sheets("Punto singular").Cells(b, 22).Value = "OUT" Then
    vano = 27
End If
End Function

