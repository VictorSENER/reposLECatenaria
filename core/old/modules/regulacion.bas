Attribute VB_Name = "regulacion"
'//
'// Rutina destinada a corregir los errores de incremento de los vanos
'// (se deberá ampliar en caso de diferenciales entre vano max y min mayores al acutal)
'//
Sub regulacion(ByRef h, ByRef a, ByRef b, ByRef k)
'//
'// Inizializar variables
'//
z = h
'//
'// Comprobación n-1
'//
While (Sheets(1).Cells(h - 1, 4).Value - Sheets(1).Cells(h - 3, 4).Value) > dist_va_max
    Sheets(1).Cells(h - 1, 4).Value = Sheets(1).Cells(h - 1, 4).Value - inc_norm_va
    Call actualizar(z, h, a, k)
Wend
'//
'// Comprobación n-2
'//
While (Sheets(1).Cells(h - 3, 4).Value - Sheets(1).Cells(h - 5, 4).Value) > dist_va_max Or _
(Sheets(1).Cells(h - 3, 4).Value - Sheets(1).Cells(h - 1, 4).Value) > dist_va_max
    Sheets(1).Cells(h - 3, 4).Value = Sheets(1).Cells(h - 3, 4).Value - inc_norm_va
    Call actualizar(z - 2, h, a, k)
Wend
'//
'// Comprobación n-3
'//
While (Sheets(1).Cells(h - 5, 4).Value - Sheets(1).Cells(h - 3, 4).Value) > dist_va_max
    Sheets(1).Cells(h - 5, 4).Value = Sheets(1).Cells(h - 5, 4).Value - inc_norm_va
    Call actualizar(z - 4, h, a, k)
Wend
'//
'// Comprobación n-5
'//
While (Sheets(1).Cells(h - 7, 4).Value - Sheets(1).Cells(h - 5, 4).Value) > dist_va_max
    Sheets(1).Cells(h - 7, 4).Value = Sheets(1).Cells(h - 7, 4).Value - inc_norm_va
    Call actualizar(z - 6, h, a, k)
Wend
End Sub
'//
'// Rutina destinada a ajustar los cambios realizados: incremento de pk, radio, vano, punto singular
'//
Sub actualizar(z, h, a, k)
While z <= h
    Sheets(1).Cells(z, 33).Value = Sheets(1).Cells(z - 1, 4).Value + Sheets(1).Cells(z - 2, 33).Value
    'Call punto_singular.sing(z + 2, k - 1, a - 1, b)
    Call radio1(z)
    Call punto_singular.sing1(z, a - 1)
    If vano.vano(Sheets(1).Cells(z - 2, 6).Value) < Sheets(1).Cells(z - 1, 4).Value Then
        Sheets(1).Cells(z - 1, 4).Value = vano.vano(Sheets(1).Cells(z - 2, 6).Value)
    End If
    z = z + 2
Wend
z = h
End Sub

