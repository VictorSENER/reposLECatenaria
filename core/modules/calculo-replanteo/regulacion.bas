Attribute VB_Name = "regulacion"
'//
'// Rutina destinada a corregir los errores de incremento de los vanos
'// (se deberá ampliar en caso de diferenciales entre vano max y min mayores al acutal)
'//
Sub regulacion(ByRef h, ByRef a)
'//
'// Inizializar variables
'//
z = h
'//
'// Comprobación n-1
'//
While (Sheets("Replanteo").Cells(h - 1, 4).Value - Sheets("Replanteo").Cells(h - 3, 4).Value) > dist_va_max
    Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Replanteo").Cells(h - 1, 4).Value - inc_norm_va
    Call actualizar(z, h, a)
Wend
'//
'// Comprobación n-2
'//
While (Sheets("Replanteo").Cells(h - 3, 4).Value - Sheets("Replanteo").Cells(h - 5, 4).Value) > dist_va_max Or _
(Sheets("Replanteo").Cells(h - 3, 4).Value - Sheets("Replanteo").Cells(h - 1, 4).Value) > dist_va_max
    Sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Replanteo").Cells(h - 3, 4).Value - inc_norm_va
    Call actualizar(z - 2, h, a)
Wend
'//
'// Comprobación n-3
'//
While (Sheets("Replanteo").Cells(h - 5, 4).Value - Sheets("Replanteo").Cells(h - 3, 4).Value) > dist_va_max
    Sheets("Replanteo").Cells(h - 5, 4).Value = Sheets("Replanteo").Cells(h - 5, 4).Value - inc_norm_va
    Call actualizar(z - 4, h, a)
Wend
'//
'// Comprobación n-5
'//
While (Sheets("Replanteo").Cells(h - 7, 4).Value - Sheets("Replanteo").Cells(h - 5, 4).Value) > dist_va_max
    Sheets("Replanteo").Cells(h - 7, 4).Value = Sheets("Replanteo").Cells(h - 7, 4).Value - inc_norm_va
    Call actualizar(z - 6, h, a)
Wend
'//
'// Comprobación n-7
'//
While (Sheets("Replanteo").Cells(h - 9, 4).Value - Sheets("Replanteo").Cells(h - 7, 4).Value) > dist_va_max
    Sheets("Replanteo").Cells(h - 9, 4).Value = Sheets("Replanteo").Cells(h - 9, 4).Value - inc_norm_va
    Call actualizar(z - 8, h, a)
Wend
End Sub
'//
'// Rutina destinada a ajustar los cambios realizados: incremento de pk, radio, vano, punto singular
'//
Sub actualizar(z, h, a)
While z <= h
aqui:
    Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z - 1, 4).Value + Sheets("Replanteo").Cells(z - 2, 33).Value
    'Call punto_singular.sing(z + 2, k - 1, a - 1, b)
    Call radio.radio1(z)
    Call punto_singular.sing1(z, a - 1, 0, 0)
    If vano.vano(Sheets("Replanteo").Cells(z - 2, 6).Value, h) < Sheets("Replanteo").Cells(z - 1, 4).Value Then
        Sheets("Replanteo").Cells(z - 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(z - 2, 6).Value, h)
        'sheets("Replanteo").Cells(z, 33).Value = sheets("Replanteo").Cells(z - 1, 4).Value + sheets("Replanteo").Cells(z - 2, 33).Value
        GoTo aqui:
        'sheets("Replanteo").Cells(z + 1, 4).Value = vano.vano(sheets("Replanteo").Cells(z - 2, 6).Value)
    ElseIf vano.vano(Sheets("Replanteo").Cells(z, 6).Value, h) < Sheets("Replanteo").Cells(z + 1, 4).Value Then
        Sheets("Replanteo").Cells(z + 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(z, 6).Value, h)
    End If
    z = z + 2
Wend
z = h
End Sub

'//
'// Rutina destinada a corregir los errores de incremento de los vanos
'// (se deberá ampliar en caso de diferenciales entre vano max y min mayores al acutal)
'//
Function long_restar(ByRef h, ByRef a) As Double
Dim n(0 To 4) As Double
'//
'// Inizializar variables
'//
z = h
n(0) = Sheets("Replanteo").Cells(h - 1, 4).Value
n(1) = Sheets("Replanteo").Cells(h - 3, 4).Value
n(2) = Sheets("Replanteo").Cells(h - 5, 4).Value
n(3) = Sheets("Replanteo").Cells(h - 7, 4).Value
n(4) = Sheets("Replanteo").Cells(h - 9, 4).Value
'//
'// Comprobación n-1
'//
While (n(0) - n(1)) > dist_va_max
    n(0) = n(0) - inc_norm_va
    long_restar = long_restar + inc_norm_va
    'sheets("Replanteo").Cells(h - 1, 4).Value = sheets("Replanteo").Cells(h - 1, 4).Value - inc_norm_va
    'Call actualizar(z, h, a)
Wend
'//
'// Comprobación n-2
'//
While (n(1) - n(2)) > dist_va_max Or (n(1) - n(0)) > dist_va_max
    n(1) = n(1) - inc_norm_va
    long_restar = long_restar + inc_norm_va
    'sheets("Replanteo").Cells(h - 3, 4).Value = sheets("Replanteo").Cells(h - 3, 4).Value - inc_norm_va
    'Call actualizar(z - 2, h, a)
Wend
'//
'// Comprobación n-3
'//
While (n(2) - n(1)) > dist_va_max
    n(2) = n(2) - inc_norm_va
    long_restar = long_restar + inc_norm_va
    'sheets("Replanteo").Cells(h - 5, 4).Value = sheets("Replanteo").Cells(h - 5, 4).Value - inc_norm_va
    'Call actualizar(z - 4, h, a)
Wend
'//
'// Comprobación n-5
'//
While (n(3) - n(2)) > dist_va_max
    n(3) = n(3) - inc_norm_va
    long_restar = long_restar + inc_norm_va
    'sheets("Replanteo").Cells(h - 7, 4).Value = sheets("Replanteo").Cells(h - 7, 4).Value - inc_norm_va
    'Call actualizar(z - 6, h, a)
Wend
'//
'// Comprobación n-7
'//
While (n(4) - n(3)) > dist_va_max
    n(4) = n(4) - inc_norm_va
    long_restar = long_restar + inc_norm_va
    'sheets("Replanteo").Cells(h - 9, 4).Value = sheets("Replanteo").Cells(h - 9, 4).Value - inc_norm_va
    'Call actualizar(z - 8, h, a)
Wend

End Function

