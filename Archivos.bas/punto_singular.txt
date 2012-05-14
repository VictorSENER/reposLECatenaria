Attribute VB_Name = "punto_singular"
Sub sing1(ByRef h, ByRef a)
Dim z As Integer
Dim pk0 As Double
' buscar en que fila de la hoja 4 se corresponde con el pk actual
dist_seg = 2

' si es uno de estos PS corregir el vano
' falta corregir los PS que tengan una longitud mayor a 9 metros
'If (Sheets(4).Cells(a, 1).Value = "Puente" Or Sheets(4).Cells(a, 1).Value = "Conducto" Or Sheets(4).Cells(a, 1).Value = "P.I." Or Sheets(4).Cells(a, 1).Value = "Drenaje" Or Sheets(4).Cells(a, 1).Value = "P.N.") Then
    z = h
    'caso puente y conducto
    If (Sheets(1).Cells(h, 33).Value >= (Sheets(4).Cells(a, 2).Value - dist_seg) And Sheets(1).Cells(h, 33).Value <= (Sheets(4).Cells(a, 21).Value + dist_seg) And _
    Sheets(4).Cells(a, 1).Value = "Puente") Then
        dist_restar = Sheets(1).Cells(h, 33).Value - (Sheets(4).Cells(a, 2).Value - dist_seg)
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        Call singular.restar(dist_restar, div, z, h, a)
    ElseIf (Sheets(1).Cells(h, 33).Value - Sheets(4).Cells(a - 1, 2).Value <= dist_seg And Sheets(4).Cells(a - 1, 1).Value = "Puente") Then
        dist_restar = Sheets(1).Cells(h, 33).Value - (Sheets(4).Cells(a - 1, 2).Value - dist_seg)
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        Call singular.restar(dist_restar, div, z, h, a - 1)
    ElseIf (Sheets(1).Cells(h, 33).Value >= (Sheets(4).Cells(a, 2).Value - dist_seg) And Sheets(1).Cells(h, 33).Value <= (Sheets(4).Cells(a, 21).Value + dist_seg) _
    And (Sheets(4).Cells(a, 1).Value = "Conducto" Or Sheets(4).Cells(a, 1).Value = "P.I." Or Sheets(4).Cells(a, 1).Value = "Drenaje" Or Sheets(4).Cells(a, 1).Value = "P.N.")) Then
        dist_restar = Sheets(1).Cells(h, 33).Value - (Sheets(4).Cells(a, 2).Value - dist_seg)
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        Call singular.restar(dist_restar, div, z, h, a)
    ElseIf ((Sheets(1).Cells(h, 33).Value - Sheets(4).Cells(a - 1, 21).Value) <= dist_seg And (Sheets(4).Cells(a - 1, 1).Value = "Conducto" Or Sheets(4).Cells(a - 1, 1).Value = "P.I." Or Sheets(4).Cells(a - 1, 1).Value = "Drenaje" Or Sheets(4).Cells(a - 1, 1).Value = "P.N.")) Then
        dist_restar = Sheets(1).Cells(h, 33).Value - (Sheets(4).Cells(a - 1, 2).Value - dist_seg)
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        Call singular.restar(dist_restar, div, z, h, a - 1)
    End If
    
       
'End If

While Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(a, 21).Value And Sheets(4).Cells(a, 23).Value <> "FINAL"
    a = a + 1
Wend
End Sub

Sub sing(ByRef h, ByRef k, ByRef a, ByRef b)
If a <> 3 Then
While Sheets(1).Cells(h, 33).Value >= Sheets(4).Cells(a, 21).Value And Sheets(4).Cells(a, 23).Value <> "FINAL" _
And Sheets(4).Cells(a + 1, 21).Value - Sheets(4).Cells(a, 21).Value <= va_max And Sheets(4).Cells(a + 1, 1).Value = "Aguja"
    a = a + 1
Wend
End If
Sheets(1).Cells(5, 1).Value = Sheets(4).Cells(a, 2).Value
' buscar en que fila de la hoja 4 se corresponde con el pk actual

If Sheets(4).Cells(a, 1) = "7 > P.S. > 5,2 m" Then
caca = 0
End If

If (Sheets(1).Cells(h - 2, 33).Value < Sheets(4).Cells(a, 21).Value And _
Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(a, 2).Value And _
(Sheets(4).Cells(a, 1) = "7 > P.S. > 5,2 m" Or Sheets(4).Cells(a, 1) = "PuenteXL")) Then
    
    If Sheets(4).Cells(a, 1) = "7 > P.S. > 5,2 m" And Sheets(4).Cells(a + 1, 1) = "Aguja" And Sheets(4).Cells(a + 1, 2) - Sheets(4).Cells(a, 2) < 63 Then
        a = a + 1
    Else
        Call singular.paso_superior(h, k, a)
    End If

End If
If Sheets(1).Cells(h, 33).Value >= Sheets(4).Cells(a, 3).Value _
And Sheets(1).Cells(h, 33).Value <= Sheets(4).Cells(a, 21).Value _
And Sheets(4).Cells(a, 1) = "Viaducto" Then
    Call singular.Viaducto(h, k, b, a)
End If

If (Sheets(1).Cells(h, 33).Value >= Sheets(4).Cells(a, 2).Value And Sheets(1).Cells(h, 33).Value <= Sheets(4).Cells(a, 21).Value _
And Sheets(4).Cells(a, 1) = "Tunel") Then
    If Sheets(1).Cells(h - 2, 33).Value >= (Sheets(4).Cells(a, 2).Value - dist_va_max) And _
     Sheets(4).Cells(a, 1).Value = "Tunel" And Sheets(1).Cells(h - 2, 38).Value <> "Tunel" Then
        dist_restar = Sheets(1).Cells(h - 2, 33).Value - (Sheets(4).Cells(a, 2).Value - dist_va_max) '+ (inc_norm_va / 2)
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        z = h - 2
        Call singular.restar(dist_restar, div, z, h, a)
        
    End If
        Sheets(1).Cells(h, 25).Value = Sheets(4).Cells(a, 23).Value
        Sheets(1).Cells(h, 38).Value = Sheets(4).Cells(a, 1).Value
    With Sheets(1).Range(Sheets(1).Cells(h, 25), Sheets(1).Cells(h + 1, 25))
    .MergeCells = True
    .Borders(xlEdgeLeft).LineStyle = 2
    .Borders(xlEdgeLeft).ColorIndex = 15
    .Borders(xlEdgeTop).LineStyle = 2
    .Borders(xlEdgeTop).ColorIndex = 15
    .Borders(xlEdgeBottom).LineStyle = 2
    .Borders(xlEdgeBottom).ColorIndex = 15
    .Borders(xlEdgeRight).LineStyle = 2
    .Borders(xlEdgeRight).ColorIndex = 15
    End With
    If Sheets(1).Cells(h - 1, 4).Value > va_max_tunel Then
        Sheets(1).Cells(h - 1, 4).Value = va_max_tunel
        Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 2, 33) + Sheets(1).Cells(h - 1, 4)
        Call radio.radio1(h)
    End If
End If

If (Sheets(1).Cells(h - 2, 33).Value < Sheets(4).Cells(a, 2).Value And Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(a, 2).Value And _
Sheets(4).Cells(a, 1) = "Aguja") Then

Call singular.aguja(h, k, b, a)
End If
'If Sheets(1).Cells(h, 33).Value >= Sheets(4).Cells(a, 2).Value And Sheets(1).Cells(h, 33).Value <= Sheets(4).Cells(a, 21).Value And _
'Sheets(4).Cells(a, 1).Value = "Zona" Then
'Call singular.Zona(h, a)
'End If
While Sheets(1).Cells(h, 33).Value >= Sheets(4).Cells(a, 21).Value And Sheets(4).Cells(a, 23).Value <> "FINAL"
    a = a + 1
Wend
End Sub

