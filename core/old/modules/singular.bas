Attribute VB_Name = "singular"
Sub Viaducto(ByRef h, ByRef k, ByRef b, ByRef a)
Dim vano01 As Double, vano12 As Double, vano0 As Double
Dim z As Integer
Dim pkref As Double, pk0 As Double, pk1 As Double, pk2 As Double, dist_restar As Double
b = 3
vano1 = Sheets(1).Cells(h + 1, 4).Value
pk0 = Sheets(1).Cells(h, 33).Value
pk1 = Sheets(4).Cells(a, b).Value
pk2 = Sheets(4).Cells(a, b + 1).Value
vano0 = Sheets(1).Cells(h - 1, 4).Value
dist_restar = pk0 - pk1
If dist_restar < (vano1 - (pk2 - pk1)) And pk2 <> 0 Then
    h = h + 2
    dist_restar = dist_restar + (pk2 - pk1) + inc_norm_va
    Sheets(1).Cells(h, 33).Value = pk1
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    Sheets(1).Cells(h + 1, 4).Value = pk2 - pk1
    Sheets(1).Cells(h - 1, 4).Value = (pk2 - pk1) + inc_norm_va
    z = h - 2
ElseIf pk2 = 0 Then
    dist_restar = pk0 - pk1
    Sheets(1).Cells(h, 33).Value = pk1
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    Call radio.radio1(h)
    Sheets(1).Cells(h + 1, 4).Value = vano.vano(Sheets(1).Cells(h, 6).Value)
    z = h
Else
    dist_restar = pk0 - pk1
    Sheets(1).Cells(h, 33).Value = pk1
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    Sheets(1).Cells(h + 1, 4).Value = pk2 - pk1
    z = h
End If
Sheets(1).Cells(h - 2, 25).Value = Sheets(4).Cells(a, 24).Value
    With Sheets(1).Range(Sheets(1).Cells(h - 2, 25), Sheets(1).Cells(h - 1, 25))
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
Call restar(dist_restar, div, z, h, a)

While Not IsEmpty(Sheets(4).Cells(a, b))
    Sheets(1).Cells(h, 25).Value = Sheets(4).Cells(a, 25).Value
    'Sheets(1).Range(Sheets(1).Cells(h, 25), Sheets(1).Cells(h + 1, 25)).MergeCells = True
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
    h = h + 2
    b = b + 1
        If Not IsEmpty(Sheets(4).Cells(a, b).Value) Then
            Sheets(1).Cells(h, 33).Value = Sheets(4).Cells(a, b).Value
            Call radio.radio1(h)
            Sheets(1).Cells(h - 1, 4).Value = Sheets(1).Cells(h, 33).Value - Sheets(1).Cells(h - 2, 33).Value
        Else
            Sheets(1).Cells(h - 1, 4).Value = vano.vano(Sheets(1).Cells(h - 2, 6).Value)
        End If
Wend
Sheets(1).Cells(h, 25).Value = Sheets(4).Cells(a, 26).Value
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
h = h - 2
a = a + 1
End Sub
Sub paso_superior(ByRef h, ByRef k, ByRef a)
Dim l1 As Double, l2 As Double, vano0 As Double, dist_restar As Double, pmedio As Double
Dim pk1 As Double, pk2 As Double, div As Double, vano12 As Double, vanoref As Double
Dim z As Integer
l1 = Sheets(4).Cells(a, 2).Value - Sheets(1).Cells(h - 2, 33).Value
l2 = Sheets(1).Cells(h, 33).Value - Sheets(4).Cells(a, 21).Value
vano0 = Sheets(1).Cells(h + 1, 4).Value
pmedio = (vano0 - (Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value)) / 2
pk1 = Sheets(4).Cells(a, 2).Value - pmedio
pk2 = Sheets(4).Cells(a, 21).Value + pmedio
dist_restar = Sheets(1).Cells(h - 2, 33).Value - pk1
vano12 = pk2 - pk1
z = h
If dist_restar <= 0 Then
    h = h + 2
    Sheets(1).Cells(h, 33).Value = pk2
    Call radio.radio1(h)
    Sheets(1).Cells(h - 2, 33).Value = pk1
    Call radio.radio1(h - 2)
        If vano.vano(Sheets(1).Cells(h - 2, 6).Value) < Sheets(1).Cells(h - 1, 4).Value Then
            vano0 = vano0 - inc_norm_va
            pmedio = (vano0 - (Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value)) / 2
            pk1 = Sheets(4).Cells(a, 2).Value - pmedio
            pk2 = Sheets(4).Cells(a, 21).Value + pmedio
            dist_restar = Sheets(1).Cells(h - 2, 33).Value - pk1 - inc_norm_va
            Sheets(1).Cells(h - 1, 4).Value = vano0
        End If
    dist_restar = Sheets(1).Cells(h - 3, 4).Value - (dist_restar * (-1))
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
Else
    Sheets(1).Cells(h - 4, 33).Value = Sheets(1).Cells(h - 4, 33).Value - dist_restar
    Call radio.radio(h - 4)
        If vano.vano(Sheets(1).Cells(h - 4, 6).Value) < Sheets(1).Cells(h - 3, 4).Value Then
            vano0 = vano0 - inc_norm_va
            pmedio = (vano0 - (Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value)) / 2
            pk1 = Sheets(4).Cells(a, 2).Value - pmedio
            pk2 = Sheets(4).Cells(a, 21).Value + pmedio
            dist_restar = Sheets(1).Cells(h - 2, 33).Value - pk1
            Sheets(1).Cells(h - 1, 4).Value = vano0
        End If
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    Sheets(1).Cells(h, 33).Value = pk2
    Call radio.radio1(h)
    Sheets(1).Cells(h - 2, 33).Value = pk1
    Call radio.radio1(h - 2)
End If
Call restar(dist_restar, div, z - 2, h, a)

h = h - 2
a = a + 1
End Sub
Sub aguja(ByRef h, ByRef k, ByRef b, ByRef a)
Dim dist_restar As Double
Dim pk1 As Double, pk0 As Double, vano12 As Double, vanoref As Double
Dim z As Integer
    pk1 = Sheets(4).Cells(a, 2).Value
pk0 = Sheets(1).Cells(h, 33).Value
'caso particular de tener un paso superior bajo antes de la aguja
If Sheets(4).Cells(a - 1, 1).Value = "7 > P.S. > 5,2 m" And _
Sheets(4).Cells(a, 2).Value - Sheets(4).Cells(a - 1, 21).Value < (va_max) _
And Sheets(4).Cells(a, 2).Value - Sheets(4).Cells(a - 1, 21).Value > (inc_norm_va * 6) Then
    'hace falta añadir una celda
    If Sheets(1).Cells(h - 2, 33).Value < (Sheets(4).Cells(a - 1, 21).Value + dist_va_max) Then
        pk0 = Sheets(1).Cells(h - 2, 33).Value
        h = h + 2
        Sheets(1).Cells(h - 1, 4).Value = pk1 - (Sheets(4).Cells(a - 1, 21).Value + dist_va_max)
        Sheets(1).Cells(h - 3, 4).Value = Sheets(4).Cells(a - 1, 21).Value - Sheets(4).Cells(a - 1, 2).Value + (4 * inc_norm_va)
        Sheets(1).Cells(h - 4, 33).Value = pk1 - Sheets(1).Cells(h - 1, 4).Value - Sheets(1).Cells(h - 3, 4).Value
        Call radio.radio1(h - 4)
        Sheets(1).Cells(h - 2, 33).Value = pk1 - Sheets(1).Cells(h - 1, 4).Value
        Call radio.radio1(h - 2)
        Sheets(1).Cells(h, 33).Value = pk1
        Call radio.radio1(h)
        Sheets(1).Cells(h + 1, 4).Value = vano.vano(Sheets(1).Cells(h, 6).Value)
        dist_restar = pk0 - Sheets(1).Cells(h - 4, 33).Value
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        z = h - 4
    'no hace falta añadir celda
    Else
        pk0 = Sheets(1).Cells(h - 4, 33).Value
        Sheets(1).Cells(h - 1, 4).Value = pk1 - (Sheets(4).Cells(a - 1, 21).Value + dist_va_max)
        Sheets(1).Cells(h - 3, 4).Value = Sheets(4).Cells(a - 1, 21).Value - Sheets(4).Cells(a - 1, 2).Value + 2 * dist_va_max
        Sheets(1).Cells(h - 4, 33).Value = pk1 - Sheets(1).Cells(h - 1, 4).Value - Sheets(1).Cells(h - 3, 4).Value
        Call radio.radio1(h - 4)
        Sheets(1).Cells(h - 2, 33).Value = pk1 - Sheets(1).Cells(h - 1, 4).Value
        Call radio.radio1(h - 2)
        Sheets(1).Cells(h, 33).Value = pk1
        Call radio.radio1(h)
        Sheets(1).Cells(h + 1, 4).Value = vano.vano(Sheets(1).Cells(h, 6).Value)
        dist_restar = pk0 - Sheets(1).Cells(h - 4, 33).Value
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        z = h - 4
        Sheets(1).Cells(h - 6, 16).Value = "Anc.Aigu." '!!!!! No debe ser texto
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Aigu." '!!!!! No debe ser texto
        GoTo salto
    End If
'caso particular de tener un puente despues de la aguja
ElseIf Sheets(4).Cells(a + 1, 1).Value = "Puente" And _
    Sheets(4).Cells(a + 1, 2).Value - Sheets(4).Cells(a, 21).Value < va_max _
    And Sheets(4).Cells(a, 22).Value > dist_va_max Then
    vano12 = Sheets(4).Cells(a + 1, 2).Value - Sheets(4).Cells(a, 2).Value - 2
    vanoref = vano12 + dist_va_max
    dist_restar = pk0 - pk1 - (Sheets(1).Cells(h - 1, 4).Value - vanoref)
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    Sheets(1).Cells(h - 1, 4).Value = vanoref
    Sheets(1).Cells(h, 33).Value = pk1
    Call radio.radio1(h)
    Sheets(1).Cells(h + 1, 4).Value = vano12
    Sheets(1).Cells(h + 2, 33).Value = pk1 + vano12
    Call radio.radio1(h)
    z = h - 2
'para el resto de los casos
Else
    dist_restar = pk0 - pk1
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    z = h
End If
b = 0
'escribir texto para las agujas
    If Sheets(4).Cells(a, 22).Value = "IN" Then
        Sheets(1).Cells(h - 4, 16).Value = "Anc.Aigu." '!!!!! No debe ser texto
salto:
        Sheets(1).Cells(h - 2, 16).Value = "Inter.Aigu." '!!!!! No debe ser texto
        Sheets(1).Cells(h, 16).Value = "Axe.Aigu." '!!!!! No debe ser texto
        Sheets(1).Cells(h + 1, 25).Value = Sheets(4).Cells(a, 23).Value & " - " & Sheets(4).Cells(a, 4).Value
        Sheets(1).Cells(h + 1, 35).Value = Sheets(4).Cells(a, 5).Value
        z_var = h + 1
    Else
        Sheets(1).Cells(h, 16).Value = "Axe.Aigu." '!!!!! No debe ser texto
        Sheets(1).Cells(h + 2, 16).Value = "Inter.Aigu." '!!!!! No debe ser texto
        Sheets(1).Cells(h + 4, 16).Value = "Anc.Aigu." '!!!!! No debe ser texto
        Sheets(1).Cells(h, 25).Value = Sheets(4).Cells(a, 23).Value & " - " & Sheets(4).Cells(a, 4).Value
        Sheets(1).Cells(h + 1, 35).Value = Sheets(4).Cells(a, 5).Value
        z_var = h
    End If
    With Sheets(1).Range(Sheets(1).Cells(z_var, 25), Sheets(1).Cells(z_var, 25))
    .Borders(xlEdgeLeft).LineStyle = 2
    .Borders(xlEdgeLeft).ColorIndex = 15
    .Borders(xlEdgeTop).LineStyle = 2
    .Borders(xlEdgeTop).ColorIndex = 15
    .Borders(xlEdgeBottom).LineStyle = 2
    .Borders(xlEdgeBottom).ColorIndex = 15
    .Borders(xlEdgeRight).LineStyle = 2
    .Borders(xlEdgeRight).ColorIndex = 15
    .MergeCells = True
    End With
      
    'If Sheets(1).Cells(z - 1, 4).Value > Sheets(1).Cells(z + 1, 4).Value And _
        'Sheets(1).Cells(z - 1, 4).Value > Sheets(1).Cells(z - 3, 4).Value And dist_restar > dist_va_max Then
        'Call restar(dist_restar, div, z, h, a)
    'ElseIf Sheets(1).Cells(z - 1, 4).Value >= Sheets(1).Cells(z + 1, 4).Value _
    'And Sheets(1).Cells(z - 1, 4).Value >= Sheets(1).Cells(z - 3, 4).Value And dist_restar > dist_va_max Then
        'Call restar(dist_restar, div, z, h, a)
    'Else
    Call restar(dist_restar, div, z, h, a)
    'End If

'While ((Sheets(1).Cells(z - 3, 4).Value - Sheets(1).Cells(z - 1, 4).Value) >= dist_va_max And _
   '(Sheets(1).Cells(z - 3, 4).Value - Sheets(1).Cells(z - 1, 4).Value) >= -dist_va_max) _
   'And Not (Sheets(1).Cells(z - 3, 4).Value = Sheets(1).Cells(z - 1, 4).Value)
        'z = z - 2
'Wend
    'Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - dist_restar
    'z = z - 2
'While z <= h
    'Sheets(1).Cells(z, 33).Value = Sheets(1).Cells(z - 1, 4) + Sheets(1).Cells(z - 2, 33)
    'Call radio.radio1(z)
    'z = z + 2
'Wend
Sheets(1).Cells(h + 2, 33).Value = Sheets(1).Cells(h + 1, 4) + Sheets(1).Cells(h, 33)
Call radio.radio1(h + 2)
a = a + 1
End Sub
Sub Zona(ByRef h, ByRef a)
Sheets(1).Cells(h, 16).Value = "Anc.Neutre"
Sheets(1).Cells(h - 1, 4).Value = 27
Sheets(1).Cells(h - 2, 16).Value = "Inter.Neutre"
Sheets(1).Cells(h - 3, 4).Value = 27
Sheets(1).Cells(h - 4, 16).Value = "Inter.Neutre"
Sheets(1).Cells(h - 5, 4).Value = 36
Sheets(1).Cells(h - 6, 16).Value = "Axe.Neutre"
Sheets(1).Cells(h - 7, 4).Value = 27
Sheets(1).Cells(h - 8, 16).Value = "Inter.Neutre"
Sheets(1).Cells(h - 9, 4).Value = 27
Sheets(1).Cells(h - 10, 16).Value = "Inter.Neutre"
Sheets(1).Cells(h - 11, 4).Value = 36
Sheets(1).Cells(h - 12, 16).Value = "Anc.Neutre"
Sheets(1).Cells(h - 13, 4).Value = 45
Sheets(1).Cells(h - 15, 4).Value = 54
z = h - 14
While z <> h
Sheets(1).Cells(z, 33).Value = Sheets(1).Cells(z - 1, 4).Value + Sheets(1).Cells(z - 2, 33).Value
z = z + 2
Sheets(1).Cells(z, 25).Value = Sheets(4).Cells(a, 23).Value
Wend
Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
a = a + 1
End Sub

Private Function two(ByRef z, ByRef dist_restar, ByRef n, ByRef h, ByRef div)
Sheets(1).Cells(z, 33).Value = Sheets(1).Cells(z, 33).Value - dist_restar
Call radio.radio1(z)
vano_nuevo = vano.vano(Sheets(1).Cells(z, 6).Value)
If vano_nuevo < Sheets(1).Cells(z + 1, 4).Value Then
    dist_restar = dist_restar - n - (Sheets(1).Cells(z + 1, 4).Value - vano_nuevo)
    div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
    Sheets(1).Cells(z + 1, 4).Value = vano_nuevo
Else
dist_restar = dist_restar - n
End If
If Sheets(1).Cells(z + 3, 4).Value - Sheets(1).Cells(z + 1, 4).Value > 9 And z < h - 2 Then
    Sheets(1).Cells(z + 3, 4).Value = Sheets(1).Cells(z + 3, 4).Value - inc_norm_va
    dist_restar = dist_restar - inc_norm_va
    'Call two(z + 2, dist_restar, inc_norm_va)
End If
End Function
Sub restar(dist_restar, div, z, h, a)
While dist_restar > 0
       
    'if Sheets(1).Cells(z - 1, 4).Value
        'If dist_restar <= 2 * inc_norm_va And Sheets(1).Cells(z - 3, 4).Value - (Sheets(1).Cells(z - 1, 4).Value) = dist_va_max Then
        'ElseIf z = h Then
                'Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value + dist_va_max
                'Call two(z, dist_restar, dist_va_max, h, div)
        If Sheets(1).Cells(z - 1, 4).Value - (Sheets(1).Cells(z + 1, 4).Value) > dist_va_max And _
            z >= h - 4 And dist_restar > dist_va_max Then
            While Sheets(1).Cells(z - 1, 4).Value - (Sheets(1).Cells(z + 1, 4).Value) > dist_va_max
                Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - dist_va_max
                Call two(z, dist_restar, dist_va_max, h, div)
            Wend
        ElseIf Abs(Sheets(1).Cells(z - 1, 4).Value - (Sheets(1).Cells(z + 1, 4).Value)) >= dist_va_max And _
            z >= h And dist_restar > inc_norm_va Then
                Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - inc_norm_va
                Call two(z, dist_restar, inc_norm_va, h, div)
        ElseIf Sheets(1).Cells(z - 1, 4).Value >= (Sheets(1).Cells(z - 3, 4).Value) And dist_restar >= dist_va_max _
        And Sheets(1).Cells(z - 1, 4).Value >= Sheets(1).Cells(z + 1, 4).Value And Sheets(1).Cells(z - 1, 4).Value >= (va_max - 2 * inc_norm_va) Then
            Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - dist_va_max
            Call two(z, dist_restar, dist_va_max, h, div)
        ElseIf dist_restar < inc_norm_va And Sheets(1).Cells(z + 1, 4).Value - Sheets(1).Cells(z - 1, 4).Value < dist_va_max And Sheets(1).Cells(z - 3, 4).Value - Sheets(1).Cells(z - 1, 4).Value < div Then
            Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - div
            dist_restar = dist_restar - div
        ElseIf dist_restar > inc_norm_va And Sheets(1).Cells(z + 1, 4).Value - Sheets(1).Cells(z - 1, 4).Value <= dist_va_max And Sheets(1).Cells(z - 3, 4).Value - Sheets(1).Cells(z - 1, 4).Value <= dist_va_max _
        And (Sheets(1).Cells(z - 3, 4).Value - Sheets(1).Cells(z - 1, 4).Value <= inc_norm_va Or dist_restar > dist_va_max * 1.5) Then
            Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - inc_norm_va
            Call two(z, dist_restar, inc_norm_va, h, div)
        ElseIf dist_restar = inc_norm_va And Sheets(1).Cells(z + 1, 4).Value - Sheets(1).Cells(z - 1, 4).Value < dist_va_max And Sheets(1).Cells(z - 3, 4).Value - Sheets(1).Cells(z - 1, 4).Value < dist_va_max Then
            Sheets(1).Cells(z - 1, 4).Value = Sheets(1).Cells(z - 1, 4).Value - inc_norm_va
            Call two(z, dist_restar, inc_norm_va, h, div)
        End If
        z = z - 2
    Wend
While z <= h
    Sheets(1).Cells(z, 33).Value = Sheets(1).Cells(z - 1, 4) + Sheets(1).Cells(z - 2, 33)
    Call radio.radio1(z)
    'Call punto_singular.sing1(z, a - 1)
    z = z + 2
Wend
End Sub




