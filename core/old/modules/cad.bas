Attribute VB_Name = "cad"
Sub posicion()
h = 10
While Not IsEmpty(Sheets(1).Cells(h, 33).Value)
    a = 3
    While Sheets(1).Cells(h, 33).Value >= Sheets(5).Cells(a, 2).Value And Not IsEmpty(Sheets(5).Cells(a, 2).Value)
        a = a + 1
    Wend
    Sheets(1).Cells(h, 30).Value = Sheets(5).Cells(a, 3).Value

    
    h = h + 2
Wend
End Sub
Sub esfuerzo()
Dim Ry As Double, T As Double, V0 As Double, V As Double, tres As Double, res As Double
Dim cont As Double, j As Integer
h = 12
T = 2328
tres = 2328
'fin = 65550
cont = -1
j = 34
While Not IsEmpty(Sheets(1).Cells(h + 2, 33).Value)
If Sheets(1).Cells(h, 16).Value = "Axe.Antich." Then
    tres = 2328
End If
If Sheets(1).Cells(h, 16).Value = "Anc.Chevau." Or Sheets(1).Cells(h, 16).Value = "Anc.Section." Then
    cont = cont + 1
End If
If cont = 2 Then
    cont = -1
    h = h - 8
    tres = 2328
    If j = 34 Then
        j = 35
    Else: j = 34
    End If
End If
    V0 = Sheets(1).Cells(h - 1, 4).Value
    V = Sheets(1).Cells(h + 1, 4).Value
    R0 = Sheets(1).Cells(h - 2, 6).Value
    r = Sheets(1).Cells(h, 6).Value
    R1 = Sheets(1).Cells(h + 2, 6).Value
    des0 = Sheets(1).Cells(h - 2, 8).Value
    des = Sheets(1).Cells(h, 8).Value
    des1 = Sheets(1).Cells(h + 2, 8).Value

If IsEmpty(Sheets(1).Cells(h, 6).Value) Then
    Ry = T * (((des + des0) / V0) + ((des + des1) / V))
    res = T - ((Sqr((T ^ 2) - (Ry ^ 2))))
    tres = tres - res
    Sheets(1).Cells(h, j).Value = tres
    h = h + 2
Else
    Ry = T * ((((V0 ^ 2) + ((r + des) ^ 2) - ((r + des0) ^ 2)) / (2 * V0 * (r + des))) + (((V ^ 2) + ((r + des) ^ 2) - ((r + des1) ^ 2)) / (2 * V * (r + des))))
    If Ry < 0 Then
        Ry = Ry * (-1)
    End If
    res = T - ((Sqr((T ^ 2) - (Ry ^ 2))))
    tres = tres - res
    Sheets(1).Cells(h, j).Value = tres
    h = h + 2
End If
If tres <= 2136 Then
    Sheets(1).Cells(h - 2, 36).Value = "cambio"
End If
Wend
Ry = 0
End Sub
