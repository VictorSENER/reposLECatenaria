Attribute VB_Name = "revision"
Sub revision()
Dim pk1 As Double, vano0 As Double, vano1 As Double, vanox As Double
Dim z As Long, aloc As Long, bloc As Long, d As Long
Dim l1 As Double, l2 As Double, rady As Double, inicio1 As Double, inicio2 As Double
Dim radioC As Double, dev As Double, pk0 As Double, final As Double, radioC1 As Double
Dim radioC2 As Double, rady2 As Double, A1 As Double, A12 As Double, A2 As Double, C As Double
Dim contador As Integer
z = 10
aloc = 3
bloc = 2
kloc = 3
d = 1

While Not IsEmpty(Sheets(1).Cells(z + 1, 4).Value)
    pk0 = Sheets(1).Cells(z, 33).Value
    pk1 = Sheets(1).Cells(z + 2, 33).Value
    vano0 = Sheets(1).Cells(z - 1, 4).Value
    vano1 = Sheets(1).Cells(z + 1, 4).Value
    vanox = vano.vano(Sheets(1).Cells(z, 6).Value)
    contador = 0

    ' encontrar el punto singular siguiente
    While pk0 >= Sheets(4).Cells(aloc, 21).Value
        aloc = aloc + 1
    Wend
    ' calcular el radio

While contador <> 2
        pk0 = Sheets(1).Cells(z, 33).Value
        final = Sheets(3).Cells(kloc + 1, 3).Value
        inicio1 = Sheets(3).Cells(kloc, 3).Value
' buscar en que fila de la hoja 3 cae el PK actual
        While (pk0 < inicio1 Or pk0 > final) And pk0 > inicio1 And final <> 0
            kloc = kloc + 1
            final = Sheets(3).Cells(kloc + 1, 3).Value
            inicio1 = Sheets(3).Cells(kloc, 3).Value
        Wend
        If kloc = 3 Then
        
        Else
        rady2 = Sheets(3).Cells(kloc - 1, 2).Value
        A12 = Sheets(3).Cells(kloc - 1, 9).Value
        End If
        l1 = Sheets(3).Cells(kloc, 11).Value
        l2 = Sheets(3).Cells(kloc, 12).Value
        rady = Sheets(3).Cells(kloc, 2).Value
        
        inicio1 = Sheets(3).Cells(kloc, 3).Value
        inicio2 = Sheets(3).Cells(kloc, 5).Value
        A1 = Sheets(3).Cells(kloc, 9).Value
        
        A2 = Sheets(3).Cells(kloc, 10).Value
        C = Sheets(3).Cells(kloc, 7).Value

' calculo del radio correspondiente aloc la clotoide de entrada
        If pk0 >= inicio1 And pk0 < (inicio1 + l1) And Not IsEmpty(Sheets(3).Cells(kloc - 1, 6).Value) Then
            radioC = A1 ^ 2 / (pk0 - inicio1)
            If rady < 0 Then
                radioC = radioC * (-1)
            End If
            If contador = 0 Then
                If IsEmpty(Sheets(1).Cells(z - 2, 6).Value) Then
                    dev = (C * 1000) / radioC
                    'pk = vano(Sheets(1).Cells(z - 2, 6).Value) + Sheets(1).Cells(z - 2, 33).Value
                    contador = contador + 1
                Else
                    If Sheets(1).Cells(z - 2, 6).Value < 0 Then
                        radioC1 = Sheets(1).Cells(z - 2, 6).Value * (-1)
                    End If
                    If radioC < 0 Then
                        radioC2 = radioC * (-1)
                    End If
                    If radioC1 >= radioC2 Then
                        dev = (C * 1000) / radioC
                        ' pk = vano(Sheets(1).Cells(z - 2, 6).Value) + Sheets(1).Cells(z - 2, 33).Value
                        contador = contador + 1
                    Else
                        'pk = Sheets(1).Cells(z - 1, 4).Value + Sheets(1).Cells(z - 2, 33).Value
                        contador = contador + 1
                    End If
                End If
            ElseIf contador = 1 Then
                dev = (C * 1000) / radioC
                contador = contador + 1
            Else
                contador = 2
            End If
                
' calculo del radio correspondiente aloc la curva
ElseIf pk0 >= (inicio1 + l1) And pk0 < inicio2 Then
    radioC = rady
    If rady < 0 Then
        rady = rady * (-1)
    End If
    If Sheets(1).Cells(z - 2, 6).Value < 0 Then
        radio_0 = Sheets(1).Cells(z - 2, 6).Value * (-1)
    ElseIf Sheets(1).Cells(z - 2, 6).Value > 0 Then
        radio_0 = Sheets(1).Cells(z - 2, 6).Value
    Else
        radio_0 = r_re
    End If
    If inicio = pk0 Then
        dev = (C * 1000) / radioC
        contador = 2
    ElseIf contador = 0 Then
        dev = (C * 1000) / radioC
        contador = contador + 1
    ElseIf rady >= radio_0 Then
        dev = (C * 1000) / radioC
        contador = 2

    ElseIf contador = 1 Then
        dev = (C * 1000) / radioC
        contador = contador + 1
    Else
        contador = 2
    End If

' calculo del radio correspondiente aloc la clotoide de salida
ElseIf pk0 > inicio2 And pk0 < inicio2 + l2 Then
    radioC = A2 ^ 2 / (l2 - (pk0 - inicio2))
    contador = 2
    If rady < 0 Then
        radioC = radioC * (-1)
    End If
    dev = (C * 1000) / radioC

' calculo del radio correspondiente aloc la clotoide entre dos curvas
ElseIf IsEmpty(Sheets(3).Cells(kloc - 1, 6).Value) And pk0 < (inicio2 + l2) Then
    contador = 2
    If rady < 0 Then
        radioC1 = rady * (-1)
    Else: radioC1 = rady
    End If
    If rady2 < 0 Then
        radioC2 = rady2 * (-1)
    Else: radioC2 = rady2
    End If
    If radioC1 < radioC2 Then
        lmin = A1 ^ 2 / radioC1
        radioC = (A1 ^ 2) / (lmin - ((inicio1 + l1) - pk0))
        If rady < 0 Then
            radioC = radioC * (-1)
        End If
    Else
        lmin = A1 ^ 2 / radioC2
        radioC = (A1 ^ 2) / (lmin - (pk0 - inicio1))
        If rady2 < 0 Then
            radioC = radioC * (-1)
        End If
    End If
        dev = (C * 1000) / radioC
Else
contador = 2
End If
Wend
If radioC < 0 Then
    radiofinal = radioC * (-1)
Else
    radiofinal = radioC
End If
If radiofinal > 15000 Then
    radioC = 0
End If
    ' verificaci?n de las variaciones del vano
    If ((vano0 - dist_va_max) > vano1 Or (vano0 + dist_va_max) < vano1) And z > 10 Then
        Sheets(1).Cells(z + 1, 4).Font.Color = vbRed
        Sheets(7).Cells(bloc, 1).Value = d
        Sheets(7).Cells(bloc, 2).Value = "Error en la verificaci?n de la diferencia entre vanos"
        Sheets(7).Cells(bloc, 3).Value = pk0
        Sheets(7).Cells(bloc, 4).Value = z + 1
        Sheets(7).Cells(bloc, 5).Value = 4
        bloc = bloc + 1
    End If
    ' verificaci?n del incremento del pk
    If pk1 <> (pk0 + vano1) Then
        Sheets(1).Cells(z + 2, 33).Font.Color = vbRed
        Sheets(7).Cells(bloc, 1).Value = d
        Sheets(7).Cells(bloc, 2).Value = "Error en la verificaci?n del incremento del PK"
        Sheets(7).Cells(bloc, 3).Value = pk0
        Sheets(7).Cells(bloc, 4).Value = z
        Sheets(7).Cells(bloc, 5).Value = 33
        bloc = bloc + 1
    End If
    ' verificaci?n ubicaci?n sobre puntos singulares

    Select Case Sheets(4).Cells(aloc, 1).Value
        Case Is = "PuenteXL", "Puente", "P.S. > 7 m", "7 > P.S. > 5,2 m", "Conducto", "P.N.", "P.I.", "Drenaje"
            If pk0 > Sheets(4).Cells(aloc, 2).Value And pk0 < Sheets(4).Cells(aloc, 21).Value Then
                Sheets(1).Cells(z, 33).Font.Color = vbBlue
                Sheets(7).Cells(bloc, 1).Value = d
                Sheets(7).Cells(bloc, 2).Value = "Error en la verificaci?n de la ubicaci?n sobre puntos singulares"
                Sheets(7).Cells(bloc, 3).Value = pk0
                Sheets(7).Cells(bloc, 4).Value = z
                Sheets(7).Cells(bloc, 5).Value = 33
                bloc = bloc + 1
                d = d + 1
            End If
        Case Is = "Aguja"
            If pk0 = Sheets(4).Cells(aloc, 2).Value Then
                Sheets(7).Cells(bloc, 2).Value = "Aguja: " & Sheets(4).Cells(aloc, 3).Value & " instalada correctamente"
                bloc = bloc + 1
            End If
    End Select
    'verificaci?n del vano
    If vanox < Sheets(1).Cells(z + 1, 4).Value Then
        Sheets(1).Cells(z + 1, 4).Font.Color = vbBlue
        Sheets(7).Cells(bloc, 1).Value = d
        Sheets(7).Cells(bloc, 2).Value = "Error en la verificaci?n del vano respecto aloc su radio"
        Sheets(7).Cells(bloc, 3).Value = pk0
        Sheets(7).Cells(bloc, 4).Value = z + 1
        Sheets(7).Cells(bloc, 5).Value = 4
        Sheets(7).Cells(bloc, 6).Value = vanox
        bloc = bloc + 1
        d = d + 1
    End If
    ' verificaci?n del radio
    
    If radioC = Sheets(1).Cells(z - 2, 6).Value Then
        If Sheets(7).Cells(bloc - 1, 7).Value = "R" And Sheets(7).Cells(bloc - 1, 4).Value = z - 2 Then
            Sheets(1).Cells(z - 2, 6).Font.Color = vbBlack
            bloc = bloc - 1
            d = d - 1
            Sheets(7).Cells(bloc, 1).EntireRow.Delete
        ElseIf bloc > 2 Then
        
            If Sheets(7).Cells(bloc - 2, 7).Value = "R" And Sheets(7).Cells(bloc - 2, 4).Value = z - 2 Then
                Sheets(1).Cells(z - 2, 6).Font.Color = vbBlack
                bloc = bloc - 1
                d = d - 1
                Sheets(7).Cells(bloc, 1).Value = d - 1
                Sheets(7).Cells(bloc - 1, 1).EntireRow.Delete
            End If
        End If
    End If
    If radioC <> Sheets(1).Cells(z, 6).Value Then
        Sheets(1).Cells(z, 6).Font.Color = vbRed
        Sheets(7).Cells(bloc, 1).Value = d
        Sheets(7).Cells(bloc, 2).Value = "Error en la verificaci?n del radio respecto aloc su PK"
        Sheets(7).Cells(bloc, 3).Value = pk0
        Sheets(7).Cells(bloc, 4).Value = z
        Sheets(7).Cells(bloc, 5).Value = 6
        Sheets(7).Cells(bloc, 6).Value = radioC
        Sheets(7).Cells(bloc, 7).Value = "R"
        bloc = bloc + 1
        d = d + 1
    End If

    z = z + 2
    radioC = 0
Wend
'If bloc > 2 Then
    'AVISO.Show
    'Sheets(7).Activate
'End If
End Sub

