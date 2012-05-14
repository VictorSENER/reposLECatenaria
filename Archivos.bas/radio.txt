Attribute VB_Name = "radio"
'//
'// Rutina destinada a encontrar la ubicación del PK respecto a su trazado (clotoide entrada,
'// radio, clotoide de salida o clotoide entre curvas)
'//
Sub radio(ByRef h)
Dim l1 As Double, l2 As Double, rady As Double, inicio1 As Double, inicio2 As Double
Dim radioC As Double, dev As Double, pk0 As Double, final As Double, radioC1 As Double
Dim radioC2 As Double, rady2 As Double, A1 As Double, A12 As Double, A2 As Double, C As Double
Dim contador As Integer
'//
'// Inicializar variables
'//
k = 3
contador = 0
'//
'// Mientras el bucle sea menor o igual a 2. Necesario realizar el calculo dos veces en clotoides de entrada,
'// curvas y tramos entre clotoides.
'//
While contador <> 2
    '//
    '// Inicializar variables generales
    '//
    pk0 = Sheets(1).Cells(h, 33).Value
    final = Sheets(3).Cells(k + 1, 3).Value
    inicio1 = Sheets(3).Cells(k, 3).Value
    '//
    '// Buscar en que fila de la hoja 3 cae el PK actual
    '//
    While (pk0 < inicio1 Or pk0 > final) And pk0 > inicio1 And final <> 0
        k = k + 1
        final = Sheets(3).Cells(k + 1, 3).Value
        inicio1 = Sheets(3).Cells(k, 3).Value
    Wend
    '//
    '// Inicializar variables del trazado
    '//
    l1 = Sheets(3).Cells(k, 11).Value
    l2 = Sheets(3).Cells(k, 12).Value
    rady = Sheets(3).Cells(k, 2).Value
    inicio1 = Sheets(3).Cells(k, 3).Value
    inicio2 = Sheets(3).Cells(k, 5).Value
    A1 = Sheets(3).Cells(k, 9).Value
    A2 = Sheets(3).Cells(k, 10).Value
    C = Sheets(3).Cells(k, 7).Value
    '//
    '// Recoger datos del radio anterior. Si es nulo lo consideramos como recta
    '//
    If Sheets(1).Cells(h - 2, 6).Value <> 0 And k <> 3 Then
        radio_0 = Abs(Sheets(1).Cells(h - 2, 6).Value)
    Else
        radio_0 = r_re
    End If
    '//
    '// Caso inicial (coincide con la primera fila). No existen datos anteriores
    '//
    If k <> 3 Then
        A12 = Sheets(3).Cells(k - 1, 9).Value
        rady2 = Sheets(3).Cells(k - 1, 2).Value
    End If
    '//
    '// calculo del radio correspondiente a la clotoide de entrada
    '//
    If pk0 >= inicio1 And pk0 < (inicio1 + l1) And Not IsEmpty(Sheets(3).Cells(k - 1, 6).Value) Then
        radioC = A1 ^ 2 / (pk0 - inicio1)
        '//
        '// calculo si radio calculado menor a radio considerado como recta
        '//
        If radioC < r_re Then
            If rady < 0 Then
                radioC = radioC * (-1)
            End If
            '//
            '// calculo caso inicial (pk actual = inicio)
            '//
            If inicio = pk0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                contador = 2
            '//
            '// calculo caso primer bucle no realizado
            '//
            ElseIf contador = 0 Then
                '//
                '// calculo caso celda anterior vacia
                '//
                If IsEmpty(Sheets(1).Cells(h - 2, 6).Value) Then
                    Sheets(1).Cells(h - 2, 6).Value = radioC
                    dev = (C * 1000) / radioC
                    Sheets(1).Cells(h - 2, 7).Value = dev
                    Sheets(1).Cells(h - 1, 4).Value = vano.vano(Sheets(1).Cells(h - 2, 6).Value)
                    Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
                    contador = contador + 1
                '//
                '// calculo caso celda anterior llena
                '//
                Else
                    radioC1 = Abs(Sheets(1).Cells(h - 2, 6).Value)
                    radioC2 = Abs(radioC)
                    '//
                    '// calculo caso radio anterior mayor a radio calculado
                    '//
                    If radioC1 >= radioC2 Then
                        Sheets(1).Cells(h - 2, 6).Value = radioC
                        dev = (C * 1000) / radioC
                        Sheets(1).Cells(h - 2, 7).Value = dev
                        Sheets(1).Cells(h - 1, 4).Value = vano.vano(Sheets(1).Cells(h - 2, 6).Value)
                        Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
                        contador = contador + 1
                    '//
                    '// calculo caso radio anterior menor a radio calculado
                    '//
                    Else
                        Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
                        contador = contador + 1
                    End If
                End If
            '//
            '// calculo caso primer bucle realizado
            '//
            ElseIf contador = 1 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                '//
                '// calculo caso radio anterior mayor a radio calculado
                '//
                If radioC1 >= radioC2 Then
                    Sheets(1).Cells(h - 2, 6).Value = radioC
                    Sheets(1).Cells(h - 2, 7).Value = dev
                End If
                contador = contador + 1
            End If
        '//
        '// calculo si radio calculado mayor a radio considerado como recta
        '//
            Else
                contador = 2
        End If
    '//
    '// calculo del radio correspondiente a la curva
    '//
    ElseIf pk0 >= (inicio1 + l1) And pk0 < inicio2 Then
        radioC = rady
        If rady < 0 Then
            rady = rady * (-1)
        End If
        '//
        '// calculo si radio calculado menor a radio considerado como recta
        '//
        If rady < r_re Then
            '//
            '// calculo caso inicial
            '//
            If inicio = pk0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            '//
            '// calculo caso radio calculado mayor radio anterior
            '//
            ElseIf rady >= radio_0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            '//
            '// calculo caso primer bucle no realizado
            '//
            ElseIf contador = 0 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                Sheets(1).Cells(h - 1, 4).Value = vano.vano(Sheets(1).Cells(h - 2, 6).Value)
                pk0 = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
                Sheets(1).Cells(h, 33).Value = pk0
                contador = contador + 1
            '//
            '// calculo caso primer bucle realizado
            '//
            ElseIf contador = 1 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = contador + 1
            End If
        '//
        '// calculo si radio calculado mayor a radio considerado como recta
        '//
        Else
            contador = 2
        End If
    '//
    '// calculo del radio correspondiente a la clotoide de salida
    '//
    ElseIf pk0 > inicio2 And pk0 < inicio2 + l2 Then
        radioC = A2 ^ 2 / (l2 - (pk0 - inicio2))
        contador = 2
        '//
        '// calculo si radio calculado menor a radio considerado como recta
        '//
        If radioC < r_re Then
            If rady < 0 Then
                radioC = radioC * (-1)
            End If
            Sheets(1).Cells(h, 6).Value = radioC
            dev = (C * 1000) / radioC
            Sheets(1).Cells(h, 7).Value = dev
        End If
    '//
    '// calculo del radio correspondiente a la clotoide entre dos curvas
    '//
    ElseIf IsEmpty(Sheets(3).Cells(k - 1, 6).Value) And pk0 < (inicio2 + l2) Then
        radioC1 = Abs(rady)
        radioC2 = Abs(rady2)
        '//
        '// Elección del radio menor de las dos clotoides
        '//
        If radioC1 < radioC2 And l1 <> 0 Then
            lmin = A1 ^ 2 / radioC1
            radiot = (A1 ^ 2) / (lmin - ((inicio1 + l1) - pk0))
            If rady < 0 Then
                radioC = radiot * (-1)
            Else: radioC = radiot
            End If
        Else
            lmin = A1 ^ 2 / radioC2
            radiot = (A1 ^ 2) / (lmin - (pk0 - inicio1))
            If rady2 < 0 Then
                radioC = radiot * (-1)
            Else: radioC = radiot
            End If
        End If
        '//
        '// calculo si radio calculado menor a radio considerado como recta y
        '// radio clotoide 1 menor a radio clotoide 2
        '//
        If radiot < r_re And radioC1 < radioC2 Then
            '//
            '// calculo caso inicial
            '//
            If inicio = pk0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            '//
            '// calculo especifico radio calculado = 0
            '//
            ElseIf radiot = 0 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC2
                dev = (C * 1000) / radioC2
                Sheets(1).Cells(h - 2, 7).Value = dev
                contador = contador + 1
            '//
            '// calculo caso radio calculado mayor a radio anterior y primer bucle no realizado
            '//
            ElseIf radiot > radio_0 And contador = 0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            '//
            '// calculo caso primer bucle no realizado
            '//
            ElseIf contador = 0 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                Sheets(1).Cells(h - 1, 4).Value = vano.vano(Sheets(1).Cells(h - 2, 6).Value)
                pk0 = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
                Sheets(1).Cells(h, 33).Value = pk0
                contador = contador + 1
            '//
            '// calculo caso primer bucle realizado
            '//
            ElseIf contador = 1 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                contador = contador + 1
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            End If
        '//
        '// calculo si radio calculado menor a radio considerado como recta y
        '// radio clotoide 1 mayor a radio clotoide 2
        '//
        ElseIf rady < r_re And radioC1 > radioC2 Then
            Sheets(1).Cells(h, 6).Value = radioC
            dev = (C * 1000) / radioC
            Sheets(1).Cells(h, 7).Value = dev
            contador = 2
        Else
            contador = 2
        End If
    '//
    '// calculo del radio correspondiente a recta
    '//
    Else
    contador = 2
    End If
Wend
'//
'// indicar el sentido de la curva (para autocad)
'//

If radioC > 0 Then
    Sheets(1).Cells(h, 29).Value = "pos"
ElseIf radioC < 0 Then
    Sheets(1).Cells(h, 29).Value = "neg"
End If
End Sub
'//
'// Rutina destinada a encontrar la ubicación del PK respecto a su trazado (clotoide entrada,
'// radio, clotoide de salida o clotoide entre curvas) y calcular el radio
'//
Sub radio1(ByRef h)
Dim l1 As Double, l2 As Double, rady As Double, inicio1 As Double, inicio2 As Double
Dim radioC As Double, dev As Double, pk0 As Double, final As Double, radioC1 As Double
Dim radioC2 As Double, rady2 As Double, A1 As Double, A12 As Double, A2 As Double, C As Double
Dim contador As Integer
'//
'// Inicializar variables
'//
k = 3
contador = 0
final = Sheets(3).Cells(k + 1, 3).Value
inicio1 = Sheets(3).Cells(k, 3).Value
'//
'// Mientras el bucle sea menor o igual a 2
'//
While contador <> 2
    '//
    '// Inicializar variables generales
    '//
    pk0 = Sheets(1).Cells(h, 33).Value
    final = Sheets(3).Cells(k + 1, 3).Value
    inicio1 = Sheets(3).Cells(k, 3).Value
    '//
    '// Buscar en que fila de la hoja 3 cae el PK actual
    '//
    While (pk0 < inicio1 Or pk0 > final) And pk0 > inicio1 And final <> 0
        k = k + 1
        final = Sheets(3).Cells(k + 1, 3).Value
        inicio1 = Sheets(3).Cells(k, 3).Value
    Wend
    '//
    '// Inicializar variables del trazado
    '//
    l1 = Sheets(3).Cells(k, 11).Value
    l2 = Sheets(3).Cells(k, 12).Value
    rady = Sheets(3).Cells(k, 2).Value
    inicio1 = Sheets(3).Cells(k, 3).Value
    inicio2 = Sheets(3).Cells(k, 5).Value
    A1 = Sheets(3).Cells(k, 9).Value
    A2 = Sheets(3).Cells(k, 10).Value
    C = Sheets(3).Cells(k, 7).Value
    '//
    '// Caso inicial
    '//
    If k <> 3 Then
        A12 = Sheets(3).Cells(k - 1, 9).Value
        rady2 = Sheets(3).Cells(k - 1, 2).Value
    End If
    '//
    '// Recoger dato radio anterior
    '//
    If Sheets(1).Cells(h - 2, 6).Value < 0 Or Sheets(1).Cells(h - 2, 6).Value > 0 _
    And k <> 3 Then
        radio_0 = Abs(Sheets(1).Cells(h - 2, 6).Value)
    Else
        radio_0 = r_re
    End If
    '//
    '// el funcionamiento es igual que en radio
    '// calculo del radio correspondiente a la clotoide de entrada
    '//
    If pk0 >= inicio1 And pk0 < (inicio1 + l1) And Not IsEmpty(Sheets(3).Cells(k - 1, 6).Value) Then
        radioC = A1 ^ 2 / (pk0 - inicio1)
        radioa = Sheets(1).Cells(h - 2, 6).Value
        If Sheets(1).Cells(h - 2, 6).Value < 0 Then
            radioC1 = Sheets(1).Cells(h - 2, 6).Value * (-1)
        Else: radioC1 = Sheets(1).Cells(h - 2, 6).Value
        End If
        If radioC < 0 Then
            radioC2 = radioC * (-1)
        Else: radioC2 = radioC
        End If
        If radioC < r_re Then
            If rady < 0 Then
                radioC = radioC * (-1)
            End If
            If inicio = pk0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                contador = 2
            ElseIf contador = 0 Then
                Sheets(1).Cells(h - 2, 6).Value = ""
                Sheets(1).Cells(h - 2, 7).Value = ""
                If IsEmpty(Sheets(1).Cells(h - 2, 6).Value) Then
                    Sheets(1).Cells(h - 2, 6).Value = radioC
                    dev = (C * 1000) / radioC
                    Sheets(1).Cells(h - 2, 7).Value = dev
                    contador = contador + 1
                End If
            ElseIf contador = 1 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                If radioC1 >= radioC2 Then
                    Sheets(1).Cells(h - 2, 6).Value = radioC
                    Sheets(1).Cells(h - 2, 7).Value = dev
                End If
                contador = contador + 1
            End If
            Else
                If (Sheets(1).Cells(h - 2, 33).Value < Sheets(3).Cells(k - 1, 6).Value _
                And radioC1 < radioC2) Or k = 3 Then
                    Sheets(1).Cells(h - 2, 6).Value = ""
                    Sheets(1).Cells(h - 2, 7).Value = ""
                
                ElseIf (radioC1 <> 0 And radioC1 < radioC2) Then
                    Sheets(1).Cells(h - 2, 6).Value = ""
                    Sheets(1).Cells(h - 2, 7).Value = ""
                End If
                contador = 2
        End If
    '//
    '// calculo del radio correspondiente a la curva
    '//
    ElseIf pk0 >= (inicio1 + l1) And pk0 < inicio2 Then
        radioC = rady
        If rady < 0 Then
            rady = rady * (-1)
        End If
    
        If rady < r_re Then
            If inicio = pk0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            'ElseIf rady > radio_0 And contador = 0 Then
                'Sheets(1).Cells(h, 6).Value = radioC
                'dev = (C * 1000) / radioC
                'Sheets(1).Cells(h, 7).Value = dev
                'contador = 2
            ElseIf contador = 0 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                pk0 = vano.vano(Sheets(1).Cells(h - 2, 6).Value) + Sheets(1).Cells(h - 2, 33).Value
                contador = contador + 1
            ElseIf contador = 1 Then
            Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                pk0 = vano.vano(Sheets(1).Cells(h - 2, 6).Value) + Sheets(1).Cells(h - 2, 33).Value
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = contador + 1
            End If
            Else
                contador = 2
        End If
    '//
    '// calculo del radio correspondiente a la clotoide de salida
    '//
    ElseIf pk0 > inicio2 And pk0 < inicio2 + l2 Then
        radioC = A2 ^ 2 / (l2 - (pk0 - inicio2))
        contador = 2
        If radioC < r_re Then
            If rady < 0 Then
                radioC = radioC * (-1)
            End If
            Sheets(1).Cells(h, 6).Value = radioC
            dev = (C * 1000) / radioC
            Sheets(1).Cells(h, 7).Value = dev
        Else
            Sheets(1).Cells(h, 6).Value = ""
            Sheets(1).Cells(h, 7).Value = ""
        End If
    '//
    '// calculo del radio correspondiente a la clotoide entre dos curvas
    '//
    ElseIf IsEmpty(Sheets(3).Cells(k - 1, 6).Value) And pk0 < (inicio2 + l2) Then
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
            radiot = (A1 ^ 2) / (lmin - ((inicio1 + l1) - pk0))
            If rady < 0 Then
                radioC = radiot * (-1)
            Else: radioC = radiot
            End If
        Else
            lmin = A1 ^ 2 / radioC2
            radiot = (A1 ^ 2) / (lmin - (pk0 - inicio1))
            If rady2 < 0 Then
                radioC = radiot * (-1)
            Else: radioC = radiot
            End If
        End If
        If radioC1 < r_re And radioC1 < radioC2 Then
            If inicio = pk0 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
            ElseIf contador = 0 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                pk0 = vano.vano(Sheets(1).Cells(h - 2, 6).Value) + Sheets(1).Cells(h - 2, 33).Value
                contador = contador + 1
            ElseIf contador = 1 Then
                Sheets(1).Cells(h - 2, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h - 2, 7).Value = dev
                pk0 = vano.vano(Sheets(1).Cells(h - 2, 6).Value) + Sheets(1).Cells(h - 2, 33).Value
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = contador + 1
            End If
        ElseIf rady < r_re And radioC1 > radioC2 Then
                Sheets(1).Cells(h, 6).Value = radioC
                dev = (C * 1000) / radioC
                Sheets(1).Cells(h, 7).Value = dev
                contador = 2
        Else
            contador = 2
        End If
        
    Else
    contador = 2
    Sheets(1).Cells(h, 6).Value = ""
    Sheets(1).Cells(h, 7).Value = ""
        If pk0 + 4.5 >= Sheets(3).Cells(k + 1, 3).Value And _
        pk0 + 4.5 < (Sheets(3).Cells(k + 1, 3).Value + Sheets(3).Cells(k + 1, 11).Value) Then
            Sheets(1).Cells(h - 2, 6).Value = ""
            Sheets(1).Cells(h - 2, 7).Value = ""
        End If
    End If
Wend
'indicar el sentido de la curva (para autocad)
If radioC > 0 Then
Sheets(1).Cells(h, 29).Value = "pos"
ElseIf radioC < 0 Then
Sheets(1).Cells(h, 29).Value = "neg"
End If
End Sub

