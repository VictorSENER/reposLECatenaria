Attribute VB_Name = "altura"
'//
'// Rutina destinada a calcular los incrementos de altura en puntos singulares
'//
Sub altura(nombre_catVB)
Dim Amx As Double, Amx1 As Double
Dim z As Integer, aloc As Integer, s As Integer
'//
'// Cargar variables catenaria
'//
Call cargar.datos_lac(nombre_catVB)
'//
'// inicializar variables locales
'//
z = 10
aloc = 3
While Sheets("Punto singular").Cells(aloc, 2).Value < Sheets("Replanteo").Cells(z, 33).Value
    aloc = aloc + 1
Wend
'//
'// inicio de la rutina principal
'//
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
    '//
    '// Encontrar el punto singular para cada PK
    '//
    While Sheets("Punto singular").Cells(aloc, 23).Value <> "FINAL" And Sheets("Punto singular").Cells(aloc, 1).Value <> "Tunel" And Sheets("Punto singular").Cells(aloc, 1).Value <> "7 > P.S. > 5,2 m" And Sheets("Punto singular").Cells(aloc, 1).Value <> "P.N." And Sheets("Punto singular").Cells(aloc, 1).Value <> "Marquesina"
        aloc = aloc + 1
    Wend
    '//
    '// Actuar si pasamos por un paso a nivel (se debe ir incrementando la altura hasta
    '// la altura máxima y decrementarla hasta la nominal al pasar el paso a nivel)
    '//
    If Sheets("Punto singular").Cells(aloc, 1) = ("P.N.") And Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Replanteo").Cells(z - 2, 33).Value < Sheets("Punto singular").Cells(aloc, 2).Value Then
        '//
        '// Inicializar altura a valor maximo en PK actual
        '// Calculo del primer decremento (mas pequeño) hacia PK anteriores
        '//
        s = z
        Amx = alt_max
        s = s - 2
        Sheets("Replanteo").Cells(s, 10).Value = alt_max
        Amx = alt_max - Abs((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) / 100) * 100
        '//
        '// Calculo de la altura a decrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx >= alt_nom And Not IsEmpty(Sheets("Replanteo").Cells(s, 33).Value)
            s = s - 2
            Sheets("Replanteo").Cells(s, 10).Value = Amx
            Amx = Amx - Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                If Amx < alt_nom Then
                    Amx1 = Amx + Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                        If ((Amx1 - alt_nom) * 1000) / Sheets("Replanteo").Cells(s - 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s - 2
                            Amx1 = Amx1 - Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                            Sheets("Replanteo").Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
        '//
        '// Inicializar altura a valor maximo en PK actual
        '// Calculo del primer decremento (mas pequeño) hacia PK siguientes
        '//
        s = z
        Sheets("Replanteo").Cells(s, 10).Value = alt_max
        Amx = alt_max - Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a decrementar hacia PK siguientes hasta llegar a la altura nominal
        '//
        While Amx >= alt_nom And Not IsEmpty(Sheets("Replanteo").Cells(s, 33).Value)
            s = s + 2
            Sheets("Replanteo").Cells(s, 10).Value = Amx
            Amx = Amx - Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                If Amx < alt_nom Then
                    Amx1 = Amx + Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                        If ((Amx1 - alt_nom) * 1000) / Sheets("Replanteo").Cells(s + 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s + 2
                            Amx1 = Amx1 - Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                            Sheets("Replanteo").Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
        z = s
        aloc = aloc + 1
    '//
    '// Actuar si pasamos por un paso superior bajo (se debe ir decrementando la altura hasta
    '// la altura mínima al llegar al paso e incrementar la altura hasta la nominal una vez pasado)
    '//
    ElseIf (Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Replanteo").Cells(z, 33).Value <= Sheets("Punto singular").Cells(aloc, 21).Value _
    And Sheets("Punto singular").Cells(aloc, 1) = "7 > P.S. > 5,2 m") Or (Sheets("Replanteo").Cells(z - 2, 33).Value < Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(aloc, 2).Value And _
    Sheets("Punto singular").Cells(aloc, 1) = "7 > P.S. > 5,2 m") Then
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas pequeño) hacia PK anteriores
        '//
        s = z
        s = s - 2
        Sheets("Replanteo").Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a incrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx <= alt_nom And Not IsEmpty(Sheets("Replanteo").Cells(s, 33).Value)
            s = s - 2
            Sheets("Replanteo").Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                If Amx > alt_nom Then
                    Amx1 = Amx - ((inc_max_alt_hc * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000)
                        If ((alt_nom - Amx1) * 1000) / Sheets("Replanteo").Cells(s - 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s - 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                            Sheets("Replanteo").Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas pequeño) hacia PK siguientes
        '//
        s = z
        Sheets("Replanteo").Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a incrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx <= alt_nom And Not IsEmpty(Sheets("Replanteo").Cells(s, 33).Value)
            s = s + 2
            Sheets("Replanteo").Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                If Amx > alt_nom Then
                    Amx1 = Amx - Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                        If ((alt_nom - Amx1) * 1000) / Sheets("Replanteo").Cells(s + 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s + 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                            Sheets("Replanteo").Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
        z = s
        aloc = aloc + 1
    '//
    '// Actuar si estamos dentro del tunel
    '//
    ElseIf (Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Replanteo").Cells(z, 33).Value <= Sheets("Punto singular").Cells(aloc, 21).Value _
    And (Sheets("Punto singular").Cells(aloc, 1) = "Tunel" Or Sheets("Punto singular").Cells(aloc, 1) = "Marquesina")) Then
    '//
    '// Actualizar altura a valor minimo
    '//
         Sheets("Replanteo").Cells(z, 10).Value = alt_min
        If Sheets("Punto singular").Cells(aloc, 1) = "Tunel" Then
            Sheets("Replanteo").Cells(z, 38).Value = "Tunel"
        End If
        '//
        '// Actuar en salida de tunel
        '//
        If (Sheets("Replanteo").Cells(z + 2, 33).Value >= Sheets("Punto singular").Cells(aloc, 21).Value) Then 'And Sheets("Punto singular").Cells(aloc + 1, 1).Value <> "Tunel" Then 'And Sheets("Replanteo").Cells(z - 2, 33).Value <= Sheets("Punto singular").Cells(aloc - 1, 21).Value
        'And (Sheets("Punto singular").Cells(aloc - 1, 1) = "Tunel" Or Sheets("Punto singular").Cells(aloc - 1, 1) = "Marquesina")) Then
            '//
            '// Inicializar altura a valor minimo en PK actual
            '// Calculo del primer incremento (mas pequeño) hacia PK siguientes
            '//
            s = z
            Sheets("Replanteo").Cells(s, 10).Value = alt_min
            Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
    
            '//
            '// Calculo de la altura a incrementar hacia PK siguientes hasta llegar a la altura nominal
            '//
            While (Amx <= alt_nom And Not IsEmpty(Sheets("Replanteo").Cells(s, 33).Value)) And Not (Sheets("Replanteo").Cells(s + 4, 33).Value > Sheets("Punto singular").Cells(aloc + 1, 2).Value And Sheets("Punto singular").Cells(aloc + 1, 1).Value = "Tunel")
                'If (Sheets("Replanteo").Cells(s + 4, 33).Value > Sheets("Punto singular").Cells(aloc + 1, 2).Value And Sheets("Punto singular").Cells(aloc + 1, 1).Value = "Tunel") Then
                    'algo = 0
                'Else
                s = s + 2
                Sheets("Replanteo").Cells(s, 10).Value = Amx
                Amx = Amx + Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                    If Amx > alt_nom Then
                        Amx1 = Amx - Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                            If ((alt_nom - Amx1) * 1000) / Sheets("Replanteo").Cells(s + 1, 4).Value > (inc_max_alt_hc / 2) Then
                                s = s + 2
                                Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s + 1, 4).Value) / 1000) * 100) / 100
                                Sheets("Replanteo").Cells(s, 10).Value = Amx1
                            End If
                    End If
                'End If
            Wend
            z = s
            aloc = aloc + 1
            
        End If
    '//
    '// Actuar en entrada de tunel
    '//
    ElseIf (Sheets("Replanteo").Cells(z, 33).Value <= Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Replanteo").Cells(z + 2, 33).Value > Sheets("Punto singular").Cells(aloc, 2).Value _
    And (Sheets("Punto singular").Cells(aloc, 1) = "Tunel" Or Sheets("Punto singular").Cells(aloc, 1) = "Marquesina")) Then
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas pequeño) hacia PK anteriores
        '//
        s = z + 2
        Sheets("Replanteo").Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a incrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While (Amx <= alt_nom And Not IsEmpty(Sheets("Replanteo").Cells(s, 33).Value)) And Not (Amx >= Sheets("Replanteo").Cells(s - 2, 10).Value And Not IsEmpty(Sheets("Replanteo").Cells(s - 2, 10).Value))
            'If Amx >= Sheets("Replanteo").Cells(s - 2, 10).Value And Not IsEmpty(Sheets("Replanteo").Cells(s - 2, 10).Value) Then
                'algo = 0
            'Else
            s = s - 2
            Sheets("Replanteo").Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                If Amx >= alt_nom Then
                    Amx1 = Amx - ((inc_max_alt_hc * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000)
                        If ((alt_nom - Amx1) * 1000) / Sheets("Replanteo").Cells(s - 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s - 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets("Replanteo").Cells(s - 1, 4).Value) / 1000) * 100) / 100
                            Sheets("Replanteo").Cells(s, 10).Value = Amx1
                        End If
                End If
            'End If
        Wend
    Else
        Sheets("Replanteo").Cells(z, 10).Value = alt_nom
    End If
'//
'// Incrementar una fila del replanteo
'//
z = z + 2
Wend
End Sub


