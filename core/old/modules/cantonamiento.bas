Attribute VB_Name = "cantonamiento"
'//
'// Declaración de variables publicas (para durante)
'//
Public INI As Double, finish As Double
Public tipo_singular_out As String, tipo_singular_in As String
Public hini As Integer, acan As Integer
Sub canton_durante(ByRef h, ByRef C, ByRef k, ByRef a)

If h = 10 Then
        Sheets(1).Cells(h, 16).Value = "Anc.Chevau."
        INI = Sheets(1).Cells(h, 33).Value
        hini = h
        tipo_singular_in = "inicio"
        acan = 3
End If

    While (Sheets(4).Cells(acan, 1).Value <> "Tunel" And Sheets(4).Cells(acan, 22).Value <> "IN" And Sheets(4).Cells(acan, 23).Value <> "FINAL" _
     And Sheets(4).Cells(acan, 22).Value <> "OUT" And Sheets(4).Cells(acan, 1).Value <> "Zona") Or (Sheets(4).Cells(acan, 2).Value < Sheets(1).Cells(hini, 33).Value) _
     'Or (Sheets(4).Cells(acan, 2).Value < Sheets(1).Cells(h, 33).Value)
     
        acan = acan + 1
        tipo_singular_out = Sheets(4).Cells(acan, 1).Value
        INI = Sheets(1).Cells(hini, 33).Value
        finish = Sheets(4).Cells(acan, 2).Value
    Wend
   
    Select Case tipo_singular_out
        Case Is = "Tunel"
            If tipo_singular_in = tipo_singular_out And Sheets(4).Cells(acan + 2, 1).Value <> "Aguja" Then
                tipo_singular_out = Sheets(4).Cells(acan, 1).Value
                finish = Sheets(4).Cells(acan, 2).Value
                tipo_singular_in = "inicio"
             ElseIf tipo_singular_in = tipo_singular_out And Sheets(4).Cells(acan + 2, 1).Value = "Aguja" Then
                acan = acan + 2
                tipo_singular_out = Sheets(4).Cells(acan, 1).Value
                finish = Sheets(4).Cells(acan, 2).Value
            ElseIf tipo_singular_out = tipo_singular_in _
            And Sheets(4).Cells(acan, 21).Value - Sheets(4).Cells(acan, 2).Value < dist_max_canton Then
                acan = acan + 1
            
            End If
        Case Is = "Aguja"
            If tipo_singular_in = "Aguja" And ((Sheets(4).Cells(acan, 2).Value + 325) - Sheets(4).Cells(hini, 33).Value > dist_max_canton) _
             And (finish >= Sheets(4).Cells(acan - 1, 2).Value And finish <= Sheets(4).Cells(acan, 2).Value) Then
                finish = Sheets(1).Cells(hini, 33).Value + (((Sheets(4).Cells(acan, 2).Value + 432) - Sheets(1).Cells(hini, 33).Value) / 2)
            ElseIf tipo_singular_in = "inicio" And finish > Sheets(4).Cells(acan, 2).Value Then
                finish = Sheets(4).Cells(acan, 2).Value + 432

            ElseIf tipo_singular_in = "Tunel" Then
                finish = Sheets(4).Cells(acan, 2).Value
            
            End If
        Case Is = "Zona"
            finish = Sheets(4).Cells(acan, 2).Value + 40
    End Select


If INI + dist_max_canton < Sheets(1).Cells(h, 33).Value Or finish < Sheets(1).Cells(h, 33).Value Then

    'tipo_singular_in = "inicio"
    If finish < Sheets(1).Cells(hini, 33).Value + 600 And Sheets(4).Cells(acan, 1).Value = "Aguja" Then
        tipo_singular_in = Sheets(4).Cells(acan, 1).Value
        acan = acan + 1
    ElseIf finish < Sheets(1).Cells(hini, 33).Value + 600 And Sheets(4).Cells(acan, 1).Value = "Tunel" Then
            tipo_singular_in = Sheets(4).Cells(acan, 1).Value
        'acan = acan + 1
    ElseIf INI + dist_max_canton < Sheets(1).Cells(h, 33).Value Then
        marca = 0
        z = h
        Call one(z, tipo_singular_in)
        Call pintar(z - 2, hini, 27)
        Call two(z, tipo_singular_in, tipo_singular_out, acan, marca)
    ElseIf finish < Sheets(1).Cells(h, 33).Value Then
        z = h
        Select Case tipo_singular_out
            Case Is = "Aguja"
                If Sheets(4).Cells(acan, 22).Value = "IN" Then
                    z = h - 10
                Else
                    z = h
                End If
            Case Is = "Zona"
                z = h + 2
                tipo_singular_in = Sheets(4).Cells(acan, 1).Value
            Case Is = "Tunel"
                'tipo_singular_out = Sheets(4).Cells(acan, 1).Value
        End Select
        marca = 1
        Call one(z, tipo_singular_in)
        Call pintar(z - 2, hini, 27)
        Call two(z, tipo_singular_in, tipo_singular_out, acan, marca)
        If Sheets(4).Cells(acan, 22).Value <> "OUT" Then
            acan = acan + 1
            tipo_singular_out = Sheets(4).Cells(acan, 1).Value
        Else
            finish = Sheets(4).Cells(acan, 2).Value + 325
        End If

    End If
End If
End Sub
'//
'// Rutina destinada a realizar el cantonamiento al final del replanteo
'//
Sub canton_final(nombre_cat, fin)
Dim total As Double, ncanton As Double, lcanton As Double, corte As Double, prin As Double, error As Double
Dim a As Integer, h As Integer, hini As Integer, contador As Integer
Dim algo As Double
Dim resultado As Integer
'//
'// inicializar variables
'//
Sheets(1).Activate
'Call cargar.datos_acces(nombre_cat)
h = 10
z = 10
a = 4
alfa = 4
alfa1 = 4
beta = h

If h = 10 Then
        Sheets(1).Cells(h, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h + 2, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h + 4, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h + 6, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h + 8, 16).Value = "Anc.Chevau."
        INI = Sheets(1).Cells(h, 33).Value
        prin = Sheets(1).Cells(h, 33).Value
        hini = h
End If
tipo_singular_in = "inicio"

'//
'// Inicio de la rutina, realizar hasta encontrar una celda vacia
'//
While Not IsEmpty(Sheets(1).Cells(h, 33).Value) And Not IsEmpty(Sheets(1).Cells(beta, 33).Value)
    '//
    '// Encontrar puntos singulares (tuneles, agujas y zonas neutras)
    '//
    While (Sheets(4).Cells(a, 1).Value <> "Tunel" And Sheets(4).Cells(a, 22).Value <> "IN" And Sheets(4).Cells(a, 23).Value <> "FINAL" _
    And Sheets(4).Cells(a, 22).Value <> "OUT" And Sheets(4).Cells(a, 1).Value <> "Zona") Or (Sheets(4).Cells(a, 2).Value < Sheets(1).Cells(hini, 33).Value)
        a = a + 1
    Wend
    '//
    '// Inicializar variables locales
    '//
    tipo_singular_out = Sheets(4).Cells(a, 1).Value
    INI = Sheets(1).Cells(hini, 33).Value
    prin = Sheets(1).Cells(hini, 33).Value
    '//
    '// Escoger el tratamiento del tramo siguiente
    '//
    Select Case tipo_singular_out
        Case Is = "Tunel"
            If tipo_singular_in = tipo_singular_out And Sheets(4).Cells(a + 1, 1).Value = "Aguja" Then
                a = a + 1
                tipo_singular_out = Sheets(4).Cells(a, 1).Value
             ElseIf tipo_singular_in = tipo_singular_out And Sheets(4).Cells(a + 2, 1).Value = "Aguja" Then
                a = a + 2
                tipo_singular_out = Sheets(4).Cells(a, 1).Value
            End If
    End Select
    beta = h
    '//
    '// Encontrar puntos singulares siguiente (tuneles, agujas y zonas neutras)
    '//
    While (Sheets(1).Cells(beta, 33).Value < Sheets(4).Cells(a, 2).Value _
     Or (tipo_singular_out = "Tunel" And tipo_singular_in = "Tunel" And Sheets(1).Cells(beta, 33).Value < Sheets(4).Cells(a, 21).Value)) _
     And Not IsEmpty(Sheets(1).Cells(beta, 33).Value)
        beta = beta + 2
    Wend
    '//
    '// Calcular y escoger el Pk final del seccionamiento
    '//

    If Sheets(4).Cells(a, 22).Value = "OUT" Then
        final2 = Sheets(1).Cells(beta + 16, 33).Value
    ElseIf Sheets(4).Cells(a, 22).Value = "IN" Then
        final2 = Sheets(1).Cells(beta - 8, 33).Value
    ElseIf tipo_singular_in = "Tunel" And tipo_singular_out = "Tunel" Then
        final2 = Sheets(1).Cells(beta + 10, 33).Value
    ElseIf Sheets(4).Cells(a, 1).Value = "Tunel" And Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(a, 2).Value _
    And (Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value > dist_max_canton) Then
        final2 = Sheets(4).Cells(a + 2, 2).Value
    ElseIf Sheets(4).Cells(a, 1).Value = "Tunel" Then
        final2 = Sheets(1).Cells(beta - 2, 33).Value
    ElseIf tipo_singular_out = "Zona" Then
        final2 = Sheets(1).Cells(beta - 6, 33).Value
    End If
    total = final2 - INI
    ncanton1 = (total \ dist_max_canton) + 1
    lcanton = ((total) / ncanton1)
    ncanton = 0
    '//
    '// Calcular el numero de seccionamientos y la longitud de cada uno de ellos
    '//
    While ncanton1 <> ncanton
        z = hini
        total = final2 - INI
        ncanton = ncanton1
        lcanton = ((total) / ncanton)
        corte = INI + lcanton
        '//
        '// Calcular el incremento de distancia en los seccionamientos
        '//
        While Sheets(1).Cells(z, 33).Value <= final2 And Sheets(1).Cells(z, 33).Value < fin
            If Val(corte) < Val(Sheets(1).Cells(z, 33).Value) Then
                total = total + (Sheets(1).Cells(z - 2, 33).Value - Sheets(1).Cells(z - 10, 33).Value) + (Sheets(1).Cells(z, 33).Value - corte) + 10
                corte = corte + lcanton
            End If
            z = z + 2
        Wend
        ncanton1 = (total \ dist_max_canton) + 1
    Wend
    '//
    '// Calcular la longitud media del cantonamiento, el final del seccionamiento y punto fijo
    '//
    lcanton = ((total) / ncanton)
    corte = INI + lcanton
    fijo = corte - (lcanton / 2)
    '//
    '// Avanzar hasta llegar al final del seccionamiento o final del tramo
    '//
    While Sheets(1).Cells(h, 33).Value <= final2 And Sheets(1).Cells(h, 33).Value < fin
        '//
        '// Seccionamientos
        '//
        If corte <= (Sheets(1).Cells(h, 33).Value) Then
        '//
        '// Escribir información de cantonamiento: longitud cantonamiento, longitud punto fijo,
        '// inicialización siguiente cantonamiento, escribir tipo de seccionamiento en 4 vanos.
        '//
        Select Case ncanton
           Case Is > 2
                h = h - 2
                Call ancla(alfa, h)
                Call pintar(h, hini, 27)
                Call pintar(h, hfijo - 1, 26)
                Sheets(1).Cells(h, 27).Value = Sheets(1).Cells(h, 33).Value - prin
                Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h, 27).Value
                Sheets(1).Cells(h, 26).Value = Sheets(1).Cells(h, 33).Value - puntofijo
                Sheets(1).Cells(hfijo - 1, 26).Value = Sheets(1).Cells(h, 26).Value
                corte = Sheets(1).Cells(h - 8, 33).Value + lcanton
                prin = Sheets(1).Cells(h - 8, 33).Value
                hini = h - 8
                fijo = prin + (lcanton / 2)
                ncanton = ncanton - 1
                '//
                '// caso particular para cantonamiento en tunel largo
                '//
                If tipo_singular_in = "Tunel" Then
                    Sheets(1).Cells(h, 16).Value = "Anc.Chevau.sans AT"
                    Sheets(1).Cells(h - 2, 16).Value = "Inter.Chevau."
                    Sheets(1).Cells(h - 4, 16).Value = "Axe.Chevau."
                    Sheets(1).Cells(h - 6, 16).Value = "Inter.Chevau."
                    Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau.sans AT"
                Else
                    Sheets(1).Cells(h, 16).Value = "Anc.Chevau."
                    Sheets(1).Cells(h - 2, 16).Value = "Inter.Chevau."
                    Sheets(1).Cells(h - 4, 16).Value = "Axe.Chevau."
                    Sheets(1).Cells(h - 6, 16).Value = "Inter.Chevau."
                    Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau."
                End If
                Sheets(1).Range(Sheets(1).Cells(h, 26), Sheets(1).Cells(h, 26)).Interior.ColorIndex = 8
                Sheets(1).Range(Sheets(1).Cells(hfijo - 1, 26), Sheets(1).Cells(hfijo - 1, 26)).Interior.ColorIndex = 8
                h = h + 2
            Case Is = 2
                h = h - 2
                Call ancla(alfa, h)
                Call pintar(h, hini, 27)
                Call pintar(h, hfijo - 1, 26)
                Sheets(1).Cells(h, 27).Value = Sheets(1).Cells(h, 33).Value - prin
                Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h, 27).Value
                Sheets(1).Cells(h, 26).Value = Sheets(1).Cells(h, 33).Value - puntofijo
                Sheets(1).Cells(hfijo - 1, 26).Value = Sheets(1).Cells(h, 26).Value
                prin = Sheets(1).Cells(h - 8, 33).Value
                hini = h - 8
                Sheets(1).Cells(h, 16).Value = "Anc.Chevau."
                Sheets(1).Cells(h - 2, 16).Value = "Inter.Chevau."
                Sheets(1).Cells(h - 4, 16).Value = "Axe.Chevau."
                Sheets(1).Cells(h - 6, 16).Value = "Inter.Chevau."
                Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau."
    
                Select Case tipo_singular_out
                    Case Is = "Tunel"
                        '//
                        '// caso particular para cantonamiento dentro tunel largo
                        '//
                        If tipo_singular_out = tipo_singular_in Then
                            Sheets(1).Cells(h, 16).Value = "Anc.Chevau.sans AT"
                            Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau.sans AT"
                        End If
                        lcanton = final2 - prin
                    Case Is = "Aguja"
                        '//
                        '// caso particular para cantonamiento con final en aguja
                        '//
                        If Sheets(4).Cells(a, 22).Value = "IN" Then
                            lcanton = final2 - prin
                        '//
                        '// caso particular para cantonamiento entre agujas
                        '//
                        ElseIf Sheets(4).Cells(a, 22).Value = "OUT" Then
                            lcanton = final2 - prin
                        End If
                        '//
                        '// caso particular para cantonamiento entre tunel y aguja
                        '//
                        If tipo_singular_in = "Tunel" Then
                            Sheets(1).Cells(h, 16).Value = "Anc.Chevau.sans AT"
                            Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau.sans AT"
                        End If
                    Case Is = "Zona"
                        lcanton = final2 - prin
                End Select
                corte = Sheets(1).Cells(h - 8, 33).Value + lcanton
                fijo = prin + (lcanton / 2)
                ncanton = ncanton - 1
                Sheets(1).Range(Sheets(1).Cells(h, 26), Sheets(1).Cells(h, 26)).Interior.ColorIndex = 8
                Sheets(1).Range(Sheets(1).Cells(hfijo - 1, 26), Sheets(1).Cells(hfijo - 1, 26)).Interior.ColorIndex = 8
                h = h + 2
            Case Is = 1
                
                Call ancla(alfa, h)
                Call pintar(h, hini, 27)

                        Sheets(1).Cells(h, 27).Value = Sheets(1).Cells(h, 33).Value - prin
                        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h, 27).Value
                Select Case tipo_singular_out
                
                    Case Is = "Zona"
                        ncanton = ncanton - 1
                        hini = h - 12
                        a = a + 1
                    Case Is = "Tunel"
                        Sheets(1).Cells(h, 16).Value = "Anc.Chevau."
                        Sheets(1).Cells(h - 2, 16).Value = "Inter.Chevau."
                        Sheets(1).Cells(h - 4, 16).Value = "Axe.Chevau."
                        Sheets(1).Cells(h - 6, 16).Value = "Inter.Chevau."
                        Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau."
                        ncanton = ncanton - 1
                        hini = h - 8
                        prin = Sheets(1).Cells(h - 8, 33).Value
                        'If tipo_singular_in = tipo_singular_out And Sheets(4).Cells(a, 22).Value > dist_max_canton Then
                            'Sheets(1).Cells(h, 16).Value = "Anc.Chevau.sans AT"
                        'ElseIf tipo_singular_out = "Tunel" And Sheets(4).Cells(a, 22).Value > dist_max_canton Then
                            'Sheets(1).Cells(h - 8, 16).Value = "Anc.Chevau.sans AT"
                        'End If
                        If tipo_singular_in = tipo_singular_out Then
                            a = a + 1
                        End If
                    Case Is = "Aguja"
                        Sheets(1).Cells(h, 16).Value = "Anc.Section."
                        Sheets(1).Cells(h - 2, 16).Value = "Inter.Section."
                        Sheets(1).Cells(h - 4, 16).Value = "Axe.Section."
                        Sheets(1).Cells(h - 6, 16).Value = "Inter.Section."
                        Sheets(1).Cells(h - 8, 16).Value = "Anc.Section."
                        ncanton = ncanton - 1
                        hini = h - 8
                        prin = Sheets(1).Cells(h - 8, 33).Value
                        'If tipo_singular_in = "Tunel" And (Sheets(4).Cells(a - 1, 22).Value > dist_max_canton Or _
                        'Sheets(4).Cells(a - 2, 22).Value > dist_max_canton) Then
                            'Sheets(1).Cells(h, 16).Value = "Anc.Section.sans AT"
                        'End If
                        
                        If Sheets(4).Cells(a, 22).Value = "IN" Then
                            a = a + 1
                        End If
                    End Select
                If lcanton >= 700 Then '' !!!!!!!!!Introducir por variable
                    Call pintar(h, hfijo - 1, 26)
                    Sheets(1).Range(Sheets(1).Cells(h, 26), Sheets(1).Cells(h, 26)).Interior.ColorIndex = 8
                    Sheets(1).Range(Sheets(1).Cells(hfijo - 1, 26), Sheets(1).Cells(hfijo - 1, 26)).Interior.ColorIndex = 8
                    Sheets(1).Cells(h, 26).Value = Sheets(1).Cells(h, 33).Value - puntofijo
                    Sheets(1).Cells(hfijo - 1, 26).Value = Sheets(1).Cells(h, 26).Value
                Else
                    Sheets(1).Cells(h, 16).Value = "Anc.Chevau.sans AT"
                
                End If
                    tipo_singular_in = Sheets(4).Cells(a, 1).Value
                End Select
     
            End If
        '//
        '// Punto fijo
        '// Escribir información de punto fijo: longitud, longitud punto fijo,
        '// inicialización siguiente punto fijo, escribir tipo de punto fijo
        '//
        If fijo <= (Sheets(1).Cells(h, 33).Value) Then
                com1 = ""
                com2 = ""
                com3 = ""
            If lcanton <= 700 Then '!!!!!!!!!Introducir por variable
                hfijo = h
            Else
                If Not IsEmpty(Sheets(1).Cells(h, 16).Value) Then
                    com1 = " + " & Sheets(1).Cells(h, 16).Value
                End If
                If Not IsEmpty(Sheets(1).Cells(h - 2, 16).Value) Then
                    com2 = " + " & Sheets(1).Cells(h - 2, 16).Value
                End If
                If Not IsEmpty(Sheets(1).Cells(h - 4, 16).Value) Then
                    com3 = " + " & Sheets(1).Cells(h - 4, 16).Value
                End If
                Call ancla2(alfa, h, corte)
                If tipo_singular_in = "Tunel" Then
                    Sheets(1).Cells(h - 2, 16).Value = "Axe.Antich." & com2
                Else
                    Sheets(1).Cells(h - 4, 16).Value = "Anc.Antich." & com3
                    Sheets(1).Cells(h - 2, 16).Value = "Axe.Antich." & com2
                    Sheets(1).Cells(h, 16).Value = "Anc.Antich." & com1
                End If
                hfijo = h
                fijo = Sheets(1).Cells(h - 2, 33).Value + 2000
                puntofijo = Sheets(1).Cells(h - 2, 33).Value
                Sheets(1).Cells(h - 3, 26).Value = puntofijo - prin
                Sheets(1).Range(Sheets(1).Cells(h - 3, 26), Sheets(1).Cells(h - 3, 26)).Interior.ColorIndex = 6
                Sheets(1).Cells(hini, 26).Value = Sheets(1).Cells(h - 3, 26).Value
                Call pintar(h - 3, hini, 26)
                Sheets(1).Range(Sheets(1).Cells(hini, 26), Sheets(1).Cells(hini, 26)).Interior.ColorIndex = 6
            End If
        End If
        h = h + 2
    Wend
Wend
End Sub
'///
'///Rutina destinada a encontrar anclajes de puntos fijos que caen encima de puntos singulares
'///
Sub ancla2(ByRef alfa, ByRef h As Integer, ByRef corte)

While Sheets(1).Cells(h - 4, 33).Value > Sheets(4).Cells(alfa, 21).Value
    alfa = alfa + 1
Wend

If Sheets(1).Cells(h - 4, 33).Value - Sheets(4).Cells(alfa - 1, 21).Value <= 8.5 And Not Sheets(1).Cells(h - 4, 38).Value = "Tunel" Then
    h = h - 2
    corte = corte - Sheets(1).Cells(h - 1, 4).Value
End If
While Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(alfa, 2).Value
    alfa = alfa + 1
Wend
If Sheets(4).Cells(alfa, 2).Value - Sheets(1).Cells(h, 33).Value <= 8.5 And Not Sheets(1).Cells(h, 38).Value = "Tunel" Then
    h = h - 2
    corte = corte - Sheets(1).Cells(h - 1, 4).Value
End If
End Sub
'///
'///Rutina destinada a encontrar anclajes de seccionamientos que caen encima de puntos singulares
'///
Sub ancla(ByRef alfa, ByRef h As Integer)

While Sheets(1).Cells(h - 8, 33).Value > Sheets(4).Cells(alfa, 21).Value
    alfa = alfa + 1
Wend

If Sheets(1).Cells(h - 8, 33).Value - Sheets(4).Cells(alfa - 1, 21).Value <= 8.5 And Not Sheets(1).Cells(h - 8, 38).Value = "Tunel" Then
    h = h - 2
End If
While Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(alfa, 2).Value
    alfa = alfa + 1
Wend
If Sheets(4).Cells(alfa, 2).Value - Sheets(1).Cells(h, 33).Value <= 8.5 And Not Sheets(1).Cells(h, 38).Value = "Tunel" Then
    h = h - 2
End If
End Sub
'//
'// Rutina destinada a dibujar flechas para aclarar los cantonamientos
'//
Sub pintar(ByRef z, zini, columna)

Top = Sheets(1).Range(Sheets(1).Cells(zini + 1, columna), Sheets(1).Cells(z - 1, columna)).Top
Lef = Sheets(1).Range(Sheets(1).Cells(zini + 1, columna), Sheets(1).Cells(z - 1, columna)).Left
Heigh = Sheets(1).Range(Sheets(1).Cells(zini + 1, columna), Sheets(1).Cells(z - 1, columna)).height
columnW = Sheets(1).Range(Sheets(1).Cells(zini + 1, columna), Sheets(1).Cells(z - 1, columna)).Width
If columna = 27 Then
    If Sheets(1).Range(Sheets(1).Cells(zini + 8, columna), Sheets(1).Cells(zini + 8, columna)).Interior.ColorIndex = 3 Then
        Sheets(1).Range(Sheets(1).Cells(z, columna), Sheets(1).Cells(z, columna)).Interior.ColorIndex = 4
        Sheets(1).Range(Sheets(1).Cells(zini, columna), Sheets(1).Cells(zini, columna)).Interior.ColorIndex = 4

        ActiveSheet.Shapes.AddLine(Lef + (columnW / 3), Top, Lef + (columnW / 3), Top + Heigh).Select
    Else
        Sheets(1).Range(Sheets(1).Cells(z, columna), Sheets(1).Cells(z, columna)).Interior.ColorIndex = 3
        Sheets(1).Range(Sheets(1).Cells(zini, columna), Sheets(1).Cells(zini, columna)).Interior.ColorIndex = 3
        Sheets(1).Shapes.AddLine(Lef + (2 * columnW / 3), Top, Lef + (2 * columnW / 3), Top + Heigh).Select
    End If
Else
    If Sheets(1).Range(Sheets(1).Cells(zini - 2, columna), Sheets(1).Cells(zini - 2, columna)).Interior.ColorIndex = 6 Then

        Sheets(1).Shapes.AddLine(Lef + (columnW / 3), Top, Lef + (columnW / 3), Top + Heigh).Select
    Else

        Sheets(1).Shapes.AddLine(Lef + (2 * columnW / 3), Top, Lef + (2 * columnW / 3), Top + Heigh).Select
    End If
End If

Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
Selection.ShapeRange.Line.EndArrowheadLength = msoArrowheadLengthMedium
Selection.ShapeRange.Line.EndArrowheadWidth = msoArrowheadWidthMedium
Selection.ShapeRange.Flip msoFlipHorizontal
Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
Selection.ShapeRange.Line.BeginArrowheadLength = msoArrowheadLengthMedium
Selection.ShapeRange.Line.BeginArrowheadWidth = msoArrowheadWidthMedium
End Sub
'//
'// Rutina de durante (a pulir)
'//
Private Sub one(z, ByRef tipo_singular_in)
Dim C As Integer
C = 10
h = z
aone = 3
While Sheets(1).Cells(h, 33).Value > Sheets(4).Cells(aone, 21).Value And Sheets(4).Cells(aone, 23).Value <> "FINAL"
    aone = aone + 1
Wend

    While C >= 0 And tipo_singular_in <> "Zona"
        If Sheets(1).Cells(h - C - 1, 4).Value > va_max_sm Then
            Sheets(1).Cells(h - C - 1, 4).Value = va_max_sm
        End If
            Sheets(1).Cells(h - C, 33).Value = Sheets(1).Cells(h - C - 2, 33).Value + Sheets(1).Cells(h - C - 1, 4).Value
            Call radio.radio1(h - C)
            Call punto_singular.sing1(h - C, aone - 1)
        C = C - 2
    Wend

    Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 2, 33).Value + Sheets(1).Cells(h - 1, 4).Value
    Call radio.radio1(h)
    Call punto_singular.sing1(h, aone - 1)
While tipo_singular_in = "Aguja" And Not IsEmpty(Sheets(1).Cells(h, 33).Value)
    Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 2, 33).Value + Sheets(1).Cells(h - 1, 4).Value
    Call radio.radio1(h)
    Call punto_singular.sing1(h, aone - 1)
    h = h + 2
Wend
End Sub
'//
'// Rutina de durante (a pulir)
'//
Sub two(ByRef h, tipo_singular_in, tipo_singular_out, a, marca)
If marca = 0 Then
    If (tipo_singular_out = "Aguja" And tipo_singular_in = "Tunel") And Not IsEmpty(Sheets(1).Cells(h - 2, 25).Value) Or _
    (tipo_singular_in = tipo_singular_out And tipo_singular_in = "Tunel") Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau.sans AT"
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau.sans AT"
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
    ElseIf (tipo_singular_out = "Aguja" And tipo_singular_in = "Tunel") And _
    finish < Sheets(1).Cells(h, 33).Value + 600 Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Section.sans AT"
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Section."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        acan = acan + 1
        tipo_singular_in = "Aguja"
    ElseIf (tipo_singular_out = "Aguja" And tipo_singular_in = "inicio") And _
    finish < Sheets(1).Cells(h, 33).Value + 600 Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Section."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        acan = acan + 1
        tipo_singular_in = "Aguja"
    ElseIf (tipo_singular_out = "Tunel" And (tipo_singular_in = "Aguja" Or tipo_singular_in = "inicio")) And _
    finish < Sheets(1).Cells(h, 33).Value + 600 _
    And Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value > dist_max_canton Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau.sans AT"
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
    ElseIf (tipo_singular_out = "Tunel" And (tipo_singular_in = "Aguja" Or tipo_singular_in = "inicio")) And _
    finish < Sheets(1).Cells(h, 33).Value + 600 _
    And Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value < dist_max_canton Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau"
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        acan = acan + 1
    Else
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
    End If
ElseIf marca = 1 Then
    If tipo_singular_out = "Tunel" _
    And Sheets(4).Cells(a, 21).Value - Sheets(4).Cells(a, 2).Value > dist_max_canton Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau.sans AT"
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        tipo_singular_in = "Tunel"
    ElseIf tipo_singular_out = "Zona" Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Neutre"
        Sheets(1).Cells(h - 3, 4).Value = 27
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Neutre"
        Sheets(1).Cells(h - 5, 4).Value = 27
        Sheets(1).Cells(h - 6, 16).Value = "Inter.Neutre"
        Sheets(1).Cells(h - 7, 4).Value = 36
        Sheets(1).Cells(h - 8, 16).Value = "Axe.Neutre"
        Sheets(1).Cells(h - 9, 4).Value = 27
        Sheets(1).Cells(h - 10, 16).Value = "Inter.Neutre"
        Sheets(1).Cells(h - 11, 4).Value = 27
        Sheets(1).Cells(h - 12, 16).Value = "Inter.Neutre"
        Sheets(1).Cells(h - 13, 4).Value = 36
        Sheets(1).Cells(h - 14, 16).Value = "Anc.Neutre"
        Sheets(1).Cells(h - 15, 4).Value = 45
        
        z = h - 14
        While z <> h
            Sheets(1).Cells(z, 33).Value = Sheets(1).Cells(z - 1, 4).Value + Sheets(1).Cells(z - 2, 33).Value
            Sheets(1).Cells(z, 25).Value = Sheets(4).Cells(a, 23).Value
            z = z + 2
        Wend
        Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 1, 4).Value + Sheets(1).Cells(h - 2, 33).Value
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 14, 33).Value
        hinifijo = hini
        hini = h - 14
    ElseIf tipo_singular_out = "Aguja" And tipo_singular_in = "Tunel" Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Section.sans AT"
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Section."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        tipo_singular_in = "Aguja"

    ElseIf tipo_singular_out = "Aguja" And tipo_singular_in = "inicio" And Sheets(4).Cells(a, 2).Value < Sheets(1).Cells(h, 33).Value Or _
    tipo_singular_out = "inicio" And finish < Sheets(1).Cells(h, 33).Value + 600 Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Section."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        'tipo_singular_in = "Aguja"
    ElseIf tipo_singular_out = "Aguja" And tipo_singular_in = "Aguja" And Sheets(4).Cells(a, 2).Value < Sheets(1).Cells(h, 33).Value Or _
    tipo_singular_in = "inicio" And tipo_singular_out = "Aguja" And finish < Sheets(1).Cells(h, 33).Value + 600 Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Section."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Section."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Section."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        tipo_singular_in = "Aguja"
    ElseIf tipo_singular_out = "Aguja" And tipo_singular_in = "Aguja" Then
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
        tipo_singular_in = "inicio"

    Else
        Sheets(1).Cells(h - 2, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 4, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 6, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(h - 8, 16).Value = "Inter.Chevau."
        Sheets(1).Cells(h - 10, 16).Value = "Anc.Chevau."
        Sheets(1).Cells(h - 2, 27).Value = Sheets(1).Cells(h - 2, 33).Value - INI
        Sheets(1).Cells(hini, 27).Value = Sheets(1).Cells(h - 2, 27).Value
        inifijo = INI
        INI = Sheets(1).Cells(h - 10, 33).Value
        hinifijo = hini
        hini = h - 10
    End If
End If
    hfijo = h
If Sheets(1).Cells(h - 2, 27).Value > 600 Then
    While Sheets(1).Cells(h - 2, 33).Value - (Sheets(1).Cells(h - 2, 27).Value / 2) < Sheets(1).Cells(hfijo, 33).Value
        hfijo = hfijo - 2
    Wend
    Sheets(1).Cells(hfijo - 2, 16).Value = "Anc.Antich."
    Sheets(1).Cells(hfijo, 16).Value = "Axe.Antich."
    Sheets(1).Cells(hfijo + 2, 16).Value = "Anc.Antich."
    Sheets(1).Cells(hfijo - 2, 26).Value = Sheets(1).Cells(hfijo, 33).Value - inifijo
    Sheets(1).Cells(hfijo + 2, 26).Value = Sheets(1).Cells(h - 2, 33).Value - Sheets(1).Cells(hfijo, 33).Value
    Sheets(1).Cells(hinifijo, 26).Value = Sheets(1).Cells(hfijo - 2, 26).Value
    Sheets(1).Cells(h - 2, 26).Value = Sheets(1).Cells(hfijo + 2, 26).Value
    Sheets(1).Range(Sheets(1).Cells(hfijo - 2, 26), Sheets(1).Cells(hfijo - 2, 26)).Interior.ColorIndex = 6
    Sheets(1).Range(Sheets(1).Cells(hinifijo, 26), Sheets(1).Cells(hinifijo, 26)).Interior.ColorIndex = 6
    Sheets(1).Range(Sheets(1).Cells(h - 2, 26), Sheets(1).Cells(h - 2, 26)).Interior.ColorIndex = 8
    Sheets(1).Range(Sheets(1).Cells(hfijo + 2, 26), Sheets(1).Cells(hfijo + 2, 26)).Interior.ColorIndex = 8
    Sheets(1).Cells(hini, 26).Value = Sheets(1).Cells(h - 3, 26).Value
    Call pintar(hfijo - 2, hinifijo, 26)
    Call pintar(hfijo + 3, h - 3, 26)
End If

End Sub

