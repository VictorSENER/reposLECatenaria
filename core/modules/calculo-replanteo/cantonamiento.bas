Attribute VB_Name = "cantonamiento"
'//
'// Rutina destinada a realizar el cantonamiento al final del replanteo
'//
Sub canton_final(nombre_cat)
Dim total As Double, ncanton As Double, lcanton As Double, corte As Double, prin As Double, error As Double
Dim a As Integer, h As Integer, hini As Integer, contador As Integer
Dim algo As Double
Dim resultado As Integer
'//
'// inicializar variables
'//
Sheets("Replanteo").Activate
Call cargar.datos_lac(nombre_cat)
h = 10
beta = 10
z = 10
a = 4
alfa = 4
If h = 10 Then
    If Sheets("Replanteo").Cells(h + 1, 4).Value < 40.5 And Sheets("Replanteo").Cells(h + 3, 4).Value < 40.5 And Sheets("Replanteo").Cells(h + 5, 4).Value < 40.5 _
    And Sheets("Replanteo").Cells(h + 7, 4).Value < 40.5 Then
        Call com(h + 10, anc_sla_con, semi_eje_sla, eje_sla)
    Else
        Call com(h + 8, anc_sla_con, semi_eje_sla, eje_sla)
    End If
    ini = Sheets("Replanteo").Cells(h, 33).Value
    prin = Sheets("Replanteo").Cells(h, 33).Value
    hini = h
End If
tipo_singular_in = "inicio"

'//
'// Inicio de la rutina, realizar hasta encontrar una celda vacia
'//
aseg = a

While Not IsEmpty(Sheets("Replanteo").Cells(h, 33).Value) And Not IsEmpty(Sheets("Replanteo").Cells(beta + 2, 33).Value)
    '//
    '// Encontrar puntos singulares (tuneles largos y agujas y zonas neutras)
    '//Or Sheets("Punto singular").Cells(a, 22).Value < 300

    
    While ((Sheets("Punto singular").Cells(a, 1).Value <> "Tunel") And Sheets("Punto singular").Cells(a, 1).Value <> "Aguja" And Sheets("Punto singular").Cells(a, 23).Value <> "FINAL" And Sheets("Punto singular").Cells(a, 1).Value <> "Desvío" And Sheets("Punto singular").Cells(a, 1).Value <> "Viaducto" And Sheets("Punto singular").Cells(a, 1).Value <> "Marquesina") _
     Or (Sheets("Punto singular").Cells(a, 2).Value < Sheets("Replanteo").Cells(hini, 33).Value) Or (Sheets("Punto singular").Cells(aseg, 2).Value + 1000 > Sheets("Punto singular").Cells(a, 21).Value And (Sheets("Punto singular").Cells(a, 1).Value = "Viaducto" Or Sheets("Punto singular").Cells(a, 1).Value = "Marquesina"))
        a = a + 1
    Wend
    '//
    '// Inicializar variables locales
    '//

    tipo_singular_out = Sheets("Punto singular").Cells(a, 1).Value

    ini = Sheets("Replanteo").Cells(hini, 33).Value
    prin = Sheets("Replanteo").Cells(hini, 33).Value
    '//
    '// Escoger el tratamiento del tramo siguiente
    '//
    Select Case tipo_singular_out
        Case Is = "Tunel"
            If tipo_singular_in = tipo_singular_out And Sheets("Punto singular").Cells(a + 2, 1).Value = "Aguja" Then
                a = a + 2
                tipo_singular_out = Sheets("Punto singular").Cells(a, 1).Value
             'ElseIf tipo_singular_in = tipo_singular_out And Sheets("Punto singular").Cells(a + 4, 1).Value = "Aguja" Then
                'a = a + 4
                'tipo_singular_out = Sheets("Punto singular").Cells(a, 1).Value
            'ElseIf tipo_singular_in = tipo_singular_out And Sheets("Punto singular").Cells(a + 1, 1).Value = "Tunel" Then
                'a = a + 1
                'tipo_singular_out = Sheets("Punto singular").Cells(a, 1).Value
            End If
        Case Is = ""
            algo = 0
    End Select
    beta = h
    '//
    '// Encontrar correspondencia ubicación de PK y puntos singulares siguiente
    '//
    While (Sheets("Replanteo").Cells(beta, 33).Value < Sheets("Punto singular").Cells(a, 2).Value _
     Or (tipo_singular_out = "Tunel" And tipo_singular_in = "Tunel" And Sheets("Replanteo").Cells(beta, 33).Value < Sheets("Punto singular").Cells(a, 21).Value)) _
     And Not IsEmpty(Sheets("Replanteo").Cells(beta, 33).Value)
        beta = beta + 2
    Wend
    '//
    '// Calcular y escoger el Pk final del seccionamiento
    '//
    aplus = 4
    If Sheets("Punto singular").Cells(a, 22).Value = "OUT" Then
        While Not IsEmpty(Sheets("Replanteo").Cells(beta, 16).Value)
            beta = beta + 2
        Wend
        
        final2 = Sheets("Replanteo").Cells(beta + 10, 33).Value
        
        
        'final2 = Sheets("Replanteo").Cells(beta + 18, 33).Value
        
        '/// Se ha quitado este requisito después de la reunión del 7-1-14 en la ONCF
        'While Sheets("Replanteo").Cells(beta, 33).Value - Sheets("Punto singular").Cells(a, 2).Value <= 350
            'beta = beta + 2
        'Wend
        
        'final2 = Sheets("Replanteo").Cells(beta, 33).Value
        aplus = a
        While Sheets("Punto singular").Cells(aplus, 1).Value <> "Estacion"
            aplus = aplus - 1
        Wend
    ElseIf Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
        While Not IsEmpty(Sheets("Replanteo").Cells(beta, 16).Value)
            beta = beta - 2
        Wend
        If Sheets("Punto singular").Cells(a, 3).Value = "Fes Bab Fetouh" Then
        
            beta = beta + 6
        End If
        final2 = Sheets("Replanteo").Cells(beta, 33).Value
        '/// Se ha quitado este requisito después de la reunión del 7-1-14 en la ONCF
        'While Sheets("Punto singular").Cells(a, 2).Value - Sheets("Replanteo").Cells(beta, 33).Value <= 350
            'beta = beta - 2
        'Wend
        'If Sheets("Replanteo").Cells(beta - 1, 4).Value > 40.5 And Sheets("Replanteo").Cells(beta - 3, 4).Value > 40.5 And Sheets("Replanteo").Cells(beta - 5, 4).Value > 40.5 Then
            'final2 = Sheets("Replanteo").Cells(beta + 8, 33).Value
        'Else
            'final2 = Sheets("Replanteo").Cells(beta + 10, 33).Value
        'End If
        
    'ElseIf (Sheets("Punto singular").Cells(a, 1).Value = "Tunel" And Sheets("Punto singular").Cells(a + 2, 1).Value = "Aguja" And (Sheets("Punto singular").Cells(a, 21).Value - Sheets("Punto singular").Cells(a + 2, 2).Value < 150)) Then
        'a = a + 2
        'final2 = Sheets("Punto singular").Cells(a, 2).Value - 31.5
        'tipo_singular_out = "Aguja"
    ElseIf tipo_singular_in = tipo_singular_out And Sheets("Punto singular").Cells(a + 1, 1).Value = "Tunel" Then
        final2 = Sheets("Replanteo").Cells(beta + 8, 33).Value
    ElseIf tipo_singular_in = "Tunel" And tipo_singular_out = "Tunel" Then
        final2 = Sheets("Replanteo").Cells(beta + 10, 33).Value
    ElseIf Sheets("Punto singular").Cells(a, 1).Value = "Tunel" And Sheets("Replanteo").Cells(h, 33).Value > Sheets("Punto singular").Cells(a, 2).Value _
    And (Sheets("Punto singular").Cells(a, 21).Value - Sheets("Punto singular").Cells(a, 2).Value > dist_max_canton) Then
        final2 = Sheets("Punto singular").Cells(a + 2, 2).Value
    ElseIf Sheets("Punto singular").Cells(a, 1).Value = "Tunel" Then
        final2 = Sheets("Replanteo").Cells(beta - 4, 33).Value
    ElseIf tipo_singular_out = "Zona" Then
        final2 = Sheets("Replanteo").Cells(beta - 6, 33).Value
    ElseIf tipo_singular_out = "Viaducto" Then
        final2 = Sheets("Replanteo").Cells(beta - 4, 33).Value
    ElseIf tipo_singular_out = "Marquesina" Then
        final2 = Sheets("Replanteo").Cells(beta - 2, 33).Value
    Else
        final2 = Sheets("Replanteo").Cells(beta, 33).Value
    End If
    total = final2 - ini
    ncanton1 = (total \ dist_max_canton) + 1
    lcanton = ((total) / ncanton1)
    ncanton = 0

    
    '//
    '//Comprobar que el final del cantonamiento está dentro del tramo a considerar
    '//
    If IsEmpty(final2) Then
        GoTo finalizar
    End If
    '//
    '// Calcular el numero de seccionamientos y la longitud de cada uno de ellos
    '//
    While ncanton1 <> ncanton
        z = hini
        total = final2 - ini
        ncanton = ncanton1
        lcanton = ((total) / ncanton)
        corte = ini + lcanton
        '//
        '// Calcular el incremento de distancia en los seccionamientos
        '//
        While Sheets("Replanteo").Cells(z, 33).Value <= final2 And Sheets("Replanteo").Cells(z, 33).Value < final
            If Val(corte) < Val(Sheets("Replanteo").Cells(z, 33).Value) Then
                If Sheets("Replanteo").Cells(z - 1, 4).Value > 54 And Sheets("Replanteo").Cells(z - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(z - 5, 4).Value >= 54 Then
                    total = total + (Sheets("Replanteo").Cells(z - 2, 33).Value - Sheets("Replanteo").Cells(z - 8, 33).Value) + (Sheets("Replanteo").Cells(z, 33).Value - corte) + 10
                    corte = corte + lcanton
                ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(z - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(z - 5, 4).Value > 31.5 Then
                    total = total + (Sheets("Replanteo").Cells(z - 2, 33).Value - Sheets("Replanteo").Cells(z - 12, 33).Value) + (Sheets("Replanteo").Cells(z, 33).Value - corte) + 10
                    corte = corte + lcanton
                Else
                    total = total + (Sheets("Replanteo").Cells(z - 2, 33).Value - Sheets("Replanteo").Cells(z - 10, 33).Value) + (Sheets("Replanteo").Cells(z, 33).Value - corte) + 10
                    corte = corte + lcanton
                End If
            End If
            z = z + 2
        Wend
        ncanton1 = (total \ dist_max_canton) + 1
    Wend
    '//
    '// Calcular la longitud media del cantonamiento, el final del seccionamiento y punto fijo
    '//
    If tipo_singular_in = "Tunel" Or (final2 - ini < dist_max_canton + 100 And Sheets("Punto singular").Cells(a, 22).Value = "OUT") Then
        ncanton = 1
        total = final2 - ini
        lcanton = total
        corte = ini + lcanton
        fijo = corte - (lcanton / 2)
    Else
        lcanton = ((total) / ncanton)
        corte = ini + lcanton
        fijo = corte - (lcanton / 2)
    End If
    '//
    '// Avanzar hasta llegar al final del seccionamiento o final del tramo
    '//
    While Sheets("Replanteo").Cells(h, 33).Value <= final2 And Sheets("Replanteo").Cells(h, 33).Value < final
        '//
        '// Seccionamientos
        '//
        If corte <= (Sheets("Replanteo").Cells(h, 33).Value) Then
        '///
        '/// Comprobar si el edificio de la estación cae delante de un seccionamiento
        '///
        If Sheets("Replanteo").Cells(h - 2, 33).Value > Sheets("Punto singular").Cells(aplus, 21).Value And (Sheets("Replanteo").Cells(h - 10, 33).Value - 10) < Sheets("Punto singular").Cells(aplus, 2).Value Or _
        (Sheets("Replanteo").Cells(h - 10, 33).Value - 10) < Sheets("Punto singular").Cells(aplus, 21).Value And Sheets("Replanteo").Cells(h - 2, 33).Value > Sheets("Punto singular").Cells(aplus, 21).Value Then
               zeta = h
               While Sheets("Replanteo").Cells(zeta - 10, 33).Value - 10 < Sheets("Punto singular").Cells(aplus, 21).Value
                        zeta = zeta + 2
                Wend
        h = zeta
        End If
        
        
        
        '//
        '// Escribir información de cantonamiento: longitud cantonamiento, longitud punto fijo,
        '// inicialización siguiente cantonamiento, escribir tipo de seccionamiento en 4 vanos.
        '//
        Select Case ncanton
           Case Is > 2
                If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                    Call ancla(alfa, h - 2, corte, h - 6)
                ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                    Call ancla(alfa, h, corte, h - 8)
                Else
                    Call ancla(alfa, h - 2, corte, h - 10)
                End If
                Call pintar(h - 2, hini, 27)
                Call pintar(h - 2, hfijo - 1, 26)
                Call inici(h - 2, prin, puntofijo, hini, hfijo, lcanton, ncanton, final2, corte, fijo)
                '//
                '// caso particular para cantonamiento en tunel largo
                '//
                If tipo_singular_in = "Tunel" Then
                    'sheets("Replanteo").Cells(h - 10, 17).Value = "sin"
                    'sheets("Replanteo").Cells(h - 2, 17).Value = "sin"
                    Call com(h - 2, anc_sm_sin, semi_eje_sm, eje_sm)
                '///
                '/// resto de casos
                '///
                Else
                    Call com(h - 2, anc_sm_con, semi_eje_sm, eje_sm)
                    
                End If
                Sheets("Replanteo").Range(Sheets("Replanteo").Cells(h - 2, 26), Sheets("Replanteo").Cells(h - 2, 26)).Interior.ColorIndex = 8
                Sheets("Replanteo").Range(Sheets("Replanteo").Cells(hfijo - 1, 26), Sheets("Replanteo").Cells(hfijo - 1, 26)).Interior.ColorIndex = 8
            '///
            '/// se realiza la misma tarea que en el caso anterior pero se ajusta el final del último seccionamiento
            '///
            Case Is = 2
                If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                    Call ancla(alfa, h - 2, corte, h - 6)
                ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                    Call ancla(alfa, h, corte, h - 8)
                Else
                    Call ancla(alfa, h, corte, h - 10)
                End If
                Call pintar(h - 2, hini, 27)
                Call pintar(h - 2, hfijo - 1, 26)
                Call inici(h - 2, prin, puntofijo, hini, hfijo, lcanton, ncanton, final2, corte, fijo)
                '//
                '// caso particular para cantonamiento dentro tunel largo
                '//
                If tipo_singular_out = tipo_singular_in And tipo_singular_out = "Tunel" Then
                    'sheets("Replanteo").Cells(h - 10, 17).Value = "sin"
                    'sheets("Replanteo").Cells(h - 2, 17).Value = "sin"
                    Call com(h - 2, anc_sm_sin, semi_eje_sm, eje_sm)
                '///
                '/// caso particular de
                '///
                ElseIf tipo_singular_in = "Tunel" And tipo_singular_out = "Aguja" Then
                    'sheets("Replanteo").Cells(h - 10, 17).Value = "sin"
                    'sheets("Replanteo").Cells(h - 2, 17).Value = "sin"
                    Call com(h - 2, anc_sm_sin, semi_eje_sm, eje_sm)
                Else
                    Call com(h - 2, anc_sm_con, semi_eje_sm, eje_sm)
                End If
                Sheets("Replanteo").Range(Sheets("Replanteo").Cells(h - 2, 26), Sheets("Replanteo").Cells(h - 2, 26)).Interior.ColorIndex = 8
                Sheets("Replanteo").Range(Sheets("Replanteo").Cells(hfijo - 1, 26), Sheets("Replanteo").Cells(hfijo - 1, 26)).Interior.ColorIndex = 8
            Case Is = 1
                If tipo_singular_out = "Aguja" And Sheets("Punto singular").Cells(a, 22).Value = "OUT" Then
                    If Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 7, 4).Value >= 54 _
                    And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 8, 6).Value) Then
                        h = h - 2
                        Call ancla(alfa, h, corte, h - 8)
                    ElseIf Sheets("Replanteo").Cells(h - 3, 4).Value >= 40.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 40.5 And Sheets("Replanteo").Cells(h - 7, 4).Value >= 40.5 Then
                        h = h - 2
                        Call ancla(alfa, h, corte, h - 8)
                    Else
                        Call ancla(alfa, h, corte, h - 10)
                    End If

                Else
                    If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 Then
                        Call ancla(alfa, h, corte, h - 6)
                    ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                        Call ancla(alfa, h, corte, h - 8)
                    Else
                        Call ancla(alfa, h, corte, h - 10)
                    End If
                End If
                Call pintar(h, hini, 27)
                Sheets("Replanteo").Cells(h, 27).Value = Sheets("Replanteo").Cells(h, 33).Value - prin
                Sheets("Replanteo").Cells(hini, 27).Value = Sheets("Replanteo").Cells(h, 27).Value

                Select Case tipo_singular_out
                    '///
                    '/// caso particular salida de tunel
                    '///
                    Case Is = "Tunel"
                        Call com(h, anc_sm_con, semi_eje_sm, eje_sm)
                        ncanton = ncanton - 1
                        If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                        And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                            hini = h - 6
                        ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                            hini = h - 8
                        Else
                            hini = h - 10
                        End If
                        prin = Sheets("Replanteo").Cells(hini, 33).Value
                        aseg = a
                        If tipo_singular_in = tipo_singular_out Then
                            a = a + 1
                        End If
                    '///
                    '/// caso particular aguja
                    '///
                    Case Is = "Aguja"
                        If Sheets("Replanteo").Cells(h, 38).Value = "Tunel" Then
                            Call com(h, anc_sla_sin, semi_eje_sla, eje_sla)
                        Else
                            Call com(h, anc_sla_con, semi_eje_sla, eje_sla)
                        End If
                        ncanton = ncanton - 1
                        If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                        And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                            hini = h - 8
                        ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 40.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 40.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 40.5 Then
                            hini = h - 8
                        Else
                            hini = h - 10
                        End If
                        prin = Sheets("Replanteo").Cells(hini, 33).Value
                        aseg = a
                        If Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
                            a = a + 1
                        End If
                    Case Is = "Desvío"
                         Call com(h, anc_sm_con, semi_eje_sm, eje_sm)
                        ncanton = ncanton - 1
                        If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                        And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                            hini = h - 6
                        ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                            hini = h - 8
                        Else
                            hini = h - 10
                        End If
                        prin = Sheets("Replanteo").Cells(hini, 33).Value
                        aseg = a
                        If Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
                            a = a + 1
                        End If
                    Case Is = "Viaducto"
                         Call com(h, anc_sm_con, semi_eje_sm, eje_sm)
                        ncanton = ncanton - 1
                        If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                        And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                            hini = h - 6
                        ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                            hini = h - 8
                        Else
                            hini = h - 10
                        End If
                        prin = Sheets("Replanteo").Cells(hini, 33).Value
                        'If Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
                            aseg = a
                            a = a + 1
                        'End If
                    Case Is = "Marquesina"
                         Call com(h, anc_sm_con, semi_eje_sm, eje_sm)
                        ncanton = ncanton - 1
                        If Sheets("Replanteo").Cells(h - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 54 _
                        And IsEmpty(Sheets("Replanteo").Cells(h - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(h - 6, 6).Value) Then
                            hini = h - 6
                        ElseIf Sheets("Replanteo").Cells(h - 1, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 3, 4).Value >= 31.5 And Sheets("Replanteo").Cells(h - 5, 4).Value >= 31.5 Then
                            hini = h - 8
                        Else
                            hini = h - 10
                        End If
                        prin = Sheets("Replanteo").Cells(hini, 33).Value
                        aseg = a
                        a = a + 1
                    End Select
                '///
                '/// particularidad para actuar dependiente de si la longitud final es mayor o menor a lg max semicanton
                '///
                If lcanton >= 700 Then '' !!!!!!!!!Introducir por variable
                    Call pintar(h, hfijo - 1, 26)
                    Sheets("Replanteo").Range(Sheets("Replanteo").Cells(h, 26), Sheets("Replanteo").Cells(h, 26)).Interior.ColorIndex = 8
                    Sheets("Replanteo").Range(Sheets("Replanteo").Cells(hfijo - 1, 26), Sheets("Replanteo").Cells(hfijo - 1, 26)).Interior.ColorIndex = 8
                    Sheets("Replanteo").Cells(h, 26).Value = Sheets("Replanteo").Cells(h, 33).Value - puntofijo
                    Sheets("Replanteo").Cells(hfijo - 1, 26).Value = Sheets("Replanteo").Cells(h, 26).Value
                Else
                    If Sheets("Replanteo").Cells(h - 2, 16).Value = semi_eje_sla Then
                        Sheets("Replanteo").Cells(h, 16).Value = anc_sla_sin
                    Else
                        Sheets("Replanteo").Cells(h, 16).Value = anc_sm_sin
                    End If
                End If
                    tipo_singular_in = Sheets("Punto singular").Cells(a, 1).Value
            End Select
     
        End If
        '//
        '// Punto fijo
        '// Escribir información de punto fijo: longitud, longitud punto fijo,
        '// inicialización siguiente punto fijo, escribir tipo de punto fijo
        '//
        If fijo <= (Sheets("Replanteo").Cells(h, 33).Value) Then
                com1 = ""
                com2 = ""
                com3 = ""
            If lcanton <= 700 Then '!!!!!!!!!Introducir por variable
                hfijo = h
            Else
                If ((Sheets("Replanteo").Cells(h - 4, 38).Value = "Tunel" And Sheets("Replanteo").Cells(h - 2, 38).Value <> "Tunel") _
                Or (Sheets("Replanteo").Cells(h - 3, 25).Value = via And Sheets("Replanteo").Cells(h - 1, 25).Value <> via)) Then
                    h = h + 2
                ElseIf (Sheets("Replanteo").Cells(h - 6, 38).Value = "Tunel" And Sheets("Replanteo").Cells(h - 4, 38).Value = "Tunel" And Sheets("Replanteo").Cells(h - 2, 38).Value <> "Tunel") _
                Or (Sheets("Replanteo").Cells(h - 5, 25).Value = via And Sheets("Replanteo").Cells(h - 3, 25).Value = via And Sheets("Replanteo").Cells(h - 1, 25).Value <> via) Then
                    h = h + 4

                End If
                Call ancla(alfa, h, corte, h - 4)
                If Not IsEmpty(Sheets("Replanteo").Cells(h, 16).Value) Then
                    com1 = " + " & Sheets("Replanteo").Cells(h, 16).Value
                    'sheets("Replanteo").Cells(h, 16).Value = anc_pf & com1
                End If
                If Not IsEmpty(Sheets("Replanteo").Cells(h - 2, 16).Value) Then
                    com2 = " + " & Sheets("Replanteo").Cells(h - 2, 16).Value
                    'sheets("Replanteo").Cells(h - 2, 16).Value = eje_pf & com2
                End If
                If Not IsEmpty(Sheets("Replanteo").Cells(h - 4, 16).Value) Then
                    com3 = " + " & Sheets("Replanteo").Cells(h - 4, 16).Value
                    'sheets("Replanteo").Cells(h - 4, 16).Value = anc_pf & com3
                End If
                
                If tipo_singular_in = "Tunel" And (Sheets("Replanteo").Cells(h - 4, 38).Value = "Tunel" And Sheets("Replanteo").Cells(h - 2, 38).Value <> "Tunel") Then
                    Sheets("Replanteo").Cells(h - 2, 16).Value = eje_pf & com2
                Else
                    Sheets("Replanteo").Cells(h - 4, 16).Value = anc_pf & com3
                    Sheets("Replanteo").Cells(h - 4, 24).Value = "C2"
                    Sheets("Replanteo").Cells(h - 2, 16).Value = eje_pf & com2
                    Sheets("Replanteo").Cells(h, 16).Value = anc_pf & com1
                    Sheets("Replanteo").Cells(h, 24).Value = "C2"
                End If
                
                hfijo = h
                fijo = Sheets("Replanteo").Cells(h - 2, 33).Value + 2000
                puntofijo = Sheets("Replanteo").Cells(h - 2, 33).Value
                Sheets("Replanteo").Cells(h - 3, 26).Value = puntofijo - prin
                Sheets("Replanteo").Range(Sheets("Replanteo").Cells(h - 3, 26), Sheets("Replanteo").Cells(h - 3, 26)).Interior.ColorIndex = 6
                Sheets("Replanteo").Cells(hini, 26).Value = Sheets("Replanteo").Cells(h - 3, 26).Value
                Call pintar(h - 3, hini, 26)
                Sheets("Replanteo").Range(Sheets("Replanteo").Cells(hini, 26), Sheets("Replanteo").Cells(hini, 26)).Interior.ColorIndex = 6
            End If
        End If
        Call txt.progress("5", "14", "Distribución de los cantones", Sheets("Replanteo").Cells(h, 33).Value - inicio, final - inicio)

        h = h + 2
    Wend
Wend
finalizar:
End Sub
Sub inici(b, prin, puntofijo, hini, hfijo, lcanton, ncanton, final2, corte, fijo)
Sheets("Replanteo").Cells(b, 27).Value = Sheets("Replanteo").Cells(b, 33).Value - prin
Sheets("Replanteo").Cells(hini, 27).Value = Sheets("Replanteo").Cells(b, 27).Value
Sheets("Replanteo").Cells(b, 26).Value = Sheets("Replanteo").Cells(b, 33).Value - puntofijo
Sheets("Replanteo").Cells(hfijo - 1, 26).Value = Sheets("Replanteo").Cells(b, 26).Value

If Sheets("Replanteo").Cells(b - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(b - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(b - 5, 4).Value >= 54 _
And IsEmpty(Sheets("Replanteo").Cells(b - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(b - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(b - 6, 6).Value) Then
    hini = b - 6
    prin = Sheets("Replanteo").Cells(b - 6, 33).Value
ElseIf Sheets("Replanteo").Cells(b - 1, 4).Value < 31.5 Or Sheets("Replanteo").Cells(b - 3, 4).Value < 31.5 Or Sheets("Replanteo").Cells(b - 5, 4).Value < 31.5 _
Or Sheets("Replanteo").Cells(b - 7, 4).Value < 31.5 Then
    hini = b - 10
    prin = Sheets("Replanteo").Cells(b - 10, 33).Value
Else
    hini = b - 8
    prin = Sheets("Replanteo").Cells(b - 8, 33).Value
End If
If ncanton = 2 Then
    lcanton = final2 - prin
End If
corte = Sheets("Replanteo").Cells(hini, 33).Value + lcanton
fijo = prin + (lcanton / 2)
ncanton = ncanton - 1
End Sub

Sub com(b, anc, semi, eje)

If b = 412 And (Sheets("Replanteo").Cells(b - 1, 4).Value < 40.5 Or Sheets("Replanteo").Cells(b - 3, 4).Value < 40.5 Or Sheets("Replanteo").Cells(b - 5, 4).Value < 40.5 _
Or Sheets("Replanteo").Cells(b - 7, 4).Value < 40.5) And (anc = anc_sla_con Or anc = anc_sla_sin) Then
    Sheets("Replanteo").Cells(b, 16).Value = anc & " + " & semi_eje_aguj
    Sheets("Replanteo").Cells(b, 24).Value = "C2"
    Sheets("Replanteo").Cells(b - 2, 16).Value = semi & " + " & anc_aguj
    Sheets("Replanteo").Cells(b - 4, 16).Value = eje
    Sheets("Replanteo").Cells(b - 6, 16).Value = eje
    Sheets("Replanteo").Cells(b - 8, 16).Value = semi
    Sheets("Replanteo").Cells(b - 10, 16).Value = anc
    Sheets("Replanteo").Cells(b - 10, 24).Value = "C2"

ElseIf Sheets("Replanteo").Cells(b - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(b - 3, 4).Value >= 54 And Sheets("Replanteo").Cells(b - 5, 4).Value >= 54 _
And (anc = anc_sm_con Or anc = anc_sm_sin) And IsEmpty(Sheets("Replanteo").Cells(b - 2, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(b - 4, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(b - 6, 6).Value) Then
    Sheets("Replanteo").Cells(b, 16).Value = anc
    Sheets("Replanteo").Cells(b, 24).Value = "C2"
    Sheets("Replanteo").Cells(b - 2, 16).Value = semi
    Sheets("Replanteo").Cells(b - 4, 16).Value = semi
    Sheets("Replanteo").Cells(b - 6, 16).Value = anc
    Sheets("Replanteo").Cells(b - 6, 24).Value = "C2"
ElseIf (Sheets("Replanteo").Cells(b - 1, 4).Value < 40.5 Or Sheets("Replanteo").Cells(b - 3, 4).Value < 40.5 Or Sheets("Replanteo").Cells(b - 5, 4).Value < 40.5 _
Or Sheets("Replanteo").Cells(b - 7, 4).Value < 40.5) And (anc = anc_sla_con Or anc = anc_sla_sin) Then
    Sheets("Replanteo").Cells(b, 16).Value = anc
    Sheets("Replanteo").Cells(b, 24).Value = "C2"
    Sheets("Replanteo").Cells(b - 2, 16).Value = semi
    Sheets("Replanteo").Cells(b - 4, 16).Value = eje
    Sheets("Replanteo").Cells(b - 6, 16).Value = eje
    Sheets("Replanteo").Cells(b - 8, 16).Value = semi
    Sheets("Replanteo").Cells(b - 10, 16).Value = anc
    Sheets("Replanteo").Cells(b - 10, 24).Value = "C2"
ElseIf Sheets("Replanteo").Cells(b - 1, 4).Value < 31.5 Or Sheets("Replanteo").Cells(b - 3, 4).Value < 31.5 Or Sheets("Replanteo").Cells(b - 5, 4).Value < 31.5 _
Or Sheets("Replanteo").Cells(b - 7, 4).Value < 31.5 Then
    Sheets("Replanteo").Cells(b, 16).Value = anc
    Sheets("Replanteo").Cells(b, 24).Value = "C2"
    Sheets("Replanteo").Cells(b - 2, 16).Value = semi
    Sheets("Replanteo").Cells(b - 4, 16).Value = eje
    Sheets("Replanteo").Cells(b - 6, 16).Value = eje
    Sheets("Replanteo").Cells(b - 8, 16).Value = semi
    Sheets("Replanteo").Cells(b - 10, 16).Value = anc
    Sheets("Replanteo").Cells(b - 10, 24).Value = "C2"
Else
    Sheets("Replanteo").Cells(b, 16).Value = anc
    Sheets("Replanteo").Cells(b, 24).Value = "C2"
    Sheets("Replanteo").Cells(b - 2, 16).Value = semi
    Sheets("Replanteo").Cells(b - 4, 16).Value = eje
    Sheets("Replanteo").Cells(b - 6, 16).Value = semi
    Sheets("Replanteo").Cells(b - 8, 16).Value = anc
    Sheets("Replanteo").Cells(b - 8, 24).Value = "C2"
End If
End Sub


'///
'///Rutina destinada a encontrar anclajes de puntos fijos que caen encima de puntos singulares
'///
Sub ancla(ByRef alfa, ByRef h As Integer, ByRef corte, b As Integer)

While Sheets("Replanteo").Cells(b, 33).Value > Sheets("Punto singular").Cells(alfa, 21).Value
    alfa = alfa + 1
Wend

If Sheets("Replanteo").Cells(b, 33).Value - Sheets("Punto singular").Cells(alfa - 1, 21).Value <= 8.5 And Not Sheets("Replanteo").Cells(b, 38).Value = "Tunel" And Not Sheets("Replanteo").Cells(b - 1, 25).Value = via Then
    h = h - 2
    corte = corte - Sheets("Replanteo").Cells(h - 1, 4).Value
End If
While Sheets("Replanteo").Cells(h, 33).Value > Sheets("Punto singular").Cells(alfa, 2).Value
    alfa = alfa + 1
Wend
If Sheets("Punto singular").Cells(alfa, 2).Value - Sheets("Replanteo").Cells(h, 33).Value <= 8.5 And Not Sheets("Replanteo").Cells(h, 38).Value = "Tunel" Then
    h = h - 2
    corte = corte - Sheets("Replanteo").Cells(h - 1, 4).Value
End If
End Sub

'//
'// Rutina destinada a dibujar flechas para aclarar los cantonamientos
'//
Sub pintar(ByRef z, zini, columna)

Top = Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, columna), Sheets("Replanteo").Cells(z - 1, columna)).Top
Lef = Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, columna), Sheets("Replanteo").Cells(z - 1, columna)).Left
Heigh = Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, columna), Sheets("Replanteo").Cells(z - 1, columna)).height
columnW = Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, columna), Sheets("Replanteo").Cells(z - 1, columna)).Width

If columna = 27 Then
    If Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 8, columna), Sheets("Replanteo").Cells(zini + 8, columna)).Interior.ColorIndex = 3 Or _
    Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 6, columna), Sheets("Replanteo").Cells(zini + 6, columna)).Interior.ColorIndex = 3 Or _
    Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 10, columna), Sheets("Replanteo").Cells(zini + 10, columna)).Interior.ColorIndex = 3 Then
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, columna), Sheets("Replanteo").Cells(z, columna)).Interior.ColorIndex = 4
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini, columna), Sheets("Replanteo").Cells(zini, columna)).Interior.ColorIndex = 4

        ActiveSheet.Shapes.AddLine(Lef + (columnW / 3), Top, Lef + (columnW / 3), Top + Heigh).Select
    Else
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, columna), Sheets("Replanteo").Cells(z, columna)).Interior.ColorIndex = 3
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini, columna), Sheets("Replanteo").Cells(zini, columna)).Interior.ColorIndex = 3
        Sheets("Replanteo").Shapes.AddLine(Lef + (2 * columnW / 3), Top, Lef + (2 * columnW / 3), Top + Heigh).Select
    End If
Else
    If Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini - 2, columna), Sheets("Replanteo").Cells(zini - 2, columna)).Interior.ColorIndex = 6 Then

        Sheets("Replanteo").Shapes.AddLine(Lef + (columnW / 3), Top, Lef + (columnW / 3), Top + Heigh).Select
    Else

        Sheets("Replanteo").Shapes.AddLine(Lef + (2 * columnW / 3), Top, Lef + (2 * columnW / 3), Top + Heigh).Select
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
