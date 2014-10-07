Attribute VB_Name = "punto_singular"
Sub sing1(ByRef h, ByRef a, marcador, dist_restar)
Dim z As Integer
Dim dist_seg As Integer
'///
'/// inicialización de variables
'///
dist_seg = 2.5
dist_pn = 7
dist_in_tun = 18
z = h
rest = 0
inicio:
    '///
    '/// comprobación ubicación PK actual cae encima de puente, drenaje, paso inferior, paso a nivel y conducto
    '/// se calcula la distancia a restar y se llama al módulo singular.restar
    '///
    If (Sheets("Replanteo").Cells(h, 33).Value >= Abs((Sheets("Punto singular").Cells(a, 2).Value - dist_pn)) And Sheets("Replanteo").Cells(h, 33).Value <= (Sheets("Punto singular").Cells(a, 21).Value + dist_pn) _
    And (Sheets("Punto singular").Cells(a, 1).Value = "P.N.")) Then
            dist_restar1 = Sheets("Replanteo").Cells(h, 33).Value - (Sheets("Punto singular").Cells(a, 2).Value - dist_pn)
            Call singular.restar(dist_restar1, z, h, a)
    
    ElseIf (Sheets("Replanteo").Cells(h, 33).Value >= Abs((Sheets("Punto singular").Cells(a, 2).Value - dist_seg)) And Sheets("Replanteo").Cells(h, 33).Value <= (Sheets("Punto singular").Cells(a, 21).Value + dist_seg) _
    And (Sheets("Punto singular").Cells(a, 1).Value = "Conducto" Or Sheets("Punto singular").Cells(a, 1).Value = "P.I." Or Sheets("Punto singular").Cells(a, 1).Value = "Drenaje" Or Sheets("Punto singular").Cells(a, 1).Value = "P.N." _
    Or Sheets("Punto singular").Cells(a, 1).Value = "Puente" Or Sheets("Punto singular").Cells(a, 1).Value = "P.S. > 7 m")) Then
        '///
        '/// caso puente y aguja continuados
        '///

        'If marcador = 0 And (Sheets("Punto singular").Cells(a, 1).Value = "Puente" And Sheets("Punto singular").Cells(a + 1, 1).Value = "Aguja") Then
            'algo = 0
        '///
        '/// resto de casos
        '///
        If marcador = 0 And (Sheets("Punto singular").Cells(a, 3).Value <> "saltar") Then
            dist_restar1 = Sheets("Replanteo").Cells(h, 33).Value - (Sheets("Punto singular").Cells(a, 2).Value - dist_seg)
            Call singular.restar(dist_restar1, z, h, a)
        'ElseIf Sheets("Punto singular").Cells(a, 3).Value = "adelante" Then
            'dist_restar1 = sheets("Replanteo").Cells(h, 33).Value - (Sheets("Punto singular").Cells(a, 2).Value - dist_seg)
            'Call singular.restar(dist_restar1, z, h, a)
            'algo = 0
            
        '///
        '/// restando nos encontramos con un punto singular
        '///
        ElseIf marcador = 1 Then
            dist_restar1 = Sheets("Replanteo").Cells(h, 33).Value - (Sheets("Punto singular").Cells(a, 2).Value - dist_seg)
            dist_restar2 = (Sheets("Punto singular").Cells(a, 21).Value + dist_seg) - Sheets("Replanteo").Cells(h, 33).Value
            Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h, 33).Value - dist_restar2
            Call radio.radio1(h)

            dist_restar = dist_restar + dist_restar2
            div1 = dist_restar2 - (Int(dist_restar2 / inc_norm_va) * inc_norm_va)
            rest = 0
            z = z + 2
            'longitud = sheets("Replanteo").Cells(z - 2, 33).Value + sheets("Replanteo").Cells(z - 1, 4).Value + div1
            
            While dist_restar2 > 0
                vano_algo = vano.vano(Sheets("Replanteo").Cells(z - 2, 6).Value, h - 2)
                If dist_restar2 < inc_norm_va And vano_algo > Sheets("Replanteo").Cells(z - 1, 4).Value + div1 Then
                    Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value + div1
                    Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z - 2, 33).Value + Sheets("Replanteo").Cells(z - 1, 4).Value
                    Call radio.radio1(z)
                    dist_restar2 = dist_restar2 - div1
                
                ElseIf vano_algo >= Sheets("Replanteo").Cells(z - 1, 4).Value + inc_norm_va Then
                    Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value + inc_norm_va
    
                    Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z - 2, 33).Value + Sheets("Replanteo").Cells(z - 1, 4).Value
                    Call radio.radio1(z)
                    z = z + 2
                    dist_restar2 = dist_restar2 - inc_norm_va
                Else
                    Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z - 2, 33).Value + Sheets("Replanteo").Cells(z - 1, 4).Value
                    Call radio.radio1(z)
                    z = z + 2
                End If

            Wend

        End If
    '///
    
    ElseIf Sheets("Replanteo").Cells(h, 33).Value >= Abs((Sheets("Punto singular").Cells(a, 2).Value - dist_in_tun)) And Sheets("Punto singular").Cells(a, 2).Value - Sheets("Replanteo").Cells(h, 33).Value > dist_max_va And Sheets("Punto singular").Cells(a, 1).Value = "Tunel" Then
    dist_restar1 = Sheets("Replanteo").Cells(h, 33).Value - (Sheets("Punto singular").Cells(a, 2).Value - dist_in_tun)
    Call singular.restar(dist_restar1, z, h, a)

    '/// comprobación ubicación PK anterior cae encima de puente, drenaje, paso inferior, paso a nivel y conducto
    '/// se calcula la distancia a restar y se llama al módulo singular.restar si se viene del modulo restar
    '///
    ElseIf (Abs((Sheets("Replanteo").Cells(h, 33).Value - Sheets("Punto singular").Cells(a - 1, 21).Value)) <= dist_seg And (Sheets("Punto singular").Cells(a - 1, 1).Value = "Conducto" Or _
    Sheets("Punto singular").Cells(a - 1, 1).Value = "P.I." Or Sheets("Punto singular").Cells(a - 1, 1).Value = "Drenaje" Or Sheets("Punto singular").Cells(a - 1, 1).Value = "P.N." _
    Or Sheets("Punto singular").Cells(a - 1, 1).Value = "Puente")) Then
        If marcador = 0 And Sheets("Punto singular").Cells(a - 1, 3).Value <> "saltar" Then
            dist_restar = Sheets("Replanteo").Cells(h, 33).Value - (Sheets("Punto singular").Cells(a - 1, 2).Value - dist_seg)
            Call singular.restar(dist_restar, z, h, a - 1)

        End If
    End If
'///
'/// incrementar la fila del punto singular si es necesario
'///
While Sheets("Replanteo").Cells(h, 33).Value > Sheets("Punto singular").Cells(a, 21).Value And Sheets("Punto singular").Cells(a, 23).Value <> "FINAL"
    a = a + 1
Wend
End Sub

Sub sing(ByRef h, ByRef a, ByRef k)
dist_in_tun = 18
'///
'/// incrementar la fila del punto singular si es necesario
'///
While (Sheets("Replanteo").Cells(h, 33).Value >= Sheets("Punto singular").Cells(a, 21).Value And Sheets("Punto singular").Cells(a, 23).Value <> "FINAL" _
And Sheets("Punto singular").Cells(a + 1, 21).Value - Sheets("Punto singular").Cells(a, 21).Value <= va_max And Sheets("Punto singular").Cells(a, 1).Value <> "Aguja") _
 Or Sheets("Punto singular").Cells(a, 1).Value = "Señalización"
    a = a + 1
Wend
Sheets("Replanteo").Cells(5, 1).Value = Sheets("Punto singular").Cells(a, 2).Value
'If 42700 < Sheets("Replanteo").Cells(h, 33).Value Then
   ' algo = 0
'End If
'///
'/// comprobación ubicación PK actual cae encima de puente largo o un paso superior bajo
'/// se llama al módulo singular.paso_superior
'///
If (Sheets("Replanteo").Cells(h - 2, 33).Value < Sheets("Punto singular").Cells(a, 21).Value And _
Sheets("Replanteo").Cells(h, 33).Value > Sheets("Punto singular").Cells(a, 2).Value And _
(Sheets("Punto singular").Cells(a, 1) = "7 > P.S. > 5,2 m" Or Sheets("Punto singular").Cells(a, 1) = "PuenteXL")) Then
    '///
    '/// caso particular de una aguja muy cerca de un paso superior bajo, se realiza conjuntamente con la aguja
    '///
    If Sheets("Punto singular").Cells(a, 1) = "7 > P.S. > 5,2 m" And Sheets("Punto singular").Cells(a + 2, 1) = "Aguja" And Sheets("Punto singular").Cells(a + 2, 2) - Sheets("Punto singular").Cells(a, 2) < va_max Then
        a = a + 1
    '///
    '/// resto de casos
    '///
    ElseIf Sheets("Punto singular").Cells(a, 1) = "7 > P.S. > 5,2 m" And Sheets("Punto singular").Cells(a - 1, 1) = "Aguja" And Sheets("Punto singular").Cells(a, 2) - Sheets("Punto singular").Cells(a - 1, 2) < va_max Then
        a = a + 1
    'ElseIf Sheets("Punto singular").Cells(a, 1) = "PuenteXL" And Sheets("Punto singular").Cells(a, 3) = "saltar" Then
        'algo = 0
    Else
            Call singular.paso_superior(h, a)
    End If
'///
'/// comprobación ubicación PK actual cae encima de un viaducto
'/// se llama al módulo singular.viaducto
'///
ElseIf Sheets("Replanteo").Cells(h, 33).Value >= Sheets("Punto singular").Cells(a, 3).Value _
And Sheets("Punto singular").Cells(a, 1) = "Viaducto" Then
    Call singular.Viaducto(h, a)

ElseIf Sheets("Replanteo").Cells(h, 33).Value >= Sheets("Punto singular").Cells(a + 1, 3).Value _
And Sheets("Punto singular").Cells(a + 1, 1) = "Viaducto" And Sheets("Punto singular").Cells(a, 1) = "Tunel" And Sheets("Punto singular").Cells(a + 2, 1) = "Tunel" Then
    Call singular.Viaducto(h, a + 1)

'///
'/// comprobación ubicación PK actual cae encima de un tunel
'/// se calcula la distancia a restar y se llama al módulo singular.restar
'///Or Sheets("Punto singular").Cells(a, 1) = "Marquesina")
ElseIf (Sheets("Replanteo").Cells(h, 33).Value >= Sheets("Punto singular").Cells(a, 2).Value And Sheets("Replanteo").Cells(h, 33).Value <= Sheets("Punto singular").Cells(a, 21).Value _
And Sheets("Punto singular").Cells(a, 1) = "Tunel") Then
        Sheets("Replanteo").Cells(h, 38).Value = Sheets("Punto singular").Cells(a, 1).Value
    '///
    '/// se comprueba si el vano calculado anteriormente es mayor al permitido dentro de túnel
    '///
    If Sheets("Replanteo").Cells(h - 1, 4).Value > va_max_tunel Then
        Sheets("Replanteo").Cells(h - 1, 4).Value = va_max_tunel
        Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h - 2, 33) + Sheets("Replanteo").Cells(h - 1, 4)
        Call radio.radio1(h)
    End If
    If Sheets("Replanteo").Cells(h, 33).Value + va_max_tunel >= Sheets("Punto singular").Cells(a, 21).Value And Sheets("Replanteo").Cells(h, 33).Value + va_max_tunel + 9 <= Abs((Sheets("Punto singular").Cells(a, 21).Value + dist_in_tun)) Then
        long_tun = Sheets("Punto singular").Cells(a, 4).Value
        
        dist_restar = (Sheets("Replanteo").Cells(h, 33).Value + va_max_tunel) - (Sheets("Punto singular").Cells(a, 21).Value - long_tun)
        'Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Replanteo").Cells(h - 1, 4).Value - dist_restar
        Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h, 33).Value - dist_restar
        Call radio.radio1(h)
        'dist_restar = (Sheets("Replanteo").Cells(h, 33).Value + va_max_tunel) - (Sheets("Punto singular").Cells(a, 21).Value - 5)
        Call restar(dist_restar, h - 2, h, a)
    End If
'///
'/// comprobación ubicación PK actual cae encima de una aguja
'/// se llama al módulo singular.aguja
'///
ElseIf (Sheets("Replanteo").Cells(h - 2, 33).Value < Sheets("Punto singular").Cells(a, 2).Value And Sheets("Replanteo").Cells(h, 33).Value > Sheets("Punto singular").Cells(a, 2).Value And _
    (Sheets("Punto singular").Cells(a, 1) = "Aguja" Or Sheets("Punto singular").Cells(a, 1) = "Desvío")) Then
    Call singular.aguja(h, a, k)
'///
'/// comprobación que solamente la ubicación PK actual cae encima de una marquesina
'/// se llama al módulo singular.marquesina
'///
ElseIf Sheets("Replanteo").Cells(h, 33).Value >= Sheets("Punto singular").Cells(a, 2).Value And Sheets("Punto singular").Cells(a, 1) = "Marquesina" Then
    Call singular.marquise(h, a)
    'punto_medio = Sheets("Punto singular").Cells(a, 2).Value + ((Sheets("Punto singular").Cells(a, 21).Value - Sheets("Punto singular").Cells(a, 2).Value) / 2)
    'dist_restar1 = Sheets("Replanteo").Cells(h, 33).Value - punto_medio
    'Call singular.restar(dist_restar1, h, h, a)
    'Call singular.aguja(h, a, k)


End If
'///
'/// incrementar la fila del punto singular si es necesario
'///
While Sheets("Replanteo").Cells(h, 33).Value >= Sheets("Punto singular").Cells(a, 21).Value + 7 And Sheets("Punto singular").Cells(a, 23).Value <> "FINAL"
    a = a + 1
Wend
End Sub

