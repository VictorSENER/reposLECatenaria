Attribute VB_Name = "singular"
Sub marquise(ByRef h, ByRef a)
'///
'/// inicialización de variables
'///
vano_cal = Sheets("Punto singular").Cells(a, 4).Value
z = h
Sheets("Replanteo").Cells(h + 2, 33).Value = Sheets("Punto singular").Cells(a, 2).Value + 5
Call radio.radio1(h + 2)
While vano_cal <= Sheets("Replanteo").Cells(z - 1, 4).Value
    Sheets("Replanteo").Cells(z + 1, 4).Value = vano_cal
    pk0 = Sheets("Replanteo").Cells(z, 33).Value
    Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z + 2, 33).Value - Sheets("Replanteo").Cells(z + 1, 4).Value
    Call radio.radio1(z)
    vano_cal = vano_cal + dist_va_max
    z = z - 2
Wend
dist_restar = pk0 - Sheets("Replanteo").Cells(z + 2, 33).Value
'Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z, 33).Value + dist_restar
Call restar(dist_restar, z, h, a)
h = h + 2
While Sheets("Replanteo").Cells(h, 33).Value < Sheets("Punto singular").Cells(a, 21).Value
    Sheets("Replanteo").Cells(h + 1, 4).Value = 10
    Sheets("Replanteo").Cells(h + 2, 33).Value = Sheets("Replanteo").Cells(h, 33).Value + Sheets("Replanteo").Cells(h + 1, 4).Value
    Call radio.radio1(h + 2)
    Sheets("Replanteo").Cells(h, 38).Value = "Marquesina"
    h = h + 2
Wend
h = h - 4
a = a + 1
End Sub
Sub Viaducto(ByRef h, ByRef a)
Dim vano01 As Double, vano12 As Double, vano0 As Double
Dim z As Integer
Dim pkref As Double, pk0 As Double, pk1 As Double, pk2 As Double, dist_restar As Double
'///
'/// inicializacion de variables
'///
b = 3
vano1 = Sheets("Replanteo").Cells(h + 1, 4).Value
pk0 = Sheets("Replanteo").Cells(h, 33).Value
pk1 = Sheets("Punto singular").Cells(a, b).Value
If Not IsEmpty(Sheets("Punto singular").Cells(a, b + 1).Value) Then

    pk2 = Sheets("Punto singular").Cells(a, b + 1).Value
Else
    pk2 = 1000000000
End If
vano0 = Sheets("Replanteo").Cells(h - 1, 4).Value
dist_restar = pk0 - pk1
vano2 = pk2 - pk1
vano_actual = vano2
suma_restar = 0
'///
'/// comprobación si es necesario generar un espacio para poder ubicar el poste en la primera pila
'///
z = h
While Sheets("Replanteo").Cells(z + 1, 4).Value > (vano_actual + dist_va_max)
    vano_actual = vano_actual + dist_va_max
    suma_restar = suma_restar + Sheets("Replanteo").Cells(z + 1, 4).Value - vano_actual
    z = z - 2

Wend

'If dist_restar < (vano0 - dist_va_max) And suma_restar > dist_restar Then 'esta mal
    'vano_norm = (Int((dist_va_max + (pk2 - pk1)) / inc_norm_va) * inc_norm_va)
    'z = h
    'While dist_restar >= 0 And vano_norm < 54 And pk2 <> 0
        'dist_restada = Sheets("Replanteo").Cells(z - 1, 4).Value - vano_norm
        'dist_restar = dist_restar - dist_restada
        'vano_norm = vano_norm + dist_va_max
        'z = z - 2
    'Wend
'End If
'///
'/// si la comprobación da una distancia negativa quiere decir que es necesario generar un espacio
'///
If dist_restar < (vano0 - dist_va_max) And suma_restar > dist_restar Then
    h = h + 2
    Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h - 2, 33).Value + Sheets("Replanteo").Cells(h - 1, 4).Value
    Call radio.radio1(h)
    Sheets("Replanteo").Cells(h + 1, 4).Value = pk2 - pk1
    dist_restar = (Sheets("Replanteo").Cells(h, 33).Value - pk1)
'///
'/// tratamiento especial para viaductos con una sola pila
'///
ElseIf pk2 = 0 Then
    dist_restar = pk0 - pk1
Else
    dist_restar = pk0 - pk1
    Sheets("Replanteo").Cells(h + 1, 4).Value = pk2 - pk1
End If
z = h
'///
'/// llamada a restar
'///
Call restar(dist_restar, z, h, a)
'///
'/// ubicación de un poste en cada pila del viaducto
'///
While Not IsEmpty(Sheets("Punto singular").Cells(a, b))
    Sheets("Replanteo").Cells(h, 38).Value = "Viaducto"
    h = h + 2
    b = b + 1
        If Not IsEmpty(Sheets("Punto singular").Cells(a, b).Value) Then
            Sheets("Replanteo").Cells(h, 33).Value = Sheets("Punto singular").Cells(a, b).Value
            Call radio.radio1(h)
            Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Replanteo").Cells(h, 33).Value - Sheets("Replanteo").Cells(h - 2, 33).Value
        Else
            Sheets("Replanteo").Cells(h - 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(h - 2, 6).Value, h - 2)
        End If
    
Wend
h = h - 2
a = a + 1
End Sub

Sub paso_superior(ByRef h, ByRef a)
Dim L1 As Double, l2 As Double, vano0 As Double, dist_restar As Double, pmedio As Double
Dim pk1 As Double, pk2 As Double, div As Double, vano12 As Double, vanoref As Double
Dim z As Integer
'///
'/// inicializacion de variables
'///
L1 = Sheets("Punto singular").Cells(a, 2).Value - Sheets("Replanteo").Cells(h - 2, 33).Value
l2 = Sheets("Replanteo").Cells(h, 33).Value - Sheets("Punto singular").Cells(a, 21).Value
vano0 = Sheets("Replanteo").Cells(h + 1, 4).Value
pmedio = (vano0 - (Sheets("Punto singular").Cells(a, 21).Value - Sheets("Punto singular").Cells(a, 2).Value)) / 2
pk1 = Sheets("Punto singular").Cells(a, 2).Value - pmedio
pk2 = Sheets("Punto singular").Cells(a, 21).Value + pmedio
dist_restar = Sheets("Replanteo").Cells(h - 2, 33).Value - pk1
vano12 = pk2 - pk1
z = h
'///
'/// si la comprobación da una distancia negativa quiere decir que es necesario generar un espacio
'///
If dist_restar <= 0 Then
    h = h + 2
    dist_restar = Sheets("Replanteo").Cells(h - 3, 4).Value - Abs(dist_restar)
Else
    Sheets("Replanteo").Cells(h - 1, 4).Value = vano0
End If
'///
'/// ubicación del poste posterior
'///
Sheets("Replanteo").Cells(h, 33).Value = pk2
Call radio.radio1(h)
z = h - 2
'///
'/// llamada a restar
'///
Call restar(dist_restar, z, h, a)
If Sheets("Punto singular").Cells(a, 3).Value = "adelante" Then
    Sheets("Replanteo").Cells(h + 1, 4).Value = Sheets("Punto singular").Cells(a, 4).Value
    Sheets("Replanteo").Cells(h + 2, 33).Value = Sheets("Replanteo").Cells(h + 1, 4).Value + Sheets("Replanteo").Cells(h, 33).Value
    Call radio.radio1(h + 2)
    h = h + 2
End If
h = h - 2
a = a + 1
End Sub
Sub aguja(ByRef h, ByRef a, ByRef k)
Dim dist_restar As Double
Dim pk1 As Double, pk0 As Double, vano12 As Double, vanoref As Double
Dim z As Integer

'///
'/// inicializacion de variables
'///
pk1 = Sheets("Punto singular").Cells(a, 2).Value
pk0 = Sheets("Replanteo").Cells(h, 33).Value
dist_restar = pk0 - pk1
z = h
If Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) And Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
    dist_restar = dist_restar - (Sheets("Replanteo").Cells(h - 1, 4).Value - Sheets("Punto singular").Cells(a, 6).Value)
ElseIf Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) And Sheets("Punto singular").Cells(a, 22).Value = "OUT" _
And Sheets("Punto singular").Cells(a, 6).Value + dist_va_max < Sheets("Replanteo").Cells(h - 1, 4).Value Then
    dist_restar = dist_restar - (Sheets("Replanteo").Cells(h - 1, 4).Value - Sheets("Punto singular").Cells(a, 6).Value) + dist_va_max
End If


'///
'/// caso particular de tener un paso superior bajo antes de la aguja !!! se debe automatizar para todos los casos
'///
If Sheets("Punto singular").Cells(a - 2, 1).Value = "7 > P.S. > 5,2 m" And _
Sheets("Punto singular").Cells(a, 2).Value - Sheets("Punto singular").Cells(a - 2, 21).Value < (va_max) _
And Sheets("Punto singular").Cells(a, 2).Value - Sheets("Punto singular").Cells(a - 2, 21).Value > (inc_norm_va * 6) Then
    '///
    '/// hace falta añadir una celda
    '///
    If Sheets("Replanteo").Cells(h - 2, 33).Value < (Sheets("Punto singular").Cells(a - 2, 21).Value + dist_va_max) Then
        h = h + 2
        pk0 = Sheets("Replanteo").Cells(h - 4, 33).Value
        Sheets("Replanteo").Cells(h - 1, 4).Value = pk1 - (Sheets("Punto singular").Cells(a - 1, 21).Value + dist_va_max)
        Sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Punto singular").Cells(a - 1, 21).Value - Sheets("Punto singular").Cells(a - 1, 2).Value + (4 * inc_norm_va)
        pk_restar = pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value - Sheets("Replanteo").Cells(h - 3, 4).Value
        Sheets("Replanteo").Cells(h - 2, 33).Value = pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value
        Call radio.radio1(h - 2)
        Sheets("Replanteo").Cells(h, 33).Value = pk1
        Call radio.radio1(h)
        Sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(h, 6).Value, h)
        dist_restar = pk0 - pk_restar
        z = h - 4
    '///
    '/// no hace falta añadir una celda
    '///
    Else

        pk0 = Sheets("Replanteo").Cells(h - 4, 33).Value
        Sheets("Replanteo").Cells(h - 1, 4).Value = pk1 - (Sheets("Punto singular").Cells(a - 2, 21).Value + dist_va_max)
        Sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Punto singular").Cells(a - 2, 21).Value - Sheets("Punto singular").Cells(a - 2, 2).Value + 2 * dist_va_max
        quitar = pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value - Sheets("Replanteo").Cells(h - 3, 4).Value
        Call radio.radio1(h - 4)
        Sheets("Replanteo").Cells(h - 2, 33).Value = pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value
        Call radio.radio1(h - 2)
        Sheets("Replanteo").Cells(h, 33).Value = pk1
        Call radio.radio1(h)
        Sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(h, 6).Value, h)
        dist_restar = pk0 - quitar
        z = h - 4
        Sheets("Replanteo").Cells(h - 6, 16).Value = anc_aguj
        Sheets("Replanteo").Cells(h - 4, 16).Value = semi_eje_aguj
        Sheets("Replanteo").Cells(h - 2, 16).Value = semi_eje_aguj
        Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
        Call restar(dist_restar, z, h, a)
        GoTo salto
    End If
'///
'/// caso particular de tener un puente despues de la aguja !!! Automatizar para todos los casos
'///
ElseIf (Sheets("Punto singular").Cells(a + 2, 1).Value = "Puente" Or Sheets("Punto singular").Cells(a + 2, 1).Value = "7 > P.S. > 5,2 m") And _
Sheets("Punto singular").Cells(a + 2, 2).Value - Sheets("Punto singular").Cells(a, 21).Value < va_max And Sheets("Punto singular").Cells(a, 7).Value <> "Forzado" _
And Sheets("Punto singular").Cells(a, 22).Value > dist_va_max Then
    vano12 = Sheets("Punto singular").Cells(a + 2, 2).Value - Sheets("Punto singular").Cells(a, 2).Value - 2
    vanoref = vano12 + dist_va_max
    dist_restar = pk0 - pk1 - (Sheets("Replanteo").Cells(h - 1, 4).Value - vanoref)
    Sheets("Replanteo").Cells(h - 1, 4).Value = vanoref
    Sheets("Replanteo").Cells(h, 33).Value = pk1
    Call radio.radio1(h)
    Sheets("Replanteo").Cells(h + 1, 4).Value = vano12
    Sheets("Replanteo").Cells(h + 2, 33).Value = pk1 + vano12
    Call radio.radio1(h)
    z = h - 2
'///
'/// caso particular de tener un puente antes de la aguja
'///
'ElseIf Sheets("Punto singular").Cells(a - 1, 1).Value = "Puente" And _
'Sheets("Punto singular").Cells(a, 21).Value - Sheets("Punto singular").Cells(a - 1, 2).Value < 2 * va_max Then
    'z = h
    'dist_restar_comp = dist_restar
    'While dist_restar_comp > 0
        'If (sheets("Replanteo").Cells(z, 33).Value - dist_restar_comp) > Sheets("Punto singular").Cells(a - 1, 2).Value And (sheets("Replanteo").Cells(z, 33).Value - dist_restar_comp) < Sheets("Punto singular").Cells(a - 1, 21).Value Then
            
            
            'z = h
            'While Sheets("Punto singular").Cells(a - 1, 2).Value <= sheets("Replanteo").Cells(z, 33).Value

                'z = z - 2
            'Wend
            'dist_restar = dist_restar + sheets("Replanteo").Cells(z + 1, 4).Value
        'End If
        'dist_restar_comp = dist_restar_comp - inc_norm_va
        'z = z - 2
    'Wend
    'z = h
'ElseIf pk1 + va_max > Sheets("Trazado").Cells(k + 1, 3).Value And sheets("Replanteo").Cells(h - 1, 4).Value - vano.vano(Sheets("Trazado").Cells(k + 1, 2).Value) > dist_va_max _
'And Sheets("Trazado").Cells(k + 1, 4).Value - (pk1 + va_max) < dist_va_max * 4 And dist_restar > 0 Then
    'sheets("Replanteo").Cells(h - 1, 4).Value = vano.vano(Sheets("Trazado").Cells(k + 1, 2).Value) + dist_va_max
    'sheets("Replanteo").Cells(h, 33).Value = sheets("Replanteo").Cells(h - 2, 33).Value + sheets("Replanteo").Cells(h - 1, 4).Value
    'sheets("Replanteo").Cells(h, 33).Value = pk1
    'dist_restar = dist_restar + dist_va_max
    'z = h - 2
    'GoTo fin
    
ElseIf Not Not dist_restar < 0 And Sheets("Punto singular").Cells(a, 22).Value = "OUT" And Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) Then
    h = h + 2
    pk0 = Sheets("Replanteo").Cells(h - 2, 33).Value
    Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Punto singular").Cells(a, 6).Value + dist_va_max
    'sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Punto singular").Cells(a - 1, 21).Value - Sheets("Punto singular").Cells(a - 1, 2).Value + (4 * inc_norm_va)
    'pk_restar = pk1 - sheets("Replanteo").Cells(h - 1, 4).Value - sheets("Replanteo").Cells(h - 3, 4).Value
    'sheets("Replanteo").Cells(h - 2, 33).Value = pk1 - sheets("Replanteo").Cells(h - 1, 4).Value
    'Call radio.radio1(h - 2)
    Sheets("Replanteo").Cells(h, 33).Value = pk1
    Call radio.radio1(h)
    Sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(h, 6).Value, h)
    dist_restar = pk0 - (pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value)
    z = h - 2
ElseIf Not Not dist_restar > 0 And Sheets("Punto singular").Cells(a, 22).Value = "OUT" And Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) And _
Sheets("Punto singular").Cells(a, 6).Value + 2 * dist_va_max < Sheets("Replanteo").Cells(h - 1, 4).Value Then
    'h = h + 2
    pk0 = Sheets("Replanteo").Cells(h, 33).Value
    Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Punto singular").Cells(a, 6).Value + dist_va_max
    Sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Punto singular").Cells(a, 6).Value + 2 * dist_va_max
    'pk_restar = pk1 - sheets("Replanteo").Cells(h - 1, 4).Value - sheets("Replanteo").Cells(h - 3, 4).Value
    Sheets("Replanteo").Cells(h, 33).Value = pk1
    Call radio.radio1(h)
    Sheets("Replanteo").Cells(h - 2, 33).Value = pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value
    Call radio.radio1(h - 2)

    'sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(sheets("Replanteo").Cells(h, 6).Value)
    dist_restar = pk0 - (pk1 + dist_va_max)
    Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
    Sheets("Replanteo").Cells(h + 2, 16).Value = semi_eje_aguj
    Sheets("Replanteo").Cells(h + 4, 16).Value = semi_eje_aguj
    Sheets("Replanteo").Cells(h + 6, 16).Value = anc_aguj
    Sheets("Replanteo").Cells(h + 1, 35).Value = Sheets("Punto singular").Cells(a, 5).Value
    
    z = h - 4
    GoTo fin
ElseIf Not Not dist_restar > 0 And Sheets("Punto singular").Cells(a, 22).Value = "OUT" And Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) And _
Sheets("Punto singular").Cells(a, 6).Value + dist_va_max + inc_norm_va <= Sheets("Replanteo").Cells(h - 1, 4).Value Then
    'h = h + 2
    'pk0 = sheets("Replanteo").Cells(h, 33).Value
    dif = Sheets("Replanteo").Cells(h - 1, 4).Value - (Sheets("Punto singular").Cells(a, 6).Value + dist_va_max)
    Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Punto singular").Cells(a, 6).Value + dist_va_max
    'sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Punto singular").Cells(a, 6).Value + dist_va_max + inc_norm_va
    'pk_restar = pk1 - sheets("Replanteo").Cells(h - 1, 4).Value - sheets("Replanteo").Cells(h - 3, 4).Value
    Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h - 2, 33).Value + Sheets("Replanteo").Cells(h - 1, 4).Value
    'Call radio.radio1(h)
    'sheets("Replanteo").Cells(h - 2, 33).Value = pk1 - sheets("Replanteo").Cells(h - 1, 4).Value
    'Call radio.radio1(h - 2)

    'sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(sheets("Replanteo").Cells(h, 6).Value)
    'dist_restar = dist_restar + dif
    'If Sheets("Punto singular").Cells(a, 6).Value <= 30.5 Then
    
    
    'sheets("Replanteo").Cells(h, 16).Value = eje_aguj
    'sheets("Replanteo").Cells(h + 2, 16).Value = semi_eje_aguj
    'sheets("Replanteo").Cells(h + 4, 16).Value = semi_eje_aguj
    'sheets("Replanteo").Cells(h + 6, 16).Value = anc_aguj
    'sheets("Replanteo").Cells(h + 1, 35).Value = Sheets("Punto singular").Cells(a, 5).Value
    
    'z = h - 2
    'GoTo fin
    'End If

'///
'/// resto de los casos
'///
'Else
    'dist_restar = pk0 - pk1
    'z = h
    
'ElseIf dist_restar < 0 Then
   ' h = h + 2
    'Sheets("Replanteo").Cells(h, 33).Value = pk1
    'Call radio.radio1(h)
    'Sheets("Replanteo").Cells(h - 1, 4).Value = Sheets("Punto singular").Cells(a, 6).Value
    'sheets("Replanteo").Cells(h - 3, 4).Value = Sheets("Punto singular").Cells(a - 1, 21).Value - Sheets("Punto singular").Cells(a - 1, 2).Value + (4 * inc_norm_va)
    'pk_restar = pk1 - sheets("Replanteo").Cells(h - 1, 4).Value - sheets("Replanteo").Cells(h - 3, 4).Value
    'Sheets("Replanteo").Cells(h - 2, 33).Value = pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value
    'Call radio.radio1(h - 2)
    'Sheets("Replanteo").Cells(h - 3, 4).Value = vano.vano(Sheets("Replanteo").Cells(h - 2, 33).Value - Sheets("Replanteo").Cells(h - 3, 4).Value)

    'Sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(h, 6).Value)
    'dist_restar = pk0 - (pk1 - Sheets("Replanteo").Cells(h - 1, 4).Value)
    'z = h - 4

End If

'///
'/// llamar a restar
'///
fin:

Call restar(dist_restar, z, h, a)
'///
'/// escribir texto para las agujas
'///
    If Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
       ' If Sheets("Replanteo").Cells(h - 1, 4).Value <= 31.5 Then
            'Sheets("Replanteo").Cells(h - 6, 16).Value = anc_aguj
            'Sheets("Replanteo").Cells(h - 4, 16).Value = semi_eje_aguj
            'Sheets("Replanteo").Cells(h - 2, 16).Value = semi_eje_aguj
            'Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
        'Else
            'Sheets("Replanteo").Cells(h - 4, 16).Value = anc_aguj
salto:
            Sheets("Replanteo").Cells(h - 2, 16).Value = semi_eje_aguj
            Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
        'End If
        Sheets("Replanteo").Cells(h + 1, 35).Value = Sheets("Punto singular").Cells(a, 5).Value
        Sheets("Replanteo").Cells(h, 25).Value = aguj & " - " & Sheets("Punto singular").Cells(a, 4).Value & " - " & Sheets("Punto singular").Cells(a, 3).Value
        Sheets("Replanteo").Cells(h, 56).Value = Sheets("Punto singular").Cells(a, 3).Value
        z_var = h + 1
    Else
        'If Sheets("Replanteo").Cells(h + 1, 4).Value <= 31.5 Then
            'Sheets("Replanteo").Cells(h + 6, 16).Value = anc_aguj
            'Sheets("Replanteo").Cells(h + 4, 16).Value = semi_eje_aguj
            Sheets("Replanteo").Cells(h + 2, 16).Value = semi_eje_aguj
            Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
        'Else
                
            'Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
            'Sheets("Replanteo").Cells(h + 2, 16).Value = semi_eje_aguj
            'Sheets("Replanteo").Cells(h + 4, 16).Value = anc_aguj
        'End If
        Sheets("Replanteo").Cells(h + 1, 35).Value = Sheets("Punto singular").Cells(a, 5).Value
        Sheets("Replanteo").Cells(h, 25).Value = aguj & " - " & Sheets("Punto singular").Cells(a, 4).Value & " - " & Sheets("Punto singular").Cells(a, 3).Value
        Sheets("Replanteo").Cells(h, 56).Value = Sheets("Punto singular").Cells(a, 3).Value
        z_var = h
    End If


If Sheets("Punto singular").Cells(a, 7).Value = "Forzado" Then
    estacion = Sheets("Punto singular").Cells(a, 3).Value
    cont = 2
    While Not IsEmpty(Sheets(estacion).Cells(cont, 1).Value)
        Sheets("Replanteo").Cells(h + 1, 4).Value = Sheets(estacion).Cells(cont, 1).Value
        Sheets("Replanteo").Cells(h + 2, 33).Value = Sheets("Replanteo").Cells(h + 1, 4) + Sheets("Replanteo").Cells(h, 33)
        Call radio.radio1(h)
        cont = cont + 1
        h = h + 2
    Wend
    a = a + 1
    While Sheets("Punto singular").Cells(a, 1).Value <> "Aguja"
        a = a + 1
    Wend
    'If Sheets("Replanteo").Cells(h - 1, 4).Value <= 31.5 Then
        'Sheets("Replanteo").Cells(h + 6, 16).Value = anc_aguj
        'Sheets("Replanteo").Cells(h + 4, 16).Value = semi_eje_aguj
        Sheets("Replanteo").Cells(h + 2, 16).Value = semi_eje_aguj
        Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
    'Else
    
        'Sheets("Replanteo").Cells(h, 16).Value = eje_aguj
        'Sheets("Replanteo").Cells(h + 2, 16).Value = semi_eje_aguj
        'Sheets("Replanteo").Cells(h + 4, 16).Value = anc_aguj
    'End If
    Sheets("Replanteo").Cells(h + 1, 35).Value = Sheets("Punto singular").Cells(a, 5).Value
    'Sheets("Replanteo").Cells(h + 1, 25).Value = Sheets("Punto singular").Cells(a, 4).Value
    Sheets("Replanteo").Cells(h, 25).Value = aguj & " - " & Sheets("Punto singular").Cells(a, 4).Value & " - " & Sheets("Punto singular").Cells(a, 3).Value
    Sheets("Replanteo").Cells(h, 56).Value = Sheets("Punto singular").Cells(a, 3).Value
    If Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) Then
        GoTo jump
    End If
    h = h - 2
    'sheets("Replanteo").Cells(h, 33).Value = Sheets("Punto singular").Cells(a, 2).Value
Else
    Sheets("Replanteo").Cells(h + 2, 33).Value = Sheets("Replanteo").Cells(h + 1, 4) + Sheets("Replanteo").Cells(h, 33)
    Call radio.radio1(h + 2)
End If

If Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) And Sheets("Punto singular").Cells(a, 22).Value = "OUT" Then
jump:
    Sheets("Replanteo").Cells(h + 1, 4).Value = Sheets("Punto singular").Cells(a, 6).Value
    Sheets("Replanteo").Cells(h + 2, 33).Value = Sheets("Replanteo").Cells(h + 1, 4) + Sheets("Replanteo").Cells(h, 33)
    Call radio.radio1(h + 2)
    h = h + 2
    Sheets("Replanteo").Cells(h + 1, 4).Value = vano.vano(Sheets("Replanteo").Cells(h, 6).Value, h)
End If

a = a + 1
End Sub
Sub Zona(ByRef h, ByRef a)
'///
'/// se deberá implementar cuando se realiza una catenaria para 25 kVA
'///
End Sub

Private Function two(ByRef z, ByRef dist_restar, n, ByRef h, ByRef div, a)
Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z, 33).Value - dist_restar ' + div
algo = Sheets("Replanteo").Cells(z, 33).Value
If dist_restar > inc_norm_va Then
    Call punto_singular.sing1(z, a - 1, 1, dist_restar)
End If
Call radio.radio1(z)
'If algo <> sheets("Replanteo").Cells(z, 33).Value Then
   ' dist_restar2 = sheets("Replanteo").Cells(z, 33).Value - algo
    'sheets("Replanteo").Cells(z, 33).Value = sheets("Replanteo").Cells(z, 33).Value + dist_restar2
    'z = z - 2
'End If

vano_nuevo = vano.vano(Sheets("Replanteo").Cells(z, 6).Value, h)
If vano_nuevo < Sheets("Replanteo").Cells(z + 1, 4).Value Then
    If z <> h Then
        dist_restar = dist_restar - n - (Sheets("Replanteo").Cells(z + 1, 4).Value - vano_nuevo)
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        Sheets("Replanteo").Cells(z + 1, 4).Value = vano_nuevo
        Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z + 2, 33).Value - Sheets("Replanteo").Cells(z + 1, 4).Value
        Call radio.radio1(z)
        If Sheets("Replanteo").Cells(z + 3, 4).Value - vano_nuevo > dist_va_max And h - z > 2 Then
            dist_restar = dist_restar - n
            Sheets("Replanteo").Cells(z + 3, 4).Value = Sheets("Replanteo").Cells(z + 3, 4).Value - n
            Sheets("Replanteo").Cells(z + 2, 33).Value = Sheets("Replanteo").Cells(z + 4, 33).Value - Sheets("Replanteo").Cells(z + 3, 4).Value
            Call radio.radio(z + 2)
            Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z + 2, 33).Value - Sheets("Replanteo").Cells(z + 1, 4).Value
            Call radio.radio(z)
            ex = 0
        End If
    Else
        dist_restar = dist_restar - n
        Sheets("Replanteo").Cells(z + 1, 4).Value = vano_nuevo
    End If
Else 'If algo = sheets("Replanteo").Cells(z, 33).Value Then
    dist_restar = dist_restar - n
End If
End Function
Sub restar(dist_restar, z, h, a)

'div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
rest = 0
    If Not IsEmpty(Sheets("Punto singular").Cells(a, 6).Value) And Sheets("Punto singular").Cells(a, 22).Value = "IN" Then
        Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z, 33).Value - (Sheets("Replanteo").Cells(z - 1, 4).Value - Sheets("Punto singular").Cells(a, 6).Value) - dist_restar
        Call radio.radio1(z)
        Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Punto singular").Cells(a, 6).Value
        z = z - 2
    End If
    While dist_restar > 0
        div = dist_restar - (Int(dist_restar / inc_norm_va) * inc_norm_va)
        quitar = 2 * dist_va_max + inc_norm_va
        quitar2 = 2 * dist_va_max
        If IsEmpty(Sheets("Replanteo").Cells(z + 2, 33).Value) And dist_restar >= quitar + inc_norm_va And Sheets("Replanteo").Cells(z - 1, 4).Value - (Sheets("Replanteo").Cells(z - 3, 4).Value) = dist_va_max _
        And Sheets("Replanteo").Cells(z - 1, 4).Value >= 54 Then
            quitar = 2 * dist_va_max + inc_norm_va
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - quitar
            Call two(z, dist_restar, quitar, h, div, a)
        ElseIf IsEmpty(Sheets("Replanteo").Cells(z + 2, 33).Value) And dist_restar >= quitar2 And Sheets("Replanteo").Cells(z - 1, 4).Value - (Sheets("Replanteo").Cells(z - 3, 4).Value) = dist_va_max _
        And Sheets("Replanteo").Cells(z - 1, 4).Value >= 54 Then
            quitar2 = 2 * dist_va_max
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - quitar2
            Call two(z, dist_restar, quitar2, h, div, a)
        
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value - (Sheets("Replanteo").Cells(z + 1, 4).Value) > dist_va_max And _
            z >= h - 4 And dist_restar > dist_va_max Then
            rest = 0
            While Sheets("Replanteo").Cells(z - 1, 4).Value - (Sheets("Replanteo").Cells(z + 1, 4).Value) > dist_va_max And _
            dist_restar > inc_norm_va
                Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - inc_norm_va
                rest = rest + inc_norm_va
            Wend
            Call two(z, dist_restar, rest, h, div, a)
        ElseIf Abs(Sheets("Replanteo").Cells(z - 1, 4).Value - (Sheets("Replanteo").Cells(z + 1, 4).Value)) >= dist_va_max And _
            z >= h And dist_restar > inc_norm_va Then
                Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - inc_norm_va
                Call two(z, dist_restar, inc_norm_va, h, div, a)
        'ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value >= (Sheets("Replanteo").Cells(z - 3, 4).Value) And dist_restar >= dist_va_max _
        'And Sheets("Replanteo").Cells(z - 1, 4).Value >= Sheets("Replanteo").Cells(z + 1, 4).Value And Sheets("Replanteo").Cells(z - 1, 4).Value > 31.5 Then
            'Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - dist_va_max
            'Call two(z, dist_restar, dist_va_max, h, div, a) '///
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value >= (Sheets("Replanteo").Cells(z - 3, 4).Value) And dist_restar >= dist_va_max _
        And Sheets("Replanteo").Cells(z - 1, 4).Value >= Sheets("Replanteo").Cells(z + 1, 4).Value And Sheets("Replanteo").Cells(z - 1, 4).Value >= (va_max - 2 * inc_norm_va) Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - dist_va_max
            Call two(z, dist_restar, dist_va_max, h, div, a) '///
        ElseIf dist_restar < inc_norm_va And Sheets("Replanteo").Cells(z + 1, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value < dist_va_max And Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value < div Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - div
            Call two(z, dist_restar, div, h, div, a)
        ElseIf dist_restar < dist_va_max And Sheets("Replanteo").Cells(z + 1, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value < dist_va_max And Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value < div _
        And Sheets("Replanteo").Cells(z - 1, 4).Value > 27 Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - dist_restar
            Call two(z, dist_restar, dist_restar, h, div, a)
        'ElseIf dist_restar > dist_va_max And Sheets("Replanteo").Cells(z + 1, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value <= dist_va_max And Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value <= dist_va_max _
        'And (Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value <= dist_va_max Or dist_restar > dist_va_max * 1.5) Then
            'Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - dist_va_max
            'Call two(z, dist_restar, inc_norm_va, h, div, a)
        ElseIf dist_restar > inc_norm_va And Sheets("Replanteo").Cells(z + 1, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value <= dist_va_max And Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value <= dist_va_max _
        And (Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value <= inc_norm_va Or dist_restar > dist_va_max * 1.5) Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - inc_norm_va
            Call two(z, dist_restar, inc_norm_va, h, div, a)
        ElseIf dist_restar = inc_norm_va And Sheets("Replanteo").Cells(z + 1, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value < dist_va_max And Sheets("Replanteo").Cells(z - 3, 4).Value - (Sheets("Replanteo").Cells(z - 1, 4).Value - inc_norm_va) <= dist_va_max Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - inc_norm_va
            Call two(z, dist_restar, inc_norm_va, h, div, a)
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value > Sheets("Replanteo").Cells(z + 1, 4).Value And inc_norm_va < dist_restar And _
        (dist_restar - inc_norm_va) > (Sheets("Replanteo").Cells(z - 5, 4).Value - Sheets("Replanteo").Cells(z - 3, 4).Value) And _
        (Sheets("Replanteo").Cells(z - 3, 4).Value - (Sheets("Replanteo").Cells(z - 1, 4).Value + inc_norm_va) > dist_va_max Or dist_restar > dist_va_max) Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - inc_norm_va
            Call two(z, dist_restar, inc_norm_va, h, div, a)
        ElseIf dist_restar < inc_norm_va And Sheets("Replanteo").Cells(z - 3, 4).Value - Sheets("Replanteo").Cells(z - 1, 4).Value < div And _
         IsEmpty(Sheets("Replanteo").Cells(z + 1, 33).Value) Then
            Sheets("Replanteo").Cells(z - 1, 4).Value = Sheets("Replanteo").Cells(z - 1, 4).Value - div
            Call two(z, dist_restar, div, h, div, a)
        Else
            Call two(z, dist_restar, 0, h, div, a)
        End If
        z = z - 2
        
    Wend
    Sheets("Replanteo").Cells(z, 33).Value = Sheets("Replanteo").Cells(z - 1, 4) + Sheets("Replanteo").Cells(z - 2, 33)
End Sub

