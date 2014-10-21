Attribute VB_Name = "tabla_vanos"
Sub tabla_vanos(nombre_catVB, poli, ventoso)

'//
'//AÑADIR EN UN IF NUEVAS TABLAS DE VANOS
'//

'//
'//CATENARIA MARRUECOS 1x25 KV (CALCULADA)
'//

If nombre_catVB = "Marruecos" Then
'//
'//LECTURA BASE DE DATOS
'//
nombre_catVB = "Marruecos"
Call cargar.datos_lac(nombre_catVB)
Call cargar.datos_cable
'//
'//CÁLCULO FUERZA VIENTO
'//

If n_hc = 2 Then
    coef_viento = 1.2 * 1.6         '/// tabla 5.7 CLER
Else
    coef_viento = 1.2               '/// tabla 5.7 CLER
End If

'//
'//Para el cálculo de la fuerza de viento se ha considerado un 15% de afectación en las péndolas
'//

fuerza_viento = 0.625 * coef_viento * (vw ^ 2) * (diam_hc + diam_sust + diam_pend * 0.15) '/// formula 5.27 de CLER ¿de donde sale el 0,15?

'//
'//ZONA RECTA
'//

    d_max_cu_var = d_max_cu
    d_media = (d_max_cu + d_max_cu_var) / 2
    t_tot = (n_hc * t_hc + t_sust) * 9.81 * 0.89  '/// que significa el 0.89?
    
    va_max_re = 2 * Sqr((t_tot / fuerza_viento) * (d_max_ad + Sqr((d_max_ad ^ 2) - (d_max_re ^ 2)))) '/// fórmula 5.87 CLER
    
    va_max_re = Int(va_max_re)
    
    va_max_var = va_max
    
    While va_max_re < va_max_var
        While (va_max_re / inc_norm_va) <> Int(va_max_re / inc_norm_va)
            va_max_re = va_max_re - 0.5
        Wend
    
        va_max_var = va_max_re

    Wend
    '///
    '/// realizado Edgar
    '///
    i = 3
    va_max_cu = va_max_var
    r = r_re
    Sheets("Vano").Cells(i, 2) = 100000
    Sheets("Vano").Cells(i, 1) = va_max_var
    Sheets("Vano").Cells(i, 3) = r_re
    Sheets("Vano").Cells(i, 4) = d_max_re
    Sheets("Vano").Cells(i, 5) = -d_max_cu
    Sheets("Vano").Cells(i + 1, 2) = r_re
    i = i + 1


     d_max_cu_var = 2 * ((va_max_var ^ 2) / (8 * r)) - d_max_cu
     d_max_cu_fix = -d_max_cu
    While d_max_cu_var <= d_max_cu

        d_max_cu_var = 2 * ((va_max_var ^ 2) / (8 * r)) - d_max_cu
        While d_max_cu_var <= d_max_cu_fix + 0.05
            r = r - 10
            d_max_cu_var = 2 * ((va_max_var ^ 2) / (8 * r)) - d_max_cu
        Wend

        If d_max_cu_var > d_max_fic + 0.05 Then
            r = r + 10
            d_max_cu_var = 2 * ((va_max_var ^ 2) / (8 * r)) - d_max_cu
        End If
        If d_max_cu_var > d_max_cu Then
            va_max_var = va_max_var - inc_norm_va
            GoTo inicio
        End If
        Sheets("Vano").Cells(i, 3) = r
        Sheets("Vano").Cells(i + 1, 2) = r
        Sheets("Vano").Cells(i, 4) = d_max_cu
        Sheets("Vano").Cells(i, 5) = d_max_cu_var
        Sheets("Vano").Cells(i, 1) = va_max_var
            
        d_max_cu_fix = d_max_cu_fix + 0.05
        i = i + 1
    Wend
    va_max_var = va_max_var - inc_norm_va
    
'///
'//
'//ZONA CURVA
'//

inicio:

    While va_max_cu >= va_max_var And r >= r_min_traz
        r = r - 10
        va_max_cu = Sqr((8 * r * t_tot * (d_max_ad + d_max_cu)) / (fuerza_viento * r + t_tot)) '/// formula 5.88 CLER
    Wend
        
    Sheets("Vano").Cells(i + 1, 2) = r
    Sheets("Vano").Cells(i, 4) = d_max_cu
    Sheets("Vano").Cells(i, 5) = d_max_cu
    
    If r < r_min_traz Then
        va_max_cu = Sqr((8 * r_min_traz * t_tot * (d_max_ad + d_max_cu)) / (fuerza_viento * r + t_tot)) '/// formula 5.88 CLER
        If va_max_cu >= va_max_var Then
            Sheets("Vano").Cells(i, 3) = r_min_traz
            Sheets("Vano").Cells(i, 1) = va_max_var
            GoTo fin
        Else
            Sheets("Vano").Cells(i, 3) = r + 50
            Sheets("Vano").Cells(i, 1) = va_max_var
            GoTo fin
        End If
    GoTo fin
    End If
    
    Sheets("Vano").Cells(i, 3) = r
    Sheets("Vano").Cells(i, 1) = va_max_var
    i = i + 1
    va_max_cu = va_max_var - inc_norm_va
    
        While (va_max_cu / inc_norm_va) <> Int(va_max_cu / inc_norm_va)
            va_max_cu = va_max_cu + 0.5
        Wend
    
    va_max_var = va_max_cu
    
    GoTo inicio
    
fin:

'//
'//FINALIZACIÓN TABLA
'//

Sheets("Vano").Range(Sheets("Vano").Cells(i + 1, 1), Sheets("Vano").Cells(i + 10, 5)).ClearContents


'//
'//CATENARIA MARRUECOS 3.000 Vcc (FIJADA POR MEMORANDUM)
'//

ElseIf nombre_catVB = "Marruecos 3.000 Vcc" And ventoso(poli) = "si" Then



Sheets("Vano").Cells(3, 1).Value = 54
Sheets("Vano").Cells(3, 2).Value = 100000
Sheets("Vano").Cells(3, 3).Value = 7500
Sheets("Vano").Cells(3, 4).Value = 0.2
Sheets("Vano").Cells(3, 5).Value = -0.2

Sheets("Vano").Cells(4, 1).Value = 54
Sheets("Vano").Cells(4, 2).Value = 7500
Sheets("Vano").Cells(4, 3).Value = 5000
Sheets("Vano").Cells(4, 4).Value = 0.2
Sheets("Vano").Cells(4, 5).Value = -0.04

Sheets("Vano").Cells(5, 1).Value = 54
Sheets("Vano").Cells(5, 2).Value = 5000
Sheets("Vano").Cells(5, 3).Value = 4000
Sheets("Vano").Cells(5, 4).Value = 0.2
Sheets("Vano").Cells(5, 5).Value = 0.04

Sheets("Vano").Cells(6, 1).Value = 54
Sheets("Vano").Cells(6, 2).Value = 4000
Sheets("Vano").Cells(6, 3).Value = 3000
Sheets("Vano").Cells(6, 4).Value = 0.2
Sheets("Vano").Cells(6, 5).Value = 0.12

Sheets("Vano").Cells(7, 1).Value = 54
Sheets("Vano").Cells(7, 2).Value = 3000
Sheets("Vano").Cells(7, 3).Value = 1350
Sheets("Vano").Cells(7, 4).Value = 0.2
Sheets("Vano").Cells(7, 5).Value = 0.2

Sheets("Vano").Cells(8, 1).Value = 49.5
Sheets("Vano").Cells(8, 2).Value = 1350
Sheets("Vano").Cells(8, 3).Value = 1100
Sheets("Vano").Cells(8, 4).Value = 0.2
Sheets("Vano").Cells(8, 5).Value = 0.2

Sheets("Vano").Cells(9, 1).Value = 45
Sheets("Vano").Cells(9, 2).Value = 1100
Sheets("Vano").Cells(9, 3).Value = 850
Sheets("Vano").Cells(9, 4).Value = 0.2
Sheets("Vano").Cells(9, 5).Value = 0.2

Sheets("Vano").Cells(10, 1).Value = 40.5
Sheets("Vano").Cells(10, 2).Value = 850
Sheets("Vano").Cells(10, 3).Value = 650
Sheets("Vano").Cells(10, 4).Value = 0.2
Sheets("Vano").Cells(10, 5).Value = 0.2

Sheets("Vano").Cells(11, 1).Value = 36
Sheets("Vano").Cells(11, 2).Value = 650
Sheets("Vano").Cells(11, 3).Value = 500
Sheets("Vano").Cells(11, 4).Value = 0.2
Sheets("Vano").Cells(11, 5).Value = 0.2

Sheets("Vano").Cells(12, 1).Value = 31.5
Sheets("Vano").Cells(12, 2).Value = 500
Sheets("Vano").Cells(12, 3).Value = 350
Sheets("Vano").Cells(12, 4).Value = 0.2
Sheets("Vano").Cells(12, 5).Value = 0.2

Sheets("Vano").Cells(13, 1).Value = 27
Sheets("Vano").Cells(13, 2).Value = 350
Sheets("Vano").Cells(13, 3).Value = 300
Sheets("Vano").Cells(13, 4).Value = 0.2
Sheets("Vano").Cells(13, 5).Value = 0.2


'//
'//CATENARIA MARRUECOS 3.000 Vcc (FIJADA POR MEMORANDUM)
'//

ElseIf nombre_catVB = "Marruecos 3.000 Vcc" And ventoso(poli) = "no" Then
Sheets("Vano").Cells(3, 1).Value = 63
Sheets("Vano").Cells(3, 2).Value = 100000
Sheets("Vano").Cells(3, 3).Value = 7500
Sheets("Vano").Cells(3, 4).Value = 0.2
Sheets("Vano").Cells(3, 5).Value = -0.2

Sheets("Vano").Cells(4, 1).Value = 63
Sheets("Vano").Cells(4, 2).Value = 7500
Sheets("Vano").Cells(4, 3).Value = 5000
Sheets("Vano").Cells(4, 4).Value = -0.2
Sheets("Vano").Cells(4, 5).Value = -0.04

Sheets("Vano").Cells(5, 1).Value = 63
Sheets("Vano").Cells(5, 2).Value = 5000
Sheets("Vano").Cells(5, 3).Value = 4000
Sheets("Vano").Cells(5, 4).Value = 0.2
Sheets("Vano").Cells(5, 5).Value = 0.04

Sheets("Vano").Cells(6, 1).Value = 63
Sheets("Vano").Cells(6, 2).Value = 4000
Sheets("Vano").Cells(6, 3).Value = 3000
Sheets("Vano").Cells(6, 4).Value = 0.2
Sheets("Vano").Cells(6, 5).Value = 0.12

Sheets("Vano").Cells(7, 1).Value = 63
Sheets("Vano").Cells(7, 2).Value = 3000
Sheets("Vano").Cells(7, 3).Value = 2600
Sheets("Vano").Cells(7, 4).Value = 0.2
Sheets("Vano").Cells(7, 5).Value = 0.2

Sheets("Vano").Cells(8, 1).Value = 58.5
Sheets("Vano").Cells(8, 2).Value = 2600
Sheets("Vano").Cells(8, 3).Value = 2500
Sheets("Vano").Cells(8, 4).Value = 0.2
Sheets("Vano").Cells(8, 5).Value = 0.15

Sheets("Vano").Cells(9, 1).Value = 58.5
Sheets("Vano").Cells(9, 2).Value = 2500
Sheets("Vano").Cells(9, 3).Value = 1900
Sheets("Vano").Cells(9, 4).Value = 0.2
Sheets("Vano").Cells(9, 5).Value = 0.2

Sheets("Vano").Cells(10, 1).Value = 54
Sheets("Vano").Cells(10, 2).Value = 1900
Sheets("Vano").Cells(10, 3).Value = 1400
Sheets("Vano").Cells(10, 4).Value = 0.2
Sheets("Vano").Cells(10, 5).Value = 0.2

Sheets("Vano").Cells(11, 1).Value = 49.5
Sheets("Vano").Cells(11, 2).Value = 1400
Sheets("Vano").Cells(11, 3).Value = 1100
Sheets("Vano").Cells(11, 4).Value = 0.2
Sheets("Vano").Cells(11, 5).Value = 0.2

Sheets("Vano").Cells(12, 1).Value = 45
Sheets("Vano").Cells(12, 2).Value = 1100
Sheets("Vano").Cells(12, 3).Value = 850
Sheets("Vano").Cells(12, 4).Value = 0.2
Sheets("Vano").Cells(12, 5).Value = 0.2

Sheets("Vano").Cells(13, 1).Value = 40.5
Sheets("Vano").Cells(13, 2).Value = 850
Sheets("Vano").Cells(13, 3).Value = 650
Sheets("Vano").Cells(13, 4).Value = 0.2
Sheets("Vano").Cells(13, 5).Value = 0.2

Sheets("Vano").Cells(14, 1).Value = 36
Sheets("Vano").Cells(14, 2).Value = 650
Sheets("Vano").Cells(14, 3).Value = 500
Sheets("Vano").Cells(14, 4).Value = 0.2
Sheets("Vano").Cells(14, 5).Value = 0.2

Sheets("Vano").Cells(15, 1).Value = 31.5
Sheets("Vano").Cells(15, 2).Value = 500
Sheets("Vano").Cells(15, 3).Value = 350
Sheets("Vano").Cells(15, 4).Value = 0.2
Sheets("Vano").Cells(15, 5).Value = 0.2

Sheets("Vano").Cells(16, 1).Value = 27
Sheets("Vano").Cells(16, 2).Value = 350
Sheets("Vano").Cells(16, 3).Value = 300
Sheets("Vano").Cells(16, 4).Value = 0.2
Sheets("Vano").Cells(16, 5).Value = 0.2
End If
End Sub

