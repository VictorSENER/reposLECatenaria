Attribute VB_Name = "tabla_vanos"
Sub tabla_vanos(nombre_catVB)

Dim cont(10) As Long

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

Call cargar.datos_acces(nombre_catVB)

'//
'//CÁLCULO FUERZA VIENTO
'//

coef_viento = 1.2

'//
'//Para el cálculo de la fuerza de viento se ha considerado un 15% de afectación en las péndolas
'//

fuerza_viento = 0.625 * coef_viento * (vw ^ 2) * (diam_hc + diam_sust + diam_pend * 0.15)

'//
'//ZONA RECTA
'//

    d_max_cu_var = d_max_cu
    i = 1
    
    Sheets(2).Cells(3, 2) = 100000
    
    cont(2) = 15000
    cont(3) = 7500
    cont(4) = 5000
    cont(5) = 4000
    cont(6) = 3000
    cont(7) = 2500
   
    d_media = (d_max_cu + d_max_cu_var) / 2
    t_tot = (n_hc * t_hc + t_sust) * 9.81 * 0.89
    
    va_max_re = 2 * Sqr((t_tot / fuerza_viento) * (d_max_ad + Sqr((d_max_ad ^ 2) - (d_max_re ^ 2))))
    
    va_max_re = Int(va_max_re)
    
    va_max_var = va_max
    
    While va_max_re < va_max_var
        While (va_max_re / inc_norm_va) <> Int(va_max_re / inc_norm_va)
            va_max_re = va_max_re - 0.5
        Wend
    
        va_max_var = va_max_re

    Wend
      
    Sheets(2).Cells(3, 4) = d_max_re
    Sheets(2).Cells(3, 5) = -d_max_re
    
    i = 1
    r = 3000
    va_max_cu = va_max_var
    
    j = 2
    r = r_re

n:
    Sheets(2).Cells(i + 2, j + 3) = d_max_cu
    
    While r > 2500
    r = cont(i + 1)
    d_max_cu_var = 2 * ((va_max_var ^ 2) / (8 * r)) - d_max_cu
    Sheets(2).Cells(i + 2, j + 1) = cont(i + 1)
    Sheets(2).Cells(i + 3, j) = cont(i + 1)
    Sheets(2).Cells(i + 3, j + 2) = d_max_cu
    Sheets(2).Cells(i + 3, j + 3) = d_max_cu_var
    Sheets(2).Cells(i + 2, j - 1) = va_max_var
    i = i + 1
    
        If r <= 2500 Then
        GoTo n
        End If
    
    Wend
        
    r = 2500
   
'//
'//ZONA CURVA
'//

inicio:

    While va_max_cu >= va_max_var And r >= r_min_traz
        r = r - 50
        va_max_cu = Sqr((8 * r * t_tot * (d_max_ad + d_max_cu)) / (fuerza_viento * r + t_tot))
    Wend
        
    Sheets(2).Cells(i + 3, 2) = r + 50
    Sheets(2).Cells(i + 3, 4) = d_max_cu
    Sheets(2).Cells(i + 3, 5) = d_max_cu
    
    If r < r_min_traz Then
    
        va_max_cu = Sqr((8 * r_min_traz * t_tot * (d_max_ad + d_max_cu)) / (fuerza_viento * r + t_tot))
        If va_max_cu >= va_max_var Then
            Sheets(2).Cells(i + 2, 3) = r_min_traz
            Sheets(2).Cells(i + 2, 1) = va_max_var
            GoTo fin
        Else
            Sheets(2).Cells(i + 2, 3) = r + 50
            Sheets(2).Cells(i + 2, 1) = va_max_var
            GoTo fin
        End If
    GoTo fin
    End If
    
    Sheets(2).Cells(i + 2, 3) = r + 50
    Sheets(2).Cells(i + 2, 1) = va_max_var
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

Sheets(2).Range(Sheets(2).Cells(i + 3, 1), Sheets(2).Cells(i + 10, 5)).ClearContents


'//
'//CATENARIA MARRUECOS 3.000 Vcc (FIJADA POR MEMORANDUM)
'//

ElseIf nombre_catVB = "Marruecos 3.000 Vcc" Then

Sheets(2).Cells(3, 1).Value = 54
Sheets(2).Cells(3, 2).Value = 100000
Sheets(2).Cells(3, 3).Value = 7500
Sheets(2).Cells(3, 4).Value = 0.2
Sheets(2).Cells(3, 5).Value = -0.2

Sheets(2).Cells(4, 1).Value = 54
Sheets(2).Cells(4, 2).Value = 7500
Sheets(2).Cells(4, 3).Value = 5000
Sheets(2).Cells(4, 4).Value = 0.2
Sheets(2).Cells(4, 5).Value = -0.04

Sheets(2).Cells(5, 1).Value = 54
Sheets(2).Cells(5, 2).Value = 5000
Sheets(2).Cells(5, 3).Value = 4000
Sheets(2).Cells(5, 4).Value = 0.2
Sheets(2).Cells(5, 5).Value = 0.04

Sheets(2).Cells(6, 1).Value = 54
Sheets(2).Cells(6, 2).Value = 4000
Sheets(2).Cells(6, 3).Value = 3000
Sheets(2).Cells(6, 4).Value = 0.2
Sheets(2).Cells(6, 5).Value = 0.12

Sheets(2).Cells(7, 1).Value = 54
Sheets(2).Cells(7, 2).Value = 3000
Sheets(2).Cells(7, 3).Value = 1350
Sheets(2).Cells(7, 4).Value = 0.2
Sheets(2).Cells(7, 5).Value = 0.2

Sheets(2).Cells(8, 1).Value = 49.5
Sheets(2).Cells(8, 2).Value = 1350
Sheets(2).Cells(8, 3).Value = 1100
Sheets(2).Cells(8, 4).Value = 0.2
Sheets(2).Cells(8, 5).Value = 0.2

Sheets(2).Cells(9, 1).Value = 45
Sheets(2).Cells(9, 2).Value = 1100
Sheets(2).Cells(9, 3).Value = 850
Sheets(2).Cells(9, 4).Value = 0.2
Sheets(2).Cells(9, 5).Value = 0.2

Sheets(2).Cells(10, 1).Value = 40.5
Sheets(2).Cells(10, 2).Value = 850
Sheets(2).Cells(10, 3).Value = 650
Sheets(2).Cells(10, 4).Value = 0.2
Sheets(2).Cells(10, 5).Value = 0.2

Sheets(2).Cells(11, 1).Value = 36
Sheets(2).Cells(11, 2).Value = 650
Sheets(2).Cells(11, 3).Value = 500
Sheets(2).Cells(11, 4).Value = 0.2
Sheets(2).Cells(11, 5).Value = 0.2

Sheets(2).Cells(12, 1).Value = 31.5
Sheets(2).Cells(12, 2).Value = 500
Sheets(2).Cells(12, 3).Value = 350
Sheets(2).Cells(12, 4).Value = 0.2
Sheets(2).Cells(12, 5).Value = 0.2

Sheets(2).Cells(13, 1).Value = 27
Sheets(2).Cells(13, 2).Value = 350
Sheets(2).Cells(13, 3).Value = 300
Sheets(2).Cells(13, 4).Value = 0.2
Sheets(2).Cells(13, 5).Value = 0.2

End If
End Sub

