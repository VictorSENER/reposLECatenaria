Attribute VB_Name = "momento"
' queda la duda de si el momento de atirantado en recta se debe sumar en los momentos de curva

Sub momento(nombre_catVB, fila, CAD)
Dim coef_viento As Double
Dim va_ini As Double, va_fin As Double
Dim r As Double, d_0 As Double, d_1 As Double, d_2 As Double, d_3 As Double, d_4 As Double
Dim tip As String
Dim fuerza_at_re_ejey_sust As Double, fuerza_at_cu_ejey_sust As Double, fuerza_at_ejey_sust As Double
Dim fuerza_at_re_ejex_sust As Double, fuerza_at_cu_ejex_sust As Double, fuerza_at_ejex_sust As Double
Dim fuerza_at_re_ejey_hc As Double, fuerza_at_cu_ejey_hc As Double, fuerza_at_ejey_hc As Double, fuerza_at_re_ejex_hc As Double
Dim fuerza_at_cu_ejex_hc As Double, fuerza_at_ejex_hc As Double, fuerza_at_ejey_feed_pos As Double, fuerza_at_cu_ejey_feed_pos As Double
Dim fuerza_at_ejex_feed_pos As Double, fuerza_at_cu_ejex_feed_pos As Double, fuerza_at_ejey_feed_neg As Double, fuerza_at_cu_ejey_feed_neg As Double
Dim fuerza_at_ejex_feed_neg As Double, fuerza_at_cu_ejex_feed_neg As Double, fuerza_at_ejey_cdpa As Double, fuerza_at_cu_ejey_cdpa As Double
Dim fuerza_at_ejex_cdpa As Double, fuerza_at_cu_ejex_cdpa As Double, mom_sust As Double, mom_sust_1 As Double, mom_hc As Double
Dim mom_hc_1 As Double, mom_feed_pos As Double, mom_feed_neg As Double, mom_cdpa As Double, mom_equip As Double
Dim mom_sust_anc As Double, mom_hc_anc As Double, mom_tot As Double, mom_tot_1 As Double, mom_tot_2 As Double
Dim pres_viento_sust As Double, pres_viento_hc As Double, pres_viento_feed_pos As Double, pres_viento_feed_neg As Double
Dim pres_viento_pto_fijo As Double, pres_viento_cdpa As Double, pres_viento_pend As Double, fuerza_viento_sust As Double
Dim fuerza_viento_hc As Double, fuerza_viento_feed_pos As Double, fuerza_viento_feed_neg As Double, fuerza_viento_pto_fijo As Double
Dim fuerza_viento_cdpa As Double, fuerza_viento_pend As Double, masa_sust As Double, masa_hc As Double, masa_feed_pos As Double
Dim masa_feed_neg As Double, masa_pto_fijo As Double, masa_cdpa As Double, masa_pend As Double, masa_equip As Double
Dim el_hc_0 As Double, el_hc_1 As Double, el_hc_2 As Double, el_hc_3 As Double, el_hc_4 As Double, el_hc_5 As Double
'//
'//LECTURA BASE DE DATOS
'//
Call cargar.datos_lac(nombre_catVB)
Call cargar.datos_cable

   
'///
'///
'///
p_sust1 = p_sust * 9.81 / 10
p_hc1 = p_hc * 9.81 / 10
p_cdpa1 = p_cdpa * 9.81 / 10
p_feed_pos1 = p_feed_pos * 9.81 / 10
p_pto_fijo1 = p_pto_fijo * 9.81 / 10

'//
'//LECTURA DATOS REPLANTEO
'//

While Not IsEmpty(Sheets("Replanteo").Cells(fila + 2, 33).Value)
If Sheets("Replanteo").Cells(fila, 38).Value <> "Tunel" And Sheets("Replanteo").Cells(fila, 38).Value <> "Marquesina" And Sheets("Replanteo").Cells(fila, 38).Value <> "Viaducto" Then

    If CAD = True Then
        tip_poste = Sheets("Replanteo").Cells(fila, 18).Value
            If tip_poste = "X1-7.8" Or tip_poste = "X2-7.8" Or tip_poste = "X3-7.8" Then
                base_poste = 0.46
                cabeza_poste = 0.2
                alt_nenc_poste = 7.8
                tgx = ((base_poste - cabeza_poste) / 2) / alt_nenc_poste
            ElseIf tip_poste = "X1-8" Or tip_poste = "X2-8" Or tip_poste = "X3-8" Then
                base_poste = 0.46
                cabeza_poste = 0.2
                alt_nenc_poste = 8
                tgx = ((base_poste - cabeza_poste) / 2) / alt_nenc_poste
            ElseIf tip_poste = "X1-8.5" Or tip_poste = "X2-8.5" Or tip_poste = "X3-8.5" Then
                base_poste = 0.46
                cabeza_poste = 0.2
                alt_nenc_poste = 8.5
                tgx = ((base_poste - cabeza_poste) / 2) / alt_nenc_poste
            ElseIf tip_poste = "Z1" Then
                base_poste = 0.525
                cabeza_poste = 0.25
                alt_nenc_poste = 9.75
                tgx = ((base_poste - cabeza_poste) / 2) / alt_nenc_poste
            ElseIf tip_poste = "X3A" Or tip_poste = "Z3" Or tip_poste = "Z5" Then
                base_poste = 0.5
                cabeza_poste = 0.25
                alt_nenc_poste = 8.75
                tgx = ((base_poste - cabeza_poste) / 2) / alt_nenc_poste
            ElseIf tip_poste = "Z6bis" Then
                base_poste = 0.7
                cabeza_poste = 0.35
                alt_nenc_poste = 10.25
                tgx = ((base_poste - cabeza_poste) / 2) / alt_nenc_poste
            End If
    
        ancho_medio_poste = base_poste / 2 'valor estándar supuesto
        sup_perf_max_poste = alt_nenc_poste * 0.1 ' corregir esto
        arasa = Sheets("Replanteo").Cells(fila, 20).Value
    Else
        ancho_medio_poste = base_poste / 2 'valor estándar supuesto
        sup_perf_max_poste = alt_nenc_poste * 0.1 ' corregir esto
        arasa = Sheets("Replanteo").Cells(fila, 20).Value
        alt_nenc_poste = 7.8
    End If
        
        alt_nom_1 = Sheets("Replanteo").Cells(fila, 10).Value
        alt_nom_2 = Sheets("Replanteo").Cells(fila + 2, 10).Value
        
        dist_carril_poste_1 = Sheets("Replanteo").Cells(fila, 5).Value
        dist_carril_poste_2 = Sheets("Replanteo").Cells(fila + 2, 5).Value
        lado = Sheets("Replanteo").Cells(fila, 30).Value
        va_fin = Sheets("Replanteo").Cells(fila + 1, 4).Value
        
        d_1 = Sheets("Replanteo").Cells(fila, 8).Value
        d_2 = Sheets("Replanteo").Cells(fila + 2, 8).Value
        d_4 = Sheets("Replanteo").Cells(fila, 9).Value
        d_5 = Sheets("Replanteo").Cells(fila + 2, 9).Value
    If CAD = False Then
        If Sheets("Replanteo").Cells(fila, 16).Value = anc_aguj & " + " & semi_eje_sla Then
            tip_1 = semi_eje_sla
            tip_pf_1 = anc_aguj
        ElseIf Sheets("Replanteo").Cells(fila, 16).Value = semi_eje_aguj & " + " & anc_sla_con Then
            tip_1 = anc_sla_con
            tip_pf_1 = semi_eje_aguj
        ElseIf Len(Sheets("Replanteo").Cells(fila, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(fila, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(fila, 16).Value = anc_sm_sin) Then
            tip_1 = Mid(Sheets("Replanteo").Cells(fila, 16).Value, 15)
            tip_pf_1 = Mid(Sheets("Replanteo").Cells(fila, 16).Value, 1, 11)
        Else
            tip_1 = Sheets("Replanteo").Cells(fila, 16).Value
            tip_pf_1 = Sheets("Replanteo").Cells(fila, 16).Value
        End If
        If Sheets("Replanteo").Cells(fila - 2, 16).Value = anc_aguj & " + " & semi_eje_sla Then
            tip_0 = semi_eje_sla
            tip_pf_0 = anc_aguj
        ElseIf Sheets("Replanteo").Cells(fila - 2, 16).Value = semi_eje_aguj & " + " & anc_sla_con Then
            tip_0 = anc_sla_con
            tip_pf_0 = semi_eje_aguj
        ElseIf Len(Sheets("Replanteo").Cells(fila - 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(fila - 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(fila - 2, 16).Value = anc_sm_sin) Then
            tip_0 = Mid(Sheets("Replanteo").Cells(fila - 2, 16).Value, 15)
            tip_pf_0 = Mid(Sheets("Replanteo").Cells(fila - 2, 16).Value, 1, 11)
        Else
            tip_0 = Sheets("Replanteo").Cells(fila - 2, 16).Value
            tip_pf_0 = Sheets("Replanteo").Cells(fila - 2, 16).Value
        End If
        If Sheets("Replanteo").Cells(fila + 2, 16).Value = anc_aguj & " + " & semi_eje_sla Then
            tip_2 = semi_eje_sla
            tip_pf_2 = anc_aguj
        ElseIf Sheets("Replanteo").Cells(fila + 2, 16).Value = semi_eje_aguj & " + " & anc_sla_con Then
            tip_2 = anc_sla_con
            tip_pf_2 = semi_eje_aguj
        ElseIf Len(Sheets("Replanteo").Cells(fila + 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(fila + 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(fila + 2, 16).Value = anc_sm_sin) Then
            tip_2 = Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 15)
            tip_pf_2 = Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 1, 11)
        Else
            tip_2 = Sheets("Replanteo").Cells(fila + 2, 16).Value
            tip_pf_2 = Sheets("Replanteo").Cells(fila + 2, 16).Value
        End If
    End If
    If fila <> 10 Then
        va_ini = Sheets("Replanteo").Cells(fila - 1, 4).Value
        d_0 = Sheets("Replanteo").Cells(fila - 2, 8).Value
        d_3 = Sheets("Replanteo").Cells(fila - 2, 9).Value
        alt_nom_0 = Sheets("Replanteo").Cells(fila - 2, 10).Value
        dist_carril_poste_0 = Sheets("Replanteo").Cells(fila - 2, 5).Value
    Else
        va_ini = va_fin
        d_0 = d_1
        d_3 = d_4
        alt_nom_0 = alt_nom
        dist_carril_poste_0 = Sheets("Replanteo").Cells(fila, 5).Value
    End If
    r = Sheets("Replanteo").Cells(fila, 6).Value
    
    If lado = "D" Then
        d_0 = -d_0
        d_1 = -d_1
        d_2 = -d_2
        d_3 = -d_3
        d_4 = -d_4
        d_5 = -d_5
    End If
    If r = 0 Then
        r = 1000000000000#
        ayuda = 1
    ElseIf (lado = "D" And r > 0) Or (lado = "G" And r < 0) Then
        ayuda = -1
    Else
        ayuda = 1
    End If
    
    r = Abs(r)
'///
'/// Cálculo distancias entre hilos de contacto y punto medio del poste
'///
    dist_hc_poste_0 = (ancho_via / 2) + ancho_carril + dist_carril_poste_0 - d_0 + ancho_medio_poste
    dist_hc_poste_1 = (ancho_via / 2) + ancho_carril + dist_carril_poste_1 - d_1 + ancho_medio_poste
    dist_hc_poste_2 = (ancho_via / 2) + ancho_carril + dist_carril_poste_2 - d_2 + ancho_medio_poste
    dist_hc_poste_3 = (ancho_via / 2) + ancho_carril + dist_carril_poste_0 - d_3 + ancho_medio_poste
    dist_hc_poste_4 = (ancho_via / 2) + ancho_carril + dist_carril_poste_1 - d_4 + ancho_medio_poste
    dist_hc_poste_5 = (ancho_via / 2) + ancho_carril + dist_carril_poste_2 - d_5 + ancho_medio_poste

    
    If tip_2 <> "" And (tip_1 = anc_sm_sin Or tip_1 = anc_sla_sin Or tip_1 = anc_sla_con Or tip_1 = anc_sm_con) Then
        el_hc_5 = Sheets("Replanteo").Cells(fila + 2, 46).Value
        el_hc_0 = 0
        
    ElseIf (tip_1 = eje_sm And tip_2 = eje_sm) Or (tip_1 = eje_sla And tip_2 = eje_sla) Then
        el_hc_2 = Sheets("Replanteo").Cells(fila - 2, 46).Value
        el_hc_3 = Sheets("Replanteo").Cells(fila - 2, 46).Value

    ElseIf (tip_1 = eje_sm And tip_0 = eje_sm) Or (tip_1 = eje_sla And tip_0 = eje_sla) Then
        el_hc_2 = Sheets("Replanteo").Cells(fila + 2, 40).Value
        el_hc_0 = 0
    ElseIf (tip_1 = eje_sm Or tip_1 = eje_sla) And (tip_2 = semi_eje_sm Or tip_2 = semi_eje_sla) Then
        el_hc_2 = Sheets("Replanteo").Cells(fila + 2, 40).Value
        el_hc_3 = Sheets("Replanteo").Cells(fila - 2, 46).Value
        el_hc_0 = 0
    ElseIf (tip_1 = semi_eje_sm And tip_2 = eje_sm) Or (tip_1 = semi_eje_sla And tip_2 = eje_sla) Then
        el_hc_4 = Sheets("Replanteo").Cells(fila, 46).Value
        el_hc_0 = 0
    ElseIf (tip_1 = semi_eje_sm And tip_0 = eje_sm) Or (tip_1 = semi_eje_sla And tip_0 = eje_sla) Then
        el_hc_1 = Sheets("Replanteo").Cells(fila, 40).Value
        el_hc_0 = 0
    ElseIf tip_2 = "" And (tip_1 = anc_sm_sin Or tip_1 = anc_sla_sin Or tip_1 = anc_sla_con Or tip_1 = anc_sm_con) Then
        el_hc_0 = Sheets("Replanteo").Cells(fila - 2, 40).Value
        el_hc_2 = 0

    Else
        el_hc_0 = 0
        el_hc_2 = 0
    End If
        masa_equip = p_medio_equip_t
        dist_horiz_equip = dist_horiz_equip_t
    '///
    '/// Variación alt_cat
    '///
    If Not IsEmpty(Sheets("Replanteo").Cells(fila, 39).Value) Then
        alt_cat_1 = Sheets("Replanteo").Cells(fila, 39).Value
        alt_cat_4 = Sheets("Replanteo").Cells(fila, 45).Value
    Else
        alt_cat_1 = alt_cat
    End If
    If Not IsEmpty(Sheets("Replanteo").Cells(fila - 2, 39).Value) Then
        alt_cat_0 = Sheets("Replanteo").Cells(fila - 2, 39).Value
        alt_cat_3 = Sheets("Replanteo").Cells(fila - 2, 45).Value
    Else
        alt_cat_0 = alt_cat
    End If
    If Not IsEmpty(Sheets("Replanteo").Cells(fila + 2, 39).Value) Then
        alt_cat_2 = Sheets("Replanteo").Cells(fila + 2, 39).Value
        alt_cat_5 = Sheets("Replanteo").Cells(fila + 2, 45).Value
    Else
        alt_cat_2 = alt_cat
    End If
    
    
    var_0 = 1
    var_1 = 1
    n_sust = 1
    If posicion_feed_pos = "apoyado" Then
        dist_horiz_feed_pos1 = 0
    ElseIf posicion_feed_pos = "Suspendido (lado exterior)" Then
        dist_horiz_feed_pos1 = -(dist_horiz_feed_pos)
    ElseIf posicion_feed_pos = "Suspendido (lado vía)" Then
        dist_horiz_feed_pos1 = (dist_horiz_feed_pos)
    ElseIf posicion_feed_pos = "NO HAY" Then
        var_0 = 0
    End If
    
    If posicion_feed_neg = "apoyado" Then
        dist_horiz_feed_neg = 0
    ElseIf posicion_feed_neg = "Suspendido (lado exterior)" Then
        dist_horiz_feed_neg = -(dist_horiz_feed_neg)
    ElseIf posicion_feed_neg = "Suspendido (lado vía)" Then
        dist_horiz_feed_neg = (dist_horiz_feed_neg)
    ElseIf posicion_feed_neg = "NO HAY" Then
        var_1 = 0
    End If


'//
'//CÁLCULO FUERZA VIENTO
'//

    fuerza_viento_sust = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_ini + va_fin) / 2

    fuerza_viento_hc = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_ini + va_fin) / 2

    fuerza_viento_feed_pos = fuerza_viento_cyc(diam_feed_pos, sec_feed_pos, n_feed_pos, "feed_pos", adm_lin_poste) * (va_ini + va_fin) / 2
    
    fuerza_viento_feed_neg = fuerza_viento_cyc(diam_feed_neg, sec_feed_neg, n_feed_neg, "feed_neg", adm_lin_poste) * (va_ini + va_fin) / 2
   
    'fuerza_viento_pto_fijo = fuerza_viento_cyc(diam_pto_fijo, sec_pto_fijo, n_pto_fijo, "pto_fijo", adm_lin_poste) * (va_ini + va_fin) / 2
    
    fuerza_viento_cdpa = fuerza_viento_cyc(diam_cdpa, sec_cdpa, n_cdpa, "cdpa", adm_lin_poste) * (va_ini + va_fin) / 2
   
    'fuerza_viento_pend = fuerza_viento_cyc(diam_pend, sec_pend, n_pend, "pend", adm_lin_poste) * (va_ini + va_fin) / 2 * 0.15 'multiplicar por un porcentaje a concretar
    
    fuerza_viento_poste = fuerza_viento_sup(adm_lin_poste) * sup_perf_max_poste
    
'//
'//CÁLCULO PESO CONDUCTORES
'//
    
    masa_sust = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_1) - (alt_nom_0 + alt_cat_0)) / va_ini) + (((alt_nom_1 + alt_cat_1) - (alt_nom_2 + alt_cat_2)) / va_fin)))
    masa_hc = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - (alt_nom_0 + el_hc_0)) / va_ini) + ((alt_nom_1 - alt_nom_2) / va_fin)))
    masa_feed_pos = n_feed_pos * p_feed_pos1 * (va_ini + va_fin) / 2
    masa_feed_neg = n_feed_neg * p_feed_neg * (va_ini + va_fin) / 2
    'masa_pto_fijo = p_pto_fijo * (va_ini + va_fin) / 2
    masa_cdpa = n_cdpa * p_cdpa1 * (va_ini + va_fin) / 2
    'masa_pend = p_pend * (va_ini + va_fin) / 2 'multiplicar por un porcentaje a concretar

'//
'//CÁLCULO FUERZA RADIAL
'//

    fuerza_at_ejey_sust = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_1 - d_0) / (va_ini) + (d_1 - d_2) / (va_fin))
    'fuerza_at_ejex_sust = n_sust * t_sust * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_ejey_hc = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_1 - d_0) / (va_ini) + (d_1 - d_2) / (va_fin))
    'fuerza_at_ejex_hc = n_hc * t_hc * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_ejey_feed_pos = n_feed_pos * t_feed_pos * ((va_ini + va_fin) / (2 * r))
    'fuerza_at_ejex_feed_pos = n_feed_pos * t_feed_pos * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_ejey_feed_neg = n_feed_neg * t_feed_neg * ((va_ini + va_fin) / (2 * r))
    'fuerza_at_ejex_feed_neg = n_feed_neg * t_feed_neg * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_ejey_cdpa = n_cdpa * t_cdpa * ((va_ini + va_fin) / (2 * r))
    'fuerza_at_ejex_cdpa = n_cdpa * t_cdpa * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
'//
'//COMPARACIÓN ESFUERZO RADIAL
'//
    '///
    '/// Cálculo del angulo del feeder
    '///
    angulo = ayuda * Atn(fuerza_at_ejey_feed_pos / masa_feed_pos) * 180 / 3.1415 ' cambiar para feeder
    Sheets("Replanteo").Cells(fila, 34).Value = angulo
    
    '//
    '//CÁLCULO DISTANCIAS
    '//
    If Abs(angulo) <= 40 Then
        dist_vert_feed_pos1 = alt_nenc_poste - 0.512 - Abs(Cos(angulo * 3.1415 / 180) * 0.6515)
        dist_horiz_feed_pos1 = 1.075 - (sin(angulo * 3.1415 / 180) * 0.6515) + (base_poste / 2) - tgx * (dist_vert_feed_pos1)
    Else
        dist_vert_feed_pos1 = alt_nenc_poste - 0.872 - Abs(Cos(angulo * 3.1415 / 180) * 0.6415)
        dist_horiz_feed_pos1 = 0.925 - (sin(angulo * 3.1415 / 180) * 0.6415) + (base_poste / 2) - tgx * (dist_vert_feed_pos1)
    End If
    dist_horiz_cdpa = (base_poste / 2) - tgx * (dist_vert_cdpa_pos1)
    dist_vert_feed_neg1 = dist_vert_feed_neg + arasa
    dist_vert_cdpa1 = dist_vert_cdpa + arasa
    dist_horiz_equip1 = dist_horiz_equip_t
    dist_vert_pf_anc1 = 6.8
    dist_horiz_cdpa = (base_poste / 2) - tgx * (dist_vert_cdpa1)
    
'//
'//MOMENTO POSTE SIMPLE
'//

    If tip_1 = "" Then
        
        mom_sust_1 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        
        mom_hc_1 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
                        
        mom_feed_pos_1 = (ayuda * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        
        mom_feed_neg = (fuerza_viento_feed_neg * dist_vert_feed_neg1) - (masa_feed_neg * dist_horiz_feed_neg) + (fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
                
        mom_cdpa_1 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        
        mom_equip = (masa_equip * dist_horiz_equip)
        
        mom_poste_1 = (ayuda * fuerza_viento_poste * (alt_nenc_poste / 2))
        mom_poste_2 = (fuerza_viento_poste * (alt_nenc_poste / 2))
        
        mom_tot_1 = (mom_sust_1 + mom_hc_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_poste_1)
        mom_tot_2 = (mom_sust_2 + mom_hc_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_poste_2)
        
        mom_tot = MAX(mom_tot_1, mom_tot_2)
'//
'//MOMENTO POSTE SECCIONAMIENTO MECÁNICO
'//

    ElseIf tip_1 = anc_sm_con Or tip_1 = anc_sm_sin Or tip_1 = anc_sla_sin Or tip_1 = anc_sla_con Then
        
        If tip_2 = semi_eje_sm Or tip_2 = semi_eje_sla Then
            
            masa_sust = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_1) - (alt_nom_0 + alt_cat_0)) / va_ini) + (((alt_nom_1 + alt_cat_1) - (alt_nom_2 + alt_cat_2)) / va_fin)))
            masa_hc = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - (alt_nom_0)) / va_ini) + ((alt_nom_1 - alt_nom_2) / va_fin)))
                        
            dist_vert_hc_anc1 = Sheets("Replanteo").Cells(fila + 1, 46).Value + alt_nom_2
            dist_vert_sust_anc1 = dist_vert_hc_anc1 + 0.5
        
            masa_sust_2 = (n_sust * p_sust1 * (va_fin) / 2) + (n_sust * t_sust * ((dist_vert_sust_anc1 - (alt_nom_2 + alt_cat_5)) / va_fin))
            masa_hc_2 = (n_hc * p_hc1 * (va_fin) / 2) + (n_hc * t_hc * ((dist_vert_hc_anc1 - (alt_nom_2 + el_hc_5)) / va_fin))
            
            fuerza_ejey_sust_anc = t_sust * sin(Atn(((dist_hc_poste_5 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin)))
            fuerza_ejey_hc_anc = n_hc * t_hc * sin(Atn(((dist_hc_poste_5 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin)))
            
            fuerza_viento_sust_s = Cos(Atn(((dist_hc_poste_5 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin))) * fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_fin / 2)
            fuerza_viento_hc_s = Cos(Atn(((dist_hc_poste_5 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin))) * fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_fin / 2)
            
        ElseIf tip_0 = semi_eje_sm Or tip_0 = semi_eje_sla Then
            '/// corregir la fuerza de atirantado en seccionamientos
            
            masa_sust = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_1) - (alt_nom_0 + alt_cat_3)) / va_ini) + (((alt_nom_1 + alt_cat_1) - (alt_nom_2 + alt_cat_2)) / va_fin)))
            masa_hc = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - alt_nom_0) / va_ini) + ((alt_nom_1 - alt_nom_2) / va_fin)))
    
            dist_vert_hc_anc1 = Sheets("Replanteo").Cells(fila - 1, 48).Value + alt_nom_0
            dist_vert_sust_anc1 = dist_vert_hc_anc1 + 0.5
            
            fuerza_at_ejey_sust = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_1 - d_3) / (va_ini) + (d_1 - d_2) / (va_fin))
            fuerza_at_ejey_hc = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_1 - d_3) / (va_ini) + (d_1 - d_2) / (va_fin))
            
            masa_sust_2 = (n_sust * p_sust1 * (va_ini) / 2) + (n_sust * t_sust * ((dist_vert_sust_anc1 - (alt_nom_0 + alt_cat_0)) / va_ini))
            masa_hc_2 = (n_hc * p_hc1 * (va_ini) / 2) + (n_hc * t_hc * ((dist_vert_hc_anc1 - (alt_nom_0 + el_hc_0)) / va_ini))
            
            fuerza_ejey_sust_anc = t_sust * sin(Atn(((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini)))
            fuerza_ejey_hc_anc = n_hc * t_hc * sin(Atn(((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini)))
            
            fuerza_viento_sust_s = Cos(Atn(((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini))) * fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_ini / 2)
            fuerza_viento_hc_s = Cos(Atn(((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini))) * fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_ini / 2)
        End If
        
        mom_sust_1 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        
        mom_hc_1 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        
        mom_feed_pos_1 = (ayuda * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        
        'mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        
        mom_cdpa_1 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        
        mom_equip = masa_equip * dist_horiz_equip
        
        mom_sust_anc_1 = ((ayuda * fuerza_viento_sust_s) * dist_vert_sust_anc1 + arasa) + (fuerza_ejey_sust_anc * dist_vert_sust_anc1 + arasa) + (masa_sust_2 * 0.2)
        mom_sust_anc_2 = ((fuerza_viento_sust_s) * dist_vert_sust_anc1 + arasa) + (fuerza_ejey_sust_anc * dist_vert_sust_anc1 + arasa) + (masa_sust_2 * 0.2)
        
        mom_hc_anc_1 = ((ayuda * fuerza_viento_hc_s) * dist_vert_hc_anc1 + arasa) + (fuerza_ejey_hc_anc * dist_vert_hc_anc1 + arasa) + (masa_hc_2 * 0.2)
        mom_hc_anc_2 = ((fuerza_viento_hc_s) * dist_vert_hc_anc1 + arasa) + (fuerza_ejey_hc_anc * dist_vert_hc_anc1 + arasa) + (masa_hc_2 * 0.2)
        
        mom_poste_1 = ayuda * fuerza_viento_poste * (alt_nenc_poste / 2)
        mom_poste_2 = (fuerza_viento_poste * (alt_nenc_poste / 2))
        
        mom_tot_1 = (mom_sust_1 + mom_hc_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_sust_anc_1 + mom_hc_anc_1 + mom_poste_1)
        mom_tot_2 = (mom_sust_2 + mom_hc_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_sust_anc_2 + mom_hc_anc_2 + mom_poste_2)

        mom_tot = MAX(mom_tot_1, mom_tot_2)
'///
'/// Cálculo del momento en semi-ejes de seccionamiento mecánico o seccionamiento de lámina de aire
'///
    ElseIf tip_1 = semi_eje_sm Or tip_1 = semi_eje_sla Then ' elegir que momento es mas desfavorable (fuera o dentro)
'///
'/// Cálculo del momento en el primer semi-eje de seccionamiento mecánico o seccionamiento de lámina de aire
'///
        If tip_2 = eje_sm Or tip_2 = eje_sla Then
            
            masa_sust = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_1) - (alt_nom_0 + alt_cat_0)) / va_ini) + (((alt_nom_1 + alt_cat_1) - (alt_nom_2 + alt_cat_2)) / va_fin)))
            masa_hc = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - alt_nom_0) / va_ini) + ((alt_nom_1 - alt_nom_2) / va_fin)))
            
            dist_vert_hc_anc1 = Sheets("Replanteo").Cells(fila - 1, 46).Value + alt_nom_1
            dist_vert_sust_anc1 = dist_vert_hc_anc1 + 0.5
            
            masa_sust_2 = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * (((alt_nom_1 + alt_cat_4) - (alt_nom_2 + alt_cat_5)) / va_fin)) + (n_sust * t_sust * (((alt_nom_1 + alt_cat_4) - dist_vert_sust_anc1) / va_ini))
            masa_hc_2 = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * ((alt_nom_1 + el_hc_4 - (alt_nom_2)) / va_fin)) + (n_hc * t_hc * (((alt_nom_1 + el_hc_4) - dist_vert_hc_anc1) / va_ini))
            
            fuerza_at_ejey_sust_sm = (n_sust * t_sust * ((va_fin / (2 * r)) + (d_4 - d_5) / (va_fin))) - (n_sust * t_sust * sin(Atn(((dist_hc_poste_4 + (dist_carril_poste_0 - dist_carril_poste_1)) / va_ini))))
            fuerza_at_ejey_hc_sm = (n_hc * t_hc * ((va_fin / (2 * r)) + (d_4 - d_5) / (va_fin))) - (n_hc * t_hc * sin(Atn((dist_hc_poste_4 + (dist_carril_poste_0 - dist_carril_poste_1)) / va_ini)))
            
            fuerza_viento_sust_sm = (Cos(Atn(((dist_hc_poste_4 + (dist_carril_poste_0 - dist_carril_poste_1)) / va_ini)))) * fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_ini + va_fin) / 2
            fuerza_viento_hc_sm = (Cos(Atn(((dist_hc_poste_4 + (dist_carril_poste_0 - dist_carril_poste_1)) / va_ini)))) * fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_ini + va_fin) / 2
            
            mom_sust_1 = (-1 * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
            mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
            
            mom_hc_1 = (-1 * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
            mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
            
            mom_sust_se_sm_el_1 = (-1 * fuerza_viento_sust_sm * (alt_nom_1 + alt_cat_4 + arasa)) + (masa_sust_2 * dist_hc_poste_4) + (fuerza_at_ejey_sust_sm * (alt_cat_4 + arasa + alt_nom_1))
            mom_sust_se_sm_el_2 = (fuerza_viento_sust_sm * (alt_nom_1 + alt_cat_4 + arasa)) + (masa_sust_2 * dist_hc_poste_4) + (fuerza_at_ejey_sust_sm * (alt_cat_4 + arasa + alt_nom_1))
                        
            mom_hc_se_sm_el_1 = (-1 * fuerza_viento_hc_sm * (alt_nom_1 + arasa + el_hc_4)) + (masa_hc_2 * dist_hc_poste_4) + (fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc_4))
            mom_hc_se_sm_el_2 = (fuerza_viento_hc_sm * (alt_nom_1 + arasa + el_hc_4)) + (masa_hc_2 * dist_hc_poste_4) + (fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc_4))
            
            mom_feed_pos_1 = (-1 * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
            mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
            
            mom_cdpa_1 = (-1 * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
            mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
            
            mom_equip = 2 * masa_equip * dist_horiz_equip
            
            mom_poste_1 = -1 * fuerza_viento_poste * (alt_nenc_poste / 2)
            mom_poste_2 = fuerza_viento_poste * (alt_nenc_poste / 2)
            
            mom_tot_1 = (mom_sust_1 + mom_sust_se_sm_el_1 + mom_hc_1 + mom_hc_se_sm_el_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_poste_1)
            mom_tot_2 = (mom_sust_2 + mom_sust_se_sm_el_2 + mom_hc_2 + mom_hc_se_sm_el_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_poste_2)
            
            mom_tot = MAX(mom_tot_1, mom_tot_2)
'///
'/// Cálculo del momento en el segundo semi-eje de seccionamiento mecánico o seccionamiento de lámina de aire
'///
        Else
            masa_sust = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_4) - (alt_nom_0 + alt_cat_3)) / va_ini) + (((alt_nom_1 + alt_cat_4) - (alt_nom_2 + alt_cat_2)) / va_fin)))
            masa_hc = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - alt_nom_0) / va_ini) + ((alt_nom_1 - alt_nom_2) / va_fin)))
            
            dist_vert_hc_anc1 = Sheets("Replanteo").Cells(fila + 1, 48).Value + alt_nom_1
            dist_vert_sust_anc1 = dist_vert_hc_anc1 + 0.5
            
            fuerza_at_ejey_sust = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_4 - d_3) / (va_ini) + (d_4 - d_2) / (va_fin))
            fuerza_at_ejey_hc = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_4 - d_3) / (va_ini) + (d_4 - d_2) / (va_fin))
            
            masa_sust_2 = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * (((alt_nom_1 + alt_cat_1) - dist_vert_sust_anc1) / va_fin)) + (n_sust * t_sust * (((alt_nom_1 + alt_cat_1) - (alt_nom_0 + alt_cat_0)) / va_ini))
            masa_hc_2 = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 + el_hc_1) - dist_vert_hc_anc1) / va_fin)) + (n_hc * t_hc * (((alt_nom_1 + el_hc_1) - alt_nom_0) / va_ini))
                                    
            fuerza_at_ejey_sust_sm = n_sust * t_sust * (((va_ini) / (2 * r)) + (d_1 - d_0) / (va_ini)) - n_sust * t_sust * sin(Atn(((dist_hc_poste_1 + (dist_carril_poste_2 - dist_carril_poste_1)) / va_fin)))
            'algo = n_hc * t_hc * sin(Atn(((dist_hc_poste_1 + (dist_carril_poste_2 - dist_carril_poste_1)) / va_fin)))
            
            fuerza_at_ejey_hc_sm = (n_hc * t_hc * ((va_ini) / (2 * r) + (d_1 - d_0) / (va_ini))) - n_hc * t_hc * sin(Atn(((dist_hc_poste_1 + (dist_carril_poste_2 - dist_carril_poste_1)) / va_fin)))
            
            fuerza_viento_sust_sm = Cos(Atn(((dist_hc_poste_1 + (dist_carril_poste_2 - dist_carril_poste_1)) / va_fin))) * fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_ini) / 2 + fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_fin) / 2
            fuerza_viento_hc_sm = Cos(Atn(((dist_hc_poste_1 + (dist_carril_poste_2 - dist_carril_poste_1)) / va_fin))) * fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_ini) / 2 + fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_fin) / 2
            
            mom_sust_1 = (-1 * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_4)) + (masa_sust * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_4))
            mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_4)) + (masa_sust * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_4))
            
            mom_sust_se_sm_el_1 = (-1 * fuerza_viento_sust_sm * (alt_nom_1 + alt_cat_1 + arasa)) + (masa_sust_2 * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust_sm * (alt_cat_1 + arasa + alt_nom_1))
            mom_sust_se_sm_el_2 = (fuerza_viento_sust_sm * (alt_nom_1 + alt_cat_1 + arasa)) + (masa_sust_2 * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust_sm * (alt_cat_1 + arasa + alt_nom_1))
            
            mom_hc_1 = (-1 * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
            mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
            
            mom_hc_se_sm_el_1 = (-1 * fuerza_viento_hc_sm * (alt_nom_1 + arasa + el_hc_1)) + (masa_hc_2 * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc_1))
            mom_hc_se_sm_el_2 = (fuerza_viento_hc_sm * (alt_nom_1 + arasa + el_hc_1)) + (masa_hc_2 * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc_1))
            
            mom_feed_pos_1 = (-1 * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
            mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
            
            mom_cdpa_1 = (-1 * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
            mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
            
            mom_equip = 2 * masa_equip * dist_horiz_equip
            
            mom_poste_1 = -1 * fuerza_viento_poste * (alt_nenc_poste / 2)
            mom_poste_2 = fuerza_viento_poste * (alt_nenc_poste / 2)
            
            mom_tot_1 = (mom_sust_1 + mom_sust_se_sm_el_1 + mom_hc_1 + mom_hc_se_sm_el_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_poste_1)
            mom_tot_2 = (mom_sust_2 + mom_sust_se_sm_el_2 + mom_hc_2 + mom_hc_se_sm_el_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_poste_2)
            
            mom_tot = MAX(mom_tot_1, mom_tot_2)
        End If
'///
'/// Cálculo del momento en el eje de seccionamiento mecánico o seccionamiento de lámina de aire
'///
    ElseIf (tip_1 = eje_sm Or tip_1 = eje_sla) Then
    
            masa_sust = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_1) - (alt_nom_0 + alt_cat_0)) / va_ini) + (((alt_nom_1 + alt_cat_1) - (alt_nom_2 + alt_cat_2)) / va_fin)))
            masa_hc = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - alt_nom_0) / va_ini) + ((alt_nom_1 - (alt_nom_2 + el_hc_2)) / va_fin)))
            
            fuerza_at_ejey_sust = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_1 - d_0) / (va_ini) + (d_1 - d_2) / (va_fin))
            fuerza_at_ejey_hc = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_1 - d_0) / (va_ini) + (d_1 - d_2) / (va_fin))
            
            masa_sust_2 = (n_sust * p_sust1 * (va_ini + va_fin) / 2) + (n_sust * t_sust * ((((alt_nom_1 + alt_cat_4) - (alt_nom_0 + alt_cat_3)) / va_ini) + (((alt_nom_1 + alt_cat_4) - (alt_nom_2 + alt_cat_5)) / va_fin)))
            masa_hc_2 = (n_hc * p_hc1 * (va_ini + va_fin) / 2) + (n_hc * t_hc * (((alt_nom_1 - (alt_nom_0 + el_hc_3)) / va_ini) + ((alt_nom_1 - alt_nom_2) / va_fin)))
                                    
            fuerza_at_ejey_sust_sm = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_4 - d_3) / (va_ini) + (d_4 - d_5) / (va_fin))
            fuerza_at_ejey_hc_sm = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_4 - d_3) / (va_ini) + (d_4 - d_5) / (va_fin))
            
            fuerza_viento_sust_sm = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_ini + va_fin) / 2
            fuerza_viento_hc_sm = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_ini + va_fin) / 2
            
            mom_sust_1 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
            mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
            
            mom_hc_1 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
            mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
            
            mom_sust_se_sm_el_1 = (ayuda * fuerza_viento_sust_sm * (alt_nom_1 + alt_cat_4 + arasa)) + (masa_sust_2 * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_sust_sm * (alt_cat_4 + arasa + alt_nom_1))
            mom_sust_se_sm_el_2 = (fuerza_viento_sust_sm * (alt_nom_1 + alt_cat_4 + arasa)) + (masa_sust_2 * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_sust_sm * (alt_cat_4 + arasa + alt_nom_1))
                        
            mom_hc_se_sm_el_1 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc_2 * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa))
            mom_hc_se_sm_el_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc_2 * dist_hc_poste_4) + (ayuda * fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa))
            
            mom_feed_pos_1 = (ayuda * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
            mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
            
            mom_cdpa_1 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
            mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
            
            mom_equip = 2 * masa_equip * dist_horiz_equip
            
            mom_poste_1 = ayuda * fuerza_viento_poste * (alt_nenc_poste / 2)
            mom_poste_2 = fuerza_viento_poste * (alt_nenc_poste / 2)
            
            mom_tot_1 = (mom_sust_1 + mom_sust_se_sm_el_1 + mom_hc_1 + mom_hc_se_sm_el_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_poste_1)
            mom_tot_2 = (mom_sust_2 + mom_sust_se_sm_el_2 + mom_hc_2 + mom_hc_se_sm_el_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_poste_2)
            
            mom_tot = MAX(mom_tot_1, mom_tot_2)
          
    
'//
'//MOMENTO POSTE PUNTO FIJO
'//
    
    ElseIf tip_pf_1 = anc_pf Or tip_1 = anc_pf Then
        'dist_carril_poste_pos = Sheets("Replanteo").Cells(fila + 2, 5).Value
        If Sheets("Replanteo").Cells(fila + 2, 16).Value = eje_pf Or Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 1, 14) = eje_pf Then
            
            fuerza_ejey_pf_anc = t_pto_fijo * sin(Atn((dist_hc_poste_2 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin))
            
            fuerza_viento_pf = fuerza_viento_cyc(diam_pto_fijo, sec_pto_fijo, n_pto_fijo, "p_fijo", adm_lin_poste) * (va_fin) * Cos(Atn((dist_hc_poste_2 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin))
            
            masa_pf = (p_pto_fijo1 * (va_fin) / 2) + (t_pto_fijo * ((dist_vert_pf_anc1 - (alt_nom_2)) / va_fin))
        Else

            fuerza_ejey_pf_anc = t_pto_fijo * sin(Atn((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini))
            
            fuerza_viento_pf = fuerza_viento_cyc(diam_pto_fijo, sec_pto_fijo, 1, "p_fijo", adm_lin_poste) * (va_ini) * Cos(Atn((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini))
            
            masa_pf = (p_pto_fijo1 * (va_ini) / 2) + (t_pto_fijo * ((dist_vert_pf_anc1 - (alt_nom_0)) / va_ini))
        End If
        
        mom_sust_1 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        
        mom_hc_1 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
                        
        mom_feed_pos_1 = (ayuda * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        
        mom_feed_neg = (fuerza_viento_feed_neg * dist_vert_feed_neg1) - (masa_feed_neg * dist_horiz_feed_neg) + (fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
                
        mom_cdpa_1 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        
        mom_equip = (masa_equip * dist_horiz_equip)
        
        mom_poste_1 = (ayuda * fuerza_viento_poste * (alt_nenc_poste / 2))
        mom_poste_2 = (fuerza_viento_poste * (alt_nenc_poste / 2))
        
        mom_pf_anc_1 = (0.5 * ayuda * fuerza_viento_pf * (dist_vert_pf_anc1 + arasa)) + (fuerza_ejey_pf_anc * (dist_vert_pf_anc1 + arasa)) + (masa_pf * 0.2)
        mom_pf_anc_2 = 0.5 * (fuerza_viento_pf * (dist_vert_pf_anc1 + arasa)) + (fuerza_ejey_pf_anc * (dist_vert_pf_anc1 + arasa)) + (masa_pf * 0.2)
                   
        mom_tot_1 = (mom_sust_1 + mom_hc_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_pf_anc_1 + mom_poste_1)
        mom_tot_2 = (mom_sust_2 + mom_hc_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_pf_anc_2 + mom_poste_2)
        
        mom_tot = MAX(mom_tot_1, mom_tot_2)
         
    ElseIf tip_1 = eje_pf Or tip_pf_1 = eje_pf Then
        
        fuerza_ejey_pf_anc = t_pto_fijo * sin(Atn(((dist_hc_poste_2 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin))) + t_pto_fijo * sin(Atn(((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini)))
        
        fuerza_viento_pf = Cos(Atn(((dist_hc_poste_0 + (dist_carril_poste_1 - dist_carril_poste_0)) / va_ini))) * fuerza_viento_cyc(diam_pto_fijo, sec_pto_fijo, n_pto_fijo, "p_fijo", adm_lin_poste) * (va_fin) + Cos(Atn(((dist_hc_poste_2 + (dist_carril_poste_1 - dist_carril_poste_2)) / va_fin))) * fuerza_viento_cyc(diam_pto_fijo, sec_pto_fijo, 1, "p_fijo", adm_lin_poste) * (va_ini)
        
        masa_pf = (p_pto_fijo1 * (va_fin) / 2) + (t_pto_fijo * ((dist_vert_pf_anc1 - (alt_nom_2)) / va_fin)) + (p_pto_fijo1 * (va_ini) / 2) + (t_pto_fijo * ((dist_vert_pf_anc1 - (alt_nom_0)) / va_ini))
        
        mom_sust_1 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        
        mom_hc_1 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste_1) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
                        
        mom_feed_pos_1 = (ayuda * fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_pos_2 = (fuerza_viento_feed_pos * dist_vert_feed_pos1) - (masa_feed_pos * dist_horiz_feed_pos1) + (ayuda * fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        
        mom_feed_neg = (fuerza_viento_feed_neg * dist_vert_feed_neg1) - (masa_feed_neg * dist_horiz_feed_neg) + (fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
                
        mom_cdpa_1 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        
        mom_equip = (masa_equip * dist_horiz_equip)
        
        mom_poste_1 = (ayuda * fuerza_viento_poste * (alt_nenc_poste / 2))
        mom_poste_2 = (fuerza_viento_poste * (alt_nenc_poste / 2))
        
        mom_pf_anc_1 = (0.5 * ayuda * fuerza_viento_pf * (dist_vert_pf_anc1 + arasa)) + (fuerza_ejey_pf_anc * (dist_vert_pf_anc1 + arasa)) + (masa_pf * 0.2)
        mom_pf_anc_2 = (0.5 * fuerza_viento_pf * (dist_vert_pf_anc1 + arasa)) + (fuerza_ejey_pf_anc * (dist_vert_pf_anc1 + arasa)) + (masa_pf * 0.2)
        
        mom_tot_1 = (mom_sust_1 + mom_hc_1 + mom_feed_pos_1 + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_pf_anc_1 + mom_poste_1)
        mom_tot_2 = (mom_sust_2 + mom_hc_2 + mom_feed_pos_2 + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_pf_anc_2 + mom_poste_2)
        
        mom_tot = MAX(mom_tot_1, mom_tot_2)
        
'//
'//MOMENTO POSTE AGUJA
'//
    
    ElseIf tip_1 = anc_aguj Or tip_pf_1 = anc_aguj Then
           
           
        If Sheets("Replanteo").Cells(fila + 2, 16).Value = semi_eje_aguj Or Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 15) = semi_eje_aguj Then
            fuerza_ejey_sust_anc = t_sust * ((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_4 + ancho_medio_poste) - dist_carril_poste + dist_carril_poste_pos) / va_fin)
            fuerza_ejey_hc_anc = n_hc * t_hc * ((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_4 + ancho_medio_poste) - dist_carril_poste + dist_carril_poste_pos) / va_fin)
            fuerza_viento_sust_sm = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_fin)
            fuerza_viento_hc_sm = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_fin)
        ElseIf Sheets("Replanteo").Cells(fila - 2, 16).Value = semi_eje_aguj Then
        
            fuerza_at_ejey_sust = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_1 - d_5) / (va_ini) + (d_1 - d_2) / (va_fin))
            fuerza_at_ejey_hc = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_1 - d_5) / (va_ini) + (d_1 - d_2) / (va_fin))
            fuerza_at_ejey_cdpa = n_cdpa * t_cdpa * ((va_ini + va_fin) / (2 * r) + (d_1 - d_5) / (va_ini) + (d_1 - d_2) / (va_fin))
            fuerza_viento_sust_sm = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_fin)
            fuerza_viento_hc_sm = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_fin)
        End If
        mom_sust_1 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_hc_1 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        mom_cdpa_1 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc_1 = 0.5 * (fuerza_viento_sust_sm * dist_vert_sust_anc) + (fuerza_ejey_sust_anc * dist_vert_sust_anc)
        mom_sust_anc_2 = 0.5 * (ayuda * fuerza_viento_sust_sm * dist_vert_sust_anc) + (fuerza_ejey_sust_anc * dist_vert_sust_anc)
        mom_hc_anc_1 = 0.5 * (fuerza_viento_hc_sm * dist_vert_hc_anc) + (fuerza_ejey_hc_anc * dist_vert_hc_anc)
        mom_hc_anc_2 = 0.5 * (ayuda * fuerza_viento_hc_sm * dist_vert_hc_anc) + (fuerza_ejey_hc_anc * dist_vert_hc_anc)
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)

        
        mom_tot_1 = (mom_sust_1 + mom_hc_1 + mom_feed_pos + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_sust_anc_1 + mom_hc_anc_1 + mom_poste)
        mom_tot_2 = (mom_sust_2 + mom_hc_2 + mom_feed_pos + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_sust_anc_2 + mom_hc_anc_2 + mom_poste)

        mom_tot = MAX(mom_tot_1, mom_tot_2)
    
    ElseIf tip_1 = semi_eje_aguj Or tip_pf_1 = semi_eje_aguj Then
    
        'mom_sust = fuerza_viento_sust * (alt_nom + arasa + alt_cat) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + arasa + alt_cat)
        'mom_sust_se_ag_el = fuerza_viento_sust * (alt_cat_se_ag_el + arasa + alt_nom) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_sust * (alt_cat_se_ag_el + arasa + alt_nom)
        'mom_hc = fuerza_viento_hc * (alt_nom + arasa + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + arasa + el_hc)
        'mom_hc_se_ag_el = fuerza_viento_hc * (alt_nom + arasa + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_nom + arasa + el_hc)
        'mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 - masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        'mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        'mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa1 + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa1
        'mom_equip = -2 * masa_equip * dist_horiz_equip
        'mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        'mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        'mom_tot = mom_sust + mom_sust_se_ag_el + mom_hc + mom_hc_se_ag_el + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
     dist_carril_poste_pos = Sheets("Replanteo").Cells(fila + 2, 5).Value
        If tip_2 = eje_aguj Or tip_pf_2 = eje_aguj Then
            fuerza_at_ejey_sust_sm = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_3 - d_4) / (va_fin))
            fuerza_at_ejey_hc_sm = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_3 - d_4) / (va_fin))
            fuerza_viento_sust_sm = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_fin)
            fuerza_viento_hc_sm = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_fin)
            el_hc = Sheets("Replanteo").Cells(fila, 46).Value
        ElseIf tip_2 = anc_aguj Or tip_pf_2 = anc_aguj Then
            'fuerza_ejey_sust_anc = t_sust * sin(Atn((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_2 + ancho_medio_poste)) / va_fin))
            'fuerza_ejey_hc_anc = n_hc * t_hc * sin(Atn((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_2 + ancho_medio_poste)) / va_fin))
            fuerza_ejey_sust_anc = t_sust * ((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_2 + ancho_medio_poste) - dist_carril_poste + dist_carril_poste_pos) / va_fin)
            fuerza_ejey_hc_anc = n_hc * t_hc * ((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_2 + ancho_medio_poste) - dist_carril_poste + dist_carril_poste_pos) / va_fin)
            fuerza_at_ejey_sust_sm = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_1 - d_5) / (va_fin))
            fuerza_at_ejey_hc_sm = n_hc * t_hc * ((va_ini + va_fin) / (2 * r) + (d_1 - d_5) / (va_fin))
            fuerza_viento_sust_sm = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_fin)
            fuerza_viento_hc_sm = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_fin)
            el_hc = Sheets("Replanteo").Cells(fila, 46).Value
        Else
            algo = 0
        
        End If
        mom_sust_1 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * dist_hc_poste) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_se_sm_el_1 = (fuerza_viento_sust * (alt_cat_4 + arasa + alt_nom_1)) + (masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste + dist_elect_sm)) + (fuerza_at_ejey_sust_sm * (alt_cat_3 + arasa + alt_nom_1))
        mom_sust_se_sm_el_2 = (ayuda * fuerza_viento_sust * (alt_cat_se_sm_el + arasa + alt_nom_1)) + (masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste + dist_elect_sm)) + (fuerza_at_ejey_sust_sm * (alt_cat_se_sm_el + arasa + alt_nom_1))
        mom_hc_1 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * dist_hc_poste) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_se_sm_el_1 = (fuerza_viento_hc * (alt_nom_1 + arasa + el_hc)) + (masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste + dist_elect_sm)) + (fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc))
        mom_hc_se_sm_el_2 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa + el_hc)) + (masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste + dist_elect_sm)) + (fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc))
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 - masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        mom_cdpa_1 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_equip = 2 * masa_equip * dist_horiz_equip
        mom_sust_anc_1 = 0.5 * (fuerza_viento_sust_sm * dist_vert_sust_anc) - (fuerza_ejey_sust_anc * dist_vert_sust_anc)
        mom_sust_anc_2 = 0.5 * (ayuda * fuerza_viento_sust_sm * dist_vert_sust_anc) - (fuerza_ejey_sust_anc * dist_vert_sust_anc)
        mom_hc_anc_1 = 0.5 * (fuerza_viento_hc_sm * dist_vert_hc_anc) - (fuerza_ejey_hc_anc * dist_vert_hc_anc)
        mom_hc_anc_2 = 0.5 * (ayuda * fuerza_viento_hc_sm * dist_vert_hc_anc) - (fuerza_ejey_hc_anc * dist_vert_hc_anc)
        mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot_1 = (mom_sust_1 + mom_sust_se_sm_el_1 + mom_hc_1 + mom_hc_se_sm_el_1 + mom_feed_pos + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_sust_anc_1 + mom_hc_anc_1 + mom_poste)
        mom_tot_2 = (mom_sust_2 + mom_sust_se_sm_el_2 + mom_hc_2 + mom_hc_se_sm_el_2 + mom_feed_pos + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_sust_anc_2 + mom_hc_anc_2 + mom_poste)
        
        mom_tot = MAX(mom_tot_1, mom_tot_2)
    
    
    ElseIf tip_1 = eje_aguj Or tip_pf_1 = eje_aguj Then
    
        'mom_sust = fuerza_viento_sust * (alt_nom + arasa + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + arasa + alt_cat)
        'mom_sust_e_ag = fuerza_viento_sust * (alt_cat_e_ag + arasa + alt_nom) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_sust * (alt_cat_e_ag + arasa + alt_nom)
        'mom_hc = fuerza_viento_hc * (alt_nom + arasa + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + arasa + el_hc)
        'mom_hc_e_ag = fuerza_viento_hc * (alt_nom + arasa + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_nom + arasa + el_hc)
        'mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        'mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        'mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa1 - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa1
        'mom_equip = 2 * masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        'mom_tot = mom_sust + mom_sust_e_ag + mom_hc + mom_hc_e_ag + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste
        el_hc = Sheets("Replanteo").Cells(fila, 46).Value

        fuerza_at_ejey_sust_sm = n_sust * t_sust * (((va_ini + va_fin) / (2 * r)) + (d_3 - d_5) / (va_ini))
        fuerza_at_ejey_hc_sm = n_hc * t_hc * (((va_ini + va_fin) / (2 * r)) + (d_3 - d_5) / (va_ini))
        
        
        mom_sust_1 = (fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste)) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_2 = (ayuda * fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1)) + (masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste)) + (ayuda * fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1))
        mom_sust_e_sm_1 = (fuerza_viento_sust * (alt_cat_se_sm_el + arasa + alt_nom_1)) + (masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste)) + (ayuda * fuerza_at_ejey_sust_sm * (alt_cat_se_sm_el + arasa + alt_nom_1))
        mom_sust_e_sm_2 = (ayuda * fuerza_viento_sust * (alt_cat_se_sm_el + arasa + alt_nom_1)) + (masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste)) + (ayuda * fuerza_at_ejey_sust_sm * (alt_cat_se_sm_el + arasa + alt_nom_1))
        mom_hc_1 = (fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste)) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_2 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa)) + (masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste)) + (ayuda * fuerza_at_ejey_hc * (alt_nom_1 + arasa))
        mom_hc_e_sm_1 = (fuerza_viento_hc * (alt_nom_1 + arasa + el_hc)) + (masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste + dist_elect_sm)) + (ayuda * fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc))
        mom_hc_e_sm_2 = (ayuda * fuerza_viento_hc * (alt_nom_1 + arasa + el_hc)) + (masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_3 + ancho_medio_poste + dist_elect_sm)) + (ayuda * fuerza_at_ejey_hc_sm * (alt_nom_1 + arasa + el_hc))
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        mom_cdpa_1 = (fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_cdpa_2 = (ayuda * fuerza_viento_cdpa * dist_vert_cdpa1) - (masa_cdpa * dist_horiz_cdpa) + (ayuda * fuerza_at_ejey_cdpa * dist_vert_cdpa1)
        mom_equip = 2 * masa_equip * dist_horiz_equip
        
        mom_tot_1 = (mom_sust_1 + mom_sust_e_sm_1 + mom_hc_1 + mom_hc_e_sm_1 + mom_feed_pos + mom_feed_neg + mom_cdpa_1 + mom_equip + mom_poste)
        mom_tot_2 = (mom_sust_2 + mom_sust_e_sm_2 + mom_hc_2 + mom_hc_e_sm_2 + mom_feed_pos + mom_feed_neg + mom_cdpa_2 + mom_equip + mom_poste)
    
        mom_tot = MAX(mom_tot_1, mom_tot_2)
    
'//
'//MOMENTO POSTE ZONA NEUTRA
'//
    
    ElseIf tip = anc_neutra Then
    
        mom_sust = fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1)
        mom_hc = fuerza_viento_hc * (alt_nom_1 + arasa + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom_1 + arasa + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa1 - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa1
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
    
    ElseIf tip = semi_eje_neutra Then
    
        mom_sust = fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1)
        mom_sust_se_zn_el = fuerza_viento_sust * (alt_cat_se_zn_el + arasa + alt_nom_1) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_sust * (alt_cat_se_zn_el + arasa + alt_nom_1)
        mom_hc = fuerza_viento_hc * (alt_nom_1 + arasa + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom_1 + arasa + el_hc)
        mom_hc_se_zn_el = fuerza_viento_hc * (alt_nom_1 + arasa + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_hc * (alt_nom_1 + arasa + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 - masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa1 + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa1
        mom_equip = -2 * masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_se_zn_el + mom_hc + mom_hc_se_zn_el + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    

    ElseIf tip = eje_neutra Then
    
        mom_sust = fuerza_viento_sust * (alt_nom_1 + arasa + alt_cat_1) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom_1 + arasa + alt_cat_1)
        mom_sust_e_zn = fuerza_viento_sust * (alt_cat_e_zn + arasa + alt_nom_1) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_sust * (alt_cat_e_zn + arasa + alt_nom_1)
        mom_hc = fuerza_viento_hc * (alt_nom_1 + arasa + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom_1 + arasa + el_hc)
        mom_hc_e_zn = fuerza_viento_hc * (alt_nom_1 + arasa + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_hc * (alt_nom_1 + arasa + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos1 + masa_feed_pos * dist_horiz_feed_pos1 + fuerza_at_ejey_feed_pos * dist_vert_feed_pos1)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg1 + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg1)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa1 - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa1
        mom_equip = 2 * masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_e_zn + mom_hc + mom_hc_e_zn + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste

    End If

'//
'//INSERCIÓN MOMENTO EN REPLANTEO
'//

    Sheets("Replanteo").Cells(fila, 19) = mom_tot
End If

If CAD = True Then
    GoTo fin
Else
    fila = fila + 2
End If
Wend
fin:
End Sub

Public Function fuerza_viento_cyc(ByRef diam, ByRef sec, ByRef n, ByRef tip, ByRef adm_lin) As Double

    If adm_lin_poste = "ADIF" Then
        
        If sec <= 107 Then
            coef_viento = 1.2
        End If
        
        If sec > 107 And sec < 150 Then
            coef_viento = 1.1
        End If
        
        If sec >= 150 Then
            coef_viento = 1
        End If
        
        If n > 1 And tip = "hc" Then
            If sep_hc > (6 * diam) Then
                coef_viento = 2 * coef_viento
            Else
                coef_viento = 1.6 * coef_viento
            End If
        End If
        
        If n > 1 And (tip = "hc" Or ((tip = "feed_pos" Or tip = "feed_neg") And al = "C.Continua")) Then
            If sep_hc > (6 * diam) Then
                coef_viento = 2 * coef_viento
            Else
                coef_viento = 1.6 * coef_viento
            End If
        End If
      
        fuerza_viento_cyc = 0.5 * 1.25 * (vw ^ 2) * coef_viento * diam / 9.81
        'kg/m
      
      ElseIf adm_lin_poste = "ONCF" Then
        
        coef_viento = 1.2
               
        fuerza_viento_cyc = coef_viento * 68.1663 * diam
        
        'daN/m
      End If

End Function

Public Function fuerza_viento_sup(ByRef adm_lin) As Double

    If adm_lin_poste = "ADIF" Then
        
        fuerza_viento_sup = 100
        
    ElseIf adm_lin_poste = "ONCF" Then
        
        fuerza_viento_sup = 87.93
    
    End If
        
        'daN/m2
End Function


Public Function ASENO(X As Single) As Single
Dim Resul As Single
Resul = CSng(Atn(X / Sqr(-X * X + 1)))
ASENO = Resul
End Function

Public Function ACOS(X As Single) As Single
Dim Resul As Single
Resul = CSng(Atn(Sqr(-X * X + 1) / X))
ACOS = Resul
End Function

Public Function MAX(X As Double, Y As Double) As Double
Dim Resul As Double
If Abs(X) > Abs(Y) Then
Resul = X
ElseIf Abs(X) <= Abs(Y) Then
Resul = Y
End If
MAX = Resul
End Function

