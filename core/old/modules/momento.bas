Attribute VB_Name = "momento"
' queda la duda de si el momento de atirantado en recta se debe sumar en los momentos de curva

Sub momento(nombre_catVB)
Dim coef_viento As Double
Dim va_ini As Double, va_fin As Double
Dim r As Double, d_0 As Double, d_1 As Double, d_2 As Double
Dim tip As String
Dim fuerza_at_re_ejey_sust As Double, fuerza_at_cu_ejey_sust As Double, fuerza_at_ejey_sust As Double
Dim fuerza_at_re_ejex_sust As Double, fuerza_at_cu_ejex_sust As Double, fuerza_at_ejex_sust As Double
Dim fuerza_at_re_ejey_hc As Double
Dim fuerza_at_cu_ejey_hc As Double
Dim fuerza_at_ejey_hc As Double
Dim fuerza_at_re_ejex_hc As Double
Dim fuerza_at_cu_ejex_hc As Double
Dim fuerza_at_ejex_hc As Double
Dim fuerza_at_ejey_feed_pos As Double
Dim fuerza_at_cu_ejey_feed_pos As Double
Dim fuerza_at_ejex_feed_pos As Double
Dim fuerza_at_cu_ejex_feed_pos As Double
Dim fuerza_at_ejey_feed_neg As Double
Dim fuerza_at_cu_ejey_feed_neg As Double
Dim fuerza_at_ejex_feed_neg As Double
Dim fuerza_at_cu_ejex_feed_neg As Double
Dim fuerza_at_ejey_cdpa As Double
Dim fuerza_at_cu_ejey_cdpa As Double
Dim fuerza_at_ejex_cdpa As Double
Dim fuerza_at_cu_ejex_cdpa As Double
Dim mom_sust As Double
Dim mom_sust_1 As Double
Dim mom_hc As Double
Dim mom_hc_1 As Double
Dim mom_feed_pos As Double
Dim mom_feed_neg As Double
Dim mom_cdpa As Double
Dim mom_equip As Double
Dim mom_sust_anc As Double
Dim mom_hc_anc As Double
Dim pres_viento_sust As Double
Dim pres_viento_hc As Double
Dim pres_viento_feed_pos As Double
Dim pres_viento_feed_neg As Double
Dim pres_viento_pto_fijo As Double
Dim pres_viento_cdpa As Double
Dim pres_viento_pend As Double
Dim fuerza_viento_sust As Double
Dim fuerza_viento_hc As Double
Dim fuerza_viento_feed_pos As Double
Dim fuerza_viento_feed_neg As Double
Dim fuerza_viento_pto_fijo As Double
Dim fuerza_viento_cdpa As Double
Dim fuerza_viento_pend As Double
Dim masa_sust As Double
Dim masa_hc As Double
Dim masa_feed_pos As Double
Dim masa_feed_neg As Double
Dim masa_pto_fijo As Double
Dim masa_cdpa As Double
Dim masa_pend As Double
Dim masa_equip As Double

'//
'//LECTURA BASE DE DATOS
'//
Call cargar.datos_acces(nombre_catVB)

ancho_medio_poste = 0.2 'valor est?ndar supuesto

z = 10
   
'//
'//LECTURA DATOS REPLANTEO
'//

While Not IsEmpty(Sheets(1).Cells(z + 2, 33).Value)
If Sheets(1).Cells(z, 25).Value <> "Tunnel" Then
    alt_nom = Hoja1.Cells(z, 10)
    dist_carril_poste = Hoja1.Cells(z, 5)
    
    va_fin = Hoja1.Cells(z + 1, 4).Value
    r = Hoja1.Cells(z, 6).Value
    If r = 0 Then
        r = 100000000
    ElseIf r < 0 Then
        r = r * (-1)
    End If
    d_1 = Hoja1.Cells(z, 8).Value
    d_2 = Hoja1.Cells(z + 2, 8).Value
    tip = Hoja1.Cells(z, 16).Value
    If z <> 10 Then
        va_ini = Hoja1.Cells(z - 1, 4).Value
        d_0 = Hoja1.Cells(z - 2, 8).Value
    Else
    va_ini = va_fin
    d_0 = d_1
    End If
    
    If d_1 > 0 Then
        dist_elect_sm = -(dist_elect_sm)
        dist_elect_sla = -(dist_elect_sla)
        masa_equip = p_medio_equip_comp
        dist_horiz_equip = dist_horiz_equip_comp
    Else
        masa_equip = p_medio_equip_t
        dist_horiz_equip = dist_horiz_equip_t
    End If

'//
'//C?LCULO DISTANCIAS
'//

    dist_vert_feed_pos = dist_vert_feed_pos + dist_base_poste_pmr
    dist_vert_feed_neg = dist_vert_feed_neg + dist_base_poste_pmr
    dist_vert_cdpa = dist_vert_cdpa + dist_base_poste_pmr
    dist_horiz_equip = dist_horiz_equip
    dist_vert_hc_anc = dist_vert_hc_anc + dist_base_poste_pmr
    dist_vert_sust_anc = dist_vert_sust_anc + dist_base_poste_pmr
    
    var_0 = 1
    var_1 = 1
    n_sust = 1
    If posicion_feed_pos = "apoyado" Then
        dist_horiz_feed_pos = 0
    End If
    If posicion_feed_pos = "Suspendido (lado exterior)" Then
        dist_horiz_feed_pos = -(dist_horiz_feed_pos)
    End If
    If posicion_feed_pos = "Suspendido (lado v?a)" Then
        dist_horiz_feed_pos = (dist_horiz_feed_pos)
    End If
    If posicion_feed_pos = "NO HAY" Then
        var_0 = 0
    End If
    
    If posicion_feed_neg = "apoyado" Then
        dist_horiz_feed_neg = 0
    End If
    If posicion_feed_neg = "Suspendido (lado exterior)" Then
        dist_horiz_feed_neg = -(dist_horiz_feed_neg)
    End If
    If posicion_feed_neg = "Suspendido (lado v?a)" Then
        dist_horiz_feed_neg = (dist_horiz_feed_neg)
    End If
    If posicion_feed_neg = "NO HAY" Then
        var_1 = 0
    End If


'//
'//C?LCULO FUERZA VIENTO
'//

    fuerza_viento_sust = fuerza_viento_cyc(diam_sust, sec_sust, n_sust, "sust", adm_lin_poste) * (va_ini + va_fin) / 2
    fuerza_viento_hc = fuerza_viento_cyc(diam_hc, sec_hc, n_hc, "hc", adm_lin_poste) * (va_ini + va_fin) / 2
    fuerza_viento_feed_pos = fuerza_viento_cyc(diam_feed_pos, sec_feed_pos, n_feed_pos, "feed_pos", adm_lin_poste) * (va_ini + va_fin) / 2
    fuerza_viento_feed_neg = fuerza_viento_cyc(diam_feed_neg, sec_feed_neg, n_feed_neg, "feed_neg", adm_lin_poste) * (va_ini + va_fin) / 2
    fuerza_viento_pto_fijo = fuerza_viento_cyc(diam_pto_fijo, sec_pto_fijo, n_pto_fijo, "pto_fijo", adm_lin_poste) * (va_ini + va_fin) / 2
    fuerza_viento_cdpa = fuerza_viento_cyc(diam_cdpa, sec_cdpa, n_cdpa, "cdpa", adm_lin_poste) * (va_ini + va_fin) / 2
    fuerza_viento_pend = fuerza_viento_cyc(diam_pend, sec_pend, n_pend, "pend", adm_lin_poste) * (va_ini + va_fin) / 2 * 0.15 'multiplicar por un porcentaje a concretar
    fuerza_viento_poste = fuerza_viento_sup(adm_lin_poste) * sup_perf_max_poste
    
'//
'//C?LCULO PESO CONDUCTORES
'//
    
    masa_sust = n_sust * p_sust * (va_ini + va_fin) / 2
    masa_hc = n_hc * p_hc * (va_ini + va_fin) / 2
    masa_feed_pos = n_feed_pos * p_feed_pos * (va_ini + va_fin) / 2
    masa_feed_neg = n_feed_neg * p_feed_neg * (va_ini + va_fin) / 2
    masa_pto_fijo = p_pto_fijo * (va_ini + va_fin) / 2
    masa_cdpa = n_cdpa * p_cdpa * (va_ini + va_fin) / 2
    masa_pend = p_pend * (va_ini + va_fin) / 2 'multiplicar por un porcentaje a concretar
    
'//
'//C?LCULO FUERZA RADIAL EN RECTA
'//

    fuerza_at_re_ejey_sust = n_sust * t_sust * ((Abs(d_1) + Abs(d_0)) / (va_ini) + (Abs(d_1) + Abs(d_2)) / (va_fin))
    fuerza_at_re_ejex_sust = n_sust * t_sust * (Sqr(1 - ((Abs(d_1) + Abs(d_2)) / (va_fin)) ^ 2) - Sqr(1 - ((Abs(d_1) + Abs(d_0)) / (va_ini)) ^ 2))
    fuerza_at_re_ejey_hc = n_hc * t_hc * ((Abs(d_1) + Abs(d_0)) / (va_ini) + (Abs(d_1) + Abs(d_2)) / (va_fin))
    fuerza_at_re_ejex_hc = n_hc * t_hc * (Sqr(1 - ((Abs(d_1) + Abs(d_2)) / (va_fin)) ^ 2) - Sqr(1 - ((Abs(d_1) + Abs(d_0)) / (va_ini)) ^ 2))

'//
'//C?LCULO FUERZA RADIAL EN CURVA
'//

    fuerza_at_cu_ejey_sust = n_sust * t_sust * (((va_ini ^ 2 + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) + ((va_fin ^ 2 + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))))
    fuerza_at_cu_ejex_sust = n_sust * t_sust * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_cu_ejey_hc = n_hc * t_hc * (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1))) + ((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1))))
    fuerza_at_cu_ejex_hc = n_hc * t_hc * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_cu_ejey_feed_pos = n_feed_pos * t_feed_pos * (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1))) + ((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1))))
    fuerza_at_cu_ejex_feed_pos = n_feed_pos * t_feed_pos * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_cu_ejey_feed_neg = n_feed_neg * t_feed_neg * (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1))) + ((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1))))
    fuerza_at_cu_ejex_feed_neg = n_feed_neg * t_feed_neg * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))
    fuerza_at_cu_ejey_cdpa = n_cdpa * t_cdpa * (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1))) + ((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1))))
    fuerza_at_cu_ejex_cdpa = n_cdpa * t_cdpa * (Sqr(1 - (((va_ini ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_0)) ^ 2) / (2 * va_ini * (r + Abs(d_1)))) ^ 2) - Sqr(1 - (((va_fin ^ 2) + (r + Abs(d_1)) ^ 2 - (r + Abs(d_2)) ^ 2) / (2 * va_fin * (r + Abs(d_1)))) ^ 2))

'//
'//COMPARACI?N ESFUERZO RADIAL
'//

    fuerza_at_ejey_sust = MAX(fuerza_at_re_ejey_sust, fuerza_at_cu_ejey_sust)
    fuerza_at_ejex_sust = MAX(fuerza_at_re_ejex_sust, fuerza_at_cu_ejex_sust)
    fuerza_at_ejey_hc = MAX(fuerza_at_re_ejey_hc, fuerza_at_cu_ejey_hc)
    fuerza_at_ejex_hc = MAX(fuerza_at_re_ejex_hc, fuerza_at_cu_ejex_hc)
    fuerza_at_ejey_feed_pos = fuerza_at_cu_ejey_feed_pos
    fuerza_at_ejex_feed_pos = fuerza_at_cu_ejex_feed_pos
    fuerza_at_ejey_feed_neg = fuerza_at_cu_ejey_feed_neg
    fuerza_at_ejex_feed_neg = fuerza_at_cu_ejex_feed_neg
    fuerza_at_ejey_cdpa = fuerza_at_cu_ejey_cdpa
    fuerza_at_ejex_cdpa = fuerza_at_cu_ejex_cdpa
    
'//
'//C?LCULO FUERZA ANCLAJE
'//

If IsEmpty(Sheets(1).Cells(z + 2, 33).Value) Then
    fuerza_ejey_sust_anc = t_sust * Sin(Atn((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + d_1) / va_ini))
    fuerza_ejey_hc_anc = t_hc * Sin(Atn((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + d_1) / va_ini))
Else
    fuerza_ejey_sust_anc = t_sust * Sin(Atn((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + d_1) / va_fin))
    fuerza_ejey_hc_anc = t_hc * Sin(Atn((((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + d_1) / va_fin))
End If
    
'//
'//MOMENTO POSTE SIMPLE
'//

    If tip = "" Then
           
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste
        

'//
'//MOMENTO POSTE SECCIONAMIENTO MEC?NICO
'//

    ElseIf tip = "Anc.Chevau." Or tip = "Anc.Chevau.sans AT" Or tip = "Anc.Section.sans AT" Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    

    
    ElseIf tip = "Inter.Chevau." Then ' elegir que momento es mas desfavorable (fuera o dentro)
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_2 = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_se_sm_el = fuerza_viento_sust * (alt_cat_se_sm_el + dist_base_poste_pmr + alt_nom) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_sust * (alt_cat_se_sm_el + dist_base_poste_pmr + alt_nom)
        mom_sust_se_sm_el_2 = fuerza_viento_sust * (alt_cat_se_sm_el + dist_base_poste_pmr + alt_nom) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_sust * (alt_cat_se_sm_el + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_2 = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_se_sm_el = fuerza_viento_hc * (alt_cat_se_sm_el + alt_nom + dist_base_poste_pmr) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_cat_se_sm_el + alt_nom + dist_base_poste_pmr)
        mom_hc_se_sm_el_2 = fuerza_viento_hc * (alt_cat_se_sm_el + alt_nom + dist_base_poste_pmr) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_cat_se_sm_el + alt_nom + dist_base_poste_pmr)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos - masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_pos_2 = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_feed_neg_2 = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_cdpa_2 = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = -2 * masa_equip * dist_horiz_equip
        mom_equip_2 = 2 * masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_sust_anc_2 = 0.5 * fuerza_viento_sust * dist_vert_sust_anc - fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        mom_hc_anc_2 = 0.5 * fuerza_viento_hc * dist_vert_hc_anc - fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_se_sm_el + mom_hc + mom_hc_se_sm_el + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
        mom_tot_2 = mom_sust_2 + mom_sust_se_sm_el_2 + mom_hc_2 + mom_hc_se_sm_el_2 + mom_feed_pos_2 + mom_feed_neg_2 + mom_cdpa_2 + mom_equip_2 + mom_sust_anc_2 + mom_hc_anc_2 + mom_poste_2
        If mom_tot_2 > mom_tot Then
            mom_tot = mom_tot_2
        End If
    
   
    ElseIf tip = "Axe.Chevau." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_e_sm = fuerza_viento_sust * (alt_cat_e_sm + dist_base_poste_pmr + alt_nom) + masa_sust * dist_horiz_e_sm + fuerza_at_ejey_sust * (alt_cat_e_sm + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_e_sm = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = 2 * masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_e_sm + mom_hc + mom_hc_e_sm + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste
    
    
'//
'//MOMENTO POSTE SECCIONAMIENTO EL?CTRICO
'//
    
    ElseIf tip = "Anc.Section." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
   
    ElseIf tip = "Inter.Section." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_se_sla_el = fuerza_viento_sust * (alt_cat_se_sla_el + dist_base_poste_pmr + alt_nom) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_sust * ((alt_cat_se_sla_el + dist_base_poste_pmr + alt_nom))
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_se_sla_el = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos - masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = -2 * masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_se_sla_el + mom_hc + mom_hc_se_sla_el + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
    
    ElseIf tip = "Axe.Section." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_e_sla = fuerza_viento_sust * (alt_cat_e_sla + dist_base_poste_pmr + alt_nom) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_sust * (alt_cat_e_sla + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_e_sla = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = 2 * masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_e_sla + mom_hc + mom_hc_e_sla + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste
    
    
'//
'//MOMENTO POSTE PUNTO FIJO
'//
    
    ElseIf tip = "Anc.Antich." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
            
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_poste
    
         
    ElseIf tip = "Axe.Antich." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos - masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = -masa_equip * dist_horiz_equip
        mom_sust_anc = fuerza_viento_sust * dist_vert_sust_anc + 2 * fuerza_ejey_sust_anc * dist_vert_sust_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_poste
    
'//
'//MOMENTO POSTE AGUJA
'//
    
    ElseIf tip = "Anc.Aigu" Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
    
    ElseIf tip = "Inter.Aigu" Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_se_ag_el = fuerza_viento_sust * (alt_cat_se_ag_el + dist_base_poste_pmr + alt_nom) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_sust * (alt_cat_se_ag_el + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_se_ag_el = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos - masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = -2 * masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_se_ag_el + mom_hc + mom_hc_se_ag_el + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
    
    ElseIf tip = "Axe.Aigu." Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_e_ag = fuerza_viento_sust * (alt_cat_e_ag + dist_base_poste_pmr + alt_nom) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_sust * (alt_cat_e_ag + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_e_ag = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sm) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = 2 * masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_e_ag + mom_hc + mom_hc_e_ag + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste

    
'//
'//MOMENTO POSTE ZONA NEUTRA
'//
    
    ElseIf tip = "Anc.Neutre" Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_hc + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    
    
    ElseIf tip = "Inter.Neutre" Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_se_zn_el = fuerza_viento_sust * (alt_cat_se_zn_el + dist_base_poste_pmr + alt_nom) - masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_sust * (alt_cat_se_zn_el + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_se_zn_el = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) - masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos - masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg - masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa + masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = -2 * masa_equip * dist_horiz_equip
        mom_sust_anc = 0.5 * fuerza_viento_sust * dist_vert_sust_anc + fuerza_ejey_sust_anc * dist_vert_sust_anc
        mom_hc_anc = 0.5 * fuerza_viento_hc * dist_vert_hc_anc + fuerza_ejey_hc_anc * dist_vert_hc_anc
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_se_zn_el + mom_hc + mom_hc_se_zn_el + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_sust_anc + mom_hc_anc + mom_poste
    

    ElseIf tip = "Axe.Neutre" Then
    
        mom_sust = fuerza_viento_sust * (alt_nom + dist_base_poste_pmr + alt_cat) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_sust * (alt_nom + dist_base_poste_pmr + alt_cat)
        mom_sust_e_zn = fuerza_viento_sust * (alt_cat_e_zn + dist_base_poste_pmr + alt_nom) + masa_sust * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_sust * (alt_cat_e_zn + dist_base_poste_pmr + alt_nom)
        mom_hc = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_hc_e_zn = fuerza_viento_hc * (alt_nom + dist_base_poste_pmr + el_hc) + masa_hc * ((ancho_via / 2) + ancho_carril + dist_carril_poste + d_1 + ancho_medio_poste + dist_elect_sla) + fuerza_at_ejey_hc * (alt_nom + dist_base_poste_pmr + el_hc)
        mom_feed_pos = var_0 * (fuerza_viento_feed_pos * dist_vert_feed_pos + masa_feed_pos * dist_horiz_feed_pos + fuerza_at_ejey_feed_pos * dist_vert_feed_pos)
        mom_feed_neg = var_1 * (fuerza_viento_feed_neg * dist_vert_feed_neg + masa_feed_neg * dist_horiz_feed_neg + fuerza_at_ejey_feed_neg * dist_vert_feed_neg)
        mom_cdpa = fuerza_viento_cdpa * dist_vert_cdpa - masa_cdpa * dist_horiz_cdpa + fuerza_at_ejey_cdpa * dist_vert_cdpa
        mom_equip = 2 * masa_equip * dist_horiz_equip
        'mom_poste = fuerza_viento_poste * (alt_nenc_poste / 2)
        
        mom_tot = mom_sust + mom_sust_e_zn + mom_hc + mom_hc_e_zn + mom_feed_pos + mom_feed_neg + mom_cdpa + mom_equip + mom_poste
    
    End If

'//
'//INSERCI?N MOMENTO EN REPLANTEO
'//

    Hoja1.Cells(z, 19) = mom_tot
    End If
    z = z + 2
    Wend
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
        
        coef_viento = 1
        
        If n > 1 And (tip = "hc" Or ((tip = "feed_pos" Or tip = "feed_neg") And al = "C.Continua")) Then
            If sep_hc > (6 * diam) Then
                coef_viento = 2
            Else
                coef_viento = 2
            End If
        End If
        
        fuerza_viento_cyc = coef_viento * 68.2 * diam
        
        'kg/m
      End If

End Function

Public Function fuerza_viento_sup(ByRef adm_lin) As Double

    If adm_lin_poste = "ADIF" Then
        
        fuerza_viento_sup = 100
        
    ElseIf adm_lin_poste = "ONCF" Then
        
        fuerza_viento_sup = 100
    
    End If
        
        'kg/m2
End Function


Public Function ASENO(x As Single) As Single
Dim Resul As Single
Resul = CSng(Atn(x / Sqr(-x * x + 1)))
ASENO = Resul
End Function

Public Function ACOS(x As Single) As Single
Dim Resul As Single
Resul = CSng(Atn(Sqr(-x * x + 1) / x))
ACOS = Resul
End Function

Public Function MAX(x As Double, y As Double) As Double
Dim Resul As Double
If x > y Then
Resul = x
ElseIf x <= y Then
Resul = y
End If
MAX = Resul
End Function

