Attribute VB_Name = "eleccion"
Sub postes(nombre_catVB, fila, CAD)
    
'//
'//LECTURA BASE DE DATOS
'//
    Call cargar.datos_lac(nombre_catVB)
    Call cargar.datos_poste
    
    i = 0
'///
'///CALCULO MOMENTO POSTE
'///
    'While cim(i, 1) <> ""
        'If IsEmpty(post(i, 16)) Then
            'K = (post(i, 16)/(2*post(i,9))*(0.5-(post(i,9)*post(i,10)/
        'Else
        'i = i + 1
    'Wend
    'i = 0
'//
'//ELECCIÓN POSTE EN FUNCIÓN MOMENTO CALCULADO
'//



    While Not IsEmpty(Sheets("Replanteo").Cells(fila, 33).Value)
        If CAD = False Then
        If Len(Sheets("Replanteo").Cells(fila, 16).Value) >= 19 Then
            tip_1 = Mid(Sheets("Replanteo").Cells(fila, 16).Value, 15)
            tip_pf_1 = Mid(Sheets("Replanteo").Cells(fila, 16).Value, 1, 11)
        Else
            tip_1 = Sheets("Replanteo").Cells(fila, 16).Value
            tip_pf_1 = Sheets("Replanteo").Cells(fila, 16).Value
        End If
        If Len(Sheets("Replanteo").Cells(fila - 2, 16).Value) >= 19 Then
            tip_0 = Mid(Sheets("Replanteo").Cells(fila - 2, 16).Value, 15)
            tip_pf_0 = Mid(Sheets("Replanteo").Cells(fila - 2, 16).Value, 1, 11)
        Else
            tip_0 = Sheets("Replanteo").Cells(fila - 2, 16).Value
            tip_pf_0 = Sheets("Replanteo").Cells(fila - 2, 16).Value
        End If
        If Len(Sheets("Replanteo").Cells(fila + 2, 16).Value) >= 19 Then
            tip_2 = Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 15)
            tip_pf_2 = Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 1, 11)
        Else
            tip_2 = Sheets("Replanteo").Cells(fila + 2, 16).Value
            tip_pf_2 = Sheets("Replanteo").Cells(fila + 2, 16).Value
        End If
        End If
        
        
        
        If Sheets("Replanteo").Cells(fila, 38).Value = "Tunel" Or Sheets("Replanteo").Cells(fila, 38).Value = "Marquesina" Or Sheets("Replanteo").Cells(fila, 38).Value = "Viaducto" Then '/// retocar
            tipo_poste = ""
        Else
        mom_poste_calc = Abs(Sheets("Replanteo").Cells(fila, 19))
        alt_nenc_poste = post(i, 8)
        
        If Not IsEmpty(Sheets("Replanteo").Cells(fila, 39).Value) Then
            alt_cat_poste = MAX(Sheets("Replanteo").Cells(fila, 39).Value, Sheets("Replanteo").Cells(fila, 45).Value)
        Else
            alt_cat_poste = alt_cat
        End If
        var_altura = Sheets("Replanteo").Cells(fila, 10).Value + Sheets("Replanteo").Cells(fila, 20).Value + alt_cat_poste + 0.27 + 0.22 + 0.2
        'If (alt_nenc_poste + 0.7 - 0.256 - Sheets("Replanteo").Cells(fila, 10).Value - Sheets("Replanteo").Cells(fila, 20).Value - alt_cat_poste - 0.27) < 0.15 Then
            'var_altura = 9.75
        'ElseIf tip_1 = semi_eje_sm Or tip_1 = eje_sm Or tip_1 = semi_eje_sla Or tip_1 = eje_sla Or tip_1 = semi_eje_aguj Or tip_1 = eje_aguj Then
            'If (alt_nenc_poste + 0.7 - 0.256 - (Sheets("Replanteo").Cells(fila, 10).Value) - (Sheets("Replanteo").Cells(fila, 20).Value) - alt_cat_poste - 0.27) < 0.15 Then
                'var_altura = 9.75
            'ElseIf (alt_nenc_poste + 0.7 - 0.256 - (Sheets("Replanteo").Cells(fila, 10).Value) - (Sheets("Replanteo").Cells(fila, 20).Value) - alt_cat_poste - 0.27) < 0.15 Then
                'var_altura = 9.75
            'Else
                'var_altura = 8.5
            'End If
        'ElseIf (alt_nenc_poste + 0.2 - 0.256 - Sheets("Replanteo").Cells(fila, 10).Value - Sheets("Replanteo").Cells(fila, 20).Value - alt_cat_poste - 0.27) < 0.15 Then
            'var_altura = 8.5
        
        'ElseIf (alt_nenc_poste - 0.256 - Sheets("Replanteo").Cells(fila, 10).Value - Sheets("Replanteo").Cells(fila, 20).Value - alt_cat_poste - 0.27) < 0.15 Then
            'var_altura = 8
        'Else
            'var_altura = post(i, 8)
        
        'End If
        
        While post(i, 16) < mom_poste_calc Or var_altura > alt_nenc_poste
            i = i + 1
            alt_nenc_poste = post(i, 8)
        Wend
        mom_poste_var = post(i, 16)
        alt_nenc_poste = post(i, 8)
        tip_poste = post(i, 2)
        cod_poste = post(i, 17)
X:
        i = 0
        While post(i, 16) <> 0
            If post(i, 16) < mom_poste_var And post(i, 16) > mom_poste_calc And post(i, 9) <= alt_nenc_poste And post(i, 9) >= var_altura Then
                'If alt_nenc_poste > poste(i, 9) Then
                    mom_poste_var = post(i, 16)
                    alt_nenc_poste = post(i, 9)
                    tip_poste = post(i, 2)
                    cod_poste = post(i, 17)
                GoTo X
                    'Else
                'GoTo X
                'End If
            Else: i = i + 1
        End If
        Wend
        
'//
'//INSERECIÓN POSTE EN REPLANTEO
'//
        Sheets("Replanteo").Cells(fila, 35).Value = mom_poste_var
        Sheets("Replanteo").Cells(fila, 36).Value = alt_nenc_poste
        Sheets("Replanteo").Cells(fila, 18).Value = tip_poste
        Sheets("Replanteo").Cells(fila, 51).Value = cod_poste
        
        If (tip_1 = anc_sm_con Or tip_1 = anc_sla_con Or tip_1 = anc_aguj Or tip_1 = anc_sm_sin Or tip_1 = anc_sla_sin) _
         And mom_poste_var <= 7100 And Mid(Sheets("Replanteo").Cells(fila, 18).Value, 1, 1) <> "Z" Then 'post(i, 16) Then
            algo = Mid(Sheets("Replanteo").Cells(fila, 18).Value, 3)
            Sheets("Replanteo").Cells(fila, 18).Value = "X3" & algo
        End If
    End If
    If CAD = True Then
        GoTo acad
    Else
        fila = fila + 2
        i = 0
    End If

    Wend
acad:
End Sub
Sub cimentaciones(nombre_catVB, idiomaVB, cimenta, z, CAD)
'//
'//LECTURA BASE DE DATOS
'//
    
Call cargar.datos_lac(nombre_catVB)
Call cargar.datos_cim
Call cargar.punto_singular(idiomaVB)
'//
'//CÁLCULO CIMENTACIÓN
'//

Dim desm_terrap_mac As String
Dim res_lat As Double
Dim res_arr As Double
Dim alt_tot_mac As Double
Dim mom_vuelco_pos As Double
Dim mom_vuelco_neg As Double

dist_ad_mac = 4 'metros desde el extremo del macizo hasta el final del terreno (se deberá recoger del perfil)
desn_ad_mac = 0.3 'metros de diferencia entre terrenos (se debe recoger del perfil)
incl_ad_mac = 10 'grados (o pendiente, 19º->1:3, 33º->1:1.5) de inclinación del terreno (se debe recoger del perfil)

Const alt_base_desm = 2
Const alt_base_terrap = 2
Const tg_alpha = 0.005
Const p_esp_horm = 2200
'valores impuestos por ADIF a variar si las condiciones son peores
coef_compres_base_desm = 6 ' valor a leer del excel
coef_compres_base_terrap = 4 'valor a leer del excel
ang_roz = 14 'Valor del SiReCa 'ángulo de rozamiento interno del terreno
p_esp_terr = 1400 'Valor obtenido del SiReCa (kg/m3)
cap_lat = 10000 'kg/m2 capacidad lateral del terreno, obtenido del SiReCa
'no se tiene en cuenta el peso del poste, puesto que lo que haría es incrementar el mom_base (mom_vuelco), la elección de la cimentación se hace partiendo del momento del poste calculado, por lo que la cimentación que se escojerá será la que soporte ese momento, si aplicasemos el peso del poste, el mom_vuelco incrementaría.

i = 0
'z = 10
'desm_terrap_mac = cimenta
While cim(i, 1) <> ""

mac_mac = cim(i, 0)

desm_terrap_mac = cim(i, 2)

    If mac_mac = "Paralelepípedo" Then
                     
            l_ent_mac = cim(i, 4) 'ancho cimentación en sentido perpendicular a la vía
            ancho_ent_mac = cim(i, 5) 'ancho cimentación en sentido paralelo a la vía
            l_tot_mac = cim(i, 6) 'ancho cimentación total en sentido perpendicular a la vía
            alt_ent_mac = cim(i, 7) ' altura enterrada del macizo
            alt_nent_mac = cim(i, 11) 'altura no enterrada del macizo
            v_tot_mac = cim(i, 13) ' volumen total del macizo
            diam_mac = cim(i, 14) 'ancho cimentación en sentido perpendicular a la vía
            alt_tot_mac = alt_ent_mac + alt_nent_mac 'altura total del macizo
            p_tot_mac = l_ent_mac * ancho_ent_mac * alt_tot_mac * p_esp_horm 'peso total macizo
                
                If desm_terrap_mac = "desmonte" Then
                        If alt_ent_mac <= 2 Then
                            coef_compres_var = coef_compres_base_desm * (alt_ent_mac / alt_base_desm)
                        Else: coef_compres_var = coef_compres_base_desm
                        End If
                    mom_lat = (1000000 / 36) * tg_alpha * ancho_ent_mac * coef_compres_var * (alt_ent_mac ^ 3)
                    mom_base = p_tot_mac * ((l_ent_mac / 2) - (1 / 3000) * Sqr((2 * p_tot_mac) / (ancho_ent_mac * coef_compres_var * tg_alpha)))
                    coef_pond = mom_lat / mom_base
                        If coef_pond <= 1 Then
                            coef_pond = 0.4167 * ((mom_lat / mom_base) ^ 2) - 0.9167 * (mom_lat / mom_base) + 1.5
                        Else: coef_pond = 1
                        End If
                    mom_vuelco = (mom_lat + mom_base) / coef_pond
                    If cim(i, 15) = "" Then
                        cim(i, 15) = mom_vuelco
                    End If
                
                ElseIf desm_terrap_mac = "terraplén" Then
                    p_tot_mac_2 = 0.5 * (l_tot_mac - l_ent_mac) * ancho_ent_mac * alt_tot_mac * p_esp_horm
                        If alt_ent_mac <= 2 Then
                            coef_compres_var = coef_compres_base_terrap * (alt_ent_mac / alt_base_terrap)
                        Else: coef_compres_var = coef_compres_base_terrap
                        End If
                    mom_lat_pos = (4000000 / 243) * tg_alpha * ancho_ent_mac * coef_compres_var * (alt_ent_mac ^ 3)
                    mom_base_pos = p_tot_mac_1 * (l_ent_mac / 2) + p_tot_mac_2 * (l_tot_mac - (2 / 3) * (l_tot_mac - l_ent_mac))
                    coef_pond = mom_lat_pos / mom_base_pos
                        If coef_pond <= 1 Then
                            coef_pond = 0.4167 * ((mom_lat_pos / mom_base_pos) ^ 2) - 0.9167 * (mom_lat_pos / mom_base_pos) + 1.5
                        Else: coef_pond = 1
                        End If
                    mom_vuelco_pos = (mom_lat_pos + mom_base_pos) / coef_pond
                    param_z = 0.001 * (Sqr((2 * (p_tot_mac_1 + p_tot_mac_2)) / (tg_alpha * coef_compres_var * ancho_ent_mac)))
                    mom_lat_neg = (4000000 / 243) * tg_alpha * ancho_ent_mac * coef_compres_var * (alt_ent_mac ^ 3)
                    mom_base_neg = p_tot_mac_1 * (l_tot_mac - (l_ent_mac / 2)) + p_tot_mac_2 * (((2 / 3) * (l_tot_mac - l_ent_mac)) - param_z) + (p_tot_mac_1 + p_tot_mac_2) * (1 / 3) * param_z
                    coef_pond = mom_lat_neg / mom_base_neg
                        If coef_pond <= 1 Then
                            coef_pond = 0.4167 * ((mom_lat_neg / mom_base_neg) ^ 2) - 0.9167 * (mom_lat_neg / mom_base_neg) + 1.5
                        Else: coef_pond = 1
                        End If
                    mom_vuelco_neg = mom_base_neg / coef_pond
                    mom_vuelco_neg = mom_vuelco_neg / 1.5
                    mom_vuelco = MAX(mom_vuelco_pos, mom_vuelco_neg)
                    If cim(i, 15) = "" Then
                        cim(i, 15) = mom_vuelco
                    End If
                ElseIf desm_terrap_mac = "anclaje" Then
                    v_terr = 2 * (alt_ent_mac * Tan(ang_roz * (3.1416 / 180)) * (l_ent_mac + ancho_ent_mac)) * alt_ent_mac + (4 / 3) * ((alt_ent_mac * Tan(ang_roz * (3.1416 / 180))) ^ 2) * alt_ent_mac
                    res_lat = cap_lat * l_ent_mac * alt_ent_mac
                    res_arr = v_tot_mac * p_esp_horm + v_terr * p_esp_terr
                    fuerza_anc = Sqr(2) * MIN(res_lat, res_arr)
                    mom_vuelco = 0
                    cim(i, 16) = fuerza_anc
                End If
            'a partir de este momento hay que desplazar el momento a la base del poste para saber cuanto vale, para ello es necesario conocer el tipo de poste y su altura.
                    
    ElseIf mac_mac = "Cilíndrico" Then
    
        If incl_ad_mac > 33 Then
            'no se debe usar los macizos cilíndricos
        End If
        diam_mac = cim(i, 14) 'ancho cimentación en sentido perpendicular a la vía
        alt_ent_mac = cim(i, 7) ' altura enterrada del macizo
        alt_nent_mac = cim(i, 11) 'altura no enterrada del macizo
        'p_esp_horm = 2200 'Valor obtenido del SiReCa (kg/m3)
        diam_mac = cim(i, 14) 'ancho cimentación en sentido perpendicular a la vía
        v_tot_mac = cim(i, 13)
        alt_tot_mac = alt_ent_mac + alt_nent_mac 'altura total del macizo
        If desm_terrap_mac = "desmonte" Then
            coef = coef_compres_base_desm
        ElseIf desm_terrap_mac = "terraplén" Then
            coef = coef_compres_base_terrap
        End If
        If alt_ent_mac <= 2 Then
            coef_compres_var = coef * (alt_ent_mac / alt_base_desm)
            Else: coef_compres_var = coef
        End If
        mom_lat = (1000000 / 36) * tg_alpha * coef_compres_var * diam_mac * (alt_ent_mac ^ 3)
        coef_pond = 1
        mom_vuelco = (mom_lat) / coef_pond
        If cim(i, 15) = "" Then
            cim(i, 15) = mom_vuelco
        End If
    ElseIf desm_terrap_mac = "anclaje" Then
        v_terr = 3.1416 * (((alt_ent_mac ^ 3) / 3) * Tan(ang_roz * (3.1416 / 180)) + (diam_mac / 2) * (alt_ent_mac ^ 2) * Tan(ang_roz * (3.1416 / 180)))
                    
                    'Hay que revisar la resistencia lateral
                    'res_lat = cap_lat * diam_mac * 3.1416 * diam_mac * alt_ent_mac
        res_arr = v_tot_mac * p_esp_horm + v_terr * p_esp_terr
        fuerza_anc = Sqr(2) * MIN(res_lat, res_arr)
        mom_vuelco = 0
        cim(i, 17) = fuerza_anc
    End If
            'a partir de este momento hay que desplazar el momento a la base del poste para saber cuanto vale, para ello es necesario conocer el tipo de poste y su altura.
    
    If cim(i, 15) = "" Then
        cim(i, 19) = mom_vuelco * (Sheets("Replanteo").Cells(z, 36) / (Sheets("Replanteo").Cells(z, 36) + alt_nent_mac + (2 / 3) * alt_ent_mac))
    Else
        cim(i, 19) = cim(i, 15)
    End If
    i = i + 1
    Wend
     
'//
'//ELECCIÓN CIMENTACIÓN EN FUNCIÓN MOMENTO POSTE
'//
  
    'z = 10
    i = 0
    
    While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
        If Sheets("Replanteo").Cells(z, 38).Value = "Tunel" Or Sheets("Replanteo").Cells(z, 38).Value = "Viaducto" Or Sheets("Replanteo").Cells(z, 38).Value = "Marquesina" Then '/// retocar
        Else
            'if sheets("Replanteo").Cells(z, XXX).Value = terraplen then
                'buscar = terraplen
            
            'else then
            
                'buscar = "desmonte"
            'End If
        mom_poste_var = Abs(Sheets("Replanteo").Cells(z, 19))
        
        While cim(i, 19) < mom_poste_var Or cim(i, 2) <> cimenta
            i = i + 1
        Wend
        
        mom_cim_var = cim(i, 19)
        tip_mac = cim(i, 3)
        vol_tot_mac = cim(i, 13)
X:
        i = 0
        
        While cim(i, 19) <> 0 And cim(i, 2) = cimenta
            If cim(i, 19) < mom_cim_var And cim(i, 19) > mom_poste_var Then
                mom_cim_var = cim(i, 19)
                tip_mac = cim(i, 3)
                vol_tot_mac = cim(i, 13)
                GoTo X
            Else: i = i + 1
            End If
        Wend
'//
'//INSERCIÓN CIMENTACIÓN EN REPLANTEO
'//
        Sheets("Replanteo").Cells(z, 37) = mom_cim_var
        Sheets("Replanteo").Cells(z, 22) = tip_mac
        Sheets("Replanteo").Cells(z, 23) = vol_tot_mac
        End If
    If CAD = True Then
        GoTo fin
    Else
        z = z + 2
        i = 0
    End If
    Wend
fin:
  End Sub
Public Function MAX(X As Double, Y As Double) As Double
Dim Resul As Double
If X > Y Then
Resul = X
ElseIf X <= Y Then
Resul = Y
End If
MAX = Resul
End Function
Public Function MIN(X As Double, Y As Double) As Double
Dim Resul As Double
If X > Y Then
Resul = Y
ElseIf X <= Y Then
Resul = X
End If
MIN = Resul
End Function

 



