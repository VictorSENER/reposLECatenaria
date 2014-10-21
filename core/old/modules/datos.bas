Attribute VB_Name = "datos"
Public caso As String
Public uno As Integer
Public sist As String, al As String, alt_nom As Double, alt_min As Double, alt_max As Long, alt_cat As Double, va_max_sla As Double, inc_max_alt_hc As Double, n_min_va_sm As Double, n_min_va_sla As Double, ancho_via As Double, d_max_re As Double, d_max_cu As Double, d_max_ad As Double, el_max_pant As Double, vw As Double, fl_max_centro_va As Double, dist_carril_poste As Double, dist_base_poste_pmr As Double, dist_elect_sm As Double, dist_elect_sla As Double, l_zc_max As Double, l_zc_min As Double, l_zn As Double, r_min_traz As Double, hc As String, sust As String, cdpa As String, cdte As String, feed_pos As String, feed_neg As String, pto_fijo As String, pend As String, anc As String, posicion_feed_neg As String, n_hc As Long, n_cdpa As Long, n_feed_pos As Long, n_feed_neg As Long, t_hc As Double, t_sust As Double, t_cdpa As Double, t_
Public t_feed_neg As Double, t_pto_fijo As Double, adm_lin_poste As String, tip_poste As String, num_poste As String, adm_lin_mac As String, tip_mac As String, tubo_men As String, tubo_tir As String, cola_anc As String, aisl_feed_pos As String, aisl_feed_neg As String, dist_ap_prim_pend As Long, dist_prim_seg_pend As Long, dist_max_pend As Long, idioma As String
Public sec_hc As Double, diam_hc As Double, p_hc As Double, res_max_hc As Double, coef_dil_hc As Double, mod_elast_hc As Double, carga_rot_hc As Double, norma_hc As String, origen_1_hc As String, origen_2_hc As String
Public sec_sust As Double, diam_sust As Double, p_sust As Double, res_max_sust As Double, coef_dil_sust As Double, mod_elast_sust As Double, carga_rot_sust As Double, norma_sust As String, origen_1_sust As String, origen_2_sust As String
Public sec_cdpa As Double, diam_cdpa As Double, p_cdpa As Double, res_max_cdpa As Double, coef_dil_cdpa As Double, mod_elast_cdpa As Double, carga_rot_cdpa As Double, norma_cdpa As String, origen_1_cdpa As String, origen_2_cdpa As String
Public sec_pto_fijo As Double, diam_pto_fijo As Double, p_pto_fijo, res_max_pto_fijo As Double, coef_dil_pto_fijo As Double, mod_elast_pto_fijo As Double, carga_rot_pto_fijo As Double, norma_pto_fijo As String, origen_1_pto_fijo As String, origen_2_pto_fijo As String
Public diam_feed_pos As Double, p_feed_pos As Double, res_max_feed_pos As Double, coef_dil_feed_pos As Double, mod_elast_feed_pos As Double, carga_rot_feed_pos As Double, norma_feed_pos As String, origen_1_feed_pos As String, origen_2_feed_pos As String
Public sec_feed_neg As Double, diam_feed_neg As Double, p_feed_neg As Double, res_max_feed_neg As Double, coef_dil_feed_neg As Double, mod_elast_feed_neg As Double, carga_rot_feed_neg As Double, norma_feed_neg As String, origen_1_feed_neg As String, origen_2_feed_neg As String
Public sec_cdte As Double, diam_cdte As Double, p_cdte As Double, res_max_cdte As Double, coef_dil_cdte As Double, mod_elast_cdte As Double, carga_rot_cdte As Double, norma_cdte As String, origen_1_cdte As String, origen_2_cdte As String
Public sec_pend As Double, diam_pend As Double, p_pend As Double, res_max_pend As Double, coef_dil_pend As Double, mod_elast_pend As Double, carga_rot_pend As Double, norma_pend As String, origen_1_pend As String, origen_2_pend As String
Public dist_vert_hc As Double, dist_horiz_hc As Double, dist_vert_sust As Double, dist_horiz_sust As Double, sep_hc As Double
Public dist_vert_feed_pos As Double, dist_horiz_feed_pos, dist_vert_feed_neg As Double, dist_horiz_feed_neg, dist_vert_cdpa As Double, dist_horiz_cdpa, dist_horiz_equip As Double, dist_vert_hc_anc, dist_vert_sust_anc As Double, dist_vert_hc_se_sm_el, dist_horiz_hc_se_sm_el As Double, dist_vert_sust_se_sm_el, dist_horiz_sust_se_sm_el As Double, dist_vert_hc_e_sm, dist_horiz_hc_e_sm As Double, dist_vert_sust_e_sm, dist_horiz_sust_e_sm As Double, dist_vert_hc_se_sla_el, dist_horiz_hc_se_sla_el As Double, dist_vert_sust_se_sla_el, dist_horiz_sust_se_sla_el As Double, dist_vert_hc_e_sla, dist_horiz_hc_e_sla As Double, dist_vert_sust_e_sla, dist_horiz_sust_e_sla As Double, dist_vert_hc_se_ag_el, dist_horiz_hc_se_ag_el As Double, dist_vert_sust_se_ag_el, dist_horiz_sust_se_ag_el As Double, dist_vert_hc_e_ag, dist_horiz_hc_e_ag As Double, dist_vert_sust_e_ag, dist_horiz_sust_e_ag As Double
Public dist_vert_hc_se_zn_el As Double, dist_horiz_hc_se_zn_el As Double, dist_vert_sust_se_zn_el As Double, dist_horiz_sust_se_zn_el As Double, dist_vert_hc_e_zn, dist_horiz_hc_e_zn As Double, dist_vert_sust_e_zn, dist_horiz_sust_e_zn As Double
Sub datos_acces(nombre_catVB)

Dim oConn As ADODB.Connection
Dim oRead As ADODB.Recordset
Dim strDB, strSQL As String
Dim strTabla As String
Dim lngTablas As Long
Dim i As Long
    nombre_cat = nombre_catVB
    'elegir uno de estas dos rutas al archivo Access
    strDB = "C:\Documents and Settings\23370\Escritorio\SiReCa\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb" 'si en otra carpeta
    'nombre de la tabla del archivo Access
    strTabla = "Datos"
    'crear la conexi�n
    Set oConn = New ADODB.Connection
    Set oRead = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oRead.Open strSQL, oConn
    
    'copiar datos a la hoja
    
    While Not oRead.EOF
    If oRead("nombre_cat") = nombre_cat Then
    
                sist = oRead("sist")
                al = oRead("al")
                alt_nom = oRead("alt_nom")
                alt_min = oRead("alt_min")
                alt_max = oRead("alt_max")
                alt_cat = oRead("alt_cat")
                dist_va_max = oRead("dist_max_va")
                dist_max_canton = oRead("dist_max_canton")
                va_max = oRead("va_max")
                va_max_sm = oRead("va_max_sm")
                va_max_sla = oRead("va_max_sla")
                va_max_tunel = oRead("va_max_tunel")
                inc_norm_va = oRead("inc_norm_va")
                inc_max_alt_hc = oRead("inc_max_alt_hc")
                n_min_va_sm = oRead("n_min_va_sm")
                n_min_va_sla = oRead("n_min_va_sla")
                ancho_via = oRead("ancho_via")
                d_max_re = oRead("d_max_re")
                d_max_cu = oRead("d_max_cu")
                r_re = oRead("r_re")
                d_max_ad = oRead("d_max_ad")
                el_max_pant = oRead("el_max_pant")
                vw = oRead("vw")
                fl_max_centro_va = oRead("fl_max_centro_va")
                dist_carril_poste = oRead("dist_carril_poste")
                dist_base_poste_pmr = oRead("dist_base_poste_pmr")
                dist_elect_sm = oRead("dist_elect_sm")
                dist_elect_sla = oRead("dist_elect_sla")
                l_zc_max = oRead("l_zc_max")
                l_zc_min = oRead("l_zc_min")
                l_zn = oRead("l_zn")
                r_min_traz = oRead("r_min_traz")
                hc = oRead("hc")
                sust = oRead("sust")
                cdpa = oRead("cdpa")
                cdte = oRead("cdte")
                feed_pos = oRead("feed_pos")
                feed_neg = oRead("feed_neg")
                pto_fijo = oRead("pto_fijo")
                pend = oRead("pend")
                anc = oRead("anc")
                posicion_feed_pos = oRead("posicion_feed_pos")
                posicion_feed_neg = oRead("posicion_feed_neg")
                n_hc = oRead("n_hc")
                n_cdpa = oRead("n_cdpa")
                n_feed_pos = oRead("n_feed_pos")
                n_feed_neg = oRead("n_feed_neg")
                t_hc = oRead("t_hc")
                t_sust = oRead("t_sust")
                t_cdpa = oRead("t_cdpa")
                t_feed_pos = oRead("t_feed_pos")
                t_feed_neg = oRead("t_feed_neg")
                t_pto_fijo = oRead("t_pto_fijo")
                adm_lin_poste = oRead("adm_lin_poste")
                tip_poste = oRead("tip_poste")
                num_poste = oRead("num_poste")
                adm_lin_mac = oRead("adm_lin_mac")
                tip_mac = oRead("tip_mac")
                tubo_men = oRead("tubo_men")
                tubo_tir = oRead("tubo_tir")
                cola_anc = oRead("cola_anc")
                aisl_feed_pos = oRead("feed_pos")
                aisl_feed_neg = oRead("feed_neg")
                dist_ap_prim_pend = oRead("dist_ap_prim_pend")
                dist_prim_seg_pend = oRead("dist_prim_seg_pend")
                dist_max_pend = oRead("dist_max_pend")
                idioma = oRead("idioma")
                dist_vert_hc = oRead("dist_vert_hc")
                dist_horiz_hc = oRead("dist_horiz_hc")
                dist_vert_sust = oRead("dist_vert_sust")
                dist_horiz_sust = oRead("dist_horiz_sust")
                dist_vert_feed_pos = oRead("dist_vert_feed_pos")
                dist_horiz_feed_pos = oRead("dist_horiz_feed_pos")
                dist_vert_feed_neg = oRead("dist_vert_feed_neg")
                dist_horiz_feed_neg = oRead("dist_horiz_feed_neg")
                dist_vert_cdpa = oRead("dist_vert_cdpa")
                dist_horiz_cdpa = oRead("dist_horiz_cdpa")
                dist_horiz_equip = oRead("dist_horiz_equip")
                dist_vert_hc_anc = oRead("dist_vert_hc_anc")
                dist_vert_sust_anc = oRead("dist_vert_sust_anc")
                dist_vert_hc_se_sm_el = oRead("dist_vert_hc_se_sm_el")
                dist_horiz_hc_se_sm_el = oRead("dist_horiz_hc_se_sm_el")
                dist_vert_sust_se_sm_el = oRead("dist_vert_sust_se_sm_el")
                dist_horiz_sust_se_sm_el = oRead("dist_horiz_sust_se_sm_el")
                dist_vert_hc_e_sm = oRead("dist_vert_hc_e_sm")
                dist_horiz_hc_e_sm = oRead("dist_horiz_hc_e_sm")
                dist_vert_sust_e_sm = oRead("dist_vert_sust_e_sm")
                dist_horiz_sust_e_sm = oRead("dist_horiz_sust_e_sm")
                dist_vert_hc_se_sla_el = oRead("dist_vert_hc_se_sla_el")
                dist_horiz_hc_se_sla_el = oRead("dist_horiz_hc_se_sla_el")
                dist_vert_sust_se_sla_el = oRead("dist_vert_sust_se_sla_el")
                dist_horiz_sust_se_sla_el = oRead("dist_horiz_sust_se_sla_el")
                dist_vert_hc_e_sla = oRead("dist_vert_hc_e_sla")
                dist_horiz_hc_e_sla = oRead("dist_horiz_hc_e_sla")
                dist_vert_sust_e_sla = oRead("dist_vert_sust_e_sla")
                dist_horiz_sust_e_sla = oRead("dist_horiz_sust_e_sla")
                dist_vert_hc_se_ag_el = oRead("dist_vert_hc_se_ag_el")
                dist_horiz_hc_se_ag_el = oRead("dist_horiz_hc_se_ag_el")
                dist_vert_sust_se_ag_el = oRead("dist_vert_sust_se_ag_el")
                dist_horiz_sust_se_ag_el = oRead("dist_horiz_sust_se_ag_el")
                dist_vert_hc_e_ag = oRead("dist_vert_hc_e_ag")
                dist_horiz_hc_e_ag = oRead("dist_horiz_hc_e_ag")
                dist_vert_sust_e_ag = oRead("dist_vert_sust_e_ag")
                dist_horiz_sust_e_ag = oRead("dist_horiz_sust_e_ag")
                dist_vert_hc_se_zn_el = oRead("dist_vert_hc_se_zn_el")
                dist_horiz_hc_se_zn_el = oRead("dist_horiz_hc_se_zn_el")
                dist_vert_sust_se_zn_el = oRead("dist_vert_sust_se_zn_el")
                dist_horiz_sust_se_zn_el = oRead("dist_horiz_sust_se_zn_el")
                dist_vert_hc_e_zn = oRead("dist_vert_hc_e_zn")
                dist_horiz_hc_e_zn = oRead("dist_horiz_hc_e_zn")
                dist_vert_sust_e_zn = oRead("dist_vert_sust_e_zn")
                dist_horiz_sust_e_zn = oRead("dist_horiz_sust_e_zn")
                sep_hc = oRead("sep_hc")
    End If
    oRead.MoveNext
    Wend
    oRead.Close: Set oRead = Nothing
    oConn.Close: Set oConn = Nothing
    'nombre de la tabla del archivo Access
    strTabla = "Conductores_y_cables"
    'crear la conexi�n
    Set oConn = New ADODB.Connection
    Set oRead = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oRead.Open strSQL, oConn
    'copiar datos a la hoja
    While Not oRead.EOF
            If oRead("tip_cyc") = ("HC") And oRead("mat_cyc") = hc Then
                sec_hc = oRead("sec_cyc")
                diam_hc = oRead("diam_cyc")
                p_hc = oRead("p_cyc")
                res_max_hc = oRead("res_max_cyc")
                coef_dil_hc = oRead("coef_dil_cyc")
                mod_elast_hc = oRead("mod_elast_cyc")
                carga_rot_hc = oRead("carga_rot_cyc")
                norma_hc = oRead("norma_cyc")
                origen_1_hc = oRead("origen_1_cyc")
                origen_2_hc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("SUSTENTADOR") And oRead("mat_cyc") = sust Then
                sec_sust = oRead("sec_cyc")
                diam_sust = oRead("diam_cyc")
                p_sust = oRead("p_cyc")
                res_max_sust = oRead("res_max_cyc")
                coef_dil_sust = oRead("coef_dil_cyc")
                mod_elast_sust = oRead("mod_elast_cyc")
                carga_rot_sust = oRead("carga_rot_cyc")
                norma_sust = oRead("norma_cyc")
                origen_1_sust = oRead("origen_1_cyc")
                origen_2_sust = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("CDPA") And oRead("mat_cyc") = cdpa Then
                sec_cdpa = oRead("sec_cyc")
                diam_cdpa = oRead("diam_cyc")
                p_cdpa = oRead("p_cyc")
                res_max_cdpa = oRead("res_max_cyc")
                coef_dil_cdpa = oRead("coef_dil_cyc")
                mod_elast_cdpa = oRead("mod_elast_cyc")
                carga_rot_cdpa = oRead("carga_rot_cyc")
                norma_cdpa = oRead("norma_cyc")
                origen_1_cdpa = oRead("origen_1_cyc")
                origen_2_cdpa = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("PUNTO FIJO") And oRead("mat_cyc") = pto_fijo Then
                sec_pto_fijo = oRead("sec_cyc")
                diam_pto_fijo = oRead("diam_cyc")
                p_pto_fijo = oRead("p_cyc")
                res_max_pto_fijo = oRead("res_max_cyc")
                coef_dil_pto_fijo = oRead("coef_dil_cyc")
                mod_elast_pto_fijo = oRead("mod_elast_cyc")
                carga_rot_pto_fijo = oRead("carga_rot_cyc")
                norma_pto_fijo = oRead("norma_cyc")
                origen_1_pto_fijo = oRead("origen_1_cyc")
                origen_2_pto_fijo = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("FEEDER +") And oRead("mat_cyc") = feed_pos Then
                sec_feed_pos = oRead("sec_cyc")
                diam_feed_pos = oRead("diam_cyc")
                p_feed_pos = oRead("p_cyc")
                res_max_feed_pos = oRead("res_max_cyc")
                coef_dil_feed_pos = oRead("coef_dil_cyc")
                mod_elast_feed_pos = oRead("mod_elast_cyc")
                carga_rot_feed_pos = oRead("carga_rot_cyc")
                norma_feed_pos = oRead("norma_cyc")
                origen_1_feed_pos = oRead("origen_1_cyc")
                origen_2_feed_pos = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("FEEDER -") And oRead("mat_cyc") = feed_neg Then
                sec_feed_neg = oRead("sec_cyc")
                diam_feed_neg = oRead("diam_cyc")
                p_feed_neg = oRead("p_cyc")
                res_max_feed_neg = oRead("res_max_cyc")
                coef_dil_feed_neg = oRead("coef_dil_cyc")
                mod_elast_feed_neg = oRead("mod_elast_cyc")
                carga_rot_feed_neg = oRead("carga_rot_cyc")
                norma_feed_neg = oRead("norma_cyc")
                origen_1_feed_neg = oRead("origen_1_cyc")
                origen_2_feed_neg = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("CDTE") And oRead("mat_cyc") = cdte Then
                sec_cdte = oRead("sec_cyc")
                diam_cdte = oRead("diam_cyc")
                p_cdte = oRead("p_cyc")
                res_max_cdte = oRead("res_max_cyc")
                coef_dil_cdte = oRead("coef_dil_cyc")
                mod_elast_cdte = oRead("mod_elast_cyc")
                carga_rot_cdte = oRead("carga_rot_cyc")
                norma_cdte = oRead("norma_cyc")
                origen_1_cdte = oRead("origen_1_cyc")
                origen_2_cdte = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("PENDOLA") And oRead("mat_cyc") = pend Then
                sec_pend = oRead("sec_cyc")
                diam_pend = oRead("diam_cyc")
                p_pend = oRead("p_cyc")
                res_max_pend = oRead("res_max_cyc")
                coef_dil_pend = oRead("coef_dil_cyc")
                mod_elast_pend = oRead("mod_elast_cyc")
                carga_rot_pend = oRead("carga_rot_cyc")
                norma_pend = oRead("norma_cyc")
                origen_1_pend = oRead("origen_1_cyc")
                origen_2_pend = oRead("origen_2_cyc")
            End If
        oRead.MoveNext
        Wend
    'desconectar
    oRead.Close: Set oRead = Nothing
    oConn.Close: Set oConn = Nothing

End Sub

