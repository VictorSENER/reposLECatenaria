Attribute VB_Name = "cargar"
'//
'// Declaración de todas las variables públicas
'//
Public caso As String
Public sist As String, al As String, posicion_feed_pos As String
Public alt_nom As Double, alt_min As Double, alt_max As Double, alt_cat As Double, va_max_sla As Double, inc_max_alt_hc As Double, n_min_va_sm As Double, n_min_va_sla As Double, ancho_via As Double, d_max_re As Double, d_max_cu As Double, d_max_ad As Double, el_max_pant As Double, vw As Double, fl_max_centro_va As Double
Public dist_carril_poste As Double, dist_base_poste_pmr As Double, dist_elect_sm As Double, dist_elect_sla As Double
Public l_zc_max As Double, l_zc_min As Double, l_zn As Double, r_min_traz As Double, hc As String, sust As String, cdpa As String, cdte As String, feed_pos As String, feed_neg As String, pto_fijo As String, pend As String, anc As String, posicion_feed_neg As String, n_hc As Long, n_cdpa As Long, n_feed_pos As Long, n_feed_neg As Long, t_hc As Double, t_sust As Double, t_cdpa As Double
Public t_feed_neg As Double, t_feed_pos As Double, t_pto_fijo As Double, adm_lin_poste As String, tip_poste As String, num_poste As String, adm_lin_mac As String, tip_mac As String, tubo_men As String, tubo_tir As String, cola_anc As String, aisl_feed_pos As String, aisl_feed_neg As String, dist_ap_prim_pend As Double, dist_prim_seg_pend As Double, dist_max_pend As Double, idioma As String
Public sec_hc As Double, diam_hc As Double, p_hc As Double, res_max_hc As Double, coef_dil_hc As Double, mod_elast_hc As Double, carga_rot_hc As Double, norma_hc As String, origen_1_hc As String, origen_2_hc As String
Public sec_sust As Double, diam_sust As Double, p_sust As Double, res_max_sust As Double, coef_dil_sust As Double, mod_elast_sust As Double, carga_rot_sust As Double, norma_sust As String, origen_1_sust As String, origen_2_sust As String
Public sec_cdpa As Double, diam_cdpa As Double, p_cdpa As Double, res_max_cdpa As Double, coef_dil_cdpa As Double, mod_elast_cdpa As Double, carga_rot_cdpa As Double, norma_cdpa As String, origen_1_cdpa As String, origen_2_cdpa As String
Public sec_pto_fijo As Double, diam_pto_fijo As Double, p_pto_fijo, res_max_pto_fijo As Double, coef_dil_pto_fijo As Double, mod_elast_pto_fijo As Double, carga_rot_pto_fijo As Double, norma_pto_fijo As String, origen_1_pto_fijo As String, origen_2_pto_fijo As String
Public diam_feed_pos As Double, p_feed_pos As Double, res_max_feed_pos As Double, coef_dil_feed_pos As Double, mod_elast_feed_pos As Double, carga_rot_feed_pos As Double, norma_feed_pos As String, origen_1_feed_pos As String, origen_2_feed_pos As String
Public sec_feed_neg As Double, diam_feed_neg As Double, p_feed_neg As Double, res_max_feed_neg As Double, coef_dil_feed_neg As Double, mod_elast_feed_neg As Double, carga_rot_feed_neg As Double, norma_feed_neg As String, origen_1_feed_neg As String, origen_2_feed_neg As String
Public sec_cdte As Double, diam_cdte As Double, p_cdte As Double, res_max_cdte As Double, coef_dil_cdte As Double, mod_elast_cdte As Double, carga_rot_cdte As Double, norma_cdte As String, origen_1_cdte As String, origen_2_cdte As String
Public sec_pend As Double, diam_pend As Double, p_pend As Double, res_max_pend As Double, coef_dil_pend As Double, mod_elast_pend As Double, carga_rot_pend As Double, norma_pend As String, origen_1_pend As String, origen_2_pend As String
Public dist_vert_hc As Double, dist_horiz_hc As Double, dist_vert_sust As Double, dist_horiz_sust As Double, sep_hc As Double
Public dist_vert_feed_pos As Double, dist_horiz_feed_pos As Double, dist_vert_feed_neg As Double, dist_horiz_feed_neg As Double, dist_vert_cdpa As Double, dist_horiz_cdpa As Double, dist_horiz_equip As Double, dist_vert_hc_anc As Double, dist_vert_sust_anc As Double, alt_cat_se_sm_el As Double, alt_cat_e_sm As Double, alt_cat_se_sla_el As Double, alt_cat_e_sla As Double, alt_cat_se_ag_el As Double, alt_cat_e_ag As Double, alt_cat_se_zn_el As Double, alt_cat_e_zn As Double
Public el_hc As Double, ancho_carril As Double, dist_horiz_equip_t As Double, p_medio_equip_t As Double
Public dist_vert_hc_se_zn_el As Double, dist_horiz_hc_se_zn_el As Double, dist_vert_sust_se_zn_el As Double, dist_horiz_sust_se_zn_el As Double, dist_vert_hc_e_zn, dist_horiz_hc_e_zn As Double, dist_vert_sust_e_zn, dist_horiz_sust_e_zn As Double
Public anc_sm_con As String, anc_sm_sin As String, anc_sla_con As String, anc_sla_sin As String, semi_eje_sm As String, eje_aguj As String, anc_neutra As String, semi_eje_neutra As String
Public semi_eje_sla As String, eje_sm As String, eje_sla As String, anc_pf As String, eje_pf As String, anc_aguj As String, semi_eje_aguj As String, eje_neutra As String
Public d_semi_eje_sla1 As Double, d_semi_eje_sla2 As Double, d_eje_sla1 As Double, d_eje_sla2 As Double, d_semi_eje_sm1 As Double, d_semi_eje_sm2 As Double, d_eje_sm1 As Double, d_eje_sm2 As Double, dist_pant_util As Double, d_eje_aguj1 As Double, d_eje_aguj2 As Double, d_semi_eje_aguj1 As Double, d_semi_eje_aguj2 As Double
Public pas_sup As String, pue As String, con As String, tun As String, p_n As String, p_i As String, aguj As String, dren As String, via As String, est As String, pue_xl As String, mar As String, zon As String, sen As String, lin As String
Public SS As String, pot_ali As String
Public col() As String
Public lngCampos As Integer
Public cim(100, 100) As Variant
Public post(100, 100) As Variant
Public ruta_replanteo As String
Public fila_ini As Double, fila_fin As Double

'//
'// Rutina destinada a recoger los datos de catenaria de la base de datos (Acces)
'//
Sub datos_lac(nombre_catVB)
Dim oConn As ADODB.Connection
Dim oread As ADODB.Recordset
Dim strSQL As String
Dim strTabla As String
Dim lngTablas As Long
Dim i As Long
'//
'// Inicializar variables
'//
nombre_cat = nombre_catVB
'//
'// Obtener la dirección de la base de datos
'//
    'elegir uno de estas dos rutas al archivo Access
'strDB = "W:\223\D\D223041\CC_CALCULOS\SiReCa\Base de datos.accdb" 'si en otra carpeta
        
'//
'// nombre de la tabla del archivo Access
'//
    strTabla = "Datos"
'//
'// crear objetos y la conexión
'//
Set oConn = New ADODB.Connection
Set oread = New ADODB.Recordset
oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source =" & strDB & ";"
'//
'// consulta a la base de datos
'//
strSQL = "SELECT * FROM " & strTabla & ""
oread.Open strSQL, oConn
'//
'// realización de la consulta y copia de datos a las variables
'//

While Not oread.EOF
    If oread("nombre_cat") = nombre_cat Then
        
        sist = oread("sist")
        al = oread("al")
        alt_nom = oread("alt_nom")
        alt_min = oread("alt_min")
        alt_max = oread("alt_max")
        alt_cat = oread("alt_cat")
        dist_va_max = oread("dist_max_va")
        dist_max_canton = oread("dist_max_canton")
        va_max = oread("va_max")
        va_max_sm = oread("va_max_sm")
        va_max_sla = oread("va_max_sla")
        va_max_tunel = oread("va_max_tunel")
        inc_norm_va = oread("inc_norm_va")
        inc_max_alt_hc = oread("inc_max_alt_hc")
        n_min_va_sm = oread("n_min_va_sm")
        n_min_va_sla = oread("n_min_va_sla")
        ancho_via = oread("ancho_via")
        d_max_re = oread("d_max_re")
        d_max_cu = oread("d_max_cu")
        r_re = oread("r_re")
        d_max_ad = oread("d_max_ad")
        el_max_pant = oread("el_max_pant")
        vw = oread("vw")
        fl_max_centro_va = oread("fl_max_centro_va")
        dist_carril_poste = oread("dist_carril_poste")
        dist_base_poste_pmr = oread("dist_base_poste_pmr")
        dist_elect_sm = oread("dist_elect_sm")
        dist_elect_sla = oread("dist_elect_sla")
        l_zc_max = oread("l_zc_max")
        l_zc_min = oread("l_zc_min")
        l_zn = oread("l_zn")
        r_min_traz = oread("r_min_traz")
        hc = oread("hc")
        sust = oread("sust")
        cdpa = oread("cdpa")
        cdte = oread("cdte")
        feed_pos = oread("feed_pos")
        feed_neg = oread("feed_neg")
        pto_fijo = oread("pto_fijo")
        pend = oread("pend")
        anc = oread("anc")
        posicion_feed_pos = oread("posicion_feed_pos")
        posicion_feed_neg = oread("posicion_feed_neg")
        n_hc = oread("n_hc")
        n_cdpa = oread("n_cdpa")
        n_feed_pos = oread("n_feed_pos")
        n_feed_neg = oread("n_feed_neg")
        t_hc = oread("t_hc")
        t_sust = oread("t_sust")
        t_cdpa = oread("t_cdpa")
        t_feed_pos = oread("t_feed_pos")
        t_feed_neg = oread("t_feed_neg")
        t_pto_fijo = oread("t_pto_fijo")
        adm_lin_poste = oread("adm_lin_poste")
        tip_poste = oread("tip_poste")
        num_poste = oread("num_poste")
        adm_lin_mac = oread("adm_lin_mac")
        tip_mac = oread("tip_mac")
        tubo_men = oread("tubo_men")
        tubo_tir = oread("tubo_tir")
        cola_anc = oread("cola_anc")
        aisl_feed_pos = oread("feed_pos")
        aisl_feed_neg = oread("feed_neg")
        dist_ap_prim_pend = oread("dist_ap_prim_pend")
        dist_prim_seg_pend = oread("dist_prim_seg_pend")
        dist_max_pend = oread("dist_max_pend")
        idioma = oread("idioma")
        dist_vert_feed_pos = oread("dist_vert_feed_pos")
        dist_horiz_feed_pos = oread("dist_horiz_feed_pos")
        dist_vert_feed_neg = oread("dist_vert_feed_neg")
        dist_horiz_feed_neg = oread("dist_horiz_feed_neg")
        dist_vert_cdpa = oread("dist_vert_cdpa")
        dist_horiz_cdpa = oread("dist_horiz_cdpa")
        dist_horiz_equip_t = oread("dist_horiz_equip_t")
        dist_horiz_equip_comp = oread("dist_horiz_equip_comp")
        dist_vert_hc_anc = oread("dist_vert_hc_anc")
        dist_vert_sust_anc = oread("dist_vert_sust_anc")
        alt_cat_se_sm_el = oread("alt_cat_se_sm_el")
        alt_cat_e_sm = oread("alt_cat_e_sm")
        alt_cat_se_sla_el = oread("alt_cat_se_sla_el")
        alt_cat_e_sla = oread("alt_cat_e_sla")
        alt_cat_se_ag_el = oread("alt_cat_se_ag_el")
        alt_cat_e_ag = oread("alt_cat_e_ag")
        alt_cat_se_zn_el = oread("alt_cat_se_zn_el")
        alt_cat_e_zn = oread("alt_cat_e_zn")
        sep_hc = oread("sep_hc")
        p_medio_equip_t = oread("p_medio_equip_t")
        p_medio_equip_comp = oread("p_medio_equip_comp")
        el_hc = oread("el_hc")
        tip_carril = oread("tip_carril")
        ancho_carril = oread("ancho_carril")
        anc_sm_con = oread("anc_sm_con")
        anc_sm_sin = oread("anc_sm_sin")
        anc_sla_con = oread("anc_sla_con")
        anc_sla_sin = oread("anc_sla_sin")
        semi_eje_sm = oread("semi_eje_sm")
        eje_aguj = oread("eje_aguj")
        anc_neutra = oread("anc_neutra")
        semi_eje_neutra = oread("semi_eje_neutra")
        semi_eje_sla = oread("semi_eje_sla")
        eje_sm = oread("eje_sm")
        eje_sla = oread("eje_sla")
        anc_pf = oread("anc_pf")
        eje_pf = oread("eje_pf")
        anc_aguj = oread("anc_aguj")
        semi_eje_aguj = oread("semi_eje_aguj")
        eje_neutra = oread("eje_neutra")
        d_semi_eje_sla1 = oread("d_semi_eje_sla1")
        d_semi_eje_sla2 = oread("d_semi_eje_sla2")
        d_eje_sla1 = oread("d_eje_sla1")
        d_eje_sla2 = oread("d_eje_sla2")
        d_semi_eje_sm1 = oread("d_semi_eje_sm1")
        d_semi_eje_sm2 = oread("d_semi_eje_sm2")
        d_eje_sm1 = oread("d_eje_sm1")
        d_eje_sm2 = oread("d_eje_sm2")
        d_eje_aguj1 = oread("d_eje_aguj1")
        d_eje_aguj2 = oread("d_eje_aguj2")
        d_semi_eje_aguj1 = oread("d_semi_eje_aguj1")
        d_semi_eje_aguj2 = oread("d_semi_eje_aguj2")
        dist_pant_util = oread("dist_pant_util")
        GoTo fin
    End If
    i = i + 1
oread.MoveNext

Wend
fin:
'//
'// desconectar de la base de datos
'//
oread.Close: Set oread = Nothing
oConn.Close: Set oConn = Nothing
End Sub
'//
'// Rutina destinada a recoger los datos del cableado de la catenaria guardados en la base de datos (Acces)
'//
Sub datos_cable()
Dim oConn As ADODB.Connection
Dim oread As ADODB.Recordset
Dim strSQL As String
Dim strTabla As String
Dim lngTablas As Long
Dim i As Long

'//
'// nombre de la tabla del archivo Access
'//
strTabla = "Conductores_y_cables"
'//
'// crear conexión
'//
Set oConn = New ADODB.Connection
Set oread = New ADODB.Recordset
oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source =" & strDB & ";"
'//
'// consulta a la base de datos
'//
strSQL = "SELECT * FROM " & strTabla & ""
oread.Open strSQL, oConn
'//
'// copiar datos a las variables
'//
While Not oread.EOF
    If oread("tip_cyc") = ("HC") And oread("mat_cyc") = hc Then
        sec_hc = oread("sec_cyc")
        diam_hc = oread("diam_cyc")
        p_hc = oread("p_cyc")
        res_max_hc = oread("res_max_cyc")
        coef_dil_hc = oread("coef_dil_cyc")
        mod_elast_hc = oread("mod_elast_cyc")
        carga_rot_hc = oread("carga_rot_cyc")
        norma_hc = oread("norma_cyc")
        origen_1_hc = oread("origen_1_cyc")
        origen_2_hc = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("SUSTENTADOR") And oread("mat_cyc") = sust Then
        sec_sust = oread("sec_cyc")
        diam_sust = oread("diam_cyc")
        p_sust = oread("p_cyc")
        res_max_sust = oread("res_max_cyc")
        coef_dil_sust = oread("coef_dil_cyc")
        mod_elast_sust = oread("mod_elast_cyc")
        carga_rot_sust = oread("carga_rot_cyc")
        norma_sust = oread("norma_cyc")
        origen_1_sust = oread("origen_1_cyc")
        origen_2_sust = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("CDPA") And oread("mat_cyc") = cdpa Then
        sec_cdpa = oread("sec_cyc")
        diam_cdpa = oread("diam_cyc")
        p_cdpa = oread("p_cyc")
        res_max_cdpa = oread("res_max_cyc")
        coef_dil_cdpa = oread("coef_dil_cyc")
        mod_elast_cdpa = oread("mod_elast_cyc")
        carga_rot_cdpa = oread("carga_rot_cyc")
        norma_cdpa = oread("norma_cyc")
        origen_1_cdpa = oread("origen_1_cyc")
        origen_2_cdpa = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("PUNTO FIJO") And oread("mat_cyc") = pto_fijo Then
        sec_pto_fijo = oread("sec_cyc")
        diam_pto_fijo = oread("diam_cyc")
        p_pto_fijo = oread("p_cyc")
        res_max_pto_fijo = oread("res_max_cyc")
        coef_dil_pto_fijo = oread("coef_dil_cyc")
        mod_elast_pto_fijo = oread("mod_elast_cyc")
        carga_rot_pto_fijo = oread("carga_rot_cyc")
        norma_pto_fijo = oread("norma_cyc")
        origen_1_pto_fijo = oread("origen_1_cyc")
        origen_2_pto_fijo = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("FEEDER +") And oread("mat_cyc") = feed_pos Then
        sec_feed_pos = oread("sec_cyc")
        diam_feed_pos = oread("diam_cyc")
        p_feed_pos = oread("p_cyc")
        res_max_feed_pos = oread("res_max_cyc")
        coef_dil_feed_pos = oread("coef_dil_cyc")
        mod_elast_feed_pos = oread("mod_elast_cyc")
        carga_rot_feed_pos = oread("carga_rot_cyc")
        norma_feed_pos = oread("norma_cyc")
        origen_1_feed_pos = oread("origen_1_cyc")
        origen_2_feed_pos = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("FEEDER -") And oread("mat_cyc") = feed_neg Then
        sec_feed_neg = oread("sec_cyc")
        diam_feed_neg = oread("diam_cyc")
        p_feed_neg = oread("p_cyc")
        res_max_feed_neg = oread("res_max_cyc")
        coef_dil_feed_neg = oread("coef_dil_cyc")
        mod_elast_feed_neg = oread("mod_elast_cyc")
        carga_rot_feed_neg = oread("carga_rot_cyc")
        norma_feed_neg = oread("norma_cyc")
        origen_1_feed_neg = oread("origen_1_cyc")
        origen_2_feed_neg = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("CDTE") And oread("mat_cyc") = cdte Then
        sec_cdte = oread("sec_cyc")
        diam_cdte = oread("diam_cyc")
        p_cdte = oread("p_cyc")
        res_max_cdte = oread("res_max_cyc")
        coef_dil_cdte = oread("coef_dil_cyc")
        mod_elast_cdte = oread("mod_elast_cyc")
        carga_rot_cdte = oread("carga_rot_cyc")
        norma_cdte = oread("norma_cyc")
        origen_1_cdte = oread("origen_1_cyc")
        origen_2_cdte = oread("origen_2_cyc")
    End If
    If oread("tip_cyc") = ("PENDOLA") And oread("mat_cyc") = pend Then
        sec_pend = oread("sec_cyc")
        diam_pend = oread("diam_cyc")
        p_pend = oread("p_cyc")
        res_max_pend = oread("res_max_cyc")
        coef_dil_pend = oread("coef_dil_cyc")
        mod_elast_pend = oread("mod_elast_cyc")
        carga_rot_pend = oread("carga_rot_cyc")
        norma_pend = oread("norma_cyc")
        origen_1_pend = oread("origen_1_cyc")
        origen_2_pend = oread("origen_2_cyc")
    End If
oread.MoveNext
Wend
'//
'// desconectar de la base de datos
'//
oread.Close: Set oread = Nothing
oConn.Close: Set oConn = Nothing
End Sub

Sub datos_poste()

'//
'//INSERCIÓN DATOS EN HOJA ANEXA DE REPLANTEO
'//
    
    Dim oConn As ADODB.Connection
    Dim oread As ADODB.Recordset
    Dim strSQL As String
    Dim strTabla As String
    Dim lngTablas As Long
    Dim i As Long
    'elegir uno de estas dos rutas al archivo Access
    'strDB = "W:\223\D\D223041\CC_CALCULOS\SiReCa\Base de datos.accdb"
    'nombre de la tabla del archivo Access
    strTabla = "Postes"
    'crear la conexión
    Set oConn = New ADODB.Connection
    Set oread = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oread.Open strSQL, oConn
    'copiar datos a la hoja
    
    j = 0
    'mientras hayan registros
    While Not oread.EOF
 
    If oread.Fields(1).Value = adm_lin_poste And oread.Fields(0).Value = tip_poste Then
    
      lngCampos = oread.Fields.count
      For i = 0 To lngCampos - 1
          post(j, i) = oread.Fields(i).Value
      Next
      j = j + 1
    End If
    'saltar al siguiente registro
    oread.MoveNext
    Wend
    'desconectar
    oread.Close: Set oread = Nothing
    oConn.Close: Set oConn = Nothing
End Sub

Sub datos_cim()
    Dim oConn As ADODB.Connection
    Dim oread As ADODB.Recordset
    Dim strSQL As String
    Dim strTabla As String
    Dim lngTablas As Long
    Dim i As Long

    'elegir uno de estas dos rutas al archivo Access
    'strDB = "W:\223\D\D223041\CC_CALCULOS\SiReCa\Base de datos.accdb"
    'nombre de la tabla del archivo Access
    strTabla = "Macizos"
    'crear la conexión
    Set oConn = New ADODB.Connection
    Set oread = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oread.Open strSQL, oConn
    'copiar datos a la hoja
        
    j = 0
    'mientras hayan registros
    While Not oread.EOF
 
      'se debe leer del programa
      If oread.Fields(0).Value = tip_mac And oread.Fields(1).Value = adm_lin_mac And oread.Fields(2).Value = "desmonte" Then
          lngCim_des = oread.Fields.count
            For i = 0 To lngCim_des - 1
              cim(j, i) = oread.Fields(i).Value
            Next
        j = j + 1
      End If
      'saltar al siguiente registro
      oread.MoveNext

    Wend
    oread.MoveFirst

    'mientras hayan registros
    While Not oread.EOF
 
      'se debe leer del programa
      If oread.Fields(0).Value = tip_mac And oread.Fields(1).Value = adm_lin_mac And oread.Fields(2).Value = "terraplén" Then
      
          lngCim_terr = oread.Fields.count
            For i = 0 To lngCim_terr - 1
                cim(j, i) = oread.Fields(i).Value
            Next
        
          j = j + 1
      End If
      'saltar al siguiente registro
      oread.MoveNext
    Wend
    'desconectar
    oread.Close: Set oread = Nothing
    oConn.Close: Set oConn = Nothing
End Sub

Sub punto_singular(idiomaVB)
    Dim oConn As ADODB.Connection
    Dim oread As ADODB.Recordset
    Dim strSQL As String
    Dim strTabla As String
    Dim lngTablas As Long

    strTabla = "Punto_singular"
    'crear la conexión
    Set oConn = New ADODB.Connection
    Set oread = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oread.Open strSQL, oConn
    'copiar datos a la hoja
        
    'mientras hayan registros
    While Not oread.EOF
 
      'se debe leer del programa
        If oread.Fields(0).Value = idiomaVB Then
      
            lngCampos = oread.Fields.count
            pas_sup = oread("pas_sup")
            pue = oread("pue")
            con = oread("con")
            tun = oread("tun")
            p_n = oread("pn")
            p_i = oread("pi")
            aguj = oread("aguj")
            dren = oread("dren")
            via = oread("via")
            est = oread("est")
            pue_xl = oread("pue_xl")
            mar = oread("mar")
            zon = oread("zon")
            sen = oread("sin")
            lin = oread("lin")
            SS = oread("SS")
            pot_ali = oread("pot_ali")
        End If
      'saltar al siguiente registro
      oread.MoveNext
    Wend

    'desconectar
    oread.Close: Set oread = Nothing
    oConn.Close: Set oConn = Nothing
End Sub
Sub cabecera(idiomaVB)
    Dim oConn As ADODB.Connection
    Dim oread As ADODB.Recordset
    Dim strSQL As String
    Dim strTabla As String
    Dim lngTablas As Long
    
    strTabla = "Cabecera"
    'crear la conexión
    Set oConn = New ADODB.Connection
    Set oread = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oread.Open strSQL, oConn
    'copiar datos a la hoja
    'mientras hayan registros
    While Not oread.EOF
 
      'se debe leer del programa
      If oread.Fields(0).Value = idiomaVB Then
      
          lngCampos = oread.Fields.count
            ReDim col(lngCampos) As String
            For i = 1 To lngCampos - 1
            col(i) = oread("col_" & i)
            Next
        
      End If
      'saltar al siguiente registro
      oread.MoveNext
    Wend

    'desconectar
    oread.Close: Set oread = Nothing
    oConn.Close: Set oConn = Nothing
End Sub

