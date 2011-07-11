'//
'//Importar libreria
'//
Imports System.Data.OleDb
'//
'// Rutina destinada a aprovechar catenaria existente y realizar cambios
'//
Module modificar_lac
    Sub modificar_lac()
        '//
        '// crear objeto base de datos y abrir la base de datos
        '//
        Dim oConn As New OleDbConnection
        Dim oComm As OleDbCommand
        Dim oRead As OleDbDataReader
        'oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\23370\Escritorio\SiReCa\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn.Open()
        '//
        '// realiación de la consulta
        '//
        oComm = New OleDbCommand("select * from Datos", oConn)
        oRead = oComm.ExecuteReader
        While oRead.Read
            '//
            '// cargar los datos generales de la catenaria de la BBDD a la pantalla
            '//
            If (Pantalla_principal.nueva_lac = oRead("nombre_cat")) Then
                Pantalla_datos.Combo_sist.Text = oRead("sist")
                Pantalla_datos.Text_alt_nom.Text = oRead("alt_nom")
                Pantalla_datos.Text_alt_min.Text = oRead("alt_min")
                Pantalla_datos.Text_alt_max.Text = oRead("alt_max")
                Pantalla_datos.Text_alt_cat.Text = oRead("alt_cat")
                Pantalla_datos.Text_dist_max_va.Text = oRead("dist_max_va")
                Pantalla_datos.Text_dist_max_canton.Text = oRead("dist_max_canton")
                Pantalla_datos.Text_va_max.Text = oRead("va_max")
                Pantalla_datos.Text_va_max_sm.Text = oRead("va_max_sm")
                Pantalla_datos.Text_va_max_sla.Text = oRead("va_max_sla")
                Pantalla_datos.Text_va_max_tunel.Text = oRead("va_max_tunel")
                Pantalla_datos.Text_inc_norm_va.Text = oRead("inc_norm_va")
                Pantalla_datos.Text_inc_max_alt_hc.Text = oRead("inc_max_alt_hc")
                Pantalla_datos.Text_n_min_va_sm.Text = oRead("n_min_va_sm")
                Pantalla_datos.Text_n_min_va_sla.Text = oRead("n_min_va_sla")
                Pantalla_datos.Text_dist_base_poste_pmr.Text = oRead("dist_base_poste_pmr")
                Pantalla_datos.Text_ancho_via.Text = oRead("ancho_via")
                Pantalla_datos.Text_d_max_cu.Text = oRead("d_max_cu")
                Pantalla_datos.Text_d_max_re.Text = oRead("d_max_re")
                'Pantalla_datos.Text_d_max_ad.Text = oRead("d_max_ad")
                Pantalla_datos.Text_r_re.Text = oRead("r_re")
                Pantalla_datos.Text_d_max_ad.Text = oRead("d_max_ad")
                Pantalla_datos.Text_el_max_pant.Text = oRead("el_max_pant")
                Pantalla_datos.Text_vw.Text = oRead("vw")
                Pantalla_datos.Text_fl_max_centro_va.Text = oRead("fl_max_centro_va")
                Pantalla_datos.Text_dist_carril_poste.Text = oRead("dist_carril_poste")
                Pantalla_datos.Text_dist_base_poste_pmr.Text = oRead("dist_base_poste_pmr")
                Pantalla_datos.Text_dist_elect_sm.Text = oRead("dist_elect_sm")
                Pantalla_datos.Text_dist_elect_sla.Text = oRead("dist_elect_sla")
                Pantalla_datos.Text_l_zc_max.Text = oRead("l_zc_max")
                Pantalla_datos.Text_l_zc_min.Text = oRead("l_zc_min")
                Pantalla_datos.Text_l_zn.Text = oRead("l_zn")
                Pantalla_datos.Text_r_min_traz.Text = oRead("r_min_traz")
                Pantalla_datos.Combo_hc.Text = oRead("hc")
                Pantalla_datos.Combo_sust.Text = oRead("sust")
                Pantalla_datos.Combo_cdpa.Text = oRead("cdpa")
                Pantalla_datos.Combo_cdte.Text = oRead("cdte")
                Pantalla_datos.Combo_feed_pos.Text = oRead("feed_pos")
                Pantalla_datos.Combo_feed_neg.Text = oRead("feed_neg")
                Pantalla_datos.Combo_pto_fijo.Text = oRead("pto_fijo")
                Pantalla_datos.Combo_pend.Text = oRead("pend")
                Pantalla_datos.Combo_anc.Text = oRead("anc")
                Pantalla_datos.Combo_posicion_feed_pos.Text = oRead("posicion_feed_pos")
                Pantalla_datos.Combo_posicion_feed_neg.Text = oRead("posicion_feed_neg")
                Pantalla_datos.Text_n_hc.Text = oRead("n_hc")
                Pantalla_datos.Text_n_cdpa.Text = oRead("n_cdpa")
                Pantalla_datos.Text_n_feed_pos.Text = oRead("n_feed_pos")
                Pantalla_datos.Text_n_feed_neg.Text = oRead("n_feed_neg")
                Pantalla_datos.Text_t_hc.Text = oRead("t_hc")
                Pantalla_datos.Text_t_sust.Text = oRead("t_sust")
                Pantalla_datos.Text_t_cdpa.Text = oRead("t_cdpa")
                Pantalla_datos.Text_t_feed_pos.Text = oRead("t_feed_pos")
                Pantalla_datos.Text_t_feed_neg.Text = oRead("t_feed_neg")
                Pantalla_datos.Text_t_pto_fijo.Text = oRead("t_pto_fijo")
                Pantalla_datos.Combo_adm_lin_poste.Text = oRead("adm_lin_poste")
                Pantalla_datos.Text_tip_poste.Text = oRead("tip_poste")
                Pantalla_datos.Combo_num_poste.Text = oRead("num_poste")
                Pantalla_datos.Combo_adm_lin_mac.Text = oRead("adm_lin_mac")
                Pantalla_datos.Text_tip_mac.Text = oRead("tip_mac")
                Pantalla_datos.Combo_tubo_men.Text = oRead("tubo_men")
                Pantalla_datos.Combo_tubo_tir.Text = oRead("tubo_tir")
                Pantalla_datos.Combo_cola_anc.Text = oRead("cola_anc")
                Pantalla_datos.Combo_aisl_feed_pos.Text = oRead("feed_pos")
                Pantalla_datos.Combo_aisl_feed_neg.Text = oRead("feed_neg")
                Pantalla_datos.Text_dist_ap_prim_pend.Text = oRead("dist_ap_prim_pend")
                Pantalla_datos.Text_dist_prim_seg_pend.Text = oRead("dist_prim_seg_pend")
                Pantalla_datos.Text_dist_max_pend.Text = oRead("dist_max_pend")
                Pantalla_datos.Combo_idioma.Text = oRead("idioma")
                Pantalla_datos.Text_dist_vert_feed_pos.Text = oRead("dist_vert_feed_pos")
                Pantalla_datos.Text_dist_horiz_feed_pos.Text = oRead("dist_horiz_feed_pos")
                Pantalla_datos.Text_dist_vert_feed_neg.Text = oRead("dist_vert_feed_neg")
                Pantalla_datos.Text_dist_horiz_feed_neg.Text = oRead("dist_horiz_feed_neg")
                Pantalla_datos.Text_dist_vert_cdpa.Text = oRead("dist_vert_cdpa")
                Pantalla_datos.Text_dist_horiz_cdpa.Text = oRead("dist_horiz_cdpa")
                Pantalla_datos.Text_dist_horiz_equip_t.Text = oRead("dist_horiz_equip_t")
                Pantalla_datos.Text_dist_horiz_equip_comp.Text = oRead("dist_horiz_equip_comp")
                Pantalla_datos.Text_dist_vert_hc_anc.Text = oRead("dist_vert_hc_anc")
                Pantalla_datos.Text_dist_vert_sust_anc.Text = oRead("dist_vert_sust_anc")
                Pantalla_datos.Text_alt_cat_se_sm_el.Text = oRead("alt_cat_se_sm_el")
                Pantalla_datos.Text_alt_cat_e_sm.Text = oRead("alt_cat_e_sm")
                Pantalla_datos.Text_alt_cat_se_sla_el.Text = oRead("alt_cat_se_sla_el")
                Pantalla_datos.Text_alt_cat_e_sla.Text = oRead("alt_cat_e_sla")
                Pantalla_datos.Text_alt_cat_se_ag_el.Text = oRead("alt_cat_se_ag_el")
                Pantalla_datos.Text_alt_cat_e_ag.Text = oRead("alt_cat_e_ag")
                Pantalla_datos.Text_alt_cat_se_zn_el.Text = oRead("alt_cat_se_zn_el")
                Pantalla_datos.Text_alt_cat_e_zn.Text = oRead("alt_cat_e_zn")
                Pantalla_datos.Text_sep_hc.Text = oRead("sep_hc")
                Pantalla_datos.Text_p_medio_equip_t.Text = oRead("p_medio_equip_t")
                Pantalla_datos.Text_p_medio_equip_comp.Text = oRead("p_medio_equip_comp")
                Pantalla_datos.Text_el_hc.Text = oRead("el_hc")
                Pantalla_datos.Combo_tip_carril.Text = oRead("tip_carril")
                Pantalla_datos.Text_ancho_carril.Text = oRead("ancho_carril")
            End If
        End While
        '//
        '// cerrar la base de datos
        '//
        oRead.Close()
        oConn.Close()
        '//
        '// Mostrar la pantalla de datos
        '//
        Pantalla_datos.Show()
    End Sub
End Module
