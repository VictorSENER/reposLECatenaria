Imports System.Data.OleDb
Module nueva_lac
    Sub nueva_lac()
        Dim oConn As New OleDbConnection
        Dim oComm As OleDbCommand
        Dim oComm2 As OleDbCommand
        Dim oRead As OleDbDataReader
        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        'oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\23370\Escritorio\SiReCa\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        'INTRODUCIR NUEVA CATENARIA FALTA VER QUE EL NOMBRE ESCRITO NO COINCIDA

        oComm = New OleDbCommand("insert into Datos(nombre_cat, sist, al, alt_nom, alt_min, alt_max, alt_cat, dist_max_va, dist_max_canton, va_max, va_max_sm, va_max_sla, va_max_tunel, inc_norm_va, inc_max_alt_hc, n_min_va_sm, n_min_va_sla, ancho_via, d_max_re, d_max_cu, r_re, d_max_ad, el_max_pant, vw, fl_max_centro_va, dist_carril_poste, dist_base_poste_pmr, dist_elect_sm, dist_elect_sla, l_zc_max, l_zc_min, l_zn, r_min_traz, hc, sust, cdpa, cdte, feed_pos, feed_neg, pto_fijo, pend, anc, posicion_feed_pos, posicion_feed_neg, n_hc, n_cdpa, n_feed_pos, n_feed_neg, t_hc, t_sust, t_cdpa, t_feed_pos, t_feed_neg, t_pto_fijo, adm_lin_poste, tip_poste, num_poste, adm_lin_mac, tip_mac, tubo_men, tubo_tir, cola_anc, aisl_feed_pos, aisl_feed_neg, dist_ap_prim_pend, dist_prim_seg_pend, dist_max_pend, idioma, dist_vert_feed_pos, dist_horiz_feed_pos, dist_vert_feed_neg, dist_horiz_feed_neg, dist_vert_cdpa, dist_horiz_cdpa, dist_horiz_equip_t, dist_horiz_equip_comp, dist_vert_hc_anc, dist_vert_sust_anc, alt_cat_se_sm_el, alt_cat_e_sm, alt_cat_se_sla_el, alt_cat_e_sla, alt_cat_se_ag_el, alt_cat_e_ag, alt_cat_se_zn_el, alt_cat_e_zn, sep_hc, p_medio_equip_t, p_medio_equip_comp, el_hc, tip_carril, ancho_carril) values(@nombre_cat, @sist, @al, @alt_nom, @alt_min, @alt_max, @alt_cat, @dist_max_va, @dist_max_canton, @va_max, @va_max_sm, @va_max_sla, @va_max_tunel, @inc_norm_va, @inc_max_alt_hc, @n_min_va_sm, @n_min_va_sla, @ancho_via, @d_max_re, @d_max_cu, @r_re, @zona_trab_pant, @el_max_pant, @vw, @fl_max_centro_va, @dist_carril_poste, @dist_base_poste_pmr, @dist_elect_sm, @dist_elect_sla, @l_zc_max, @l_zc_min, @l_zn, @r_min_traz, @hc, @sust, @cdpa, @cdte, @feed_pos, @feed_neg, @pto_fijo, @pend, @anc, @posicion_feed_pos, @posicion_feed_neg, @n_hc, @n_cdpa, @n_feed_pos, @n_feed_neg, @t_hc, @t_sust, @t_cdpa, @t_feed_pos, @t_feed_neg, @t_pto_fijo, @adm_lin_poste, @tip_poste, @num_poste, @adm_lin_mac, @tip_mac, @tubo_men, @tubo_tir, @cola_anc, @aisl_feed_pos, @aisl_feed_neg, @dist_ap_prim_pend, @dist_prim_seg_pend, @dist_max_pend, @idioma, @dist_vert_feed_pos, @dist_horiz_feed_pos, @dist_vert_feed_neg, @dist_horiz_feed_neg, @dist_vert_cdpa, @dist_horiz_cdpa, @dist_horiz_equip_t, @dist_horiz_equip_comp, @dist_vert_hc_anc, @dist_vert_sust_anc, @alt_cat_se_sm_el, @alt_cat_e_sm, @alt_cat_se_sla_el, @alt_cat_e_sla, @alt_cat_se_ag_el, @alt_cat_e_ag, @alt_cat_se_zn_el, @alt_cat_e_zn, @sep_hc, @p_medio_equip_t, @p_medio_equip_comp, @el_hc, @tip_carril, @ancho_carril)", oConn)

        oComm2 = New OleDbCommand("select * from Datos", oConn)
        oConn.Open()

        oRead = oComm2.ExecuteReader

        If Pantalla_datos.Text_nombre_cat.Visible Then

            While oRead.Read
                If (oRead("nombre_cat") = Pantalla_datos.Text_nombre_cat.Text) Then
                    Pantalla_datos.Text_nombre_cat.BackColor = Color.Red
                    Pantalla_datos.Label2.ForeColor = Color.Red
                    Pantalla_datos.Text_nombre_cat.SelectAll()
                    MsgBox("NOMBRE REPETIDO", 48)
                    GoTo x
                End If

            End While

            oComm.Parameters.Add(New OleDbParameter("@nombre_cat", OleDbType.VarChar))
            oComm.Parameters("@nombre_cat").Value = Pantalla_datos.Text_nombre_cat.Text

        Else

            oComm.Parameters.Add(New OleDbParameter("@nombre_cat", OleDbType.VarChar))
            oComm.Parameters("@nombre_cat").Value = Pantalla_principal.nueva_lac

        End If


        oComm.Parameters.Add(New OleDbParameter("@sist", OleDbType.VarChar))
        oComm.Parameters("@sist").Value = Pantalla_datos.Combo_sist.Text

        oComm.Parameters.Add(New OleDbParameter("@al", OleDbType.VarChar))
        oComm.Parameters("@al").Value = Pantalla_datos.Text_al.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_nom", OleDbType.VarChar))
        oComm.Parameters("@alt_nom").Value = Pantalla_datos.Text_alt_nom.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_min", OleDbType.VarChar))
        oComm.Parameters("@alt_min").Value = Pantalla_datos.Text_alt_min.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_max", OleDbType.VarChar))
        oComm.Parameters("@alt_max").Value = Pantalla_datos.Text_alt_max.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat", OleDbType.VarChar))
        oComm.Parameters("@alt_cat").Value = Pantalla_datos.Text_alt_cat.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_max_va", OleDbType.VarChar))
        oComm.Parameters("@dist_max_va").Value = Pantalla_datos.Text_dist_max_va.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_max_canton", OleDbType.VarChar))
        oComm.Parameters("@dist_max_canton").Value = Pantalla_datos.Text_dist_max_canton.Text

        oComm.Parameters.Add(New OleDbParameter("@va_max", OleDbType.VarChar))
        oComm.Parameters("@va_max").Value = Pantalla_datos.Text_va_max.Text

        oComm.Parameters.Add(New OleDbParameter("@va_max_sm", OleDbType.VarChar))
        oComm.Parameters("@va_max_sm").Value = Pantalla_datos.Text_va_max_sm.Text

        oComm.Parameters.Add(New OleDbParameter("@va_max_sla", OleDbType.VarChar))
        oComm.Parameters("@va_max_sla").Value = Pantalla_datos.Text_va_max_sla.Text

        oComm.Parameters.Add(New OleDbParameter("@va_max_tunel", OleDbType.VarChar))
        oComm.Parameters("@va_max_tunel").Value = Pantalla_datos.Text_va_max_tunel.Text

        oComm.Parameters.Add(New OleDbParameter("@inc_norm_va", OleDbType.VarChar))
        oComm.Parameters("@inc_norm_va").Value = Pantalla_datos.Text_inc_norm_va.Text

        oComm.Parameters.Add(New OleDbParameter("@inc_max_alt_hc", OleDbType.VarChar))
        oComm.Parameters("@inc_max_alt_hc").Value = Pantalla_datos.Text_inc_max_alt_hc.Text

        oComm.Parameters.Add(New OleDbParameter("@n_min_va_sm", OleDbType.VarChar))
        oComm.Parameters("@n_min_va_sm").Value = Pantalla_datos.Text_n_min_va_sm.Text

        oComm.Parameters.Add(New OleDbParameter("@n_min_va_sla", OleDbType.VarChar))
        oComm.Parameters("@n_min_va_sla").Value = Pantalla_datos.Text_n_min_va_sla.Text

        oComm.Parameters.Add(New OleDbParameter("@ancho_via", OleDbType.VarChar))
        oComm.Parameters("@ancho_via").Value = Pantalla_datos.Text_ancho_via.Text

        oComm.Parameters.Add(New OleDbParameter("@d_max_re", OleDbType.VarChar))
        oComm.Parameters("@d_max_re").Value = Pantalla_datos.Text_d_max_re.Text

        oComm.Parameters.Add(New OleDbParameter("@d_max_cu", OleDbType.VarChar))
        oComm.Parameters("@d_max_cu").Value = Pantalla_datos.Text_d_max_cu.Text

        oComm.Parameters.Add(New OleDbParameter("@r_re", OleDbType.VarChar))
        oComm.Parameters("@r_re").Value = Pantalla_datos.Text_r_re.Text

        oComm.Parameters.Add(New OleDbParameter("@d_max_ad", OleDbType.VarChar))
        oComm.Parameters("@d_max_ad").Value = Pantalla_datos.Text_d_max_ad.Text

        oComm.Parameters.Add(New OleDbParameter("@el_max_pant", OleDbType.VarChar))
        oComm.Parameters("@el_max_pant").Value = Pantalla_datos.Text_el_max_pant.Text

        oComm.Parameters.Add(New OleDbParameter("@vw", OleDbType.VarChar))
        oComm.Parameters("@vw").Value = Pantalla_datos.Text_vw.Text

        oComm.Parameters.Add(New OleDbParameter("@fl_max_centro_va", OleDbType.VarChar))
        oComm.Parameters("@fl_max_centro_va").Value = Pantalla_datos.Text_fl_max_centro_va.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_carril_poste", OleDbType.VarChar))
        oComm.Parameters("@dist_carril_poste").Value = Pantalla_datos.Text_dist_carril_poste.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_base_poste_pmr", OleDbType.VarChar))
        oComm.Parameters("@dist_base_poste_pmr").Value = Pantalla_datos.Text_dist_base_poste_pmr.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_elect_sm", OleDbType.VarChar))
        oComm.Parameters("@dist_elect_sm").Value = Pantalla_datos.Text_dist_elect_sm.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_elect_sla", OleDbType.VarChar))
        oComm.Parameters("@dist_elect_sla").Value = Pantalla_datos.Text_dist_elect_sla.Text

        oComm.Parameters.Add(New OleDbParameter("@l_zc_max", OleDbType.VarChar))
        oComm.Parameters("@l_zc_max").Value = Pantalla_datos.Text_l_zc_max.Text

        oComm.Parameters.Add(New OleDbParameter("@l_zc_min", OleDbType.VarChar))
        oComm.Parameters("@l_zc_min").Value = Pantalla_datos.Text_l_zc_min.Text

        oComm.Parameters.Add(New OleDbParameter("@l_zn", OleDbType.VarChar))
        oComm.Parameters("@l_zn").Value = Pantalla_datos.Text_l_zn.Text

        oComm.Parameters.Add(New OleDbParameter("@r_min_traz", OleDbType.VarChar))
        oComm.Parameters("@r_min_traz").Value = Pantalla_datos.Text_r_min_traz.Text

        oComm.Parameters.Add(New OleDbParameter("@hc", OleDbType.VarChar))
        oComm.Parameters("@hc").Value = Pantalla_datos.Combo_hc.Text

        oComm.Parameters.Add(New OleDbParameter("@sust", OleDbType.VarChar))
        oComm.Parameters("@sust").Value = Pantalla_datos.Combo_sust.Text

        oComm.Parameters.Add(New OleDbParameter("@cdpa", OleDbType.VarChar))
        oComm.Parameters("@cdpa").Value = Pantalla_datos.Combo_cdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@cdte", OleDbType.VarChar))
        oComm.Parameters("@cdte").Value = Pantalla_datos.Combo_cdte.Text

        oComm.Parameters.Add(New OleDbParameter("@feed_pos", OleDbType.VarChar))
        oComm.Parameters("@feed_pos").Value = Pantalla_datos.Combo_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@feed_neg", OleDbType.VarChar))
        oComm.Parameters("@feed_neg").Value = Pantalla_datos.Combo_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@pto_fijo", OleDbType.VarChar))
        oComm.Parameters("@pto_fijo").Value = Pantalla_datos.Combo_pto_fijo.Text

        oComm.Parameters.Add(New OleDbParameter("@pend", OleDbType.VarChar))
        oComm.Parameters("@pend").Value = Pantalla_datos.Combo_pend.Text

        oComm.Parameters.Add(New OleDbParameter("@anc", OleDbType.VarChar))
        oComm.Parameters("@anc").Value = Pantalla_datos.Combo_anc.Text

        oComm.Parameters.Add(New OleDbParameter("@posicion_feed_pos", OleDbType.VarChar))
        oComm.Parameters("@posicion_feed_pos").Value = Pantalla_datos.Combo_posicion_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@posicion_feed_neg", OleDbType.VarChar))
        oComm.Parameters("@posicion_feed_neg").Value = Pantalla_datos.Combo_posicion_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@n_hc", OleDbType.VarChar))
        oComm.Parameters("@n_hc").Value = Pantalla_datos.Text_n_hc.Text

        oComm.Parameters.Add(New OleDbParameter("@n_cdpa", OleDbType.VarChar))
        oComm.Parameters("@n_cdpa").Value = Pantalla_datos.Text_n_cdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@n_feed_pos", OleDbType.VarChar))
        oComm.Parameters("@n_feed_pos").Value = Pantalla_datos.Text_n_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@n_feed_neg", OleDbType.VarChar))
        oComm.Parameters("@n_feed_neg").Value = Pantalla_datos.Text_n_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@t_hc", OleDbType.VarChar))
        oComm.Parameters("@t_hc").Value = Pantalla_datos.Text_t_hc.Text

        oComm.Parameters.Add(New OleDbParameter("@t_sust", OleDbType.VarChar))
        oComm.Parameters("@t_sust").Value = Pantalla_datos.Text_t_sust.Text

        oComm.Parameters.Add(New OleDbParameter("@t_cdpa", OleDbType.VarChar))
        oComm.Parameters("@t_cdpa").Value = Pantalla_datos.Text_t_cdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@t_feed_pos", OleDbType.VarChar))
        oComm.Parameters("@t_feed_pos").Value = Pantalla_datos.Text_t_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@t_feed_neg", OleDbType.VarChar))
        oComm.Parameters("@t_feed_neg").Value = Pantalla_datos.Text_t_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@t_pto_fijo", OleDbType.VarChar))
        oComm.Parameters("@t_pto_fijo").Value = Pantalla_datos.Text_t_pto_fijo.Text

        oComm.Parameters.Add(New OleDbParameter("@adm_lin_poste", OleDbType.VarChar))
        oComm.Parameters("@adm_lin_poste").Value = Pantalla_datos.Combo_adm_lin_poste.Text

        oComm.Parameters.Add(New OleDbParameter("@tip_poste", OleDbType.VarChar))
        oComm.Parameters("@tip_poste").Value = Pantalla_datos.Text_tip_poste.Text

        oComm.Parameters.Add(New OleDbParameter("@num_poste", OleDbType.VarChar))
        oComm.Parameters("@num_poste").Value = Pantalla_datos.Combo_num_poste.Text

        oComm.Parameters.Add(New OleDbParameter("@adm_lin_mac", OleDbType.VarChar))
        oComm.Parameters("@adm_lin_mac").Value = Pantalla_datos.Combo_adm_lin_mac.Text

        oComm.Parameters.Add(New OleDbParameter("@tip_mac", OleDbType.VarChar))
        oComm.Parameters("@tip_mac").Value = Pantalla_datos.Text_tip_mac.Text

        oComm.Parameters.Add(New OleDbParameter("@tubo_men", OleDbType.VarChar))
        oComm.Parameters("@tubo_men").Value = Pantalla_datos.Combo_tubo_men.Text

        oComm.Parameters.Add(New OleDbParameter("@tubo_tir", OleDbType.VarChar))
        oComm.Parameters("@tubo_tir").Value = Pantalla_datos.Combo_tubo_tir.Text

        oComm.Parameters.Add(New OleDbParameter("@cola_anc", OleDbType.VarChar))
        oComm.Parameters("@cola_anc").Value = Pantalla_datos.Combo_cola_anc.Text

        oComm.Parameters.Add(New OleDbParameter("@aisl_feed_pos", OleDbType.VarChar))
        oComm.Parameters("@aisl_feed_pos").Value = Pantalla_datos.Combo_aisl_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@aisl_feed_neg", OleDbType.VarChar))
        oComm.Parameters("@aisl_feed_neg").Value = Pantalla_datos.Combo_aisl_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_ap_prim_pend", OleDbType.VarChar))
        oComm.Parameters("@dist_ap_prim_pend").Value = Pantalla_datos.Text_dist_ap_prim_pend.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_prim_seg_pend", OleDbType.VarChar))
        oComm.Parameters("@dist_prim_seg_pend").Value = Pantalla_datos.Text_dist_prim_seg_pend.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_max_pend", OleDbType.VarChar))
        oComm.Parameters("@dist_max_pend").Value = Pantalla_datos.Text_dist_max_pend.Text

        oComm.Parameters.Add(New OleDbParameter("@idioma", OleDbType.VarChar))
        oComm.Parameters("@idioma").Value = Pantalla_datos.Combo_idioma.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_vert_feed_pos", OleDbType.VarChar))
        oComm.Parameters("@dist_vert_feed_pos").Value = Pantalla_datos.Text_dist_vert_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_horiz_feed_pos", OleDbType.VarChar))
        oComm.Parameters("@dist_horiz_feed_pos").Value = Pantalla_datos.Text_dist_horiz_feed_pos.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_vert_feed_neg", OleDbType.VarChar))
        oComm.Parameters("@dist_vert_feed_neg").Value = Pantalla_datos.Text_dist_vert_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_horiz_feed_neg", OleDbType.VarChar))
        oComm.Parameters("@dist_horiz_feed_neg").Value = Pantalla_datos.Text_dist_horiz_feed_neg.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_vert_cdpa", OleDbType.VarChar))
        oComm.Parameters("@dist_vert_cdpa").Value = Pantalla_datos.Text_dist_vert_cdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_horiz_cdpa", OleDbType.VarChar))
        oComm.Parameters("@dist_horiz_cdpa").Value = Pantalla_datos.Text_dist_horiz_cdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_horiz_equip_t", OleDbType.VarChar))
        oComm.Parameters("@dist_horiz_equip_t").Value = Pantalla_datos.Text_dist_horiz_equip_t.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_horiz_equip_comp", OleDbType.VarChar))
        oComm.Parameters("@dist_horiz_equip_comp").Value = Pantalla_datos.Text_dist_horiz_equip_comp.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_vert_hc_anc", OleDbType.VarChar))
        oComm.Parameters("@dist_vert_hc_anc").Value = Pantalla_datos.Text_dist_vert_hc_anc.Text

        oComm.Parameters.Add(New OleDbParameter("@dist_vert_sust_anc", OleDbType.VarChar))
        oComm.Parameters("@dist_vert_sust_anc").Value = Pantalla_datos.Text_dist_vert_sust_anc.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_se_sm_el", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_se_sm_el").Value = Pantalla_datos.Text_alt_cat_se_sm_el.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_e_sm", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_e_sm").Value = Pantalla_datos.Text_alt_cat_e_sm.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_se_sla_el", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_se_sla_el").Value = Pantalla_datos.Text_alt_cat_se_sla_el.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_e_sla", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_e_sla").Value = Pantalla_datos.Text_alt_cat_e_sla.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_se_ag_el", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_se_ag_el").Value = Pantalla_datos.Text_alt_cat_se_ag_el.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_e_ag", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_e_ag").Value = Pantalla_datos.Text_alt_cat_e_ag.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_se_zn_el", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_se_zn_el").Value = Pantalla_datos.Text_alt_cat_se_zn_el.Text

        oComm.Parameters.Add(New OleDbParameter("@alt_cat_e_zn", OleDbType.VarChar))
        oComm.Parameters("@alt_cat_e_zn").Value = Pantalla_datos.Text_alt_cat_e_zn.Text

        oComm.Parameters.Add(New OleDbParameter("@sep_hc", OleDbType.VarChar))
        oComm.Parameters("@sep_hc").Value = Pantalla_datos.Text_sep_hc.Text

        oComm.Parameters.Add(New OleDbParameter("@p_medio_equip_t", OleDbType.VarChar))
        oComm.Parameters("@p_medio_equip_t").Value = Pantalla_datos.Text_p_medio_equip_t.Text

        oComm.Parameters.Add(New OleDbParameter("@p_medio_equip_comp", OleDbType.VarChar))
        oComm.Parameters("@p_medio_equip_comp").Value = Pantalla_datos.Text_p_medio_equip_comp.Text

        oComm.Parameters.Add(New OleDbParameter("@el_hc", OleDbType.VarChar))
        oComm.Parameters("@el_hc").Value = Pantalla_datos.Text_el_hc.Text

        oComm.Parameters.Add(New OleDbParameter("@tip_carril", OleDbType.VarChar))
        oComm.Parameters("@tip_carril").Value = Pantalla_datos.Combo_tip_carril.Text

        oComm.Parameters.Add(New OleDbParameter("@ancho_carril", OleDbType.VarChar))
        oComm.Parameters("@ancho_carril").Value = Pantalla_datos.Text_ancho_carril.Text

        oComm2.Connection.Close()
        oComm.Connection.Open()
        On Error GoTo mserror
        oComm.ExecuteNonQuery()
        oComm.Connection.Close()
        MsgBox("RESGITRO AÑADIDO")
        cargar_lac.cargar_lac()
        Pantalla_datos.Close()
        Pantalla_principal.Label1.Hide()
        Pantalla_principal.Label2.Hide()
        Pantalla_principal.TextBox1.Hide()
        Pantalla_principal.ComboBox1.Hide()
        Pantalla_principal.Button1.Hide()
        Pantalla_principal.Button8.Hide()
        Pantalla_principal.Button9.Hide()
        Pantalla_principal.RadioButton1.Hide()
        Pantalla_principal.RadioButton2.Hide()
        Pantalla_principal.GroupBox1.Text = "Datos de catenaria introducidos"
        Pantalla_principal.GroupBox2.ForeColor = Color.Green


        Exit Sub
mserror:
        oComm.Connection.Close()
        MsgBox("FALTAN CAMPOS POR COMPLETAR")
x:

    End Sub

End Module
