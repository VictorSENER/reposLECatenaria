Imports System.Data.OleDb
Module cargar_lac
    Sub cargar_lac()
        Dim oConn As New OleDbConnection
        Dim oComm As OleDbCommand
        Dim oRead As OleDbDataReader
        'LEER NOMBRE CATENARIA Y CARGAR
        'oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\23370\Escritorio\SiReCa\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn.Open()
        oComm = New OleDbCommand("select * from Datos", oConn)
        oRead = oComm.ExecuteReader

        While oRead.Read

            'El DataReader se situa sobre el registro

            If (Pantalla_principal.nueva_lac = oRead("nombre_cat")) Then

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
                zona_trab_pant = oRead("zona_trab_pant")
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

            End If

        End While

        oRead.Close()
        oConn.Close()

        'LECTURA DE LA TABLA CONDUCTORES Y CABLES

        'oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\23370\Escritorio\SiReCa\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn.Open()
        oComm = New OleDbCommand("select * from Conductores_y_cables", oConn)
        oRead = oComm.ExecuteReader

        While oRead.Read

            If oRead("tip_cyc") = ("HC") And oRead("mat_cyc") = hc Then
                sec_hc_cyc = oRead("sec_cyc")
                diam_hc_cyc = oRead("diam_cyc")
                p_hc_cyc = oRead("p_cyc")
                res_max_hc_cyc = oRead("res_max_cyc")
                coef_dil_hc_cyc = oRead("coef_dil_cyc")
                mod_elast_hc_cyc = oRead("mod_elast_cyc")
                carga_rot_hc_cyc = oRead("carga_rot_cyc")
                norma_hc_cyc = oRead("norma_cyc")
                origen_1_hc_cyc = oRead("origen_1_cyc")
                origen_2_hc_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("SUSTENTADOR") And oRead("mat_cyc") = sust Then
                sec_sust_cyc = oRead("sec_cyc")
                diam_sust_cyc = oRead("diam_cyc")
                p_sust_cyc = oRead("p_cyc")
                res_max_sust_cyc = oRead("res_max_cyc")
                coef_dil_sust_cyc = oRead("coef_dil_cyc")
                mod_elast_sust_cyc = oRead("mod_elast_cyc")
                carga_rot_sust_cyc = oRead("carga_rot_cyc")
                norma_sust_cyc = oRead("norma_cyc")
                origen_1_sust_cyc = oRead("origen_1_cyc")
                origen_2_sust_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("CDPA") And oRead("mat_cyc") = cdpa Then
                sec_cdpa_cyc = oRead("sec_cyc")
                diam_cdpa_cyc = oRead("diam_cyc")
                p_cdpa_cyc = oRead("p_cyc")
                res_max_cdpa_cyc = oRead("res_max_cyc")
                coef_dil_cdpa_cyc = oRead("coef_dil_cyc")
                mod_elast_cdpa_cyc = oRead("mod_elast_cyc")
                carga_rot_cdpa_cyc = oRead("carga_rot_cyc")
                norma_cdpa_cyc = oRead("norma_cyc")
                origen_1_cdpa_cyc = oRead("origen_1_cyc")
                origen_2_cdpa_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("PUNTO FIJO") And oRead("mat_cyc") = pto_fijo Then
                sec_pto_fijo_cyc = oRead("sec_cyc")
                diam_pto_fijo_cyc = oRead("diam_cyc")
                p_pto_fijo_cyc = oRead("p_cyc")
                res_max_pto_fijo_cyc = oRead("res_max_cyc")
                coef_dil_pto_fijo_cyc = oRead("coef_dil_cyc")
                mod_elast_pto_fijo_cyc = oRead("mod_elast_cyc")
                carga_rot_pto_fijo_cyc = oRead("carga_rot_cyc")
                norma_pto_fijo_cyc = oRead("norma_cyc")
                origen_1_pto_fijo_cyc = oRead("origen_1_cyc")
                origen_2_pto_fijo_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("FEEDER +") And oRead("mat_cyc") = feed_pos Then
                sec_feed_pos_cyc = oRead("sec_cyc")
                diam_feed_pos_cyc = oRead("diam_cyc")
                p_feed_pos_cyc = oRead("p_cyc")
                res_max_feed_pos_cyc = oRead("res_max_cyc")
                coef_dil_feed_pos_cyc = oRead("coef_dil_cyc")
                mod_elast_feed_pos_cyc = oRead("mod_elast_cyc")
                carga_rot_feed_pos_cyc = oRead("carga_rot_cyc")
                norma_feed_pos_cyc = oRead("norma_cyc")
                origen_1_feed_pos_cyc = oRead("origen_1_cyc")
                origen_2_feed_pos_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("FEEDER -") And oRead("mat_cyc") = feed_neg Then
                sec_feed_neg_cyc = oRead("sec_cyc")
                diam_feed_neg_cyc = oRead("diam_cyc")
                p_feed_neg_cyc = oRead("p_cyc")
                res_max_feed_neg_cyc = oRead("res_max_cyc")
                coef_dil_feed_neg_cyc = oRead("coef_dil_cyc")
                mod_elast_feed_neg_cyc = oRead("mod_elast_cyc")
                carga_rot_feed_neg_cyc = oRead("carga_rot_cyc")
                norma_feed_neg_cyc = oRead("norma_cyc")
                origen_1_feed_neg_cyc = oRead("origen_1_cyc")
                origen_2_feed_neg_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("CDTE") And oRead("mat_cyc") = cdte Then
                sec_cdte_cyc = oRead("sec_cyc")
                diam_cdte_cyc = oRead("diam_cyc")
                p_cdte_cyc = oRead("p_cyc")
                res_max_cdte_cyc = oRead("res_max_cyc")
                coef_dil_cdte_cyc = oRead("coef_dil_cyc")
                mod_elast_cdte_cyc = oRead("mod_elast_cyc")
                carga_rot_cdte_cyc = oRead("carga_rot_cyc")
                norma_cdte_cyc = oRead("norma_cyc")
                origen_1_cdte_cyc = oRead("origen_1_cyc")
                origen_2_cdte_cyc = oRead("origen_2_cyc")
            End If

            If oRead("tip_cyc") = ("PENDOLA") And oRead("mat_cyc") = pend Then
                sec_pend_cyc = oRead("sec_cyc")
                diam_pend_cyc = oRead("diam_cyc")
                p_pend_cyc = oRead("p_cyc")
                res_max_pend_cyc = oRead("res_max_cyc")
                coef_dil_pend_cyc = oRead("coef_dil_cyc")
                mod_elast_pend_cyc = oRead("mod_elast_cyc")
                carga_rot_pend_cyc = oRead("carga_rot_cyc")
                norma_pend_cyc = oRead("norma_cyc")
                origen_1_pend_cyc = oRead("origen_1_cyc")
                origen_2_pend_cyc = oRead("origen_2_cyc")
            End If

        End While
        oRead.Close()
        oConn.Close()

        'LECTURA DE LA TABLA MACIZOS
        'LECTURA DE LA TABLA POSTES


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
        Pantalla_principal.Label3.Show()
        Pantalla_principal.Button2.Show()
        Pantalla_principal.GroupBox2.Show()



    End Sub
End Module
