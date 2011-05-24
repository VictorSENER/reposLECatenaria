Imports System.Data.OleDb
Public Class Pantalla_aviso
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Pantalla_datos.Show()
        Pantalla_datos.Button2.Hide()

        cargar_lac.cargar_lac()

        Pantalla_datos.Combo_sist.Enabled = False
        Pantalla_datos.Combo_sist.Enabled = False
        Pantalla_datos.Text_al.Enabled = False
        Pantalla_datos.Text_alt_nom.Enabled = False
        Pantalla_datos.Text_alt_min.Enabled = False
        Pantalla_datos.Text_alt_max.Enabled = False
        Pantalla_datos.Text_alt_cat.Enabled = False
        Pantalla_datos.Text_dist_max_va.Enabled = False
        Pantalla_datos.Text_dist_max_canton.Enabled = False
        Pantalla_datos.Text_va_max.Enabled = False
        Pantalla_datos.Text_va_max_sm.Enabled = False
        Pantalla_datos.Text_va_max_sla.Enabled = False
        Pantalla_datos.Text_va_max_tunel.Enabled = False
        Pantalla_datos.Text_inc_norm_va.Enabled = False
        Pantalla_datos.Text_inc_max_alt_hc.Enabled = False
        Pantalla_datos.Text_n_min_va_sm.Enabled = False
        Pantalla_datos.Text_n_min_va_sla.Enabled = False
        Pantalla_datos.Text_dist_base_poste_pmr.Enabled = False
        Pantalla_datos.Text_ancho_via.Enabled = False
        Pantalla_datos.Text_d_max_re.Enabled = False
        Pantalla_datos.Text_d_max_re.Enabled = False
        Pantalla_datos.Text_d_max_cu.Enabled = False
        Pantalla_datos.Text_r_re.Enabled = False
        Pantalla_datos.Text_zona_trab_pant.Enabled = False
        Pantalla_datos.Text_el_max_pant.Enabled = False
        Pantalla_datos.Text_vw.Enabled = False
        Pantalla_datos.Text_fl_max_centro_va.Enabled = False
        Pantalla_datos.Text_dist_carril_poste.Enabled = False
        Pantalla_datos.Text_dist_base_poste_pmr.Enabled = False
        Pantalla_datos.Text_dist_elect_sm.Enabled = False
        Pantalla_datos.Text_dist_elect_sla.Enabled = False
        Pantalla_datos.Text_l_zc_max.Enabled = False
        Pantalla_datos.Text_l_zc_min.Enabled = False
        Pantalla_datos.Text_l_zn.Enabled = False
        Pantalla_datos.Combo_hc.Enabled = False
        Pantalla_datos.Combo_sust.Enabled = False
        Pantalla_datos.Combo_cdpa.Enabled = False
        Pantalla_datos.Combo_cdte.Enabled = False
        Pantalla_datos.Combo_feed_pos.Enabled = False
        Pantalla_datos.Combo_feed_neg.Enabled = False
        Pantalla_datos.Combo_pto_fijo.Enabled = False
        Pantalla_datos.Combo_pend.Enabled = False
        Pantalla_datos.Combo_anc.Enabled = False
        Pantalla_datos.Combo_posicion_feed_pos.Enabled = False
        Pantalla_datos.Combo_posicion_feed_neg.Enabled = False
        Pantalla_datos.Text_n_hc.Enabled = False
        Pantalla_datos.Text_n_cdpa.Enabled = False
        Pantalla_datos.Text_n_feed_pos.Enabled = False
        Pantalla_datos.Text_n_feed_neg.Enabled = False
        Pantalla_datos.Text_t_hc.Enabled = False
        Pantalla_datos.Text_t_sust.Enabled = False
        Pantalla_datos.Text_t_cdpa.Enabled = False
        Pantalla_datos.Text_t_feed_pos.Enabled = False
        Pantalla_datos.Text_t_feed_neg.Enabled = False
        Pantalla_datos.Text_t_pto_fijo.Enabled = False
        Pantalla_datos.Combo_adm_lin_poste.Enabled = False
        Pantalla_datos.Text_tip_poste.Enabled = False
        Pantalla_datos.Combo_num_poste.Enabled = False
        Pantalla_datos.Combo_adm_lin_mac.Enabled = False
        Pantalla_datos.Text_tip_mac.Enabled = False
        Pantalla_datos.Combo_tubo_men.Enabled = False
        Pantalla_datos.Combo_tubo_tir.Enabled = False
        Pantalla_datos.Combo_cola_anc.Enabled = False
        Pantalla_datos.Combo_aisl_feed_pos.Enabled = False
        Pantalla_datos.Combo_aisl_feed_neg.Enabled = False
        Pantalla_datos.Text_dist_ap_prim_pend.Enabled = False
        Pantalla_datos.Text_dist_prim_seg_pend.Enabled = False
        Pantalla_datos.Text_dist_max_pend.Enabled = False

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Label3.Show()
        Label4.Show()
        Text_usuario.Show()
        Text_contraseña.Show()
        Button3.Show()

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim oConn As OleDbConnection
        Dim oComm As OleDbCommand
        Dim oRead As OleDbDataReader

        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        oConn.Open()
        oComm = New OleDbCommand("select * from Contraseña", oConn)
        oRead = oComm.ExecuteReader

        While oRead.Read

            If oRead("USUARIO") = Text_usuario.Text And oRead("CONTRASEÑA") = Text_contraseña.Text Then

                cargar_lac.cargar_lac()
                Pantalla_datos.Text_nombre_cat.Show()
                Pantalla_datos.Label2.Show()


            Else

                MsgBox("USUARIO Y CONTRASEÑA INCORRECTOS", 48)

            End If

            Text_usuario.Clear()
            Text_contraseña.Clear()

        End While

        Me.Close()

    End Sub

    Private Sub Pantalla_aviso_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label3.Hide()
        Label4.Hide()
        Text_usuario.Hide()
        Text_contraseña.Hide()
        Button3.Hide()
    End Sub
End Class