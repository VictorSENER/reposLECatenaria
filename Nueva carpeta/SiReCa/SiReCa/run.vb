
Module run
    ' variables publicas para tabla de replanteo
    Public caso As String
    Public uno As Integer
    Public nombre_cat As String, sist As String, al As String, alt_nom As Double, alt_min As Double, alt_max As Long, alt_cat As Double, dist_va_max As Long, dist_max_canton As Double, va_max As Double, va_max_sm As Double, va_max_sla As Double, va_max_tunel As Double, inc_norm_va As Double, inc_max_alt_hc As Double, n_min_va_sm As Double, n_min_va_sla As Double, ancho_via As Double, d_max_re As Double, d_max_cu As Double, r_re As Double, zona_trab_pant As Double, el_max_pant As Double, vw As Double, fl_max_centro_va As Double, dist_carril_poste As Double, dist_base_poste_pmr As Double, dist_elect_sm As Double, dist_elect_sla As Double, l_zc_max As Double, l_zc_min As Double, l_zn As Double, r_min_traz As Double, hc As String, sust As String, cdpa As String, cdte As String, feed_pos As String, feed_neg As String, pto_fijo As String, pend As String, anc As String, posicion_feed_pos As String, posicion_feed_neg As String, n_hc As Long, n_cdpa As Long, n_feed_pos As Long, n_feed_neg As Long, t_hc As Double, t_sust As Double, t_cdpa As Double, t_feed_pos As Double, t_feed_neg As Double, t_pto_fijo As Double, adm_lin_poste As String, tip_poste As String, num_poste As String, adm_lin_mac As String, tip_mac As String, tubo_men As String, tubo_tir As String, cola_anc As String, aisl_feed_pos As String, aisl_feed_neg As String, dist_ap_prim_pend As Long, dist_prim_seg_pend As Long, dist_max_pend As Long
    Public sec_hc_cyc As Double, diam_hc_cyc As String, p_hc_cyc As Double, res_max_hc_cyc As Double, coef_dil_hc_cyc As String, mod_elast_hc_cyc As Double, carga_rot_hc_cyc As Double, norma_hc_cyc As String, origen_1_hc_cyc As String, origen_2_hc_cyc As String
    Public sec_sust_cyc As Double, diam_sust_cyc As String, p_sust_cyc As Double, res_max_sust_cyc As Double, coef_dil_sust_cyc As String, mod_elast_sust_cyc As Double, carga_rot_sust_cyc As Double, norma_sust_cyc As String, origen_1_sust_cyc As String, origen_2_sust_cyc As String
    Public sec_cdpa_cyc As Double, diam_cdpa_cyc As String, p_cdpa_cyc As Double, res_max_cdpa_cyc As Double, coef_dil_cdpa_cyc As String, mod_elast_cdpa_cyc As Double, carga_rot_cdpa_cyc As Double, norma_cdpa_cyc As String, origen_1_cdpa_cyc As String, origen_2_cdpa_cyc As String
    Public sec_pto_fijo_cyc As Double, diam_pto_fijo_cyc As String, p_pto_fijo_cyc, res_max_pto_fijo_cyc As Double, coef_dil_pto_fijo_cyc As String, mod_elast_pto_fijo_cyc As Double, carga_rot_pto_fijo_cyc As Double, norma_pto_fijo_cyc As String, origen_1_pto_fijo_cyc As String, origen_2_pto_fijo_cyc As String
    Public sec_feed_pos_cyc As Double, diam_feed_pos_cyc As String, p_feed_pos_cyc As Double, res_max_feed_pos_cyc As Double, coef_dil_feed_pos_cyc As String, mod_elast_feed_pos_cyc As Double, carga_rot_feed_pos_cyc As Double, norma_feed_pos_cyc As String, origen_1_feed_pos_cyc As String, origen_2_feed_pos_cyc As String
    Public sec_feed_neg_cyc As Double, diam_feed_neg_cyc As String, p_feed_neg_cyc As Double, res_max_feed_neg_cyc As Double, coef_dil_feed_neg_cyc As String, mod_elast_feed_neg_cyc As Double, carga_rot_feed_neg_cyc As Double, norma_feed_neg_cyc As String, origen_1_feed_neg_cyc As String, origen_2_feed_neg_cyc As String
    Public sec_cdte_cyc As Double, diam_cdte_cyc As String, p_cdte_cyc As Double, res_max_cdte_cyc As Double, coef_dil_cdte_cyc As String, mod_elast_cdte_cyc As Double, carga_rot_cdte_cyc As Double, norma_cdte_cyc As String, origen_1_cdte_cyc As String, origen_2_cdte_cyc As String
    Public sec_pend_cyc As Double, diam_pend_cyc As String, p_pend_cyc As Double, res_max_pend_cyc As Double, coef_dil_pend_cyc As String, mod_elast_pend_cyc As Double, carga_rot_pend_cyc As Double, norma_pend_cyc As String, origen_1_pend_cyc As String, origen_2_pend_cyc As String
    Public inicio As Double, fin As Double, start As Long, l_canton As Long
    Public radio_recta As Long, va_max_sec_comp As Long
    Public va_max_sec_aire As Long, fallo As Long
    Public cadena As String
    Public n_canton2 As Long
    Public objExcel As Microsoft.Office.Interop.Excel.Application
    Public xLibro As Microsoft.Office.Interop.Excel.Workbook
    Public datos_trazado(1000, 9) As Short
    Public c As Long, h As Long, w As Long, k As Long, a As Long, b As Long
    Public tiempo As System.Int32()
    'Public WithEvents objCAD As Autodesk.AutoCAD.Interop.Common.AcadRegisteredApplication

    Public Sub run_excel(ByVal inicio, ByVal fin, ByVal ruta_replanteo, ByVal nombre_excel, ByVal ruta_trazado)
        'generar un objeto excel
        objExcel = New Microsoft.Office.Interop.Excel.Application
        'cargar las hojas del trazado
        xLibro = objExcel.Workbooks.Open(ruta_trazado)

        'objExcel.Workbooks.Add()
        'objExcel.Visible = True
        objExcel.Worksheets.Add(Before:=objExcel.Worksheets(1))
        objExcel.Worksheets.Add(After:=objExcel.Worksheets(6))
        'cargar los modulos
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\principal.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\aguja.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\altura.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\cantonamiento.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\cad.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\datos.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\descentramiento.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\num_postes.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\paso_superior.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\pk_real.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\punto_singular.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\radio.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\regulacion.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\revision.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\vano.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\viaducto.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\comentarios.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\formato.txt")

        'actualizar la barra de progreso
        Pantalla_principal.Button5.Visible = False
        Pantalla_principal.Button4.Visible = False
        Pantalla_principal.TextBox2.Visible = False
        Pantalla_principal.TextBox3.Visible = False
        Pantalla_principal.TextBox4.Visible = False
        Pantalla_principal.Label4.Visible = False
        Pantalla_principal.Label5.Visible = False
        Pantalla_principal.Label6.Visible = False
        Pantalla_principal.Label8.Visible = False

        With Pantalla_principal.ProgressBar1
            .Maximum = fin + 60
            .Minimum = inicio
            .Value = inicio
        End With
        Pantalla_principal.ProgressBar1.Visible = True
        Pantalla_principal.Label10.Visible = True
        Pantalla_principal.ProgressBar2.Visible = True
        Pantalla_principal.Label11.Visible = True
        With Pantalla_principal.ProgressBar2
            .Maximum = 10
            .Minimum = 0
            .Value = 0
        End With
        Pantalla_principal.ProgressBar2.Visible = True
        Pantalla_principal.Refresh()
        'ejecutar el programa en excel y actualizar contadores
        a = 3
        b = 3
        c = 3
        h = 10
        k = 3
        w = 1

        tiempo = objExcel.Run("principal.principal", inicio, h, w, k, a, b, c)
        'objExcel.ActiveWorkbook.SaveAs(ruta_replanteo & "\" & nombre_excel, 52)



        'eliminar modulos
        While tiempo(7) < fin
            tiempo = objExcel.Run("principal.principal", tiempo(0), tiempo(1), tiempo(2), tiempo(3), tiempo(4), tiempo(5), tiempo(6))
            Pantalla_principal.ProgressBar1.Value = tiempo(7)

        End While
        Pantalla_principal.Refresh()
        Pantalla_principal.ProgressBar2.Value = 1
        objExcel.Run("formato.formato", fin)
        Pantalla_principal.ProgressBar2.Value = 2
        objExcel.Run("pk_real.convertir_LT", fin)
        Pantalla_principal.ProgressBar2.Value = 3
        objExcel.Run("num_postes.postes", fin)
        Pantalla_principal.ProgressBar2.Value = 4
        objExcel.Run("altura.altura", fin)
        Pantalla_principal.ProgressBar2.Value = 5
        objExcel.Run("cad.esfuerzo", fin)
        Pantalla_principal.ProgressBar2.Value = 6
        'objExcel.Run("canton")                                            ' distribución de los cantones de catenaria
        objExcel.Run("descentramiento.desc", fin)
        Pantalla_principal.ProgressBar2.Value = 7
        objExcel.Run("cad.posicion", fin)
        Pantalla_principal.ProgressBar2.Value = 8
        objExcel.Run("comentarios.comentarios", fin)
        Pantalla_principal.ProgressBar2.Value = 9
        'objExcel.Run("im_pend(fin)")
        objExcel.Run("revision.revision")
        Pantalla_principal.ProgressBar2.Value = 10
        'borrar los módulos
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("principal"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("aguja"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("altura"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("cantonamiento"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("cad"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("datos"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("descentramiento"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("num_postes"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("paso_superior"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("pk_real"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("punto_singular"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("radio"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("regulacion"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("revision"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("vano"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("viaducto"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("comentarios"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("formato"))
        objExcel.DisplayAlerts = False
        xLibro.Worksheets(6).Delete()
        xLibro.Worksheets(5).Delete()
        xLibro.Worksheets(4).Delete()
        xLibro.Worksheets(3).Delete()
        xLibro.Worksheets(2).Delete()
        objExcel.DisplayAlerts = True
        objExcel.ActiveWorkbook.SaveAs(ruta_replanteo & "\" & nombre_excel, 56)
        xLibro.Close()
        objExcel.Quit()
        xLibro = Nothing
        objExcel = Nothing

    End Sub
    'Sub run_autocad(ByVal ruta_autocad)
    'objCAD = New Autodesk.AutoCAD.Interop.AcadApplication
    'objCAD.Visible = True
    'objCAD.Application.Documents.Open(ruta_autocad)
    'objCAD.VBE.ActiveVBProject.VBComponents.Import()
    'End Sub

End Module

