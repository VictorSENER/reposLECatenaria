'//
'// Rutina destinada a comunicarse con el Excel y activar su programación
'//
Module run
    '//
    '// variables publicas para tabla de replanteo
    '//
    Public caso As String
    Public uno As Integer
    Public nombre_cat As String, sist As String, al As String, alt_nom As Double, alt_min As Double, alt_max As Long, alt_cat As Double, dist_va_max As Long, dist_max_canton As Double, va_max As Double, va_max_sm As Double, va_max_sla As Double, va_max_tunel As Double, inc_norm_va As Double, inc_max_alt_hc As Double, n_min_va_sm As Double, n_min_va_sla As Double, ancho_via As Double, d_max_re As Double, d_max_cu As Double, r_re As Double, d_max_ad As Double, el_max_pant As Double, vw As Double, fl_max_centro_va As Double, dist_carril_poste As Double, dist_base_poste_pmr As Double, dist_elect_sm As Double, dist_elect_sla As Double, l_zc_max As Double, l_zc_min As Double, l_zn As Double, r_min_traz As Double, hc As String, sust As String, cdpa As String, cdte As String, feed_pos As String, feed_neg As String, pto_fijo As String, pend As String, anc As String, posicion_feed_neg As String, n_hc As Long, n_cdpa As Long, n_feed_pos As Long, n_feed_neg As Long, t_hc As Double, t_sust As Double, t_cdpa As Double, t_feed_pos As Double, t_feed_neg As Double, t_pto_fijo As Double, adm_lin_poste As String, tip_poste As String, num_poste As String, adm_lin_mac As String, tip_mac As String, tubo_men As String, tubo_tir As String, cola_anc As String, aisl_feed_pos As String, aisl_feed_neg As String, dist_ap_prim_pend As Long, dist_prim_seg_pend As Long, dist_max_pend As Long, idioma As String
    Public sec_hc As Double, diam_hc As String, p_hc As Double, res_max_hc As Double, coef_dil_hc As String, mod_elast_hc As Double, carga_rot_hc As Double, norma_hc As String, origen_1_hc As String, origen_2_hc As String
    Public sec_sust As Double, diam_sust As String, p_sust As Double, res_max_sust As Double, coef_dil_sust As String, mod_elast_sust As Double, carga_rot_sust As Double, norma_sust As String, origen_1_sust As String, origen_2_sust As String
    Public sec_cdpa As Double, diam_cdpa As String, p_cdpa As Double, res_max_cdpa As Double, coef_dil_cdpa As String, mod_elast_cdpa As Double, carga_rot_cdpa As Double, norma_cdpa As String, origen_1_cdpa As String, origen_2_cdpa As String
    Public sec_pto_fijo As Double, diam_pto_fijo As String, p_pto_fijo, res_max_pto_fijo As Double, coef_dil_pto_fijo As String, mod_elast_pto_fijo As Double, carga_rot_pto_fijo As Double, norma_pto_fijo As String, origen_1_pto_fijo As String, origen_2_pto_fijo As String
    Public sec_feed_pos As Double, diam_feed_pos As String, p_feed_pos As Double, res_max_feed_pos As Double, coef_dil_feed_pos As String, mod_elast_feed_pos As Double, carga_rot_feed_pos As Double, norma_feed_pos As String, origen_1_feed_pos As String, origen_2_feed_pos As String
    Public sec_feed_neg As Double, diam_feed_neg As String, p_feed_neg As Double, res_max_feed_neg As Double, coef_dil_feed_neg As String, mod_elast_feed_neg As Double, carga_rot_feed_neg As Double, norma_feed_neg As String, origen_1_feed_neg As String, origen_2_feed_neg As String
    Public sec_cdte As Double, diam_cdte As String, p_cdte As Double, res_max_cdte As Double, coef_dil_cdte As String, mod_elast_cdte As Double, carga_rot_cdte As Double, norma_cdte As String, origen_1_cdte As String, origen_2_cdte As String
    Public sec_pend As Double, diam_pend As String, p_pend As Double, res_max_pend As Double, coef_dil_pend As String, mod_elast_pend As Double, carga_rot_pend As Double, norma_pend As String, origen_1_pend As String, origen_2_pend As String
    Public dist_vert_feed_pos As Double, dist_horiz_feed_pos As Double, dist_vert_feed_neg As Double, dist_horiz_feed_neg As Double, dist_vert_cdpa As Double, dist_horiz_cdpa As Double, dist_horiz_equip_t As Double, dist_horiz_equip_comp As Double, dist_vert_hc_anc As Double, dist_vert_sust_anc As Double, sep_hc As Double, p_medio_equip_t As Double, p_medio_equip_comp As Double, alt_cat_se_sm_el As Double, alt_cat_e_sm As Double, alt_cat_se_sla_el As Double, alt_cat_e_sla As Double, alt_cat_se_ag_el As Double, alt_cat_e_ag As Double, alt_cat_se_zn_el As Double, alt_cat_e_zn As Double, el_hc As Double, tip_carril As String, ancho_carril As Double
    Public inicio As Double, fin As Double, start As Long
    Public objExcel As Microsoft.Office.Interop.Excel.Application
    Public xLibro As Microsoft.Office.Interop.Excel.Workbook
    Public datos_trazado(1000, 9) As Short
    Public c As Long, h As Long, w As Long, k As Long, a As Long, b As Long
    Public tiempo As System.Int32()
    Public fuerza_d(31) As Double
    'Public WithEvents objCAD As Autodesk.AutoCAD.Interop.Common.AcadRegisteredApplication
    Public fuerza_s(2) As String
    Public momento(44) As Double
    Public Sub run_excel(ByVal inicio, ByVal fin, ByVal ruta_replanteo, ByVal nombre_excel, ByVal ruta_trazado)
        '//
        '// generar un objeto excel
        '//
        objExcel = New Microsoft.Office.Interop.Excel.Application
        'objExcel.Visible = True
        '//
        '// abrir el excel de trazado
        '//
        xLibro = objExcel.Workbooks.Open(ruta_trazado)
        '//
        '// cargar y activar la rutina formato.lenguaje
        '//
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\formato.txt")
        objExcel.Run("formato.lenguaje", idioma)
        '//
        '// añadir hojas para trabajar con el excel
        '//
        objExcel.Worksheets.Add(Before:=objExcel.Worksheets(1))
        objExcel.Worksheets.Add(Before:=objExcel.Worksheets(2))
        objExcel.Worksheets.Add(After:=objExcel.Worksheets(6))
        objExcel.Worksheets.Add(After:=objExcel.Worksheets(7))
        objExcel.Worksheets.Add(After:=objExcel.Worksheets(8))
        '//
        '// cargar los modulos desde los doc's TXT
        '//
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\principal.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\punto_singular.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\altura.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\cantonamiento.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\cad.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\descentramiento.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\num_postes.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\pk_real.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\singular.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\radio.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\regulacion.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\revision.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\vano.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\comentarios.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\tabla_vanos.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\momento.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\cargar.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\eleccion.txt")
        objExcel.VBE.ActiveVBProject.VBComponents.Import("C:\Documents and Settings\23370\Escritorio\SiReCa\Archivos.bas\pendolado.txt")
        '//
        '// añadir librerias
        '//
        objExcel.VBE.ActiveVBProject.References.AddFromFile("C:\Documents and Settings\23370\Escritorio\SiReCa\msado15.dll")
        objExcel.VBE.ActiveVBProject.References.AddFromFile("C:\Documents and Settings\23370\Escritorio\SiReCa\Acrobat.dll")

        objExcel.VBE.ActiveVBProject.References.AddFromFile("C:\Documents and Settings\23370\Escritorio\SiReCa\acrodist.exe")
        '//
        '// ejecutar rutinas antes del programa principal
        '//
        objExcel.Run("tabla_vanos.tabla_vanos", nombre_cat)

        Pantalla_principal.Button5.Visible = False
        Pantalla_principal.Button4.Visible = False
        Pantalla_principal.TextBox2.Visible = False
        Pantalla_principal.TextBox3.Visible = False
        Pantalla_principal.TextBox4.Visible = False
        Pantalla_principal.Label4.Visible = False
        Pantalla_principal.Label5.Visible = False
        Pantalla_principal.Label6.Visible = False
        Pantalla_principal.Label8.Visible = False
        '//
        '// actualizar la barra de progreso
        '//
        With Pantalla_principal.ProgressBar1
            .Maximum = fin + 60
            .Minimum = inicio
            .Value = inicio
        End With
        '//
        '// actualizar la apariencia de la plantalla
        '//
        Pantalla_principal.ProgressBar1.Visible = True
        Pantalla_principal.Label10.Visible = True
        Pantalla_principal.ProgressBar2.Visible = True
        Pantalla_principal.Label11.Visible = True
        With Pantalla_principal.ProgressBar2
            If Pantalla_principal.CheckBox9.Checked = True Then
                .Maximum = 14
            Else
                .Maximum = 13
            End If
            .Minimum = 0
            .Value = 0
        End With
        Pantalla_principal.ProgressBar2.Visible = True
        Pantalla_principal.Refresh()
        '//
        '// ejecutar el programa en excel e inicializar contadores
        '//
        a = 3
        b = 3
        c = 3
        h = 10
        k = 3
        w = 1
        '//
        '// formato del contenido de las celdas
        '//
        xLibro.Worksheets(1).Columns(1).NumberFormat = "@"
        xLibro.Worksheets(1).range("A1", "AA10001").ColumnWidth = 16
        xLibro.Worksheets(1).Range("A1", "AA10001").RowHeight = 14
        '//
        '// activación del programa principal de excel
        '//
        tiempo = objExcel.Run("principal.principal", inicio, h, w, k, a, b, c, r_re, _
                              dist_va_max, inc_norm_va, va_max_tunel, va_max, dist_max_canton, _
                              va_max_sm)
        '//
        '// rutina del programa principal y actualización de barra de estado
        '//
        While tiempo(7) < fin
            tiempo = objExcel.Run("principal.principal", tiempo(0), tiempo(1), tiempo(2), tiempo(3), tiempo(4), tiempo(5), tiempo(6), _
                              r_re, dist_va_max, inc_norm_va, va_max_tunel, va_max, dist_max_canton, va_max_sm)
            Pantalla_principal.ProgressBar1.Value = tiempo(7)
        End While
        '//
        '// activación del resto de rutinas y actualización de la barra de estado
        '//
        Pantalla_principal.Label11.Text = "Módulo conversión de PK"
        Pantalla_principal.ProgressBar2.Value = 1
        Pantalla_principal.Refresh()
        objExcel.Run("pk_real.convertir_LT")
        Pantalla_principal.Label11.Text = "Módulo numeración postes"
        Pantalla_principal.ProgressBar2.Value = 2
        Pantalla_principal.Refresh()
        objExcel.Run("num_postes.postes", nombre_cat)
        Pantalla_principal.Label11.Text = "Módulo altura"
        Pantalla_principal.ProgressBar2.Value = 3
        Pantalla_principal.Refresh()
        objExcel.Run("altura.altura", nombre_cat)
        Pantalla_principal.Label11.Text = "Módulo esfuerzos"
        Pantalla_principal.ProgressBar2.Value = 4
        Pantalla_principal.Refresh()
        objExcel.Run("cad.esfuerzo")
        Pantalla_principal.Label11.Text = "Módulo descentramiento"
        Pantalla_principal.ProgressBar2.Value = 5
        Pantalla_principal.Refresh()
        objExcel.Run("cantonamiento.canton_final", nombre_cat, fin)         ' distribución de los cantones de catenaria
        objExcel.Run("descentramiento.desc")
        Pantalla_principal.Label11.Text = "Módulo posicion"
        Pantalla_principal.ProgressBar2.Value = 6
        Pantalla_principal.Refresh()
        objExcel.Run("cad.posicion")
        Pantalla_principal.Label11.Text = "Módulo comentarios"
        Pantalla_principal.ProgressBar2.Value = 7
        Pantalla_principal.Refresh()
        objExcel.Run("comentarios.comentarios")
        Pantalla_principal.Label11.Text = "Módulo momento"
        Pantalla_principal.ProgressBar2.Value = 8
        Pantalla_principal.Refresh()
        objExcel.Run("momento.momento", nombre_cat)
        Pantalla_principal.Label11.Text = "Módulo formato"
        Pantalla_principal.ProgressBar2.Value = 9
        Pantalla_principal.Refresh()
        objExcel.Run("formato.formato", idioma)
        Pantalla_principal.Label11.Text = "Módulo elección postes"
        Pantalla_principal.ProgressBar2.Value = 10
        Pantalla_principal.Refresh()
        objExcel.Run("eleccion.postes", nombre_cat)
        Pantalla_principal.Label11.Text = "Módulo elección cimentaciones"
        Pantalla_principal.ProgressBar2.Value = 11
        Pantalla_principal.Refresh()
        objExcel.Run("eleccion.cimentaciones", nombre_cat)
        Pantalla_principal.Label11.Text = "Módulo revisión"
        Pantalla_principal.ProgressBar2.Value = 12
        Pantalla_principal.Refresh()
        If Pantalla_principal.CheckBox9.Checked = True Then
            objExcel.Run("pendolado.pendolado", nombre_cat)
            Pantalla_principal.Label11.Text = "Módulo pendolado"
            Pantalla_principal.ProgressBar2.Value = 13
            Pantalla_principal.Refresh()
            objExcel.Run("revision.revision")
            Pantalla_principal.ProgressBar2.Value = 14
        Else
            objExcel.Run("revision.revision")
            Pantalla_principal.ProgressBar2.Value = 13
        End If
        '//
        '// borrar los módulos del excel
        '//
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("principal"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("singular"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("altura"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("cantonamiento"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("cad"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("descentramiento"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("num_postes"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("pk_real"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("punto_singular"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("radio"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("regulacion"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("revision"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("vano"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("comentarios"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("formato"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("tabla_vanos"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("cargar"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("eleccion"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("momento"))
        objExcel.VBE.ActiveVBProject.VBComponents.Remove(VBComponent:=objExcel.VBE.ActiveVBProject.VBComponents.Item("pendolado"))
        objExcel.DisplayAlerts = False
        '//
        '// borrar las hojas no presentables
        '//
        xLibro.Worksheets(9).Delete()
        xLibro.Worksheets(8).Delete()
        xLibro.Worksheets(6).Delete()
        xLibro.Worksheets(5).Delete()
        xLibro.Worksheets(4).Delete()
        xLibro.Worksheets(3).Delete()
        xLibro.Worksheets(2).Delete()
        xLibro.Worksheets(1).activate()
        '//
        '// cambiar el nombre de las hojas principales
        '//
        xLibro.Worksheets(1).name = "Replanteo"
        xLibro.Worksheets(2).name = "Errores"
        objExcel.DisplayAlerts = True
        '//
        '// guardar y cerrar el excel y su objeto
        '//
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

