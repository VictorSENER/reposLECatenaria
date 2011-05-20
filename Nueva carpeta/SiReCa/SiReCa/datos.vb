
Module datos
    ' variables publicas para tabla de replanteo
    Public caso As String, tip_mac As String, tip_poste As String, num_poste As String
    Public uno As Integer
    Public alt_nom As Double, alt_max As Double, alt_min As Long, dist_va_max As Long
    Public canton_max As Long, va_max_comp As Long, va_max_aire As Long, va_max_tunel As Long
    Public inc_va As Long, dist_carril_poste As Long, va_max As Long, inc_alt_hc As Long
    Public inicio As Double, fin As Double, start As Long, l_canton As Long
    Public radio_recta As Long, alt_cat As Long, va_max_sec_comp As Long
    Public va_max_sec_aire As Long, fallo As Long
    Public cadena As String
    Public n_canton2 As Long
    Public objExcel As Microsoft.Office.Interop.Excel.Application
    Public xLibro As Microsoft.Office.Interop.Excel.Workbook
    Public ws1 As New Microsoft.Office.Interop.Excel.Worksheet
    Public ws2 As New Microsoft.Office.Interop.Excel.Worksheet
    Public ws3 As New Microsoft.Office.Interop.Excel.Worksheet
    Public ws4 As New Microsoft.Office.Interop.Excel.Worksheet
    Public ws5 As New Microsoft.Office.Interop.Excel.Worksheet
    Public ws6 As New Microsoft.Office.Interop.Excel.Worksheet
    Public ws7 As New Microsoft.Office.Interop.Excel.Worksheet
    Public datos_trazado(1000, 9) As Short
    Public c As Long, h As Long, w As Long, k As Long, a As Long, b As Long
    Public tiempo As System.Int32()

    Public Sub datos_excel(ByVal inicio, ByVal fin, ByVal ruta_replanteo, ByVal nombre_excel, ByVal ruta_trazado)



        'generar un objeto excel
        objExcel = New Microsoft.Office.Interop.Excel.Application
        'cargar las hojas del trazado
        xLibro = objExcel.Workbooks.Open(ruta_trazado)
        'xLibro = objExcel_tra.Workbooks.Open("C:\Documents and Settings\23370\Escritorio\trazado.xlsx")

        'objExcel.Workbooks.Add()
        'objExcel.Visible = True
        'objExcel.Workbooks.Open("C:\Documents and Settings\23370\Escritorio\trazado.xlsx")
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

End Module

