Imports System.Data.OleDb
Public Class Pantalla_aviso

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

                modificar_lac.modificar_lac()
                Me.Close()
                Pantalla_datos.Text_nombre_cat.Show()
                Pantalla_datos.Label2.Show()


            Else

                MsgBox("USUARIO Y CONTRASEÑA INCORRECTOS", 48)

            End If

            Text_usuario.Clear()
            Text_contraseña.Clear()

        End While


    End Sub

    Private Sub Pantalla_aviso_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Direct As New DxVBLib.DirectX7

        Dim DirectD As DxVBLib.DirectDraw7

        Dim ScreenWith, ScreenHeight As Integer

        'Parametro de Resolucion Deseados

        ScreenWith = 1280

        ScreenHeight = 1024

        DirectD = Direct.DirectDrawCreate("")

        DirectD.SetDisplayMode(ScreenWith, ScreenHeight, 0, 0, DxVBLib.CONST_DDSDMFLAGS.DDSDM_DEFAULT)
    End Sub
End Class