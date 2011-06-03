Imports System.Data.OleDb

Public Class Pantalla_datos
    Dim Direct As New DxVBLib.DirectX7

    Dim DirectD As DxVBLib.DirectDraw7

    Dim ScreenWith, ScreenHeight As Integer
    Dim oConn As New OleDbConnection

    Private Sub Pantalla_datos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Feeder__TableAdapter.Fill(Me.Base_de_datosDataSet._Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.__Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Feeder__TableAdapter1.Fill(Me.Base_de_datosDataSet.__Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Macizos_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Macizos_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Macizos_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Postes_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Postes_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Postes_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Electrificación_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Electrificación_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Electrificación_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Anclaje' Puede moverla o quitarla según sea necesario.
        Me.Conductor_AnclajeTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Anclaje)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Pendola' Puede moverla o quitarla según sea necesario.
        Me.Conductor_PendolaTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Pendola)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_punto_fijo' Puede moverla o quitarla según sea necesario.
        Me.Conductor_punto_fijoTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_punto_fijo)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Feeder__TableAdapter.Fill(Me.Base_de_datosDataSet._Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Cable_de_Tierra' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Cable_de_TierraTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Cable_de_Tierra)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_CDPA' Puede moverla o quitarla según sea necesario.
        Me.Conductor_CDPATableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_CDPA)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Sustentador' Puede moverla o quitarla según sea necesario.
        Me.Conductor_SustentadorTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Sustentador)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_HC' Puede moverla o quitarla según sea necesario.
        Me.Conductor_HCTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_HC)


        





        Text_nombre_cat.Hide()
        Label2.Hide()



        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        'Parametro de Resolucion Deseados

        ScreenWith = 1280

        ScreenHeight = 1024

        DirectD = Direct.DirectDrawCreate("")

        DirectD.SetDisplayMode(ScreenWith, ScreenHeight, 0, 0, DxVBLib.CONST_DDSDMFLAGS.DDSDM_DEFAULT)
    End Sub

    Private Sub Pantalla_datos_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        oConn = Nothing

    End Sub

    

 
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call nueva_lac.nueva_lac()
    End Sub

  
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub


End Class