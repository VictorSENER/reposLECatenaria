Imports System.Data.OleDb

Public Class Pantalla_datos
    Dim Direct As New DxVBLib.DirectX7

    Dim DirectD As DxVBLib.DirectDraw7

    Dim ScreenWith, ScreenHeight As Integer
    Dim oConn As New OleDbConnection

    'Private Sub Pantalla_datos_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    'ScreenWith = 1280

    'ScreenHeight = 1024

    'DirectD = Direct.DirectDrawCreate("")

    'DirectD.SetDisplayMode(ScreenWith, ScreenHeight, 0, 0, DxVBLib.CONST_DDSDMFLAGS.DDSDM_DEFAULT)

    'End Sub

    Private Sub Pantalla_datos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Carril_Consulta' Puede moverla o quitarla según sea necesario.
        'Me.Carril_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Carril_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_Feeder__TableAdapter.Fill(Me.Base_de_datosDataSet._Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.__Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_Feeder__TableAdapter1.Fill(Me.Base_de_datosDataSet.__Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Macizos_Consulta' Puede moverla o quitarla según sea necesario.
        'Me.Macizos_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Macizos_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Postes_Consulta' Puede moverla o quitarla según sea necesario.
        'Me.Postes_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Postes_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Electrificación_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Electrificación_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Electrificación_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Anclaje' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_AnclajeTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Anclaje)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Pendola' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_PendolaTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Pendola)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_punto_fijo' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_punto_fijoTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_punto_fijo)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_Feeder__TableAdapter.Fill(Me.Base_de_datosDataSet._Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Cable_de_Tierra' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_Cable_de_TierraTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Cable_de_Tierra)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_CDPA' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_CDPATableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_CDPA)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Sustentador' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_SustentadorTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Sustentador)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_HC' Puede moverla o quitarla según sea necesario.
        'Me.Conductor_HCTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_HC)
        Text_nombre_cat.Hide()
        Label2.Hide()

        'oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        'Parametro de Resolucion Deseados


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

    Private Sub Combo_sist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo_sist.SelectedIndexChanged


        If Combo_sist.Text = "1x25 kV" Or Combo_sist.Text = "3.000 Vcc" Then
            Combo_feed_neg.Text = "NO HAY"
            Combo_feed_neg.Enabled = False
            Combo_posicion_feed_neg.Text = "NO HAY"
            Combo_posicion_feed_neg.Enabled = False
            Text_t_feed_neg.Text = "0"
            Text_t_feed_neg.Enabled = False
            Text_n_feed_neg.Text = "0"
            Text_n_feed_neg.Enabled = False
            Combo_aisl_feed_neg.Text = "NO HAY"
            Combo_aisl_feed_neg.Enabled = False
            Text_dist_vert_feed_neg.Text = "0"
            Text_dist_vert_feed_neg.Enabled = False
            Text_dist_horiz_feed_neg.Text = "0"
            Text_dist_horiz_feed_neg.Enabled = False
        Else
            Combo_feed_neg.Enabled = True
            Combo_feed_neg.Text = ""
            Combo_posicion_feed_neg.Enabled = True
            Combo_posicion_feed_neg.Text = ""
            Text_t_feed_neg.Enabled = True
            Text_t_feed_neg.Text = ""
            Text_n_feed_neg.Enabled = True
            Text_n_feed_neg.Text = ""
            Combo_aisl_feed_neg.Enabled = True
            Combo_aisl_feed_neg.Text = ""
            Text_dist_vert_feed_neg.Enabled = True
            Text_dist_vert_feed_neg.Text = ""
            Text_dist_horiz_feed_neg.Enabled = True
            Text_dist_horiz_feed_neg.Text = ""

        End If

    End Sub



    Private Sub Combo_feed_pos_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Combo_feed_pos.TextChanged
        If Combo_posicion_feed_pos.Text = "" Then
            Me.Text_t_feed_pos.Enabled = False
            Me.Text_t_feed_pos.Text = "0"
            Me.Text_n_feed_pos.Enabled = False
            Me.Text_n_feed_pos.Text = "0"
            Me.Combo_posicion_feed_pos.Enabled = False
            Me.Combo_posicion_feed_pos.Text = "NO HAY"
            Me.Combo_aisl_feed_pos.Enabled = False
            Me.Combo_aisl_feed_pos.Text = "NO HAY"
            Me.Text_dist_vert_feed_pos.Enabled = False
            Me.Text_dist_vert_feed_pos.Text = "0"
            Me.Text_dist_horiz_feed_pos.Enabled = False
            Me.Text_dist_horiz_feed_pos.Text = "0"
        Else
            Me.Text_t_feed_pos.Enabled = True
            Me.Text_t_feed_pos.Text = ""
            Me.Text_n_feed_pos.Enabled = True
            Me.Text_n_feed_pos.Text = ""
            Me.Combo_posicion_feed_pos.Enabled = True
            Me.Combo_posicion_feed_pos.Text = ""
            Me.Combo_aisl_feed_pos.Enabled = True
            Me.Combo_aisl_feed_pos.Text = ""
            Me.Text_dist_vert_feed_pos.Enabled = True
            Me.Text_dist_vert_feed_pos.Text = ""
            Me.Text_dist_horiz_feed_pos.Enabled = True
            Me.Text_dist_horiz_feed_pos.Text = ""
        End If
    End Sub
End Class