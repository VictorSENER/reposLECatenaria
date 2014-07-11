Imports System.Data.OleDb
Public Class Pantalla_principal
    Public nueva_lac As String
    Public ruta_trazado As String
    Public ruta_replanteo As String
    Public ruta_autocad As String
    Public nombre_excel As String


    Private Sub Pantalla_principal_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated


        Dim Direct As New DxVBLib.DirectX7

        Dim DirectD As DxVBLib.DirectDraw7

        Dim ScreenWith, ScreenHeight As Integer

        'Parametro de Resolucion Deseados

        ScreenWith = 1280

        ScreenHeight = 1024

        DirectD = Direct.DirectDrawCreate("")

        DirectD.SetDisplayMode(ScreenWith, ScreenHeight, 0, 0, DxVBLib.CONST_DDSDMFLAGS.DDSDM_DEFAULT)
    End Sub

    Private Sub Pantalla_principal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Nombre_Catenaria' Puede moverla o quitarla según sea necesario.
        Me.Nombre_CatenariaTableAdapter.Fill(Me.Base_de_datosDataSet.Nombre_Catenaria)

        

    End Sub


    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Me.TextBox1.Hide()
        Me.Label1.Hide()
        Me.ComboBox1.Show()
        Me.Button1.Text = "CARGAR"
        Me.Button1.Show()
        Me.Button8.Show()
        Me.Button9.Show()
        Me.Label2.Show()
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Label1.ForeColor = Color.White
        Me.TextBox1.BackColor = Color.White
        If Me.RadioButton1.Checked = True Then
            If Me.ComboBox1.Text = "" Then
                Me.Label2.ForeColor = Color.Red
                Me.ComboBox1.BackColor = Color.Red
                MsgBox("Rellenar la celda", 48)
            Else
                nueva_lac = ComboBox1.Text
                cargar_lac.cargar_lac()

            End If

        ElseIf Me.RadioButton2.Checked = True Then

            If Me.TextBox1.Text = "" Then
                Me.Label1.ForeColor = Color.Red
                Me.TextBox1.BackColor = Color.Red
                MsgBox("RELLENAR LA CELDA", 48)
            Else
                nueva_lac = TextBox1.Text
                Dim oConn As OleDbConnection
                Dim oComm As OleDbCommand
                Dim oRead As OleDbDataReader

                'oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
                oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\23370\Escritorio\SiReCa\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
                oConn.Open()
                oComm = New OleDbCommand("select * from Datos", oConn)
                oRead = oComm.ExecuteReader

                While oRead.Read
                    If oRead("nombre_cat") = nueva_lac Then
                        Me.Label1.ForeColor = Color.Red
                        Me.TextBox1.BackColor = Color.Red
                        MsgBox("NOMBRE REPETIDO", 48)
                    End If
                End While

                If (Me.Label1.ForeColor = Color.Red) = False Then
                    nueva_lac = TextBox1.Text
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Carril_Consulta' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Carril_ConsultaTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Carril_Consulta)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_Feeder__TableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet._Conductor_Feeder__)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.__Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_Feeder__TableAdapter1.Fill(Pantalla_datos.Base_de_datosDataSet.__Conductor_Feeder__)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Macizos_Consulta' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Macizos_ConsultaTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Macizos_Consulta)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Postes_Consulta' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Postes_ConsultaTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Postes_Consulta)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Electrificación_Consulta' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Electrificación_ConsultaTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Electrificación_Consulta)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Anclaje' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_AnclajeTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_Anclaje)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Pendola' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_PendolaTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_Pendola)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_punto_fijo' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_punto_fijoTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_punto_fijo)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_Feeder__TableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet._Conductor_Feeder__)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Cable_de_Tierra' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_Cable_de_TierraTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_Cable_de_Tierra)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_CDPA' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_CDPATableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_CDPA)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Sustentador' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_SustentadorTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_Sustentador)
                    'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_HC' Puede moverla o quitarla según sea necesario.
                    Pantalla_datos.Conductor_HCTableAdapter.Fill(Pantalla_datos.Base_de_datosDataSet.Conductor_HC)
                    Pantalla_datos.Show()
                    Pantalla_datos.Combo_sist.Text = ""
                    Pantalla_datos.Combo_hc.Text = ""
                    Me.Label3.Show()
                    Me.Button2.Show()
                    Me.GroupBox2.Show()

                End If
            End If
        End If

    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Me.Label2.Hide()
        Me.ComboBox1.Hide()
        Me.Button8.Hide()
        Me.Button9.Hide()
        Me.Button1.Text = "INTRODUCIR"
        Me.Button1.Show()
        Me.TextBox1.Show()
        Me.Label1.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim tipo As String
        Me.Label3.ForeColor = Color.White
        Me.Button2.ForeColor = Color.White
        tipo = "xlsx files (*.xlsx)|*xlsx| xls files (*.xls)|*.xls"
        ruta_trazado = buscar.buscar_archivo(tipo)
        If Not IsNothing(ruta_trazado) Then
            Me.Button2.Hide()
            Me.Label3.Hide()
            Me.Label3.Text = "Archivo encontrado"
            Me.TextBox2.Show()
            Me.TextBox3.Show()
            Me.TextBox4.Show()
            Me.Label4.Show()
            Me.Label5.Show()
            Me.Label6.Show()
            Me.Label8.Show()
            Me.Button4.Show()
            Me.Button5.Show()
            Me.GroupBox2.Text = "Datos del trazado introducidos"
            Me.GroupBox3.ForeColor = Color.Green
        Else
            Me.Label3.ForeColor = Color.Red
            Me.Button2.ForeColor = Color.Red
            Me.Label3.Text = "Archivo no encontrado"
            MsgBox("Archivo no encontrado", 48)
        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        ruta_replanteo = buscar.buscar_carpeta

    End Sub
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim tipo As String
        tipo = "dxf files (*.dxf)|*dxf | dwg files (*.dwg)|*.dwg"
        ruta_autocad = buscar.buscar_archivo(tipo)
        If Not IsNothing(ruta_trazado) Then
            Me.Button2.Hide()
            Me.Label3.Text = "Archivo encontrado"
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'actualizar el progressbar
        Me.TextBox2.BackColor = Color.White
        Me.TextBox3.BackColor = Color.White
        Me.TextBox4.BackColor = Color.White
        Me.Label4.ForeColor = Color.White
        Me.Label5.ForeColor = Color.White
        Me.Label6.ForeColor = Color.White
        Me.Label8.ForeColor = Color.White
        Me.Button4.ForeColor = Color.White

        If Me.TextBox2.Text = "" Or Not IsNumeric(Me.TextBox2.Text) Then
            Me.TextBox2.BackColor = Color.Red
            Me.TextBox2.Text = ""
            Me.Label4.ForeColor = Color.Red
            MsgBox("Introducir un dato numerico", 48)
        End If
        If Me.TextBox3.Text = "" Or Not IsNumeric(Me.TextBox3.Text) Then
            Me.TextBox3.BackColor = Color.Red
            Me.TextBox3.Text = ""
            Me.Label5.ForeColor = Color.Red
            MsgBox("Introducir un dato numerico", 48)
        End If
        If Me.TextBox4.Text = "" Then
            Me.TextBox4.BackColor = Color.Red
            Me.Label6.ForeColor = Color.Red
            MsgBox("Introducir nombre del archivo", 48)
        End If
        If IsNothing(ruta_replanteo) Then
            Me.Label8.ForeColor = Color.Red
            Me.Button4.ForeColor = Color.Red
            MsgBox("Elegir ruta de destino", 48)
        End If
        inicio = Me.TextBox2.Text
        fin = Me.TextBox3.Text
        If inicio > fin Then
            MsgBox("PK final debe ser mayor al PK inicial", 48)
        End If
        If Me.TextBox2.Text <> "" And Me.TextBox3.Text <> "" And Me.TextBox4.Text <> "" And Not IsNothing(ruta_replanteo) _
        And inicio < fin Then

            'recogemos las variables necesarias para excel
            inicio = Me.TextBox2.Text
            fin = Me.TextBox3.Text
            nombre_excel = Me.TextBox4.Text
            'llamar a la rutina de acceso al excel
            Call run.run_excel(inicio, fin, ruta_replanteo, nombre_excel, ruta_trazado)
            Me.GroupBox3.Text = "Replanteo realizado correctamente"
            Me.GroupBox4.ForeColor = Color.Green
            Me.ProgressBar1.Hide()
            Me.ProgressBar2.Hide()
            Me.Label10.Hide()
            Me.Label11.Hide()
            Me.CheckBox1.Show()
            Me.CheckBox2.Show()
            Me.CheckBox3.Show()
            Me.CheckBox4.Show()
            Me.CheckBox5.Show()
            Me.CheckBox6.Show()
            Me.CheckBox7.Show()
            Me.CheckBox8.Show()
            Me.Label7.Show()
            Me.Button6.Show()
            Me.Button7.Show()
        End If

    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        'Call run.run_autocad(ruta_autocad)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        nueva_lac = ComboBox1.Text
        ver_lac.ver_lac()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        nueva_lac = ComboBox1.Text
        Pantalla_aviso.Show()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Me.Close()
    End Sub


End Class