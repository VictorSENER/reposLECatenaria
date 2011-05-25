﻿Imports System.Data.OleDb
Public Class Pantalla_principal
    Public nueva_lac As String
    Public ruta_trazado As String
    Public ruta_replanteo As String
    Public ruta_autocad As String
    Public inicio As Long
    Public fin As Long
    Public nombre_excel As String
    Private Sub Pantalla_principal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Nombre_Catenaria' Puede moverla o quitarla según sea necesario.
        Me.Nombre_CatenariaTableAdapter.Fill(Me.Base_de_datosDataSet.Nombre_Catenaria)


    End Sub
    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        Me.TextBox1.Hide()
        Me.Label1.Hide()
        Me.ComboBox1.Show()
        Me.Button1.Show()
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
                Pantalla_aviso.Show()
                Me.Label3.Show()
                Me.Button2.Show()
                Me.GroupBox2.Show()
                'Me.Label2.Hide()
                'Me.ComboBox1.Hide()
                'Me.Button1.Hide()
                'Me.RadioButton1.Hide()
                'Me.RadioButton2.Hide()
                'Me.GroupBox1.Text = "Datos de catenaria introducidos"
                'Me.GroupBox2.ForeColor = Color.Green
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

                oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
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
                    Pantalla_datos.Show()
                    Me.Label3.Show()
                    Me.Button2.Show()
                    Me.GroupBox2.Show()
                    'Me.Label1.Hide()
                    'Me.TextBox1.Hide()
                    'Me.Button1.Hide()
                    'Me.RadioButton1.Hide()
                    'Me.RadioButton2.Hide()
                    'Me.GroupBox1.Text = "Datos de catenaria introducidos"
                    'Me.GroupBox2.ForeColor = Color.Green
                End If
            End If
        End If
    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Me.Label2.Hide()
        Me.ComboBox1.Hide()
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
        Close()
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
        If TextBox2.Text > TextBox3.Text Then
            MsgBox("PK final debe ser mayor al PK inicial", 48)
        End If
        If Me.TextBox2.Text <> "" And Me.TextBox3.Text <> "" And Me.TextBox4.Text <> "" And Not IsNothing(ruta_replanteo) _
        And TextBox2.Text < TextBox3.Text Then

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


End Class