Module buscar
    Function buscar_archivo(ByVal tipo As String) As String
        Dim myStream As IO.Stream = Nothing
        Dim openFileDialog1 As New OpenFileDialog()

        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = tipo
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                myStream = openFileDialog1.OpenFile()
                If (myStream IsNot Nothing) Then
                    ' Insert code to read the stream here.
                    buscar_archivo = DirectCast(myStream, System.IO.FileStream).Name
                End If
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            Finally
                ' Check this again, since we need to make sure we didn't throw an exception on open.
                If (myStream IsNot Nothing) Then
                    myStream.Close()
                End If
            End Try
        End If
    End Function

    Function buscar_carpeta(Optional ByVal Titulo As String = "...Seleccione una carpeta ", _
                                Optional ByVal Path_Inicial As Object = "") As String

        Try

            Dim objShell As Object
            Dim objFolder As Object
            Dim o_Carpeta As Object

            ' Nuevo objeto Shell.Application
            objShell = CreateObject("Shell.Application")

            Try
                'Abre el cuadro de diálogo para seleccionar
                objFolder = objShell.BrowseForFolder( _
                                        0, _
                                        Titulo, _
                                        0, _
                                        Path_Inicial)

                ' Devuelve solo el nombre de carpeta
                o_Carpeta = objFolder.Self

                ' Devuelve la ruta completa seleccionada en el diálogo
                Buscar_Carpeta = o_Carpeta.Path
            Catch ex As Exception
            End Try
            'Error
        Catch ex As Exception
            MsgBox(Err.Description, vbCritical)
            Buscar_Carpeta = vbNullString
        End Try
    End Function
End Module