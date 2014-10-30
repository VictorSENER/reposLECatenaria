Attribute VB_Name = "txt"
Sub progress(ini_fun As Integer, fin_fun As Integer, tit As String, pk_act As Double, pk_fin As Double)
    On Error Resume Next
    Set text = a_text.CreateTextFile(dir_progress)
    If Err = False Then
        If pk_act > pk_fin Then
            text.Write ini_fun & "/" & fin_fun & "/" & tit & "/" & pk_fin & "/" & pk_fin
        Else
            text.Write ini_fun & "/" & fin_fun & "/" & tit & "/" & pk_act & "/" & pk_fin
        End If
        text.Close
    Else
        Err.Clear
    End If
End Sub
Sub error(aviso As String, tit As String)
    On Error Resume Next
    If Err = False Then
        Set text_error = a_text.OpenTextFile(dir_error, ForAppending)
        text_error.WriteLine aviso & "/" & tit
        text_error.Close
    Else
        Err.Clear
    End If
        
End Sub
