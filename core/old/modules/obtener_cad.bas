Attribute VB_Name = "obtener_cad"
Sub polilinea()
Dim PntA As Variant
Dim objPoli As AcadLWPolyline
Dim objPoli2 As AcadLWPolyline
Dim enType As String
Dim entObj As AcadObject
Dim entObj2 As AcadObject
Dim coordenadasPoli As Variant
Dim AcadDoc As Object
Dim i, j As Integer
Dim OffsetObj As Variant
Set AcadDoc = GetObject(, "Autocad.Application").ActiveDocument
AcadDoc.Utility.GetEntity entObj, PntA, vbCrLf & "Seleciona el eje. Debe estar unido: "
enType = entObj.ObjectName
If enType = "AcDbPolyline" Or enType = "AcDb2dPolyline" Then
    Set objPoli = entObj
    coordenadasPoli = objPoli.Coordinates
    j = 0
    poli(j).coordx = coordenadasPoli(LBound(coordenadasPoli))
    poli(j).coordy = coordenadasPoli(LBound(coordenadasPoli) + 1)
    poli(j).dist_acum = Sireca.TextBox1.Text
    For i = LBound(coordenadasPoli) To UBound(coordenadasPoli)
        j = j + 1
        poli(j).coordx = coordenadasPoli(i)
        poli(j).coordy = coordenadasPoli(i + 1)
        poli(j).dist_acum = poli(j - 1).dist_acum + Sqr((poli(j).coordx - poli(j - 1).coordx) ^ 2 + (poli(j).coordy - poli(j - 1).coordy) ^ 2)
        If poli(j).coordx <> poli(j - 1).coordx Then
            If (poli(j).coordx - poli(j - 1).coordx) > 0 Then
                poli(j - 1).angulo_post = (180 / PI) * Atn((poli(j).coordy - poli(j - 1).coordy) / (poli(j).coordx - poli(j - 1).coordx))
            ElseIf (poli(j).coordx - poli(j - 1).coordx) < 0 Then
                poli(j - 1).angulo_post = 180 + (180 / PI) * Atn((poli(j).coordy - poli(j - 1).coordy) / (poli(j).coordx - poli(j - 1).coordx))
            End If
        ElseIf (poli(j).coordy - poli(j - 1).coordy) > 0 Then
            poli(j - 1).angulo_post = 90
        ElseIf (poli(j).coordy - poli(j - 1).coordy) < 0 Then
            poli(j - 1).angulo_post = 270
        End If
        i = i + 1
    Next i
    num_lineas_total = j
Else
    MsgBox "La entidad seleccionada no es una polilínea"
End If

End Sub
Sub datos_excel()
'Dim objExcel As Object
'Dim objLibro As Object
'Dim sheets(1) As Object
Dim fila_ini, fila_fin, columna As Long
'Set objExcel = GetObject(, "Excel.Application")
'Set objLibro = objExcel.Workbooks(1)
'Set sheets(1) = objLibro.Worksheets(1)
fila_ini = Sireca.TextBox2.Text
fila_fin = Sireca.TextBox3.Text
iposte = 0
cadena = Sheets(1).Cells(2, 2).Value
For ichevau = 0 To 999
    chevau(ichevau).poste_antich = 0
Next
ichevau = 1
With ActiveWorkbook
With .Sheets(1)
For fila = fila_ini To fila_fin
    iposte = iposte + 1
    poste(iposte).lado = Sheets(1).Cells(fila, 30).Value
    poste(iposte).etiq_1 = Sheets(1).Cells(fila, 31).Value
    poste(iposte).etiq_2 = Sheets(1).Cells(fila, 32).Value
    poste(iposte).pk_global = Sheets(1).Cells(fila, 33).Value
    If fila < fila_fin Then
        poste(iposte).vano_post = Sheets(1).Cells(fila + 1, 4).Value
    End If
    poste(iposte).tipo = Sheets(1).Cells(fila, 16).Value
    If poste(iposte).tipo = "Anc.Chevau." Or poste(iposte).tipo = "Anc.Section." Or poste(iposte).tipo = "Anc.Aigu." Or poste(iposte).tipo = "Anc.Neutre" Then
        poste(iposte).AT = True
    Else
        poste(iposte).AT = False
    End If
    Select Case poste(iposte).tipo
        Case "Anc.Chevau.sans AT"
            poste(iposte).tipo = "Anc.Chevau."
        Case "Anc.Section.sans AT"
            poste(iposte).tipo = "Anc.Section."
        Case "Anc.Aigu.sans AT"
            poste(iposte).tipo = "Anc.Aigu."
        Case "Anc.Neutre sans AT"
            poste(iposte).tipo = "Anc.Neutre"
    End Select
    If poste(iposte).tipo = "Anc.Chevau." Or poste(iposte).tipo = "Anc.Section." Or poste(iposte).tipo = "Anc.Neutre" Then
        If iposte > 3 Then
            If poste(iposte - 3).tipo = "Anc.Chevau." And poste(iposte - 1).tipo = "Inter.Chevau." Then
                chevau(ichevau).tipo = "CHEVAU"
                chevau(ichevau).AT_post = poste(iposte).AT
                chevau(ichevau).AT_ant = poste(iposte - 3).AT
                chevau(ichevau).poste_ini = iposte - 3
                chevau(ichevau).poste_fin = iposte
                chevau(ichevau).num_vanos = 3
                If poste(iposte - 1).vano_post = poste(iposte - 3).vano_post Then
                    chevau(ichevau).sim_chevau = True
                Else
                    chevau(ichevau).sim_chevau = True
                End If
                ichevau = ichevau + 1
            End If
        End If
        If iposte > 4 Then
            If (poste(iposte - 4).tipo = "Anc.Chevau." Or poste(iposte - 4).tipo = "Anc.Section.") Then
                If poste(iposte).tipo = "Anc.Chevau." Then
                    chevau(ichevau).tipo = "CHEVAU"
                End If
                If poste(iposte).tipo = "Anc.Section." Then
                    chevau(ichevau).tipo = "SECTION"
                End If
                chevau(ichevau).AT_post = poste(iposte).AT
                chevau(ichevau).AT_ant = poste(iposte - 4).AT
                chevau(ichevau).poste_ini = iposte - 4
                chevau(ichevau).poste_fin = iposte
                chevau(ichevau).num_vanos = 4
                If poste(iposte - 1).vano_post = poste(iposte - 4).vano_post And poste(iposte - 2).vano_post = poste(iposte - 3).vano_post And chevau(ichevau).AT_ant = chevau(ichevau).AT_post Then
                    chevau(ichevau).sim_chevau = True
                Else
                    chevau(ichevau).sim_chevau = False
                End If
                ichevau = ichevau + 1
            End If
        End If
        'NUEVO
        If iposte > 6 Then
            If poste(iposte - 6).tipo = "Anc.Neutre" Then
                chevau(ichevau).tipo = "ZN"
                chevau(ichevau).AT_post = poste(iposte).AT
                chevau(ichevau).AT_ant = poste(iposte - 6).AT
                chevau(ichevau).poste_ini = iposte - 6
                chevau(ichevau).poste_fin = iposte
                chevau(ichevau).num_vanos = 6
                chevau(ichevau).sim_chevau = False
                ichevau = ichevau + 1
            End If
        End If
    End If
    If poste(iposte).tipo = "Axe.Antich." Then
        chevau(ichevau - 1).poste_antich = iposte
    End If
    If poste(iposte).tipo = "Anc.Antich." And poste(iposte - 1).tipo = "Axe.Antich." Then
        If (poste(iposte - 1).vano_post = poste(iposte - 2).vano_post) Then
            chevau(ichevau - 1).sim_antich = True
        Else
            chevau(ichevau - 1).sim_antich = False
        End If
    End If
    poste(iposte).descentramiento = 1000 * Sheets(1).Cells(fila, 8).Value
    poste(iposte).descentramiento_2mens = 1000 * Sheets(1).Cells(fila, 9).Value
    poste(iposte).mensula2a = False
    If poste(iposte).tipo = "Axe.Chevau." Or poste(iposte).tipo = "Axe.Section." Or poste(iposte).tipo = "Inter.Chevau." Or poste(iposte).tipo = "Inter.Section." Or poste(iposte).tipo = "Inter.Aigu." Or poste(iposte).tipo = "Axe.Aigu." Or poste(iposte).tipo = "Inter.Neutre" Or poste(iposte).tipo = "Axe.Neutre" Then
        poste(iposte).mensula2a = True
    End If
    poste(iposte).tunel = False
    If Sheets(1).Cells(fila, 25) = "Tunnel" Then
        poste(iposte).tunel = True
    End If
    poste(iposte).implantacion = Sheets(1).Cells(fila, 5).Value
    poste(iposte).altura_HC = Round(Sheets(1).Cells(fila, 10).Value, 2)
    poste(iposte).radio = Sheets(1).Cells(fila, 6).Value
    fila = fila + 1
Next
num_postes_total = iposte
num_chevau = ichevau - 1
End With
End With
'Set sheets(1) = Nothing
'Set objLibro = Nothing
'Set objExcel = Nothing
For ichevau = 1 To num_chevau
    If num_chevau > ichevau Then
        chevau(ichevau).PM_ant = Round((poste(chevau(ichevau + 1).poste_fin).pk_global - poste(chevau(ichevau).poste_ini).pk_global) / (chevau(ichevau + 1).poste_fin - chevau(ichevau).poste_ini), 2)
        chevau(ichevau).long_post = Round(poste(chevau(ichevau + 1).poste_fin).pk_global - poste(chevau(ichevau).poste_ini).pk_global, 2)
    End If
    If ichevau > 1 Then
        chevau(ichevau).PM_post = Round((poste(chevau(ichevau).poste_fin).pk_global - poste(chevau(ichevau - 1).poste_ini).pk_global) / (chevau(ichevau).poste_fin - chevau(ichevau - 1).poste_ini), 2)
        chevau(ichevau).long_ant = Round(poste(chevau(ichevau).poste_fin).pk_global - poste(chevau(ichevau - 1).poste_ini).pk_global, 2)
    End If
Next
End Sub
Sub excel_pks()
Dim fila As Integer
'Dim objExcel As Object
'Dim objLibro As Object
'Dim objHoja As Object
'Set objExcel = GetObject(, "Excel.Application")
'Set objLibro = objExcel.Workbooks(1)
'Set objHoja = objLibro.Worksheets(6)
With ActiveWorkbook
With .Sheets(6)
        For fila = 2 To 122
            matriz_pk(Sheets(6).Cells(fila, 1).Value) = Sheets(6).Cells(fila, 2).Value
        Next
    End With
End With
End Sub
