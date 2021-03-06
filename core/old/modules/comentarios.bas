Attribute VB_Name = "comentarios"
'//
'// Rutina destinada a incluir los comentarios del trazado
'//
Sub comentarios()
Dim z As Integer, aloc As Integer, cont As Integer
Dim valor As String, guardar As String, rango As Range
'//
'// inicializar variables locales
'//
z = 10
aloc = 3
rayo = "Parafoudres"
cont = 3
Call cargar.datos_acces(nombre_catVB)
'//
'// inicio de la rutina
'//
While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
    '//
    '// buscar puntos sigulares
    '//
    While Sheets(1).Cells(z, 33).Value >= Sheets(4).Cells(aloc, 21).Value And Sheets(4).Cells(aloc, 23).Value <> "FINAL"
    aloc = aloc + 1
    Wend
    '//
    '// Si PK actual coincide con puntos singulares, escribir su respectivo comentario
    '//
    If Sheets(1).Cells(z + 2, 33).Value >= Sheets(4).Cells(aloc, 21).Value Then
        If Sheets(4).Cells(aloc, 1).Value = "Conducto" Or Sheets(4).Cells(aloc, 1).Value = "Drenaje" _
        Or Sheets(4).Cells(aloc, 1).Value = "Puente" Or Sheets(4).Cells(aloc, 1).Value = "P.I." _
        Or Sheets(4).Cells(aloc, 1).Value = "P.S. > 7 m" Or Sheets(4).Cells(aloc, 1).Value = "7 > P.S. > 5,2 m" _
        Or Sheets(4).Cells(aloc, 1).Value = "PuenteXL" Or Sheets(4).Cells(aloc, 1).Value = "P.N." Then
        valor = Sheets(4).Cells(aloc, 23).Value
            If Sheets(4).Cells(aloc, 1).Value = "Aguja" Then
                z_var = z
            Else
                z_var = z + 1
            End If
        '//
        '// Insertar comentario en excel
        '//
        If IsEmpty(Sheets(1).Cells(z_var + 1, 25).Value) Then
            Sheets(1).Cells(z_var + 1, 25).Value = valor
        Else
            Sheets(1).Cells(z_var + 1, 25).Value = Sheets(1).Cells(z_var + 1, 25).Value & " " & valor
        End If
        If Not IsEmpty(Sheets(1).Cells(z_var, 25).Value) Then
            guardar = Sheets(1).Cells(z_var, 25).Value
            Sheets(1).Cells(z_var, 25).Value = ""
            Sheets(1).Cells(z_var, 25).Value = guardar & " - " & valor
        End If
        
        '//
        '// Formato de la celda de comentarios
        '//
        With Sheets(1).Range(Sheets(1).Cells(z_var, 25), Sheets(1).Cells(z_var + 1, 25))
            .Borders(xlEdgeLeft).LineStyle = 2
            .Borders(xlEdgeLeft).ColorIndex = 15
            .Borders(xlEdgeTop).LineStyle = 2
            .Borders(xlEdgeTop).ColorIndex = 15
            .Borders(xlEdgeBottom).LineStyle = 2
            .Borders(xlEdgeBottom).ColorIndex = 15
            .Borders(xlEdgeRight).LineStyle = 2
            .Borders(xlEdgeRight).ColorIndex = 15
            .MergeCells = True
        End With
        
        End If
    End If
'//
'// relleno de la columna de altura y distancia poste carril
'//
Sheets(1).Cells(z, 10).Value = alt_nom
If Sheets(1).Cells(z, 16).Value = "Axe.Aigu." Then
    Sheets(1).Cells(z, 5).Value = 2.2 ' !!!!!!!!!!!!!!!!!FALTA VARIABLE EN BBDD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Else
    Sheets(1).Cells(z, 5).Value = dist_carril_poste
End If
Sheets(1).Cells(z, 20).Value = dist_base_poste_pmr

If Sheets(1).Cells(z, 16).Value = "Axe.Antich." And Sheets(1).Cells(z, 38).Value <> "Tunel" Then
    If rayo = "Parafoudres" Then
        Sheets(1).Cells(z, 15).Value = "Parafoudres" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z, 14).Value = "Mise ? la terre" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        rayo = "Parafoudres + DPPo"
    ElseIf rayo = "Parafoudres + DPPo" Then
        Sheets(1).Cells(z, 15).Value = "Parafoudres + DPPo" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z, 14).Value = "Mise au rail + mise ? la terre" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        rayo = "Parafoudres"
    End If
End If

If Sheets(1).Cells(z, 33).Value > tierra And Sheets(1).Cells(z, 38).Value <> "Tunel" Then
    tierra = tierra + 3000 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    Sheets(1).Cells(z, 14).Value = "Mise ? la terre"
End If
If Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Chevau." Then
    Sheets(1).Cells(z + 1, 13).Value = "667001 Rep. 51" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
ElseIf Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z - 2, 16).Value = "Axe.Chevau." Then
    Sheets(1).Cells(z - 1, 13).Value = "667001 Rep. 51" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
ElseIf Sheets(1).Cells(z, 16).Value = "Inter.Section." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Section." Then
    Sheets(1).Cells(z - 1, 13).Value = "667001 Rep. 53" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
ElseIf Sheets(1).Cells(z, 16).Value = "Inter.Section." And Sheets(1).Cells(z - 2, 16).Value = "Axe.Section." Then
    Sheets(1).Cells(z + 1, 13).Value = "667001 Rep. 53" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
ElseIf Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Chevau." Then
    Sheets(1).Cells(z + 1, 13).Value = "667001 Rep. 51" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
ElseIf Sheets(1).Cells(z, 16).Value = "Axe.Antich." Then
    Sheets(1).Cells(z - 1, 13).Value = "667001 Rep. 02" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
ElseIf Sheets(1).Cells(z, 16).Value = "Axe.Aigu." Then
    Sheets(1).Cells(z - 1, 13).Value = "667001 Rep. 22" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    Sheets(1).Cells(z + 1, 13).Value = "667001 Rep. 22" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
End If

While Not IsEmpty(Sheets(1).Cells(cont, 5).Value) Or Sheets(1).Cells(z, 33).Value >= Sheets(5).Cells(cont, 6).Value
cont = cont + 1
Wend
Sheets(1).Cells(z, 21).Value = Sheets(5).Cells(cont, 7).Value


'//
'// Incrementar fila del replanteo
'//
z = z + 2
Wend
End Sub

