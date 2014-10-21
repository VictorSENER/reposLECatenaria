Attribute VB_Name = "formato"
'//
'// Rutina destinada a dar formato al excel de salida
'//
Sub formato(idiomaVB)
Dim w As Integer, bis As Integer
Dim z As Integer
'//
'// Inicializar variables
'//

z = 8
'//
'// Introducir los títulos de las filas en el idioma adecuado
'//
Select Case idioma
    Case Is = "Frances"
        Sheets(1).Cells(z, 1).Value = "Nº du pylône"
        Sheets(1).Cells(z, 2).Value = ""
        Sheets(1).Cells(z, 3).Value = "PK (m)"
        Sheets(1).Cells(z, 4).Value = "Portée aval (m)"
        Sheets(1).Cells(z, 5).Value = "Implantation (m)"
        Sheets(1).Cells(z, 6).Value = "Rayon (m)"
        Sheets(1).Cells(z, 7).Value = "Devers (mm)"
        Sheets(1).Cells(z, 8).Value = "Desax. 1 (m)"
        Sheets(1).Cells(z, 9).Value = "Desax. 2 (m)"
        Sheets(1).Cells(z, 10).Value = "Hauteur (m)"
        Sheets(1).Cells(z, 11).Value = "Pendulage aval 1"
        Sheets(1).Cells(z, 12).Value = "Pendulage aval 2"
        Sheets(1).Cells(z, 13).Value = "Connection électrique"
        Sheets(1).Cells(z, 14).Value = "Mise au rail"
        Sheets(1).Cells(z, 15).Value = "Parafoudres"
        Sheets(1).Cells(z, 16).Value = "Axe.Chevau."
        Sheets(1).Cells(z, 17).Value = "Repère poutrelle H"
        Sheets(1).Cells(z, 18).Value = "Type de poutrelle H"
        Sheets(1).Cells(z, 19).Value = "Moment en tête pytône (daN/m)"
        Sheets(1).Cells(z, 20).Value = "Arasement fondation (m)"
        Sheets(1).Cells(z, 21).Value = "Type du terrain"
        Sheets(1).Cells(z, 22).Value = "Type de massif"
        Sheets(1).Cells(z, 23).Value = "Volume du massif (m3)"
        Sheets(1).Cells(z, 24).Value = "Massif d'ancrage"
        Sheets(1).Cells(z, 25).Value = "Observations"
        Sheets(1).Cells(z, 26).Value = "Lg. 1/2 tir anc. à axe antich. (m)"
        Sheets(1).Cells(z, 27).Value = "Lg. de tir anc. à anc. (m)"
    Case Is = "Español"
        '//
        '// pendiente de actualizar
        '//
    Case Is = "catalan"
        '//
        '// pendiente de actualizar
        '//
    Case Is = "Ingles"
        Sheets(1).Cells(z, 1).Value = "profil number"
        Sheets(1).Cells(z, 2).Value = ""
        Sheets(1).Cells(z, 3).Value = "PK (m)"
        Sheets(1).Cells(z, 4).Value = "Span (m)"
        Sheets(1).Cells(z, 5).Value = "Implantation (m)"
        Sheets(1).Cells(z, 6).Value = "Radius (m)"
        Sheets(1).Cells(z, 7).Value = "Slope (mm)"
        Sheets(1).Cells(z, 8).Value = "Lateral Offset 1 (m)"
        Sheets(1).Cells(z, 9).Value = "lateral Offset 2 (m)"
        Sheets(1).Cells(z, 10).Value = "Contact wire height (m)"
        Sheets(1).Cells(z, 11).Value = "Dropper type 1"
        Sheets(1).Cells(z, 12).Value = "Dropper type 2"
        Sheets(1).Cells(z, 13).Value = "Electrical connection"
        Sheets(1).Cells(z, 14).Value = "Connecting to rail"
        Sheets(1).Cells(z, 15).Value = "Lightning"
        Sheets(1).Cells(z, 16).Value = "Overlap"
        Sheets(1).Cells(z, 17).Value = "Repère poutrelle H"
        Sheets(1).Cells(z, 18).Value = "Mast type"
        Sheets(1).Cells(z, 19).Value = "Moment of force (daN/m)"
        Sheets(1).Cells(z, 20).Value = "Foundation height (m)"
        Sheets(1).Cells(z, 21).Value = "Soil type"
        Sheets(1).Cells(z, 22).Value = "Foundation type"
        Sheets(1).Cells(z, 23).Value = "Foundation volume (m3)"
        Sheets(1).Cells(z, 24).Value = "Foundation anchor"
        Sheets(1).Cells(z, 25).Value = "Observations"
        Sheets(1).Cells(z, 26).Value = "Lg. 1/2 section (m)"
        Sheets(1).Cells(z, 27).Value = "Lg. section (m)"
End Select
'//
'// Formato del borde de las celdas de cabecera
'//
Sheets(1).Range(Sheets(1).Cells(z, 1), Sheets(1).Cells(z + 1, 27)).Interior.ColorIndex = 15
With Sheets(1).Range(Sheets(1).Cells(z, 1), Sheets(1).Cells(z + 1, 27)).Borders(xlEdgeLeft)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets(1).Range(Sheets(1).Cells(z, 1), Sheets(1).Cells(z + 1, 27)).Borders(xlEdgeTop)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets(1).Range(Sheets(1).Cells(z, 1), Sheets(1).Cells(z + 1, 27)).Borders(xlEdgeBottom)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets(1).Range(Sheets(1).Cells(z, 1), Sheets(1).Cells(z + 1, 27)).Borders(xlEdgeRight)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets(1).Range(Sheets(1).Cells(z, 1), Sheets(1).Cells(z + 1, 27)).Borders(xlInsideVertical)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
For i = 1 To 27
    Sheets(1).Range(Sheets(1).Cells(z, i), Sheets(1).Cells(z + 1, i)).MergeCells = True
Next i

z = z + 2
'//
'// Formato de agrupación de celdas
'//
While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
    For i = 1 To 3
        Sheets(1).Range(Sheets(1).Cells(z, i), Sheets(1).Cells(z + 1, i)).MergeCells = True
    Next i
    For i = 4 To 4
        Sheets(1).Range(Sheets(1).Cells(z + 1, i), Sheets(1).Cells(z + 2, i)).MergeCells = True
    Next i
    For i = 5 To 10
        Sheets(1).Range(Sheets(1).Cells(z, i), Sheets(1).Cells(z + 1, i)).MergeCells = True
    Next i
    For i = 11 To 13
        Sheets(1).Range(Sheets(1).Cells(z + 1, i), Sheets(1).Cells(z + 2, i)).MergeCells = True
    Next i
    For i = 14 To 24
        Sheets(1).Range(Sheets(1).Cells(z, i), Sheets(1).Cells(z + 1, i)).MergeCells = True
    Next i
    Sheets(1).Range(Sheets(1).Cells(z + 1, 4), Sheets(1).Cells(z + 2, 4)).MergeCells = True
    z = z + 2
Wend
'//
'// Formato del borde de las celdas de replanteo
'//
With Sheets(1).Range(Sheets(1).Cells(10, 1), Sheets(1).Cells(z, 27)).Borders(xlEdgeLeft)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets(1).Range(Sheets(1).Cells(10, 1), Sheets(1).Cells(z, 27)).Borders(xlEdgeBottom)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets(1).Range(Sheets(1).Cells(10, 1), Sheets(1).Cells(z, 27)).Borders(xlEdgeRight)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets(1).Range(Sheets(1).Cells(10, 1), Sheets(1).Cells(z, 27)).Borders(xlInsideVertical)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets(1).Range(Sheets(1).Cells(10, 1), Sheets(1).Cells(z, 24)).Borders(xlInsideHorizontal)
    .LineStyle = 2
    .ColorIndex = 15
End With
'//
'// Formato de los datos
'//
Sheets(1).Columns(3).NumberFormat = "0+000.0"
Sheets(1).Columns(6).NumberFormat = "0.00"
Sheets(1).Columns(7).NumberFormat = "0.00"
Sheets(1).Columns(8).NumberFormat = "0.00"
Sheets(1).Columns(10).NumberFormat = "0.00"
Sheets(1).Columns(19).NumberFormat = "0.00"
Sheets(1).Columns(23).NumberFormat = "0.00"
'//
'// Formato de centrar el texto
'//
Sheets(1).Range("A1", "AA10001").HorizontalAlignment = xlHAlignCenter
Sheets(1).Range("A1", "AA10001").VerticalAlignment = xlHAlignCenter
Sheets(1).Range("A8", "AA10001").WrapText = True
'//
'// Formato de ocultar columnas
'//
'Sheets(1).Columns("AB:AO").EntireColumn.Hidden = True
'Sheets(1).Columns("B").EntireColumn.Hidden = True
End Sub
'//
'// Rutina destinada a establecer el comentario del punto singular en el idioma adecuado
'//
Sub lenguaje(idiomaVB)
'//
'// Inicializar variables
'//
a = 3
'//
'// Mientras no se hayan comentado todos los puntos singulares
'//
While Not IsEmpty(Sheets(4).Cells(a, 1).Value)
punto = Sheets(4).Cells(a, 1).Value
If idiomaVB = "Frances" Then
    '//
    '// Traducir los comentarios al idioma adecuado
    '//
    Select Case punto
        Case Is = "P.S. > 7 m"
            Sheets(4).Cells(a, 23).Value = "Passage superieur nº " & Sheets(4).Cells(a, 3).Value
        Case Is = "Puente"
            Sheets(4).Cells(a, 23).Value = "Pont"
        Case Is = "7 > P.S. > 5,2 m"
            Sheets(4).Cells(a, 23).Value = "Passage superieur nº " & Sheets(4).Cells(a, 3).Value
        Case Is = "Conducto"
            Sheets(4).Cells(a, 23).Value = "Buse"
        Case Is = "Tunel"
            Sheets(4).Cells(a, 23).Value = "Tunnel nº " & Sheets(4).Cells(a, 3).Value
        Case Is = "P.N."
            Sheets(4).Cells(a, 23).Value = "Passage à niveau"
        Case Is = "P.I."
            Sheets(4).Cells(a, 23).Value = "Passage inférieur"
        Case Is = "Aguja"
            Sheets(4).Cells(a, 23).Value = "Aiguillage"
        Case Is = "Drenaje"
            Sheets(4).Cells(a, 23).Value = "Dallot"
        Case Is = "Viaducto"
            Sheets(4).Cells(a, 23).Value = "Viaduc"
            Sheets(4).Cells(a, 24).Value = "Commencement Viaduc"
            Sheets(4).Cells(a, 25).Value = "Pilier Viaduc"
            Sheets(4).Cells(a, 26).Value = "Final Viaduc"
        Case Is = "PuenteXL"
            Sheets(4).Cells(a, 23).Value = "Pont longue"
        Case Is = "Zona"
            Sheets(4).Cells(a, 23).Value = "Zone neutre"
    End Select
ElseIf idioma = "castellano" Then
'//
'// pendiente de actualizar
'//
ElseIf idioma = "catalan" Then
'//
'// pendiente de actualizar
'//
ElseIf idioma = "ingles" Then
'//
'// pendiente de actualizar
'//
End If
'//
'// Incrementar fila de los puntos singulares
'//
a = a + 1
Wend
End Sub

