Attribute VB_Name = "formato"
'//
'// Rutina destinada a dar formato al excel de salida
'//
Sub formato(idiomaVB)
Dim w As Integer, bis As Integer
Dim z As Integer

formato_dar.Hide

zini = CInt(formato_dar.TextBox1)
zfin = CInt(formato_dar.TextBox2)
z = zini
'//
'// Inicializar variables
'//
Call cargar.cabecera(idiomaVB)

'//
'// Introducir los títulos de las filas en el idioma adecuado
'//

For i = 1 To lngCampos - 1
    Sheets("Replanteo").Cells(8, i).Value = col(i)
Next

'//
'// Formato del borde de las celdas de cabecera
'//
Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, 1), Sheets("Replanteo").Cells(8 + 1, 27)).Interior.ColorIndex = 15
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, 1), Sheets("Replanteo").Cells(8 + 1, 27)).Borders(xlEdgeLeft)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, 1), Sheets("Replanteo").Cells(8 + 1, 27)).Borders(xlEdgeTop)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, 1), Sheets("Replanteo").Cells(8 + 1, 27)).Borders(xlEdgeBottom)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, 1), Sheets("Replanteo").Cells(8 + 1, 27)).Borders(xlEdgeRight)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, 1), Sheets("Replanteo").Cells(8 + 1, 27)).Borders(xlInsideVertical)
.LineStyle = xlContinuous
.Weight = xlMedium
End With
For i = 1 To 27
    Sheets("Replanteo").Range(Sheets("Replanteo").Cells(8, i), Sheets("Replanteo").Cells(8 + 1, i)).MergeCells = True
Next i


'//
'// Formato de agrupación de celdas
'//

While zini <= zfin
    For i = 1 To 3
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini, i), Sheets("Replanteo").Cells(zini + 1, i)).MergeCells = True
    Next i
    For i = 4 To 4
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, i), Sheets("Replanteo").Cells(zini + 2, i)).MergeCells = True
    Next i
    For i = 5 To 10
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini, i), Sheets("Replanteo").Cells(zini + 1, i)).MergeCells = True
    Next i
    For i = 11 To 13
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, i), Sheets("Replanteo").Cells(zini + 2, i)).MergeCells = True
    Next i
    For i = 14 To 24
        Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini, i), Sheets("Replanteo").Cells(zini + 1, i)).MergeCells = True
    Next i
    Sheets("Replanteo").Range(Sheets("Replanteo").Cells(zini + 1, 4), Sheets("Replanteo").Cells(zini + 2, 4)).MergeCells = True
    zini = zini + 2
Wend
'//
'// Formato del borde de las celdas de replanteo
'//
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, 1), Sheets("Replanteo").Cells(zfin, 27)).Borders(xlEdgeLeft)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, 1), Sheets("Replanteo").Cells(zfin, 27)).Borders(xlEdgeBottom)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, 1), Sheets("Replanteo").Cells(zfin, 27)).Borders(xlEdgeRight)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, 1), Sheets("Replanteo").Cells(zfin, 27)).Borders(xlInsideVertical)
    .LineStyle = 2
    .ColorIndex = 15
End With
With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z, 1), Sheets("Replanteo").Cells(zfin, 24)).Borders(xlInsideHorizontal)
    .LineStyle = 2
    .ColorIndex = 15
End With
'//
'// Formato de los datos
'//
Sheets("Replanteo").Columns(3).NumberFormat = "0+000.0"
Sheets("Replanteo").Columns(4).NumberFormat = "0.00"
Sheets("Replanteo").Columns(6).NumberFormat = "0.00"
Sheets("Replanteo").Columns(7).NumberFormat = "0.00"
Sheets("Replanteo").Columns(8).NumberFormat = "0.00"
Sheets("Replanteo").Columns(9).NumberFormat = "0.00"
Sheets("Replanteo").Columns(10).NumberFormat = "0.00"
Sheets("Replanteo").Columns(19).NumberFormat = "0.00"
Sheets("Replanteo").Columns(20).NumberFormat = "0.00"
Sheets("Replanteo").Columns(23).NumberFormat = "0.00"
Sheets("Replanteo").Columns(26).NumberFormat = "0.00"
Sheets("Replanteo").Columns(27).NumberFormat = "0.00"
'//
'// Formato de centrar el texto
'//
Sheets("Replanteo").Range("A1", "AA10001").HorizontalAlignment = xlHAlignCenter
Sheets("Replanteo").Range("A1", "AA10001").VerticalAlignment = xlHAlignCenter
Sheets("Replanteo").Range("A8", "AA10001").WrapText = True
'//
'// Formato de ocultar columnas
'//
Sheets("Replanteo").Columns("AB:AX").EntireColumn.Hidden = True
Sheets("Replanteo").Columns("B").EntireColumn.Hidden = True
Sheets("Replanteo").Columns("Q").EntireColumn.Hidden = True
End Sub
'//
'// Rutina destinada a establecer el comentario del punto singular en el idioma adecuado
'//
Sub lenguaje(idiomaVB)
'//
'// Inicializar variables
'//
a = 4
Call cargar.punto_singular(idiomaVB)
Sheets("Punto singular").Cells(a - 2, 23).Value = idiomaVB
'//
'// Mientras no se hayan comentado todos los puntos singulares
'//
While Not IsEmpty(Sheets("Punto singular").Cells(a, 1).Value)
punto = Sheets("Punto singular").Cells(a, 1).Value

'If idiomaVB = "Frances" Then
    '//
    '// Traducir los comentarios al idioma adecuado
    '//
    Select Case punto
        Case Is = "P.S. > 7 m"
            Sheets("Punto singular").Cells(a, 23).Value = pas_sup & " nº " & Sheets("Punto singular").Cells(a, 3).Value
        Case Is = "Puente"
            Sheets("Punto singular").Cells(a, 23).Value = pue
        Case Is = "7 > P.S. > 5,2 m"
            Sheets("Punto singular").Cells(a, 23).Value = pas_sup & " nº " & Sheets("Punto singular").Cells(a, 3).Value
        Case Is = "Conducto"
            Sheets("Punto singular").Cells(a, 23).Value = con
        Case Is = "Tunel"
            Sheets("Punto singular").Cells(a, 23).Value = tun & " nº " & Sheets("Punto singular").Cells(a, 3).Value
        Case Is = "P.N."
            Sheets("Punto singular").Cells(a, 23).Value = p_n
        Case Is = "P.I."
            Sheets("Punto singular").Cells(a, 23).Value = p_i
        Case Is = "Aguja"
            Sheets("Punto singular").Cells(a, 23).Value = aguj & " " & Sheets("Punto singular").Cells(a, 4).Value
        Case Is = "Drenaje"
            Sheets("Punto singular").Cells(a, 23).Value = dren
        Case Is = "Viaducto"
            Sheets("Punto singular").Cells(a, 23).Value = via
        Case Is = "PuenteXL"
            Sheets("Punto singular").Cells(a, 23).Value = pue_xl
        Case Is = "Zona"
            Sheets("Punto singular").Cells(a, 23).Value = zon
        Case Is = "Marquesina"
            Sheets("Punto singular").Cells(a, 23).Value = mar
        Case Is = "Estacion"
            Sheets("Punto singular").Cells(a, 23).Value = est
        Case Is = "Señalización"
            Sheets("Punto singular").Cells(a, 23).Value = sen
        Case Is = "LEHT"
            Sheets("Punto singular").Cells(a, 23).Value = lin
        Case Is = "Subestación"
            Sheets("Punto singular").Cells(a, 23).Value = SS
        Case Is = "Pórtico catenaria"
            Sheets("Punto singular").Cells(a, 23).Value = pot_ali
    End Select

'//
'// Incrementar fila de los puntos singulares
'//
a = a + 1
Wend
End Sub

