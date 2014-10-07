Attribute VB_Name = "comentarios"
'//
'// Rutina destinada a incluir los comentarios del trazado
'//
Sub puntos_singulares()
Dim z As Integer, aloc As Integer, cont As Integer
Dim valor As String, guardar As String, rango As Range
'//
'// inicializar variables locales
'//
z = 10
aloc = 3
Call cargar.datos_lac(nombre_catVB)
'//
'// inicio de la rutina
'//
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
    '//
    '// buscar puntos sigulares
    '//
   While Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Punto singular").Cells(aloc, 23).Value <> "FINAL" _
    And Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(aloc, 21).Value
        aloc = aloc + 1
    Wend
    '//
    '// Si PK actual coincide con puntos singulares, escribir su respectivo comentario
    '//
    If Sheets("Replanteo").Cells(z + 2, 33).Value > Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Punto singular").Cells(aloc, 1).Value <> "Señalización" And Sheets("Punto singular").Cells(aloc, 1).Value <> "Aguja" Then
        valor = Sheets("Punto singular").Cells(aloc, 23).Value
        '///
        '///Si es aguja no agrupar dos celdas
        '///
        If Sheets("Punto singular").Cells(aloc, 22).Value = "IN" Then
            z_var = z - 1
            z_var_i = z_var
        ElseIf Sheets("Punto singular").Cells(aloc, 22).Value = "OUT" Then
            z_var = z + 1
            z_var_i = z_var
        Else
            z_var = z + 1
        End If
        '//
        '// Insertar comentario en excel
        '//
        Sheets("Replanteo").Cells(z_var, 25).Value = valor
        '//
        '// Formato de la celda de comentarios
        '//
        'With Sheets("Replanteo").Range(Sheets("Replanteo").Cells(z_var_i, 25), Sheets("Replanteo").Cells(z_var + 1, 25))
            '.Borders(xlEdgeLeft).LineStyle = 2
            '.Borders(xlEdgeLeft).ColorIndex = 15
            '.Borders(xlEdgeTop).LineStyle = 2
            '.Borders(xlEdgeTop).ColorIndex = 15
            '.Borders(xlEdgeBottom).LineStyle = 2
            '.Borders(xlEdgeBottom).ColorIndex = 15
            '.Borders(xlEdgeRight).LineStyle = 2
            '.Borders(xlEdgeRight).ColorIndex = 15
            '.MergeCells = True
        'End With
    While Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(aloc, 2).Value And Sheets("Punto singular").Cells(aloc, 23).Value <> "FINAL" _
    And Sheets("Replanteo").Cells(z, 33).Value > Sheets("Punto singular").Cells(aloc, 21).Value
        aloc = aloc + 1
    Wend
    End If
'//
'// Incrementar fila del replanteo
'//
z = z + 2
Wend
End Sub
Sub conexiones(nombre_catVB)
Dim ddpo() As Double
'///
'///Inicializar variables
'///

a = 3
z = 10
'tierra = 1200 '/// eliminado reunión 7-1-14
lon_cdpa = 1500
repar = Sheets("Replanteo").Cells(z, 33).Value + 250
fed_lac = Sheets("Replanteo").Cells(z, 33).Value + 400
rayo = "Parafoudres" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
Call cargar.datos_lac(nombre_catVB)
cont = 0
cont_est = 0

'//
'// inicio de la rutina
'//
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
    '///
    '/// recoger ubicaciones de DDPO
    '///
    While Not IsEmpty(Sheets("Extra").Cells(a, 24).Value) And Sheets("Replanteo").Cells(z, 3).Value > Sheets("Extra").Cells(a, 24).Value
        a = a + 1
        Sheets("Replanteo").Cells(z, 15).Value = "DPPo" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets("Replanteo").Cells(z, 14).Value = "Mise au rail" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    Wend
    
    
    '///
    '/// insertar los comentarios de puesta a tierra y pararrayos
    '///
    If Sheets("Replanteo").Cells(z, 16).Value = eje_pf And Sheets("Replanteo").Cells(z, 38).Value <> "Tunel" Then
        If rayo = "Parafoudres" Then
            Sheets("Replanteo").Cells(z, 15).Value = "Parafoudres - DPPo" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets("Replanteo").Cells(z, 14).Value = "Mise à la terre - Mise au rail" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            rayo = "DPPo" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        ElseIf rayo = "DPPo" Then
            Sheets("Replanteo").Cells(z, 15).Value = "DPPo" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets("Replanteo").Cells(z, 14).Value = "Mise au rail" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            rayo = "Parafoudres" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        End If
    End If
    '///
    '/// insertar los anclajes de cdpa y feeder
    '///

    

    If Sheets("Replanteo").Cells(z, 33).Value > lon_cdpa And Sheets("Replanteo").Cells(z, 38).Value <> "Tunel" And (Sheets("Replanteo").Cells(z, 16).Value = anc_pf Or Sheets("Replanteo").Cells(z, 16).Value = anc_sm_sin) And cont < 1 Then
        lon_cdpa = Sheets("Replanteo").Cells(z, 33).Value + 1700 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets("Replanteo").Cells(z, 17).Value = "Anc. CdPA et Feeder" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf (Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_sin Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con & " + " & semi_eje_aguj) And cont = 3 Then
        lon_cdpa = Sheets("Replanteo").Cells(z - 2, 33).Value + 1700 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    End If
    '///
    '/// Variar la escala en estaciones
    '///
    If (Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_sin Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And z <> 10 And z <> 20 Then
        cont = cont + 1
    ElseIf (Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_sin Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con & " + " & semi_eje_aguj) And Sheets("Replanteo").Cells(z + 2, 16).Value = "" And cont = 3 Then
        cont = 0
    End If
    '///
    '/// insertar los códigos de conexiones de cableado (exclusivo para la catenaria francesa)
    '///
    If (Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sm And Sheets("Replanteo").Cells(z + 2, 16).Value = eje_sm) Or _
    (Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sm And Sheets("Replanteo").Cells(z + 2, 16).Value = semi_eje_sm) Then
        Sheets("Replanteo").Cells(z + 1, 13).Value = "667001-51" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf (Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sm And Sheets("Replanteo").Cells(z - 2, 16).Value = eje_sm) Or _
    (Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sm And Sheets("Replanteo").Cells(z - 2, 16).Value = semi_eje_sm) Then
        Sheets("Replanteo").Cells(z - 1, 13).Value = "667001-51" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sla And Sheets("Replanteo").Cells(z + 2, 16).Value = eje_sla Then
        Sheets("Replanteo").Cells(z - 1, 13).Value = "667001-53" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf (Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sla Or Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sla & " + " & anc_aguj) And Sheets("Replanteo").Cells(z - 2, 16).Value = eje_sla Then
        Sheets("Replanteo").Cells(z + 1, 13).Value = "667001-53" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sm And Sheets("Replanteo").Cells(z + 2, 16).Value = eje_sm Then
        Sheets("Replanteo").Cells(z + 1, 13).Value = "667001-51" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf Sheets("Replanteo").Cells(z, 16).Value = eje_pf Then
        Sheets("Replanteo").Cells(z - 1, 13).Value = "667001-02" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf Sheets("Replanteo").Cells(z, 16).Value = eje_aguj Then
        Sheets("Replanteo").Cells(z - 1, 13).Value = "667001-23" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets("Replanteo").Cells(z + 1, 13).Value = "667001-23" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    End If
    '///
    '/// insertar los comentarios conexion equipotencial
    '///
    If Sheets("Replanteo").Cells(z, 33).Value > repar Then
        If Not IsEmpty(Sheets("Replanteo").Cells(z, 16).Value) Or Not IsEmpty(Sheets("Replanteo").Cells(z - 2, 16).Value) Then
            repar = repar + 250
        Else
            repar = repar + 250 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets("Replanteo").Cells(z - 1, 13).Value = "667001-02" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        End If
    End If
    '///
    '/// Verificación del lado a implantar los datos en estacion
    '///
        
     If Sheets("Replanteo").Cells(z, 56).Value = "" And cont_est = 0 Then
        '///
        '/// insertar los comentarios conexion equipotencial
        '///
        If Sheets("Replanteo").Cells(z, 33).Value > fed_lac Then
            zbis = z
            While Not IsEmpty(Sheets("Replanteo").Cells(zbis, 16).Value) Or Not IsEmpty(Sheets("Replanteo").Cells(zbis - 1, 13).Value)
                zbis = zbis - 2
            Wend
            fed_lac = Sheets("Replanteo").Cells(zbis, 33).Value + 400 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets("Replanteo").Cells(zbis - 1, 13).Value = "667001-90" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            
        End If
    ElseIf Sheets("Replanteo").Cells(z, 56).Value <> "" And cont_est = 0 Then
        cont_est = 1


    ElseIf Sheets("Replanteo").Cells(z, 56).Value <> "" And cont_est = 1 Then
        cont_est = 0
        fed_lac = Sheets("Replanteo").Cells(z, 33).Value + 400 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        
    'ElseIf Sheets("Replanteo").Cells(z, 56).Value <> "" And cont_est = 2 Then
        'cont_est = 0
    End If
    
    
'//
'// Incrementar fila del replanteo
'//
z = z + 2
Wend
End Sub
Sub implantacion(nombre_catVB)
'///
'/// inicializar variables
'///
z = 10
cont = 3
j = 2
Call cargar.datos_lac(nombre_catVB)
'//
'// inicio de la rutina, cambio después de reunión 7-1-14. implantación = 1,7 siempre
'//
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
    If Mid(Sheets("Replanteo").Cells(z, 3).Value, 3, 3) = "bis" Then
        Replanteo = Mid(Sheets("Replanteo").Cells(z, 3).Value, 1, 2) & Mid(Sheets("Replanteo").Cells(z, 3).Value, 7)
        implanta = Mid(Sheets("Extra").Cells(j, 21).Value, 1, 2) & Mid(Sheets("Extra").Cells(j, 21).Value, 7)
        implanta2 = Mid(Sheets("Extra").Cells(j, 20).Value, 1, 2) & Mid(Sheets("Extra").Cells(j, 20).Value, 7)
    Else
        Replanteo = Sheets("Replanteo").Cells(z, 3).Value
        implanta = Sheets("Extra").Cells(j, 21).Value
        implanta2 = Sheets("Extra").Cells(j, 20).Value
    End If
    While implanta < Replanteo
        j = j + 1
    If Mid(Sheets("Replanteo").Cells(z, 3).Value, 3, 3) = "bis" Then
        Replanteo = Mid(Sheets("Replanteo").Cells(z, 3).Value, 1, 2) & Mid(Sheets("Replanteo").Cells(z, 3).Value, 7)
        implanta = Mid(Sheets("Extra").Cells(j, 21).Value, 1, 2) & Mid(Sheets("Extra").Cells(j, 21).Value, 7)
        implanta2 = Mid(Sheets("Extra").Cells(j, 20).Value, 1, 2) & Mid(Sheets("Extra").Cells(j, 20).Value, 7)
    Else
        Replanteo = Sheets("Replanteo").Cells(z, 3).Value
        implanta = Sheets("Extra").Cells(j, 21).Value
        implanta2 = Sheets("Extra").Cells(j, 20).Value
    End If
    Wend
    
    '///
    '/// caso de estar dentro de túnel
    '///

    If Sheets("Replanteo").Cells(z, 38).Value = "Tunel" Or Sheets("Replanteo").Cells(z, 38).Value = "Marquesina" Then
    '///
    '/// caso de estar dentro fuera de túnel
    '///
    Else

        If implanta2 <= Replanteo And implanta >= Replanteo Then
           
            Sheets("Replanteo").Cells(z, 5).Value = Sheets("Extra").Cells(j, 22).Value

        '//
        '// relleno de la columna de distancia poste carril
        '//
        
        ElseIf Sheets("Replanteo").Cells(z, 16).Value = eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = eje_pf & " + " & eje_aguj _
        Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & eje_aguj Then
            Sheets("Replanteo").Cells(z, 5).Value = 2.2 ' !!!!!!!!!!!!!!!!!FALTA VARIABLE EN BBDD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            
        'ElseIf (Sheets("Replanteo").Cells(z, 16).Value = eje_sm Or Sheets("Replanteo").Cells(z, 16).Value = eje_sla) _
        'And Sheets("Replanteo").Cells(z, 6).Value < 0 Then
            'Sheets("Replanteo").Cells(z, 5).Value = 2.2 ' !!!!!!!!!!!!!!!!!FALTA VARIABLE EN BBDD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Else
            Sheets("Replanteo").Cells(z, 5).Value = dist_carril_poste
        End If
        
        If Sheets("Replanteo").Cells(z, 38).Value = "Viaducto" Then
        
        Else
        '///
        '/// insertar distancia entre PMR y cimentación
        '///
        Sheets("Replanteo").Cells(z, 20).Value = dist_base_poste_pmr
        '///
        '///insertar tipo de terreno
        '///
        
        While Sheets("Replanteo").Cells(z, 33).Value <= Sheets("Extra").Cells(cont, 5).Value Or Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Extra").Cells(cont, 6).Value
    
            cont = cont + 1
        Wend
        Sheets("Replanteo").Cells(z, 21).Value = Sheets("Extra").Cells(cont, 7).Value
        End If
    End If
    z = z + 2

Wend
End Sub
