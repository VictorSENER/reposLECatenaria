Attribute VB_Name = "num_postes"
'//
'// Rutina destinada a numerar los postes
'//
Sub postes(nombre_cat)
Dim w As Integer, bis As Integer
Dim z As Integer
Dim a As Integer
Dim polp As Integer, pola As Integer
Dim pk_nolineal() As Double
'//
'// Inicializar variables y cargar los datos de catenaria
'//
a = 2
bis = 1
z = 10
polp = 0
If Sheets("Replanteo").Cells(z, 30).Value = "G" Then
    w = 1
Else
    w = 2
End If
Sheets("Replanteo").Cells(z, 1).Value = (Sheets("Replanteo").Cells(z, 3).Value \ 1000) & "-" & w
Sheets("Replanteo").Cells(z, 32).Value = w & "AF"
Sheets("Replanteo").Cells(z, 31).Value = (Sheets("Replanteo").Cells(z, 3).Value \ 1000)
w = w + 2
z = z + 2
'//
'//Buscar los kilométros no lineales
'//

While Not IsEmpty(Sheets("Pk real").Cells(a, 1).Value)
    If Sheets("Pk real").Cells(a, 1).Value = Sheets("Pk real").Cells(a - 1, 1).Value Then
        polp = polp + 3
        ReDim Preserve pk_nolineal(1 To polp)
        pk_nolineal(polp - 2) = Sheets("Pk real").Cells(a, 2).Value
        pk_nolineal(polp - 1) = Sheets("Pk real").Cells(a + 1, 2).Value
        pk_nolineal(polp) = Sheets("Pk real").Cells(a, 1).Value
        
    End If
    a = a + 1
Wend
pola = 1
If IsEmpty(Sheets("Pk real").Cells(a, 1).Value) Then
        polp = polp + 3
        ReDim Preserve pk_nolineal(1 To polp)
        pk_nolineal(polp - 2) = 0
        pk_nolineal(polp - 1) = 0
        pk_nolineal(polp) = 0
End If
'//
'// Final de replanteo?
'//
b = 4
cont = 1
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
    '/// ENCONTRAR ESTACIONES
    If z < 20 And (Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla Or Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla & " + " & anc_aguj) And cont <= 3 Then
        estacion = True
        cont = cont + 2
    ElseIf z <= 20 And cont >= 4 And Sheets("Replanteo").Cells(z - 2, 16).Cells = semi_eje_sla And Sheets("Replanteo").Cells(z - 4, 16).Cells = eje_sla Then
        estacion = False
        cont = 1
        Sheets("Replanteo").Cells(z - 3, 13).Value = "667001-90" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    ElseIf z > 20 And (Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla Or Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla & " + " & anc_aguj) And cont <= 3 Then
        estacion = True
        cont = cont + 1
        If cont = 2 Then
            Sheets("Replanteo").Cells(z + 1, 13).Value = "667001-90" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        End If
    ElseIf cont = 4 And Sheets("Replanteo").Cells(z - 2, 16).Cells = semi_eje_sla And Sheets("Replanteo").Cells(z - 4, 16).Cells = eje_sla Then
        estacion = False
        cont = 1
        Sheets("Replanteo").Cells(z - 3, 13).Value = "667001-90" '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
    End If
    
    '///
    '/// encontrar subestaciones
    '///
    While (Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Punto singular").Cells(b, 21).Value And Sheets("Punto singular").Cells(b, 23).Value <> "FINAL" _
    And (Sheets("Punto singular").Cells(b, 1).Value <> "Subestación" And Sheets("Punto singular").Cells(b, 1).Value <> "Pórtico catenaria"))
        b = b + 1
    Wend
    '///
    '/// caso de estar dentro de túnel
    '///
    If Sheets("Replanteo").Cells(z, 38).Value = "Tunel" Then
        cod_tun = "T"
    ElseIf Sheets("Replanteo").Cells(z, 38).Value = "Marquesina" Then
        cod_tun = "M"
    ElseIf Sheets("Replanteo").Cells(z, 16).Value = anc_pf Or Sheets("Replanteo").Cells(z, 16).Value = anc_aguj Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_sin _
    Or Sheets("Replanteo").Cells(z, 16).Value = anc_sm_con Or Sheets("Replanteo").Cells(z, 16).Value = anc_sm_sin Or Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sla & " + " & anc_aguj Or Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con & " + " & semi_eje_aguj _
    Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & semi_eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & anc_aguj Then
        cod_tun = "A"
    ElseIf ((Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla Or Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla & " + " & anc_aguj) And cont = 2) Then
        cod_tun = "A"
        Sheets("Replanteo").Cells(z, 17).Value = "Anc. Feeder Alim."
    ElseIf cont >= 4 And Sheets("Replanteo").Cells(z, 16).Cells = semi_eje_sla And Sheets("Replanteo").Cells(z - 2, 16).Cells = eje_sla Then
        cod_tun = "A"
        Sheets("Replanteo").Cells(z, 17).Value = "Anc. Feeder Alim."
    Else
        cod_tun = ""
    End If
    If estacion = True Then
    
        cod_tun = cod_tun & "F"
    End If
    '//
    '// Caso particular de existencia de PK BIS
    '//
    If pk_nolineal(pola) <= Sheets("Replanteo").Cells(z, 33).Value And Sheets("Replanteo").Cells(z, 33).Value < pk_nolineal(pola + 1) Then
            Sheets("Replanteo").Cells(z, 31).Value = pk_nolineal(pola + 2) & "bis" '/// mequedo aqui
            Sheets("Replanteo").Cells(z, 32).Value = bis
            Sheets("Replanteo").Cells(z, 1).Value = pk_nolineal(pola + 2) & "bis -" & bis
            bis = bis + 2
            w = 0
        If Sheets("Replanteo").Cells(z, 33).Value > pk_nolineal(pola + 1) And pola + 2 < polp Then
            pola = pola + 3
            bis = 1
        End If
    '///
    '/// caso no existencia de PK bis
    '///
    Else
        If (Sheets("Punto singular").Cells(b, 1).Value = "Subestación" Or Sheets("Punto singular").Cells(b, 1).Value = "Pórtico catenaria") And Sheets("Replanteo").Cells(z, 33).Value >= Sheets("Punto singular").Cells(b, 2).Value Then
    
            w = w + 2
            b = b + 1
        End If
    
        If w = 0 And Sheets("Replanteo").Cells(z, 30).Value = "G" Then
            w = 1
        ElseIf w = 0 And Sheets("Replanteo").Cells(z, 30).Value = "D" Then
            w = 2
        ElseIf (Sheets("Replanteo").Cells(z, 3).Value \ 1000) > (Sheets("Replanteo").Cells(z - 2, 3).Value \ 1000) And Sheets("Replanteo").Cells(z, 30).Value = "G" Then
            w = 1
        ElseIf (Sheets("Replanteo").Cells(z, 3).Value \ 1000) > (Sheets("Replanteo").Cells(z - 2, 3).Value \ 1000) And Sheets("Replanteo").Cells(z, 30).Value = "D" Then
            w = 2
        End If
        Sheets("Replanteo").Cells(z, 1).Value = (Sheets("Replanteo").Cells(z, 3).Value \ 1000) & "-" & w & cod_tun
        Sheets("Replanteo").Cells(z, 32).Value = w & cod_tun
        Sheets("Replanteo").Cells(z, 31).Value = (Sheets("Replanteo").Cells(z, 3).Value \ 1000)
        w = w + 2
    End If
    '///
    '/// adecuación de la numeración ante el cambio de lado de la vía
    '///
    If Sheets("Replanteo").Cells(z, 30).Value <> Sheets("Replanteo").Cells(z + 2, 30).Value Then
        If w Mod 2 = 0 Then
            w = w - 1
        Else
            w = w + 1
        End If
    End If
Set text = a_text.CreateTextFile(dir_progress)
text.WriteLine "14" & "/" & "14" & "/" & "Numeración postes" & "/" & Sheets("Replanteo").Cells(z, 33).Value & "/" & final
text.Close
    '//
    '// Incrementar fila del replanteo
    '//
    z = z + 2
Wend
End Sub
