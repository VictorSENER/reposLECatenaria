Attribute VB_Name = "descentramiento"
'//
'// Rutina destinada a elegir el descentramiento correspondiente al radio en la tabla de vanos
'//
Sub desc(nombre_catVB)
Dim z As Integer, M As Integer
Dim descentramiento As Double, rady As Double, l As Double
Dim cote As String
'//
'// Inicializar variables
'//
d_aguja = 0.1
z = 10
Section = "Normal"
Call cargar.datos_lac(nombre_catVB)
Sheets("Replanteo").Cells(1, 1).Value = nombre_catVB
Sheets("Replanteo").Cells(1, 2).Value = adm_lin_poste
'//
'// Mientras no lleguemos al final del replanteo
'//
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
'//
'// Inicializar variable local
'//
M = 3
rady = Abs(Sheets("Replanteo").Cells(z, 6).Value)
'//
'// Buscar en que fila de la hoja 2 se encuentra el radio que buscamos
'//
If Not IsEmpty(Sheets("Replanteo").Cells(z, 6).Value) Then
    While rady < Sheets("Vano").Cells(M, 3).Value
        M = M + 1
    Wend
Else
        M = 3
End If
descentramiento = Sheets("Vano").Cells(M, 5).Value
a = 3
cote = Sheets("Replanteo").Cells(z, 30).Value
'//
'// el descentramiento varia segun el sentido de giro de la curva
'// el descentramiento anterior y posterior varian de signo en recta
'//
If Not IsEmpty(Sheets("Replanteo").Cells(z, 6).Value) Then
    If Sheets("Replanteo").Cells(z, 6).Value >= 0 Then
        Sheets("Replanteo").Cells(z, 8).Value = descentramiento
    Else
      Sheets("Replanteo").Cells(z, 8).Value = -descentramiento
    End If
ElseIf IsEmpty(Sheets("Replanteo").Cells(z, 6).Value) Then
    If Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
        Sheets("Replanteo").Cells(z, 8).Value = -descentramiento
    Else
        Sheets("Replanteo").Cells(z, 8).Value = descentramiento
    End If
End If
'//
'//insertar descentramiento en seccionamientos de lamina de aire y compensación
'//los datos utilizados estan guardados en la base de datos
'//
If Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sla Then
    '///
    '/// seccionamientos de 3 vanos
    '///
    If Sheets("Replanteo").Cells(z - 1, 4).Value > 54 And Sheets("Replanteo").Cells(z + 1, 4).Value >= 54 And Sheets("Replanteo").Cells(z + 3, 4).Value >= 54 And Sheets("Replanteo").Cells(z + 2, 16).Value = semi_eje_sla Then
        If Section = "Inverso" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            'sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_max_re - d_max_ad
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_max_re
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_max_re
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_max_re + d_max_ad
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_max_re
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Normal"
        ElseIf Section = "Normal" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            'sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = d_max_re + d_max_ad
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_max_re
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_max_re
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_max_re - d_max_ad
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_max_re
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
            Section = "Inverso"
        Else
            algo = 0
        End If
        z = z + 2
    '///
    '/// seccionamientos de 5 vanos
    '///
    
    ElseIf (Sheets("Replanteo").Cells(z - 1, 4).Value <= 40.5 Or Sheets("Replanteo").Cells(z + 3, 4).Value <= 40.5 Or Sheets("Replanteo").Cells(z + 3, 4).Value <= 40.5 _
    Or Sheets("Replanteo").Cells(z + 5, 4).Value <= 40.5) And (Sheets("Replanteo").Cells(z + 6, 16).Value = semi_eje_sla Or Sheets("Replanteo").Cells(z + 6, 16).Value = semi_eje_sla & " + " & anc_aguj) Then
        If Section = "Inverso" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Normal"
        ElseIf Section = "Normal" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 8).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Inverso"
    ElseIf Section = "Normal" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sla1 + 0.4
            Sheets("Replanteo").Cells(z + 6, 9).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Inverso"
        ElseIf Section = "Normal" And cote = "G" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value <= 0 Then
            'Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sla1
            'Sheets("Replanteo").Cells(z, 9).Value = -d_eje_sla1
            'Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla1
            'Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla2
            'Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sla1
            'Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sla2
            'Sheets("Replanteo").Cells(z + 6, 9).Value = -d_semi_eje_sla1
            'Sheets("Replanteo").Cells(z + 6, 8).Value = d_semi_eje_sla2
            
            Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla1 + 0.4
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
            Section = "Inverso"
        ElseIf Section = "Inverso" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value <= 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla1 + 0.4
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso" ' verificar
            Section = "Normal"
            
        ElseIf Section = "Normal" And cote = "G" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value >= 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
            Section = "Inverso"
        ElseIf Section = "Inverso" And cote = "G" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value >= 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 6, 8).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Normal"
        ElseIf Section = "Inverso" And cote = "G" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 Then ' And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sla1 + 0.4
            Sheets("Replanteo").Cells(z + 6, 9).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Normal"
        Else
            algo = 0
        End If
        z = z + 6
    '///
    '/// seccionamientos de 4 vanos
    '///
    ElseIf Sheets("Replanteo").Cells(z + 4, 16).Value = semi_eje_sla Or Sheets("Replanteo").Cells(z + 4, 16).Value = anc_aguj & " + " & semi_eje_sla Then
    
        If Section = "Inverso" And cote = "G" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_eje_sla2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Normal"
        ElseIf Section = "Normal" And cote = "G" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value <= 0 Then ' And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_eje_sla2
            Sheets("Replanteo").Cells(z, 9).Value = -d_eje_sla2 + 0.4
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
            Section = "Inverso"
        ElseIf Section = "Normal" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_eje_sla2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Inverso"
        ElseIf Section = "Normal" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla2
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_eje_sla2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
            Section = "Inverso"
        ElseIf Section = "Inverso" And cote = "D" And Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 Then 'And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z, 9).Value = d_eje_sla1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sla2
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sla1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
            Section = "Normal"
        
        Else
            Sheets("Replanteo").Cells(z, 9).Value = "ERROR"
            Sheets("Replanteo").Cells(z + 2, 9).Value = "ERROR"
            Sheets("Replanteo").Cells(z + 4, 9).Value = "ERROR"
        End If
    z = z + 4
    End If

ElseIf Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sm Then
    '///
    '/// seccionamientos de 3 vanos
    '///
    If Sheets("Replanteo").Cells(z - 1, 4).Value >= 54 And Sheets("Replanteo").Cells(z + 1, 4).Value >= 54 And Sheets("Replanteo").Cells(z + 3, 4).Value >= 54 _
    And IsEmpty(Sheets("Replanteo").Cells(z, 6).Value) And IsEmpty(Sheets("Replanteo").Cells(z + 2, 6).Value) And Sheets("Replanteo").Cells(z + 4, 16).Value = anc_sm_con Then
        If Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_max_re
            Sheets("Replanteo").Cells(z, 9).Value = -d_max_ad
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_max_ad
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_max_re
            'sheets("Replanteo").Cells(z + 4, 8).Value = d_eje_sla1
            'sheets("Replanteo").Cells(z + 4, 9).Value = d_eje_sla2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
        ElseIf Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_max_re
            Sheets("Replanteo").Cells(z, 9).Value = d_max_ad
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_max_ad
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_max_re
            'sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sm2
            'sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
       Else
       algo = 0
       End If
        z = z + 2
    '///
    '/// seccionamientos de 5 vanos
    '///
    ElseIf (Sheets("Replanteo").Cells(z - 1, 4).Value <= 31.5 Or Sheets("Replanteo").Cells(z + 3, 4).Value <= 31.5 Or Sheets("Replanteo").Cells(z + 3, 4).Value <= 31.5 _
    Or Sheets("Replanteo").Cells(z + 5, 4).Value <= 31.5) And Sheets("Replanteo").Cells(z + 6, 16).Value = semi_eje_sm Then
        If Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 6, 8).Value = d_eje_sm1
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z - 2, 8).Value <= 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = -d_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 6, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z - 2, 8).Value <= 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = d_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 6, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_eje_sm1
            Sheets("Replanteo").Cells(z + 6, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
        Else
            algo = 0
        End If
        z = z + 6
    '///
    '/// seccionamientos de 4 vanos
    '///
    Else
        If Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value <= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_eje_sm2
            Sheets("Replanteo").Cells(z, 9).Value = d_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value <= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_eje_sm2
            Sheets("Replanteo").Cells(z, 9).Value = -d_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"

        ElseIf Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value <= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"

        ElseIf Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value <= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 8).Value = -d_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 9).Value = -d_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Normal"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_eje_sm2
            Sheets("Replanteo").Cells(z, 9).Value = -d_eje_sm1
            Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
        ElseIf Sheets("Replanteo").Cells(z, 6).Value <= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 _
        And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_sm2
            Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_sm2 + 0.2
            Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z + 4, 8).Value = d_semi_eje_sm2
            Sheets("Replanteo").Cells(z + 4, 9).Value = d_semi_eje_sm1
            Sheets("Replanteo").Cells(z - 2, 47).Value = "Inverso"
        Else
        Sheets("Replanteo").Cells(z, 9).Value = "ERROR"
        Sheets("Replanteo").Cells(z + 2, 9).Value = "ERROR"
        Sheets("Replanteo").Cells(z + 4, 9).Value = "ERROR"
        End If
    If Sheets("Replanteo").Cells(z, 6).Value = 0 And Sheets("Replanteo").Cells(z + 2, 6).Value = 0 And Sheets("Replanteo").Cells(z + 4, 6).Value = 0 _
      And Sheets("Replanteo").Cells(z - 2, 8).Value > 0 Then
            Sheets("Replanteo").Cells(z + 6, 8).Value = -d_max_re
            z = z + 2
    ElseIf Sheets("Replanteo").Cells(z, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 2, 6).Value >= 0 And Sheets("Replanteo").Cells(z + 4, 6).Value >= 0 _
      And Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
            Sheets("Replanteo").Cells(z + 6, 8).Value = d_max_re
            z = z + 2
    End If
    z = z + 4
    End If


'//
'//insertar descentramiento en agujas
'//
ElseIf (Sheets("Replanteo").Cells(z, 16).Value = eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = eje_pf & " + " & eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & eje_aguj) And _
(Sheets("Replanteo").Cells(z - 2, 16).Value = semi_eje_aguj Or Sheets("Replanteo").Cells(z - 2, 16).Value = anc_sla_con & " + " & semi_eje_aguj Or Sheets("Replanteo").Cells(z - 2, 16).Value = anc_pf & " + " & semi_eje_aguj Or Sheets("Replanteo").Cells(z - 2, 16).Value = eje_pf & " + " & semi_eje_aguj Or Sheets("Replanteo").Cells(z - 2, 16).Value = semi_eje_sla & " + " & anc_aguj) Then

    'While sheets("Replanteo").Cells(z, 25).Value <> sheets("Extra").Cells(a, 13).Value
        'a = a + 1
    'Wend
    'If Sheets("Replanteo").Cells(z - 2, 8).Value < 0 Then
        'd_semi_eje_aguja = d_aguja
    'Else
        'd_semi_eje_aguja = d_aguja * -1
    'End If

    'l = (sheets("Extra").Cells(a, 14).Value * sheets("Extra").Cells(a, 15).Value) / 2
    'long_min = (dist_pant_util + 112) / 1000
    'long_min_2 = l - d_max_re
    'If long_min > long_min_2 Then
        'd_aguja = long_min
    'Else
        'd_aguja = long_min_2
    'End If
    If Sheets("Replanteo").Cells(z + 1, 35).Value = "I" Then
        Sheets("Replanteo").Cells(z, 8).Value = d_eje_aguj1
        Sheets("Replanteo").Cells(z, 9).Value = d_eje_aguj2
        Sheets("Replanteo").Cells(z - 2, 8).Value = d_semi_eje_aguj1
        Sheets("Replanteo").Cells(z - 2, 9).Value = d_semi_eje_aguj2

        If Sheets("Replanteo").Cells(z - 4, 16).Value = semi_eje_aguj Then
            Sheets("Replanteo").Cells(z - 4, 8).Value = d_semi_eje_aguj1
            Sheets("Replanteo").Cells(z - 4, 9).Value = d_semi_eje_aguj2
        End If
    Else
        Sheets("Replanteo").Cells(z, 8).Value = -d_eje_aguj1
        Sheets("Replanteo").Cells(z, 9).Value = -d_eje_aguj2
        Sheets("Replanteo").Cells(z - 2, 8).Value = -d_semi_eje_aguj1
        Sheets("Replanteo").Cells(z - 2, 9).Value = -d_semi_eje_aguj2
        If Sheets("Replanteo").Cells(z - 4, 16).Value = semi_eje_aguj Then
            Sheets("Replanteo").Cells(z - 4, 8).Value = -d_semi_eje_aguj1
            Sheets("Replanteo").Cells(z - 4, 9).Value = -d_semi_eje_aguj2
        End If
    End If
ElseIf (Sheets("Replanteo").Cells(z - 2, 16).Value = eje_aguj Or Sheets("Replanteo").Cells(z - 2, 16).Value = eje_pf & " + " & eje_aguj Or Sheets("Replanteo").Cells(z - 2, 16).Value = anc_pf & " + " & eje_aguj) And (Sheets("Replanteo").Cells(z, 16).Value = semi_eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & semi_eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = eje_pf & " + " & semi_eje_aguj) Then
    If Sheets("Replanteo").Cells(z - 1, 35).Value = "I" Then
        Sheets("Replanteo").Cells(z - 2, 8).Value = d_eje_aguj1
        Sheets("Replanteo").Cells(z - 2, 9).Value = d_eje_aguj2
        Sheets("Replanteo").Cells(z, 8).Value = d_semi_eje_aguj1
        Sheets("Replanteo").Cells(z, 9).Value = d_semi_eje_aguj2
            If Sheets("Replanteo").Cells(z + 2, 16).Value = semi_eje_aguj Then
                Sheets("Replanteo").Cells(z + 2, 8).Value = d_semi_eje_aguj1
                Sheets("Replanteo").Cells(z + 2, 9).Value = d_semi_eje_aguj2
            End If
    Else
        Sheets("Replanteo").Cells(z - 2, 8).Value = -d_eje_aguj1
        Sheets("Replanteo").Cells(z - 2, 9).Value = -d_eje_aguj2
        Sheets("Replanteo").Cells(z, 8).Value = -d_semi_eje_aguj1
        Sheets("Replanteo").Cells(z, 9).Value = -d_semi_eje_aguj2
            If Sheets("Replanteo").Cells(z + 2, 16).Value = semi_eje_aguj Or Sheets("Replanteo").Cells(z + 2, 16).Value = eje_pf & " + " & semi_eje_aguj Then
                Sheets("Replanteo").Cells(z + 2, 8).Value = -d_semi_eje_aguj1
                Sheets("Replanteo").Cells(z + 2, 9).Value = -d_semi_eje_aguj2
            End If
    End If
    'z = z + 2
End If

If IsEmpty(Sheets("Replanteo").Cells(z, 8).Value) Then

    algo = 0
    
End If
Call txt.progress("6", "14", "Descentramiento", Sheets("Replanteo").Cells(z, 33).Value - inicio, final - inicio)


'//
'// Incrementar fila del replanteo
'//

z = z + 2
Wend
End Sub

