Attribute VB_Name = "descentramiento"
'//
'// Rutina destinada a elegir el descentramiento correspondiente al radio en la tabla de vanos
'//
Sub desc()
Dim z As Integer, m As Integer
Dim descentramiento As Double, rady As Double
'//
'// Inicializar variables
'//
z = 10
tierra = 3000
'//
'// Mientras no lleguemos al final del replanteo
'//
While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
'//
'// Inicializar variable local
'//
m = 3

If Sheets(1).Cells(z, 6).Value < 0 Then
    rady = Sheets(1).Cells(z, 6).Value * (-1)
Else
    rady = Sheets(1).Cells(z, 6).Value
End If
'//
'// Buscar en que fila de la hoja 2 se encuentra el radio que buscamos
'//
If Not IsEmpty(Sheets(1).Cells(z, 6).Value) Then
    While rady < Sheets(2).Cells(m, 3).Value
        m = m + 1
    Wend
Else
        m = 3
End If
descentramiento = Sheets(2).Cells(m, 5).Value
'//
'// el descentramiento varia segun el sentido de giro de la curva
'// el descentramiento anterior y posterior varian de signo en recta
'//
If Not IsEmpty(Sheets(1).Cells(z, 6).Value) Then
    If Sheets(1).Cells(z, 6).Value >= 0 Then
        Sheets(1).Cells(z, 8).Value = descentramiento
    Else
      Sheets(1).Cells(z, 8).Value = -descentramiento
    End If
ElseIf IsEmpty(Sheets(1).Cells(z, 6).Value) Then
    If Sheets(1).Cells(z - 2, 8).Value < 0 Then
        Sheets(1).Cells(z, 8).Value = -descentramiento
    Else
        Sheets(1).Cells(z, 8).Value = descentramiento
    End If
End If
'//
'//insertar descentramiento en seccionamientos de lamina de aire y compensación
'//
If Sheets(1).Cells(z, 16).Value = "Inter.Section." Then
    If Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value >= 0 And Sheets(1).Cells(z + 4, 6).Value >= 0 _
    And Sheets(1).Cells(z - 2, 8).Value > 0 Then
        Sheets(1).Cells(z, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z, 9).Value = -0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 2, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 2, 9).Value = -0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 4, 8).Value = 0.65 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 4, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z - 2, 47).Value = "Normal"
    ElseIf Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value >= 0 And Sheets(1).Cells(z + 4, 6).Value >= 0 _
      And Sheets(1).Cells(z - 2, 8).Value < 0 Then
            Sheets(1).Cells(z, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = 0.65 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = -0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = -0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Inverso"
    ElseIf Sheets(1).Cells(z, 6).Value < 0 And Sheets(1).Cells(z + 2, 6).Value < 0 And Sheets(1).Cells(z + 4, 6).Value < 0 _
    And Sheets(1).Cells(z - 2, 8).Value < 0 Then
        Sheets(1).Cells(z, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z, 9).Value = 0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 2, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 2, 9).Value = 0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 4, 8).Value = -0.65 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 4, 9).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z - 2, 47).Value = "Normal"
    ElseIf Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value >= 0 And Sheets(1).Cells(z + 4, 6).Value < 0 _
    And Sheets(1).Cells(z - 2, 8).Value > 0 Then
        Sheets(1).Cells(z, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z, 9).Value = -0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 2, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 2, 9).Value = -0.15 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 4, 8).Value = -0.65 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z + 4, 9).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
        Sheets(1).Cells(z - 2, 47).Value = "Normal"
    Else
        Sheets(1).Cells(z, 9).Value = "ERROR"
        Sheets(1).Cells(z + 2, 9).Value = "ERROR"
        Sheets(1).Cells(z + 4, 9).Value = "ERROR"
    End If
z = z + 4
End If

If Sheets(1).Cells(z, 16).Value = "Inter.Chevau." Then
      If Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value >= 0 And Sheets(1).Cells(z + 4, 6).Value >= 0 _
      And Sheets(1).Cells(z - 2, 8).Value > 0 Then
            Sheets(1).Cells(z, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = 0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = 0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = 0.45 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Normal"
    ElseIf Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value >= 0 And Sheets(1).Cells(z + 4, 6).Value >= 0 _
      And Sheets(1).Cells(z - 2, 8).Value < 0 Then
            Sheets(1).Cells(z, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = 0.45 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = 0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = 0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Inverso"
       ElseIf Sheets(1).Cells(z, 6).Value <= 0 And Sheets(1).Cells(z + 2, 6).Value <= 0 And Sheets(1).Cells(z + 4, 6).Value <= 0 _
      And Sheets(1).Cells(z - 2, 8).Value < 0 Then
            Sheets(1).Cells(z, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = -0.45 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Normal"
        ElseIf Sheets(1).Cells(z, 6).Value <= 0 And Sheets(1).Cells(z + 2, 6).Value <= 0 And Sheets(1).Cells(z + 4, 6).Value <= 0 _
      And Sheets(1).Cells(z - 2, 8).Value > 0 Then
            Sheets(1).Cells(z, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = -0.45 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Inverso"
        ElseIf Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value <= 0 And Sheets(1).Cells(z + 4, 6).Value <= 0 _
      And Sheets(1).Cells(z - 2, 8).Value > 0 Then
            Sheets(1).Cells(z, 8).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = 0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = -0.45 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Normal"
       ElseIf Sheets(1).Cells(z, 6).Value <= 0 And Sheets(1).Cells(z + 2, 6).Value <= 0 And Sheets(1).Cells(z + 4, 6).Value >= 0 _
      And Sheets(1).Cells(z - 2, 8).Value < 0 Then
            Sheets(1).Cells(z, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z, 9).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 8).Value = -0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 2, 9).Value = -0.05 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 8).Value = 0.45 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z + 4, 9).Value = 0.25 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            Sheets(1).Cells(z - 2, 47).Value = "Normal"
        Else
        Sheets(1).Cells(z, 9).Value = "ERROR"
        Sheets(1).Cells(z + 2, 9).Value = "ERROR"
        Sheets(1).Cells(z + 4, 9).Value = "ERROR"
        End If
    If Sheets(1).Cells(z, 6).Value = 0 And Sheets(1).Cells(z + 2, 6).Value = 0 And Sheets(1).Cells(z + 4, 6).Value = 0 _
      And Sheets(1).Cells(z - 2, 8).Value > 0 Then
            Sheets(1).Cells(z + 6, 8).Value = -0.2 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            z = z + 2
    ElseIf Sheets(1).Cells(z, 6).Value >= 0 And Sheets(1).Cells(z + 2, 6).Value >= 0 And Sheets(1).Cells(z + 4, 6).Value >= 0 _
      And Sheets(1).Cells(z - 2, 8).Value < 0 Then
            Sheets(1).Cells(z + 6, 8).Value = 0.2 '!!!!!!!!!!!!!!!!! debe ser variable desde BBDD
            z = z + 2
    End If
z = z + 4
End If



'//
'//insertar descentramiento en agujas
'//
If Sheets(1).Cells(z, 16).Value = "Axe.Aigu." And Sheets(1).Cells(z - 2, 16).Value = "Inter.Aigu." Then
    If Sheets(1).Cells(z + 1, 35).Value = "I" Then
        Sheets(1).Cells(z, 9).Value = Sheets(1).Cells(z, 8).Value - d_max_re
        Sheets(1).Cells(z - 2, 9).Value = Sheets(1).Cells(z - 2, 8).Value - d_max_re
    Else
        Sheets(1).Cells(z, 9).Value = Sheets(1).Cells(z, 8).Value + d_max_re
        Sheets(1).Cells(z - 2, 9).Value = Sheets(1).Cells(z - 2, 8).Value + d_max_re
    End If
ElseIf Sheets(1).Cells(z - 2, 16).Value = "Axe.Aigu." And Sheets(1).Cells(z, 16).Value = "Inter.Aigu." Then
    If Sheets(1).Cells(z - 1, 35).Value = "I" Then
        Sheets(1).Cells(z - 2, 9).Value = Sheets(1).Cells(z - 2, 8).Value - d_max_re
        Sheets(1).Cells(z, 9).Value = Sheets(1).Cells(z, 8).Value - d_max_re
    Else
        Sheets(1).Cells(z - 2, 9).Value = Sheets(1).Cells(z - 2, 8).Value + d_max_re
        Sheets(1).Cells(z, 9).Value = Sheets(1).Cells(z, 8).Value + d_max_re
    End If
End If
'//
'// Incrementar fila del replanteo
'//

z = z + 2
Wend
End Sub

