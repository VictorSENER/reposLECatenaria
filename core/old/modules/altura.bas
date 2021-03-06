Attribute VB_Name = "altura"
'//
'// Rutina destinada a calcular los incrementos de altura en puntos singulares
'//
Sub altura(nombre_catVB)
Dim Amx As Double, Amx1 As Double
Dim z As Integer, aloc As Integer, s As Integer
'//
'// Cargar variables catenaria
'//
Call cargar.datos_acces(nombre_catVB)
'//
'// inicializar variables locales
'//
z = 10
aloc = 3
'//
'// inicio de la rutina principal
'//
While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
    '//
    '// Encontrar el punto singular para cada PK
    '//
    While (Sheets(1).Cells(z, 33).Value >= Sheets(4).Cells(aloc, 21).Value) And Sheets(4).Cells(aloc, 23).Value <> "FINAL"
        aloc = aloc + 1
    Wend
    '//
    '// Actuar si pasamos por un paso a nivel (se debe ir incrementando la altura hasta
    '// la altura m?xima y decrementarla hasta la nominal al pasar el paso a nivel)
    '//
    If Sheets(4).Cells(aloc - 1, 1) = ("P.N.") And Sheets(1).Cells(z, 33).Value > Sheets(4).Cells(aloc - 1, 2).Value And Sheets(1).Cells(z - 2, 33).Value < Sheets(4).Cells(aloc - 1, 2).Value Then
        '//
        '// Inicializar altura a valor maximo en PK actual
        '// Calculo del primer decremento (mas peque?o) hacia PK anteriores
        '//
        s = z
        Amx = alt_max
        s = s - 2
        Sheets(1).Cells(s, 10).Value = alt_max
        Amx = alt_max - Abs((((inc_max_alt_hc / 2) * Sheets(1).Cells(s - 1, 4).Value) / 1000) / 100) * 100
        '//
        '// Calculo de la altura a decrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx >= alt_nom And Not IsEmpty(Sheets(1).Cells(s, 33).Value)
            s = s - 2
            Sheets(1).Cells(s, 10).Value = Amx
            Amx = Amx - Int(((inc_max_alt_hc * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                If Amx < alt_nom Then
                    Amx1 = Amx + Int(((inc_max_alt_hc * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                        If ((Amx1 - alt_nom) * 1000) / Sheets(1).Cells(s - 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s - 2
                            Amx1 = Amx1 - Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                            Sheets(1).Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
        '//
        '// Inicializar altura a valor maximo en PK actual
        '// Calculo del primer decremento (mas peque?o) hacia PK siguientes
        '//
        s = z
        Sheets(1).Cells(s, 10).Value = alt_max
        Amx = alt_max - Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a decrementar hacia PK siguientes hasta llegar a la altura nominal
        '//
        While Amx >= alt_nom And Not IsEmpty(Sheets(1).Cells(s, 33).Value)
            s = s + 2
            Sheets(1).Cells(s, 10).Value = Amx
            Amx = Amx - Int(((inc_max_alt_hc * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                If Amx < alt_nom Then
                    Amx1 = Amx + Int(((inc_max_alt_hc * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                        If ((Amx1 - alt_nom) * 1000) / Sheets(1).Cells(s + 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s + 2
                            Amx1 = Amx1 - Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                            Sheets(1).Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
    '//
    '// Actuar si pasamos por un paso superior bajo (se debe ir decrementando la altura hasta
    '// la altura m?nima al llegar al paso e incrementar la altura hasta la nominal una vez pasado)
    '//
    ElseIf (Sheets(1).Cells(z, 33).Value >= Sheets(4).Cells(aloc - 1, 2).Value And Sheets(1).Cells(z, 33).Value <= Sheets(4).Cells(aloc - 1, 21).Value _
    And Sheets(4).Cells(aloc - 1, 1) = "7 > P.S. > 5,2 m") Or (Sheets(1).Cells(z - 2, 33).Value < Sheets(4).Cells(aloc - 1, 2).Value And Sheets(1).Cells(z, 33).Value > Sheets(4).Cells(aloc - 1, 2).Value And _
    Sheets(4).Cells(aloc - 1, 1) = "7 > P.S. > 5,2 m") Then
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas peque?o) hacia PK anteriores
        '//
        s = z
        s = s - 2
        Sheets(1).Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a incrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx <= alt_nom And Not IsEmpty(Sheets(1).Cells(s, 33).Value)
            s = s - 2
            Sheets(1).Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                If Amx > alt_nom Then
                    Amx1 = Amx - ((inc_max_alt_hc * Sheets(1).Cells(s - 1, 4).Value) / 1000)
                        If ((alt_nom - Amx1) * 1000) / Sheets(1).Cells(s - 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s - 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                            Sheets(1).Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas peque?o) hacia PK siguientes
        '//
        s = z
        Sheets(1).Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a incrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx <= alt_nom And Not IsEmpty(Sheets(1).Cells(s, 33).Value)
            s = s + 2
            Sheets(1).Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                If Amx > alt_nom Then
                    Amx1 = Amx - Int(((inc_max_alt_hc * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                        If ((alt_nom - Amx1) * 1000) / Sheets(1).Cells(s + 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s + 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                            Sheets(1).Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
    '//
    '// Actuar si estamos dentro del tunel
    '//
    ElseIf (Sheets(1).Cells(z, 33).Value >= Sheets(4).Cells(aloc, 2).Value And Sheets(1).Cells(z, 33).Value <= Sheets(4).Cells(aloc, 21).Value _
    And Sheets(4).Cells(aloc, 1) = "Tunel") Then
    '//
    '// Actualizar altura a valor minimo
    '//
         Sheets(1).Cells(z, 10).Value = alt_min
    '//
    '// Actuar en salida de tunel
    '//
    ElseIf (Sheets(1).Cells(z, 33).Value >= Sheets(4).Cells(aloc - 1, 2).Value And Sheets(1).Cells(z - 2, 33).Value <= Sheets(4).Cells(aloc - 1, 21).Value _
    And Sheets(4).Cells(aloc - 1, 1) = "Tunel") Then
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas peque?o) hacia PK siguientes
        '//
        s = z
        Sheets(1).Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100

        '//
        '// Calculo de la altura a incrementar hacia PK siguientes hasta llegar a la altura nominal
        '//
        While Amx <= alt_nom And Not IsEmpty(Sheets(1).Cells(s, 33).Value)
            s = s + 2
            Sheets(1).Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                If Amx > alt_nom Then
                    Amx1 = Amx - Int(((inc_max_alt_hc * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                        If ((alt_nom - Amx1) * 1000) / Sheets(1).Cells(s + 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s + 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s + 1, 4).Value) / 1000) * 100) / 100
                            Sheets(1).Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
    '//
    '// Actuar en entrada de tunel
    '//
    ElseIf (Sheets(1).Cells(z, 33).Value <= Sheets(4).Cells(aloc, 2).Value And Sheets(1).Cells(z + 2, 33).Value > Sheets(4).Cells(aloc, 2).Value _
    And Sheets(4).Cells(aloc, 1) = "Tunel") Then
        '//
        '// Inicializar altura a valor minimo en PK actual
        '// Calculo del primer incremento (mas peque?o) hacia PK anteriores
        '//
        s = z
        Sheets(1).Cells(s, 10).Value = alt_min
        Amx = alt_min + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
        '//
        '// Calculo de la altura a incrementar hacia PK anteriores hasta llegar a la altura nominal
        '//
        While Amx <= alt_nom And Not IsEmpty(Sheets(1).Cells(s, 33).Value)
            s = s - 2
            Sheets(1).Cells(s, 10).Value = Amx
            Amx = Amx + Int(((inc_max_alt_hc * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                If Amx > alt_nom Then
                    Amx1 = Amx - ((inc_max_alt_hc * Sheets(1).Cells(s - 1, 4).Value) / 1000)
                        If ((alt_nom - Amx1) * 1000) / Sheets(1).Cells(s - 1, 4).Value > (inc_max_alt_hc / 2) Then
                            s = s - 2
                            Amx1 = Amx1 + Int((((inc_max_alt_hc / 2) * Sheets(1).Cells(s - 1, 4).Value) / 1000) * 100) / 100
                            Sheets(1).Cells(s, 10).Value = Amx1
                        End If
                End If
        Wend
    End If
'//
'// Incrementar una fila del replanteo
'//
z = z + 2
Wend
End Sub


