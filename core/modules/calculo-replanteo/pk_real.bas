Attribute VB_Name = "pk_real"
Sub convertir_LT()       ' función que convierte los PK lineales a los de trazado
Dim ipk As Double
Dim ceros As String
Dim beta As Double
Dim matriz_pk(500) As Double
Dim pk_nolineal() As Double
'///
'/// inicializar variables
'///

fila = 2
h = 10
polil = 0
a = 2
ipk = 0
'///
'/// recoger datos tablas pk_reales
'///
While Not IsEmpty(Sheets("Pk real").Cells(a, 1).Value)
    If Sheets("Pk real").Cells(a, 1).Value = Sheets("Pk real").Cells(a - 1, 1).Value Then
        polil = polil + 3
        ReDim Preserve pk_nolineal(1 To polil)
        pk_nolineal(polil - 2) = Sheets("Pk real").Cells(a, 2).Value
        pk_nolineal(polil - 1) = Sheets("Pk real").Cells(a + 1, 2).Value
        pk_nolineal(polil) = Sheets("Pk real").Cells(a, 1).Value
    End If
    matriz_pk(Sheets("Pk real").Cells(fila, 1).Value) = Sheets("Pk real").Cells(fila, 2).Value
    fila = fila + 1
    a = a + 1
Wend
pola = 1
a = 2
If IsEmpty(Sheets("Pk real").Cells(a, 2).Value) Then
        polil = polil + 3
        ReDim Preserve pk_nolineal(1 To polil)
        pk_nolineal(polil - 2) = 0
        pk_nolineal(polil - 1) = 0
        pk_nolineal(polil) = 0
        ipk = Sheets("Pk real").Cells(fila, 1).Value
        matriz_pk(0) = 0
End If
'///
'/// ¿fin replanteo?
'///

While Not IsEmpty(Sheets("Replanteo").Cells(h, 33).Value)
    '///
    '/// caso pk repetido
    '///
    If pk_nolineal(pola) <= Sheets("Replanteo").Cells(h, 33).Value And Sheets("Replanteo").Cells(h, 33).Value < pk_nolineal(pola + 1) Then
        If Round((Sheets("Replanteo").Cells(h, 33).Value - pk_nolineal(pola)), 2) < 100 Then
            If Round((Sheets("Replanteo").Cells(h, 33).Value - pk_nolineal(pola)), 2) < 10 Then
                ceros = "00"
            End If
        ceros = "0"
        Else
        ceros = ""
        End If
        Sheets("Replanteo").Cells(h, 3).Value = pk_nolineal(pola + 2) & "bis" & "+" & ceros & Round((Sheets("Replanteo").Cells(h, 33).Value - pk_nolineal(pola)), 2)

    '///
    '/// caso normal
    '///
    Else
         Do While matriz_pk(ipk) < Sheets("Replanteo").Cells(h, 33).Value And Not IsEmpty(Sheets("Pk real").Cells(a, 2).Value)
                ipk = ipk + 1
         Loop
        If IsEmpty(Sheets("Pk real").Cells(a, 2).Value) Then
            Sheets("Replanteo").Cells(h, 3).Value = Sheets("Replanteo").Cells(h, 33).Value
        Else
            Sheets("Replanteo").Cells(h, 3).Value = (1000 * CDbl(ipk - 1) + Sheets("Replanteo").Cells(h, 33).Value - matriz_pk(ipk - 1))
        End If
    End If
    '///
    '/// incrementar datos pk repetido
    '///
    If Sheets("Replanteo").Cells(h, 33).Value > pk_nolineal(pola + 1) And pola + 2 < polil Then
        pola = pola + 3
    End If
    
    Set text = a_text.CreateTextFile(dir_progress)
    text.WriteLine "3" & "/" & "14" & "/" & "Convertir el PK lineal a PK de trazado" & "/" & Sheets("Replanteo").Cells(h, 33).Value & "/" & final
    text.Close
'///
'/// incrementar fila repanteo
'///
h = h + 2
Wend
End Sub
