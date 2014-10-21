Attribute VB_Name = "pk_real"
Sub convertir_LT()       ' función que convierte los PK lineales a los de trazado
Dim ipk As Integer
Dim ceros As String
Dim beta As Double
Dim matriz_pk(500) As Double
With ActiveWorkbook
With .Sheets(6)
        For fila = 2 To 122
            matriz_pk(Sheets(6).Cells(fila, 1).Value) = Sheets(6).Cells(fila, 2).Value
        Next
    End With
End With
h = 10

While Not IsEmpty(Sheets(1).Cells(h, 33).Value)
ipk = 0
    If 55453.6631 <= Sheets(1).Cells(h, 33).Value And Sheets(1).Cells(h, 33).Value < 56453.5677 Then
        If Round((Sheets(1).Cells(h, 33).Value - 55453.6631), 2) < 100 Then
            If Round((Sheets(1).Cells(h, 33).Value - 55453.6631), 2) < 10 Then
                ceros = "00"
            End If
        ceros = "0"
        Else
        ceros = ""
        End If
        Sheets(1).Cells(h, 3).Value = "55bis" & "+" & ceros & Round((Sheets(1).Cells(h, 33).Value - 55453.6631), 2)
    Else
        Do While matriz_pk(ipk) < Sheets(1).Cells(h, 33).Value
            ipk = ipk + 1
        Loop
    ipk = ipk - 1
    Sheets(1).Cells(h, 3).Value = (1000 * CDbl(ipk) + Sheets(1).Cells(h, 33).Value - matriz_pk(ipk))
    End If
h = h + 2
Wend
End Sub
