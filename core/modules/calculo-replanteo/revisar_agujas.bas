Attribute VB_Name = "revisar_agujas"
Sub revisar_agujas()
fila = 10
While Not IsEmpty(Sheets("Replanteo").Cells(fila, 33).Value)
    If Sheets("Replanteo").Cells(fila, 16).Value = eje_aguj And Sheets("Replanteo").Cells(fila - 2, 16).Value = semi_eje_aguj _
    And Sheets("Replanteo").Cells(fila - 4, 16).Value = "" Then
        If Sheets("Replanteo").Cells(fila - 1, 4).Value <= 31.5 Then
            Sheets("Replanteo").Cells(fila - 4, 16).Value = semi_eje_aguj
            Sheets("Replanteo").Cells(fila - 6, 16).Value = anc_aguj
            Sheets("Replanteo").Cells(fila - 6, 24).Value = "C2"
        Else
            Sheets("Replanteo").Cells(fila - 4, 16).Value = anc_aguj
            Sheets("Replanteo").Cells(fila - 4, 24).Value = "C2"
        End If
    ElseIf Sheets("Replanteo").Cells(fila, 16).Value = eje_aguj And Sheets("Replanteo").Cells(fila + 2, 16).Value = semi_eje_aguj _
    And Sheets("Replanteo").Cells(fila + 4, 16).Value = "" Then
        If Sheets("Replanteo").Cells(fila + 1, 4).Value <= 31.5 Then
            Sheets("Replanteo").Cells(fila + 4, 16).Value = semi_eje_aguj
            Sheets("Replanteo").Cells(fila + 6, 16).Value = anc_aguj
            Sheets("Replanteo").Cells(fila + 6, 24).Value = "C2"
        Else
            Sheets("Replanteo").Cells(fila + 4, 16).Value = anc_aguj
            Sheets("Replanteo").Cells(fila + 4, 24).Value = "C2"
        End If
    
    End If
    fila = fila + 2
Wend
End Sub
