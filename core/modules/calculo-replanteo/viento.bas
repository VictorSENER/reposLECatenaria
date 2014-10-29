Attribute VB_Name = "viento"
Function viento() As Variant
Dim ventoso() As Variant
'///
'///recolectar los tramos ventosos o no ventosos
'///
aloc = 3
pol = 0
While Not IsEmpty(Sheets("Extra").Cells(aloc, 11).Value)
    pol = pol + 3
    ReDim Preserve ventoso(1 To pol)
    ventoso(pol - 2) = Sheets("Extra").Cells(aloc, 9).Value
    ventoso(pol - 1) = Sheets("Extra").Cells(aloc, 10).Value
    ventoso(pol) = Sheets("Extra").Cells(aloc, 11).Value
    aloc = aloc + 1
Wend

viento = ventoso
End Function
