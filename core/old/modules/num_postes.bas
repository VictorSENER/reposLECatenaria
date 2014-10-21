Attribute VB_Name = "num_postes"
'//
'// Rutina destinada a numerar los postes
'//
Sub postes(nombre_catVB)
Dim w As Integer, bis As Integer
Dim z As Integer
'//
'// Inicializar variables y cargar los datos de catenaria
'//
Call cargar.datos_acces(nombre_catVB)

bis = 1
z = 10

If Sheets(1).Cells(z, 30).Value = "G" Then
    w = 1
Else
    w = 2
End If


'//
'// Final de replanteo?
'//
While Not IsEmpty(Sheets(1).Cells(z, 33).Value)

'//
'// Caso particular de existencia de PK BIS
'//
If 55453.6631 <= Sheets(1).Cells(z, 33).Value And Sheets(1).Cells(z, 33).Value < 56453.5677 Then
    If Not (55453.6631 <= Sheets(1).Cells(z, 33).Value And Sheets(1).Cells(z, 33).Value < 56453.5677) Then
    Sheets(1).Cells(z, 1) = (Sheets(1).Cells(z, 3).Value \ 1000) & "-" & w
    Sheets(1).Cells(z, 32) = w
    Sheets(1).Cells(z, 31) = (Sheets(1).Cells(z, 3).Value \ 1000)
    w = 1
    Else
    Sheets(1).Cells(z, 32).Value = bis
    Sheets(1).Cells(z, 1).Value = "55bis" & "-" & bis
    Sheets(1).Cells(z, 31).Value = "55bis"
    bis = bis + 2
    w = 0
    End If
Else
'//
'// se comparan PK para saber si se debe seguir contando o empezar de 0 la numeraci�n del poste
'// w = n� de poste
'//
On Error Resume Next
    If Err Then
        GoTo aqui
    End If
If w = 0 Then
    w = 1
    GoTo aqui
End If
If (Sheets(1).Cells(z + 2, 3).Value \ 1000) <= (Sheets(1).Cells(z, 3).Value \ 1000) Then
    Sheets(1).Cells(z, 1).Value = (Sheets(1).Cells(z, 3).Value \ 1000) & "-" & w
    Sheets(1).Cells(z, 32).Value = w
    Sheets(1).Cells(z, 31).Value = (Sheets(1).Cells(z, 3).Value \ 1000)
    w = w + 2
Else
aqui:
    Sheets(1).Cells(z, 1).Value = (Sheets(1).Cells(z, 3).Value \ 1000) & "-" & w
    Sheets(1).Cells(z, 32).Value = w
    Sheets(1).Cells(z, 31).Value = (Sheets(1).Cells(z, 3).Value \ 1000)
    If Sheets(1).Cells(z, 30).Value = "G" Then
        w = 1
    Else
        w = 2
    End If
    
End If

End If

'//
'// Incrementar fila del replanteo
'//

If Sheets(1).Cells(z, 30).Value <> Sheets(1).Cells(z + 2, 30).Value Then
    If w Mod 2 = 0 Then
        w = w - 1
    Else
        w = w + 1
    End If
End If

z = z + 2
Wend
End Sub
