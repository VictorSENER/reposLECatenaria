Attribute VB_Name = "postes"
Public pol As Integer
Sub ubicacion_postes(catenaria As String, ventoso)

Dim a As Integer, h As Integer


pol = 3
Call formato.lenguaje(idioma)
a = 4
h = 10
While Sheets("Replanteo").Cells(h, 33).Value < final

    marcador = 0
    On Error Resume Next
    While pkini > ventoso(pol - 1)
        
    Wend
        
    If ventoso(pol - 1) >= Sheets("Replanteo").Cells(h, 33).Value Then
        Sheets("Vano").Range("A3:E20").ClearContents
        Call tabla_vanos.tabla_vanos(catenaria, pol, ventoso)
    Else
        pol = pol + 3
        Sheets("Vano").Range("A3:E20").ClearContents
        Call tabla_vanos.tabla_vanos(catenaria, pol, ventoso)
    End If
    
aqui:
    '//
    '// Inicializar variable al inicio de la rutina
    '//
    If h = 10 Then
        Sheets("Replanteo").Cells(h, 33) = inicio
    End If
    '//
    '// Rutina general del programa
    '// radio + vano + regulación vano + cantonamiento + punto singular + incrementar PK y fila
    '//
    k = radio.radio(h)
    vano_pri = vano.vano(Sheets("Replanteo").Cells(h, 6).Value, h)
    '///
    '/// Mejora rendimiento al no entrar tantas veces en el módulo regulación
    '///
    If vano_pri > Sheets("Replanteo").Cells(h - 1, 4).Value + dist_va_max And h <> 10 Then
        Sheets("Replanteo").Cells(h + 1, 4).Value = Sheets("Replanteo").Cells(h - 1, 4).Value + dist_va_max
    Else
        Sheets("Replanteo").Cells(h + 1, 4).Value = vano_pri
    End If
    '//
    '// Empezar a regular cuando se hayan realizado 3 bucles
    '//
    res = regulacion.long_restar(h, a)
    If res > 27 And h > 16 Then
        Call regulacion.regulacion(h, a)
    ElseIf res <> 0 And h > 16 Then
        Call regulacion.regulacion(h, a)
    End If
    Call punto_singular.sing(h, a, k)
    Call punto_singular.sing1(h, a, marcador, 0)
    h = h + 2
    Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h - 1, 4) + Sheets("Replanteo").Cells(h - 2, 33)
    
    Call txt.progress("1", "14", "Replanteo de los postes", Sheets("Replanteo").Cells(h, 33).Value - inicio, final - inicio)

Wend

End Sub
