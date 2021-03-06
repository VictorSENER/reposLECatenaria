Attribute VB_Name = "principal"
'//
'// declaración de variables publicas
'//
Public dist_va_max As Double
Public dist_max_canton As Double, va_max_sm As Double, va_max_tunel As Double
Public inc_norm_va As Double, va_max As Double
Public inicio As Double, r_re As Double
Public uno As Integer

'//
'// Función principal. Es la responsable de la rutina general y de la comunicación con VB Studio
'//
Function principal(inicioVB, hVB, wVB, kVB, aVB, bVB, cVB, r_reVB, dist_va_maxVB, _
inc_norm_vaVB, va_max_tunelVB, va_maxVB, dist_max_cantonVB, va_max_smVB) As Long()
'//
'// Recolección de datos
'//
inicio = inicioVB
h = hVB
w = wVB
k = kVB
a = aVB
b = bVB
C = cVB
r_re = r_reVB
dist_va_max = dist_va_maxVB
inc_norm_va = inc_norm_vaVB
va_max_tunel = va_max_tunelVB
va_max = va_maxVB
dist_max_canton = dist_max_cantonVB
va_max_sm = va_max_smVB
'//
'// Inicializar variable al inicio de la rutina
'//
If h = 10 Then
    uno = 1
    Sheets(1).Cells(10, 33) = inicio
    'Call cantonamiento.canton_durante(h, C, k, a)
End If

'//
'// Rutina general del programa
'// radio + vano + regulación vano + cantonamiento + punto singular + incrementar PK y fila
'//
Call radio.radio(h)
Sheets(1).Cells(h + 1, 4).Value = vano.vano(Sheets(1).Cells(h, 6).Value)
'//
'// Empezar a regular cuando se hayan realizado 3 bucles
'//
If h > 16 Then
    Call regulacion.regulacion(h, a, b, k)
    If va_max <> va_max_sm Then
        Call cantonamiento.canton_durante(h, C, k, a)
    End If
End If
Call punto_singular.sing(h, k, a, b)
Call punto_singular.sing1(h, a)
h = h + 2
Sheets(1).Cells(h, 33).Value = Sheets(1).Cells(h - 1, 4) + Sheets(1).Cells(h - 2, 33)
'//
'// Declaración de variable y comunicación con VB Studio
'//
Dim x(7) As Long
    x(0) = inicio
    x(1) = h
    x(2) = w
    x(3) = k
    x(4) = a
    x(5) = b
    x(6) = C
    x(7) = Sheets(1).Cells(h, 33).Value
principal = x
End Function
