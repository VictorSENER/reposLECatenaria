Attribute VB_Name = "pendolado"
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'// The SendMessage function sends the specified message to a window or windows.
'// The function calls the window procedure for the specified window and does not
'// return until the window procedure has processed the message.
'// The PostMessage function, in contrast, posts a message to a thread?s message
'// queue and returns immediately.
'//
'// PARAMETERS:
'//
'// hwnd
'// Identifies the window whose window procedure will receive the message.
'// If this parameter is HWND_BROADCAST, the message is sent to all top-level
'// windows in the system, including disabled or invisible unowned windows,
'// overlapped windows, and pop-up windows; but the message is not sent to child windows.

'// Msg
'// Specifies the message to be sent.

'// wParam
'// Specifies additional message-specific information.

'// lParam
'// Specifies additional message-specific information.

'//////////////////////////////////////////////////////////////////////////
'// The IsWindow function determines whether the specified window handle
'// identifies an existing window.
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'// PARAMETERS:
'// hWnd
'// Specifies the window handle.

'//////////////////////////////////////////////////////////////////////////
'//
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As Long, ByVal bErase As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Function fncScreenUpdating(State As Boolean, Optional Window_hWnd As Long = 0)
Const WM_SETREDRAW = &HB
Const WM_PAINT = &HF

If Window_hWnd = 0 Then
    Window_hWnd = GetDesktopWindow()
Else
    If IsWindow(hwnd:=Window_hWnd) = False Then
        Exit Function
    End If
End If

If State = True Then
    Call SendMessage(hwnd:=Window_hWnd, wMsg:=WM_SETREDRAW, wParam:=1, lParam:=0)
    Call InvalidateRect(hwnd:=Window_hWnd, lpRect:=0, bErase:=True)
    Call UpdateWindow(hwnd:=Window_hWnd)
Else
    Call SendMessage(hwnd:=Window_hWnd, wMsg:=WM_SETREDRAW, wParam:=0, lParam:=0)
End If

End Function

Sub pendolado(nombre_catVB, ruta_replanteoVB)
Dim fso As Object
Dim dist(100) As Double, fuerza(40) As Double, mom(40) As Double
Dim n_pend As Long, cont As Long, n(40) As Long
Dim fl_hc(40) As Double, fl_sust(40) As Double, el_hc_der(40) As Double, el_hc_izq(40) As Double
Dim fuerza_der(40) As Double, dist_der(40) As Double, l_pend_der(40) As Double, acum(40) As Double
Dim dist_ap_prim_pend_izq As Double, dist_ap_prim_pend_der As Double, dist_prim_seg_pend_izq As Double
Dim dist_prim_seg_pend_der As Double, va_var_sla As Double, l_sup_pend As Double, l_inf_pend As Double
Dim va As Double, el_hc_ini As Double, el_hc_fin As Double, p_aisl As Double, var_dist As Double
Dim p_sust_var As Double, va_min As Double, var_0 As Double, va_var As Double
Dim var_1 As Double, p_var_0 As Double, p_var_1 As Double, p_var_2 As Double, p_var_3 As Double
Dim p_sust_ap As Double, x_var As Double, p_hc_ap As Double, p_pend_equip As Double, p_hc_ap_izq As Double
Dim p_hc_ap_der As Double, p_ap_tot_izq As Double, p_ap_tot_der_var As Double, p_ap_tot_der As Double
Dim t_horiz_sust As Double, var_comp As Double, d As String, e As String, st As Long
Dim a As Integer, dfijo As String, efijo As String, aisl As Integer, it As Integer, it_var As Integer
Dim npi As Integer, i As Integer, j As Integer, var_izq_der As Integer, b As Integer
Dim tip As String, PDFFijoFileName As String, tip_var As String, PSFileName As String
Dim PDFFileName As String, TXTFileName As String
Dim pk_ini_var As String, pk_fin_var As String, tip_pend As String, tip_pend_var As String


'//
'//LECTURA BASE DE DATOS
'//
Call cargar.datos_acces(nombre_catVB)

Dim myPDF As PdfDistiller
Set myPDF = New PdfDistiller
Set fso = CreateObject("Scripting.FileSystemObject")

Dim plantilla As Object
Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False
plantilla.Workbooks.Open "W:\223\D\D223041\IN_INFORMES\plantilla.xlsm"

a = 10
dfijo = Workbooks(1).Worksheets(1).Cells(a, 1).Value
efijo = Workbooks(1).Worksheets(1).Cells(a + 2, 1).Value
PDFFijoFileName = ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf"
aisl = 0
'//
'//SELECCI?N TIPOLOG?A PENDOLADO
'//

'//
'//S?lo un pendolado para un vano
'//
ini1:

Call cargar.datos_acces(nombre_catVB)

va = Workbooks(1).Sheets(1).Cells(a + 1, 4)

plantilla.Sheets(1).Range("D4:D9").ClearContents
plantilla.Sheets(1).Cells(6, 10).ClearContents
plantilla.Sheets(1).Cells(8, 10).ClearContents
plantilla.Sheets(1).Range("D12:D10001").ClearContents
plantilla.Sheets(1).Range("E12:E10001").ClearContents
plantilla.Sheets(1).Range("F12:F10001").ClearContents
plantilla.Sheets(1).Range("G11:G10001").ClearContents
plantilla.Sheets(1).Range("H11:H10001").ClearContents

plantilla.Sheets(1).Cells(4, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 4)


'//
'//Formato para que se muestre en el encabezado de cada ficha
'//
pk_ini_var = Int((Workbooks(1).Sheets(1).Cells(a, 3) + 1) / 1000) & "+" & Int(Workbooks(1).Sheets(1).Cells(a, 3) + 1)
pk_fin_var = Int((Workbooks(1).Sheets(1).Cells(a + 2, 3) + 1) / 1000) & "+" & Int(Workbooks(1).Sheets(1).Cells(a + 2, 3) + 1)

plantilla.Sheets(1).Name = pk_ini_var & " - " & pk_fin_var

'//
'//Lectura tipo de pendolado
'//

it = 0
st = 0

If n_hc = 2 Then

    dist_max_pend = dist_max_pend / 2

End If

tip_pend = Workbooks(1).Sheets(1).Cells(a + 1, 11)
tip_pend_var = Workbooks(1).Sheets(1).Cells(a + 1, 12)


If tip_pend_var <> "" Then

    it = 1
    
End If

If tip_pend <> "" And tip_pend_var <> "" Then

    plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 39)
    plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 41)
    plantilla.Sheets(1).Cells(7, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 40)
    plantilla.Sheets(1).Cells(8, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 42)
    dist_ap_prim_pend_izq = Workbooks(1).Sheets(1).Cells(a + 1, 43)
    dist_ap_prim_pend_der = Workbooks(1).Sheets(1).Cells(a + 1, 44)
    plantilla.Sheets(1).Cells(6, 10) = Workbooks(1).Sheets(1).Cells(a, 1)
    plantilla.Sheets(1).Cells(8, 10) = Workbooks(1).Sheets(1).Cells(a + 2, 1)
    
Else
    
    plantilla.Sheets(1).Cells(5, 4) = alt_cat
    plantilla.Sheets(1).Cells(6, 4) = alt_cat
    plantilla.Sheets(1).Cells(7, 4) = 0
    plantilla.Sheets(1).Cells(8, 4) = 0
    plantilla.Sheets(1).Cells(6, 10) = Workbooks(1).Sheets(1).Cells(a, 1)
    plantilla.Sheets(1).Cells(8, 10) = Workbooks(1).Sheets(1).Cells(a + 2, 1)
    dist_ap_prim_pend_izq = dist_ap_prim_pend
    dist_ap_prim_pend_der = dist_ap_prim_pend

End If

ini2:

If Workbooks(1).Worksheets(1).Cells(a, 1).Value = "" Then

    GoTo final

End If

plantilla.Sheets(1).Cells(7, 5) = tip_pend
  
'//
'//C?LCULO GEOM?TRICO
'//

el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
el_hc_fin = plantilla.Sheets(1).Cells(8, 4)


    dist_prim_seg_pend_izq = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 4
    
    If dist_prim_seg_pend_izq > 2.25 Then
    
        dist_prim_seg_pend_izq = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 8
    
    End If
    
    dist_prim_seg_pend_der = (va - 4.5 * (Int((va / 4.5) + 1) - 0.99)) / 4
    
    If dist_prim_seg_pend_der > 2.25 Then
    
        dist_prim_seg_pend_der = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 8
    
    End If


'//
'//Longitud de cabeza superior e inferior de la p?ndola
'//

'//l_sup_pend depende si el sustentador es de 153 36mm, si es de 93 34.55mm

l_sup_pend = 0.036
l_inf_pend = 0.0336

'//
'//Elecci?n aislador en caso de que haya
'//
If aisl = 1 Then
    
    If cola_anc = "Cer?mico" Then
            p_aisl = 15
            p_sust_var = p_sust * va
            p_sust = (p_sust_var + p_aisl) / va
    
    ElseIf cola_anc = "Sint?tico" Then
            p_aisl = 3
            p_sust_var = p_sust * va
            p_sust = (p_sust_var + p_aisl) / va
            
    ElseIf cola_anc = "Vidrio" Then
            p_aisl = 4.5
            p_sust_var = p_sust * va
            p_sust = (p_sust_var + p_aisl) / va
    
    End If
    
End If

'//
'//Se debe distinguir entre posible casos con elevaci?n a un lado, elevaci?n en ambos, etc.
'//

If el_hc_ini <> 0 And el_hc_fin <> 0 Then

    If el_hc_ini > el_hc_fin Then
        el_hc_ini = el_hc_ini - el_hc_fin
        el_hc_fin = 0
    Else
        el_hc_fin = el_hc_fin - el_hc_ini
        el_hc_ini = 0
    End If
    
End If



'PRUEBAAAAAAAAAAAAAS!!!!!!!!

'valores a leer del Sireca
p_sust = 1.407 'kg/m
p_hc = 0.95 'kg/m
p_pend = 0.101 'kg/m

n_hc = 2
t_sust = 1400 'kg
t_hc = 1000 'kg
fl_max_centro_va = va / 1000


'//
'//Flecha m?xima en centro de vano impuesta
'//

p_pend_equip = 0.13  '0.08 + 0.15 + 0.04 'kg
fl_max_centro_va = va / 1000
 
If (dist_ap_prim_pend_izq = dist_ap_prim_pend_der) And (el_hc_ini = el_hc_fin) Then

    dist_ap_prim_pend = dist_ap_prim_pend_izq
    dist_prim_seg_pend = dist_prim_seg_pend_izq

    va_min = 2 * (dist_ap_prim_pend + dist_prim_seg_pend)
    var_0 = 2 * (dist_max_pend - dist_prim_seg_pend) + va_min
    
    If va < va_min Then
    
        GoTo x
    
    Else
        
        If va_min <= va And va <= var_0 Then
        
            dist(1) = dist_ap_prim_pend
            dist(2) = (va - (2 * dist_ap_prim_pend)) / 2
            dist(3) = dist(2)
            dist(4) = dist(1)
            
        Else:
            va_var = va - va_min
                
                If (va_var / 2) < dist_max_pend Then
                    dist(1) = dist_ap_prim_pend
                    dist(2) = dist_prim_seg_pend
                    If va_var > dist_max_pend Then
                        dist(3) = va_var / 2
                        dist(4) = va_var / 2
                        dist(5) = dist(2)
                        dist(6) = dist(1)
                    Else
                        dist(3) = va_var
                        dist(4) = dist(2)
                        dist(5) = dist(1)
                    End If
                                
                ElseIf (va_var / 2) >= dist_max_pend Then
                        If (va_var) / dist_max_pend >= 1 Then
                            If Int((va_var / dist_max_pend)) = (va_var) / dist_max_pend Then
                                npi = Int((va_var / dist_max_pend))
                            
                            Else: npi = Int((va_var / dist_max_pend)) - 1
                            End If
                        Else: npi = 0
                        
                        End If
                        dist(1) = dist_ap_prim_pend
                        dist(2) = dist_prim_seg_pend
                            
                            If Int((va_var / dist_max_pend)) = (va_var) / dist_max_pend Then
                                dist(3) = dist_max_pend
                                npi = npi
                            Else: dist(3) = (va_var - dist_max_pend * npi) / 2
                            End If
                                       
                        i = 1
                        If Int((va_var / dist_max_pend)) = (va_var) / dist_max_pend Then
                            While i <= npi - 2
                                dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        Else
                            While i <= npi
                                dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        End If
                        dist(i + 3) = dist(3)
                        dist(i + 4) = dist(2)
                        dist(i + 5) = dist(1)
                
                End If
            
          End If
        
        i = 1
        j = 11
        While dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 7) = dist(i)
            var_dist = var_dist + plantilla.Sheets(1).Cells(j, 7)
            plantilla.Sheets(1).Cells(j, 8) = var_dist
            i = i + 1
            j = j + 2
        Wend
        
        i = 2
        j = 12
        While dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 4) = var_1 + 1
            plantilla.Sheets(1).Cells(9, 4) = var_1 + 1
            var_1 = var_1 + 1
            i = i + 1
            j = j + 2
        Wend
        
    End If

Else
    
    va_min = (dist_ap_prim_pend_izq + dist_ap_prim_pend_der + dist_prim_seg_pend_izq + dist_prim_seg_pend_der)
    var_0 = (2 * dist_max_pend - dist_prim_seg_pend_izq - dist_prim_seg_pend_der) + va_min
        
    End If
    
    If va < va_min Then
    
        GoTo x
    
    Else
        
        If va_min <= va And va <= var_0 Then
        
            dist(1) = dist_ap_prim_pend_izq
            dist(2) = (va - (2 * dist_ap_prim_pend)) / 2
            dist(3) = dist(2)
            dist(4) = dist_ap_prim_pend_der
        
        Else:
            va_var = va - va_min
                
                If (va_var / 2) < dist_max_pend Then
                    dist(1) = dist_ap_prim_pend_izq
                    dist(2) = dist_prim_seg_pend_izq
                    If va_var > dist_max_pend Then
                        dist(3) = va_var / 2
                        dist(4) = va_var / 2
                        dist(5) = dist_prim_seg_pend_der
                        dist(6) = dist_ap_prim_pend_der
                    Else
                        dist(3) = va_var
                        dist(4) = dist_prim_seg_pend_der
                        dist(5) = dist_ap_prim_pend_der
                    End If
                                
                ElseIf (va_var / 2) >= dist_max_pend Then
                        If (va_var) / dist_max_pend >= 1 Then
                            If Int((va_var / dist_max_pend)) = (va_var) / dist_max_pend Then
                                npi = Int((va_var / dist_max_pend))
                            
                            Else: npi = Int((va_var / dist_max_pend)) - 1
                            End If
                        Else: npi = 0
                        
                        End If
                        dist(1) = dist_ap_prim_pend_izq
                        dist(2) = dist_prim_seg_pend_izq
                            
                            If Int((va_var / dist_max_pend)) = (va_var) / dist_max_pend Then
                                dist(3) = dist_max_pend
                                npi = npi
                            Else: dist(3) = (va_var - dist_max_pend * npi) / 2
                            End If
                                       
                        i = 1
                        If Int((va_var / dist_max_pend)) = (va_var) / dist_max_pend Then
                            While i <= npi - 2
                                dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        Else
                            While i <= npi
                                dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        End If
                        dist(i + 3) = dist(3)
                        dist(i + 4) = dist_prim_seg_pend_der
                        dist(i + 5) = dist_ap_prim_pend_der
                                        
                End If
            
          End If
        
        i = 1
        j = 11
        var_dist = 0
        While dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 7) = dist(i)
            acum(i) = acum(i - 1) + plantilla.Sheets(1).Cells(j, 7)
            var_dist = var_dist + plantilla.Sheets(1).Cells(j, 7)
            plantilla.Sheets(1).Cells(j, 8) = var_dist
            i = i + 1
            j = j + 2
        Wend
        
        i = 2
        j = 12
        var_1 = 0
        While dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 4) = var_1 + 1
            plantilla.Sheets(1).Cells(9, 4) = var_1 + 1
            var_1 = var_1 + 1
            i = i + 1
            j = j + 2
        Wend
        
    End If
    
'//
'//Consideraci?n de la reducci?n de peso en caso de flecha intencional
'//
p_var_0 = p_hc * n_hc * dist(1)

If (el_hc_ini <> 0 Or el_hc_fin <> 0) Or (el_hc_ini = el_hc_fin And el_hc_ini <> 0) Then
    p_var_1 = 0
    
Else
    p_var_1 = (fl_max_centro_va * 2 * t_hc) / (((va - 2 * dist(1)) / 2) ^ 2)

End If

p_var_2 = p_hc - p_var_1

p_var_3 = (p_sust * va) / 2

p_sust_ap = (p_sust * va) / 2

'//
'//C?lculo de las elevaciones
'//

'//
'//Elevaci?n hilo de contacto (a un lado)
'//

If el_hc_fin <> 0 Or el_hc_ini <> 0 Then
i = 1
cont = 1
n_pend = 1

If el_hc_fin = 0 Then
    el_hc_fin = el_hc_ini
End If
        
    x_var = Sqr((el_hc_fin * 2 * t_hc) / p_hc)
    x_var = va - x_var
    
            While cont <= i And n_pend <= plantilla.Sheets(1).Cells(9, 4)
                n(cont) = 1
                cont = cont + 1
            
            If x_var < (n(1) * dist(1) + n(2) * dist(2) + n(3) * dist(3) + n(4) * dist(4) + n(5) * dist(5) + n(6) * dist(6) + n(7) * dist(7) + n(8) * dist(8) + n(9) * dist(9) + n(10) * dist(10) + n(11) * dist(11) + n(12) * dist(12) + n(13) * dist(13) + n(14) * dist(14) + n(15) * dist(15) + n(16) * dist(16) + n(17) * dist(17) + n(18) * dist(18) + n(19) * dist(19) + n(20) * dist(20) + n(21) * dist(21) + n(22) * dist(22) + n(23) * dist(23) + n(24) * dist(24) + n(25) * dist(25) + n(26) * dist(26) + n(27) * dist(27) + n(28) * dist(28)) Then
                
                el_hc_der(i) = (p_hc * (n(1) * dist(1) + n(2) * dist(2) + n(3) * dist(3) + n(4) * dist(4) + n(5) * dist(5) + n(6) * dist(6) + n(7) * dist(7) + n(8) * dist(8) + n(9) * dist(9) + n(10) * dist(10) + n(11) * dist(11) + n(12) * dist(12) + n(13) * dist(13) + n(14) * dist(14) + n(15) * dist(15) + n(16) * dist(16) + n(17) * dist(17) + n(18) * dist(18) + n(19) * dist(19) + n(20) * dist(20) + n(21) * dist(21) + n(22) * dist(22) + n(23) * dist(23) + n(24) * dist(24) + n(25) * dist(25) + n(26) * dist(26) + n(27) * dist(27) + n(28) * dist(28) - x_var * n(1)) ^ 2) / (2 * t_hc)
            
            Else
            
                el_hc_der(i) = 0
            
            End If
                         
            i = i + 1
            n_pend = n_pend + 1
                    
            Wend
                            
End If

'//
'//C?LCULO REACCIONES EN CADA P?NDOLA
'//

i = 1
n_pend = 1
p_hc_ap = 0
el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
el_hc_fin = plantilla.Sheets(1).Cells(8, 4)

If el_hc_ini <> 0 And el_hc_fin <> 0 And el_hc_ini <> el_hc_fin Then

    If el_hc_ini > el_hc_fin Then
        el_hc_ini = el_hc_ini - el_hc_fin
        el_hc_fin = 0
    Else
        el_hc_fin = el_hc_fin - el_hc_ini
        el_hc_ini = 0
    End If
    
End If

cont = 1
While cont <= 30
    n(cont) = 0
    cont = cont + 1
Wend

'//
'//C?lculo de las reacciones sin elevaci?n del hilo de contacto
'//

If el_hc_ini = 0 And el_hc_fin = 0 Or el_hc_ini = el_hc_fin Then
    While n_pend <= plantilla.Sheets(1).Cells(9, 4)
    
        If n_pend = 1 Then
        
            fuerza(i) = n_hc * p_hc * dist(i) + n_hc * p_var_2 * (dist(i + 1) / 2) + p_pend * ((dist(i) + dist(i + 1)) / 2) + p_pend_equip + n_hc * p_var_1 * ((va - 2 * dist(i)) / 2)
        
        ElseIf i = plantilla.Sheets(1).Cells(9, 4) Then
            
            fuerza(i) = n_hc * p_hc * dist(i + 1) + n_hc * p_var_2 * (dist(i) / 2) + p_pend * ((dist(i) + dist(i + 1)) / 2) + p_pend_equip + n_hc * p_var_1 * ((va - 2 * dist(i + 1)) / 2)
         
        Else
         
            fuerza(i) = n_hc * p_var_2 * (dist(i) / 2) + n_hc * p_var_2 * (dist(i + 1) / 2) + p_pend * ((dist(i) + dist(i + 1)) / 2) + p_pend_equip
            
        
        End If
    
        p_hc_ap = p_hc_ap + fuerza(i)
        i = i + 1
        n_pend = n_pend + 1
    
    Wend

'//
'//C?lculo de las reacciones con elevaci?n del hilo de contacto
'//
ElseIf el_hc_ini <> 0 Or el_hc_fin <> 0 And el_hc_ini <> el_hc_fin Then
       
    j = 11
    i = 1
    n_pend = 1
   
        If el_hc_fin = 0 Then
            el_hc_fin = el_hc_ini
            var_izq_der = 1
        End If
    
    While n_pend <= plantilla.Sheets(1).Cells(9, 4)
              
        fuerza(i) = 0
        If n_pend = 1 Then
                
              fuerza(i) = (dist(i) / 2) * (n_hc * p_hc + p_pend)
        
        End If
        
        If x_var >= plantilla.Sheets(1).Cells(j + 2, 8) Then
            
            fuerza(i) = fuerza(i) + n_hc * p_hc * ((dist(i) + dist(i + 1)) / 2) + p_pend * ((dist(i) + dist(i + 1)) / 2) + p_pend_equip
            
        ElseIf x_var > acum(i) And x_var < acum(i + 1) Then
            
            If (acum(i) + acum(i + 1)) / 2 < x_var Then
                    
                fuerza(i) = fuerza(i) + n_hc * p_hc * ((dist(i) + dist(i + 1)) / 2) + p_pend * ((dist(i) + dist(i + 1)) / 2) + p_pend_equip
                    
            Else
                    
                fuerza(i) = fuerza(i) + n_hc * p_hc * (dist(i) / 2) + n_hc * p_hc * ((x_var) - acum(i)) + p_pend * (dist(i) / 2) + p_pend * ((x_var) - acum(i)) + p_pend_equip
        
            End If
         
        ElseIf x_var > acum(i - 1) And x_var < acum(i) Then
            
            If (acum(i - 1) + acum(i)) / 2 < x_var Then
                
                fuerza(i) = n_hc * p_hc * (x_var - (acum(i - 1) + dist(i) / 2)) + p_pend * (x_var - (acum(i - 1) + dist(i) / 2)) + p_pend_equip
                
            Else
            
                fuerza(i) = p_pend * (dist(i) + dist(i + 1) / 2) + p_pend_equip
            
            End If
            
        ElseIf x_var < acum(i - 1) Then
        
            fuerza(i) = p_pend * (dist(i) + dist(i + 1) / 2) + p_pend_equip
            
        End If
              
        If plantilla.Sheets(1).Cells(j, 8) <= va / 2 Then
            p_hc_ap_izq = p_hc_ap_izq + fuerza(i)
            
        Else
            p_hc_ap_der = p_hc_ap_der + fuerza(i)
        
         End If
            i = i + 1
            n_pend = n_pend + 1
            j = j + 2
    
    Wend
                    
End If

'//
'//B?squeda de las reacciones en los apoyos
'//

i = 1
cont = 1
n_pend = plantilla.Sheets(1).Cells(9, 4)

While cont <= plantilla.Sheets(1).Cells(9, 4)
    n(cont) = 1
    cont = cont + 1
Wend

p_ap_tot_izq = p_sust_ap + (n(1) * fuerza(1) * acum(1) + n(2) * fuerza(2) * acum(2) + n(3) * fuerza(3) * acum(3) + n(4) * fuerza(4) * acum(4) + n(5) * fuerza(5) * acum(5) + n(6) * fuerza(6) * acum(6) + n(7) * fuerza(7) * acum(7) + n(8) * fuerza(8) * acum(8) + n(9) * fuerza(9) * acum(9) + n(10) * fuerza(10) * acum(10) + n(11) * fuerza(11) * acum(11) + n(12) * fuerza(12) * acum(12) + n(13) * fuerza(13) * acum(13) + n(14) * fuerza(14) * acum(14) + n(15) * fuerza(15) * acum(15) + n(16) * fuerza(16) * acum(16) + n(17) * fuerza(17) * acum(17) + n(18) * fuerza(18) * acum(18) + n(19) * fuerza(19) * acum(19) + n(20) * fuerza(20) * acum(20) + n(21) * fuerza(21) * acum(21) + n(22) * fuerza(22) * acum(22) + n(23) * fuerza(23) * acum(23) + n(24) * fuerza(24) * acum(24)) / va

i = 1
While i <= plantilla.Sheets(1).Cells(9, 4)
    p_ap_tot_der_var = p_ap_tot_der_var + n(i) * fuerza(i) * (acum(n_pend + 1) - acum(i))
        
    i = i + 1
    
Wend

p_ap_tot_der = p_sust_ap + p_ap_tot_der_var / va

cont = 1
While cont <= 30
    n(cont) = 0
    cont = cont + 1
Wend

'//
'//Distintas alturas de catenaria
'//
         
If plantilla.Sheets(1).Cells(5, 4) <> plantilla.Sheets(1).Cells(6, 4) Then
    
    p_ap_tot_izq = (p_ap_tot_der) + (t_sust * (plantilla.Sheets(1).Cells(6, 4) - plantilla.Sheets(1).Cells(5, 4)) / va)
    p_ap_tot_der = (p_ap_tot_der) + (t_sust * (plantilla.Sheets(1).Cells(5, 4) - plantilla.Sheets(1).Cells(6, 4)) / va)
    
End If

'//
'//C?LCULO MOMENTO EN CADA P?NDOLA
'//

i = 1
n_pend = 1
p_hc_ap = 0
cont = 1
j = 11

While n_pend <= plantilla.Sheets(1).Cells(9, 4)

    If n_pend = 1 Then
    
        mom(i) = p_ap_tot_der * acum(i) - (p_sust / 2) * (acum(i) ^ 2)
   
    ElseIf i = plantilla.Sheets(1).Cells(9, 4) And plantilla.Sheets(1).Cells(5, 4) <> plantilla.Sheets(1).Cells(6, 4) Then
    mom(i) = (p_ap_tot_izq) * plantilla.Sheets(1).Cells(j + 2, 7) - (p_sust / 2) * (plantilla.Sheets(1).Cells(j + 2, 7) ^ 2)
   
    Else
     
        While cont <= i - 1
            n(cont) = 1
            cont = cont + 1
        Wend
        
        mom(i) = (p_ap_tot_der) * acum(i) - (p_sust / 2) * (acum(i) ^ 2) - (n(1) * fuerza(1) * (acum(i) - (dist(1))) + n(2) * fuerza(2) * (acum(i) - (dist(1) + dist(2))) + n(3) * fuerza(3) * (acum(i) - (dist(1) + dist(2) + dist(3))) + n(4) * fuerza(4) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4))) + n(5) * fuerza(5) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5))) + n(6) * fuerza(6) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6))) + n(7) * fuerza(7) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7))) + n(8) * fuerza(8) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8))) + n(9) * fuerza(9) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9))) + n(10) * fuerza(10) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10))) + _
        n(11) * fuerza(11) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11))) + n(12) * fuerza(12) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12))) + n(13) * fuerza(13) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13))) + n(14) * fuerza(14) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14))) + n(15) * fuerza(15) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15))) + _
        n(16) * fuerza(16) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16))) + _
        n(17) * fuerza(17) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17))) + n(18) * fuerza(18) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18))) + n(19) * fuerza(19) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19))) + _
        n(20) * fuerza(20) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20))) + n(21) * fuerza(21) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21))) + n(22) * fuerza(22) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22))) + _
        n(23) * fuerza(23) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22) + dist(23))) + _
        n(24) * fuerza(24) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22) + dist(23) + dist(24))) + n(25) * fuerza(25) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22) + dist(23) + dist(24) + dist(25))) + n(26) * fuerza(26) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22) + dist(23) + dist(24) + dist(25) + dist(26))) + _
        n(27) * fuerza(27) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22) + dist(23) + dist(24) + dist(25) + dist(26) + dist(27))) + n(27) * fuerza(27) * (acum(i) - (dist(1) + dist(2) + dist(3) + dist(4) + dist(5) + dist(6) + dist(7) + dist(8) + dist(9) + dist(10) + dist(11) + dist(12) + dist(13) + dist(14) + dist(15) + dist(16) + dist(17) + dist(18) + dist(19) + dist(20) + dist(21) + dist(22) + dist(23) + dist(24) + dist(25) + dist(26) + dist(27) + dist(28))))
        
    
    End If
    
    i = i + 1
    j = j + 2
    n_pend = n_pend + 1

Wend

'//
'//C?LCULO FLECHA SUSTENTADOR E HILO DE CONTACTO
'//

'//
'//Flecha hilo de contacto
'//

If el_hc_ini = 0 And el_hc_fin = 0 Then

i = 1
n_pend = 1
j = 11
While n_pend <= plantilla.Sheets(1).Cells(9, 4)

If n_pend = 1 Then
    
        fl_hc(i) = 0
    
    ElseIf i = plantilla.Sheets(1).Cells(9, 4) Then
        
        fl_hc(i) = fl_hc(1)
    
    Else
     
        fl_hc(i) = (p_var_1 / (2 * t_hc)) * (acum(i) - dist(1)) * ((va - 2 * dist(1)) - (acum(i) - dist(1)))
    
    End If
       
    i = i + 1
    j = j + 2
    n_pend = n_pend + 1

Wend

ElseIf el_hc_ini <> 0 And el_hc_fin <> 0 Then
n_pend = 1
i = 1

    While n_pend <= plantilla.Sheets(1).Cells(9, 4)
        fl_hc(i) = 0
        
        i = i + 1
        n_pend = n_pend + 1

    Wend

End If

'Descomposici?n fuerza
t_horiz_sust = Sqr((t_sust ^ 2) - (p_ap_tot_der ^ 2))

'//
'//Flecha sustentador
'//
i = 1
n_pend = 1
While n_pend <= plantilla.Sheets(1).Cells(9, 4)

    fl_sust(i) = mom(i) / t_horiz_sust

n_pend = n_pend + 1
i = i + 1

Wend

'//
'//C?LCULO LONGITUD P?NDOLAS
'//
i = 1
n_pend = 1
j = 12
While n_pend <= plantilla.Sheets(1).Cells(9, 4)

    If n_pend = 1 Then
    
        plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(5, 4) - fl_sust(i) + fl_hc(i) - el_hc_izq(i) - el_hc_der(i)
        plantilla.Worksheets(1).Cells(j, 5) = plantilla.Worksheets(1).Cells(j, 6) - l_sup_pend - l_inf_pend
    
    ElseIf i = plantilla.Sheets(1).Cells(9, 4) Then
        
        plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(6, 4) - fl_sust(i) + fl_hc(i) - el_hc_izq(i) - el_hc_der(i)
        plantilla.Worksheets(1).Cells(j, 5) = plantilla.Worksheets(1).Cells(j, 6) - l_sup_pend - l_inf_pend
    
    Else
     
        plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(5, 4) - fl_sust(i) + fl_hc(i) - el_hc_izq(i) - el_hc_der(i)
        plantilla.Worksheets(1).Cells(j, 5) = plantilla.Worksheets(1).Cells(j, 6) - l_sup_pend - l_inf_pend
    
    End If
    
n_pend = n_pend + 1
i = i + 1
j = j + 2

Wend

'//
'//Si la elevaci?n es a la izquierda
'//
                    
el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
el_hc_fin = plantilla.Sheets(1).Cells(8, 4)

If el_hc_ini <> 0 And el_hc_fin <> 0 Then

    If el_hc_ini > el_hc_fin Then
        el_hc_ini = el_hc_ini - el_hc_fin
        el_hc_fin = 0
    Else
        el_hc_fin = el_hc_fin - el_hc_ini
        el_hc_ini = 0
    End If
    
End If
    
    If el_hc_ini > el_hc_fin Then
        i = 1
        cont = 1
        j = 12
        n_pend = plantilla.Sheets(1).Cells(9, 4)
                                                    
        While i <= plantilla.Sheets(1).Cells(9, 4) + 1
            dist_der(i) = dist(i)
            l_pend_der(i) = plantilla.Sheets(1).Cells(j, 6)
            acum(i) = 0
            n_pend = n_pend - 1
            i = i + 1
            j = j + 2
        Wend
        i = 1
        j = 12
        acum(i) = 0
        n_pend = plantilla.Sheets(1).Cells(9, 4) + 1
        
        While i <= plantilla.Sheets(1).Cells(9, 4) + 1
            dist(i) = dist_der(n_pend)
            plantilla.Sheets(1).Cells(j - 1, 7) = dist(i)
            If l_pend_der(n_pend - 1) <> 0 Then
                plantilla.Sheets(1).Cells(j, 6) = l_pend_der(n_pend - 1)
                plantilla.Sheets(1).Cells(j, 5) = plantilla.Sheets(1).Cells(j, 6) - l_sup_pend - l_inf_pend
            End If
            
            acum(i) = acum(i - 1) + dist(i)
            plantilla.Sheets(1).Cells(j - 1, 8) = acum(i)
                               
            n_pend = n_pend - 1
            i = i + 1
            j = j + 2
        Wend
        
    End If
    
el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
el_hc_fin = plantilla.Sheets(1).Cells(8, 4)
    
If el_hc_ini <> 0 And el_hc_fin <> 0 Then

    If el_hc_ini > el_hc_fin Then
        var_comp = el_hc_fin
    Else
        var_comp = el_hc_ini
    End If
        i = 1
        j = 12
        
    While i <= plantilla.Sheets(1).Cells(9, 4)
        plantilla.Sheets(1).Cells(j, 6) = plantilla.Sheets(1).Cells(j, 6) - var_comp
        plantilla.Sheets(1).Cells(j, 5) = plantilla.Sheets(1).Cells(j, 6) - l_sup_pend - l_inf_pend
        
        i = i + 1
        j = j + 2
    Wend
    
End If
'//
'//GENERACI?N FICHA
'//
b = 1

    If st = 0 Then
    
        d = Workbooks(1).Worksheets(1).Cells(a, 1).Value
        e = Workbooks(1).Worksheets(1).Cells(a + 2, 1).Value
    
    Else
    
        d = Workbooks(1).Worksheets(1).Cells(a, 1).Value & "_1"
        e = Workbooks(1).Worksheets(1).Cells(a + 2, 1).Value & "_1"
        
    End If


    'fncScreenUpdating State:=False
    Call plantilla.Worksheets(b).PrintOut(from:=1, To:=1, Copies:=1, preview:=False, ActivePrinter:="Adobe PDF", printtofile:=True, collate:=False, prtofilename:=ruta_replanteoVB & "\" & d & " " & e & ".ps")
    'fncScreenUpdating State:=True
    PSFileName = ruta_replanteoVB & "\" & d & " " & e & ".ps"
    PDFFileName = ruta_replanteoVB & "\" & d & " " & e & ".pdf"
    TXTFileName = ruta_replanteoVB & "\" & d & " " & e & ".log"
    myPDF.FileToPDF PSFileName, PDFFileName, ""
    fso.DeleteFile PSFileName, True
    fso.DeleteFile TXTFileName, True
'//
'//INSERCI?N FICHA EN PDF GLOBAL
'//
    Call CombPDF(PDFFijoFileName, PDFFileName, ruta_replanteoVB)
'//
'//Inicializaci?n variables
'//
i = 0
While i <= 100
    dist(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    fuerza(i) = 0
    i = i + 1
Wend
i = 0
While i <= 30
    n(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    mom(i) = 0
    i = i + 1
Wend
i = 0
While i <= 30
    fl_hc(i) = 0
    i = i + 1
Wend
i = 0
While i <= 30
    fl_sust(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    el_hc_der(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    el_hc_der(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    fuerza_der(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    dist_der(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
    l_pend_der(i) = 0
    i = i + 1
Wend
i = 0
While i <= 20
acum(i) = 0
    i = i + 1
Wend

p_ap_tot_der_var = 0
p_hc_ap_izq = 0
p_hc_ap_der = 0
va_var_sla = 0
aisl = 0

If it = 1 Then
    
    plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 45)
    plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 47)
    plantilla.Sheets(1).Cells(7, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 46)
    plantilla.Sheets(1).Cells(8, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 48)
    dist_ap_prim_pend_izq = Workbooks(1).Sheets(1).Cells(a + 1, 49)
    dist_ap_prim_pend_der = Workbooks(1).Sheets(1).Cells(a + 1, 50)
    plantilla.Sheets(1).Cells(6, 10) = Workbooks(1).Sheets(1).Cells(a, 1)
    plantilla.Sheets(1).Cells(8, 10) = Workbooks(1).Sheets(1).Cells(a + 2, 1)
    plantilla.Sheets(1).Cells(7, 5) = "tip_pend_var"
    it = 0
    st = 1
          
    GoTo ini2
    
Else
    st = 0
    a = a + 2
    GoTo ini1

End If
x:
    fso.DeleteFile ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf", True
    '//
    '// cerrar objectos
    '//
    myPDF.CancelJob
    Set myPDF = Nothing
    'fso.Close
    Set fso = Nothing

    plantilla.DisplayAlerts = False
    plantilla.Workbooks.Close
    plantilla.Quit
    Set plantilla = Nothing
   
final:
   
End Sub
Sub pendolado_MT(nombre_catVB, ruta_replanteoVB)
'//
'//INSERTA DIRECTAMENTE FICHAS YA CREADAS A?ADIENDO ALGUNOS DATOS DEL REPLANTEO
'//

Dim fso As Object
Dim dist(100) As Double, fuerza(40) As Double, mom(40) As Double
Dim n_pend As Long, cont As Long, n(40) As Long
Dim fl_hc(40) As Double, fl_sust(40) As Double, el_hc_der(40) As Double, el_hc_izq(40) As Double
Dim fuerza_der(40) As Double, dist_der(40) As Double, l_pend_der(40) As Double, acum(40) As Double
Dim dist_ap_prim_pend_izq As Double, dist_ap_prim_pend_der As Double, dist_prim_seg_pend_izq As Double
Dim dist_prim_seg_pend_der As Double, va_var_sla As Double, l_sup_pend As Double, l_inf_pend As Double
Dim va As Double, el_hc_ini As Double, el_hc_fin As Double, p_aisl As Double, var_dist As Double
Dim p_sust_var As Double, va_min As Double, var_0 As Double, va_var As Double
Dim var_1 As Double, p_var_0 As Double, p_var_1 As Double, p_var_2 As Double, p_var_3 As Double
Dim p_sust_ap As Double, x_var As Double, p_hc_ap As Double, p_pend_equip As Double, p_hc_ap_izq As Double
Dim p_hc_ap_der As Double, p_ap_tot_izq As Double, p_ap_tot_der_var As Double, p_ap_tot_der As Double
Dim t_horiz_sust As Double, var_comp As Double, d As String, e As String, st As Long
Dim a As Integer, dfijo As String, efijo As String, aisl As Integer, it As Integer, it_var As Integer
Dim npi As Integer, i As Integer, j As Integer, var_izq_der As Integer, b As Integer
Dim tip As String, PDFFijoFileName As String, tip_var As String, PSFileName As String
Dim PDFFileName As String, TXTFileName As String
Dim pk_ini_var As String, pk_fin_var As String, tip_pend As String, tip_pend_var As String


Dim myPDF As PdfDistiller
Set myPDF = New PdfDistiller
Set fso = CreateObject("Scripting.FileSystemObject")

Dim plantilla As Object
Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False

a = 10
dfijo = Workbooks(1).Worksheets(1).Cells(a, 1).Value
efijo = Workbooks(1).Worksheets(1).Cells(a + 2, 1).Value
PDFFijoFileName = ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf"

ini4:

If Workbooks(1).Sheets(1).Cells(a + 2, 1) = "" Then

    GoTo final:

End If

tip_pend = Workbooks(1).Sheets(1).Cells(a + 1, 11)
tip_pend_var = Workbooks(1).Sheets(1).Cells(a + 1, 12)

If tip_pend <> "" And tip_pend_var <> "" Then

    it = 1
    
Else: it = 0
        
End If

ini3:

If (fso.FileExists("W:\210\P\P210D50\IN_INFORMES\8-Mission 3 - ?tudes d'ex?cution cat?naire\CATENARIA 3.000 Vcc\PENDOLADO\TODOS\" & tip_pend & ".xlsx")) Then

Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False
plantilla.Workbooks.Open "W:\210\P\P210D50\IN_INFORMES\8-Mission 3 - ?tudes d'ex?cution cat?naire\CATENARIA 3.000 Vcc\PENDOLADO\TODOS\" & tip_pend & ".xlsx"

Else

Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False
plantilla.Workbooks.Open "W:\210\P\P210D50\IN_INFORMES\8-Mission 3 - ?tudes d'ex?cution cat?naire\CATENARIA 3.000 Vcc\PENDOLADO\TODOS\non_defini.xlsx"

plantilla.Sheets(1).Cells(4, 4) = Workbooks(1).Sheets(1).Cells(a + 1, 4)

End If

pk_ini_var = Int((Workbooks(1).Sheets(1).Cells(a, 3) + 1) / 1000) & "+" & Int(Workbooks(1).Sheets(1).Cells(a, 3) + 1)
pk_fin_var = Int((Workbooks(1).Sheets(1).Cells(a + 2, 3) + 1) / 1000) & "+" & Int(Workbooks(1).Sheets(1).Cells(a + 2, 3) + 1)

plantilla.Sheets(1).Name = pk_ini_var & " - " & pk_fin_var

plantilla.Sheets(1).Cells(6, 10) = Workbooks(1).Sheets(1).Cells(a, 1)
plantilla.Sheets(1).Cells(8, 10) = Workbooks(1).Sheets(1).Cells(a + 2, 1)
plantilla.Sheets(1).Cells(7, 5) = tip_pend

'//
'//GENERACI?N FICHA
'//
b = 1

    If it = 1 Or it = 0 Then
    
        d = Workbooks(1).Worksheets(1).Cells(a, 1).Value
        e = Workbooks(1).Worksheets(1).Cells(a + 2, 1).Value
    
    Else: it = 2
    
        d = Workbooks(1).Worksheets(1).Cells(a, 1).Value & "_1"
        e = Workbooks(1).Worksheets(1).Cells(a + 2, 1).Value & "_1"
        
    End If


    'fncScreenUpdating State:=False
    Call plantilla.Worksheets(b).PrintOut(from:=1, To:=1, Copies:=1, preview:=False, ActivePrinter:="Adobe PDF", printtofile:=True, collate:=False, prtofilename:=ruta_replanteoVB & "\" & d & " " & e & ".ps")
    'fncScreenUpdating State:=True
    PSFileName = ruta_replanteoVB & "\" & d & " " & e & ".ps"
    PDFFileName = ruta_replanteoVB & "\" & d & " " & e & ".pdf"
    TXTFileName = ruta_replanteoVB & "\" & d & " " & e & ".log"
    myPDF.FileToPDF PSFileName, PDFFileName, ""
    fso.DeleteFile PSFileName, True
    fso.DeleteFile TXTFileName, True
'//
'//INSERCI?N FICHA EN PDF GLOBAL
'//
    Call CombPDF(PDFFijoFileName, PDFFileName, ruta_replanteoVB)


If it = 1 Then
    
    it = 2
    tip_pend = tip_pend_var
    plantilla.DisplayAlerts = False
    plantilla.Workbooks.Close
    plantilla.Quit
    Set plantilla = Nothing
    
    GoTo ini3
          
    'GoTo ini2
    
Else

    a = a + 2
    plantilla.DisplayAlerts = False
    plantilla.Workbooks.Close
    plantilla.Quit
    Set plantilla = Nothing
    
    GoTo ini4

End If

final:
     
    fso.DeleteFile ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf", True
    '//
    '// cerrar objectos
    '//
    myPDF.CancelJob
    Set myPDF = Nothing
    'fso.Close
    Set fso = Nothing
    
End Sub


Function CombPDF(PDFFijo, PDFName, ruta_replanteoVB)
Dim fso As Object
Dim AcroApp As Acrobat.CAcroApp
Dim Part1Document As Acrobat.CAcroPDDoc
Dim Part2Document As Acrobat.CAcroPDDoc
Dim numPages As Integer
Set AcroApp = CreateObject("AcroExch.App")
Set Part1Document = CreateObject("AcroExch.PDDoc")
Set Part2Document = CreateObject("AcroExch.PDDoc")
Set fso = CreateObject("Scripting.FileSystemObject")
Part1Document.Open (PDFFijo)
Part2Document.Open (PDFName)
    
numPages = Part1Document.GetNumPages()
    
If Part1Document.InsertPages(numPages - 1, Part2Document, _
0, Part2Document.GetNumPages(), True) = False Then
    Exit Function
End If
If Part1Document.Save(PDSaveFull, ruta_replanteoVB & "\pendulage.pdf") = False Then
Else
        PDFFijo = ruta_replanteoVB & "\pendulage.pdf"
        
End If
        
Part1Document.Close
Part2Document.Close
AcroApp.Exit
Set AcroApp = Nothing
Set Part1Document = Nothing
Set Part2Document = Nothing
fso.DeleteFile PDFName, True
End Function

Sub pendolado_columna()
Dim z As Integer, cont As Integer
Dim longueur As Double
Dim primero As String, primero2 As String, segundo As String, tercero As String, cuatro As String, segundo2 As String, tercero2 As String, cuarto As String, cuarto2 As String
Dim normal As Boolean
z = 10

While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
    primero = ""
    primero2 = ""
    segundo = ""
    segundo2 = ""
    tercero = ""
    tercero2 = ""
    cuarto = ""
    cuarto2 = ""
    '///
    '///normal o inverso
    '///
    If (Sheets(1).Cells(z, 16).Value = "Anc.Chevau." Or Sheets(1).Cells(z, 16).Value = "Anc.Section.") Then
        If Sheets(1).Cells(z, 8).Value > 0 Then
            normal = False
        Else
            normal = True
        End If
    End If
    '///
    '///primera letra
    '///
    If IsEmpty(Sheets(1).Cells(z, 16).Value) Or ((Sheets(1).Cells(z, 16).Value = "Anc.Chevau." Or Sheets(1).Cells(z, 16).Value = "Anc.Chevau.sans AT") And Sheets(1).Cells(z - 2, 16).Value = "Inter.Chevau.") Or _
    (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z - 2, 16).Value = "Inter.Section.") Or Sheets(1).Cells(z, 16).Value = "Anc.Antich." Or Sheets(1).Cells(z, 16).Value = "Axe.Antich." Then
        primero = "C"
    Else
        primero = "S"
    End If
    '///
    '///segunda letra
    '///
    If IsEmpty(Sheets(1).Cells(z, 16).Value) Or ((Sheets(1).Cells(z, 16).Value = "Anc.Chevau." Or Sheets(1).Cells(z, 16).Value = "Anc.Chevau.sans AT") And Sheets(1).Cells(z - 2, 16).Value = "Inter.Chevau.") Or _
    (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z - 2, 16).Value = "Inter.Section.") Or Sheets(1).Cells(z, 16).Value = "Anc.Antich." Or Sheets(1).Cells(z, 16).Value = "Axe.Antich." Then
        segundo = "S"
    ElseIf (Sheets(1).Cells(z, 16).Value = "Anc.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.") Or _
    (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Section.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Section." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Section.") _
    Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.sans AT") Then
        segundo = "K"
        If Sheets(1).Cells(z + 1, 4).Value >= 40.5 Then
            segundo2 = 1
        ElseIf (Sheets(1).Cells(z, 16).Value = "Anc.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.") Or _
        (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.sans AT") Then
            segundo2 = 3
        Else
            segundo2 = 2
        End If
    ElseIf (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Chevau.") Then
        segundo = "K"
        segundo2 = "C"
        primero2 = primero
    ElseIf Sheets(1).Cells(z, 16).Value = "Axe.Chevau." Then
        segundo = "C"
        segundo2 = "K"
        primero2 = primero
    ElseIf (Sheets(1).Cells(z, 16).Value = "Inter.Section." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Section.") Then
        segundo = "K"
        segundo2 = "S"
        primero2 = primero
    ElseIf Sheets(1).Cells(z, 16).Value = "Axe.Section." Then
        segundo = "S"
        segundo2 = "K"
        primero2 = primero
    End If
    '///
    '///tercera letra
    '///
    
    If IsEmpty(Sheets(1).Cells(z, 16).Value) Or ((Sheets(1).Cells(z, 16).Value = "Anc.Chevau." Or Sheets(1).Cells(z, 16).Value = "Anc.Chevau.sans AT") And Sheets(1).Cells(z - 2, 16).Value = "Inter.Chevau.") Or _
    (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z - 2, 16).Value = "Inter.Section.") Or Sheets(1).Cells(z, 16).Value = "Anc.Antich." Or Sheets(1).Cells(z, 16).Value = "Axe.Antich." Then
        tercero = "n"
    ElseIf (Sheets(1).Cells(z, 16).Value = "Anc.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." Or Sheets(1).Cells(z, 16).Value = "Axe.Chevau.") _
    Or (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Section.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Section." Or Sheets(1).Cells(z, 16).Value = "Axe.Section.") _
    Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.sans AT") Then
            cont = 1
            longueur = 63 '// debe venir de una variable
            While cont <= 8 And Sheets(1).Cells(z + 1, 4).Value <> longueur
                longueur = longueur - 4.5 'inc_norm_va
                cont = cont + 1
            Wend
            If cont = 9 Then
                tercero = Round(Sheets(1).Cells(z + 1, 4).Value, 2)
                tercero2 = tercero
            Else
                tercero = cont
                tercero2 = tercero
            End If
    End If
    If (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Section.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Section." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Section.") _
    Or (Sheets(1).Cells(z, 16).Value = "Anc.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.") Then
        tercero2 = "A"
        primero2 = primero
    End If
    '///
    '///cuarta letra
    '///
    If IsEmpty(Sheets(1).Cells(z, 16).Value) Or ((Sheets(1).Cells(z, 16).Value = "Anc.Chevau." Or Sheets(1).Cells(z, 16).Value = "Anc.Chevau.sans AT") And Sheets(1).Cells(z - 2, 16).Value = "Inter.Chevau.") Or _
    (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z - 2, 16).Value = "Inter.Section.") Or Sheets(1).Cells(z, 16).Value = "Anc.Antich." Or Sheets(1).Cells(z, 16).Value = "Axe.Antich." Then
        cuarto = Round(Sheets(1).Cells(z + 1, 4).Value, 2)
    'ElseIf normal = True Then
    ElseIf (Sheets(1).Cells(z, 16).Value = "Anc.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Anc.Section." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Section.") Then
            cuarto = "a"
            cuarto2 = Round(Sheets(1).Cells(z + 1, 4).Value, 2)
            Sheets(1).Cells(z + 1, 39).Value = 1.4
            Sheets(1).Cells(z + 1, 40).Value = 0
            Sheets(1).Cells(z + 1, 41).Value = 1.4
            Sheets(1).Cells(z + 1, 42).Value = 0
            Sheets(1).Cells(z + 1, 43).Value = 1.125
            Sheets(1).Cells(z + 1, 44).Value = 2.5
    ElseIf (Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau." And Sheets(1).Cells(z, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z + 2, 16).Value = "Anc.Section." And Sheets(1).Cells(z, 16).Value = "Inter.Section.") _
    Or (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Anc.Chevau.sans AT") Then
            cuarto = "b"
            cuarto2 = Round(Sheets(1).Cells(z + 1, 4).Value, 2)
            Sheets(1).Cells(z + 1, 39).Value = 1.4
            Sheets(1).Cells(z + 1, 40).Value = 0
            Sheets(1).Cells(z + 1, 41).Value = 1.4
            Sheets(1).Cells(z + 1, 42).Value = 0
            Sheets(1).Cells(z + 1, 43).Value = 1.125
            Sheets(1).Cells(z + 1, 44).Value = 2.5
    ElseIf (Sheets(1).Cells(z, 16).Value = "Inter.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Inter.Section." And Sheets(1).Cells(z + 2, 16).Value = "Axe.Section.") Then
        If normal = True Then
            cuarto = "e"
            Sheets(1).Cells(z + 1, 39).Value = 1.4
            Sheets(1).Cells(z + 1, 40).Value = 0
            Sheets(1).Cells(z + 1, 41).Value = 1.3
            Sheets(1).Cells(z + 1, 42).Value = 0
            Sheets(1).Cells(z + 1, 43).Value = 2.5
            Sheets(1).Cells(z + 1, 44).Value = 2.5
            Sheets(1).Cells(z + 1, 45).Value = 1.8
            Sheets(1).Cells(z + 1, 47).Value = 2
            Sheets(1).Cells(z + 1, 48).Value = 0
            Sheets(1).Cells(z + 1, 49).Value = 2.5
            Sheets(1).Cells(z + 1, 50).Value = 2.5
            If Sheets(1).Cells(z + 1, 4).Value >= 40.5 Then
                cuarto2 = "g"
                Sheets(1).Cells(z + 1, 44).Value = 0.5
            ElseIf segundo2 = "C" Then
                cuarto2 = "g1"
                Sheets(1).Cells(z + 1, 44).Value = 0.3
            ElseIf segundo2 = "S" Then
                cuarto2 = "g1"
                Sheets(1).Cells(z + 1, 44).Value = 0.35
            End If
        Else
            cuarto = "k"
            Sheets(1).Cells(z + 1, 39).Value = 1.4
            Sheets(1).Cells(z + 1, 40).Value = 0
            Sheets(1).Cells(z + 1, 41).Value = 2
            Sheets(1).Cells(z + 1, 42).Value = 0
            Sheets(1).Cells(z + 1, 43).Value = 2.5
            Sheets(1).Cells(z + 1, 44).Value = 2.5
            Sheets(1).Cells(z + 1, 45).Value = 1.8
            Sheets(1).Cells(z + 1, 47).Value = 1.3
            Sheets(1).Cells(z + 1, 48).Value = 0
            Sheets(1).Cells(z + 1, 49).Value = 2.5
            Sheets(1).Cells(z + 1, 50).Value = 2.5
            If Sheets(1).Cells(z + 1, 4).Value >= 40.5 Then
                cuarto2 = "i"
                Sheets(1).Cells(z + 1, 44).Value = 0.5
            ElseIf segundo2 = "C" Then
                cuarto2 = "i1"
                Sheets(1).Cells(z + 1, 44).Value = 0.3
            ElseIf segundo2 = "S" Then
                cuarto2 = "i1"
                Sheets(1).Cells(z + 1, 44).Value = 0.35
            End If
        End If
    ElseIf (Sheets(1).Cells(z, 16).Value = "Axe.Chevau." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Chevau.") Or (Sheets(1).Cells(z, 16).Value = "Axe.Section." And Sheets(1).Cells(z + 2, 16).Value = "Inter.Section.") Then
        If normal = True Then
            Sheets(1).Cells(z + 1, 39).Value = 1.3
            Sheets(1).Cells(z + 1, 40).Value = 0
            Sheets(1).Cells(z + 1, 41).Value = 1.8
            cuarto2 = "h"
            Sheets(1).Cells(z + 1, 45).Value = 2
            Sheets(1).Cells(z + 1, 46).Value = 0
            Sheets(1).Cells(z + 1, 47).Value = 1.4
            Sheets(1).Cells(z + 1, 48).Value = 0
            Sheets(1).Cells(z + 1, 43).Value = 2.5
            Sheets(1).Cells(z + 1, 44).Value = 2.5
            Sheets(1).Cells(z + 1, 49).Value = 2.5
            Sheets(1).Cells(z + 1, 50).Value = 2.5
            If Sheets(1).Cells(z + 1, 4).Value >= 40.5 Then
                cuarto = "f"
                Sheets(1).Cells(z + 1, 42).Value = 0.5
            ElseIf segundo = "C" Then
                cuarto = "f1"
                Sheets(1).Cells(z + 1, 42).Value = 0.3
            ElseIf segundo = "S" Then
                cuarto = "f1"
                Sheets(1).Cells(z + 1, 42).Value = 0.35
            End If
        Else
            Sheets(1).Cells(z + 1, 39).Value = 2
            Sheets(1).Cells(z + 1, 40).Value = 0
            Sheets(1).Cells(z + 1, 41).Value = 1.8
            cuarto2 = "j"
            Sheets(1).Cells(z + 1, 45).Value = 1.3
            Sheets(1).Cells(z + 1, 46).Value = 0
            Sheets(1).Cells(z + 1, 47).Value = 1.4
            Sheets(1).Cells(z + 1, 48).Value = 0
            Sheets(1).Cells(z + 1, 43).Value = 2.5
            Sheets(1).Cells(z + 1, 44).Value = 2.5
            Sheets(1).Cells(z + 1, 49).Value = 2.5
            Sheets(1).Cells(z + 1, 50).Value = 2.5
            If Sheets(1).Cells(z + 1, 4).Value >= 40.5 Then
                cuarto = "l"
                Sheets(1).Cells(z + 1, 42).Value = 0.5
            ElseIf segundo = "C" Then
                cuarto = "l1"
                Sheets(1).Cells(z + 1, 42).Value = 0.3
            ElseIf segundo = "S" Then
                cuarto = "l1"
                Sheets(1).Cells(z + 1, 42).Value = 0.35
            End If
        End If
    End If

Sheets(1).Cells(z + 1, 11).Value = primero & segundo & tercero & cuarto
Sheets(1).Cells(z + 1, 12).Value = primero2 & segundo2 & tercero2 & cuarto2
z = z + 2
Wend
End Sub
