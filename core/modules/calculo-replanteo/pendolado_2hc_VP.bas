Attribute VB_Name = "pendolado_2hc_VP"
Public anc_2hc_1hc As Integer, contador_pend_long_VP As Double, contador_pend_long_tot_VP As Double, contador_pend_VP As Integer, a As Integer, va As Double
Public contador_pend_VP_anc As Integer
Public mypdf As Object, plantilla As Object, plantilla_control As Object, fso As Object, numpages As Integer
Public PDFFijoFileName As String, contador_hojas As Integer, TXTFileName As String, PSFileName As String, PDFFileName As String
Public st As Long
Public cambio As Integer

Option Explicit
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    '(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'// The SendMessage function sends the specified message to a window or windows.
'// The function calls the window procedure for the specified window and does not
'// return until the window procedure has processed the message.
'// The PostMessage function, in contrast, posts a message to a thread’s message
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
'Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'// PARAMETERS:
'// hWnd
'// Specifies the window handle.

'//////////////////////////////////////////////////////////////////////////
'//
'Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
    'lpRect As Long, ByVal bErase As Long) As Long

'Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

'Private Declare Function GetDesktopWindow Lib "user32" () As Long

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

Sub pendolado_2hc_VP(nombre_catVB, ruta_replanteoVB, fila_ini, fila_fin, cadena_general)

Dim Dist(100) As Double, fuerza(40) As Double, mom(40) As Double, acum_aisl(10) As Integer, fuerza_var(40) As Double, aisl_n_var(10) As Double, acum_aisl_var(10) As Double
Dim n_pend As Long, cont As Long, n(40) As Long, aisl_n(10) As Integer, desnivel_cambio(40) As Double
Dim fl_hc(40) As Double, fl_sust(40) As Double, el_hc(40) As Double
Dim fuerza_der(40) As Double, dist_der(40) As Double, l_pend_der(40) As Double, acum(40) As Double
Dim dist_ap_prim_pend_izq As Double, dist_ap_prim_pend_der As Double, dist_prim_seg_pend_izq As Double
Dim dist_prim_seg_pend_der As Double, va_var_sla As Double, l_sup_pend As Double, l_inf_pend As Double
Dim el_hc_ini As Double, el_hc_fin As Double, p_aisl As Double, var_dist As Double
Dim p_sust_var As Double, va_min As Double, var_0 As Double, va_var As Double
Dim var_1 As Double, p_var_0 As Double, p_var_1 As Double, p_var_2 As Double, p_var_3 As Double
Dim p_sust_ap As Double, x_var As Double, p_hc_ap As Double, p_pend_equip As Double, p_hc_ap_izq As Double
Dim p_hc_ap_der As Double, p_ap_tot_izq As Double, p_ap_tot_der_var As Double, p_ap_tot_der As Double
Dim t_horiz_sust As Double, var_comp As Double, d As String, e As String
Dim dfijo As String, efijo As String, aisl As Integer, it As Integer, it_var As Integer
Dim npi As Integer, i As Integer, j As Integer, var_izq_der As Integer, b As Integer
Dim tip As String, tip_var As String
Dim pk_ini_var As String, pk_fin_var As String, tip_pend As String, tip_pend_var As String
Dim ang(30) As Double, desnivel As Double, desnivel_0 As Double, desnivel_1 As Double, var_5 As Integer, var_6 As Integer, desn_contador As Integer, desnivel_rasante As Double, desnivel_alt_cat As Double, tangente_desnivel As Double
Dim pk_ini As Double, pk_fin As Double, d_var As Double, h_var As Double
Dim long_pend(50) As Double, alt_pend_min As Double, alt_pend_calc As Double
Dim p_ap_tot_der_aux As Double, p_ap_tot_izq_var As Double, p_ap_tot_izq_aux As Double
Dim va_calculado As Double, distancia_var As Double, var_col As Integer, ceros As String
Dim documento As String, nombre_cat As String
Dim el_hc_var As Double, alt_cat_ini As Double, alt_cat_fin As Double, alt_cat_var As Double
Dim ruta_replanteo As String
Dim fecha As Date
Dim aisl_sla As Integer, dist_aisl_1 As Double, dist_aisl_2 As Double
Dim tip_0 As String, tip_1 As String, tip_2 As String

nombre_catVB = "Marruecos 3.000 Vcc"
ruta_replanteoVB = "C:\Users\23370\Desktop\D50"

If anc_2hc_1hc = 2 Then
    
    anc_2hc_1hc = 0
    GoTo fin_anclaje_1hc
    
End If

'//
'//LECTURA BASE DE DATOS
'//
Call cargar.datos_lac(nombre_catVB)

'Dim mypdf As PdfDistiller
Set mypdf = New PdfDistiller
Set fso = CreateObject("Scripting.FileSystemObject")

'Dim plantilla As Object
Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = True

'Dim plantilla_control As Object
Set plantilla_control = CreateObject("Excel.Application")
plantilla_control.Visible = True

'//
'//HOJA DE CONTROL
'//
plantilla_control.Workbooks.Open "W:\223\D\D223041\IN_INFORMES\plantilla_control.xlsm"
plantilla_control.Sheets(1).Range("A9:B10001").ClearContents
'plantilla_control.Sheets(1).Range("B9:B10001").ClearContents
plantilla_control.Sheets(1).Range("C6:N6").ClearContents
plantilla_control.Sheets(1).Range("C9:C10001").ClearContents
'//
'//AÑADIR FECHA CREACIÓN
'//
fecha = Date

plantilla_control.Sheets("control").Cells(6, 3) = "07/05/2013"
plantilla_control.Sheets("control").Cells(6, 4) = "28/02/2014"
'//
'//INICIALIZACIÓN CONTADOR
'//
contador_pend_long_tot_VP = 0
contador_pend_long_VP = 0
contador_pend_VP = 0
contador_pend_VP_anc = 0
Workbooks(1).Sheets("Material").Cells(6, 12) = 0
Workbooks(1).Sheets("Material").Cells(6, 11) = 0
Workbooks(1).Sheets("Material").Cells(4, 11) = 0
Workbooks(1).Sheets("Material").Cells(2, 11) = 0
dist_aisl_1 = 0
'//
'//
'//


'//
'//PAGINADO FICHAS
'//
contador_hojas = 1
'//
'//
'//
a = fila_ini
While a < fila_fin

plantilla_control.Sheets(1).PageSetup.CenterHeader = vbCrLf & "&""Trebuchet MS,Bold""&12 " & "SOMMAIRE"
With plantilla_control.Sheets("control").PageSetup
    .RightFooter = "&""Arial,Bold""&12 "
End With
plantilla_control.Sheets("control").Cells(contador_hojas + 8, 1) = "FOLIO" & " " & contador_hojas & " - " & Workbooks(1).Sheets("Replanteo").Cells(a, 1) & " / " & Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1) & " - " & Workbooks(1).Sheets("Replanteo").Cells(a + 1, 11)
plantilla_control.Sheets("control").Cells(contador_hojas + 8, 3) = "+"
contador_hojas = contador_hojas + 1
If Not IsEmpty(Workbooks(1).Sheets("Replanteo").Cells(a + 1, 12).Value) Then
    plantilla_control.Sheets("control").Cells(contador_hojas + 8, 1) = "FOLIO" & " " & contador_hojas & " - " & Workbooks(1).Sheets("Replanteo").Cells(a, 1) & " / " & Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1) & " - " & Workbooks(1).Sheets("Replanteo").Cells(a + 1, 12)
    plantilla_control.Sheets("control").Cells(contador_hojas + 8, 3) = "+"
    contador_hojas = contador_hojas + 1
End If
a = a + 2

Wend
Call codificacion.codificacion("pendulage", a, cadena_general)
plantilla_control.Sheets(1).PageSetup.RightFooter = codigo


'//
'//SELECCIÓN TIPOLOGÍA PENDOLADO
'//

'//
'//Sólo un pendolado para un vano
'//

'//
'//FILA INICIO
'//
a = fila_ini

documento = "pendulage"
Call codificacion.codificacion(documento, a, cadena_general)

'//
'//FILA INICIO
'//

dfijo = Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value
efijo = Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 1).Value

PDFFijoFileName = ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf"
aisl = 0
ini1:

If a = fila_fin Then
    
    GoTo final
    
End If

'//
'//SALAR SI HAY MARQUISE METALIQUE
'//

While Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 11).Value = ""

    a = a + 2

Wend

'//
'//FILA FIN
'//


'//
'//FILA INICIO
'//

plantilla.Workbooks.Open "W:\223\D\D223041\IN_INFORMES\plantilla_pendolado.xlsm"


Call cargar.datos_lac(nombre_catVB)

If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 52) = "" Then

    va = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 4)
    
Else

    va = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 52)
    
End If

plantilla.Sheets(1).Range("D4:D9").ClearContents
plantilla.Sheets(1).Cells(6, 11).ClearContents
plantilla.Sheets(1).Cells(8, 11).ClearContents
plantilla.Sheets(1).Range("D12:D10001").ClearContents
plantilla.Sheets(1).Range("E12:E10001").ClearContents
plantilla.Sheets(1).Range("F12:F10001").ClearContents
plantilla.Sheets(1).Range("G11:G10001").ClearContents
plantilla.Sheets(1).Range("H11:H10001").ClearContents
plantilla.Sheets(1).Range("I11:I10001").ClearContents
plantilla.Sheets(1).Range("J11:J10001").ClearContents


'//
'//Formato para que se muestre en el encabezado de cada ficha
'//

'pk_ini_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) & "+" & (Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000)

        'If Round(Workbooks(1).Sheets("Replanteo").Cells(a, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000, 2) < 100 Then
            'ceros = "0"
            'If Round(Workbooks(1).Sheets("Replanteo").Cells(a, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000, 2) < 10 Then
            'ceros = "00"
            'End If
        'Else
            'ceros = ""
        'End If
        'pk_ini_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) & "+" & ceros & (Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000)

'pk_fin_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) & "+" & (Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000)

        'If Round(Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000, 2) < 100 Then
            'ceros = "0"
            'If Round(Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000, 2) < 10 Then
            'ceros = "00"
            'End If
        'Else
            'ceros = ""
        'End If
        'pk_fin_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) & "+" & ceros & (Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000)
        
'//
'//Nueva función
'//

pk_ini_var = Workbooks(1).Sheets("Replanteo").Cells(a, 3).text
pk_fin_var = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3).text

        
plantilla.Sheets(1).Name = pk_ini_var & " - " & pk_fin_var

plantilla.Sheets(1).Cells(3, 7).Value = pk_ini_var & " - " & pk_fin_var

plantilla.Sheets(1).Cells(4, 11) = codigo

plantilla.Sheets(1).Cells(2, 5) = "LIGNE: " & nombre_tramo


'//
'//Lectura tipo de pendolado
'//

it = 0
st = 0

If n_hc = 2 Then

    dist_max_pend = dist_max_pend / 2

End If

tip_pend = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 11)
tip_pend_var = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 12)




If tip_pend_var <> "" Then

    it = 1
    
End If

'En el caso de que encuentre aguja que no haga el pendolado de la parte correspondiente a la aguja

tip_0 = Workbooks(1).Sheets("Replanteo").Cells(a - 2, 16)
tip_1 = Workbooks(1).Sheets("Replanteo").Cells(a, 16)
tip_2 = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 16)

'//
'//MIRA SI EL CÓDIGO ES MÁS LARGO DE LO QUE TOCA Y RECOGE LA PARTE QUE INTERESA
'//
If Len(Sheets("Replanteo").Cells(a, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(a, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(a, 16).Value = anc_sm_sin) Then
    tip_1 = Mid(Sheets("Replanteo").Cells(a, 16).Value, 15)
    'tip_pf_1 = Mid(Sheets("Replanteo").Cells(a, 16).Value, 1, 11)
Else
    tip_1 = Sheets("Replanteo").Cells(a, 16).Value
    'tip_pf_1 = Sheets("Replanteo").Cells(a, 16).Value
End If
If Len(Sheets("Replanteo").Cells(a - 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(a - 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(a - 2, 16).Value = anc_sm_sin) Then
    tip_0 = Mid(Sheets("Replanteo").Cells(a - 2, 16).Value, 15)
    'tip_pf_0 = Mid(Sheets("Replanteo").Cells(a - 2, 16).Value, 1, 11)
Else
    tip_0 = Sheets("Replanteo").Cells(a - 2, 16).Value
    'tip_pf_0 = Sheets("Replanteo").Cells(a - 2, 16).Value
End If
If Len(Sheets("Replanteo").Cells(a + 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(a + 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(a + 2, 16).Value = anc_sm_sin) Then
    tip_2 = Mid(Sheets("Replanteo").Cells(a + 2, 16).Value, 15)
    'tip_pf_2 = Mid(Sheets("Replanteo").Cells(a + 2, 16).Value, 1, 11)
Else
    tip_2 = Sheets("Replanteo").Cells(a + 2, 16).Value
    'tip_pf_2 = Sheets("Replanteo").Cells(a + 2, 16).Value
End If


If tip_1 = "Anc.Aigu." Or tip_1 = "Axe.Aigu." Or tip_1 = "Inter.Aigu." Then

    it = 0

End If


If it = 1 Then

    'If Workbooks(1).Sheets("Replanteo").Cells(a, 39) = "" And Workbooks(1).Sheets("Replanteo").Cells(a, 41) = "" Then
    
        'plantilla.Sheets(1).Cells(5, 4) = alt_cat
        'plantilla.Sheets(1).Cells(6, 4) = alt_cat
        'plantilla.Sheets(1).Cells(8, 4) = 0
        'plantilla.Sheets(1).Cells(9, 4) = 0
        'dist_ap_prim_pend_izq = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 43)
        'dist_ap_prim_pend_der = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 44)
        'plantilla.Sheets(1).Cells(6, 11) = Workbooks(1).Sheets("Replanteo").Cells(a, 1)
        'plantilla.Sheets(1).Cells(8, 11) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1)
       
    'Else
        plantilla.Sheets(1).Cells(4, 4) = va
        plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 39)
        plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 41)
        plantilla.Sheets(1).Cells(7, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 40)
        plantilla.Sheets(1).Cells(8, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 42)
        dist_ap_prim_pend_izq = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 43)
        dist_ap_prim_pend_der = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 44)
        plantilla.Sheets(1).Cells(6, 11) = Workbooks(1).Sheets("Replanteo").Cells(a, 1)
        plantilla.Sheets(1).Cells(8, 11) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1)
        plantilla.Sheets(1).Cells(7, 5) = tip_pend
    'End If
    
ElseIf Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 43) = "var" Then
        
    plantilla.Sheets(1).Cells(4, 4) = va
    plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 39)
    plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 41)
    plantilla.Sheets(1).Cells(7, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 40)
    plantilla.Sheets(1).Cells(8, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 42)
    plantilla.Sheets(1).Cells(6, 11) = Workbooks(1).Sheets("Replanteo").Cells(a, 1)
    plantilla.Sheets(1).Cells(8, 11) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1)
    plantilla.Sheets(1).Cells(7, 5) = tip_pend
    
    'valores a leer del Sireca
    p_sust = 1.387 'kg/m
    p_hc = 0.932 'kg/m
    p_pend = 0.101 'kg/m
    p_pend_equip = 0.08 + 0.15 'kg
    
    n_hc = 2
    t_sust = 1400 'kg
    t_hc = 1000 'kg
    fl_max_centro_va = va / 1000
    
    el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
    el_hc_fin = plantilla.Sheets(1).Cells(8, 4)
    alt_cat_ini = plantilla.Sheets(1).Cells(5, 4)
    alt_cat_fin = plantilla.Sheets(1).Cells(6, 4)

    l_sup_pend = 0.036
    l_inf_pend = 0.0336
    
    GoTo anclaje_2hc_top
    
Else
    plantilla.Sheets(1).Cells(4, 4) = va
    plantilla.Sheets(1).Cells(5, 4) = alt_cat
    plantilla.Sheets(1).Cells(6, 4) = alt_cat
    plantilla.Sheets(1).Cells(7, 4) = 0
    plantilla.Sheets(1).Cells(8, 4) = 0
    plantilla.Sheets(1).Cells(6, 11) = Workbooks(1).Sheets("Replanteo").Cells(a, 1)
    plantilla.Sheets(1).Cells(8, 11) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1)
    dist_ap_prim_pend_izq = dist_ap_prim_pend
    dist_ap_prim_pend_der = dist_ap_prim_pend
    plantilla.Sheets(1).Cells(7, 5) = tip_pend

    '//
    '//SI HAY TUNEL DEBE LEER OTRAS DISTANCIAS ENTRE HILO Y SUSTENTADOR
    '//
    If Workbooks(1).Worksheets("Replanteo").Cells(a, 38).Value = "" And Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 38).Value = "Tunel" Then
        
        'plantilla.Sheets(1).Cells(5, 4) = alt_cat
        plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 39)
    
    ElseIf Workbooks(1).Worksheets("Replanteo").Cells(a, 38).Value = "Tunel" And Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 38).Value = "Tunel" Then
                
        plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets("Replanteo").Cells(a, 39)
        plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 39)
    
    ElseIf Workbooks(1).Worksheets("Replanteo").Cells(a, 38).Value = "Tunel" And Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 38).Value = "" Then
        
        plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets("Replanteo").Cells(a, 39)
        'plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 4, 39)
        
    End If

End If


'//
'//CÁLCULO GEOMÉTRICO
'//

    
ini2:

If Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value = "" Then

    GoTo final

End If


'//
'//Longitud de cabeza superior e inferior de la péndola
'//

'//l_sup_pend depende si el sustentador es de 153 36mm, si es de 93 34.55mm

el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
el_hc_fin = plantilla.Sheets(1).Cells(8, 4)
alt_cat_ini = plantilla.Sheets(1).Cells(5, 4)
alt_cat_fin = plantilla.Sheets(1).Cells(6, 4)

l_sup_pend = 0.036
l_inf_pend = 0.0336

'//
'//Elección aislador en caso de que haya
'//
If aisl = 1 Then
    
    If cola_anc = "Cerámico" Then
            'p_aisl = 15
            'p_sust_var = p_sust * va
            'p_sust = (p_sust_var + p_aisl) / va
    
    ElseIf cola_anc = "Sintético" Then
            'p_aisl = 3
            'p_sust_var = p_sust * va
            'p_sust = (p_sust_var + p_aisl) / va
            
    ElseIf cola_anc = "Vidrio" Then
            'p_aisl = 4.5
            'p_sust_var = p_sust * va
            'p_sust = (p_sust_var + p_aisl) / va
    
    End If
    
End If

'//
'//Se debe distinguir entre posible casos con elevación a un lado, elevación en ambos, etc.
'//

'If el_hc_ini <> 0 And el_hc_fin <> 0 Then

    'If el_hc_ini > el_hc_fin Then
    '    el_hc_ini = el_hc_ini - el_hc_fin
    '    el_hc_fin = 0
    'Else
    '    el_hc_fin = el_hc_fin - el_hc_ini
    '    el_hc_ini = 0
    'End If
    
'End If



'PRUEBAAAAAAAAAAAAAS!!!!!!!!

'valores a leer del Sireca
p_sust = 1.387 'kg/m
p_hc = 0.932 'kg/m
p_pend = 0.101 'kg/m
p_pend_equip = 0.08 + 0.15 'kg

n_hc = 2
t_sust = 1400 'kg
t_hc = 1000 'kg
fl_max_centro_va = va / 1000

If n_hc = 1 Then

    plantilla.Sheets(1).Cells(5, 7) = "Caténaire Légere"

ElseIf n_hc = 2 Then

    plantilla.Sheets(1).Cells(5, 7) = "Caténaire Simple"

End If

'//
'//Flecha máxima en centro de vano impuesta
'//

p_pend_equip = 0.08 + 0.15 'kg
fl_max_centro_va = va / 1000

    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 55 - it) <> "" Then
    
        dist_max_pend = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 55 - it)
        dist_prim_seg_pend_izq = 4.1
        dist_prim_seg_pend_der = 4.1
        
        'En seccionamiento eléctrico (en seccionamiento eléctrico el aisl irá a 0,75m de la primera péndola, al girarlo todo pq la elevación va a derecha para el cálculo, deberemos ponerlo a 0,75m de la última. (en doble hilo afcará a última y penúltima))
        If Workbooks(1).Sheets("Replanteo").Cells(a, 16) = "Inter.Section." Or Workbooks(1).Sheets("Replanteo").Cells(a + 2, 16) = "Inter.Section." Then
            p_aisl = 3.071 'daN
            aisl_sla = 1
            dist_aisl_1 = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 49) - 0.75
            dist_aisl_2 = va - Workbooks(1).Sheets("Replanteo").Cells(a + 1, 49) + 0.75
        End If
    
    Else
    
    '//
    '//Cálculo distancia primera segunda pendola
    '//
        
        dist_max_pend = 2.25
        
        dist_prim_seg_pend_izq = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 4
        
        If dist_prim_seg_pend_izq > 2.25 Then
        
            dist_prim_seg_pend_izq = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 8
        
        End If
        
        dist_prim_seg_pend_der = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 4
        
        If dist_prim_seg_pend_der > 2.25 Then
        
            dist_prim_seg_pend_der = (va - 4.5 * (Int((va / 4.5) + 0.99) - 2)) / 8
        
        End If
    
    End If

If (dist_ap_prim_pend_izq = dist_ap_prim_pend_der) And (el_hc_ini = el_hc_fin) Then

iteracion1:

    dist_ap_prim_pend = dist_ap_prim_pend_izq
    dist_prim_seg_pend = dist_prim_seg_pend_izq

    va_min = 2 * (dist_ap_prim_pend + dist_prim_seg_pend)
    var_0 = 2 * (dist_max_pend - dist_prim_seg_pend) + va_min
    
    If va < va_min Then
    
        GoTo X
    
    Else
        
        If va_min <= va And va <= var_0 Then
        
            Dist(1) = dist_ap_prim_pend
            Dist(2) = (va - (2 * dist_ap_prim_pend)) / 2
            Dist(3) = Dist(2)
            Dist(4) = Dist(1)
            
        Else:
            va_var = va - va_min
                
                If (va_var / 2) < dist_max_pend Then
                    Dist(1) = dist_ap_prim_pend
                    Dist(2) = dist_prim_seg_pend
                    If va_var > dist_max_pend Then
                        Dist(3) = va_var / 2
                        Dist(4) = va_var / 2
                        Dist(5) = Dist(2)
                        Dist(6) = Dist(1)
                    Else
                        Dist(3) = va_var
                        Dist(4) = Dist(2)
                        Dist(5) = Dist(1)
                    End If
                                
                ElseIf (va_var / 2) >= dist_max_pend Then
                        If (va_var) / dist_max_pend >= 1 Then
                            If Abs(Int((va_var / dist_max_pend)) - ((va_var) / dist_max_pend)) < 0.00001 Then
                                npi = Int((va_var / dist_max_pend))
                            
                            Else: npi = Int((va_var / dist_max_pend)) - 1
                            End If
                        Else: npi = 0
                        
                        End If
                        Dist(1) = dist_ap_prim_pend
                        Dist(2) = dist_prim_seg_pend
                            
                            If Abs(Int((va_var / dist_max_pend)) - ((va_var) / dist_max_pend)) < 0.00001 Then
                                Dist(3) = dist_max_pend
                                npi = npi
                            Else: Dist(3) = (va_var - dist_max_pend * npi) / 2
                            End If
                                       
                        i = 1
                        If Abs(Int((va_var / dist_max_pend)) - ((va_var) / dist_max_pend)) < 0.00001 Then
                            While i <= npi - 2
                                Dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        Else
                            While i <= npi
                                Dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        End If
                        Dist(i + 3) = Dist(3)
                        Dist(i + 4) = Dist(2)
                        Dist(i + 5) = Dist(1)
                                
                    'i = 2
                    'var_1 = 0
                    'While Dist(i) <> 0
                        'var_1 = var_1 + 1
                        'i = i + 1
                    'Wend
                    
                End If
                
                i = 2
                var_1 = 0
                While Dist(i) <> 0
                        var_1 = var_1 + 1
                        i = i + 1
                Wend
                
                
        End If
        
                i = 2
                var_1 = 0
                While Dist(i) <> 0
                        var_1 = var_1 + 1
                        i = i + 1
                Wend
        
    End If
'//
'//FASE DE PRUEBAS
'//
                
                If var_1 Mod 2 <> 0 And dist_max_pend <> 4.5 Then
                
                    var_1 = var_1 - 1
                
                    va_calculado = va - (dist_ap_prim_pend_izq + dist_ap_prim_pend_der + dist_max_pend * (var_1 - 5))
                
                    dist_prim_seg_pend_izq = va_calculado / 4
                    dist_prim_seg_pend_der = va_calculado / 4
                    
                    If dist_prim_seg_pend_izq > dist_max_pend Then
                        dist_prim_seg_pend_izq = va_calculado / 8
                        dist_prim_seg_pend_der = va_calculado / 8
                    End If
                    
                    cont = 1
                    While cont <= 100
                        Dist(cont) = 0
                        cont = cont + 1
                    Wend
                    
                    GoTo iteracion1
                
                End If
 
                
                'If Dist(2) > Dist(3) Then
                    'distancia_var = Dist(2)
                    'Dist(2) = Dist(3)
                    'Dist(3) = distancia_var
                    'distancia_var = Dist(var_1 - 1)
                    'Dist(var_1 - 1) = Dist(var_1)
                    'Dist(var_1) = distancia_var
                'End If

                    
                    distancia_var = 0
                    
                If Dist(2) <> Dist(3) Then
                    distancia_var = (Dist(2) + Dist(3)) / 2
                    Dist(2) = distancia_var
                    Dist(3) = distancia_var
                    Dist(var_1 - 1) = distancia_var
                    Dist(var_1) = distancia_var
                End If
                
    
          
          
'//
'//FASE DE PRUEBAS
'//
        i = 1
        j = 12
        var_dist = 0
        While Dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 8) = Dist(i)
            acum(i) = acum(i - 1) + plantilla.Sheets(1).Cells(j, 8)
            var_dist = var_dist + plantilla.Sheets(1).Cells(j, 8)
            plantilla.Sheets(1).Cells(j, 9) = var_dist
            i = i + 1
            j = j + 2
        Wend
        
        i = 2
        j = 13
        var_1 = 0
        While Dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 4) = var_1 + 1
            plantilla.Sheets(1).Cells(9, 4) = var_1 + 1
            var_1 = var_1 + 1
            i = i + 1
            j = j + 2
        Wend
    
Else

iteracion2:

    va_min = (dist_ap_prim_pend_izq + dist_ap_prim_pend_der + dist_prim_seg_pend_izq + dist_prim_seg_pend_der)
    var_0 = (2 * dist_max_pend - dist_prim_seg_pend_izq - dist_prim_seg_pend_der) + va_min
        
    If va < va_min Then
    
        GoTo X
    
    Else
        
    If va_min <= va And va <= var_0 Then
        
            Dist(1) = dist_ap_prim_pend_izq
            Dist(2) = (va - (dist_ap_prim_pend_izq + dist_ap_prim_pend_der)) / 2
            Dist(3) = Dist(2)
            Dist(4) = dist_ap_prim_pend_der
        
    Else:
            va_var = va - va_min
                
                If (va_var / 2) < dist_max_pend Then
                    Dist(1) = dist_ap_prim_pend_izq
                    Dist(2) = dist_prim_seg_pend_izq
                    If va_var > dist_max_pend Then
                        Dist(3) = va_var / 2
                        Dist(4) = va_var / 2
                        Dist(5) = dist_prim_seg_pend_der
                        Dist(6) = dist_ap_prim_pend_der
                    Else
                        Dist(3) = va_var
                        Dist(4) = dist_prim_seg_pend_der
                        Dist(5) = dist_ap_prim_pend_der
                    End If
                                
                ElseIf (va_var / 2) >= dist_max_pend Then
                        If (va_var) / dist_max_pend >= 1 Then
                            If Abs(Int(va_var / dist_max_pend) - (va_var / dist_max_pend)) < 0.00001 Then
                                npi = Int((va_var / dist_max_pend))
                            
                            Else: npi = Int((va_var / dist_max_pend)) - 1
                            End If
                        Else: npi = 0
                        
                        End If
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_prim_seg_pend_izq
                            
                            If Abs(Int((va_var / dist_max_pend)) - ((va_var) / dist_max_pend)) < 0.00001 Then
                                Dist(3) = dist_max_pend
                                npi = npi
                            Else: Dist(3) = (va_var - dist_max_pend * npi) / 2
                            End If
                                       
                        i = 1
                        If Abs(Int((va_var / dist_max_pend)) - ((va_var) / dist_max_pend)) < 0.00001 Then
                            While i <= npi - 2
                                Dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend
                        Else
                            While i <= npi
                                Dist(i + 3) = dist_max_pend
                                i = i + 1
                            Wend

                        End If
                        Dist(i + 3) = Dist(3)
                        Dist(i + 4) = dist_prim_seg_pend_der
                        Dist(i + 5) = dist_ap_prim_pend_der
                                        
                End If
                i = 2
                var_1 = 0
                While Dist(i) <> 0
                    var_1 = var_1 + 1
                    i = i + 1
                Wend
    End If
'//
'//FASE DE PRUEBAS
'//
                
                If var_1 Mod 2 <> 0 And dist_max_pend <> 4.5 Then
                
                    var_1 = var_1 - 1
                
                    va_calculado = va - (dist_ap_prim_pend_izq + dist_ap_prim_pend_der + dist_max_pend * (var_1 - 5))
                
                    dist_prim_seg_pend_izq = va_calculado / 4
                    dist_prim_seg_pend_der = va_calculado / 4
                                        
                    If dist_prim_seg_pend_izq > dist_max_pend Then
                        dist_prim_seg_pend_izq = va_calculado / 8
                        dist_prim_seg_pend_der = va_calculado / 8
                    End If
                    
                    cont = 1
                    While cont <= 100
                        Dist(cont) = 0
                        cont = cont + 1
                    Wend
                    
                    GoTo iteracion2
                    
                End If
                    
                i = 2
                var_1 = 0
                While Dist(i) <> 0
                    var_1 = var_1 + 1
                    i = i + 1
                Wend
                
                'If Dist(2) > Dist(3) Then
                    'distancia_var = Dist(2)
                    'Dist(2) = Dist(3)
                    'Dist(3) = distancia_var
                    'distancia_var = Dist(var_1 - 1)
                    'Dist(var_1 - 1) = Dist(var_1)
                    'Dist(var_1) = distancia_var
                'End If
                distancia_var = 0
                If Dist(2) <> Dist(3) Then
                    distancia_var = (Dist(2) + Dist(3)) / 2
                    Dist(2) = distancia_var
                    Dist(3) = distancia_var
                    Dist(var_1 - 1) = distancia_var
                    Dist(var_1) = distancia_var
                End If
                
End If
                
'//
'//FASE DE PRUEBAS
'//
                

        
         
ini3:

        i = 1
        j = 12
        var_dist = 0
        While Dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 8) = Dist(i)
            acum(i) = acum(i - 1) + plantilla.Sheets(1).Cells(j, 8)
            var_dist = var_dist + plantilla.Sheets(1).Cells(j, 8)
            plantilla.Sheets(1).Cells(j, 9) = var_dist
            i = i + 1
            j = j + 2
        Wend
        
        i = 2
        j = 13
        var_1 = 0
        While Dist(i) <> 0
            plantilla.Sheets(1).Cells(j, 4) = var_1 + 1
            plantilla.Sheets(1).Cells(9, 4) = var_1 + 1
            var_1 = var_1 + 1
            i = i + 1
            j = j + 2
        Wend
        
    End If
    
    
  
'//
'//Lectura densivel
'//

var_5 = 2
var_6 = 13
h_var = 0
d_var = 0
desn_contador = 0

plantilla.Sheets(1).Cells(var_6 - 2, 7) = 0

If a >= 3124 And a <= 3168 Then
    i = 1
    j = 0
    While i <= plantilla.Sheets(1).Cells(9, 4) + 1
        
        plantilla.Sheets(1).Cells(13 + j, 7) = 0
        j = j + 2
        i = i + 1
    Wend
        
GoTo fin_desnivel_bis

End If

pk_ini = Workbooks(1).Sheets("Replanteo").Cells(a, 3) + plantilla.Sheets(1).Cells(var_6 - 1, 8)
pk_fin = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)

'pk_ini = 56794 + plantilla.Sheets(1).Cells(var_6 - 1, 8)
'pk_fin = pk_ini + 25


ini_desnivel:

While pk_ini >= Workbooks(1).Sheets("Desnivel").Cells(var_5, 1)
    
    var_5 = var_5 + 2

Wend

If d_var = 0 Then

     d_var = plantilla.Sheets(1).Cells(var_6 - 1, 8)

End If


While pk_ini <= Workbooks(1).Sheets("Desnivel").Cells(var_5, 1)

    h_var = d_var * Workbooks(1).Sheets("Desnivel").Cells(var_5 - 1, 3) + h_var
    
    plantilla.Sheets(1).Cells(var_6, 7) = h_var + plantilla.Sheets(1).Cells(var_6 - 2, 7)
    
    var_6 = var_6 + 2
    pk_ini = pk_ini + plantilla.Sheets(1).Cells(var_6 - 1, 8)
    h_var = 0
    d_var = plantilla.Sheets(1).Cells(var_6 - 1, 8)
    desn_contador = desn_contador + 1
    
    If desn_contador >= var_1 + 1 Then

        GoTo fin_desnivel
    
    End If
    
Wend

    h_var = Workbooks(1).Sheets("Desnivel").Cells(var_5 - 1, 3) * (Workbooks(1).Sheets("Desnivel").Cells(var_5, 1) - (pk_ini - plantilla.Sheets(1).Cells(var_6 - 1, 8)))
    d_var = pk_ini - Workbooks(1).Sheets("Desnivel").Cells(var_5, 1)
    GoTo ini_desnivel
    
fin_desnivel:
fin_desnivel_bis:

desnivel_rasante = plantilla.Sheets(1).Cells(var_6 - 2, 7)
desnivel_alt_cat = plantilla.Sheets(1).Cells(6, 4) - plantilla.Sheets(1).Cells(5, 4)

desnivel = desnivel_rasante + desnivel_alt_cat

'//
'//APROXIMACIÓN DESNIVEL ENTE DOS POSTES (EN LA PARTE ANTERIOR SE CALCULA EN CADA PUNTO LA VARIACIÓN, ES MEJOR HACER LA APROXIMACIÓN PARA EL CASO QUE NOS TOCA Y PASAR DE LA TOPO INTERMEDIA)
'//

'tangente_desnivel = desnivel_rasante / va

'n_pend = 1
'i = 1
'j = 13
'While n_pend <= plantilla.Sheets(1).Cells(9, 4)

    'plantilla.Sheets(1).Cells(j, 7) = acum(i) * tangente_desnivel
    'j = j + 2
    'i = i + 1
    'n_pend = n_pend + 1

'Wend


'//
'//DESNIVEL IMPUESTO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'//
'//

'desnivel = (plantilla.Sheets(1).Cells(6, 4) - plantilla.Sheets(1).Cells(5, 4))

'//
'//DESNIVEL IMPUESTO!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'//


'//
'//Cálculo longitud media pendolas
'//

alt_pend_min = 0.25 'AÑADIR EN BASE DE DATOS

'i = 1

'While i <= (var_1 / 2)

    'alt_pend_calc = (plantilla.Sheets(1).Cells(5, 4) + plantilla.Sheets(1).Cells(6, 4)) / 2
    'long_pend(i) = alt_pend_calc - ((acum(i) / ((va - (Dist(1) + Dist(var_1 + 1))) / 2))) * (alt_pend_calc - alt_pend_min)
    'i = i + 1

'Wend

'i = 0
'j = 1

'While j <= (var_1 / 2)

    'long_pend(var_1 - i) = long_pend(j)
    
    'i = i + 1
    'j = j + 1

'Wend

i = 13
j = 1
While j <= (var_1)
    '//
    long_pend(j) = 1
    '//
    i = i + 2
    j = j + 1
Wend

'//
'//Consideración de la reducción de peso en caso de flecha intencional
'//
p_var_0 = p_hc * n_hc * Dist(1)

If (el_hc_ini <> 0 Or el_hc_fin <> 0) Or (el_hc_ini = el_hc_fin And el_hc_ini <> 0) Or (va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) <= 0 Then
    p_var_1 = 0
 
ElseIf plantilla.Sheets(1).Cells(9, 4) = 3 And (Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPI" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF") Then
    
    p_var_1 = (fl_max_centro_va * 2 * t_hc) / (((va - Dist(1) - Dist(var_1 + 1)) / 2) ^ 2)

 
Else
    p_var_1 = (fl_max_centro_va * 2 * t_hc) / (((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) / 2) ^ 2)

End If

p_var_2 = p_hc - p_var_1

    If p_var_2 < 0 Then 'SIGNIFICA QUE LA FLECHA INTENCIONAL GENERADA POR EL PESO SERÁ INFERIOR A LA IMPUESTA Y POR LO TANTO CON EL PESO DEL HILO LA FLECHA QUE SE GENERARÁ SERÁ IGUAL O INFERIOR POR LO QUE SE DEJA EL PESO DEL HILO DIRECTAMENTE, DE LO CONTRARIO, AL SER NEGATIVA ESTARIAMOS IMPONIENDO MAYOR PESO PARA FORZAR UNA FLECHA MAYOR
        
        p_var_1 = 0
        p_var_2 = p_hc

    End If

p_var_3 = (p_sust * va) / 2

p_sust_ap = (p_sust * va) / 2

'//
'//Cálculo de las elevaciones
'//

'//
'//Elevación hilo de contacto (a un lado)
'//

If el_hc_fin <> 0 Or el_hc_ini <> 0 And el_hc_ini > el_hc_fin Or (Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 43) = "var" And Dist(1) > Dist(plantilla.Sheets(1).Cells(9, 4))) Then
i = 1
cont = 1
n_pend = 1

'If el_hc_fin = 0 Then
'    el_hc_fin = el_hc_ini
'End If
'EN EL CASO SIGUIENTE SE INVIERTE LA TABLE Y LUEGO SE VUELVE A GENERAR EN SU ORDEN INICIAL

'//
'//INICIO CAMBIO DE ORDEN
'//

alt_cat_ini = plantilla.Sheets(1).Cells(5, 4)
alt_cat_fin = plantilla.Sheets(1).Cells(6, 4)
el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
el_hc_fin = plantilla.Sheets(1).Cells(8, 4)

If el_hc_ini > el_hc_fin Or (Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 43) = "var" And Dist(1) > Dist(plantilla.Sheets(1).Cells(9, 4))) Then

    el_hc_var = el_hc_fin
    el_hc_fin = el_hc_ini
    el_hc_ini = el_hc_var
    alt_cat_var = alt_cat_fin
    alt_cat_fin = alt_cat_ini
    alt_cat_ini = alt_cat_var
      
    cambio = 1
    
    i = 1
    cont = 1
    j = 13
    n_pend = plantilla.Sheets(1).Cells(9, 4)
                                                    
    While i <= plantilla.Sheets(1).Cells(9, 4) + 1
        dist_der(i) = Dist(i)
        l_pend_der(i) = plantilla.Sheets(1).Cells(j, 5)
        
        acum(i) = 0
        n_pend = n_pend - 1
        i = i + 1
        j = j + 2
    Wend
    
    i = 1
    cont = 1
    j = 13
    n_pend = plantilla.Sheets(1).Cells(9, 4)
    While i <= plantilla.Sheets(1).Cells(9, 4) + 2
        desnivel_cambio(i) = plantilla.Sheets(1).Cells(j - 2, 7)
        n_pend = n_pend - 1
        i = i + 1
        j = j + 2
    Wend
    
    i = 1
    j = 13
    acum(i) = 0
    n_pend = plantilla.Sheets(1).Cells(9, 4) + 1
    While i <= plantilla.Sheets(1).Cells(9, 4) + 2
        plantilla.Sheets(1).Cells(j - 2, 7) = desnivel_cambio(n_pend + 1)
        n_pend = n_pend - 1
        i = i + 1
        j = j + 2
    Wend
   
    i = 1
    j = 13
    acum(i) = 0
    n_pend = plantilla.Sheets(1).Cells(9, 4) + 1
        
       
    desnivel_rasante = -plantilla.Sheets(1).Cells(j - 2, 7)
    desnivel_alt_cat = plantilla.Sheets(1).Cells(5, 4) - plantilla.Sheets(1).Cells(6, 4)

    desnivel = desnivel_rasante + desnivel_alt_cat
        
    While i <= plantilla.Sheets(1).Cells(9, 4) + 1
        Dist(i) = dist_der(n_pend)
        plantilla.Sheets(1).Cells(j - 1, 8) = Dist(i)
       
        If l_pend_der(n_pend - 1) <> 0 Then
            plantilla.Sheets(1).Cells(j, 6) = l_pend_der(n_pend - 1)
            
            plantilla.Sheets(1).Cells(j, 5) = plantilla.Sheets(1).Cells(j, 6) - l_sup_pend - l_inf_pend
            End If
            
            acum(i) = acum(i - 1) + Dist(i)
            plantilla.Sheets(1).Cells(j - 1, 9) = acum(i)
                               
            n_pend = n_pend - 1
            i = i + 1
            j = j + 2
    Wend
 
End If

'//
'//FIN DE CAMBIO DE ORDEN
'//
        
    'x_var = Sqr(((el_hc_fin - el_hc_ini) * 2 * t_hc) / p_hc)
    'x_var = va - x_var
    
            'While cont <= i And n_pend <= plantilla.Sheets(1).Cells(9, 4)
            '    n(cont) = 1
            '    cont = cont + 1
            
            'If x_var < (n(1) * Dist(1) + n(2) * Dist(2) + n(3) * Dist(3) + n(4) * Dist(4) + n(5) * Dist(5) + n(6) * Dist(6) + n(7) * Dist(7) + n(8) * Dist(8) + n(9) * Dist(9) + n(10) * Dist(10) + n(11) * Dist(11) + n(12) * Dist(12) + n(13) * Dist(13) + n(14) * Dist(14) + n(15) * Dist(15) + n(16) * Dist(16) + n(17) * Dist(17) + n(18) * Dist(18) + n(19) * Dist(19) + n(20) * Dist(20) + n(21) * Dist(21) + n(22) * Dist(22) + n(23) * Dist(23) + n(24) * Dist(24) + n(25) * Dist(25) + n(26) * Dist(26) + n(27) * Dist(27) + n(28) * Dist(28)) Then
                
            '    el_hc_der(i) = (p_hc * (n(1) * Dist(1) + n(2) * Dist(2) + n(3) * Dist(3) + n(4) * Dist(4) + n(5) * Dist(5) + n(6) * Dist(6) + n(7) * Dist(7) + n(8) * Dist(8) + n(9) * Dist(9) + n(10) * Dist(10) + n(11) * Dist(11) + n(12) * Dist(12) + n(13) * Dist(13) + n(14) * Dist(14) + n(15) * Dist(15) + n(16) * Dist(16) + n(17) * Dist(17) + n(18) * Dist(18) + n(19) * Dist(19) + n(20) * Dist(20) + n(21) * Dist(21) + n(22) * Dist(22) + n(23) * Dist(23) + n(24) * Dist(24) + n(25) * Dist(25) + n(26) * Dist(26) + n(27) * Dist(27) + n(28) * Dist(28) - x_var * n(1)) ^ 2) / (2 * t_hc)
            
            'Else
            
            '    el_hc_der(i) = 0
            
            'End If
                         
            'i = i + 1
            'n_pend = n_pend + 1
                    
            'Wend
                            
End If

'//
'//CÁLCULO REACCIONES EN CADA PÉNDOLA
'//

i = 1
n_pend = 1
p_hc_ap = 0

'If el_hc_ini <> 0 And el_hc_fin <> 0 And el_hc_ini <> el_hc_fin Then

    'If el_hc_ini > el_hc_fin Then
        'el_hc_ini = el_hc_ini - el_hc_fin
        'el_hc_fin = 0
    'Else
        'el_hc_fin = el_hc_fin - el_hc_ini
        'el_hc_ini = 0
   ' End If
    
'End If
j = 0
cont = 1
While cont <= 30
    n(cont) = 0
    cont = cont + 1
Wend

'//
'//Cálculo de las reacciones sin elevación del hilo de contacto
'//

If (el_hc_ini = 0 And el_hc_fin = 0 Or el_hc_ini = el_hc_fin) And (Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 43) <> "var") Then

    While n_pend <= plantilla.Sheets(1).Cells(9, 4)
    
        If n_pend = 1 Then
        
            fuerza(i) = p_hc * Dist(i) + p_var_2 * (Dist(i + 1) + Dist(i + 2)) / 2 + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) / 2)
            
        ElseIf n_pend = 2 Then
        
            fuerza(i) = p_hc * (Dist(i - 1) + Dist(i)) + p_var_2 * (Dist(i + 1) + Dist(i + 2)) / 2 + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) / 2)
            
        ElseIf i = plantilla.Sheets(1).Cells(9, 4) - 1 Then
        
            fuerza(i) = p_hc * (Dist(i + 1) + Dist(i + 2)) + p_var_2 * (Dist(i) + Dist(i - 1)) / 2 + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) / 2)
        
        ElseIf i = plantilla.Sheets(1).Cells(9, 4) Then
            
            fuerza(i) = p_hc * Dist(i + 1) + p_var_2 * (Dist(i) + Dist(i - 1)) / 2 + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) / 2)

        Else
         
            fuerza(i) = p_var_2 * (Dist(i) + Dist(i - 1)) / 2 + p_var_2 * (Dist(i + 1) + Dist(i + 2)) / 2 + p_pend * long_pend(i) + p_pend_equip
        
        End If

        p_hc_ap = p_hc_ap + fuerza(i)
        i = i + 1
        n_pend = n_pend + 1
    
    Wend

'//
'//Cálculo de las reacciones con elevación del hilo de contacto
'//
Else
       
    j = 12
    i = 1
    n_pend = 1
    x_var = Sqr(((el_hc_fin - el_hc_ini) * 2 * t_hc) / p_hc)
    x_var = va - x_var
    
    While n_pend <= plantilla.Sheets(1).Cells(9, 4)
              
        fuerza(i) = 0
        If n_pend = 1 Then
              
            If x_var > acum(i) Then
             
                If x_var >= acum(i) + (acum(i + 2) - acum(i)) / 2 Then
                    
                    fuerza(i) = p_hc * (acum(i) + (acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(1))
                    
                    '//Contempla el caso de anclaje en toperas
                                      
                    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPI" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                                             
                            If plantilla.Sheets(1).Cells(9, 4) = 2 Then
                            
                                fuerza(i) = p_hc * (acum(i)) + p_var_2 * ((acum(i + 2) - acum(i))) + p_pend * long_pend(i) + p_pend_equip '+ (p_aisl) * (aisl_n(1))
                                
                            ElseIf plantilla.Sheets(1).Cells(9, 4) = 3 Then
                               
                               fuerza(i) = p_hc * (acum(i)) + p_var_2 * ((acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1 + 1)) / 2) '+ (p_aisl) * (aisl_n(1))
                               
                            ElseIf plantilla.Sheets(1).Cells(9, 4) = 4 Then
                               
                               fuerza(i) = p_hc * (acum(i)) + p_var_2 * ((acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) / 2) '+ (p_aisl) * (aisl_n(1))
                               
                               
                            End If
                                                
                    End If
                    
                    
                Else
                
                    fuerza(i) = p_hc * (acum(i) + (x_var - acum(i))) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(1))
                 
                End If
                 
            ElseIf x_var <= acum(i) Then
                
                    fuerza(i) = p_hc * x_var + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(1))
                    
            End If
                  
        ElseIf n_pend = 2 Then
              
            If x_var > acum(i) Then
             
                If x_var >= acum(i) + (acum(i + 2) - acum(i)) / 2 Then
                    
                    fuerza(i) = p_hc * (acum(i) + (acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(2))
                    
                    '//Contempla el caso de anclaje en toperas
                                      
                    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPI" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                        
                            If plantilla.Sheets(1).Cells(9, 4) = 2 Then
                            
                                fuerza(i) = p_hc * (acum(i)) + p_var_2 * ((acum(i + 1) - acum(i))) + p_pend * long_pend(i) + p_pend_equip '+ (p_aisl) * (aisl_n(1))
                            
                            ElseIf plantilla.Sheets(1).Cells(9, 4) = 3 Then
                               
                                fuerza(i) = p_hc * (acum(var_1 + 1)) + p_pend * long_pend(i) + p_pend_equip  '+ (p_aisl) * (aisl_n(1))
                               
                            ElseIf plantilla.Sheets(1).Cells(9, 4) = 4 Then
                               
                                fuerza(i) = p_hc * (acum(i)) + p_var_2 * ((acum(i + 1) - acum(i))) + p_pend * long_pend(i) + p_pend_equip '+ (p_aisl) * (aisl_n(1))
                               
                            End If
                    
                    End If
                    
                Else
                
                    fuerza(i) = p_hc * (acum(i) + (x_var - acum(i))) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(2))
                 
                End If
                 
            ElseIf x_var <= acum(i) Then
                
                    fuerza(i) = p_hc * x_var + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(2))
                    
            End If
               
        ElseIf n_pend = plantilla.Sheets(1).Cells(9, 4) And ((aisl_n(4) = 1 Or aisl_sla = 1) Or Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 43) = "var") Then
        
            If x_var > acum(i) Then
             
                If x_var >= acum(i) + (acum(i + 2) - acum(i)) / 2 Then
                    
                    fuerza(i) = p_hc * ((acum(i) - acum(i - 2)) / 2 + (acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(4)) + aisl_sla * (p_aisl / 2)
                    
                    '//Contempla el caso de anclaje en toperas
                                      
                    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPI" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                                
                        If plantilla.Sheets(1).Cells(9, 4) = 3 Then
                               
                            fuerza(i) = p_hc * (acum(i + 1) - acum(i)) + p_var_2 * ((acum(i) - acum(i - 2)) / 2) + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1 + 1)) / 2) '+ (p_aisl) * (aisl_n(1))
                           
                        ElseIf plantilla.Sheets(1).Cells(9, 4) = 4 Then
                           
                            fuerza(i) = p_hc * (acum(i + 1) - acum(i)) + p_var_2 * ((acum(i) - acum(i - 2)) / 2) + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1 + 1)) / 2) '+ (p_aisl) * (aisl_n(1))
                           
                        End If
                                    
                        
                    End If
                    
                Else
                
                    fuerza(i) = p_hc * ((acum(i) - acum(i - 2)) / 2 + (x_var - acum(i))) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(4)) + aisl_sla * (p_aisl / 2)
                 
                End If
                 
            ElseIf x_var <= acum(i) Then
            
                 If x_var > acum(i - 2) + (acum(i) - acum(i - 2)) / 2 Then
                
                    fuerza(i) = p_hc * ((x_var - acum(i - 2)) - (acum(i) - acum(i - 2)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(4)) + aisl_sla * (p_aisl / 2)
                    
                 Else
                 
                    fuerza(i) = p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(4)) + aisl_sla * (p_aisl / 2)
                 
                 End If
                
            End If
            
        ElseIf (n_pend = plantilla.Sheets(1).Cells(9, 4) - 1) And ((aisl_n(3) = 1 Or aisl_sla = 1) Or Workbooks(1).Worksheets("Replanteo").Cells(a + 1, 43) = "var") Then
        
            If x_var > acum(i) Then
             
                If x_var >= acum(i) + (acum(i + 2) - acum(i)) / 2 Then
                    
                    fuerza(i) = p_hc * ((acum(i) - acum(i - 2)) / 2 + (acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3)) + aisl_sla * (p_aisl / 2)
                 
                    '//Contempla el caso de anclaje en toperas
                                      
                    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPI" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                        
                        If plantilla.Sheets(1).Cells(9, 4) = 3 Then
                        
                            fuerza(i) = p_var_2 * ((acum(i) - acum(i - 2)) / 2) + p_hc * (acum(i + 2) - acum(i)) + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1 + 1) - Dist(var_1)) / 2) + (p_aisl / 2) * (aisl_n(3))
                    
                        ElseIf plantilla.Sheets(1).Cells(9, 4) = 4 Then
                               
                            fuerza(i) = p_hc * (acum(i + 2) - acum(i)) + p_var_2 * ((acum(i) - acum(i - 2)) / 2) + p_pend * long_pend(i) + p_pend_equip + p_var_1 * ((va - Dist(1) - Dist(var_1 + 1)) / 2) '+ (p_aisl) * (aisl_n(1))
                               
                        End If
                    
                    End If
                 
                Else
                
                    fuerza(i) = p_hc * ((acum(i) - acum(i - 2)) / 2 + (x_var - acum(i))) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3)) + aisl_sla * (p_aisl / 2)
                 
                End If
                 
            ElseIf x_var <= acum(i) Then
                
                If x_var > acum(i - 2) + (acum(i) - acum(i - 2)) / 2 Then
                
                    fuerza(i) = p_hc * ((x_var - acum(i - 2)) - (acum(i) - acum(i - 2)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3)) + aisl_sla * (p_aisl / 2)
                
                Else
                 
                    fuerza(i) = p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3)) + aisl_sla * (p_aisl / 2)
                 
                End If
                    
            End If
        
        Else
        
            If x_var > acum(i) Then
             
                If x_var >= acum(i) + (acum(i + 2) - acum(i)) / 2 Then
                    
                    fuerza(i) = p_hc * ((acum(i) - acum(i - 2)) / 2 + (acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3))
                    
                    '//Contempla el caso de anclaje en toperas
                                      
                    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPI" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                   
                        fuerza(i) = p_var_2 * ((acum(i) - acum(i - 2)) / 2 + (acum(i + 2) - acum(i)) / 2) + p_pend * long_pend(i) + p_pend_equip '+ (p_aisl/2) * (aisl_n(1))
                    
                    End If
                    
                Else
                
                    fuerza(i) = p_hc * ((acum(i) - acum(i - 2)) / 2 + (x_var - acum(i))) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3))
                 
                End If
                 
            ElseIf x_var <= acum(i) Then
                
                If x_var > acum(i - 2) + (acum(i) - acum(i - 2)) / 2 Then
                
                    fuerza(i) = p_hc * ((x_var - acum(i - 2)) - (acum(i) - acum(i - 2)) / 2) + p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3))
                
                Else
                 
                    fuerza(i) = p_pend * long_pend(i) + p_pend_equip + (p_aisl / 2) * (aisl_n(3))
                 
                End If
                    
            End If
            
        End If
             
        i = i + 1
        n_pend = n_pend + 1
        j = j + 2
    
    Wend
                    


End If

'//
'//Búsqueda de las reacciones en los apoyos
'//

i = 1
cont = 1
n_pend = plantilla.Sheets(1).Cells(9, 4)

While cont <= plantilla.Sheets(1).Cells(9, 4)
    n(cont) = 1
    cont = cont + 1
Wend

'If (aisl_n(1) = 1 Or aisl_n(2) = 1 Or aisl_n(3) = 1 Or aisl_n(4) = 1 Or aisl_n(5) = 1 Or aisl_n(6) = 1) Then

    'While i <= plantilla.Sheets(1).Cells(9, 4)
        'fuerza_var(i) = fuerza(i)
        'aisl_n_var(i) = aisl_n(i)
        'acum_aisl_var(i) = acum_aisl(i)
        'i = i + 1
    'Wend
    'i = 1
    'While i <= plantilla.Sheets(1).Cells(9, 4)
        'fuerza(n_pend) = fuerza_var(i)
        'aisl_n(n_pend) = aisl_n_var(i)
        'acum_aisl(n_pend) = acum_aisl_var(i)
        'i = i + 1
        'n_pend = n_pend - 1
    'Wend
'End If
i = 1
While i <= 10
    If acum_aisl(i) = 0 Then
    
        aisl_n(i) = 0
    
    End If
    
    i = i + 1
Wend

i = 1
cont = 1
n_pend = plantilla.Sheets(1).Cells(9, 4)

p_ap_tot_der = 0
p_ap_tot_der_aux = p_sust * (va / 2) + (desnivel_alt_cat / va) * Sqr(t_sust ^ 2 - p_ap_tot_der ^ 2) + aisl_sla * dist_aisl_2 * p_aisl * (1 / va) + (p_aisl * ((aisl_n(1) * (acum(acum_aisl(1)) - 2)) + (aisl_n(2) * (acum(acum_aisl(2)) + 2)) + (aisl_n(3) * (acum(acum_aisl(3)) + 2)) + (aisl_n(4) * (acum(acum_aisl(4)) + 2)))) / va + (1 / va) * (n(1) * fuerza(1) * acum(1) + n(2) * fuerza(2) * acum(2) + n(3) * fuerza(3) * acum(3) + n(4) * fuerza(4) * acum(4) + n(5) * fuerza(5) * acum(5) + n(6) * fuerza(6) * acum(6) + n(7) * fuerza(7) * acum(7) + n(8) * fuerza(8) * acum(8) + n(9) * fuerza(9) * acum(9) + n(10) * fuerza(10) * acum(10) _
+ n(11) * fuerza(11) * acum(11) + n(12) * fuerza(12) * acum(12) + n(13) * fuerza(13) * acum(13) + n(14) * fuerza(14) * acum(14) + n(15) * fuerza(15) * acum(15) + n(16) * fuerza(16) * acum(16) + n(17) * fuerza(17) * acum(17) + n(18) * fuerza(18) * acum(18) + n(19) * fuerza(19) * acum(19) + n(20) * fuerza(20) * acum(20) + n(21) * fuerza(21) * acum(21) + n(22) * fuerza(22) * acum(22) + n(23) * fuerza(23) * acum(23) + n(24) * fuerza(24) * acum(24))

While Abs(p_ap_tot_der_aux - p_ap_tot_der) > 0.001
    p_ap_tot_der = (p_ap_tot_der_aux + p_ap_tot_der) / 2
    p_ap_tot_der_aux = p_sust * (va / 2) + (desnivel_alt_cat / va) * Sqr(t_sust ^ 2 - p_ap_tot_der ^ 2) + aisl_sla * dist_aisl_2 * p_aisl * (1 / va) + (p_aisl * ((aisl_n(1) * (acum(acum_aisl(1)) - 2)) + (aisl_n(2) * (acum(acum_aisl(2)) + 2)) + (aisl_n(3) * (acum(acum_aisl(3)) + 2)) + (aisl_n(4) * (acum(acum_aisl(4)) + 2)))) / va + (1 / va) * (n(1) * fuerza(1) * acum(1) + n(2) * fuerza(2) * acum(2) + n(3) * fuerza(3) * acum(3) + n(4) * fuerza(4) * acum(4) + n(5) * fuerza(5) * acum(5) + n(6) * fuerza(6) * acum(6) + n(7) * fuerza(7) * acum(7) + n(8) * fuerza(8) * acum(8) + n(9) * fuerza(9) * acum(9) + n(10) * fuerza(10) * acum(10) _
    + n(11) * fuerza(11) * acum(11) + n(12) * fuerza(12) * acum(12) + n(13) * fuerza(13) * acum(13) + n(14) * fuerza(14) * acum(14) + n(15) * fuerza(15) * acum(15) + n(16) * fuerza(16) * acum(16) + n(17) * fuerza(17) * acum(17) + n(18) * fuerza(18) * acum(18) + n(19) * fuerza(19) * acum(19) + n(20) * fuerza(20) * acum(20) + n(21) * fuerza(21) * acum(21) + n(22) * fuerza(22) * acum(22) + n(23) * fuerza(23) * acum(23) + n(24) * fuerza(24) * acum(24))
Wend
p_ap_tot_der = p_ap_tot_der_aux

p_ap_tot_izq_var = 0
i = 1
While i <= plantilla.Sheets(1).Cells(9, 4)
    p_ap_tot_izq_var = p_ap_tot_izq_var + n(i) * fuerza(i) * (acum(n_pend + 1) - acum(i))
        
    i = i + 1
    
Wend

'p_ap_tot_der = p_sust * (va / 2) + p_ap_tot_der_var / va + (1 / va) * (t_hc * desnivel)
p_ap_tot_izq = 0
p_ap_tot_izq_aux = p_sust * (va / 2) - (desnivel_alt_cat / va) * Sqr(t_sust ^ 2 - p_ap_tot_izq ^ 2) + (1 / va) * p_ap_tot_izq_var + aisl_sla * dist_aisl_1 * p_aisl * (1 / va) + (p_aisl * ((aisl_n(1) * ((acum(n_pend + 1) - acum(acum_aisl(1))) - 2)) + (aisl_n(2) * ((acum(n_pend + 1) - acum(acum_aisl(2))) - 2)) + (aisl_n(3) * ((acum(n_pend + 1) - acum(acum_aisl(3))) - 2)) + (aisl_n(4) * ((acum(n_pend + 1) - acum(acum_aisl(4))) - 2)))) / va

While Abs(p_ap_tot_izq_aux - p_ap_tot_izq) > 0.00001
    p_ap_tot_izq = (p_ap_tot_izq_aux + p_ap_tot_izq) / 2
    p_ap_tot_izq_aux = p_sust * (va / 2) - (desnivel_alt_cat / va) * Sqr(t_sust ^ 2 - p_ap_tot_izq ^ 2) + (1 / va) * p_ap_tot_izq_var + aisl_sla * dist_aisl_1 * p_aisl * (1 / va) + (p_aisl * ((aisl_n(1) * ((acum(n_pend + 1) - acum(acum_aisl(1))) - 2)) + (aisl_n(2) * ((acum(n_pend + 1) - acum(acum_aisl(2))) - 2)) + (aisl_n(3) * ((acum(n_pend + 1) - acum(acum_aisl(3))) - 2)) + (aisl_n(4) * ((acum(n_pend + 1) - acum(acum_aisl(4))) - 2)))) / va
Wend
p_ap_tot_izq = p_ap_tot_izq_aux

'If cambio = 1 Then

    'p_ap_tot_var = p_ap_tot_izq
    'p_ap_tot_izq = p_ap_tot_der
    'p_ap_tot_der = p_ap_tot_var
    
'End If


cont = 0
While cont <= 30
    n(cont) = 0
    cont = cont + 1
Wend

'//
'//Distintas alturas de catenaria
'//
         
'If plantilla.Sheets(1).Cells(5, 4) <> plantilla.Sheets(1).Cells(6, 4) Then
    
'    p_ap_tot_izq = (p_ap_tot_izq) + (t_sust * (plantilla.Sheets(1).Cells(6, 4) - plantilla.Sheets(1).Cells(5, 4)) / va)
'    p_ap_tot_der = (p_ap_tot_der) + (t_sust * (plantilla.Sheets(1).Cells(5, 4) - plantilla.Sheets(1).Cells(6, 4)) / va)
    
'End If

'//
'//CÁLCULO MOMENTO EN CADA PÉNDOLA
'//

i = 1
n_pend = 1
p_hc_ap = 0
cont = 1
j = 12

While n_pend <= plantilla.Sheets(1).Cells(9, 4)

    'If n_pend = 1 Then
    
    '    mom(i) = p_ap_tot_izq * acum(i) - (p_sust / 2) * (acum(i) ^ 2)
   
    'ElseIf i = plantilla.Sheets(1).Cells(9, 4) And plantilla.Sheets(1).Cells(5, 4) <> plantilla.Sheets(1).Cells(6, 4) Then
    
    '    mom(i) = (p_ap_tot_der) * plantilla.Sheets(1).Cells(12 + var_1 * 2, 8) - (p_sust / 2) * (plantilla.Sheets(1).Cells(j + 2, 8) ^ 2)
  
    'Else
     
        While cont <= i - 1
            n(cont) = 1
            cont = cont + 1
        Wend
        
        mom(i) = (p_ap_tot_izq) * acum(i) - (p_sust / 2) * (acum(i) ^ 2) - (n(1) * fuerza(1) * (acum(i) - (Dist(1))) + n(2) * fuerza(2) * (acum(i) - (Dist(1) + Dist(2))) + n(3) * fuerza(3) * (acum(i) - (Dist(1) + Dist(2) + Dist(3))) + n(4) * fuerza(4) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4))) + n(5) * fuerza(5) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5))) + n(6) * fuerza(6) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6))) + _
        n(7) * fuerza(7) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7))) + n(8) * fuerza(8) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8))) + n(9) * fuerza(9) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9))) + n(10) * fuerza(10) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10))) + _
        n(11) * fuerza(11) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11))) + n(12) * fuerza(12) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12))) + n(13) * fuerza(13) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13))) + n(14) * fuerza(14) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14))) + n(15) * fuerza(15) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15))) + _
        n(16) * fuerza(16) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16))) + _
        n(17) * fuerza(17) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17))) + n(18) * fuerza(18) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18))) + n(19) * fuerza(19) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19))) + _
        n(20) * fuerza(20) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20))) + n(21) * fuerza(21) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21))) + n(22) * fuerza(22) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22))) + _
        n(23) * fuerza(23) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23))) + _
        n(24) * fuerza(24) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24))) + n(25) * fuerza(25) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25))) + n(26) * fuerza(26) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25) + Dist(26))) + _
        n(27) * fuerza(27) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25) + Dist(26) + Dist(27))) + n(27) * fuerza(27) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25) + Dist(26) + Dist(27) + Dist(28))))
        
        'mom(i) = (p_ap_tot_izq) * acum(i) - (p_sust / 2) * (acum(i) ^ 2) - (p_aisl * ((n(3) * aisl_n(1) * (0)) + (n(1) * aisl_n(2) * (acum(acum_aisl(2) + 1) - acum(acum_aisl(2)) + 2)) + (n(2) * aisl_n(3) * (0)) + (n(1) * aisl_n(4) * (acum(acum_aisl(4) + 1) - acum(acum_aisl(4)) + 2)) + (n(1) * aisl_n(5) * (0)) + (n(1) * aisl_n(6) * (acum(acum_aisl(6) + 1) - acum(acum_aisl(6)) + 2)))) - (n(1) * fuerza(1) * (acum(i) - (Dist(1))) + n(2) * fuerza(2) * (acum(i) - (Dist(1) + Dist(2))) + n(3) * fuerza(3) * (acum(i) - (Dist(1) + Dist(2) + Dist(3))) + n(4) * fuerza(4) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4))) + n(5) * fuerza(5) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5))) + n(6) * fuerza(6) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6))) + _
        'n(7) * fuerza(7) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7))) + n(8) * fuerza(8) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8))) + n(9) * fuerza(9) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9))) + n(10) * fuerza(10) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10))) + _
        'n(11) * fuerza(11) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11))) + n(12) * fuerza(12) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12))) + n(13) * fuerza(13) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13))) + n(14) * fuerza(14) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14))) + n(15) * fuerza(15) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15))) + _
        'n(16) * fuerza(16) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16))) + _
        'n(17) * fuerza(17) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17))) + n(18) * fuerza(18) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18))) + n(19) * fuerza(19) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19))) + _
        'n(20) * fuerza(20) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20))) + n(21) * fuerza(21) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21))) + n(22) * fuerza(22) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22))) + _
        'n(23) * fuerza(23) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23))) + _
        'n(24) * fuerza(24) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24))) + n(25) * fuerza(25) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25))) + n(26) * fuerza(26) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25) + Dist(26))) + _
        'n(27) * fuerza(27) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25) + Dist(26) + Dist(27))) + n(27) * fuerza(27) * (acum(i) - (Dist(1) + Dist(2) + Dist(3) + Dist(4) + Dist(5) + Dist(6) + Dist(7) + Dist(8) + Dist(9) + Dist(10) + Dist(11) + Dist(12) + Dist(13) + Dist(14) + Dist(15) + Dist(16) + Dist(17) + Dist(18) + Dist(19) + Dist(20) + Dist(21) + Dist(22) + Dist(23) + Dist(24) + Dist(25) + Dist(26) + Dist(27) + Dist(28))))
    
    'End If
    
    i = i + 1
    j = j + 2
    n_pend = n_pend + 1

Wend

'//
'//CÁLCULO FLECHA SUSTENTADOR E HILO DE CONTACTO
'//

'//
'//Flecha hilo de contacto
'//

If el_hc_ini = 0 And el_hc_fin = 0 Then

i = 1
n_pend = 1
j = 12
While n_pend <= plantilla.Sheets(1).Cells(9, 4)

If n_pend = 1 Then
    
        fl_hc(i) = 0
        
    ElseIf i = 2 Then
    
        fl_hc(i) = 0
    
    ElseIf i = plantilla.Sheets(1).Cells(9, 4) - 1 Then
    
        fl_hc(i) = fl_hc(2)
        
    ElseIf i = plantilla.Sheets(1).Cells(9, 4) Then
        
        fl_hc(i) = fl_hc(1)
    
    Else
    
        If n_pend Mod 2 <> 0 Then 'impar
        
            fl_hc(i) = (p_var_1 / (2 * t_hc)) * (acum(i) - Dist(1)) * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) - (acum(i) - Dist(1)))
            
        Else 'par
        
            fl_hc(i) = (p_var_1 / (2 * t_hc)) * (acum(i) - Dist(1) - Dist(2)) * ((va - Dist(1) - Dist(var_1) - Dist(var_1 + 1)) - (acum(i) - Dist(1) - Dist(2)))
    
        End If
        
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

'Descomposición fuerza
t_horiz_sust = Sqr((t_sust ^ 2) - (p_ap_tot_izq ^ 2))

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
'//CÁLCULO ELEVACIONES
'//

'//
'//CÁLCULO ELEVACIÓN
'//
If el_hc_ini <> 0 Or el_hc_fin <> 0 And el_hc_ini <> el_hc_fin Then

    'If cambio = 0 Then
    
        'i = 1
        'n_pend = 1
        'x_var = Sqr(((el_hc_fin - el_hc_ini) * 2 * t_hc) / p_hc)
        'x_var = va - x_var
        
        'While n_pend <= plantilla.Sheets(1).Cells(9, 4)
        
            'el_hc(i) = ((acum(i) - x_var) ^ 2 * p_hc) / (2 * t_hc)
            'i = i + 1
            'n_pend = n_pend + 1
    
        'Wend
    
    'ElseIf cambio = 1 Then
    
        i = 1
        n_pend = 1
        x_var = Sqr(((el_hc_fin - el_hc_ini) * 2 * t_hc) / p_hc)
        x_var = va - x_var
        While n_pend <= plantilla.Sheets(1).Cells(9, 4)
            If x_var > acum(i) Then
             
                  
                     el_hc(i) = 0
                  
            ElseIf x_var <= acum(i) Then
                
                    el_hc(i) = ((acum(n_pend) - x_var) ^ 2 * p_hc) / (2 * t_hc)
                    
            End If

            i = i + 1
            n_pend = n_pend + 1
    
        Wend
    


End If



'//
'//CÁLCULO LONGITUD PÉNDOLAS
'//
i = 1
n_pend = 1
j = 13

While n_pend <= plantilla.Sheets(1).Cells(9, 4)

    'If n_pend = 1 Then
    
    If cambio = 0 Then
   
        If plantilla.Worksheets(1).Cells(j, 7) < 0 Then
        
            plantilla.Worksheets(1).Cells(j, 5) = alt_cat_ini - el_hc_ini - fl_sust(i) + fl_hc(i) - el_hc(i) '- (Abs(desnivel_rasante) - Abs(plantilla.Worksheets(1).Cells(j, 7)))
            plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
            '//
            plantilla.Worksheets(1).Cells(j, 10) = plantilla.Worksheets(1).Cells(j, 6) - 0.028
            '//
        
        ElseIf plantilla.Worksheets(1).Cells(j, 7) >= 0 Then
        
            plantilla.Worksheets(1).Cells(j, 5) = alt_cat_ini - el_hc_ini - fl_sust(i) + fl_hc(i) - el_hc(i) '- plantilla.Worksheets(1).Cells(j, 7)
            plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
            '//
            plantilla.Worksheets(1).Cells(j, 10) = plantilla.Worksheets(1).Cells(j, 6) - 0.028
        
        End If
        
    ElseIf cambio = 1 Then
    
        If plantilla.Worksheets(1).Cells(j, 7) < 0 Then
        
            plantilla.Worksheets(1).Cells(j, 5) = alt_cat_ini - el_hc_ini - fl_sust(i) + fl_hc(i) - el_hc(i) '- (Abs(desnivel_rasante) - Abs(plantilla.Worksheets(1).Cells(j, 7)))
            plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
            '//
            plantilla.Worksheets(1).Cells(j, 10) = plantilla.Worksheets(1).Cells(j, 6) - 0.028
            '//
        
        ElseIf plantilla.Worksheets(1).Cells(j, 7) >= 0 Then
        
            plantilla.Worksheets(1).Cells(j, 5) = alt_cat_ini - el_hc_ini - fl_sust(i) + fl_hc(i) - el_hc(i) '- plantilla.Worksheets(1).Cells(j, 7)
            plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
            '//
            plantilla.Worksheets(1).Cells(j, 10) = plantilla.Worksheets(1).Cells(j, 6) - 0.028
        
        End If
    End If
    'ElseIf i = plantilla.Sheets(1).Cells(9, 4) Then
        
        'plantilla.Worksheets(1).Cells(j, 5) = plantilla.Worksheets(1).Cells(6, 4) - fl_sust(i) + fl_hc(i) - el_hc_izq(i) - el_hc_der(i)
        'plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
        '//
        'plantilla.Worksheets(1).Cells(j, 10) = plantilla.Worksheets(1).Cells(j, 6) - 0.028
        '//
    
    'Else
     
        'plantilla.Worksheets(1).Cells(j, 5) = plantilla.Worksheets(1).Cells(5, 4) - fl_sust(i) + fl_hc(i) - el_hc_izq(i) - el_hc_der(i)
        'plantilla.Worksheets(1).Cells(j, 6) = plantilla.Worksheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
        '//
        'plantilla.Worksheets(1).Cells(j, 10) = plantilla.Worksheets(1).Cells(j, 6) - 0.028
        '//
        
    'End If
    
n_pend = n_pend + 1
i = i + 1
j = j + 2

Wend


'//
'//SE VUELVE A ORDENAR ADECUADAMENTE
'//
If cambio = 1 Then

    el_hc_var = el_hc_fin
    el_hc_fin = el_hc_ini
    el_hc_ini = el_hc_var
    alt_cat_var = alt_cat_fin
    alt_cat_fin = alt_cat_ini
    alt_cat_ini = alt_cat_var
    
    cambio = 0
    
    i = 1
    cont = 1
    j = 13
    n_pend = plantilla.Sheets(1).Cells(9, 4)
                                                    
    While i <= plantilla.Sheets(1).Cells(9, 4) + 1
        dist_der(i) = Dist(i)
        l_pend_der(i) = plantilla.Sheets(1).Cells(j, 5)
        
        acum(i) = 0
        n_pend = n_pend - 1
        i = i + 1
        j = j + 2
    Wend
    
    i = 1
    cont = 1
    j = 13
    n_pend = plantilla.Sheets(1).Cells(9, 4)
    While i <= plantilla.Sheets(1).Cells(9, 4) + 2
        desnivel_cambio(i) = plantilla.Sheets(1).Cells(j - 2, 7)
        n_pend = n_pend - 1
        i = i + 1
        j = j + 2
    Wend
    
    i = 1
    j = 13
    acum(i) = 0
    n_pend = plantilla.Sheets(1).Cells(9, 4) + 1
    While i <= plantilla.Sheets(1).Cells(9, 4) + 2
        plantilla.Sheets(1).Cells(j - 2, 7) = desnivel_cambio(n_pend + 1)
        n_pend = n_pend - 1
        i = i + 1
        j = j + 2
    Wend
    
    i = 1
    j = 13
    acum(i) = 0
    n_pend = plantilla.Sheets(1).Cells(9, 4) + 1
        
    While i <= plantilla.Sheets(1).Cells(9, 4) + 1
        Dist(i) = dist_der(n_pend)
        plantilla.Sheets(1).Cells(j - 1, 8) = Dist(i)
       
        If l_pend_der(n_pend - 1) <> 0 Then
            plantilla.Sheets(1).Cells(j, 5) = l_pend_der(n_pend - 1)
            
            plantilla.Sheets(1).Cells(j, 6) = plantilla.Sheets(1).Cells(j, 5) - l_sup_pend - l_inf_pend
            plantilla.Sheets(1).Cells(j, 10) = plantilla.Sheets(1).Cells(j, 6) - 0.028
            End If
            
            acum(i) = acum(i - 1) + Dist(i)
            plantilla.Sheets(1).Cells(j - 1, 9) = acum(i)
                               
            n_pend = n_pend - 1
            i = i + 1
            j = j + 2
    Wend
    
    '//
    '//Lectura desnivel
    '//
    
    'var_5 = 2
    'var_6 = 13
    'h_var = 0
    'd_var = 0
    'desn_contador = 0
    
    'plantilla.Sheets(1).Cells(var_6 - 2, 7) = 0
    
    'pk_ini = Workbooks(1).Sheets("Replanteo").Cells(a, 3) + plantilla.Sheets(1).Cells(var_6 - 1, 8)
    'pk_fin = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)
    
'ini_desnivel_cambio2:
    
    'While pk_ini >= Workbooks(1).Sheets("Desnivel").Cells(var_5, 1)
        
        'var_5 = var_5 + 2
    
   ' Wend
    
    'If d_var = 0 Then
    
         'd_var = plantilla.Sheets(1).Cells(var_6 - 1, 8)
    
    'End If
    
    
    'While pk_ini <= Workbooks(1).Sheets("Desnivel").Cells(var_5, 1)
    
        'h_var = d_var * Workbooks(1).Sheets("Desnivel").Cells(var_5 - 1, 3) + h_var
        
        'plantilla.Sheets(1).Cells(var_6, 7) = h_var + plantilla.Sheets(1).Cells(var_6 - 2, 7)
        
        'var_6 = var_6 + 2
        'pk_ini = pk_ini + plantilla.Sheets(1).Cells(var_6 - 1, 8)
        'h_var = 0
        'd_var = plantilla.Sheets(1).Cells(var_6 - 1, 8)
        'desn_contador = desn_contador + 1
        
        'If desn_contador >= var_1 + 1 Then
    
            'GoTo fin_desnivel_cambio2
        
        'End If
        
    'Wend
    
        'h_var = Workbooks(1).Sheets("Desnivel").Cells(var_5 - 1, 3) * (Workbooks(1).Sheets("Desnivel").Cells(var_5, 1) - (pk_ini - plantilla.Sheets(1).Cells(var_6 - 1, 8)))
        'd_var = pk_ini - Workbooks(1).Sheets("Desnivel").Cells(var_5, 1)
        'GoTo ini_desnivel_cambio2
        
'fin_desnivel_cambio2:
    
    'desnivel_rasante = plantilla.Sheets(1).Cells(var_6 - 2, 7)
    'desnivel_alt_cat = plantilla.Sheets(1).Cells(6, 4) - plantilla.Sheets(1).Cells(5, 4)
    
    'desnivel = desnivel_rasante + desnivel_alt_cat
    
  
        
End If

'//
'//FIN DEL CAMBIO
'//

cambio = 0



'//
'//FORMATO PLANTILLA
'//

For i = 13 To var_1 * 2 + 11

    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 4), plantilla.Sheets(1).Cells(i + 1, 6)).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 4), plantilla.Sheets(1).Cells(i + 1, 6)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 4), plantilla.Sheets(1).Cells(i + 1, 6)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 4), plantilla.Sheets(1).Cells(i + 1, 6)).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 4), plantilla.Sheets(1).Cells(i + 1, 6)).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 10), plantilla.Sheets(1).Cells(i + 1, 10)).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 10), plantilla.Sheets(1).Cells(i + 1, 10)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 10), plantilla.Sheets(1).Cells(i + 1, 10)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 10), plantilla.Sheets(1).Cells(i + 1, 10)).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 10), plantilla.Sheets(1).Cells(i + 1, 10)).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With

Next i

For i = 11 To var_1 * 2 + 13

    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 7), plantilla.Sheets(1).Cells(i + 1, 7)).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 7), plantilla.Sheets(1).Cells(i + 1, 7)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 7), plantilla.Sheets(1).Cells(i + 1, 7)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 7), plantilla.Sheets(1).Cells(i + 1, 7)).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 7), plantilla.Sheets(1).Cells(i + 1, 7)).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With

Next i

For i = 12 To var_1 * 2 + 12

    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 8), plantilla.Sheets(1).Cells(i + 1, 9)).Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 8), plantilla.Sheets(1).Cells(i + 1, 9)).Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 8), plantilla.Sheets(1).Cells(i + 1, 9)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 8), plantilla.Sheets(1).Cells(i + 1, 9)).Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With
    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(i, 8), plantilla.Sheets(1).Cells(i + 1, 9)).Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Weight = xlThin
    End With

Next i

    With plantilla.Sheets(1).Range(plantilla.Sheets(1).Cells(10, 4), plantilla.Sheets(1).Cells(10, 10)).Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Weight = xlMedium
    End With
    

'//
'//
'//

plantilla.Sheets(1).PageSetup.RightFooter = contador_hojas

plantilla_control.Sheets(1).PageSetup.CenterHeader = vbCrLf & "&""Trebuchet MS,Bold""&12 " & "SOMMAIRE"
With plantilla_control.Sheets("control").PageSetup
    .RightFooter = "&""Arial,Bold""&12 "
End With
plantilla_control.Sheets("control").Cells(contador_hojas + 8, 1) = "FOLIO" & " " & contador_hojas & " - " & plantilla.Sheets(1).Cells(6, 11) & " / " & plantilla.Sheets(1).Cells(8, 11) & " - " & plantilla.Sheets(1).Cells(7, 5)
plantilla_control.Sheets("control").Cells(contador_hojas + 8, 3) = "+"
contador_hojas = contador_hojas + 1


'//
'//GENERACIÓN FICHA
'//
b = 1

    If st = 0 Then
    
        d = Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value
        e = Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 1).Value
    
    Else
    
        d = Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value & "_1"
        e = Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 1).Value & "_1"
        
    End If


    'fncScreenUpdating State:=False
    Call plantilla.Worksheets(b).PrintOut(from:=1, To:=1, Copies:=1, preview:=False, ActivePrinter:="Adobe PDF", printtofile:=True, collate:=False, prtofilename:=ruta_replanteoVB & "\" & d & " " & e & ".ps")
    'fncScreenUpdating State:=True
    PSFileName = ruta_replanteoVB & "\" & d & " " & e & ".ps"
    PDFFileName = ruta_replanteoVB & "\" & d & " " & e & ".pdf"
    TXTFileName = ruta_replanteoVB & "\" & d & " " & e & ".log"
    mypdf.FileToPDF PSFileName, PDFFileName, ""
    fso.DeleteFile PSFileName, True
    fso.DeleteFile TXTFileName, True
'//
'//INSERCIÓN FICHA EN PDF GLOBAL
'//
    Call CombPDF(PDFFijoFileName, PDFFileName, ruta_replanteoVB)
'//
'//Inicialización variables
'//
i = 0
While i <= 100
    Dist(i) = 0
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
While i <= 30
    el_hc(i) = 0
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
i = 0
While i <= 10
    aisl_n(i) = 0
    i = i + 1
Wend
i = 0
While i <= 10
    acum_aisl(i) = 0
    i = i + 1
Wend
i = 0
While i <= 40
    fuerza_var(i) = 0
    i = i + 1
Wend
i = 0
While i <= 40
    desnivel_cambio(i) = 0
    i = i + 1
Wend
i = 0
While i <= 10
    aisl_n_var(i) = 0
    i = i + 1
Wend
i = 0
While i <= 10
    acum_aisl_var(i) = 0
    i = i + 1
Wend

aisl_sla = 0

'//
'//MEDICIONES PENDOLADO
'//

'Lg. Entraxe Conducteur TOTAL VÍA PRINCIPAL
i = 1
j = 13
While i <= plantilla.Sheets(1).Cells(9, 4)

    contador_pend_long_tot_VP = plantilla.Worksheets(1).Cells(j, 5) + contador_pend_long_tot_VP
    'Lg. Entraxe Cosse VÍA PRINCIPAL
    contador_pend_long_VP = plantilla.Worksheets(1).Cells(j, 6) + contador_pend_long_VP
    'Número péndolas VÍA PRINCIPAL
    i = i + 1
    j = j + 2

Wend
contador_pend_VP = plantilla.Worksheets(1).Cells(9, 4) + contador_pend_VP

Workbooks(1).Sheets("Material").Cells(6, 12) = contador_pend_long_tot_VP
Workbooks(1).Sheets("Material").Cells(6, 11) = contador_pend_long_VP
Workbooks(1).Sheets("Material").Cells(4, 11) = contador_pend_VP
'//
'//
'//

fin_anclaje_1hc:

plantilla.DisplayAlerts = False
plantilla.Workbooks.Close


'//
'//ENTRA SI HAY VANO DOBLE IT=1
'//

If it = 1 Then
    
    plantilla.Workbooks.Open "W:\223\D\D223041\IN_INFORMES\plantilla_pendolado.xlsm"
    
    'pk_ini_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) & "+" & (Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000)

        'If Round(Workbooks(1).Sheets("Replanteo").Cells(a, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000, 2) < 100 Then
            'ceros = "0"
            'If Round(Workbooks(1).Sheets("Replanteo").Cells(a, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000, 2) < 10 Then
            'ceros = "00"
            'End If
        'Else
            'ceros = ""
        'End If
        'pk_ini_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) & "+" & ceros & (Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3)) / 1000) * 1000)

    'pk_fin_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) & "+" & (Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000)

        'If Round(Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000, 2) < 100 Then
            'ceros = "0"
            'If Round(Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000, 2) < 10 Then
            'ceros = "00"
            'End If
        'Else
            'ceros = ""
        'End If
        'pk_fin_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) & "+" & ceros & (Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3))) - Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3)) / 1000) * 1000)
        
    pk_ini_var = Workbooks(1).Sheets("Replanteo").Cells(a, 3).text
    pk_fin_var = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3).text
    
    plantilla.Sheets(1).Name = pk_ini_var & " - " & pk_fin_var

    plantilla.Sheets(1).Cells(3, 7).Value = pk_ini_var & " - " & pk_fin_var

    plantilla.Sheets(1).Cells(4, 11) = codigo
    
    plantilla.Sheets(1).Cells(2, 5) = "LIGNE: " & nombre_tramo
        
    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 52) = "" Then

        va = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 4)
    
    Else

        va = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 53)
    
    End If
    plantilla.Sheets(1).Cells(4, 4) = va
    plantilla.Sheets(1).Cells(5, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 45)
    plantilla.Sheets(1).Cells(6, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 47)
    plantilla.Sheets(1).Cells(7, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 46)
    plantilla.Sheets(1).Cells(8, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 48)
    plantilla.Sheets(1).Cells(6, 11) = Workbooks(1).Sheets("Replanteo").Cells(a, 1)
    plantilla.Sheets(1).Cells(8, 11) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1)
    plantilla.Sheets(1).Cells(7, 5) = tip_pend_var
    el_hc_ini = plantilla.Sheets(1).Cells(7, 4)
    el_hc_fin = plantilla.Sheets(1).Cells(8, 4)
    
    it = 0
    st = 1
    
    If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 55 - it) <> "" Then
    
        dist_max_pend = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 55 - it)
        dist_prim_seg_pend_izq = 4.1
        dist_prim_seg_pend_der = 4.1
        
        'En seccionamiento eléctrico (en seccionamiento eléctrico el aisl irá a 0,75m de la primera péndola, al girarlo todo pq la elevación va a derecha para el cálculo, deberemos ponerlo a 0,75m de la última. (en doble hilo afcará a última y penúltima))
        If Workbooks(1).Sheets("Replanteo").Cells(a, 16) = "Inter.Section." Or Workbooks(1).Sheets("Replanteo").Cells(a + 2, 16) = "Inter.Section." Then
            p_aisl = 3.071 'daN
            aisl_sla = 1
            dist_aisl_1 = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 49) - 0.75
            dist_aisl_2 = va - Workbooks(1).Sheets("Replanteo").Cells(a + 1, 49) + 0.75
        End If
        
    End If
    
    
anclaje_2hc_top:
    
    '//
    '//GEOMETRÍA PARA PENDOLADO ANCLAJE
    '//
    
        If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 49) = "var" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 50) = "var" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 43) = "var" Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 44) = "var" Then
            
            If Workbooks(1).Sheets("Replanteo").Cells(a, 16) = "Anc.Section." Or Workbooks(1).Sheets("Replanteo").Cells(a + 2, 16) = "Anc.Section." Or Workbooks(1).Sheets("Replanteo").Cells(a, 16) = "Anc.Section.sans AT" Or Workbooks(1).Sheets("Replanteo").Cells(a + 2, 16) = "Anc.Section.sans AT" Then 'Anclaje de seccionamiento lleva cable de acero desde el semieje al anclaje
                
                anc_2hc_1hc = 1
                Call pendolado_1hc_VP.pendolado_1hc_VP(nombre_cat, ruta_replanteo, fila_ini, fila_fin)
                
            End If
            
            If anc_2hc_1hc = 3 Then
            
                GoTo fin_anc_1hc
            End If
            
            i = 1
            While i <= 10
                aisl_n(i) = 0
                i = i + 1
            Wend
                       
            If va / 5 > 9.6 And va / 5 < 16 Then
                If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 46) < Workbooks(1).Sheets("Replanteo").Cells(a + 1, 48) Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                    If va / 5 - Int(va / 5) < 0.00001 Then
                        
                        dist_ap_prim_pend_der = Int(va / 5) + 1
                        dist_max_pend = (va - dist_ap_prim_pend_der) / 4
                        Dist(1) = dist_max_pend
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_max_pend
                        Dist(5) = dist_ap_prim_pend_der
                        aisl_n(3) = 1
                        aisl_n(4) = 1
                        acum_aisl(4) = 4
                        
                     Else
                    
                        dist_ap_prim_pend_izq = Int(va / 5)
                        dist_max_pend = Int(va / 5)
                        dist_ap_prim_pend_der = va - dist_max_pend * 4
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_max_pend
                        Dist(5) = dist_ap_prim_pend_der
                        aisl_n(3) = 1
                        aisl_n(4) = 1
                        acum_aisl(4) = 4
                    
                    End If
                    
                Else
                
                    If va / 5 - Int(va / 5) < 0.00001 Then
                        
                        dist_ap_prim_pend_izq = Int(va / 5) + 1
                        dist_max_pend = (va - dist_ap_prim_pend_izq) / 4
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_max_pend
                        Dist(5) = dist_max_pend
                        aisl_n(3) = 1
                        aisl_n(4) = 1
                        acum_aisl(4) = 4
                        
                    Else
                
                        dist_ap_prim_pend_der = Int(va / 5)
                        dist_max_pend = Int(va / 5)
                        dist_ap_prim_pend_izq = va - dist_max_pend * 4
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_max_pend
                        Dist(5) = dist_ap_prim_pend_der
                        aisl_n(3) = 1
                        aisl_n(4) = 1
                        acum_aisl(4) = 4
                    
                    End If
                                       
                End If
                
                dist_prim_seg_pend_izq = 0
                dist_prim_seg_pend_der = 0
                p_aisl = 3.071 'daN
                
                GoTo ini3
                
            ElseIf va / 4 > 8 And va / 4 < 12 Then 'intervalo de 32 a 48
                If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 46) < Workbooks(1).Sheets("Replanteo").Cells(a + 1, 48) Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                    If va / 4 - Int(va / 4) < 0.00001 Then
                        
                        dist_ap_prim_pend_der = Int(va / 4) + 1
                        dist_max_pend = (va - dist_ap_prim_pend_der) / 3
                        Dist(1) = dist_max_pend
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_ap_prim_pend_der
                        aisl_n(2) = 1
                        aisl_n(3) = 1
                        acum_aisl(3) = 3
                        
                     Else
                        
                        dist_ap_prim_pend_izq = Int(va / 4)
                        dist_max_pend = Int(va / 4)
                        dist_ap_prim_pend_der = va - dist_max_pend * 3
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_ap_prim_pend_der
                        aisl_n(2) = 1
                        aisl_n(3) = 1
                        acum_aisl(3) = 3
                    
                    End If
                    
                Else
                    
                    If va / 4 - Int(va / 4) < 0.00001 Then
                        
                        dist_ap_prim_pend_izq = Int(va / 4) + 1
                        dist_max_pend = (va - dist_ap_prim_pend_izq) / 3
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_max_pend
                        aisl_n(2) = 1
                        aisl_n(3) = 1
                        acum_aisl(3) = 3
                     Else
                    
                        dist_ap_prim_pend_der = Int(va / 4)
                        dist_max_pend = Int(va / 4)
                        dist_ap_prim_pend_izq = va - dist_max_pend * 3
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        Dist(4) = dist_ap_prim_pend_der
                        aisl_n(2) = 1
                        aisl_n(3) = 1
                        acum_aisl(3) = 3
                        
                    End If
                
                End If
                
                dist_prim_seg_pend_izq = 0
                dist_prim_seg_pend_der = 0
                p_aisl = 3.071 'daN
                
                GoTo ini3
                
            ElseIf va / 3 > 2 And va / 3 < 12 Then
                If Workbooks(1).Sheets("Replanteo").Cells(a + 1, 46) < Workbooks(1).Sheets("Replanteo").Cells(a + 1, 48) Or Workbooks(1).Sheets("Replanteo").Cells(a + 1, 56) = "TOPF" Then
                    If va / 3 - Int(va / 3) < 0.00001 Then
                        
                        dist_ap_prim_pend_der = Int(va / 3) + 1
                        dist_max_pend = (va - dist_ap_prim_pend_der) / 2
                        Dist(1) = dist_max_pend
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_ap_prim_pend_der
                        aisl_n(1) = 1
                        aisl_n(2) = 1
                        acum_aisl(2) = 2
                        
                     Else
                    
                    
                        dist_ap_prim_pend_izq = Int(va / 3)
                        dist_max_pend = Int(va / 3)
                        dist_ap_prim_pend_der = va - dist_max_pend * 2
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_ap_prim_pend_der
                        aisl_n(1) = 1
                        aisl_n(2) = 1
                        acum_aisl(2) = 2
                    
                    End If
                Else
                    
                    If va / 3 - Int(va / 3) < 0.00001 Then
                        
                        dist_ap_prim_pend_izq = Int(va / 3) + 1
                        dist_max_pend = (va - dist_ap_prim_pend_izq) / 2
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_max_pend
                        aisl_n(1) = 1
                        aisl_n(2) = 1
                        acum_aisl(2) = 2
                        
                        
                     Else
                
                        dist_ap_prim_pend_der = Int(va / 3)
                        dist_max_pend = Int(va / 3)
                        dist_ap_prim_pend_izq = va - dist_max_pend * 2
                        Dist(1) = dist_ap_prim_pend_izq
                        Dist(2) = dist_max_pend
                        Dist(3) = dist_ap_prim_pend_der
                        aisl_n(1) = 1
                        aisl_n(2) = 1
                        acum_aisl(2) = 2
                    
                    End If
                    
                End If
                
                dist_prim_seg_pend_izq = 0
                dist_prim_seg_pend_der = 0
                p_aisl = 3.071 'daN
                
                GoTo ini3
                
             End If
             
    '//
    '//PARA PENDOLADO ANCLAJE, SECCIONAMIENTOS, AGUJAS.
    '//
        Else
        
            dist_ap_prim_pend_izq = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 49)
            dist_ap_prim_pend_der = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 50)
        
        End If
          
    GoTo ini2
    
Else
    st = 0
    it = 0
    a = a + 2
    GoTo ini1

End If
X:
    fso.DeleteFile ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf", True
    '//
    '// cerrar objectos
    '//
    mypdf.CancelJob
    Set mypdf = Nothing
    'fso.Close
    Set fso = Nothing

    plantilla.DisplayAlerts = False
    plantilla.Workbooks.Close
    plantilla.Quit
    Set plantilla = Nothing
   
final:

    plantilla.Workbooks.Close
    plantilla.Quit
    '//
    '//ESCRITURA PIE DE HOJA CONTROL
    '//
    'plantilla_control.Sheets(1).Cells(contador_hojas + 8, 1) = "LEGENDE"
    'plantilla_control.Sheets(1).Cells(contador_hojas + 8, 2) = "+ FEUILLE CRÉE"
    'plantilla_control.Sheets(1).Cells(contador_hojas + 9, 2) = "- FEUILLE SUPPRIMÉE"
    'plantilla_control.Sheets(1).Cells(contador_hojas + 10, 2) = "'= FEUILLE MODIFIÉE"
    'plantilla_control.Sheets(1).Cells(contador_hojas + 11, 2) = "0 FEUILLE INCHANGÉE"
    plantilla_control.Sheets(1).PageSetup.RightFooter = codigo
    '//
    '//
    '//
    plantilla_control.DisplayAlerts = False
    Call plantilla_control.ActiveWorkbook.Close(True, "C:\Users\23370\Desktop\D50\" & nombre_tramo)
    'plantilla_control.SaveAs ("C:\Users\23370\Desktop\D50\" & nombre_tramo)
    'Workbooks("plantilla_control.xlsm").SaveAs ("C:\Users\23370\Desktop\D50\" & nombre_tramo)
    plantilla_control.Workbooks.Close
    plantilla_control.Quit
    Set plantilla = Nothing
    Set plantilla_control = Nothing
  
fin_anc_1hc:
  
End Sub
Sub pendolado_MT(nombre_catVB, ruta_replanteoVB)
'//
'//INSERTA DIRECTAMENTE FICHAS YA CREADAS AÑADIENDO ALGUNOS DATOS DEL REPLANTEO
'//

Dim fso As Object
Dim Dist(100) As Double, fuerza(40) As Double, mom(40) As Double
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
Dim tip As String, tip_var As String, PSFileName As String
Dim PDFFileName As String, TXTFileName As String
Dim pk_ini_var As String, pk_fin_var As String, tip_pend As String, tip_pend_var As String


Dim mypdf As PdfDistiller
Set mypdf = New PdfDistiller
Set fso = CreateObject("Scripting.FileSystemObject")

Dim plantilla As Object
Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False

a = 10
dfijo = Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value
efijo = Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 1).Value
PDFFijoFileName = ruta_replanteoVB & "\" & dfijo & " " & efijo & ".pdf"

ini4:

If Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1) = "" Then

    GoTo final:

End If

tip_pend = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 11)
tip_pend_var = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 12)

If tip_pend <> "" And tip_pend_var <> "" Then

    it = 1
    
Else: it = 0
        
End If

ini3:

If (fso.FileExists("W:\210\P\P210D50\IN_INFORMES\8-Mission 3 - Études d'exécution caténaire\CATENARIA 3.000 Vcc\PENDOLADO\TODOS\" & tip_pend & ".xlsx")) Then

Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False
plantilla.Workbooks.Open "W:\210\P\P210D50\IN_INFORMES\8-Mission 3 - Études d'exécution caténaire\CATENARIA 3.000 Vcc\PENDOLADO\TODOS\" & tip_pend & ".xlsx"

Else

Set plantilla = CreateObject("Excel.Application")
plantilla.Visible = False
plantilla.Workbooks.Open "W:\210\P\P210D50\IN_INFORMES\8-Mission 3 - Études d'exécution caténaire\CATENARIA 3.000 Vcc\PENDOLADO\TODOS\non_defini.xlsx"

plantilla.Sheets(1).Cells(4, 4) = Workbooks(1).Sheets("Replanteo").Cells(a + 1, 4)

End If

pk_ini_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a, 3) + 1) / 1000) & "+" & Int(Workbooks(1).Sheets("Replanteo").Cells(a, 3) + 1)
pk_fin_var = Int((Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3) + 1) / 1000) & "+" & Int(Workbooks(1).Sheets("Replanteo").Cells(a + 2, 3) + 1)

plantilla.Sheets(1).Name = pk_ini_var & " - " & pk_fin_var

plantilla.Sheets(1).Cells(6, 10) = Workbooks(1).Sheets("Replanteo").Cells(a, 1)
plantilla.Sheets(1).Cells(8, 10) = Workbooks(1).Sheets("Replanteo").Cells(a + 2, 1)
plantilla.Sheets(1).Cells(7, 5) = tip_pend

'//
'//GENERACIÓN FICHA
'//
b = 1

    If it = 1 Or it = 0 Then
    
        d = Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value
        e = Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 1).Value
    
    Else: it = 2
    
        d = Workbooks(1).Worksheets("Replanteo").Cells(a, 1).Value & "_1"
        e = Workbooks(1).Worksheets("Replanteo").Cells(a + 2, 1).Value & "_1"
        
    End If


    'fncScreenUpdating State:=False
    Call plantilla.Worksheets(b).PrintOut(from:=1, To:=1, Copies:=1, preview:=False, ActivePrinter:="Adobe PDF", printtofile:=True, collate:=False, prtofilename:=ruta_replanteoVB & "\" & d & " " & e & ".ps")
    'fncScreenUpdating State:=True
    PSFileName = ruta_replanteoVB & "\" & d & " " & e & ".ps"
    PDFFileName = ruta_replanteoVB & "\" & d & " " & e & ".pdf"
    TXTFileName = ruta_replanteoVB & "\" & d & " " & e & ".log"
    mypdf.FileToPDF PSFileName, PDFFileName, ""
    fso.DeleteFile PSFileName, True
    fso.DeleteFile TXTFileName, True
'//
'//INSERCIÓN FICHA EN PDF GLOBAL
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
    mypdf.CancelJob
    Set mypdf = Nothing
    'fso.Close
    Set fso = Nothing
    
End Sub


Function CombPDF(PDFfijo, PDFName, ruta_replanteoVB)
'Dim fso As Object
Dim AcroApp As Acrobat.CAcroApp
Dim Part1Document As Acrobat.CAcroPDDoc
Dim Part2Document As Acrobat.CAcroPDDoc
'Dim numPages As Integer
Set AcroApp = CreateObject("AcroExch.App")
Set Part1Document = CreateObject("AcroExch.PDDoc")
Set Part2Document = CreateObject("AcroExch.PDDoc")
Set fso = CreateObject("Scripting.FileSystemObject")
Part1Document.Open (PDFfijo)
Part2Document.Open (PDFName)
    
numpages = Part1Document.GetNumPages()
    
If Part1Document.InsertPages(numpages - 1, Part2Document, _
0, Part2Document.GetNumPages(), True) = False Then
    Exit Function
End If
If Part1Document.Save(PDSaveFull, ruta_replanteoVB & "\pendulage.pdf") = False Then
Else
        PDFfijo = ruta_replanteoVB & "\pendulage.pdf"
        
End If
        
Part1Document.Close
Part2Document.Close
AcroApp.Exit
Set AcroApp = Nothing
Set Part1Document = Nothing
Set Part2Document = Nothing
fso.DeleteFile PDFName, True
End Function

Sub pendolado_columna(nombre_catVB)
Dim z As Integer, cont As Integer, z1 As Integer
Dim longueur As Double
Dim primero As String, primero2 As String, segundo As String, tercero As String, cuatro As String, segundo2 As String, tercero2 As String, cuarto As String, cuarto2 As String
Dim normal As Boolean, cas_tun As Boolean
Dim algo As Variant
dim tip_1 as String, tip_pf_1 as string, tip_0 as String, tip_pf_0 as string, tip_2 as String, tip_pf_2 as string
z = 10
Call cargar.datos_lac(nombre_catVB)
While Not IsEmpty(Sheets("Replanteo").Cells(z, 33).Value)
If z = 3994 Then

    algo = 0
End If
    primero = ""
    primero2 = ""
    segundo = ""
    segundo2 = ""
    tercero = ""
    tercero2 = ""
    cuarto = ""
    cuarto2 = ""
        If Sheets("Replanteo").Cells(z, 16).Value = anc_sla_con & " + " & semi_eje_aguj Then
            tip_1 = semi_eje_aguj
            tip_pf_1 = anc_sla_con
        ElseIf Sheets("Replanteo").Cells(z, 16).Value = semi_eje_sla & " + " & anc_aguj Then
            tip_1 = semi_eje_sla
            tip_pf_1 = anc_aguj
        ElseIf Len(Sheets("Replanteo").Cells(z, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(z, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(z, 16).Value = anc_sm_sin) Then
            tip_1 = Mid(Sheets("Replanteo").Cells(z, 16).Value, 15)
            tip_pf_1 = Mid(Sheets("Replanteo").Cells(z, 16).Value, 1, 11)
        Else
            tip_1 = Sheets("Replanteo").Cells(z, 16).Value
            tip_pf_1 = Sheets("Replanteo").Cells(z, 16).Value
        End If
        
        If Sheets("Replanteo").Cells(z - 2, 16).Value = semi_eje_sla & " + " & anc_aguj Then
            tip_0 = anc_aguj
            tip_pf_0 = semi_eje_sla
            
        ElseIf Len(Sheets("Replanteo").Cells(z - 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(z - 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(z - 2, 16).Value = anc_sm_sin) Then
            tip_0 = Mid(Sheets("Replanteo").Cells(z - 2, 16).Value, 15)
            tip_pf_0 = Mid(Sheets("Replanteo").Cells(z - 2, 16).Value, 1, 11)
        Else
            tip_0 = Sheets("Replanteo").Cells(z - 2, 16).Value
            tip_pf_0 = Sheets("Replanteo").Cells(z - 2, 16).Value
        End If
        If Sheets("Replanteo").Cells(z + 2, 16).Value = anc_sla_con & " + " & semi_eje_aguj Then
            tip_2 = anc_sla_con
            tip_pf_2 = semi_eje_aguj
        ElseIf Sheets("Replanteo").Cells(z + 2, 16).Value = semi_eje_sla & " + " & anc_aguj Then
            tip_2 = semi_eje_sla
            tip_pf_2 = anc_aguj
        ElseIf Len(Sheets("Replanteo").Cells(z + 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(z + 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(z + 2, 16).Value = anc_sm_sin) Then
            tip_2 = Mid(Sheets("Replanteo").Cells(z + 2, 16).Value, 15)
            tip_pf_2 = Mid(Sheets("Replanteo").Cells(z + 2, 16).Value, 1, 11)
        Else
            tip_2 = Sheets("Replanteo").Cells(z + 2, 16).Value
            tip_pf_2 = Sheets("Replanteo").Cells(z + 2, 16).Value
        End If
    
    
    
    
    '///
    '/// Verificar si el cantón actual tiene o no compensación mecánica
    '///
    
    If tip_pf_1 = eje_pf Then
        z1 = z + 2
            While IsEmpty(Sheets("Replanteo").Cells(z1, 16).Value)
                z1 = z1 + 2
            Wend
        
        If Sheets("Replanteo").Cells(z1, 16).Value = anc_sm_sin Then
            cas_tun = True
        Else
            cas_tun = False
        End If
    
    End If

    '///
    '///normal o inverso
    '///
    If (tip_1 = anc_sm_con Or tip_1 = anc_sla_con) Then
        If Sheets("Replanteo").Cells(z, 47).Value <> "Normal" Then
            normal = False
        Else
            normal = True
        End If
    End If
If Sheets("Replanteo").Cells(z, 38).Value = "Marquesina" And Sheets("Replanteo").Cells(z + 2, 38).Value = "Marquesina" Then
Else
    '///
    '///primera letra
    '///
   

    
    If (IsEmpty(Sheets("Replanteo").Cells(z, 16).Value) Or ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_0 = semi_eje_sm) Or _
    ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_0 = semi_eje_sla) Or tip_1 = anc_pf Or tip_1 = eje_pf) _
    Or (tip_1 = anc_aguj And tip_0 = semi_eje_aguj) Or (tip_1 = eje_aguj And tip_0 = semi_eje_aguj) Then

        primero = "C"
    Else
        If cas_tun = True Then
            primero = "T"
            primero2 = "T"
        
        ElseIf (tip_1 = anc_aguj And tip_2 = semi_eje_aguj) Or (tip_1 = semi_eje_aguj And tip_2 = anc_aguj) Or _
        (tip_1 = semi_eje_aguj And (tip_2 = eje_aguj Or Sheets("Replanteo").Cells(z + 2, 16).Value = anc_pf & " + " & semi_eje_aguj)) Or (tip_1 = eje_aguj And tip_2 = semi_eje_aguj) Or _
        (tip_1 = eje_aguj And tip_2 <> semi_eje_aguj) Or (tip_1 = semi_eje_aguj And tip_2 = semi_eje_aguj) Then
            primero = "C"
            primero2 = "S"
        Else
            primero = "S"
            primero2 = "S"
        End If
    End If
    '///
    '///segunda letra
    '///
    
    If IsEmpty(Sheets("Replanteo").Cells(z, 16).Value) Or ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_0 = semi_eje_sm) Or _
    ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_0 = semi_eje_sla) Or tip_1 = anc_pf Or tip_1 = eje_pf Or _
    (tip_1 = eje_aguj And (tip_2 <> semi_eje_aguj)) Or _
    (tip_1 = anc_aguj And tip_2 <> semi_eje_aguj) Then 'Or Sheets("Replanteo").Cells(z, 16).Value = anc_pf & " + " & eje_aguj Or Sheets("Replanteo").Cells(z, 16).Value = eje_pf & " + " & eje_aguj) And (Sheets("Replanteo").Cells(z + 2, 16).Value <> semi_eje_aguj And Sheets("Replanteo").Cells(z + 2, 16).Value <> eje_pf & " + " & semi_eje_aguj And Sheets("Replanteo").Cells(z + 2, 16).Value <> anc_pf & " + " & semi_eje_aguj)) Then
        If cas_tun = True Then
            segundo = "T"
        Else
            segundo = "S"
        End If
    ElseIf (tip_1 = anc_aguj And tip_2 = semi_eje_aguj) Then
        segundo = "S"
        If Sheets("Replanteo").Cells(z + 3, 4).Value >= 40.5 Then
            segundo2 = 1
        ElseIf Sheets("Replanteo").Cells(z + 3, 4).Value > 31.5 Then
            segundo2 = 3
        ElseIf Sheets("Replanteo").Cells(z + 3, 4).Value <= 31.5 Then
            segundo2 = 1
        End If
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = anc_aguj) Then
            segundo = "S"
        If Sheets("Replanteo").Cells(z - 3, 4).Value >= 40.5 Then
            segundo2 = 1
        ElseIf Sheets("Replanteo").Cells(z - 3, 4).Value > 31.5 Then
            segundo2 = 3
        ElseIf Sheets("Replanteo").Cells(z - 3, 4).Value <= 31.5 Then
            segundo2 = 1
        End If
    
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = eje_aguj) Or (tip_1 = eje_aguj And tip_2 = semi_eje_aguj) Then
            segundo = "S"
            segundo2 = "P01-"
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = semi_eje_aguj) Or (tip_1 = semi_eje_aguj And tip_2 = semi_eje_aguj) Then
            segundo = "S"
            segundo2 = "P00-"
    ElseIf ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_2 = semi_eje_sm) Or _
    ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_2 = semi_eje_sla) Or _
    (tip_1 = anc_aguj And tip_2 = semi_eje_aguj) Then
        segundo = "K"
        If Sheets("Replanteo").Cells(z + 3, 4).Value >= 40.5 Then
            segundo2 = 1
        ElseIf Sheets("Replanteo").Cells(z + 3, 4).Value >= 31.5 And ((tip_1 = anc_sla_con And tip_2 = semi_eje_sla) Or _
        (tip_1 = anc_sla_sin And tip_2 = semi_eje_sla)) Then
            segundo2 = 2
        ElseIf Sheets("Replanteo").Cells(z + 3, 4).Value < 31.5 And ((tip_1 = anc_sla_con And tip_2 = semi_eje_sla) Or _
        (tip_1 = anc_sla_sin And tip_2 = semi_eje_sla)) Then
            segundo2 = 4
        ElseIf Sheets("Replanteo").Cells(z + 3, 4).Value >= 31.5 And ((tip_1 = anc_sm_con And tip_2 = semi_eje_sm) Or _
        (tip_1 = anc_sm_sin And tip_2 = semi_eje_sm)) Then
            segundo2 = 3
        ElseIf Sheets("Replanteo").Cells(z + 3, 4).Value < 31.5 And ((tip_1 = anc_sm_con And tip_2 = semi_eje_sm) Or _
        (tip_1 = anc_sm_sin And tip_2 = semi_eje_sm)) Then
            segundo2 = 4
        End If
    ElseIf (tip_1 = semi_eje_sm And (tip_2 = anc_sm_con Or tip_2 = anc_sm_sin)) Or _
    (tip_1 = semi_eje_sla And (tip_2 = anc_sla_con Or tip_2 = anc_sla_sin)) _
    Or (tip_1 = semi_eje_sm And tip_2 = anc_sm_sin) Or (tip_1 = semi_eje_aguj And tip_2 = anc_aguj) Then
        segundo = "K"
        If Sheets("Replanteo").Cells(z - 1, 4).Value >= 40.5 Then
            segundo2 = 1
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value >= 31.5 And ((tip_1 = semi_eje_sla And tip_2 = anc_sla_con) Or _
        (tip_1 = semi_eje_sla And tip_2 = anc_sla_sin)) Then
            segundo2 = 2
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value < 31.5 And ((tip_1 = semi_eje_sla And tip_2 = anc_sla_con) Or _
        (tip_1 = semi_eje_sla And tip_2 = anc_sla_sin)) Then
            segundo2 = 4
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value >= 31.5 And ((tip_1 = semi_eje_sm And tip_2 = anc_sm_con) Or _
        (tip_1 = semi_eje_sm And tip_2 = anc_sm_sin)) Then
            segundo2 = 3
        ElseIf Sheets("Replanteo").Cells(z - 1, 4).Value < 31.5 And ((tip_1 = semi_eje_sm And tip_2 = anc_sm_con) Or _
        (tip_1 = semi_eje_sm And tip_2 = anc_sm_sin)) Then
            segundo2 = 4
        End If
                
    ElseIf (tip_1 = semi_eje_sm And tip_2 = eje_sm) Then
        segundo = "K"
        segundo2 = "C"
    ElseIf tip_1 = eje_sm And tip_2 = eje_sm Then
        segundo = "C"
        segundo2 = "C"
    ElseIf tip_1 = eje_sla And tip_2 = eje_sla Then
        segundo = "S"
        segundo2 = "S"
    ElseIf tip_1 = eje_sm Then
        segundo = "C"
        segundo2 = "K"
    ElseIf (tip_1 = semi_eje_sla And tip_2 = eje_sla) Then
        segundo = "K"
        segundo2 = "S"
    ElseIf tip_1 = eje_sla And tip_2 = semi_eje_sla Then
        segundo = "S"
        segundo2 = "K"
    ElseIf tip_1 = eje_sla And tip_2 = eje_sla Then
        segundo = "S"
        segundo2 = "S"
    ElseIf (tip_0 = anc_sm_con Or tip_0 = anc_sm_sin) And tip_1 = semi_eje_sm And tip_2 = semi_eje_sm And normal = True Then
        segundo = "C"
        segundo2 = "C"
    ElseIf (tip_0 = anc_sm_con Or tip_0 = anc_sm_sin) And tip_1 = semi_eje_sm And tip_2 = semi_eje_sm And normal = False Then
        segundo = "C"
        segundo2 = "C"
    Else
        algo = 0
    End If
    '///
    '///tercera letra
    '///
    If IsEmpty(Sheets("Replanteo").Cells(z, 16).Value) Or ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_0 = semi_eje_sm) Or _
    ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_0 = semi_eje_sla) Or tip_1 = anc_pf Or tip_1 = eje_pf Or _
    (tip_1 = anc_aguj And tip_2 = semi_eje_aguj) Or (tip_1 = semi_eje_aguj And tip_2 = anc_aguj) Or _
    (tip_1 = eje_aguj And tip_2 <> semi_eje_aguj) Or (tip_1 = anc_aguj And tip_2 <> semi_eje_aguj) Or _
    (tip_1 = semi_eje_aguj And tip_2 = eje_aguj) Or (tip_1 = eje_aguj And tip_2 = semi_eje_aguj) Or _
    (tip_1 = semi_eje_aguj And tip_2 = semi_eje_aguj) Then
    
        tercero = "n"
    ElseIf ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_2 = semi_eje_sm) Or (tip_1 = semi_eje_sm Or tip_1 = eje_sm) _
    Or ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_2 = semi_eje_sla) Or (tip_1 = semi_eje_sla Or tip_1 = eje_sla) _
    Or (tip_1 = semi_eje_sm And tip_2 = anc_sm_sin) Then
            cont = 1
            longueur = 63 '// debe venir de una variable
            While cont <= 10 And Sheets("Replanteo").Cells(z + 1, 4).Value <> longueur
                longueur = longueur - 4.5 'inc_norm_va
                cont = cont + 1
            Wend
            If cont = 11 Then
                tercero = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
                tercero2 = tercero
            Else
                tercero = cont
                tercero2 = tercero
            End If

    End If
    If (tip_1 = anc_sla_con And tip_2 = semi_eje_sla) Or (tip_1 = semi_eje_sla And tip_2 = anc_sla_con) _
    Or (tip_1 = anc_sm_con And tip_2 = semi_eje_sm) Or (tip_1 = semi_eje_sm And tip_2 = anc_sm_con) _
    Or (tip_1 = anc_aguj And tip_2 = semi_eje_aguj) Or (tip_1 = semi_eje_aguj And tip_2 = anc_aguj) Then
        tercero2 = "A"
    ElseIf (tip_1 = anc_sla_sin And tip_2 = semi_eje_sla) Or (tip_1 = semi_eje_sla And tip_2 = anc_sla_sin) _
    Or (tip_1 = anc_sm_sin And tip_2 = semi_eje_sm) Or (tip_1 = semi_eje_sm And tip_2 = anc_sm_sin) Then
        tercero2 = "A"
    End If
    '///
    '///cuarta letra
    '///
    If IsEmpty(Sheets("Replanteo").Cells(z, 16).Value) Or ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_0 = semi_eje_sm) Or _
    ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_0 = semi_eje_sla) Or tip_1 = anc_pf Or tip_1 = eje_pf Or _
    (tip_1 = eje_aguj And tip_2 <> semi_eje_aguj) Or _
    (tip_1 = anc_aguj And tip_2 <> semi_eje_aguj) Then
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
'///
'/// Vano semieje aguja - eje aguja
'///
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = eje_aguj) Then
        Call sea_ea(z)
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
'///
'/// Vano eje aguja - semieje aguja
'///
    ElseIf (tip_1 = eje_aguj And tip_2 = semi_eje_aguj) Then
        Call ea_sea(z)
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
'///
'/// Vano anclaje aguja - semieje aguja
'///
    ElseIf (tip_1 = anc_aguj And tip_2 = semi_eje_aguj) Then
        Call aa_sea(z, segundo2)
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)

'///
'/// Vano semieje aguja - anclaje aguja
'///
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = anc_aguj) Then
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 40).Value = 0
        Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 42).Value = 0
        Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 44).Value = 1.125
        Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 49).Value = "var"
        Sheets("Replanteo").Cells(z + 1, 50).Value = "var"
        Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z, 39).Value = 1.4
        Sheets("Replanteo").Cells(z, 40).Value = 0
        If segundo2 = 1 Then
            Sheets("Replanteo").Cells(z, 45).Value = 1.8
            Sheets("Replanteo").Cells(z, 46).Value = 0.5
                If Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                    Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
                    Sheets("Replanteo").Cells(z + 1, 48).Value = 0.8
                    Sheets("Replanteo").Cells(z + 1, 47).Value = 0.8 + 0.5
                Else
                    Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
                    Sheets("Replanteo").Cells(z + 1, 48).Value = 0.65
                    Sheets("Replanteo").Cells(z + 1, 47).Value = 0.65 + 0.5
                End If
        ElseIf segundo2 = 3 Then
            Sheets("Replanteo").Cells(z, 45).Value = 1.8
            Sheets("Replanteo").Cells(z, 46).Value = 0.3
                If Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                    Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
                    Sheets("Replanteo").Cells(z + 1, 48).Value = 0.6
                    Sheets("Replanteo").Cells(z + 1, 47).Value = 0.6 + 0.5
                ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value < 31.5 Then
                    Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
                    Sheets("Replanteo").Cells(z + 1, 48).Value = 0.45
                    Sheets("Replanteo").Cells(z + 1, 47).Value = 0.45 + 0.5
                End If
        End If
'///
'/// Vano semieje aguja - semieje aguja
'///
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = semi_eje_aguj) And _
    tip_0 = anc_aguj Then
        Call sea_sea(z)
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)

'///
'/// Vano semieje aguja - semieje aguja
'///
    ElseIf (tip_1 = semi_eje_aguj And tip_2 = semi_eje_aguj) And _
    tip_0 = eje_aguj Then
        cuarto = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
        Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 40).Value = 0
        Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 42).Value = 0
        Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 47).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
        Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 48).Value = 0.5
        Sheets("Replanteo").Cells(z, 39).Value = 1.4
        Sheets("Replanteo").Cells(z, 40).Value = 0
        Sheets("Replanteo").Cells(z, 45).Value = 1.8
        Sheets("Replanteo").Cells(z, 46).Value = 0.3
'///
'/// Vano anclaje sm/sla - semieje sm/sla
'///
    ElseIf ((tip_1 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_2 = semi_eje_sm) Or ((tip_1 = anc_sla_con Or tip_1 = anc_sla_sin) And tip_2 = semi_eje_sla) Then
            cuarto = "a"
            cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 42).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 1.125
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 47).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 49).Value = "var"
            Sheets("Replanteo").Cells(z + 1, 50).Value = "var"
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.5
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
            If segundo2 = 1 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.5
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.8
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.8 + 0.5
            ElseIf segundo2 = 2 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.35
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.65
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.65 + 0.5
            ElseIf segundo2 = 3 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.3
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.6
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.6 + 0.5
            ElseIf segundo2 = 4 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.2
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.35
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.35 + 0.5
            End If
'///
'/// Vano semieje sm/sla - anclaje sm/sla
'///
    ElseIf ((tip_2 = anc_sm_con Or tip_1 = anc_sm_sin) And tip_1 = semi_eje_sm) Or ((tip_2 = anc_sla_con Or tip_2 = anc_sla_sin) And tip_1 = semi_eje_sla) _
    Or (tip_1 = semi_eje_sm And tip_2 = anc_sm_sin) Then
            cuarto = "b"
            cuarto2 = Round(Sheets("Replanteo").Cells(z + 1, 4).Value, 2)
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 42).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 1.125
            Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 49).Value = "var"
            Sheets("Replanteo").Cells(z + 1, 50).Value = "var"
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.5
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
            If segundo2 = 1 Then
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.8
                Sheets("Replanteo").Cells(z + 1, 47).Value = 0.8 + 0.5
            ElseIf segundo2 = 2 Then
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.35
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.65
                Sheets("Replanteo").Cells(z + 1, 47).Value = 0.65 + 0.5
            ElseIf segundo2 = 3 Then
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.6
                Sheets("Replanteo").Cells(z + 1, 47).Value = 0.6 + 0.5
            ElseIf segundo2 = 4 Then
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.2
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.35
                Sheets("Replanteo").Cells(z + 1, 47).Value = 0.35 + 0.5
            End If
'///
'/// Vano semieje sm/sla - eje sm/sla
'///
    ElseIf (tip_1 = semi_eje_sm And tip_2 = eje_sm) Or (tip_1 = semi_eje_sla And tip_2 = eje_sla) Then
        If normal = True Then
            Sheets("Replanteo").Cells(z, 39).Value = 1.4
            Sheets("Replanteo").Cells(z, 40).Value = 0
            Sheets("Replanteo").Cells(z, 45).Value = 1.8
            cuarto = "e"
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 42).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 47).Value = 2
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0
            Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3

            If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
                cuarto2 = "g"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
                Sheets("Replanteo").Cells(z, 45).Value = 1.8
                Sheets("Replanteo").Cells(z, 46).Value = 0.5
            ElseIf segundo2 = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto2 = "g1"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
                Sheets("Replanteo").Cells(z, 46).Value = 0.3
            ElseIf segundo2 = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto2 = "g2"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.2
                Sheets("Replanteo").Cells(z, 46).Value = 0.2
            ElseIf segundo2 = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto2 = "g1"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.35
                Sheets("Replanteo").Cells(z, 46).Value = 0.35
            ElseIf segundo2 = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto2 = "g2"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.2
                Sheets("Replanteo").Cells(z, 46).Value = 0.2
            End If
        Else
            cuarto = "k"
            Sheets("Replanteo").Cells(z, 39).Value = 1.4
            Sheets("Replanteo").Cells(z, 40).Value = 0
            Sheets("Replanteo").Cells(z, 45).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 2
            Sheets("Replanteo").Cells(z + 1, 42).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 47).Value = 1.3
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0
            Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3
            If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
                cuarto2 = "i"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
                Sheets("Replanteo").Cells(z, 46).Value = 0.5
            ElseIf segundo2 = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto2 = "i1"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
                Sheets("Replanteo").Cells(z, 46).Value = 0.3
            ElseIf segundo2 = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto2 = "i2"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.2
                Sheets("Replanteo").Cells(z, 46).Value = 0.2
            ElseIf segundo2 = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto2 = "i1"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.35
                Sheets("Replanteo").Cells(z, 46).Value = 0.35
            ElseIf segundo2 = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto2 = "i2"
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.2
                Sheets("Replanteo").Cells(z, 46).Value = 0.2
            End If
        End If
'///
'/// Vano eje sm/sla - semieje sm/sla
'///
    ElseIf (tip_1 = eje_sm And tip_2 = semi_eje_sm) Or (tip_1 = eje_sla And tip_2 = semi_eje_sla) Then
        If normal = True Then
            Sheets("Replanteo").Cells(z, 39).Value = 1.3
            Sheets("Replanteo").Cells(z, 40).Value = 0
            cuarto2 = "h"
            Sheets("Replanteo").Cells(z, 45).Value = 2
            Sheets("Replanteo").Cells(z, 46).Value = 0
            Sheets("Replanteo").Cells(z + 2, 39).Value = 1.8
            Sheets("Replanteo").Cells(z + 2, 40).Value = 0
            Sheets("Replanteo").Cells(z + 2, 45).Value = 1.4
            Sheets("Replanteo").Cells(z + 2, 46).Value = 0
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.3
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 45).Value = 2
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0
            Sheets("Replanteo").Cells(z + 1, 47).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3

            If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
                cuarto = "f"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.5
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.5
            ElseIf segundo = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto = "f1"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.3
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.3
            ElseIf segundo = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto = "f2"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.2
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.2
            ElseIf segundo = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto = "f1"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.35
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.35
            ElseIf segundo = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto = "f2"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.2
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.2
            End If
        Else
            Sheets("Replanteo").Cells(z, 39).Value = 2
            cuarto2 = "j"
            Sheets("Replanteo").Cells(z, 45).Value = 1.3
            Sheets("Replanteo").Cells(z, 46).Value = 0
            Sheets("Replanteo").Cells(z + 2, 45).Value = 1.4
            Sheets("Replanteo").Cells(z + 2, 46).Value = 0
            Sheets("Replanteo").Cells(z + 2, 39).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 39).Value = 2
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 45).Value = 1.3
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0
            Sheets("Replanteo").Cells(z + 1, 47).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3

            If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
                cuarto = "l"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.5
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.5
            ElseIf segundo = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto = "l1"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.3
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.3
            ElseIf segundo = "C" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto = "l2"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.2
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.2
            ElseIf segundo = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                cuarto = "l1"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.35
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.35
            ElseIf segundo = "S" And Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                cuarto = "l2"
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0.2
                Sheets("Replanteo").Cells(z + 2, 40).Value = 0.2
            End If
        End If
'///
'/// Vano eje sla - eje sla
'///
     ElseIf (tip_1 = eje_sla And tip_2 = eje_sla) Or (tip_1 = eje_sm And tip_2 = eje_sm) Then
            If normal = True Then
                'segundo = "K"
                'segundo2 = "K"
                Sheets("Replanteo").Cells(z, 39).Value = 1.3
                Sheets("Replanteo").Cells(z, 40).Value = 0
                Sheets("Replanteo").Cells(z, 45).Value = 2
                Sheets("Replanteo").Cells(z, 46).Value = 0
                If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
                    cuarto2 = "y"
                    cuarto = "z"
                ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                    cuarto2 = "y1"
                    cuarto = "z1"
                ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                    cuarto2 = "y2"
                    cuarto = "z2"
                End If
                Sheets("Replanteo").Cells(z + 1, 39).Value = 1.3
                Sheets("Replanteo").Cells(z + 1, 40).Value = 0
                Sheets("Replanteo").Cells(z + 1, 41).Value = 1.3
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0
                Sheets("Replanteo").Cells(z + 1, 43).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 44).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 45).Value = 2
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0
                Sheets("Replanteo").Cells(z + 1, 47).Value = 2
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0
                Sheets("Replanteo").Cells(z + 1, 49).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 50).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
                Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
                
            Else
                'segundo = "K"
                'segundo2 = "K"
                Sheets("Replanteo").Cells(z, 39).Value = 2
                Sheets("Replanteo").Cells(z, 40).Value = 0
                Sheets("Replanteo").Cells(z, 45).Value = 1.3
                Sheets("Replanteo").Cells(z, 46).Value = 0
                'Sheets("Replanteo").Cells(z + 1, 39).Value = 1.3
                'Sheets("Replanteo").Cells(z + 1, 45).Value = 2
                If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
                    cuarto2 = "v"
                    cuarto = "w"
                ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
                    cuarto2 = "v1"
                    cuarto = "w1"
                ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 22.5 Then
                    cuarto2 = "v2"
                    cuarto = "w2"
                End If
                Sheets("Replanteo").Cells(z + 1, 39).Value = 2
                Sheets("Replanteo").Cells(z + 1, 40).Value = 0
                Sheets("Replanteo").Cells(z + 1, 41).Value = 2
                Sheets("Replanteo").Cells(z + 1, 42).Value = 0
                Sheets("Replanteo").Cells(z + 1, 43).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 44).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 45).Value = 1.3
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0
                Sheets("Replanteo").Cells(z + 1, 47).Value = 1.3
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0
                Sheets("Replanteo").Cells(z + 1, 49).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 50).Value = 1.125
                Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
                Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
                
            End If
    ElseIf (tip_0 = anc_sm_con Or tip_0 = anc_sm_sin) And tip_1 = semi_eje_sm And tip_2 = semi_eje_sm And normal = True Then
            Sheets("Replanteo").Cells(z, 39).Value = 1.4
            Sheets("Replanteo").Cells(z, 40).Value = 0
            cuarto2 = "m"
            Sheets("Replanteo").Cells(z, 45).Value = 1.8
            Sheets("Replanteo").Cells(z, 46).Value = 0.3
            Sheets("Replanteo").Cells(z + 2, 39).Value = 1.8
            Sheets("Replanteo").Cells(z + 2, 40).Value = 0.3
            Sheets("Replanteo").Cells(z + 2, 45).Value = 1.4
            Sheets("Replanteo").Cells(z + 2, 46).Value = 0
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
            Sheets("Replanteo").Cells(z + 1, 47).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3
            cuarto = "n"
            Sheets("Replanteo").Cells(z + 1, 42).Value = 0.3
    ElseIf (tip_0 = anc_sm_con Or tip_0 = anc_sm_sin) And tip_1 = semi_eje_sm And tip_2 = semi_eje_sm And normal = False Then
            Sheets("Replanteo").Cells(z, 39).Value = 1.4
            Sheets("Replanteo").Cells(z, 40).Value = 0
            cuarto2 = "n"
            Sheets("Replanteo").Cells(z, 45).Value = 1.8
            Sheets("Replanteo").Cells(z, 46).Value = 0.3
            Sheets("Replanteo").Cells(z + 2, 39).Value = 1.8
            Sheets("Replanteo").Cells(z + 2, 40).Value = 0.3
            Sheets("Replanteo").Cells(z + 2, 45).Value = 1.4
            Sheets("Replanteo").Cells(z + 2, 46).Value = 0
            Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 40).Value = 0
            Sheets("Replanteo").Cells(z + 1, 41).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
            Sheets("Replanteo").Cells(z + 1, 47).Value = 1.4
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0
            Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
            Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3
            Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3
            cuarto = "m"
            Sheets("Replanteo").Cells(z + 1, 42).Value = 0.3
    
    Else
    algo = 0
    End If
End If

Sheets("Replanteo").Cells(z + 1, 11).Value = primero & segundo & tercero & cuarto
Sheets("Replanteo").Cells(z + 1, 12).Value = primero2 & segundo2 & tercero2 & cuarto2

If primero = "S" And segundo = "C" Then
    Sheets("Replanteo").Cells(z + 1, 54).Value = 4.5
ElseIf primero2 = "S" And segundo2 = "C" Then
    Sheets("Replanteo").Cells(z + 1, 55).Value = 4.5
ElseIf primero = "S" And segundo = "S" Then
    Sheets("Replanteo").Cells(z + 1, 54).Value = 4.5
ElseIf primero2 = "S" And segundo2 = "S" Then
    Sheets("Replanteo").Cells(z + 1, 55).Value = 4.5
End If
z = z + 2
Wend
End Sub
Sub sea_ea(z)

    Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
    Sheets("Replanteo").Cells(z + 1, 40).Value = 0
    Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 42).Value = 0
        Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 47).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 48).Value = 0.05
        Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value + 0.3
        Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value - 0.3
        Sheets("Replanteo").Cells(z + 2, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 2, 40).Value = 0
        Sheets("Replanteo").Cells(z + 2, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 2, 46).Value = 0.05
        If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
        ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
        Else
            Sheets("Replanteo").Cells(z + 1, 46).Value = 0.3
        End If
End Sub

Sub ea_sea(z)
        Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 40).Value = 0
        Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 42).Value = 0
        Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 47).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 46).Value = 0.05
        Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z, 39).Value = 1.4
        Sheets("Replanteo").Cells(z, 40).Value = 0
        Sheets("Replanteo").Cells(z, 45).Value = 1.8
        Sheets("Replanteo").Cells(z, 46).Value = 0.05
        If Sheets("Replanteo").Cells(z + 1, 4).Value >= 40.5 Then
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0.5
        ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 31.5 Then
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0.3
        Else
            Sheets("Replanteo").Cells(z + 1, 48).Value = 0.3
        End If
End Sub

Sub sea_sea(z)
        Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 40).Value = 0
        Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 42).Value = 0
        Sheets("Replanteo").Cells(z + 1, 43).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 47).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 48).Value = 0.3
        Sheets("Replanteo").Cells(z + 1, 49).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 50).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 46).Value = 0.5
        Sheets("Replanteo").Cells(z + 2, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 2, 40).Value = 0
        Sheets("Replanteo").Cells(z + 2, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 2, 46).Value = 0.3
End Sub
Sub aa_sea(z, segundo2)
        Sheets("Replanteo").Cells(z + 1, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 40).Value = 0
        Sheets("Replanteo").Cells(z + 1, 41).Value = 1.4
        Sheets("Replanteo").Cells(z + 1, 42).Value = 0
        Sheets("Replanteo").Cells(z + 1, 43).Value = 1.125
        Sheets("Replanteo").Cells(z + 1, 44).Value = 2.5
        Sheets("Replanteo").Cells(z + 1, 47).Value = 1.8
        Sheets("Replanteo").Cells(z + 1, 49).Value = "var"
        Sheets("Replanteo").Cells(z + 1, 50).Value = "var"
        Sheets("Replanteo").Cells(z + 1, 52).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 1, 53).Value = Sheets("Replanteo").Cells(z + 1, 4).Value
        Sheets("Replanteo").Cells(z + 2, 39).Value = 1.4
        Sheets("Replanteo").Cells(z + 2, 40).Value = 0
        If segundo2 = 1 Then
        Sheets("Replanteo").Cells(z + 2, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 2, 46).Value = 0.5
            If Sheets("Replanteo").Cells(z + 1, 4).Value >= 27 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.5
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.8
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.8 + 0.5
            Else
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.5
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.65
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.65 + 0.5
            End If
        ElseIf segundo2 = 3 Then
        Sheets("Replanteo").Cells(z + 2, 45).Value = 1.8
        Sheets("Replanteo").Cells(z + 2, 46).Value = 0.3
            If Sheets("Replanteo").Cells(z + 1, 4).Value >= 27 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.3
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.6
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.6 + 0.5
            ElseIf Sheets("Replanteo").Cells(z + 1, 4).Value >= 27 Then
                Sheets("Replanteo").Cells(z + 1, 48).Value = 0.3
                Sheets("Replanteo").Cells(z + 1, 46).Value = 0.45
                Sheets("Replanteo").Cells(z + 1, 45).Value = 0.45 + 0.5
            End If
        End If
End Sub

