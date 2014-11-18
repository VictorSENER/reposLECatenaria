Attribute VB_Name = "dibujar"
Dim iposte As Double, ipoli As Double, ichevau As Double, fila As Double, num_postes_total As Double
Dim num_chevau As Double, num_antich As Double, num_lineas_total As Double
Dim ips As Long
Const PI = 3.14159265358979
Public cadena_ruta As String
Public polinea As AcadPolyline
Dim chevau(30000) As datoschevau
Dim poste(30000) As datos1
Public poli(30000) As datos2
'Dim poli(5000) As datos
Dim p_s(30000) As punto_singular
Dim matriz_pk(3000) As Double
Dim estacion(100) As est
Public num_traz As Integer
'///
'///Variables para ubicación de datos en estaciones
'///
Private Type est
nombre As String
lado As Boolean
End Type
'///
'///Variables para puntos singulares
'///
Private Type punto_singular
pk_inicio As Double
pk_final As Double
pk_medio As Double
tipo As String
pk_mediox As Double
pk_medioy As Double
angulomedio As Double
pk_iniciox As Double
pk_inicioy As Double
anguloinicio As Double
pk_finalx As Double
pk_finaly As Double
angulofinal As Double
End Type

'///
'///Variables para seccionameintos
'///
Private Type datoschevau
poste_ini As Double
poste_fin As Double
poste_antich As Double
num_vanos As Double
tipo As String
sim_chevau As Boolean
sim_antich As Boolean
AT_ant As Boolean
AT_post As Boolean
PM_ant As Double
PM_post As Double
long_post As Double
long_ant As Double
End Type
'///
'/// variables para datos de los postes
'///
Private Type datos1
anc_cdpa As String
alt_cat(0 To 1)  As String
tipo_secc As String
pendola1 As String
pendola2 As String
conexion As String
proteccion As String
pk_global As Double
pk_terreno As String
pk_coordx As Double
pk_coordy As Double
Pk_vano_post As Double
pk_poli As Single
anguloeje As Double
alfa_vano As Double
pk_vano_postx As Double
pk_vano_posty As Double
vano_post As Double
etiq_1 As String
etiq_2 As String
descentramiento As Double
descentramiento_2mens As Double
at As Boolean
radio As Double
tipo As String
tunel As Boolean
lado As String
flecha As String
implantacion As Single
altura_HC As Single
altura_cat As Single
mensula2a As Boolean
aguja As String
lado_aguja As Boolean
End Type
'///
'///Varibales para la polilinea
'///
Private Type datos2
coordx As Double
coordy As Double
angulo_post As Double
radian_post As Double
dist_acum As Double
End Type

Private Type datos_trazado2
pk As Double
pkx As Double
pky As Double
alfa As Double
pk_centro_posterior As Double
centrox As Double
centroy As Double
alfa_centro As Double
End Type

Dim trazado(1500) As datos_trazado

Private Type datos_trazado
col(4) As datos_trazado2 'Col(1) será ORP1, col(2) será FRP1,etc...
radio As Double
devers As Double
End Type


Private Type canton
polyarray() As Double
End Type
Public pol_canton(0 To 1) As canton
Public qua As Integer
Public acaddoc As AcadApplication

Sub parametros_iniciales_SIRECA()
Sireca.TextBox2.Value = 80
Sireca.TextBox3.Value = 10000
End Sub
Function seleccionar_polilinea(ruta_autocadVB) As String

Dim AcadLayer As AcadLayer
Dim objPoli As AcadLWPolyline
Dim enType As String
Dim coordenadaspoli As Variant
Dim intCode(1) As Integer
Dim varData(1) As Variant
Dim i As Double, j As Double, k As Double
On Error Resume Next
'test for existent cad application
Set acaddoc = GetObject(, "AutoCAD.Application")
If Err Then
    Err.Clear
    'opens a new application if none is available
    Set acaddoc = AcadApplication
    acaddoc.Visible = True
    acaddoc.Documents.Close
    acaddoc.Documents.Open ruta_autocadVB
    If Err Then
        MsgBox "Error opening AutoCAD"
        Exit Function
    End If
    GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "LÍNEAS_OCULTASX2", "acadiso.lin"
    GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "CDPA", "acadiso.lin"
    GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "ACAD_ISO06W100", "acadiso.lin"
    GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "LÍNEAS_OCULTAS", "acadiso.lin"
Else
    '///
    '/// nueva posibilidad de cerrar todos los cads abiertos y abrir únicamente el adecuado.
    '///
    If acaddoc.Documents.Item(0).FullName <> ruta_autocadVB Then
        acaddoc.Documents.Close
        acaddoc.Documents.Open ruta_autocadVB
        acaddoc.Visible = True
        GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "LÍNEAS_OCULTASX2", "acad.lin"
        GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "CDPA", "acad.lin"
        GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "ACAD_ISO06W100", "acad.lin"
        GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "LÍNEAS_OCULTAS", "acadiso.lin"
    End If
    
End If

For Each AcadLayer In GetObject(, "Autocad.Application").ActiveDocument.Layers
    If AcadLayer.Name = "eje_lineal" Then
        intCode(0) = 8: varData(0) = AcadLayer.Name 'only select items on layer
        intCode(1) = 67: varData(1) = 0 'only select items in modelspace - error without this filter
        Set AcadSet = GetObject(, "Autocad.Application").ActiveDocument.SelectionSets.Add(AcadLayer.Name)
        AcadSet.Clear
        AcadSet.Select acSelectionSetAll, , , intCode, varData
        enType = AcadSet.Item(0).ObjectName
        Set objPoli = AcadSet.Item(0)
        AcadSet.Delete
        GoTo pol
    End If
Next
pol:
If enType = "AcDbPolyline" Or enType = "AcDb2dPolyline" Then
    coordenadaspoli = objPoli.Coordinates
    j = 0
    poli(j).coordx = coordenadaspoli(LBound(coordenadaspoli))
    poli(j).coordy = coordenadaspoli(LBound(coordenadaspoli) + 1)
    poli(j).dist_acum = 0
    For i = LBound(coordenadaspoli) To UBound(coordenadaspoli)
        j = j + 1
        poli(j).coordx = coordenadaspoli(i)
        poli(j).coordy = coordenadaspoli(i + 1)
        poli(j).dist_acum = poli(j - 1).dist_acum + Sqr((poli(j).coordx - poli(j - 1).coordx) ^ 2 + (poli(j).coordy - poli(j - 1).coordy) ^ 2)
        If poli(j).coordx <> poli(j - 1).coordx Then
            If (poli(j).coordx - poli(j - 1).coordx) > 0 Then
                poli(j - 1).angulo_post = (180 / PI) * Atn((poli(j).coordy - poli(j - 1).coordy) / (poli(j).coordx - poli(j - 1).coordx))
            ElseIf (poli(j).coordx - poli(j - 1).coordx) < 0 Then
                poli(j - 1).angulo_post = 180 + (180 / PI) * Atn((poli(j).coordy - poli(j - 1).coordy) / (poli(j).coordx - poli(j - 1).coordx))
            End If
            poli(j - 1).radian_post = dibujar.cuadrante(poli(j - 1).coordx, poli(j - 1).coordy, poli(j).coordx, poli(j).coordx, Atn((poli(j).coordy - poli(j - 1).coordy) / (poli(j).coordx - poli(j - 1).coordx)))
        ElseIf (poli(j).coordy - poli(j - 1).coordy) > 0 Then
            poli(j - 1).angulo_post = 90
            poli(j - 1).radian_post = PI / 2
        ElseIf (poli(j).coordy - poli(j - 1).coordy) < 0 Then
            poli(j - 1).angulo_post = 270
            poli(j - 1).radian_post = 3 * PI / 2
        End If
        i = i + 1
    Next i
    num_lineas_total = j
Else
    MsgBox "La entidad seleccionada no es una polilínea"
End If
cadena_ruta = Environ("SIRECA_HOME") & "\core\blocks\" & nombre_cat & "\"

seleccionar_polilinea = cadena_ruta
End Function
Sub Obtener_datos_Excel(inicioVB, finVB)


iposte = 0
'For ichevau = 0 To 999
    'chevau(ichevau).poste_antich = 0
'Next
ichevau = 1
With ActiveWorkbook
    fila = 10
    With .Sheets("Replanteo")
        'For fila = fila_ini To fila_fin
        While .Cells(fila, 33).Value < inicioVB
            fila = fila + 2
        Wend
        While .Cells(fila, 33).Value < finVB And Not IsEmpty(.Cells(fila, 33))
            
            iposte = iposte + 1
            poste(iposte).lado = .Cells(fila, 30).Value
            poste(iposte).etiq_1 = .Cells(fila, 31).Value
            poste(iposte).etiq_2 = .Cells(fila, 32).Value
            poste(iposte).pk_global = .Cells(fila, 33).Value
            poste(iposte).pk_terreno = .Cells(fila, 3).Value
            poste(iposte).altura_HC = .Cells(fila, 10).Value
            poste(iposte).pendola1 = .Cells(fila + 1, 11).Value
            poste(iposte).descentramiento = 1000 * .Cells(fila, 8).Value
            If .Cells(fila, 17).Value <> "" Then
                poste(iposte).anc_cdpa = .Cells(fila, 17).Value
            End If
            
            If .Cells(fila, 9).Value <> "" Then
                If .Cells(fila, 16).Value = eje_aguj Or .Cells(fila, 16).Value = anc_pf & " + " & eje_aguj Or .Cells(fila, 16).Value = eje_aguj & " + " & anc_aguj Then
                    If .Cells(fila, 9).Value < 0 Then
                        poste(iposte).descentramiento_2mens = 1000 * (.Cells(fila, 9).Value + 0.4)
                    Else
                        poste(iposte).descentramiento_2mens = 1000 * (.Cells(fila, 9).Value - 0.4)
                    End If
                Else
                    poste(iposte).descentramiento_2mens = 1000 * .Cells(fila, 9).Value
                End If
            End If
            If .Cells(fila, 47).Value <> "" Then
                poste(iposte).tipo_secc = .Cells(fila, 47).Value
            End If
            If .Cells(fila, 39).Value <> "" Then
                If .Cells(fila, 40).Value = "" Or .Cells(fila, 40).Value = 0 Then
                    poste(iposte).alt_cat(0) = .Cells(fila, 39).Value
                Else
                    poste(iposte).alt_cat(0) = .Cells(fila, 39).Value & " / " & .Cells(fila, 40).Value
                End If
            End If
            If .Cells(fila, 45).Value <> "" Then
                If .Cells(fila, 46).Value = "" Or .Cells(fila, 46).Value = 0 Then
                    poste(iposte).alt_cat(1) = .Cells(fila, 45).Value
                Else
                    poste(iposte).alt_cat(1) = .Cells(fila, 45).Value & " / " & .Cells(fila, 46).Value
                End If
            End If
            If .Cells(fila + 1, 12).Value <> "" Then
                poste(iposte).pendola2 = .Cells(fila + 1, 12).Value
            End If
            
            If .Cells(fila + 1, 13).Value <> "" Then
                poste(iposte).conexion = .Cells(fila + 1, 13).Value
            End If
            If .Cells(fila, 15).Value <> "" Then
                poste(iposte).proteccion = .Cells(fila, 15).Value
            End If
            
            'If .Cells(fila, 15).Value <> "" Then
            poste(iposte).vano_post = Round(.Cells(fila + 1, 4).Value, 2)
                'If poste(iposte - 1).vano_post = 0 Then
                'algo = 0
                'End If
            'End If
            poste(iposte).tipo = .Cells(fila, 16).Value
            If poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = anc_pf & " + " & eje_aguj Or poste(iposte).tipo = eje_pf & " + " & eje_aguj Or poste(iposte).tipo = eje_aguj & " + " & anc_sla_con Then
                poste(iposte).aguja = .Cells(fila, 56).Value
                'poste(iposte).tipo = eje_aguj
            'ElseIf poste(iposte).tipo = anc_pf & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_pf & " + " & semi_eje_aguj Then
                'poste(iposte).tipo = semi_eje_aguj
            End If
            If poste(iposte).tipo = anc_sm_con Or poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_aguj Or poste(iposte).tipo = anc_neutra Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Then
                poste(iposte).at = True
            Else
                poste(iposte).at = False
            End If
            Select Case poste(iposte).tipo
                Case anc_sm_sin
                    poste(iposte).tipo = anc_sm_con
                Case anc_sla_sin
                    poste(iposte).tipo = anc_sla_con
                Case "Anc.Aigu.sans AT"
                    poste(iposte).tipo = anc_aguj
                Case "Anc.Neutre sans AT"
                    poste(iposte).tipo = anc_neutra
            End Select
            If poste(iposte).tipo = anc_sm_con Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_neutra Then
                If iposte > 3 Then
                    If poste(iposte - 3).tipo = anc_sm_con And poste(iposte - 1).tipo = semi_eje_sm Then
                        chevau(ichevau).tipo = "CHEVAU"
                        chevau(ichevau).AT_post = poste(iposte).at
                        chevau(ichevau).AT_ant = poste(iposte - 3).at
                        chevau(ichevau).poste_ini = iposte - 3
                        chevau(ichevau).poste_fin = iposte
                        chevau(ichevau).num_vanos = 3
                        If poste(iposte - 1).vano_post = poste(iposte - 3).vano_post Then
                            chevau(ichevau).sim_chevau = True
                        Else
                            chevau(ichevau).sim_chevau = True
                        End If
                        ichevau = ichevau + 1
                    End If
                End If
                If iposte > 4 Then
                    If (poste(iposte - 4).tipo = anc_sm_con Or poste(iposte - 4).tipo = anc_sla_con) Then
                        If poste(iposte).tipo = anc_sm_con Then
                            chevau(ichevau).tipo = "CHEVAU"
                        End If
                        If poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = eje_aguj & " + " & anc_sla_con Then
                            chevau(ichevau).tipo = "SECTION"
                        End If
                        chevau(ichevau).AT_post = poste(iposte).at
                        chevau(ichevau).AT_ant = poste(iposte - 4).at
                        chevau(ichevau).poste_ini = iposte - 4
                        chevau(ichevau).poste_fin = iposte
                        chevau(ichevau).num_vanos = 4
                        If poste(iposte - 1).vano_post = poste(iposte - 4).vano_post And poste(iposte - 2).vano_post = poste(iposte - 3).vano_post And chevau(ichevau).AT_ant = chevau(ichevau).AT_post Then
                            chevau(ichevau).sim_chevau = True
                        Else
                            chevau(ichevau).sim_chevau = False
                        End If
                        ichevau = ichevau + 1
                    End If
                End If
                If iposte > 5 Then
                    If (poste(iposte - 5).tipo = anc_sm_con Or poste(iposte - 5).tipo = anc_sla_con) Then
                        If poste(iposte).tipo = anc_sm_con Then
                            chevau(ichevau).tipo = "CHEVAU"
                        End If
                        If poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Then
                            chevau(ichevau).tipo = "SECTION"
                        End If
                        chevau(ichevau).AT_post = poste(iposte).at
                        chevau(ichevau).AT_ant = poste(iposte - 5).at
                        chevau(ichevau).poste_ini = iposte - 5
                        chevau(ichevau).poste_fin = iposte
                        chevau(ichevau).num_vanos = 5
                        If poste(iposte - 1).vano_post = poste(iposte - 5).vano_post And poste(iposte - 2).vano_post = poste(iposte - 3).vano_post And chevau(ichevau).AT_ant = chevau(ichevau).AT_post Then
                            chevau(ichevau).sim_chevau = True
                        Else
                            chevau(ichevau).sim_chevau = False
                        End If
                        ichevau = ichevau + 1
                    End If
                End If
                'NUEVO
                If iposte > 6 Then
                    If poste(iposte - 6).tipo = anc_neutra Then
                        chevau(ichevau).tipo = "ZN"
                        chevau(ichevau).AT_post = poste(iposte).at
                        chevau(ichevau).AT_ant = poste(iposte - 6).at
                        chevau(ichevau).poste_ini = iposte - 6
                        chevau(ichevau).poste_fin = iposte
                        chevau(ichevau).num_vanos = 6
                        chevau(ichevau).sim_chevau = False
                        ichevau = ichevau + 1
                    End If
                End If
            End If
            If poste(iposte).tipo = eje_pf Then
                chevau(ichevau - 1).poste_antich = iposte
            End If
            If poste(iposte).tipo = anc_pf And poste(iposte - 1).tipo = eje_pf Then
                If (poste(iposte - 1).vano_post = poste(iposte - 2).vano_post) Then
                    chevau(ichevau - 1).sim_antich = True
                Else
                    chevau(ichevau - 1).sim_antich = False
                End If
            End If

            poste(iposte).mensula2a = False
            If poste(iposte).tipo = eje_sm Or poste(iposte).tipo = eje_sla Or poste(iposte).tipo = semi_eje_sm Or poste(iposte).tipo = semi_eje_sla Or poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = semi_eje_neutra Or poste(iposte).tipo = eje_neutra Or Mid(poste(iposte).tipo, 15) = eje_aguj Or Mid(poste(iposte).tipo, 15) = semi_eje_aguj _
            Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj Then
                poste(iposte).mensula2a = True
            End If
            poste(iposte).tunel = False
            If .Cells(fila, 38) = "Tunel" Then
                poste(iposte).tunel = True
            End If
            poste(iposte).implantacion = .Cells(fila, 5).Value
            poste(iposte).altura_HC = Round(.Cells(fila, 10).Value, 2)
            poste(iposte).radio = .Cells(fila, 6).Value
            fila = fila + 2
        'Next
        Wend
    num_postes_total = iposte
    num_chevau = ichevau - 1
    End With
    
'///
'///Obtener los datos de los puntos singulares en la hoja 4
'///
    With .Sheets("Punto singular")
        fila = 3
        ips = 1
        While .Cells(fila, 2).Value < inicioVB
            fila = fila + 1
        Wend
        While Not IsEmpty(.Cells(fila, 2).Value) And .Cells(fila, 2).Value < finVB
            If .Cells(fila, 1).Value = "Aguja" Then
                GoTo fin
            End If
            p_s(ips).pk_inicio = .Cells(fila, 2).Value
            p_s(ips).pk_final = .Cells(fila, 21).Value
            p_s(ips).pk_medio = (p_s(ips).pk_final + p_s(ips).pk_inicio) / 2
            p_s(ips).tipo = .Cells(fila, 23).Value
            ips = ips + 1
fin:
            fila = fila + 1
        Wend
    End With

'///
'/// Obtener los datos de la ubicación de los datos en estaciones
'///
    With .Sheets("Extra")
        fila = 3
        ipl = 1
        While Not IsEmpty(.Cells(fila, 17).Value)
            estacion(ipl).nombre = .Cells(fila, 17).Value
            estacion(ipl).lado = .Cells(fila, 18).Value
            ipl = ipl + 1
            fila = fila + 1
        Wend
    End With

'///
'/// obtener las ubicaciones de los PK no lineales
'///
Dim pkreal(400) As Double
    With .Sheets("Pk real")
        fila = 3
        ipl = 1
        While Not IsEmpty(.Cells(fila, 2).Value)
            pkreal(ipl) = .Cells(fila, 2).Value
            ipl = ipl + 1
            fila = fila + 1
        Wend
    End With
End With
'Set objHoja = Nothing
'Set objLibro = Nothing
'Set objExcel = Nothing
For ichevau = 1 To num_chevau
    If num_chevau > ichevau Then
        chevau(ichevau).PM_ant = Round((poste(chevau(ichevau + 1).poste_fin).pk_global - poste(chevau(ichevau).poste_ini).pk_global) / (chevau(ichevau + 1).poste_fin - chevau(ichevau).poste_ini), 2)
        chevau(ichevau).long_post = Round(poste(chevau(ichevau + 1).poste_fin).pk_global - poste(chevau(ichevau).poste_ini).pk_global, 2)
    End If
    If ichevau > 1 Then
        chevau(ichevau).PM_post = Round((poste(chevau(ichevau).poste_fin).pk_global - poste(chevau(ichevau - 1).poste_ini).pk_global) / (chevau(ichevau).poste_fin - chevau(ichevau - 1).poste_ini), 2)
        chevau(ichevau).long_ant = Round(poste(chevau(ichevau).poste_fin).pk_global - poste(chevau(ichevau - 1).poste_ini).pk_global, 2)
    End If
Next
End Sub
Sub Encontrar_coordenadas_pk()
For iposte = 1 To num_postes_total
    ipoli = 1
    Do While poli(ipoli).dist_acum < poste(iposte).pk_global And poste(iposte).pk_global <> 0
        ipoli = ipoli + 1
    Loop
    poste(iposte).pk_coordx = poli(ipoli - 1).coordx + (poste(iposte).pk_global - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
    poste(iposte).pk_coordy = poli(ipoli - 1).coordy + (poste(iposte).pk_global - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
    poste(iposte).anguloeje = poli(ipoli - 1).angulo_post
    If iposte <> 1 Then
        poste(iposte - 1).Pk_vano_post = poste(iposte - 1).pk_global + (poste(iposte).pk_global - poste(iposte - 1).pk_global) / 2
        ipoli = 1
        Do While poli(ipoli).dist_acum < poste(iposte - 1).Pk_vano_post
            ipoli = ipoli + 1
        Loop
        poste(iposte - 1).alfa_vano = poli(ipoli - 1).angulo_post * PI / 180
        poste(iposte - 1).pk_vano_postx = poli(ipoli - 1).coordx + (poste(iposte - 1).Pk_vano_post - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
        poste(iposte - 1).pk_vano_posty = poli(ipoli - 1).coordy + (poste(iposte - 1).Pk_vano_post - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
    End If
Next iposte
For iposte = 1 To ips
    ipoli = 1
    Do While poli(ipoli).dist_acum < p_s(iposte).pk_inicio
        ipoli = ipoli + 1
    Loop
Next iposte

ipoli = 1

For iposte = 1 To ips
    Do While poli(ipoli).dist_acum < p_s(iposte).pk_inicio
        ipoli = ipoli + 1
    Loop
    p_s(iposte).pk_iniciox = poli(ipoli - 1).coordx + (p_s(iposte).pk_inicio - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
    p_s(iposte).pk_inicioy = poli(ipoli - 1).coordy + (p_s(iposte).pk_inicio - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
    p_s(iposte).anguloinicio = poli(ipoli - 1).angulo_post
    Do While poli(ipoli).dist_acum < p_s(iposte).pk_medio
        ipoli = ipoli + 1
    Loop
    p_s(iposte).pk_mediox = poli(ipoli - 1).coordx + (p_s(iposte).pk_medio - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
    p_s(iposte).pk_medioy = poli(ipoli - 1).coordy + (p_s(iposte).pk_medio - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
    p_s(iposte).angulomedio = poli(ipoli - 1).angulo_post
    Do While poli(ipoli).dist_acum < p_s(iposte).pk_final
        ipoli = ipoli + 1
    Loop
    p_s(iposte).pk_finalx = poli(ipoli - 1).coordx + (p_s(iposte).pk_final - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
    p_s(iposte).pk_finaly = poli(ipoli - 1).coordy + (p_s(iposte).pk_final - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
    p_s(iposte).angulofinal = poli(ipoli - 1).angulo_post
Next iposte
End Sub
Function convertir_pk(pk As Double) As String
Dim ipk As Double
Dim ceros As String
'Dim beta As Double
ipk = 0
    If 55453.6631 <= pk And pk < 56453.5677 Then
        ipk = True
    Else
        Do While matriz_pk(ipk) < pk
            ipk = ipk + 1
        Loop
        ipk = ipk - 1
    End If
    If ipk = True Then
        If ((pk - 55744.4762) - 1000 * (Int((pk - 55744.4762) / 1000))) < 100 Then
            If ((pk - 55744.4762) - 1000 * (Int((pk - 55744.4762) / 1000))) < 10 Then
                ceros = "00"
            Else
                ceros = "0"
            End If
        Else
            ceros = ""
        End If
        convertir_pk = "55bis" & "+" & ceros & Round((pk - 55744.4762), 2)
    Else
        pk_te = (1000 * CDbl(ipk) + pk - matriz_pk(ipk))
    
        If (pk_te - 1000 * (Int(pk_te / 1000))) < 100 Then
            If (pk_te - 1000 * (Int(pk_te / 1000))) < 10 Then
                ceros = "00"
            Else
                ceros = "0"
            End If
        Else
            ceros = ""
        End If
        convertir_pk = Int(pk_te / 1000) & "+" & ceros & (Round(pk_te - 1000 * Int(pk_te / 1000), 2))
        'convertir_pk = (1000 * CDbl(ipk) + pk - matriz_pk(ipk))
    End If
End Function

Sub dibujar_datos_poste(cadena_ruta As String, HDC As Boolean)
Dim bloque_datos As AcadBlockReference
Dim Atributos As Variant
Dim insertionpnt(0 To 2) As Double
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-Datos_poste")
dist_eje = 12.5
ancho_poste = 0.23
eje_rail = 0.79
escala_datos = 1
cont = 0
ancho_poste = 0.46
For iposte = 1 To num_postes_total
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = -22
            ElseIf estacion(ipk).lado = False Then
                dist_eje = 22
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = 12
            cont_est = 0
            ipk = 0
        End If
        '///
        '/// Variar la escala en estaciones
        '///
    
        If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
            escala_eti = 0.5
            cont = cont + 1
        ElseIf (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont = 3 Then
            escala_eti = 1
            cont = 0
        End If
    
    
        alfa_datos = poste(iposte).anguloeje * PI / 180
        insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_datos + PI / 2)
        insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_datos + PI / 2)
        insertionpnt(2) = 0
        If poste(iposte).tipo = anc_sla_con And cont < 3 And iposte <> 1 And iposte <> 6 Then
            escala_datos = 0.75
            cont = cont + 1
        ElseIf poste(iposte - 1).tipo = anc_sla_con And poste(iposte).tipo = "" And cont = 3 Then
            escala_datos = 1
            cont = 0
        End If
    
    
        Set bloque_datos = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "datos.dwg", escala_datos, escala_datos, escala_datos, alfa_datos)
        bloque_datos.Layer = "E-Datos_poste"
        Atributos = bloque_datos.GetAttributes
        If poste(iposte).tipo = "" Then
            Atributos(0).TextString = "VC"
        Else
        Atributos(0).TextString = poste(iposte).tipo
        End If
        Atributos(1).TextString = formato_pk(poste(iposte).pk_global, iposte)
        Select Case poste(iposte).lado
            Case "G"
                Atributos(2).TextString = "X: " & Round(poste(iposte).pk_coordx + (poste(iposte).implantacion + ancho_poste + eje_rail) * Cos(alfa_datos + PI / 2), 2)
                Atributos(3).TextString = "Y: " & Round(poste(iposte).pk_coordy + (poste(iposte).implantacion + ancho_poste + eje_rail) * sin(alfa_datos + PI / 2), 2)
            Case "D"
                Atributos(2).TextString = "X: " & Round(poste(iposte).pk_coordx + (poste(iposte).implantacion + ancho_poste + eje_rail) * Cos(alfa_datos - PI / 2), 2)
                Atributos(3).TextString = "Y: " & Round(poste(iposte).pk_coordy + (poste(iposte).implantacion + ancho_poste + eje_rail) * sin(alfa_datos - PI / 2), 2)
        End Select
        If poste(iposte).tunel = True Then
            Atributos(2).TextString = "X: " & Round(poste(iposte).pk_coordx, 2)
            Atributos(3).TextString = "Y: " & Round(poste(iposte).pk_coordy, 2)
        End If
Next
End Sub
Function formato_pk(pk As Double, ipk As Double) As String
Dim ceros As String
If 55453.6631 <= pk And pk < 56453.5677 Then
    formato_pk = poste(ipk).pk_terreno
Else
    pk_te = pk
    If (pk_te - 1000 * (Int(pk_te / 1000))) < 100 Then
        If (pk_te - 1000 * (Int(pk_te / 1000))) < 10 Then
            ceros = "00"
        Else
            ceros = "0"
        End If
    Else
        ceros = ""
    End If
    formato_pk = Int(pk_te / 1000) & "+" & ceros & (Round(pk_te - 1000 * Int(pk_te / 1000), 2))
    
End If
End Function
Sub dibujar_cantones(cadena_ruta As String, HDC As Boolean)
'Dim pk_1 As Double
'Dim pk_2, ceros As String
Dim punto_fijo As Double
Dim cadena_bloque_chevau As String
Dim linea As AcadLine
Dim ini_linea(0 To 2) As Double
Dim fin_linea(0 To 2) As Double
Dim insertionpnt(0 To 2) As Double
Dim escala As Double
Dim Atributos As Variant
Dim bloque_chevau As AcadBlockReference
'Dim capa As String
'Dim count As Integer
ancho_poste = 0.46
Dim pk_eje_chevau As String
Dim alfa_chevau As Double
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-CANTONES")
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-AUX")
dist_axe_chevau = 25
dist_chevau = 20
dist_sim = 35
dist_secc = 15
cont = 0
ruta_chevau = ""
escala_chevau = 1
Call Obtener_excel_pks
For ichevau = 1 To num_chevau
        '///
        '/// Variar la escala en estaciones
        '///
        
        
        'If chevau(ichevau).tipo = "SECTION" And cont = 0 And ichevau > 1 Then
            ruta_chevau = "_est"
            cont = cont + 1
            escala_chevau = 0.5
            dist_sim = 20
            dist_secc = 8
        'ElseIf (chevau(ichevau - 1).tipo = "SECTION") And chevau(ichevau).tipo = "CHEVAU" And chevau(ichevau + 1).tipo = "CHEVAU" And cont >= 1 Then
            'ruta_chevau = ""
            'cont = 0
            'escala_chevau = 1
            'dist_sim = 35
            'dist_secc = 15
        'End If
        
        ini_linea(0) = poste(chevau(ichevau).poste_ini).pk_coordx + dist_chevau * Cos(poste(chevau(ichevau).poste_ini).anguloeje * PI / 180 + PI / 2)
        ini_linea(1) = poste(chevau(ichevau).poste_ini).pk_coordy + dist_chevau * sin(poste(chevau(ichevau).poste_ini).anguloeje * PI / 180 + PI / 2)
        ini_linea(2) = 0
        fin_linea(0) = poste(chevau(ichevau).poste_fin).pk_coordx + dist_chevau * Cos(poste(chevau(ichevau).poste_fin).anguloeje * PI / 180 + PI / 2)
        fin_linea(1) = poste(chevau(ichevau).poste_fin).pk_coordy + dist_chevau * sin(poste(chevau(ichevau).poste_fin).anguloeje * PI / 180 + PI / 2)
        fin_linea(2) = 0
        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, fin_linea)
        alfa_chevau = linea.angle
        linea.Layer = "E-AUX"
        escala = Sqr((fin_linea(1) - ini_linea(1)) ^ 2 + (fin_linea(0) - ini_linea(0)) ^ 2) / 191.1883
        Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(ini_linea, cadena_ruta & "Chevau" & ruta_chevau & ".dwg", escala, 1#, 1#, alfa_chevau)
        bloque_chevau.Layer = "E-CANTONES"
        If bloque_chevau.HasAttributes = True Then
            Atributos = bloque_chevau.GetAttributes
            If chevau(ichevau).tipo = "CHEVAU" Then
                Atributos(0).TextString = "CHEVAUCHEMENT"
            End If
            If chevau(ichevau).tipo = "SECTION" Then
                Atributos(0).TextString = "SECTIONNEMENT"
            End If
            If chevau(ichevau).tipo = "ZN" Then
                Atributos(0).TextString = "SECTION NEUTRE DE SEPARTION DE PHASES"
            End If
            If ichevau = num_chevau Then
                Atributos(1).TextString = "PM=xx.xx"
            End If
            If ichevau <> num_chevau Then
                Atributos(1).TextString = "PM=" & chevau(ichevau).PM_ant & " m"
            End If
            If ichevau = 1 Then
                Atributos(2).TextString = "PM=xx.xx" 'PM DER
            End If
             If ichevau <> 1 Then
                Atributos(2).TextString = "PM=" & chevau(ichevau).PM_post & " m"
            End If
            If chevau(ichevau).AT_ant = True Then
                Atributos(3).TextString = "Anc. Cat. Rég. avec A.T."
            End If
            If chevau(ichevau).AT_ant = False Then
                Atributos(3).TextString = "Anc. Cat. Rég. sans A.T."
            End If
            If chevau(ichevau).AT_post = True Then
                Atributos(4).TextString = "Anc. Cat. Rég. avec A.T."
            End If
            If chevau(ichevau).AT_post = False Then
                Atributos(4).TextString = "Anc. Cat. Rég. sans A.T."
            End If
            'Atributos(1).TextString = "PM=xxxx m" 'PM IZQ
            'Atributos(2).TextString = "PM=xxxx m" 'PM DER
            'Atributos(3).TextString = "Anc. Cat. Rég. avec A.T." 'ANC IZQ
            'Atributos(4).TextString = "Anc. Cat. Rég. avec A.T." ' ANC. DER
        End If
        If chevau(ichevau).num_vanos = 5 Then
        'poste(chevau(ichevau).poste_ini + 2).vano_post
            insertionpnt(0) = poste(chevau(ichevau).poste_ini + 2).pk_vano_postx + dist_axe_chevau * Cos(poste(chevau(ichevau).poste_ini + 2).alfa_vano + PI / 2)
            insertionpnt(1) = poste(chevau(ichevau).poste_ini + 2).pk_vano_posty + dist_axe_chevau * sin(poste(chevau(ichevau).poste_ini + 2).alfa_vano + PI / 2)
            'insertionpnt(0) = poste(chevau(ichevau).poste_ini + 3).pk_coordx + dist_axe_chevau * Cos(poste(chevau(ichevau).poste_ini + 3).anguloeje * PI / 180 + PI / 2)
            'insertionpnt(1) = poste(chevau(ichevau).poste_ini + 3).pk_coordy + dist_axe_chevau * sin(poste(chevau(ichevau).poste_ini + 3).anguloeje * PI / 180 + PI / 2)
            insertionpnt(2) = 0
            'alfa_chevau = poste(chevau(ichevau).poste_ini + 3).anguloeje * PI / 180
            alfa_chevau = poste(chevau(ichevau).poste_ini + 2).alfa_vano
            If chevau(ichevau).sim_chevau = True Then
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Simetria.dwg", escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Axe_Chevau.dwg", escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
            bloque_chevau.Layer = "E-CANTONES"
            On Error Resume Next
                If Err Then
                    Err.Clear
                    pk_eje_chevau = poste(chevau(ichevau).poste_ini + 2).pk_terreno
                Else
                    pk_eje_chevau = formato_pk(Round(poste(chevau(ichevau).poste_ini + 2).pk_terreno, 2) + Round(poste(chevau(ichevau).poste_ini + 2).vano_post / 2, 2), chevau(ichevau).poste_ini + 2)
                End If
            
            
            'pk_eje = Round(poste(chevau(ichevau).poste_ini + 2).pk_terreno, 2) + Round(poste(chevau(ichevau).poste_ini + 2).vano_post / 2, 2) 'CAMBIAR
            'pk_eje_chevau = formato_pk(Round(poste(chevau(ichevau).poste_ini + 2).pk_terreno, 2) + Round(poste(chevau(ichevau).poste_ini + 2).vano_post / 2, 2), chevau(ichevau).poste_ini + 2) 'CAMBIAR
        End If
        If chevau(ichevau).num_vanos = 4 Then
            insertionpnt(0) = poste(chevau(ichevau).poste_ini + 2).pk_coordx + dist_sim * Cos(poste(chevau(ichevau).poste_ini + 2).anguloeje * PI / 180 + PI / 2)
            insertionpnt(1) = poste(chevau(ichevau).poste_ini + 2).pk_coordy + dist_sim * sin(poste(chevau(ichevau).poste_ini + 2).anguloeje * PI / 180 + PI / 2)
            insertionpnt(2) = 0
            alfa_chevau = poste(chevau(ichevau).poste_ini + 2).anguloeje * PI / 180
            If chevau(ichevau).sim_chevau = True Then
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Simetria.dwg", escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Axe_Chevau.dwg", escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
            bloque_chevau.Layer = "E-CANTONES"
            
            On Error Resume Next
                If Err Then
                    Err.Clear
                    pk_eje_chevau = poste(chevau(ichevau).poste_ini + 2).pk_terreno
                Else
                    pk_eje_chevau = formato_pk(Round(poste(chevau(ichevau).poste_ini + 2).pk_terreno, 2), chevau(ichevau).poste_ini + 2)
                End If
            'pk_eje_chevau = Round(poste(chevau(ichevau).poste_ini + 2).pk_terreno, 2) 'CAMBIAR
            'pk_eje_chevau = formato_pk(Round(poste(chevau(ichevau).poste_ini + 2).pk_terreno, 2), chevau(ichevau).poste_ini + 2) 'CAMBIAR
        End If
        If chevau(ichevau).num_vanos = 3 Then
            insertionpnt(0) = poste(chevau(ichevau).poste_ini + 1).pk_vano_postx + dist_axe_chevau * Cos(poste(chevau(ichevau).poste_ini + 1).alfa_vano + PI / 2)
            insertionpnt(1) = poste(chevau(ichevau).poste_ini + 1).pk_vano_posty + dist_axe_chevau * sin(poste(chevau(ichevau).poste_ini + 1).alfa_vano + PI / 2)
            insertionpnt(2) = 0
            alfa_chevau = poste(chevau(ichevau).poste_ini + 1).alfa_vano
             If chevau(ichevau).sim_chevau = True Then
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Simetria.dwg", escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Axe_Chevau.dwg", escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
            bloque_chevau.Layer = "E-CANTONES"
            On Error Resume Next
                If Err Then
                    Err.Clear
                    pk_eje_chevau = poste(chevau(ichevau).poste_ini + 1).pk_terreno
                Else
                    pk_eje_chevau = formato_pk(Round(poste(chevau(ichevau).poste_ini + 1).pk_terreno, 2) + Round(poste(chevau(ichevau).poste_ini + 1).vano_post / 2, 2), chevau(ichevau).poste_ini + 2)
                End If
            'pk_eje_chevau = Round(poste(chevau(ichevau).poste_ini + 1).Pk_vano_post, 2)
            'pk_eje_chevau = formato_pk(Round(poste(chevau(ichevau).poste_ini + 1).pk_terreno, 2) + Round(poste(chevau(ichevau).poste_ini + 1).vano_post / 2, 2), chevau(ichevau).poste_ini + 2) 'CAMBIAR
        End If
        Atributos = bloque_chevau.GetAttributes
        Atributos(0).TextString = pk_eje_chevau
        'Atributos(0).TextString = convertir_pk_FT("lineal_terreno", pk_eje_chevau)
        insertionpnt(0) = insertionpnt(0) + dist_secc * Cos(alfa_chevau + PI / 2)
        insertionpnt(1) = insertionpnt(1) + dist_secc * sin(alfa_chevau + PI / 2)
        insertionpnt(2) = 0
        If chevau(ichevau).tipo = "CHEVAU" Then
            If ichevau = 1 Then
                Atributos(1).TextString = "xxxx.xx"
                If chevau(ichevau).AT_post = True Then
                    cadena_bloque_chevau = cadena_ruta & "Chevauchement4.dwg"
                End If
                If chevau(ichevau).AT_post = False Then
                    cadena_bloque_chevau = cadena_ruta & "Chevauchement1.dwg"
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau + PI)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            If ichevau = num_chevau Then
                Atributos(2).TextString = "xxxx.xx"
                If chevau(ichevau).AT_ant = True Then
                    cadena_bloque_chevau = cadena_ruta & "Chevauchement4.dwg"
                End If
                If chevau(ichevau).AT_ant = False Then
                    cadena_bloque_chevau = cadena_ruta & "Chevauchement1.dwg"
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            If ichevau > 1 Then
                Atributos(1).TextString = Round(chevau(ichevau).long_ant, 2)
                If chevau(ichevau).AT_post = True Then
                    If chevau(ichevau - 1).AT_ant = True Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement4.dwg"
                    End If
                    If chevau(ichevau - 1).AT_ant = False Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement2.dwg"
                    End If
                End If
                If chevau(ichevau).AT_post = False Then
                    If chevau(ichevau - 1).AT_ant = True Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement1.dwg"
                    End If
                    If chevau(ichevau - 1).AT_ant = False Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement3.dwg"
                    End If
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau + PI)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            If ichevau < num_chevau Then
                Atributos(2).TextString = Round(chevau(ichevau).long_post, 2)
                If chevau(ichevau).AT_ant = True Then
                    If chevau(ichevau + 1).AT_post = True Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement4.dwg"
                    End If
                    If chevau(ichevau + 1).AT_post = False Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement2.dwg"
                    End If
                End If
                If chevau(ichevau).AT_ant = False Then
                    If chevau(ichevau + 1).AT_post = True Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement1.dwg"
                    End If
                    If chevau(ichevau + 1).AT_post = False Then
                        cadena_bloque_chevau = cadena_ruta & "Chevauchement3.dwg"
                    End If
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
        End If
        If (chevau(ichevau).tipo = "SECTION" Or chevau(ichevau).tipo = "ZN") Then
            If ichevau = 1 Then
                Atributos(1).TextString = "xxxx.xx"
                If chevau(ichevau).AT_post = True Then
                    cadena_bloque_chevau = cadena_ruta & "Sectionnement_1.dwg"
                End If
                If chevau(ichevau).AT_post = False Then
                    cadena_bloque_chevau = cadena_ruta & "Sectionnement_2.dwg"
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau + PI)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            If ichevau = num_chevau Then
                Atributos(2).TextString = "xxxx.xx"
                If chevau(ichevau).AT_ant = True Then
                    cadena_bloque_chevau = cadena_ruta & "Sectionnement_1.dwg"
                End If
                If chevau(ichevau).AT_ant = False Then
                    cadena_bloque_chevau = cadena_ruta & "Sectionnement_2.dwg"
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            If ichevau > 1 Then
                Atributos(1).TextString = Round(chevau(ichevau).long_ant, 2)
                If chevau(ichevau).AT_post = True Then
                    If chevau(ichevau - 1).AT_ant = True Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_1.dwg"
                    End If
                    If chevau(ichevau - 1).AT_ant = False Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_3.dwg"
                    End If
                End If
                If chevau(ichevau).AT_post = False Then
                    If chevau(ichevau - 1).AT_ant = True Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_2.dwg"
                    End If
                    If chevau(ichevau - 1).AT_ant = False Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_4.dwg"
                    End If
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau + PI)
                bloque_chevau.Layer = "E-CANTONES"
            End If
            If ichevau < num_chevau Then
                Atributos(2).TextString = Round(chevau(ichevau).long_post, 2)
                If chevau(ichevau).AT_ant = True Then
                    If chevau(ichevau + 1).AT_post = True Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_1.dwg"
                    End If
                    If chevau(ichevau + 1).AT_post = False Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_3.dwg"
                    End If
                End If
                If chevau(ichevau).AT_ant = False Then
                    If chevau(ichevau + 1).AT_post = True Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_2.dwg"
                    End If
                    If chevau(ichevau + 1).AT_post = False Then
                        cadena_bloque_chevau = cadena_ruta & "Sectionnement_4.dwg"
                    End If
                End If
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_bloque_chevau, escala_chevau, escala_chevau, escala_chevau, alfa_chevau)
                bloque_chevau.Layer = "E-CANTONES"
            End If
        End If
Next
For ichevau = 0 To num_chevau
        '///
        '/// Variar la escala en estaciones
        '///
        'If chevau(ichevau).tipo = "SECTION" And cont = 0 And ichevau > 1 Then
            ruta_chevau = "_est"
            cont = cont + 1
            escala_chevau = 0.5
            dist_sim = 25
            dist_secc = 8
        'ElseIf chevau(ichevau).tipo = "SECTION" And cont >= 1 Then
            'ruta_chevau = ""
            'cont = 0
            'escala_chevau = 1
            'dist_sim = 30
            'dist_secc = 15
        'End If
    If chevau(ichevau).poste_antich <> 0 Then
        ini_linea(0) = poste(chevau(ichevau).poste_antich - 1).pk_coordx + dist_axe_chevau * Cos(poste(chevau(ichevau).poste_antich - 1).anguloeje * PI / 180 + PI / 2)
        ini_linea(1) = poste(chevau(ichevau).poste_antich - 1).pk_coordy + dist_axe_chevau * sin(poste(chevau(ichevau).poste_antich - 1).anguloeje * PI / 180 + PI / 2)
        ini_linea(2) = 0
        fin_linea(0) = poste(chevau(ichevau).poste_antich + 1).pk_coordx + dist_axe_chevau * Cos(poste(chevau(ichevau).poste_antich + 1).anguloeje * PI / 180 + PI / 2)
        fin_linea(1) = poste(chevau(ichevau).poste_antich + 1).pk_coordy + dist_axe_chevau * sin(poste(chevau(ichevau).poste_antich + 1).anguloeje * PI / 180 + PI / 2)
        fin_linea(2) = 0
        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, fin_linea)
        alfa_chevau = linea.angle
        linea.Layer = "E-AUX"
        escala = Sqr((fin_linea(1) - ini_linea(1)) ^ 2 + (fin_linea(0) - ini_linea(0)) ^ 2) / 108
        Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(ini_linea, cadena_ruta & "Antich" & ruta_chevau & ".dwg", escala, 1#, 1#, alfa_chevau)
        bloque_chevau.Layer = "E-CANTONES"
        Atributos = bloque_chevau.GetAttributes
        Atributos(0).TextString = "ANTICHEMINEMENT"
        Atributos(1).TextString = "Anc. Anticheminement"
        Atributos(2).TextString = "Anc. Anticheminement"
        insertionpnt(0) = poste(chevau(ichevau).poste_antich).pk_coordx + dist_sim * Cos(poste(chevau(ichevau).poste_antich).anguloeje * PI / 180 + PI / 2)
        insertionpnt(1) = poste(chevau(ichevau).poste_antich).pk_coordy + dist_sim * sin(poste(chevau(ichevau).poste_antich).anguloeje * PI / 180 + PI / 2)
        insertionpnt(2) = 0
        If chevau(ichevau).sim_antich = True Then
                Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Simetria.dwg", escala_chevau, escala_chevau, escala_chevau, poste(chevau(ichevau).poste_antich).anguloeje * PI / 180)
                bloque_chevau.Layer = "E-CANTONES"
        End If
        Set bloque_chevau = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Axe_Antich.dwg", escala_chevau, escala_chevau, escala_chevau, poste(chevau(ichevau).poste_antich).anguloeje * PI / 180)
        bloque_chevau.Layer = "E-CANTONES"
        Atributos = bloque_chevau.GetAttributes
        Atributos(0).TextString = convertir_pk_FT("lineal_terreno", poste(chevau(ichevau).poste_antich).pk_global)
        'Atributos(0).TextString = convertir_pk_FT("lineal_terreno", poste(chevau(ichevau).poste_antich).pk_global)
        If ichevau = 0 Then
            Atributos(1).TextString = "xxx.xx"
        End If
        If ichevau = num_chevau Then
            Atributos(2).TextString = "xxx.xx"
        End If
        If ichevau > 0 Then
            Atributos(1).TextString = Round((poste(chevau(ichevau).poste_antich).pk_global - poste(chevau(ichevau).poste_ini).pk_global), 2)
        End If
        If ichevau < num_chevau Then
            Atributos(2).TextString = Round(poste(chevau(ichevau + 1).poste_fin).pk_global - poste(chevau(ichevau).poste_antich).pk_global, 2)
        End If
    End If
Next

End Sub
Sub dibujar_descentramientos(cadena_ruta As String, HDC As Boolean)
Dim bloque_desc As AcadBlockReference
Dim Atributos As Variant
Dim insertionpnt(0 To 2) As Double
Dim dist_desc As Single
Dim alfa_desc As Double
Dim cadena_desc As String
Dim ipk As Long
escala_des = 1
ancho_poste = 0.46
dist_poste = 3
cont = 0
cont_est = 0
ipk = 0
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-DESCENTRAMIENTOS")

For iposte = 1 To num_postes_total
    alfa_desc = poste(iposte).anguloeje * PI / 180
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_des = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte - 1).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
       'escala_des = 1
        'cont = 0
    'End If
    
    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 7)
                dist_poste = 1.5
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 9)
                dist_poste = 1.5
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 5)
            dist_poste = 3
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 7)
                dist_poste = 1.5
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 9)
                dist_poste = 1.5
            End If
        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 5)
        End If
    ElseIf HDC = True Then
        If poste(iposte).lado = "G" And cont < 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 3)
            dist_poste = 1.5
        ElseIf poste(iposte).lado = "D" And cont < 1 Then
            dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            dist_poste = 1.5
        ElseIf poste(iposte).lado = "G" And cont >= 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril)
            dist_poste = 1
        ElseIf poste(iposte).lado = "D" And cont >= 1 Then
            dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 0.5)
            dist_poste = 1
        End If
    End If
   
        
        If poste(iposte).mensula2a = True Then
            insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_desc - PI / 2) - dist_poste * Cos(alfa_desc - PI)
            insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_desc - PI / 2) - dist_poste * sin(alfa_desc - PI)
            insertionpnt(2) = 0
            '///
            '/// Orden de D en agujas y el resto es diferente
            '///
            If Mid(poste(iposte).tipo, 15) = semi_eje_aguj Or Mid(poste(iposte).tipo, 15) = eje_aguj Or poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj Then
                If Mid(poste(iposte).tipo, 15) = eje_aguj Or poste(iposte).tipo = eje_aguj And poste(iposte).descentramiento_2mens > 0 Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj Then
                    descen = Round(poste(iposte).descentramiento_2mens - 400, 0)
                ElseIf Mid(poste(iposte).tipo, 15) = eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj And poste(iposte).descentramiento_2mens < 0 Then
                    descen = Round(poste(iposte).descentramiento_2mens + 400, 0)
                Else
                    descen = Round(poste(iposte).descentramiento_2mens, 0)
                End If
                If descen >= 0 Then
                cadena_desc = cadena_ruta & "Desaxement1.dwg"
                ElseIf descen < 0 Then
                cadena_desc = cadena_ruta & "Desaxement2.dwg"
                End If
                Set bloque_desc = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_desc, escala_des, escala_des, escala_des, alfa_desc)
                bloque_desc.Layer = "E-DESCENTRAMIENTOS"
                Atributos = bloque_desc.GetAttributes
                Atributos(0).TextString = descen
                insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_desc - PI / 2) + dist_poste * Cos(alfa_desc - PI)
                insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_desc - PI / 2) + dist_poste * sin(alfa_desc - PI)
                insertionpnt(2) = 0
                If poste(iposte).descentramiento >= 0 Then
                cadena_desc = cadena_ruta & "Desaxement1.dwg"
                End If
                If poste(iposte).descentramiento < 0 Then
                cadena_desc = cadena_ruta & "Desaxement2.dwg"
                End If
                Set bloque_desc = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_desc, escala_des, escala_des, escala_des, alfa_desc)
                bloque_desc.Layer = "E-DESCENTRAMIENTOS"
                Atributos = bloque_desc.GetAttributes
                Atributos(0).TextString = Abs(poste(iposte).descentramiento)
            Else
            If poste(iposte).descentramiento >= 0 Then
            cadena_desc = cadena_ruta & "Desaxement1.dwg"
            End If
            If poste(iposte).descentramiento < 0 Then
            cadena_desc = cadena_ruta & "Desaxement2.dwg"
            End If
            Set bloque_desc = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_desc, escala_des, escala_des, escala_des, alfa_desc)
            bloque_desc.Layer = "E-DESCENTRAMIENTOS"
            Atributos = bloque_desc.GetAttributes
            'If poste(iposte).descentramiento >= 0 Then
                Atributos(0).TextString = Abs(poste(iposte).descentramiento)
            'End If
            'If poste(iposte).descentramiento < 0 Then
                'Atributos(0).TextString = -poste(iposte).descentramiento
            'End If

            
            insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_desc - PI / 2) + dist_poste * Cos(alfa_desc - PI)
            insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_desc - PI / 2) + dist_poste * sin(alfa_desc - PI)
            insertionpnt(2) = 0
            If poste(iposte).descentramiento_2mens >= 0 Then
            cadena_desc = cadena_ruta & "Desaxement1.dwg"
            End If
            If poste(iposte).descentramiento_2mens < 0 Then
            cadena_desc = cadena_ruta & "Desaxement2.dwg"
            End If
            Set bloque_desc = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_desc, escala_des, escala_des, escala_des, alfa_desc)
            bloque_desc.Layer = "E-DESCENTRAMIENTOS"
            Atributos = bloque_desc.GetAttributes
            'If poste(iposte).descentramiento_2mens >= 0 Then
                Atributos(0).TextString = Abs(poste(iposte).descentramiento_2mens)
            'End If
            'If poste(iposte).descentramiento_2mens < 0 Then
                'Atributos(0).TextString = -poste(iposte).descentramiento_2mens
            'End If
            End If
        End If
        If poste(iposte).mensula2a = False Then
            insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_desc - PI / 2)
            insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_desc - PI / 2)
            insertionpnt(2) = 0
            If poste(iposte).descentramiento >= 0 Then
            cadena_desc = cadena_ruta & "Desaxement1.dwg"
            End If
            If poste(iposte).descentramiento < 0 Then
            cadena_desc = cadena_ruta & "Desaxement2.dwg"
            End If
            Set bloque_desc = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_desc, escala_des, escala_des, escala_des, alfa_desc)
            bloque_desc.Layer = "E-DESCENTRAMIENTOS"
            Atributos = bloque_desc.GetAttributes
            'If poste(iposte).descentramiento >= 0 Then
                Atributos(0).TextString = Abs(poste(iposte).descentramiento)
            'End If
            'If poste(iposte).descentramiento < 0 Then
                'Atributos(0).TextString = -poste(iposte).descentramiento
            'End If
        End If
Next
End Sub
Sub dibujar_implantacion(cadena_ruta As String, HDC As Boolean)
Dim bloque_impla As AcadBlockReference
'Dim impla As Double
Dim Atributos As Variant
Dim insertionpnt(0 To 2) As Double
Dim dist_eje As Single, dist_poste As Single
Dim alfa_impla As Double
escala_imp = 1
dist_eje = 19.5
dist_poste = 4
cont = 0
cont_est = 0
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-IMPLANTACION")
ancho_poste = 0.46
For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_imp = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
        'escala_imp = 1
        'cont = 0
    'End If
    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 17)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            End If
        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 17)
        End If
        dist_poste = 4
    ElseIf HDC = True Then
        If cont < 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril - 2)
            dist_poste = 4
        ElseIf cont >= 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril - 2)
            dist_poste = 2
        End If
    End If

    If poste(iposte).implantacion <> dist_carril_poste And poste(iposte).implantacion <> 0 Then
        alfa_impla = poste(iposte).anguloeje * PI / 180
        insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_impla - PI / 2) + dist_poste * Cos(alfa_impla - PI)
        insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_impla - PI / 2) + dist_poste * sin(alfa_impla - PI)
        insertionpnt(2) = 0
        Set bloque_impla = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Implantation.dwg", escala_imp, escala_imp, escala_imp, alfa_impla)
        bloque_impla.Layer = "E-IMPLANTACION"
        Atributos = bloque_impla.GetAttributes
        Atributos(0).TextString = "i=" & Round(poste(iposte).implantacion, 2) & " m"
        'num_impla = num_impla + 1
    End If
Next
End Sub
Sub dibujar_flechas(cadena_ruta As String, HDC As Boolean)
Dim bloque_flecha As AcadBlockReference
Dim flecha As Double
Dim Atributos As Variant
Dim radio1 As Double, radio2 As Double
Dim insertionpnt(0 To 2) As Double
Dim dist_eje As Single
Dim alfa_flecha As Double
escala_fle = 1
ancho_poste = 0.46
cont = 0
cont_est = 0
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-FLECHAS")
For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_fle = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont = 3 Then
        'escala_fle = 1
        'cont = 0
    'End If

    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            End If

        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
        End If
    
    ElseIf HDC = True Then
        If cont < 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
        ElseIf cont >= 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
        End If
    End If

    radio1 = Abs(poste(iposte).radio)
    radio2 = Abs(poste(iposte).radio)
    If radio1 = 0 And radio2 = 0 Then
        GoTo finflecha
    ElseIf radio1 = 0 Then
        radio1 = 7500
    ElseIf radio2 = 0 Then
        radio2 = 7500
    End If
    If poste(iposte).radio <> 0 Then
        flecha = Round(1000 * (poste(iposte).vano_post ^ 2) / (8 * (radio1 + radio2) / 2), 0)
        alfa_flecha = poste(iposte).alfa_vano
        If Abs(flecha) < 20 Then
            GoTo finflecha
        End If
        insertionpnt(0) = poste(iposte).pk_vano_postx + dist_eje * Cos(alfa_flecha + PI / 2)
        insertionpnt(1) = poste(iposte).pk_vano_posty + dist_eje * sin(alfa_flecha + PI / 2)
        insertionpnt(2) = 0
        If poste(iposte).radio + poste(iposte - 1).radio < 0 Then
            Set bloque_flecha = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "flèche2.dwg", escala_fle, escala_fle, escala_fle, alfa_flecha)
            bloque_flecha.Layer = "E-FLECHAS"
        Else
            Set bloque_flecha = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "flèche1.dwg", escala_fle, escala_fle, escala_fle, alfa_flecha)
            bloque_flecha.Layer = "E-FLECHAS"
        End If
        Atributos = bloque_flecha.GetAttributes
        Atributos(0).TextString = "f=" & Abs(flecha)
    End If
finflecha:
Next
End Sub
Sub dibujar_vanos(cadena_ruta As String, HDC As Boolean)
Dim alfa_vano As Double
Dim cadena_vano As String
Dim bloque_vano As AcadBlockReference
Dim dist_eje As Double
Dim insertionpnt(0 To 2) As Double
Dim Atributos As Variant
ancho_poste = 0.46
escala_vano = 1
cont = 0
cont_est = 0
ipk = 0
cadena_vano = cadena_ruta & "vano.dwg"
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-VANOS")
For iposte = 1 To num_postes_total
    alfa_vano = poste(iposte).alfa_vano
    ' Variar escala
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_vano = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont = 3 Then
        'escala_vano = 1
        'cont = 0
    'End If
    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
    
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            End If
        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
        End If
    
        ElseIf HDC = True Then
            If cont < 1 Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            ElseIf cont >= 1 Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 3)
            End If
    End If


    insertionpnt(0) = poste(iposte).pk_vano_postx + dist_eje * Cos(alfa_vano + PI / 2)
    insertionpnt(1) = poste(iposte).pk_vano_posty + dist_eje * sin(alfa_vano + PI / 2)
    insertionpnt(2) = 0

    
    
    Set bloque_vano = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_vano, escala_vano, escala_vano, escala_vano, alfa_vano)
    If bloque_vano.HasAttributes = True Then
        Atributos = bloque_vano.GetAttributes
        Atributos(0).TextString = poste(iposte).vano_post
    End If
    bloque_vano.Layer = "E-VANOS"
  Next iposte
End Sub
Sub dibujar_alturaHC(cadena_ruta As String, HDC As Boolean)
Dim i As Integer
Dim alt As Single
Dim cadena_HC, cadena_gradi_crec, cadena_gradi_plano, cadena_gradi_decrec As String
Dim bloque_HC, bloque_gradi As AcadBlockReference
Dim dist_HC_eje, alfa_HC, alfa_gradi, dist_gradi_HC, gradiente As Double
Dim insertionpnt(0 To 2) As Double
Dim Atributos As Variant
Dim ceros As String
escala_HC = 1
dist_HC_eje = 18
dist_gradi_eje = 9
cont = 0
cont_est = 0
ancho_poste = 0.46
cadena_HC = cadena_ruta & "Hauteur fil de contact.dwg"
cadena_gradi_crec = cadena_ruta & "Pente1.dwg"
cadena_gradi_plano = cadena_ruta & "Pente3.dwg"
cadena_gradi_decrec = cadena_ruta & "Pente2.dwg"
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-ALTURA HC")
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-GRADIENTE HC")
poste(0).altura_HC = 5.5

For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_HC = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte - 1).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
        'escala_HC = 1
        'cont = 0
    'End If
    If HDC = False Then
    
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_HC_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 11)
            ElseIf estacion(ipk).lado = False Then
                dist_HC_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 10.5)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_HC_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 15)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_HC_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 11)
            ElseIf estacion(ipk).lado = False Then
                dist_HC_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 10.5)
            End If
        Else
            dist_HC_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 15)
        End If
    ElseIf HDC = True Then
        If cont < 1 Then
            dist_HC_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 15)
        ElseIf cont >= 1 Then
            dist_HC_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
        End If
    End If
    If poste(iposte).altura_HC <> alt_nom Then
        alfa_HC = poste(iposte).anguloeje * PI / 180
        insertionpnt(0) = poste(iposte).pk_coordx + dist_HC_eje * Cos(alfa_HC - PI / 2)
        insertionpnt(1) = poste(iposte).pk_coordy + dist_HC_eje * sin(alfa_HC - PI / 2)
        insertionpnt(2) = 0
        Set bloque_HC = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_HC, escala_HC, escala_HC, escala_HC, alfa_HC)
        bloque_HC.Layer = "E-ALTURA HC"
        If bloque_HC.HasAttributes = True Then
            Atributos = bloque_HC.GetAttributes
            Set bloque_HC = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_HC, escala_HC, escala_HC, escala_HC, alfa_HC)
            Atributos(0).TextString = poste(iposte).altura_HC
            bloque_HC.Layer = "E-ALTURA HC"
        End If
    End If
Next
escala_HC = 1
cont = 0
cont_est = 0
For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_HC = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont = 3 Then
        'escala_HC = 1
        'cont = 0
    'End If
    If HDC = False Then
            '///
            '/// Verificación del lado a implantar los datos en estacion
            '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_gradi_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 12)
            ElseIf estacion(ipk).lado = False Then
                dist_gradi_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 12)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_gradi_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_gradi_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 12)
            ElseIf estacion(ipk).lado = False Then
                dist_gradi_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 12)
            End If
        Else
            dist_gradi_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
        End If
    ElseIf HDC = True Then
        If cont < 1 Then
            dist_gradi_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 15)
        ElseIf cont >= 1 Then
            dist_gradi_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
        End If
    End If

        If poste(iposte).altura_HC <> poste(iposte + 1).altura_HC Then
            alfa_gradi = poste(iposte).alfa_vano
            gradiente = Round(1000 * (poste(iposte + 1).altura_HC - poste(iposte).altura_HC) / poste(iposte).vano_post, 2)
            insertionpnt(0) = poste(iposte).pk_vano_postx + dist_gradi_eje * Cos(alfa_gradi - PI / 2)
            insertionpnt(1) = poste(iposte).pk_vano_posty + dist_gradi_eje * sin(alfa_gradi - PI / 2)
            insertionpnt(2) = 0
            If gradiente > 0 Then
                Set bloque_gradi = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_gradi_crec, escala_HC, escala_HC, escala_HC, alfa_gradi)
                bloque_gradi.Layer = "E-GRADIENTE HC"
            End If
            If gradiente = 0 Then
                Set bloque_gradi = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_gradi_plano, escala_HC, escala_HC, escala_HC, alfa_gradi)
                bloque_gradi.Layer = "E-GRADIENTE HC"
            End If
            If gradiente < 0 Then
                Set bloque_gradi = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_gradi_decrec, escala_HC, escala_HC, escala_HC, alfa_gradi)
                bloque_gradi.Layer = "E-GRADIENTE HC"
            End If
            If bloque_gradi.HasAttributes = True Then
                Atributos = bloque_gradi.GetAttributes
                Atributos(0).TextString = Round(gradiente, 2)
                bloque_HC.Layer = "E-GRADIENTE HC"
            End If
        End If
Next
End Sub
Sub dibujar_etiquetas(cadena_ruta As String, HDC As Boolean)
Dim cadena_etiq As String
Dim bloque_etiq As AcadBlockReference
Dim dist_eje, alfa_etiq As Double
Dim insertionpnt(0 To 2) As Double
Dim Atributos As Variant
ancho_poste = 0.46
escala_eti = 1
cont = 0
cont_est = 0
'cadena_etiq = cadena_ruta & "Rond voie principale.dwg"
cadena_etiq = cadena_ruta & "Rond voie secondaire.dwg"
cadena_anc_cdpa = cadena_ruta & "Implantation.dwg"
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-ETIQUETAS POSTES")
For iposte = 1 To num_postes_total

    ipoli = 1

    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_eti = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte - 1).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
        'escala_eti = 1
        'cont = 0
    'End If
    
    
    If HDC = False Then
           
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 0 Then
            dist_eje = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6
            
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4
            End If
        End If
    
    ElseIf HDC = True Then
        If poste(iposte).lado = "G" And cont < 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            dist_cdpa = 35
        ElseIf poste(iposte).lado = "D" And cont < 1 Then
            dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 6)
            dist_cdpa = 35
        ElseIf poste(iposte).lado = "G" And cont >= 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            dist_cdpa = 20
        ElseIf poste(iposte).lado = "D" And cont >= 1 Then
            dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 4)
            dist_cdpa = 20
        End If
    End If

    
    alfa_etiq = poste(iposte).anguloeje * PI / 180
    insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_etiq + PI / 2)
    insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_etiq + PI / 2)
    insertionpnt(2) = 0
    Set bloque_etiq = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_etiq, escala_eti, escala_eti, escala_eti, alfa_etiq)
    If bloque_etiq.HasAttributes = True Then
        Atributos = bloque_etiq.GetAttributes
        Atributos(0).TextString = poste(iposte).etiq_1
        'Atributos(1).TextString = poste(iposte).etiq_2
    End If
    bloque_etiq.Layer = "E-ETIQUETAS POSTES"
    
                '/// Añadir comentario anclajes cdpa y feeder
    If poste(iposte).anc_cdpa <> "" Then
            
        alfa_etiq = (poste(iposte).anguloeje + 90) * PI / 180
        insertionpnt(0) = poste(iposte).pk_coordx + dist_cdpa * Cos(alfa_etiq)
        insertionpnt(1) = poste(iposte).pk_coordy + dist_cdpa * sin(alfa_etiq)
        insertionpnt(2) = 0
        Set bloque_cdpa = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_anc_cdpa, escala_eti, escala_eti, escala_eti, alfa_etiq)
        Atributos = bloque_cdpa.GetAttributes
        Atributos(0).TextString = poste(iposte).anc_cdpa
        bloque_cdpa.Layer = "E-ETIQUETAS POSTES"

    End If
    
    
    
Next iposte
End Sub

Sub dibujar_postes(cadena_ruta As String, HDC As Boolean)
Dim alfa As Double
Dim bloque As AcadBlockReference
Dim circulo As AcadCircle
Dim linea As AcadLine
Dim square As AcadPolyline
Dim arco As AcadArc
Dim ini_linea(0 To 2) As Double, ini_aux(0 To 2) As Double, PA(0 To 2) As Double
Dim ini_linea2(0 To 2) As Double, fin_aux(0 To 2) As Double
Dim centro(0 To 2) As Double
Dim cadena_poste As String
Dim insertionpnt(0 To 2) As Double
Dim radio As Double
Dim accapa As AcadLayer
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-POSTES")
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers("E-POSTES")
accapa.Color = acRed
ancho_poste = 0.46
ancho_cuadrado = 1
escala_poste = 1
cont = 0
For iposte = 1 To num_postes_total

    Call txt.progress("1", "1", "Dibujado de postes", poste(iposte).pk_global - poste(1).pk_global, poste(num_postes_total).pk_global - poste(1).pk_global)
    '///
    '/// Variar la escala en estaciones
    '///
    'If poste(iposte).tipo = anc_sla_con And cont < 3 Then
        'escala_poste = 0.5
        'cont = cont + 1
    'ElseIf poste(iposte - 1).tipo = anc_sla_con And poste(iposte).tipo = "" And cont = 3 Then
        'escala_poste = 1
        'cont = 0
    'End If
    
    alfa = poste(iposte).anguloeje * PI / 180
    
    Select Case poste(iposte).lado
        Case "D"
            alfa_poste = (poste(iposte).anguloeje - 90) * PI / 180
        Case "G"
            alfa_poste = (poste(iposte).anguloeje + 90) * PI / 180
    End Select
    If poste(iposte).tipo = eje_sla Then
        ancho_mens = 1.6
    Else
        ancho_mens = 1
    End If
    If poste(iposte).tunel = False Then
        eje_poste = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril
        long_ancl = 7
    Else
        eje_poste = 1.4
        long_ancl = 3
    End If
    '///
    '///Añadir la distancia que falta del poste hasta llegar al descentramiento
    '///
    If poste(iposte).descentramiento_2mens = 0 And (poste(iposte).tipo <> semi_eje_aguj And poste(iposte).tipo <> eje_aguj And poste(iposte).tipo <> eje_pf & " + " & eje_aguj And poste(iposte).tipo <> anc_pf & " + " & eje_aguj And poste(iposte).tipo <> eje_aguj & " + " & anc_aguj) Then
        ini_aux(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2)
        ini_aux(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2)
        ini_aux(2) = 0
        centro(0) = poste(iposte).pk_coordx + eje_poste * Cos(alfa_poste)
        centro(1) = poste(iposte).pk_coordy + eje_poste * sin(alfa_poste)
        centro(2) = 0
        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_aux, centro)
        linea.Layer = "E-POSTES"
        Call cuadrado(centro, alfa, ancho_cuadrado * escala_poste, ancho_cuadrado * escala_poste, "E-POSTES")
        'Set circul = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(centro, radio)
        'circul.Layer = "E-POSTES"
        'intersectionpoint = circul.IntersectWith(linea, acExtendNone)
        intersectionpoint = polinea.IntersectWith(linea, acExtendNone)
        linea.Delete
        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_aux, intersectionpoint)
        linea.Layer = "E-POSTES"
    ElseIf poste(iposte).descentramiento_2mens <> 0 Or poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = eje_pf & " + " & eje_aguj Or poste(iposte).tipo = anc_pf & " + " & eje_aguj Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj Then
            ini_aux(0) = poste(iposte).pk_coordx
            ini_aux(1) = poste(iposte).pk_coordy
            ini_aux(2) = 0
            centro(0) = poste(iposte).pk_coordx + eje_poste * Cos(alfa_poste)
            centro(1) = poste(iposte).pk_coordy + eje_poste * sin(alfa_poste)
            centro(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_aux, centro)
            linea.Layer = "E-POSTES"
            Call cuadrado(centro, alfa, ancho_cuadrado * escala_poste, ancho_cuadrado * escala_poste, "E-POSTES")
            'Set circul = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(centro, radio)
            'circul.Layer = "E-POSTES"
            'intersectionpoint = circul.IntersectWith(linea, acExtendNone)
            intersectionpoint = polinea.IntersectWith(linea, acExtendNone)
            linea.Delete
            ini_aux(0) = intersectionpoint(0) - ancho_mens / 2 * Cos(alfa)
            ini_aux(1) = intersectionpoint(1) - ancho_mens / 2 * sin(alfa)
            ini_aux(2) = intersectionpoint(2)
            fin_aux(0) = intersectionpoint(0) + ancho_mens / 2 * Cos(alfa)
            fin_aux(1) = intersectionpoint(1) + ancho_mens / 2 * sin(alfa)
            fin_aux(2) = intersectionpoint(2)
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_aux, fin_aux)
            linea.Layer = "E-POSTES"
            If poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = eje_pf & " + " & eje_aguj Or poste(iposte).tipo = anc_pf & " + " & eje_aguj Then
                ini_linea(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento_2mens / 1000 * Cos(alfa + PI / 2) + ancho_mens / 2 * Cos(alfa)
                ini_linea(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento_2mens / 1000 * sin(alfa + PI / 2) + ancho_mens / 2 * sin(alfa)
                ini_linea(2) = 0
                ini_linea2(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
                ini_linea2(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
                ini_linea2(2) = 0
            Else
                ini_linea(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) + ancho_mens / 2 * Cos(alfa)
                ini_linea(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) + ancho_mens / 2 * sin(alfa)
                ini_linea(2) = 0
                ini_linea2(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento_2mens / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
                ini_linea2(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento_2mens / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
                ini_linea2(2) = 0
            End If
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea2, ini_aux)
            linea.Layer = "E-POSTES"
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, fin_aux)
            linea.Layer = "E-POSTES"
    End If
    '///
    '///añadir el anclaje
    '///
    If poste(iposte).tipo = anc_sm_con Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj Or poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_pf Or poste(iposte).anc_cdpa = "Anc. Feeder Alim." Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj Or _
    poste(iposte).tipo = anc_aguj Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte).tipo = anc_pf & " + " & anc_aguj Or poste(iposte).tipo = anc_pf & " + " & eje_aguj Or poste(iposte).tipo = anc_pf & " + " & semi_eje_aguj And poste(iposte - 1).tipo <> eje_pf Then
                If poste(iposte + 1).tipo = semi_eje_sm Or poste(iposte + 1).tipo = semi_eje_sla Or poste(iposte + 1).tipo = anc_sla_con & " + " & semi_eje_aguj Or (poste(iposte - 1).tipo = anc_sla_con And poste(iposte).anc_cdpa = "Anc. Feeder Alim.") Or _
                poste(iposte + 1).tipo = eje_pf Or (poste(iposte + 1).tipo = semi_eje_aguj Or poste(iposte + 1).tipo = eje_pf & " + " & semi_eje_aguj Or poste(iposte + 1).tipo = eje_pf & " + " & eje_aguj Or poste(iposte + 1).tipo = anc_pf & " + " & semi_eje_aguj) Then

                    ini_aux(0) = centro(0) - long_ancl * Cos(alfa)
                    ini_aux(1) = centro(1) - long_ancl * sin(alfa)
                    ini_aux(2) = 0
                    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
                    linea.Layer = "E-POSTES"
                    Call anclaje(ini_aux, alfa_poste, escala_poste)
                    If poste(iposte).anc_cdpa = "Anc. CdPA et Feeder" Or poste(iposte).anc_cdpa = "Anc. Feeder Alim." Then
                        ini_aux(0) = centro(0) + long_ancl * Cos(alfa)
                        ini_aux(1) = centro(1) + long_ancl * sin(alfa)
                        ini_aux(2) = 0
                        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
                        linea.Layer = "E-POSTES"
                        Call anclaje(ini_aux, alfa_poste, escala_poste)
                    End If
                ElseIf poste(iposte - 1).tipo = semi_eje_sm Or poste(iposte - 1).tipo = semi_eje_sla Or poste(iposte - 1).tipo = semi_eje_sla & " + " & anc_aguj Or (poste(iposte + 1).tipo = anc_sla_con And poste(iposte).anc_cdpa = "Anc. Feeder Alim.") Or _
                poste(iposte - 1).tipo = eje_pf Or poste(iposte - 1).tipo = semi_eje_aguj Or poste(iposte - 1).tipo = eje_pf & " + " & eje_aguj Or poste(iposte - 1).tipo = eje_pf & " + " & semi_eje_aguj Or poste(iposte - 1).tipo = anc_pf & " + " & semi_eje_aguj Then

                    ini_aux(0) = centro(0) + long_ancl * Cos(alfa)
                    ini_aux(1) = centro(1) + long_ancl * sin(alfa)
                    ini_aux(2) = 0
                    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
                    linea.Layer = "E-POSTES"
                    Call anclaje(ini_aux, alfa_poste, escala_poste)
                    If poste(iposte).anc_cdpa = "Anc. CdPA et Feeder" Or poste(iposte).anc_cdpa = "Anc. Feeder Alim." Then
                        ini_aux(0) = centro(0) - long_ancl * Cos(alfa)
                        ini_aux(1) = centro(1) - long_ancl * sin(alfa)
                        ini_aux(2) = 0
                        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
                        linea.Layer = "E-POSTES"
                        Call anclaje(ini_aux, alfa_poste, escala_poste)
                    End If
                End If
                    
    ElseIf poste(iposte).tipo = eje_pf And poste(iposte).tunel = True Then
                    ini_aux(0) = centro(0) - long_ancl * Cos(alfa_poste + 3 * PI / 4)
                    ini_aux(1) = centro(1) - long_ancl * sin(alfa_poste + 3 * PI / 4)
                    ini_aux(2) = 0
                    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
                    linea.Layer = "E-POSTES"
                    Call anclaje(ini_aux, alfa_poste, escala_poste)
                    ini_aux(0) = centro(0) + long_ancl * Cos(alfa_poste + PI / 4)
                    ini_aux(1) = centro(1) + long_ancl * sin(alfa_poste + PI / 4)
                    ini_aux(2) = 0
                    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
                    linea.Layer = "E-POSTES"
                    Call anclaje(ini_aux, alfa_poste, escala_poste)
    
    End If
    If poste(iposte).tunel = True Then
            ini_linea(0) = centro(0) + ancho_cuadrado / 2 * Cos(alfa_poste)
            ini_linea(1) = centro(1) + ancho_cuadrado / 2 * sin(alfa_poste)
            ini_linea(2) = 0
            ini_linea2(0) = centro(0) - ancho_cuadrado / 2 * Cos(alfa_poste)
            ini_linea2(1) = centro(1) - ancho_cuadrado / 2 * sin(alfa_poste)
            ini_linea2(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, ini_linea2)
            linea.Layer = "E-POSTES"
            ini_linea(0) = centro(0) + ancho_cuadrado / 2 * Cos(alfa_poste - PI / 2)
            ini_linea(1) = centro(1) + ancho_cuadrado / 2 * sin(alfa_poste - PI / 2)
            ini_linea(2) = 0
            ini_linea2(0) = centro(0) - ancho_cuadrado / 2 * Cos(alfa_poste - PI / 2)
            ini_linea2(1) = centro(1) - ancho_cuadrado / 2 * sin(alfa_poste - PI / 2)
            ini_linea2(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, ini_linea2)
            linea.Layer = "E-POSTES"
    End If
Next iposte

GetObject(, "Autocad.Application").ActiveDocument.Regen acActiveViewport

End Sub
Sub cuadrado(ini_aux, alfa, ancho_cim, alto_cim, capa)
Dim centro(0 To 2) As Double
Dim cuadro(0 To 14) As Double
    cuadro(0) = ini_aux(0) + ancho_cim / 2 * Cos(alfa) - ancho_cim / 2 * Cos(alfa - PI / 2)
    cuadro(1) = ini_aux(1) + ancho_cim / 2 * sin(alfa) - ancho_cim / 2 * sin(alfa - PI / 2)
    cuadro(2) = 0
    cuadro(3) = cuadro(0) - alto_cim * Cos(alfa)
    cuadro(4) = cuadro(1) - alto_cim * sin(alfa)
    cuadro(5) = 0
    cuadro(6) = cuadro(3) + ancho_cim * Cos(alfa - PI / 2)
    cuadro(7) = cuadro(4) + ancho_cim * sin(alfa - PI / 2)
    cuadro(8) = 0
    cuadro(9) = cuadro(6) + alto_cim * Cos(alfa)
    cuadro(10) = cuadro(7) + alto_cim * sin(alfa)
    cuadro(11) = 0
    cuadro(12) = cuadro(0)
    cuadro(13) = cuadro(1)
    cuadro(14) = 0
    Set polinea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(cuadro)
    polinea.Layer = capa
         
    If capa = "E-HILO CONTACTO" Then


        ' This example creates an associative gradient hatch in model space.
        Dim hatchObj As AcadHatch
        Dim patternName As String
        Dim PatternType As Long
        Dim bAssociativity As Boolean
        ' Define the hatch
        patternName = "CYLINDER"
        PatternType = acPreDefinedGradient '0
        bAssociativity = True
        ' Create the associative Hatch object in model space
        Set hatchObj = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddHatch(PatternType, patternName, bAssociativity, acGradientObject)
        Dim col1 As AcadAcCmColor, col2 As AcadAcCmColor
        Select Case Left(Application.Version, 2)
        Case 18
            Set col1 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.18")
            Set col2 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.18")
        Case 17
            Set col1 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.17")
            Set col2 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.17")
        Case 14
            Set col1 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.18")
            Set col2 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.18")
        End Select

        
        
        'Set col1 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.17")
     ' modificar segun versiones de autocad (2007-2009 -> 17)
        'Set col2 = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.17")
        Call col1.SetRGB(255, 0, 0)
        Call col2.SetRGB(255, 0, 0)
        hatchObj.GradientColor1 = col1
        hatchObj.GradientColor2 = col2
        ' Create the outer boundary for the hatch (a circle)
        Dim outerLoop(0 To 0) As AcadEntity
    '*********************************************************************************
    ' Creamos una region con los objectos que queramos
    '*********************************************************************************
        ' Create the region
        Dim cuadrado(0) As AcadEntity
        Set cuadrado(0) = polinea
    
        Dim regionObj As Variant
        regionObj = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddRegion(cuadrado)
        
       
        '*********************************************************************************
        ' Fin de la creacion de la region
        ' le pasamos la region creada paracrear el hatch con ella
        '*********************************************************************************
        Set outerLoop(0) = regionObj(0)
        ' Append the outerboundary to the hatch object, and display the hatch
        hatchObj.AppendOuterLoop (outerLoop)
        hatchObj.Evaluate
        hatchObj.Layer = "E-HILO CONTACTO"
        regionObj(0).Delete
        
    End If


End Sub
Sub anclaje(ini_aux, alfa, ancho_cim)
Dim linea As AcadLine
Dim cuadro(0 To 14) As Double
Dim lin1(0 To 2) As Double, lin2(0 To 2) As Double

If poste(iposte).tunel = False Then
    cuadro(0) = ini_aux(0) + ancho_cim / 2 * Cos(alfa) - ancho_cim / 2 * Cos(alfa - PI / 2)
    cuadro(1) = ini_aux(1) + ancho_cim / 2 * sin(alfa) - ancho_cim / 2 * sin(alfa - PI / 2)
    cuadro(2) = 0
    cuadro(3) = cuadro(0) - ancho_cim * Cos(alfa)
    cuadro(4) = cuadro(1) - ancho_cim * sin(alfa)
    cuadro(5) = 0
    cuadro(6) = cuadro(3) + ancho_cim * Cos(alfa - PI / 2)
    cuadro(7) = cuadro(4) + ancho_cim * sin(alfa - PI / 2)
    cuadro(8) = 0
    cuadro(9) = cuadro(6) + ancho_cim * Cos(alfa)
    cuadro(10) = cuadro(7) + ancho_cim * sin(alfa)
    cuadro(11) = 0
    cuadro(12) = cuadro(0)
    cuadro(13) = cuadro(1)
    cuadro(14) = 0
    Set polinea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(cuadro)
    polinea.Layer = "E-POSTES"
ElseIf poste(iposte).tipo = eje_pf Or poste(iposte).tipo = anc_sm_con Or poste(iposte).tipo = anc_sla_con Then
    lin1(0) = ini_aux(0) + ancho_cim / 2 * Cos(alfa)
    lin1(1) = ini_aux(1) + ancho_cim / 2 * sin(alfa)
    lin1(2) = 0
    lin2(0) = ini_aux(0) - ancho_cim / 2 * Cos(alfa)
    lin2(1) = ini_aux(1) - ancho_cim / 2 * sin(alfa)
    lin2(2) = 0
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(lin1, lin2)
    linea.Layer = "E-POSTES"
End If

End Sub

Sub dibujar_conexion(cadena_ruta As String, HDC As Boolean)
Dim accapa As AcadLayer
Dim bloque_con As AcadBlockReference
Dim circulo As AcadCircle
Dim arco As AcadArc, arco1 As AcadArc
Dim prot As Double
Dim Atributos As Variant
Dim insertionpnt(0 To 2) As Double, PA(0 To 2) As Double, PB(0 To 2) As Double, PA1(0 To 2) As Double, PB1(0 To 2) As Double
Dim insertionpnt1(0 To 2) As Double, center(0 To 2) As Double, insertionpnt2(0 To 2) As Double
Dim inter() As Double
Dim dist_eje As Single
Dim alfa_prot As Double
Dim num_prot As Integer, conta As Integer
Dim n_poste1 As Integer, n_poste2 As Integer
escala_pro = 1
dist_eje = 14.5
ancho_poste = 0.46
cont = 0
conta = 0
cont_est = 0
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-CONEXION")
accapa.Color = acCyan
num_impla = 0
dist_con = 7
iposte = 3
For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And conta < 3 And iposte <> 1 And iposte <> 6 Then
        escala_pro = 0.5
        conta = 2 ' solo para vias secundarias
        'conta = conta + 1
    'ElseIf (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And conta = 3 Then
        'escala_pro = 1
        'conta = 0
    'End If
    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 12)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
            End If
        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
        End If
    
        ElseIf HDC = True Then
            If conta < 1 Then
                dist_eje = 9
            ElseIf conta >= 1 Then
                dist_eje = 5
            End If
    End If

    If iposte = 1 Then
        cont = 1
    ElseIf iposte <= 5 And (poste(iposte).tipo = semi_eje_sla Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj) And cont <= 2 Then
        'cont = cont + 1
    ElseIf iposte <= 5 And cont >= 4 And poste(iposte - 1).tipo = semi_eje_sla And poste(iposte - 2).tipo = eje_sla Then
        cont = 1
    ElseIf iposte > 5 And (poste(iposte).tipo = semi_eje_sla Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj) And cont <= 3 Then
        cont = cont + 1
    ElseIf cont = 4 And poste(iposte - 1).tipo = semi_eje_sla And poste(iposte - 2).tipo = eje_sla Then
        cont = 1
    End If
    
    If poste(iposte).conexion <> "" Then
        alfa_vano = poste(iposte).alfa_vano
        insertionpnt2(0) = poste(iposte).pk_vano_postx + dist_eje * Cos(alfa_vano - PI / 2)
        insertionpnt2(1) = poste(iposte).pk_vano_posty + dist_eje * sin(alfa_vano - PI / 2)
        insertionpnt2(2) = 0
        Set bloque_con = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt2, cadena_ruta & "Conexion.dwg", escala_pro, escala_pro, escala_pro, alfa_vano)
        bloque_con.Layer = "E-CONEXION"
        Atributos = bloque_con.GetAttributes
        Atributos(0).TextString = poste(iposte).conexion
        num_impla = num_impla + 1
        
        '///
        '/// Añadir conexión rep02
        '///
        If (poste(iposte + 1).tipo = eje_pf Or poste(iposte + 1).tipo = eje_pf & " + " & eje_aguj Or poste(iposte + 1).tipo = eje_pf & " + " & semi_eje_aguj) Or _
        poste(iposte).conexion = "667001-02" Then
            alfa = poste(iposte).anguloeje * PI / 180
            PA(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2)
            PA(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2)
            PA(2) = 0
            alfa = poste(iposte + 1).anguloeje * PI / 180
            PB(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento / 1000 * Cos(alfa + PI / 2)
            PB(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento / 1000 * sin(alfa + PI / 2)
            PB(2) = 0
            'ang_plus = cuadrante(PA, PB)
            alpa = Atn((PB(1) - PA(1)) / (PB(0) - PA(0)))
            insertionpnt(0) = PA(0) + Cos(alpa) * 3.375
            insertionpnt(1) = PA(1) + sin(alpa) * 3.375
            insertionpnt(2) = 0
            Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(insertionpnt, 0.5, alpa, alpa - 3.14)
            arco.Layer = "E-CONEXION"
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(insertionpnt, insertionpnt2)
            linea.Layer = "E-CONEXION"
        '///
        '/// Añadir conexión rep90
        '///
        ElseIf poste(iposte).conexion = "667001-90" Then
            If ((poste(iposte + 1).tipo = semi_eje_sla Or poste(iposte + 1).tipo = semi_eje_sla & " + " & anc_aguj) And (cont = 4 Or cont = 1)) Then
                n_poste1 = iposte + 1
                n_poste2 = iposte
                pos = -1
                desc1 = poste(n_poste1).descentramiento_2mens / 1000
                desc2 = poste(n_poste2).descentramiento_2mens / 1000
            Else
                n_poste1 = iposte
                n_poste2 = iposte + 1
                pos = 1
                desc1 = poste(n_poste1).descentramiento / 1000
                desc2 = poste(n_poste2).descentramiento / 1000
            End If

            If poste(n_poste1).tunel = False Then
                eje_poste = poste(n_poste1).implantacion + (ancho_via / 2) + (ancho_poste) + ancho_carril + ancho_carril + 1.7
            Else
                eje_poste = 1.4 + 2
            End If
            If poste(n_poste2).tunel = False Then
                eje_poste1 = poste(n_poste2).implantacion + (ancho_via / 2) + (ancho_poste) + ancho_carril + ancho_carril + 1.7
            Else
                eje_poste1 = 1.4 + 2
            End If
            'eje_poste = -poste(iposte).descentramiento / 1000 + poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste) + ancho_carril + 2
            Select Case poste(n_poste1).lado
                Case "D"
                    alfa_poste = (poste(n_poste1).anguloeje - 90) * PI / 180
                Case "G"
                    alfa_poste = (poste(n_poste1).anguloeje + 90) * PI / 180
            End Select
            alfa = poste(n_poste1).anguloeje * PI / 180
            PA(0) = poste(n_poste1).pk_coordx + desc1 * Cos(alfa + PI / 2)
            PA(1) = poste(n_poste1).pk_coordy + desc1 * sin(alfa + PI / 2)
            PA(2) = 0

            PA1(0) = poste(n_poste1).pk_coordx + eje_poste * Cos(alfa_poste)
            PA1(1) = poste(n_poste1).pk_coordy + eje_poste * sin(alfa_poste)
            PA1(2) = 0
            'Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(PA1, 0.05)
            Select Case poste(n_poste2).lado
                Case "D"
                    alfa_poste1 = (poste(n_poste2).anguloeje - 90) * PI / 180
                Case "G"
                    alfa_poste1 = (poste(n_poste2).anguloeje + 90) * PI / 180
            End Select
            alfa1 = poste(n_poste + 1).anguloeje * PI / 180
            'eje_poste1 = -poste(iposte + 1).descentramiento / 1000 + poste(iposte + 1).implantacion + (ancho_via / 2) + (ancho_poste) + ancho_carril + 2
            PB(0) = poste(n_poste2).pk_coordx + desc2 * Cos(alfa1 + PI / 2)
            PB(1) = poste(n_poste2).pk_coordy + desc2 * sin(alfa1 + PI / 2)
            PB(2) = 0
            
            PB1(0) = poste(n_poste2).pk_coordx + eje_poste1 * Cos(alfa_poste1)
            PB1(1) = poste(n_poste2).pk_coordy + eje_poste1 * sin(alfa_poste1)
            PB1(2) = 0
            'Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(PB1, 0.05)
            alpa = Atn((PB(1) - PA(1)) / (PB(0) - PA(0)))
            insertionpnt(0) = PB(0) - Cos(alpa) * pos
            insertionpnt(1) = PB(1) - sin(alpa) * pos
            insertionpnt(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt, 0.05)
            circulo.Layer = "E-CONEXION"
            alpa1 = Atn((PB1(1) - PA1(1)) / (PB1(0) - PA1(0)))
            insertionpnt1(0) = PB1(0) - Cos(alpa1) * pos
            insertionpnt1(1) = PB1(1) - sin(alpa1) * pos
            insertionpnt1(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt1, 0.05)
            circulo.Layer = "E-CONEXION"
            center(0) = (insertionpnt1(0) - Cos(alfa1) * 0.5 + insertionpnt(0) - Cos(alfa) * 0.5) / 2
            center(1) = (insertionpnt1(1) - sin(alfa1) * 0.5 + insertionpnt(1) - sin(alfa) * 0.5) / 2
            center(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(insertionpnt2, insertionpnt)
            linea.Layer = "E-CONEXION"
            
            
            alpa_ini = cuadrante(center(0), center(1), insertionpnt(0), insertionpnt(1), Abs(Atn((center(1) - insertionpnt(1)) / (center(0) - insertionpnt(0)))))
            qua_ini = qua
            alpa_fin = cuadrante(center(0), center(1), insertionpnt1(0), insertionpnt1(1), Abs(Atn((center(1) - insertionpnt1(1)) / (center(0) - insertionpnt1(0)))))
            qua_fin = qua
            diam = Math.Sqr((center(0) - insertionpnt(0)) ^ 2 + ((center(1) - insertionpnt(1)) ^ 2))
            Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_ini, alpa_fin)
            Set arco1 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_fin, alpa_ini)
            If arco1.ArcLength > arco.ArcLength And pos = 1 Then
                arco.Delete
                arco1.Layer = "E-CONEXION"
            Else
                arco1.Delete
                arco.Layer = "E-CONEXION"
            End If
        
        '///
        '/// Añadir conexión rep51
        '///
        ElseIf poste(iposte + 1).tipo = eje_sm And poste(iposte).tipo = semi_eje_sm Then
            If poste(iposte + 1).tipo = eje_sla Then
                ancho_mens = 1.6
             Else
                ancho_mens = 1
            End If
            alfa = poste(iposte).anguloeje * PI / 180
            PA(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) + ancho_mens / 2 * Cos(alfa)
            PA(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) + ancho_mens / 2 * sin(alfa)
            PA(2) = 0

            PA1(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento_2mens / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
            PA1(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento_2mens / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
            PA1(2) = 0
            
            alfa1 = poste(iposte + 1).anguloeje * PI / 180
            PB(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento / 1000 * Cos(alfa1 + PI / 2) + ancho_mens / 2 * Cos(alfa1)
            PB(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento / 1000 * sin(alfa1 + PI / 2) + ancho_mens / 2 * sin(alfa1)
            PB(2) = 0
            
            PB1(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento_2mens / 1000 * Cos(alfa1 + PI / 2) - ancho_mens / 2 * Cos(alfa1)
            PB1(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento_2mens / 1000 * sin(alfa1 + PI / 2) - ancho_mens / 2 * sin(alfa1)
            PB1(2) = 0
            
            'ang_plus = cuadrante(PA, PB)
            alpa = cuadrante(PA(0), PA(1), PB(0), PB(1), Abs(Atn((PB(1) - PA(1)) / (PB(0) - PA(0)))))
            insertionpnt(0) = PA(0) + Cos(alpa) * dist_con
            insertionpnt(1) = PA(1) + sin(alpa) * dist_con
            insertionpnt(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt, 0.05)
            circulo.Layer = "E-CONEXION"
            alpa1 = cuadrante(PA1(0), PA1(1), PB1(0), PB1(1), Abs(Atn((PB1(1) - PA1(1)) / (PB1(0) - PA1(0)))))
            insertionpnt1(0) = PA1(0) + Cos(alpa1) * (dist_con + ancho_mens)
            insertionpnt1(1) = PA1(1) + sin(alpa1) * (dist_con + ancho_mens)
            insertionpnt1(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt1, 0.05)
            circulo.Layer = "E-CONEXION"
            'ang_plus = cuadrante(insertionpnt, insertionpnt1)
            center(0) = (insertionpnt1(0) + Cos(alfa) * 0.5 + insertionpnt(0) + Cos(alfa) * 0.5) / 2
            center(1) = (insertionpnt1(1) + sin(alfa) * 0.5 + insertionpnt(1) + sin(alfa) * 0.5) / 2
            center(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(insertionpnt2, center)
            linea.Layer = "E-CONEXION"
            'ang_plus = cuadrante(center, insertionpnt)
            alpa_ini = cuadrante(center(0), center(1), insertionpnt(0), insertionpnt(1), Abs(Atn((center(1) - insertionpnt(1)) / (center(0) - insertionpnt(0)))))
            alpa_fin = cuadrante(center(0), center(1), insertionpnt1(0), insertionpnt1(1), Abs(Atn((center(1) - insertionpnt1(1)) / (center(0) - insertionpnt1(0)))))
            diam = Math.Sqr((center(0) - insertionpnt(0)) ^ 2 + ((center(1) - insertionpnt(1)) ^ 2))
            Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_fin, alpa_ini)
            Set arco1 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_ini, alpa_fin)
            If arco1.ArcLength > arco.ArcLength Then
                arco.Delete
                arco1.Layer = "E-CONEXION"
            Else
                arco1.Delete
                arco.Layer = "E-CONEXION"
            End If
            
        ElseIf poste(iposte + 1).tipo = semi_eje_sm And poste(iposte).tipo = eje_sm Then
            If poste(iposte + 1).tipo = eje_sla Then
                ancho_mens = 1.6
             Else
                ancho_mens = 1
            End If
            alfa = poste(iposte).anguloeje * PI / 180
            PA(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) + ancho_mens / 2 * Cos(alfa)
            PA(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) + ancho_mens / 2 * sin(alfa)
            PA(2) = 0

            PA1(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento_2mens / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
            PA1(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento_2mens / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
            PA1(2) = 0
            
            alfa1 = poste(iposte + 1).anguloeje * PI / 180
            PB(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento / 1000 * Cos(alfa1 + PI / 2) + ancho_mens / 2 * Cos(alfa1)
            PB(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento / 1000 * sin(alfa1 + PI / 2) + ancho_mens / 2 * sin(alfa1)
            PB(2) = 0
            
            PB1(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento_2mens / 1000 * Cos(alfa1 + PI / 2) - ancho_mens / 2 * Cos(alfa1)
            PB1(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento_2mens / 1000 * sin(alfa1 + PI / 2) - ancho_mens / 2 * sin(alfa1)
            PB1(2) = 0
            'ang_plus = cuadrante(PA, PB)
            alpa = cuadrante(PA(0), PA(1), PB(0), PB(1), Abs(Atn((PB(1) - PA(1)) / (PB(0) - PA(0)))))
            insertionpnt(0) = PB(0) - Cos(alpa) * dist_con
            insertionpnt(1) = PB(1) - sin(alpa) * dist_con
            insertionpnt(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt, 0.05)
            circulo.Layer = "E-CONEXION"
            alpa1 = cuadrante(PA1(0), PA1(1), PB1(0), PB1(1), Abs(Atn((PB1(1) - PA1(1)) / (PB1(0) - PA1(0)))))
            insertionpnt1(0) = PB1(0) - Cos(alpa1) * (dist_con - ancho_mens)
            insertionpnt1(1) = PB1(1) - sin(alpa1) * (dist_con - ancho_mens)
            insertionpnt1(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt1, 0.05)
            circulo.Layer = "E-CONEXION"
            center(0) = (insertionpnt1(0) - Cos(alfa) * 0.5 + insertionpnt(0) - Cos(alfa) * 0.5) / 2
            center(1) = (insertionpnt1(1) - sin(alfa) * 0.5 + insertionpnt(1) - sin(alfa) * 0.5) / 2
            center(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(insertionpnt2, center)
            linea.Layer = "E-CONEXION"
            alpa_ini = cuadrante(center(0), center(1), insertionpnt(0), insertionpnt(1), Abs(Atn((center(1) - insertionpnt(1)) / (center(0) - insertionpnt(0)))))
            alpa_fin = cuadrante(center(0), center(1), insertionpnt1(0), insertionpnt1(1), Abs(Atn((center(1) - insertionpnt1(1)) / (center(0) - insertionpnt1(0)))))
            diam = Math.Sqr((center(0) - insertionpnt(0)) ^ 2 + ((center(1) - insertionpnt(1)) ^ 2))
            Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_ini, alpa_fin)
            Set arco1 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_fin, alpa_ini)
            If arco1.ArcLength > arco.ArcLength Then
                arco.Delete
                arco1.Layer = "E-CONEXION"
            Else
                arco1.Delete
                arco.Layer = "E-CONEXION"
            End If
        '///
        '/// Añadir conexión rep53 1
        '///
        ElseIf poste(iposte + 1).tipo = semi_eje_sla And (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin) Then
            If poste(iposte + 1).tipo = eje_sla Then
                ancho_mens = 1.6
             Else
                ancho_mens = 1
            End If
            If poste(iposte).tunel = False Then
                eje_poste = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril
            Else
                eje_poste = 1.4
            End If
            Select Case poste(iposte).lado
                Case "D"
                    alfa_poste = (poste(iposte).anguloeje - 90) * PI / 180
                Case "G"
                    alfa_poste = (poste(iposte).anguloeje + 90) * PI / 180
            End Select
            alfa = poste(iposte).anguloeje * PI / 180
            PA(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2)
            PA(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2)
            PA(2) = 0

            PA1(0) = poste(iposte).pk_coordx + eje_poste * Cos(alfa_poste)
            PA1(1) = poste(iposte).pk_coordy + eje_poste * sin(alfa_poste)
            PA1(2) = 0
            
            alfa1 = poste(iposte + 1).anguloeje * PI / 180
            PB(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento / 1000 * Cos(alfa1 + PI / 2) + ancho_mens / 2 * Cos(alfa1)
            PB(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento / 1000 * sin(alfa1 + PI / 2) + ancho_mens / 2 * sin(alfa1)
            PB(2) = 0
            
            PB1(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento_2mens / 1000 * Cos(alfa1 + PI / 2) - ancho_mens / 2 * Cos(alfa1)
            PB1(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento_2mens / 1000 * sin(alfa1 + PI / 2) - ancho_mens / 2 * sin(alfa1)
            PB1(2) = 0

            alpa = Atn((PB(1) - PA(1)) / (PB(0) - PA(0)))
            insertionpnt(0) = PB(0) - Cos(alpa) * dist_con
            insertionpnt(1) = PB(1) - sin(alpa) * dist_con
            insertionpnt(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt, 0.05)
            circulo.Layer = "E-CONEXION"
            alpa1 = Atn((PB1(1) - PA1(1)) / (PB1(0) - PA1(0)))
            insertionpnt1(0) = PB1(0) - Cos(alpa1) * (dist_con - ancho_mens)
            insertionpnt1(1) = PB1(1) - sin(alpa1) * (dist_con - ancho_mens)
            insertionpnt1(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt1, 0.05)
            circulo.Layer = "E-CONEXION"
            center(0) = (insertionpnt1(0) - Cos(alfa) * 0.5 + insertionpnt(0) - Cos(alfa) * 0.5) / 2
            center(1) = (insertionpnt1(1) - sin(alfa) * 0.5 + insertionpnt(1) - sin(alfa) * 0.5) / 2
            center(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(center, insertionpnt2)
            linea.Layer = "E-CONEXION"
            alpa_ini = cuadrante(center(0), center(1), insertionpnt(0), insertionpnt(1), Abs(Atn((center(1) - insertionpnt(1)) / (center(0) - insertionpnt(0)))))
            qua_ini = qua
            alpa_fin = cuadrante(center(0), center(1), insertionpnt1(0), insertionpnt1(1), Abs(Atn((center(1) - insertionpnt1(1)) / (center(0) - insertionpnt1(0)))))
            qua_fin = qua
            diam = Math.Sqr((center(0) - insertionpnt(0)) ^ 2 + ((center(1) - insertionpnt(1)) ^ 2))
            Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_ini, alpa_fin)
            Set arco1 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_fin, alpa_ini)
            If arco1.ArcLength > arco.ArcLength Then
                arco.Delete
                arco1.Layer = "E-CONEXION"
            Else
                arco1.Delete
                arco.Layer = "E-CONEXION"
            End If
        '///
        '/// Añadir conexión rep53 2
        '///
        ElseIf (poste(iposte + 1).tipo = anc_sla_con Or poste(iposte + 1).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte + 1).tipo = anc_sla_sin) And (poste(iposte).tipo = semi_eje_sla Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj) Then

            If poste(iposte + 1).tipo = eje_sla Then
                ancho_mens = 1.6
             Else
                ancho_mens = 1
            End If
            If poste(iposte + 1).tunel = False Then
                eje_poste = poste(iposte + 1).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril
            Else
                eje_poste = 1.4
            End If
            Select Case poste(iposte + 1).lado
                Case "D"
                    alfa_poste = (poste(iposte + 1).anguloeje - 90) * PI / 180
                Case "G"
                    alfa_poste = (poste(iposte + 1).anguloeje + 90) * PI / 180
            End Select
            alfa = poste(iposte).anguloeje * PI / 180
            PA(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) + ancho_mens / 2 * Cos(alfa)
            PA(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) + ancho_mens / 2 * sin(alfa)
            PA(2) = 0
            
            PA1(0) = poste(iposte).pk_coordx + poste(iposte).descentramiento_2mens / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
            PA1(1) = poste(iposte).pk_coordy + poste(iposte).descentramiento_2mens / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
            PA1(2) = 0
            
            alfa1 = poste(iposte + 1).anguloeje * PI / 180
            
            PB(0) = poste(iposte + 1).pk_coordx + eje_poste * Cos(alfa_poste)
            PB(1) = poste(iposte + 1).pk_coordy + eje_poste * sin(alfa_poste)
            PB(2) = 0
                        
            PB1(0) = poste(iposte + 1).pk_coordx + poste(iposte + 1).descentramiento / 1000 * Cos(alfa1 + PI / 2)
            PB1(1) = poste(iposte + 1).pk_coordy + poste(iposte + 1).descentramiento / 1000 * sin(alfa1 + PI / 2)
            PB1(2) = 0
            
            alpa = cuadrante(PA(0), PA(1), PB(0), PB(1), Abs(Atn((PB(1) - PA(1)) / (PB(0) - PA(0)))))
            insertionpnt(0) = PA(0) + Cos(alpa) * dist_con
            insertionpnt(1) = PA(1) + sin(alpa) * dist_con
            insertionpnt(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt, 0.05)
            circulo.Layer = "E-CONEXION"
            alpa1 = cuadrante(PA1(0), PA1(1), PB1(0), PB1(1), Abs(Atn((PB1(1) - PA1(1)) / (PB1(0) - PA1(0)))))
            insertionpnt1(0) = PA1(0) + Cos(alpa1) * (dist_con + ancho_mens)
            insertionpnt1(1) = PA1(1) + sin(alpa1) * (dist_con + ancho_mens)
            insertionpnt1(2) = 0
            Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(insertionpnt1, 0.05)
            circulo.Layer = "E-CONEXION"
            center(0) = (insertionpnt1(0) + Cos(alpa1) * 0.5 + insertionpnt(0) + Cos(alpa) * 0.5) / 2
            center(1) = (insertionpnt1(1) + sin(alpa1) * 0.5 + insertionpnt(1) + sin(alpa) * 0.5) / 2
            center(2) = 0
            Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(center, insertionpnt2)
            linea.Layer = "E-CONEXION"
            alpa_ini = cuadrante(center(0), center(1), insertionpnt(0), insertionpnt(1), Abs(Atn((center(1) - insertionpnt(1)) / (center(0) - insertionpnt(0)))))
            qua_ini = qua
            alpa_fin = cuadrante(center(0), center(1), insertionpnt1(0), insertionpnt1(1), Abs(Atn((center(1) - insertionpnt1(1)) / (center(0) - insertionpnt1(0)))))
            qua_fin = qua
            diam = Math.Sqr((center(0) - insertionpnt(0)) ^ 2 + ((center(1) - insertionpnt(1)) ^ 2))
            Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_ini, alpa_fin)
            Set arco1 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(center, diam, alpa_fin, alpa_ini)
            If arco1.ArcLength > arco.ArcLength Then
                arco.Delete
                arco1.Layer = "E-CONEXION"
            Else
                arco1.Delete
                arco.Layer = "E-CONEXION"
            End If

        End If
    End If
Next
If num_impla = 0 Then
    GetObject(, "Autocad.Application").ActiveDocument.Layers.Item("E-CONEXION").Delete
End If
End Sub
Function cuadrante(PA0, PA1, PB0, PB1, alga) As Double

If PB0 > PA0 And PB1 > PA1 Then
'///cuadrante 1
cuadrante = alga
qua = 1
ElseIf PB0 < PA0 And PB1 > PA1 Then
'///cuadrante 2
cuadrante = 3.1415 - Abs(alga)
qua = 2
ElseIf PB0 < PA0 And PB1 < PA1 Then
'///cuadrante 3
cuadrante = 3.1415 + Abs(alga)
qua = 3
ElseIf PB0 > PA0 And PB1 < PA1 Then
'///cuadrante 4
cuadrante = 2 * 3.1415 - Abs(alga)
qua = 4
End If
End Function

Sub dibujar_proteccion(cadena_ruta As String, HDC As Boolean)
Dim cadena_prot As String
Dim bloque_prot As AcadBlockReference
Dim dist_prot_eje, alfa_etiq As Double
Dim insertionpnt(0 To 2) As Double
Dim Atributos As Variant
dist_eje = 21
cont = 0
cont_est = 0
escala_pro = 1
ancho_poste = 0.46
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-PROTECCION")
For iposte = 1 To num_postes_total
    '///
    '/// Verificación del lado a implantar los datos en estacion
    '///
    'If poste(iposte + 1).aguja <> "" And cont_est = 0 Then
        'While estacion(ipk).nombre <> poste(iposte + 1).aguja
            'ipk = ipk + 1
        'Wend
        'If estacion(ipk).lado = True Then
            'dist_eje = -21
        'ElseIf estacion(ipk).lado = False Then
            'dist_eje = 18
        'End If
        'cont_est = 1
    'ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
        'cont_est = 2
    'ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 2 Then
        'dist_eje = 21
        'cont_est = 0
        'ipk = 0
    'End If
    
    
    '///
    '/// Variar la escala en estaciones
    '///
    
        'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
            escala_pro = 0.5
            cont = 2 ' solo para vias secundarias
            'cont = cont + 1
            'dist_eje = 18
        'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte - 1).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
            'escala_pro = 1
            'cont = 0
            'dist_eje = 21
        'End If

    If poste(iposte).proteccion <> "" Then
            alfa_prot = poste(iposte).anguloeje * PI / 180
            insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_prot + PI / 2)
            insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_prot + PI / 2)
            insertionpnt(2) = 0
        If poste(iposte).proteccion = "Parafoudres - DPPo" Then
            cadena_prot = cadena_ruta & "Proteccion1.dwg"
            Set bloque_prot = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_prot, escala_pro, escala_pro, escala_pro, alfa_prot)
            bloque_prot.Layer = "E-PROTECCION"
            cadena_prot = cadena_ruta & "Proteccion2.dwg"
            Set bloque_prot = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_prot, escala_pro, escala_pro, escala_pro, alfa_prot)
            bloque_prot.Layer = "E-PROTECCION"
        ElseIf poste(iposte).proteccion = "DPPo" Then
            cadena_prot = cadena_ruta & "Proteccion2.dwg"
            Set bloque_prot = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_prot, escala_pro, escala_pro, escala_pro, alfa_prot)
            bloque_prot.Layer = "E-PROTECCION"
        End If
        Set bloque_prot = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_prot, escala_pro, escala_pro, escala_pro, alfa_prot)
        bloque_prot.Layer = "E-PROTECCION"
    End If
    
Next iposte
End Sub
Sub dibujar_pendola(cadena_ruta As String, HDC As Boolean)
Dim bloque_pen As AcadBlockReference
Dim pen As Double
Dim Atributos As Variant
Dim insertionpnt(0 To 2) As Double
Dim alfa_pen As Double
Dim dist_eje As Single, dist_eje2 As Single
escala_pen = 1
dist_eje = 15
dist_eje2 = 17
cont = 0
cont_est = 0
ancho_poste = 0.46
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-PENDOLA")

For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_pen = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont = 3 Then
        'escala_pen = 1
        'cont = 0
    'End If
    alfa_vano = poste(iposte).alfa_vano
    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
                dist_eje2 = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 10)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
                dist_eje2 = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 10)
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 13)
            dist_eje2 = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 16)
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
                dist_eje2 = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 10)
            ElseIf estacion(ipk).lado = False Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 8)
                dist_eje2 = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 10)
            End If
        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 13)
            dist_eje2 = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 16)
        End If
    ElseIf HDC = True Then
        If cont < 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 13)
            dist_eje2 = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 16)
        ElseIf cont >= 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 9)
            dist_eje2 = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 11)
        End If
    End If
    
    insertionpnt(0) = poste(iposte).pk_vano_postx + dist_eje * Cos(alfa_vano + PI / 2)
    insertionpnt(1) = poste(iposte).pk_vano_posty + dist_eje * sin(alfa_vano + PI / 2)
    insertionpnt(2) = 0
    Set bloque_pen = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Pendola.dwg", escala_pen, escala_pen, escala_pen, alfa_vano)
    bloque_pen.Layer = "E-PENDOLA"
    Atributos = bloque_pen.GetAttributes
    Atributos(0).TextString = poste(iposte).pendola1
    
    
    If poste(iposte).pendola2 <> "" Then
        alfa_vano = poste(iposte).alfa_vano
        insertionpnt(0) = poste(iposte).pk_vano_postx + dist_eje2 * Cos(alfa_vano + PI / 2)
        insertionpnt(1) = poste(iposte).pk_vano_posty + dist_eje2 * sin(alfa_vano + PI / 2)
        insertionpnt(2) = 0
        Set bloque_pen = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Pendola.dwg", escala_pen, escala_pen, escala_pen, alfa_vano)
        bloque_pen.Layer = "E-PENDOLA"
        Atributos = bloque_pen.GetAttributes
        Atributos(0).TextString = poste(iposte).pendola2
    End If
Next
End Sub
Sub dibujar_alt_cat(cadena_ruta As String, HDC As Boolean)
Dim bloque_alt_cat As AcadBlockReference
Dim Atributos As Variant
Dim insertionpnt(0 To 2) As Double
Dim dist_eje As Single, dist_poste As Single
Dim alfa_alt_cat As Double
escala_alt = 1
dist_eje = 19.5
dist_poste = 4
cont = 0
cont_est = 0
ancho_poste = 0.46
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-ALTURA CATENARIA")

For iposte = 1 To num_postes_total
    '///
    '/// Variar la escala en estaciones
    '///
    'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
        escala_alt = 0.5
        cont = 2 ' solo para vias secundarias
        'cont = cont + 1
    'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte - 1).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
        'escala_alt = 1
        'cont = 0
    'End If
    If HDC = False Then
        '///
        '/// Verificación del lado a implantar los datos en estacion
        '///
        If poste(iposte).aguja <> "" And cont_est = 0 Then
            While estacion(ipk).nombre <> poste(iposte).aguja
                ipk = ipk + 1
            Wend
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
                dist_poste = 2
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
                dist_poste = 2
            End If
            cont_est = 1
        ElseIf poste(iposte).aguja = estacion(ipk).nombre And cont_est = 1 Then
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 17)
            dist_poste = 4
            cont_est = 0
            ipk = 0
        ElseIf cont_est = 1 Then
            If estacion(ipk).lado = True Then
                dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
                dist_poste = 2
            ElseIf estacion(ipk).lado = False Then
                dist_eje = -(poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 14)
                dist_poste = 2
            End If
        Else
            dist_eje = (poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril + 17)
        End If
        
            
    ElseIf HDC = True Then
        If cont < 1 Then
            dist_eje = -14
            dist_poste = 2
        ElseIf cont >= 1 Then
            dist_eje = -10
            dist_poste = 1
        End If
    End If

    If poste(iposte).alt_cat(0) <> "" Then

        alfa_alt_cat = (poste(iposte).anguloeje) * PI / 180
        insertionpnt(0) = poste(iposte).pk_coordx + dist_eje * Cos(alfa_alt_cat - PI / 2) '+ dist_poste * Cos(alfa_alt_cat)
        insertionpnt(1) = poste(iposte).pk_coordy + dist_eje * sin(alfa_alt_cat - PI / 2) '+ dist_poste * sin(alfa_alt_cat)
        insertionpnt(2) = 0
        Set bloque_alt_cat = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Altura_catenaria.dwg", escala_alt, escala_alt, escala_alt, alfa_alt_cat)
        bloque_alt_cat.Layer = "E-ALTURA CATENARIA"
        Atributos = bloque_alt_cat.GetAttributes
        If poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = anc_pf & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_pf & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = anc_pf & " + " & eje_aguj Or poste(iposte).tipo = eje_pf & " + " & eje_aguj _
         Or poste(iposte).tipo = anc_aguj & " + " & semi_eje_aguj Then
            Atributos(0).TextString = "e=" & poste(iposte).alt_cat(0)
        Else
            Atributos(0).TextString = "e=" & poste(iposte).alt_cat(1)
        End If
        If poste(iposte).alt_cat(1) <> "" Then
                If poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = anc_pf & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_pf & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = anc_pf & " + " & eje_aguj Or poste(iposte).tipo = eje_pf & " + " & eje_aguj _
                 Or poste(iposte).tipo = anc_aguj & " + " & semi_eje_aguj Then
                    Atributos(1).TextString = "e=" & poste(iposte).alt_cat(1)
                Else
                    Atributos(1).TextString = "e=" & poste(iposte).alt_cat(0)
                End If
        End If
    End If
Next
End Sub
Sub dibujar_singular(cadena_ruta As String, HDC As Boolean)
Dim bloque_sing As AcadBlockReference
Dim linea As AcadLine
Dim texto As AcadText
Dim Atributos As Variant
Dim ini_linea(0 To 2) As Double
Dim fin_linea(0 To 2) As Double
Dim ini_flecha(0 To 2) As Double
Dim med_flecha(0 To 2) As Double
Dim fin_flecha(0 To 2) As Double
Dim medio_flecha(0 To 2) As Double
Dim insertionpnt(0 To 2) As Double
Dim dist_sing As Single, dist_poste As Single
Dim alfa_sing As Double
Dim accapa As AcadLayer
Dim algo As String
Dim height As Double
Call Obtener_excel_pks
escala_ps = 1
escala_fl = 0.5
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-PUNTO SINGULAR")
iposte = 1
height = 1
For i = 1 To ips - 1
    While poste(iposte).pk_global < p_s(i).pk_inicio And poste(iposte).pk_global <> 0
        iposte = iposte + 1
    
        '///
        '/// Variar la escala en estaciones
        '///
        'If (poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_sin Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj) And cont < 3 And iposte <> 1 And iposte <> 6 Then
            'escala_ps = 0.5
            'escala_fl = 0.25
            'height = 0.5
            'cont = cont + 1
        'ElseIf (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin Or poste(iposte - 1).tipo = anc_sla_con & " + " & semi_eje_aguj) And poste(iposte).tipo = "" And cont = 3 Then
            'escala_ps = 1
            'escala_fl = 0.5
            'height = 1
            'cont = 0
        'End If
    Wend
    If poste(iposte).lado = "G" Then
        dere = PI / 2
    Else
        dere = -PI / 2
    End If
    
    If p_s(i + 1).pk_medio - p_s(i).pk_medio < 15 Then
        dist_eje = 41
        dist_eje2 = 36
    Else
        dist_eje = 31
        dist_eje2 = 26
        dist_eje3 = 1.5
    End If

    alfa_poste = poste(iposte - 1).anguloeje * PI / 180
    alfa_inicio = p_s(i).anguloinicio * PI / 180
    ini_flecha(0) = p_s(i).pk_iniciox + dist_eje3 * Cos(alfa_inicio + dere)
    ini_flecha(1) = p_s(i).pk_inicioy + dist_eje3 * sin(alfa_inicio + dere)
    ini_flecha(2) = 0
    fin_flecha(0) = poste(iposte - 1).pk_coordx + dist_eje3 * Cos(alfa_poste + dere)
    fin_flecha(1) = poste(iposte - 1).pk_coordy + dist_eje3 * sin(alfa_poste + dere)
    fin_flecha(2) = 0
    med_flecha(0) = (ini_flecha(0) + fin_flecha(0)) / 2
    med_flecha(1) = (ini_flecha(1) + fin_flecha(1)) / 2
    med_flecha(2) = 0
    
    alfa_sing = p_s(i).angulomedio * PI / 180
    insertionpnt(0) = p_s(i).pk_mediox + dist_eje * Cos(alfa_sing - PI / 2)
    insertionpnt(1) = p_s(i).pk_medioy + dist_eje * sin(alfa_sing - PI / 2)
    insertionpnt(2) = 0
    ini_linea(0) = p_s(i).pk_mediox + dist_eje2 * Cos(alfa_sing - PI / 2)
    ini_linea(1) = p_s(i).pk_medioy + dist_eje2 * sin(alfa_sing - PI / 2)
    ini_linea(2) = 0
    fin_linea(0) = p_s(i).pk_mediox
    fin_linea(1) = p_s(i).pk_medioy
    fin_linea(2) = 0
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_flecha, fin_flecha)
    linea.Layer = "E-PUNTO SINGULAR"
    Set bloque_sing = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(linea.StartPoint, cadena_ruta & "Punta_flecha.dwg", escala_fl, escala_fl, escala_fl, linea.angle)
    bloque_sing.Layer = "E-PUNTO SINGULAR"
    Set bloque_sing = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(linea.EndPoint, cadena_ruta & "Punta_flecha.dwg", escala_fl, escala_fl, escala_fl, linea.angle - PI)
    bloque_sing.Layer = "E-PUNTO SINGULAR"
    algo = Round(linea.Length, 2)

    Set texto = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddText(algo, med_flecha, height)
    texto.Rotate med_flecha, linea.angle + dere
    texto.Layer = "E-PUNTO SINGULAR"
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, fin_linea)
    linea.Layer = "E-PUNTO SINGULAR"
    linea.LineType = "LÍNEAS_OCULTASX2" '"LÍNEAS_OCULTASX2"
    Set bloque_sing = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(fin_linea, cadena_ruta & "Punta_flecha_vacia.dwg", escala_fl, escala_fl, escala_fl, alfa_sing - PI / 2)
    bloque_sing.Layer = "E-PUNTO SINGULAR"
    Set bloque_sing = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "Punto_singular.dwg", escala_ps, escala_ps, escala_ps, alfa_sing)
    bloque_sing.Layer = "E-PUNTO SINGULAR"
    Atributos = bloque_sing.GetAttributes
    Atributos(0).TextString = p_s(i).tipo
    Atributos(1).TextString = "Entrée: " & convertir_pk(p_s(i).pk_inicio)
    Atributos(2).TextString = "Sortie: " & convertir_pk(p_s(i).pk_final)
    Atributos(3).TextString = "Long: " & Round((p_s(i).pk_final - p_s(i).pk_inicio), 2)
    While poste(iposte).pk_global < p_s(i).pk_final And poste(iposte).pk_global <> 0
        iposte = iposte + 1
    Wend
    alfa_poste = poste(iposte).anguloeje * PI / 180
    alfa_inicio = p_s(i).angulofinal * PI / 180
    ini_flecha(0) = p_s(i).pk_finalx + dist_eje3 * Cos(alfa_inicio + dere)
    ini_flecha(1) = p_s(i).pk_finaly + dist_eje3 * sin(alfa_inicio + dere)
    ini_flecha(2) = 0
    fin_flecha(0) = poste(iposte).pk_coordx + dist_eje3 * Cos(alfa_poste + dere)
    fin_flecha(1) = poste(iposte).pk_coordy + dist_eje3 * sin(alfa_poste + dere)
    fin_flecha(2) = 0
    med_flecha(0) = (ini_flecha(0) + fin_flecha(0)) / 2
    med_flecha(1) = (ini_flecha(1) + fin_flecha(1)) / 2
    med_flecha(2) = 0
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_flecha, fin_flecha)
    linea.Layer = "E-PUNTO SINGULAR"
    Set bloque_sing = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(linea.StartPoint, cadena_ruta & "Punta_flecha.dwg", escala_fl, escala_fl, escala_fl, linea.angle)
    bloque_sing.Layer = "E-PUNTO SINGULAR"
    Set bloque_sing = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(linea.EndPoint, cadena_ruta & "Punta_flecha.dwg", escala_fl, escala_fl, escala_fl, linea.angle - PI)
    bloque_sing.Layer = "E-PUNTO SINGULAR"
    algo = Round(linea.Length, 1)
    Set texto = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddText(algo, med_flecha, height)
    texto.Rotate med_flecha, linea.angle - dere
    texto.Layer = "E-PUNTO SINGULAR"
siguiente:
Next i
End Sub
Sub dibujar_linea()
Dim linea As AcadLine
Dim arco As AcadArc
Dim circulo As AcadCircle
Dim acHC As AcadLayer, aCCdpa As AcadLayer, acfeeder As AcadLayer
Dim poly As AcadPolyline, poly2 As AcadPolyline
Dim polyCdpa() As Double, polyfeed() As Double
Set acHC = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-HILO CONTACTO")
Set acCa_el = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-CABLEADO ELEVADO")
Set aCCdpa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-CDPA")
Set acfeeder = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-FEEDER")
Dim long_hc As Double, long_cdpa As Double
'GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "CDPA", "acadiso.lin"
'GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "ACAD_ISO06W100", "acadiso.lin"
'GetObject(, "Autocad.Application").ActiveDocument.Linetypes.Load "LÍNEAS_OCULTAS", "acadiso.lin"
cadena_etiq = cadena_ruta & "Implantation.dwg"
acHC.Color = acRed
acCa_el.Color = acRed
aCCdpa.Color = acBlue
aCCdpa.LineType = "CDPA"
acfeeder.LineType = "ACAD_ISO06W100"
acCa_el.LineType = "LÍNEAS_OCULTAS"
canton = 1
ancho_poste = 0.46
i = 0
e = 1
long_hc = 0
long_pf = 0
long_cdpa = 0
long_ais2 = 3
ancho_poste = 0.5
For iposte = 1 To num_postes_total
    long_ais = poste(iposte).vano_post / 3.4
    alfa = poste(iposte).anguloeje * PI / 180
    Select Case poste(iposte).lado
        Case "D"
            alfa_poste = (poste(iposte).anguloeje - 90) * PI / 180
        Case "G"
            alfa_poste = (poste(iposte).anguloeje + 90) * PI / 180
    End Select
    If poste(iposte).tipo = eje_sla Then
        ancho_mens = 1.6
    Else
        ancho_mens = 1
    End If
    If poste(iposte).tunel = False Then
        eje_poste = poste(iposte).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril
    Else
        eje_poste = 1.4
     End If
NumberOfEl_i = NumberOfEl_i + 3
NumberOfEl_e = NumberOfEl_e + 3
NumberOfEl_c = NumberOfEl_c + 3
ReDim Preserve pol_canton(e).polyarray(1 To NumberOfEl_e)
ReDim Preserve pol_canton(i).polyarray(1 To NumberOfEl_i)
ReDim Preserve polyCdpa(1 To NumberOfEl_c)
ReDim Preserve polyfeed(1 To NumberOfEl_c)
Dim PA(0 To 2) As Double, ini_aux(0 To 2) As Double, centro(0 To 2) As Double
'//
'// Recoger datos en postes con poste inicial de anclaje
'//

If poste(iposte).tipo_secc = "Normal" Or poste(iposte).tipo_secc = "Inverso" Then
    NumberOfEl_i = NumberOfEl_e
    pol_canton(i).polyarray = pol_canton(e).polyarray
    a = i
    i = e
    e = a
    NumberOfEl_e = 0
    NumberOfEl_e = NumberOfEl_e + 3
    ReDim Preserve pol_canton(e).polyarray(1 To NumberOfEl_e)
'///
'///Dibujar hilo exterior - interior
'///
    pol_canton(e).polyarray(NumberOfEl_e - 2) = poste(iposte).pk_coordx + eje_poste * Cos(alfa_poste)
    pol_canton(e).polyarray(NumberOfEl_e - 1) = poste(iposte).pk_coordy + eje_poste * sin(alfa_poste)
    pol_canton(e).polyarray(NumberOfEl_e) = 0
    
    pol_canton(i).polyarray(NumberOfEl_i - 2) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2)
    pol_canton(i).polyarray(NumberOfEl_i - 1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2)
    pol_canton(i).polyarray(NumberOfEl_i) = 0

'//
'// Recoger datos en postes con dos mensulas
'//

ElseIf poste(iposte).descentramiento_2mens <> 0 And (Mid(poste(iposte).tipo, 15) <> semi_eje_aguj And Mid(poste(iposte).tipo, 15) <> eje_aguj And poste(iposte).tipo <> semi_eje_aguj And poste(iposte).tipo <> eje_aguj And poste(iposte).tipo <> anc_sla_con & " + " & semi_eje_aguj And poste(iposte).tipo <> eje_aguj & " + " & anc_aguj) Then
           
    pol_canton(e).polyarray(NumberOfEl_e - 2) = poste(iposte).pk_coordx + poste(iposte).descentramiento_2mens / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
    pol_canton(e).polyarray(NumberOfEl_e - 1) = poste(iposte).pk_coordy + poste(iposte).descentramiento_2mens / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
    pol_canton(e).polyarray(NumberOfEl_e) = 0
    
    pol_canton(i).polyarray(NumberOfEl_i - 2) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) + ancho_mens / 2 * Cos(alfa)
    pol_canton(i).polyarray(NumberOfEl_i - 1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) + ancho_mens / 2 * sin(alfa)
    pol_canton(i).polyarray(NumberOfEl_i) = 0
    
    '///
    '/// Añadir aislador entre anc_sm o anc_sla y semi_eje
    '///
    If (poste(iposte).tipo = semi_eje_sm And (poste(iposte - 1).tipo = anc_sm_con Or poste(iposte - 1).tipo = anc_sm_sin)) Or _
    (poste(iposte).tipo = semi_eje_sla And (poste(iposte - 1).tipo = anc_sla_con Or poste(iposte - 1).tipo = anc_sla_sin)) Then
        alpa = cuadrante(pol_canton(e).polyarray(NumberOfEl_e - 5), pol_canton(e).polyarray(NumberOfEl_e - 4), pol_canton(e).polyarray(NumberOfEl_e - 2), pol_canton(e).polyarray(NumberOfEl_e - 1), Abs(Atn((pol_canton(e).polyarray(NumberOfEl_e - 1) - pol_canton(e).polyarray(NumberOfEl_e - 4)) / (pol_canton(e).polyarray(NumberOfEl_e - 2) - pol_canton(e).polyarray(NumberOfEl_e - 5)))))
        'alpa = Atn((pol_canton(e).polyarray(NumberOfEl_e - 4) - pol_canton(e).polyarray(NumberOfEl_e - 1)) / (pol_canton(e).polyarray(NumberOfEl_e - 5) - pol_canton(e).polyarray(NumberOfEl_e - 2)))
        PA(0) = pol_canton(e).polyarray(NumberOfEl_e - 5) + Cos(alpa) * long_ais
        PA(1) = pol_canton(e).polyarray(NumberOfEl_e - 4) + sin(alpa) * long_ais
        PA(2) = 0
        Call cuadrado(PA, alpa, 0.5, 1, "E-HILO CONTACTO")

    '///
    '/// Añadir aislador entre semi_eje_sla y eje_sla
    '///
    ElseIf (poste(iposte).tipo = eje_sla And poste(iposte - 1).tipo = semi_eje_sla) Then
        alpa = cuadrante(pol_canton(e).polyarray(NumberOfEl_e - 5), pol_canton(e).polyarray(NumberOfEl_e - 4), pol_canton(e).polyarray(NumberOfEl_e - 2), pol_canton(e).polyarray(NumberOfEl_e - 1), Abs(Atn((pol_canton(e).polyarray(NumberOfEl_e - 1) - pol_canton(e).polyarray(NumberOfEl_e - 4)) / (pol_canton(e).polyarray(NumberOfEl_e - 2) - pol_canton(e).polyarray(NumberOfEl_e - 5)))))
        'alpa = Atn((pol_canton(e).polyarray(NumberOfEl_e - 4) - pol_canton(e).polyarray(NumberOfEl_e - 1)) / (pol_canton(e).polyarray(NumberOfEl_e - 5) - pol_canton(e).polyarray(NumberOfEl_e - 2)))
        PA(0) = pol_canton(e).polyarray(NumberOfEl_e - 5) + Cos(alpa) * long_ais2
        PA(1) = pol_canton(e).polyarray(NumberOfEl_e - 4) + sin(alpa) * long_ais2
        PA(2) = 0
        Call cuadrado(PA, alpa, 0.5, 1, "E-HILO CONTACTO")

    '///
    '/// Añadir aislador entre eje_sla y semi_eje_sla
    '///
    ElseIf ((poste(iposte).tipo = semi_eje_sla Or poste(iposte).tipo = semi_eje_sla & " + " & anc_aguj) And poste(iposte - 1).tipo = eje_sla) Then
        alpa = Atn((pol_canton(i).polyarray(NumberOfEl_i - 1) - pol_canton(i).polyarray(NumberOfEl_i - 4)) / (pol_canton(i).polyarray(NumberOfEl_i - 2) - pol_canton(i).polyarray(NumberOfEl_i - 5)))
        'alpa = Atn((pol_canton(i).polyarray(NumberOfEl_i - 1) - pol_canton(i).polyarray(NumberOfEl_i - 4)) / (pol_canton(i).polyarray(NumberOfEl_i - 2) - pol_canton(i).polyarray(NumberOfEl_i - 5)))
        PA(0) = pol_canton(i).polyarray(NumberOfEl_i - 2) - Cos(alpa) * long_ais2
        PA(1) = pol_canton(i).polyarray(NumberOfEl_i - 1) - sin(alpa) * long_ais2
        PA(2) = 0
        Call cuadrado(PA, alpa, 0.5, 1, "E-HILO CONTACTO")
    End If

'//
'// Recoger datos en postes simples y dibujar el canton
'//
ElseIf poste(iposte).descentramiento_2mens = 0 Or Mid(poste(iposte).tipo, 15) = semi_eje_aguj Or Mid(poste(iposte).tipo, 15) = eje_aguj Or poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = eje_aguj Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_aguj & " + " & anc_aguj Then
    If poste(iposte).tipo = anc_sm_con Or poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Then
        If poste(iposte - 1).tipo = semi_eje_sm Or poste(iposte - 1).tipo = semi_eje_sla Or poste(iposte - 1).tipo = semi_eje_sla & " + " & anc_aguj Then
            pol_canton(i).polyarray(NumberOfEl_i - 2) = poste(iposte).pk_coordx + eje_poste * Cos(alfa_poste)
            pol_canton(i).polyarray(NumberOfEl_i - 1) = poste(iposte).pk_coordy + eje_poste * sin(alfa_poste)
            pol_canton(i).polyarray(NumberOfEl_i) = 0
            
            long_hc = long_hc + separar_lineas(NumberOfEl_i, i, canton, 2)
            
            canton = 2
            'Set poly = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(pol_canton(i).polyarray)
            'poly.Layer = "E-HILO CONTACTO"
            'long_hc = poly.Length + long_hc
            '///
            '/// Añadir aislador entre semi_eje y anc_sm o anc_sla
            '///
            If ((poste(iposte).tipo = anc_sm_con Or poste(iposte).tipo = anc_sm_sin) And poste(iposte - 1).tipo = semi_eje_sm) Or _
            ((poste(iposte).tipo = anc_sla_con Or poste(iposte).tipo = anc_sla_con & " + " & semi_eje_aguj Or poste(iposte).tipo = anc_sla_sin) And (poste(iposte - 1).tipo = semi_eje_sla Or poste(iposte - 1).tipo = semi_eje_sla & " + " & anc_aguj)) Then
                alpa = cuadrante(pol_canton(i).polyarray(NumberOfEl_i - 5), pol_canton(i).polyarray(NumberOfEl_i - 4), pol_canton(i).polyarray(NumberOfEl_i - 2), pol_canton(i).polyarray(NumberOfEl_i - 1), Abs(Atn((pol_canton(i).polyarray(NumberOfEl_i - 1) - pol_canton(i).polyarray(NumberOfEl_i - 4)) / (pol_canton(i).polyarray(NumberOfEl_i - 2) - pol_canton(i).polyarray(NumberOfEl_i - 5)))))
                'alpa = Atn((pol_canton(i).polyarray(NumberOfEl_i - 4) - pol_canton(i).polyarray(NumberOfEl_i - 1)) / (pol_canton(i).polyarray(NumberOfEl_i - 5) - pol_canton(i).polyarray(NumberOfEl_i - 2)))
                PA(0) = pol_canton(i).polyarray(NumberOfEl_i - 2) - Cos(alpa) * long_ais
                PA(1) = pol_canton(i).polyarray(NumberOfEl_i - 1) - sin(alpa) * long_ais
                PA(2) = 0
                Call cuadrado(PA, alpa, 0.5, 1, "E-HILO CONTACTO")
            End If
        
        End If
    End If
    If Mid(poste(iposte).tipo, 15) = semi_eje_aguj Or Mid(poste(iposte).tipo, 15) = eje_aguj Or poste(iposte).tipo = semi_eje_aguj Or poste(iposte).tipo = eje_aguj Then
        pol_canton(e).polyarray(NumberOfEl_e - 2) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2) - ancho_mens / 2 * Cos(alfa)
        pol_canton(e).polyarray(NumberOfEl_e - 1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2) - ancho_mens / 2 * sin(alfa)
        pol_canton(e).polyarray(NumberOfEl_e) = 0
    Else
        pol_canton(e).polyarray(NumberOfEl_e - 2) = poste(iposte).pk_coordx + poste(iposte).descentramiento / 1000 * Cos(alfa + PI / 2)
        pol_canton(e).polyarray(NumberOfEl_e - 1) = poste(iposte).pk_coordy + poste(iposte).descentramiento / 1000 * sin(alfa + PI / 2)
        pol_canton(e).polyarray(NumberOfEl_e) = 0
    End If
End If


If (poste(iposte).tipo = eje_pf Or poste(iposte).tipo = eje_pf & " + " & semi_eje_aguj Or poste(iposte).tipo = eje_pf & " + " & eje_aguj) And poste(iposte).tunel = False Then
    '///
    '///Insertar atirantado anc_pf - eje_pf
    '///
                        
    Select Case poste(iposte - 1).lado
        Case "D"
            alfa_poste = (poste(iposte - 1).anguloeje - 90) * PI / 180
        Case "G"
            alfa_poste = (poste(iposte - 1).anguloeje + 90) * PI / 180
        End Select
    eje_poste = poste(iposte - 1).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril
    ini_aux(0) = pol_canton(e).polyarray(NumberOfEl_e - 2)
    ini_aux(1) = pol_canton(e).polyarray(NumberOfEl_e - 1)
    ini_aux(2) = 0
        
    centro(0) = poste(iposte - 1).pk_coordx + eje_poste * Cos(alfa_poste)
    centro(1) = poste(iposte - 1).pk_coordy + eje_poste * sin(alfa_poste)
    centro(2) = 0
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
    linea.Layer = "E-CABLEADO ELEVADO"
    linea.LinetypeScale = 0.5
    long_pf = linea.Length + long_pf
    '///
    '/// Añadir aislador
                
    PA(0) = (centro(0) + ini_aux(0)) / 2
    PA(1) = (centro(1) + ini_aux(1)) / 2
    PA(2) = 0
    alpa = Atn((centro(1) - ini_aux(1)) / (centro(0) - ini_aux(0)))
    Call cuadrado(PA, alpa, 0.5, 1, "E-HILO CONTACTO")
    '///
    '/// Añadir conexión rep02
    '///
    'alpa = Atn((pol_canton(e).polyarray(NumberOfEl_e - 1) - pol_canton(e).polyarray(NumberOfEl_e - 4)) / (pol_canton(e).polyarray(NumberOfEl_e - 2) - pol_canton(e).polyarray(NumberOfEl_e - 5)))
    'PA(0) = pol_canton(e).polyarray(NumberOfEl_e - 5) + Cos(alpa) * 3.375
    'PA(1) = pol_canton(e).polyarray(NumberOfEl_e - 4) + (sin(alpa) * 3.375)
    'PA(2) = 0
    'Set circulo = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddCircle(PA, 1)
    'Set arco = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddArc(PA, 0.5, alpa, alpa - 3.14)
    'arco.Layer = "E-HILO CONTACTO"
    '///
    '///Insertar atirantado eje_pf - anc_pf
    '///
    Select Case poste(iposte + 1).lado
        Case "D"
            alfa_poste = (poste(iposte + 1).anguloeje - 90) * PI / 180
        Case "G"
            alfa_poste = (poste(iposte + 1).anguloeje + 90) * PI / 180
    End Select
                
    eje_poste = poste(iposte + 1).implantacion + (ancho_via / 2) + (ancho_poste / 2) + ancho_carril
    centro(0) = poste(iposte + 1).pk_coordx + eje_poste * Cos(alfa_poste)
    centro(1) = poste(iposte + 1).pk_coordy + eje_poste * sin(alfa_poste)
    centro(2) = 0
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(centro, ini_aux)
    linea.Layer = "E-CABLEADO ELEVADO"
    linea.LinetypeScale = 0.5
    long_pf = linea.Length + long_pf
    '///
    '/// Añadir aislador
                
    PA(0) = (centro(0) + ini_aux(0)) / 2
    PA(1) = (centro(1) + ini_aux(1)) / 2
    PA(2) = 0
    alpa = Atn((centro(1) - ini_aux(1)) / (centro(0) - ini_aux(0)))
    Call cuadrado(PA, alpa, 0.5, 1, "E-HILO CONTACTO")
End If
'//
'//Recoger datos del CdPA
'//
polyCdpa(NumberOfEl_c - 2) = poste(iposte).pk_coordx + (eje_poste + ancho_poste) * Cos(alfa_poste)
polyCdpa(NumberOfEl_c - 1) = poste(iposte).pk_coordy + (eje_poste + ancho_poste) * sin(alfa_poste)
polyCdpa(NumberOfEl_c) = 0

'//
'//Recoger datos del feeder
'//
polyfeed(NumberOfEl_c - 2) = poste(iposte).pk_coordx + (eje_poste + ancho_poste + 1.5) * Cos(alfa_poste)
polyfeed(NumberOfEl_c - 1) = poste(iposte).pk_coordy + (eje_poste + ancho_poste + 1.5) * sin(alfa_poste)
polyfeed(NumberOfEl_c) = 0

ini_aux(0) = polyfeed(NumberOfEl_c - 2)
ini_aux(1) = polyfeed(NumberOfEl_c - 1)
ini_aux(2) = 0
centro(0) = poste(iposte).pk_coordx + eje_poste * Cos(alfa_poste)
centro(1) = poste(iposte).pk_coordy + eje_poste * sin(alfa_poste)
centro(2) = 0
Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_aux, centro)
linea.Layer = "E-FEEDER"
ancho_cuadrado = 1
escala_cuadrado = 1
escala_poste = 1
Call cuadrado(centro, alfa, ancho_cuadrado * escala_poste, ancho_cuadrado * escala_poste, "E-FEEDER")
intersectionpoint = polinea.IntersectWith(linea, acExtendNone)
linea.Delete
polinea.Delete
Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_aux, intersectionpoint)
linea.Layer = "E-FEEDER"
Dim insertionpnt(0 To 2) As Double


Next iposte
long_hc = long_hc + separar_lineas(NumberOfEl_i, i, 2, 1)
long_hc = long_hc + separar_lineas(NumberOfEl_e, e, 2, 1)

'Set poly = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(pol_canton(e).polyarray)
'poly.Layer = "E-HILO CONTACTO"
'long_hc = poly.Length + long_hc
'Sheets("Medicion").Cells(59, 5).Value = long_hc * 2
'Sheets("Medicion").Cells(59, 6).Value = long_hc * 1.05
Set poly2 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(polyCdpa)
poly2.Layer = "E-CDPA"
poly2.LinetypeScale = 1.25
long_cdpa = poly2.Length

'Set poly = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(pol_canton(e).polyarray)
'poly.Layer = "E-FEEDER"
'long_feeder = poly.Length + long_feeder

Set poly2 = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(polyfeed)
poly2.Layer = "E-FEEDER"
poly2.LinetypeScale = 0.25
long_feeder = poly2.Length
Sheets("Material").Cells(10, 11).Value = long_hc * n_hc * 1.05
Sheets("Material").Cells(11, 11).Value = long_cdpa * n_cdpa * 1.05
Sheets("Material").Cells(12, 11).Value = (long_hc + long_feeder * n_feed_pos) * 1.05
Sheets("Material").Cells(13, 11).Value = long_pf
GetObject(, "Autocad.Application").ActiveDocument.Regen acActiveViewport
End Sub
Function separar_lineas(Number, cont, canton_ini, canton_fin)
Dim polinea_in() As Double, polinea_out() As Double, polinea() As Double
lon = 0
ReDim polinea_in(1 To 9)
For g = 1 To 6
    polinea_in(g) = pol_canton(cont).polyarray(g)
Next
polinea_in(g) = (pol_canton(cont).polyarray(g - 3) + pol_canton(cont).polyarray(g)) / 2
polinea_in(g + 1) = (pol_canton(cont).polyarray(g - 2) + pol_canton(cont).polyarray(g + 1)) / 2
polinea_in(g + 2) = 0
Set poly = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(polinea_in)
If canton_ini = 2 Then
    poly.Layer = "E-CABLEADO ELEVADO"
    poly.LinetypeScale = 0.5
Else
    poly.Layer = "E-HILO CONTACTO"
End If
lon = poly.Length + lon
ReDim polinea_out(1 To 9)

polinea_out(1) = pol_canton(cont).polyarray(Number - 2)
polinea_out(2) = pol_canton(cont).polyarray(Number - 1)
polinea_out(3) = pol_canton(cont).polyarray(Number)
polinea_out(4) = pol_canton(cont).polyarray(Number - 5)
polinea_out(5) = pol_canton(cont).polyarray(Number - 4)
polinea_out(6) = pol_canton(cont).polyarray(Number - 3)
polinea_out(7) = (pol_canton(cont).polyarray(Number - 5) + pol_canton(cont).polyarray(Number - 8)) / 2
polinea_out(8) = (pol_canton(cont).polyarray(Number - 4) + pol_canton(cont).polyarray(Number - 7)) / 2
polinea_out(9) = 0

Set poly = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(polinea_out)
If canton_fin = 2 Then
    poly.Layer = "E-CABLEADO ELEVADO"
    poly.LinetypeScale = 0.5
Else
    poly.Layer = "E-HILO CONTACTO"
End If
lon = poly.Length + lon
If Number - 6 >= 6 Then
    ReDim polinea(1 To Number - 6)
    polinea(1) = polinea_in(7)
    polinea(2) = polinea_in(8)
    polinea(3) = polinea_in(9)
    For g = 4 To Number - 9
        polinea(g) = pol_canton(cont).polyarray(g + 3)
    Next
    polinea(Number - 8) = polinea_out(7)
    polinea(Number - 7) = polinea_out(8)
    polinea(Number - 6) = polinea_out(9)
    Set poly = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddPolyline(polinea)
    poly.Layer = "E-HILO CONTACTO"
    lon = poly.Length + lon
End If
separar_lineas = lon
End Function

Public Sub borrar(capa)

Dim CurrLayer As AcadLayer
Dim SSet As AcadSelectionSet
Dim intCode(1) As Integer
Dim varData(1) As Variant


For Each CurrLayer In GetObject(, "Autocad.Application").ActiveDocument.Layers
    If CurrLayer.Name = capa Then
        intCode(0) = 8: varData(0) = CurrLayer.Name 'only select items on layer "B"
        intCode(1) = 67: varData(1) = 0 'only select items in modelspace - error without this filter
    On Error Resume Next
    GetObject(, "Autocad.Application").ActiveDocument.SelectionSets.Item(CurrLayer.Name).Delete
    On Error GoTo 0
    Set SSet = GetObject(, "Autocad.Application").ActiveDocument.SelectionSets.Add(CurrLayer.Name)
    SSet.Select acSelectionSetAll, , , intCode, varData
    SSet.Highlight True
    SSet.Erase
    SSet.Delete
    CurrLayer.Delete
    End If
Next

End Sub
Sub Dibujar_PK()
Dim accapa As AcadLayer
Dim pk_ini, pk_actual As Double
Dim coord_pk As Variant
Dim linea As AcadLine
Dim insertionpnt(0 To 2) As Double
Dim ini_linea(0 To 2) As Double
Dim fin_linea(0 To 2) As Double
pk_ini = 0
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-PK")
Dim Atributos As Variant
Dim bloque_PK As AcadBlockReference
'pk_ini = 291.5858
pk_actual = pk_ini + (Int(pk_ini / 1000) * 1000 - pk_ini) - Int((Int(pk_ini / 1000) * 1000 - pk_ini) / 20) * 20
Do While pk_actual < (poli(num_lineas_total).dist_acum)
    coord_pk = coordenadas_pk(pk_actual)
    ini_linea(0) = coord_pk(0) + 1 * Cos(coord_pk(2) - PI / 2)
    ini_linea(1) = coord_pk(1) + 1 * sin(coord_pk(2) - PI / 2)
    ini_linea(2) = 0
    fin_linea(0) = coord_pk(0) + 1 * Cos(coord_pk(2) + PI / 2)
    fin_linea(1) = coord_pk(1) + 1 * sin(coord_pk(2) + PI / 2)
    fin_linea(2) = 0
    Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, fin_linea)
    linea.Layer = "E-PK"
    If (pk_actual - Int(pk_actual / 100) * 100) = 0 Then
        Set bloque_PK = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(fin_linea, cadena_ruta & "pk.dwg", 1#, 1#, 1#, coord_pk(2) + PI / 2)
        bloque_PK.Layer = "E-PK"
        Atributos = bloque_PK.GetAttributes
        Atributos(0).TextString = formato_pk_trazado(pk_actual)
    End If
    pk_actual = pk_actual + 20
Loop
End Sub
Function coordenadas_pk(pk As Double) As Variant
    Dim coord_aux(0 To 2) As Variant
    Dim ipoli As Integer
    ipoli = 1
    Do While poli(ipoli).dist_acum < pk
        ipoli = ipoli + 1
    Loop
    coord_aux(0) = poli(ipoli - 1).coordx + (pk - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
    coord_aux(1) = poli(ipoli - 1).coordy + (pk - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
    coord_aux(2) = poli(ipoli - 1).angulo_post * PI / 180
    coordenadas_pk = coord_aux
End Function
Function formato_pk_trazado(pk As Double) As Variant
Dim ceros As String
If (pk - 1000 * (Int(pk / 1000))) < 100 Then
    If (pk - 1000 * (Int(pk / 1000))) < 10 Then
        ceros = "00"
    Else
        ceros = "0"
    End If
Else
    ceros = ""
End If
formato_pk_trazado = Int(pk / 1000) & "+" & ceros & (Round(pk - 1000 * Int(pk / 1000), decimales_pk))
End Function
Sub dibujar_datos_trazado()
Dim accapa As AcadLayer
Dim ini_linea(0 To 2) As Double
Dim fin_linea(0 To 2) As Double
Dim insertionpnt(0 To 2) As Double
Dim objPoli As AcadLWPolyline
Dim dist_trazado, alfa_bloque As Double
Dim linea As AcadLine
Dim capa As String
Dim cadena_trazado As String
Dim bloque_trazado As AcadBlockReference
Dim itrazado, jtrazado As Integer
Dim Atributos As Variant
Dim texto, texto_centro As String
Call Obtener_excel_trazado
Call Obtener_excel_pks
Call pks_trazado
dist_trazado = 25
Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-Aux")
For itrazado = 1 To num_traz
    For jtrazado = 1 To 4
        If jtrazado = 3 And trazado(itrazado).col(4).pk = 0 Then
            GoTo sigtraz
        End If
        If jtrazado = 4 And itrazado <> num_traz Then
            If trazado(itrazado).col(jtrazado).pk = trazado(itrazado + 1).col(1).pk Then
                GoTo sigtraz
            End If
        End If
        If jtrazado = 1 Then
            If trazado(itrazado).col(jtrazado).pk = trazado(itrazado).col(jtrazado + 1).pk Then
                GoTo sigcol
            End If
        End If
        ini_linea(0) = trazado(itrazado).col(jtrazado).pkx
        ini_linea(1) = trazado(itrazado).col(jtrazado).pky
        ini_linea(2) = 0
        fin_linea(0) = trazado(itrazado).col(jtrazado).pkx + dist_trazado * Cos(trazado(itrazado).col(jtrazado).alfa - PI / 2)
        fin_linea(1) = trazado(itrazado).col(jtrazado).pky + dist_trazado * sin(trazado(itrazado).col(jtrazado).alfa - PI / 2)
        fin_linea(2) = 0
        Set linea = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.AddLine(ini_linea, fin_linea)
        Set accapa = GetObject(, "Autocad.Application").ActiveDocument.Layers.Add("E-TRAZADO TERRENO") '///
        linea.Layer = "E-TRAZADO TERRENO"
        Set bloque_trazado = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(ini_linea, cadena_ruta & "punta_flecha_vacia.dwg", 0.5, 0.5, 0.5, trazado(itrazado).col(jtrazado).alfa - PI / 2)
        bloque_trazado.Layer = "E-TRAZADO TERRENO"
        If itrazado <> num_traz Or jtrazado <> 4 Then
            Set bloque_trazado = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(fin_linea, cadena_ruta & "punta_flecha.dwg", 2#, 2#, 2#, trazado(itrazado).col(jtrazado).alfa)
            bloque_trazado.Layer = "E-TRAZADO TERRENO"
        End If
        If itrazado <> 1 Or jtrazado <> 1 Then
            Set bloque_trazado = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(fin_linea, cadena_ruta & "punta_flecha.dwg", 2#, 2#, 2#, trazado(itrazado).col(jtrazado).alfa + PI)
            bloque_trazado.Layer = "E-TRAZADO TERRENO"
        End If
        Set bloque_trazado = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(fin_linea, cadena_ruta & "trazado.dwg", 1#, 1#, 1#, trazado(itrazado).col(jtrazado).alfa + PI / 2)
        bloque_trazado.Layer = "E-TRAZADO TERRENO"
        Atributos = bloque_trazado.GetAttributes
        Select Case jtrazado
            Case 1
                texto = "ORP1 Km "
                texto_centro = "Lrp = " & Round(trazado(itrazado).col(jtrazado + 1).pk - trazado(itrazado).col(jtrazado).pk, 0) & "m"
            Case 2
                texto = "FRP1 Km "
                If trazado(itrazado).col(jtrazado).pk = trazado(itrazado).col(jtrazado - 1).pk Then
                    texto = ""
                End If
                texto_centro = "Rayon = " & Round(trazado(itrazado).radio, 0) & "m  Devers = " & Round(trazado(itrazado).devers, 0) & "mm"
            Case 3
                texto = "FRP2 Km "
                texto_centro = "Lrp = " & Round(trazado(itrazado).col(jtrazado + 1).pk - trazado(itrazado).col(jtrazado).pk, 0) & "m"
            Case 4
                texto = "ORP2 Km "
                texto_centro = "ALIGNEMENT"
        End Select
        
        Select Case trazado_trabajo
            Case True
                Atributos(0).TextString = texto & formato_pk_trazado(trazado(itrazado).col(jtrazado).pk)
            Case False
                Atributos(0).TextString = texto & convertir_pk_FT("lineal_terreno", trazado(itrazado).col(jtrazado).pk)
        End Select
        insertionpnt(0) = trazado(itrazado).col(jtrazado).centrox + dist_trazado * Cos(trazado(itrazado).col(jtrazado).alfa_centro - PI / 2)
        insertionpnt(1) = trazado(itrazado).col(jtrazado).centroy + dist_trazado * sin(trazado(itrazado).col(jtrazado).alfa_centro - PI / 2)
        insertionpnt(2) = 0
        Set bloque_trazado = GetObject(, "Autocad.Application").ActiveDocument.ModelSpace.InsertBlock(insertionpnt, cadena_ruta & "trazado_centro.dwg", 1#, 1#, 1#, trazado(itrazado).col(jtrazado).alfa_centro)
        bloque_trazado.Layer = "E-TRAZADO TERRENO"
        Atributos = bloque_trazado.GetAttributes
        Atributos(0).TextString = texto_centro
sigcol:
    Next
sigtraz:
Next
End Sub
Sub Obtener_excel_trazado()
Dim fila, colu As Integer
    With Sheets("Trazado")
        fila = 3
        While Not IsEmpty(.Cells(fila, 2).Value)
            For colu = 3 To 6
                trazado(fila - 2).col(colu - 2).pk = Round(.Cells(fila, colu).Value, 3)
            Next
            If .Cells(fila, 2) >= 0 Then
                trazado(fila - 2).radio = Round(.Cells(fila, 2).Value, 0)
            ElseIf .Cells(fila, 2) < 0 Then
                trazado(fila - 2).radio = -Round(.Cells(fila, 2).Value, 0)
            End If
            trazado(fila - 2).devers = Round(.Cells(fila, 8).Value, 0)
            fila = fila + 1
        Wend
    End With

num_traz = fila - 3
End Sub
Sub pks_trazado()
Dim itrazado, jtrazado, ipoli As Integer
For itrazado = 1 To num_traz
    For jtrazado = 1 To 4
        ipoli = 1
        Do While poli(ipoli).dist_acum < trazado(itrazado).col(jtrazado).pk
            ipoli = ipoli + 1
        Loop
        trazado(itrazado).col(jtrazado).pkx = poli(ipoli - 1).coordx + (trazado(itrazado).col(jtrazado).pk - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
        trazado(itrazado).col(jtrazado).pky = poli(ipoli - 1).coordy + (trazado(itrazado).col(jtrazado).pk - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
        trazado(itrazado).col(jtrazado).alfa = poli(ipoli - 1).angulo_post * PI / 180
            If jtrazado <> 1 Then
                trazado(itrazado).col(jtrazado - 1).pk_centro_posterior = trazado(itrazado).col(jtrazado - 1).pk + (trazado(itrazado).col(jtrazado).pk - trazado(itrazado).col(jtrazado - 1).pk) / 2
                ipoli = 1
                Do While poli(ipoli).dist_acum < trazado(itrazado).col(jtrazado - 1).pk_centro_posterior
                    ipoli = ipoli + 1
                Loop
                trazado(itrazado).col(jtrazado - 1).alfa_centro = poli(ipoli - 1).angulo_post * PI / 180
                trazado(itrazado).col(jtrazado - 1).centrox = poli(ipoli - 1).coordx + (trazado(itrazado).col(jtrazado - 1).pk_centro_posterior - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
                trazado(itrazado).col(jtrazado - 1).centroy = poli(ipoli - 1).coordy + (trazado(itrazado).col(jtrazado - 1).pk_centro_posterior - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
            End If
            If jtrazado = 1 And itrazado <> 1 Then
                trazado(itrazado - 1).col(4).pk_centro_posterior = trazado(itrazado - 1).col(4).pk + (trazado(itrazado).col(jtrazado).pk - trazado(itrazado - 1).col(4).pk) / 2
                ipoli = 1
                Do While poli(ipoli).dist_acum < trazado(itrazado - 1).col(4).pk_centro_posterior
                    ipoli = ipoli + 1
                Loop
                trazado(itrazado - 1).col(4).alfa_centro = poli(ipoli - 1).angulo_post * PI / 180
                trazado(itrazado - 1).col(4).centrox = poli(ipoli - 1).coordx + (trazado(itrazado - 1).col(4).pk_centro_posterior - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180) ' - poli(ipoli - 1).dist_acum) * Cos(poli(ipoli - 1).angulo_post * PI / 180)
                trazado(itrazado - 1).col(4).centroy = poli(ipoli - 1).coordy + (trazado(itrazado - 1).col(4).pk_centro_posterior - poli(ipoli - 1).dist_acum) * sin(poli(ipoli - 1).angulo_post * PI / 180)
            End If
    Next jtrazado
Next itrazado
End Sub
Function convertir_pk_FT(conv As String, pk As Double) As String
Dim ipk As Long
Dim ceros As String
Dim beta As Double
Select Case conv
    Case "lineal_terreno"
        ipk = 0
        If 55453.6631 <= pk And pk < 56453.5677 Then
            algo = True
        Else
            Do While matriz_pk(ipk) < pk
                ipk = ipk + 1
            Loop
            ipk = ipk - 1
        End If
        If algo = True Then
            If (pk - 1000 * (Int(pk / 1000))) < 100 Then
                If (pk - 1000 * (Int(pk / 1000))) < 10 Then
                    ceros = "00"
                Else
                    ceros = "0"
                End If
            Else
                ceros = ""
            End If
            convertir_pk_FT = "55bis" & "+" & ceros & Round((pk - 55453.6631), decimales_pk)
        Else
            convertir_pk_FT = formato_pk_trazado(1000 * CDbl(ipk) + pk - matriz_pk(ipk))
        End If
    Case "terreno_lineal"
        ipk = 0
        Do While ipk < Int(pk / 1000)
            ipk = ipk + 1
        Loop
        ipk = ipk - 1
        convertir_pk_FT = formato_pk_trazado(matriz_pk(ipk) + pk - 1000 * Int(pk / 1000))
End Select
End Function
Sub Obtener_excel_pks()
Dim fila As Integer
    With Sheets("Pk real")
        fila = 2
        While Not IsEmpty(.Cells(fila, 2).Value)
            matriz_pk(.Cells(fila, 1).Value) = .Cells(fila, 2).Value
            fila = fila + 1
        Wend
    End With

End Sub
