Attribute VB_Name = "principal"

'//////////////////////////////////////////////////////////////////////////////
'// Declaración de variables globales
'//////////////////////////////////////////////////////////////////////////////
Public dist_va_max As Double
Public dist_max_canton As Double
Public va_max_sm As Double
Public va_max_tunel As Double
Public inc_norm_va As Double
Public va_max As Double
Public inicio As Double
Public r_re As Double
Public strDB As String
Public dir_progress As String, dir_error As String
Public final As Double
Public a_text As scripting.FileSystemObject
Public text As scripting.TextStream


'//////////////////////////////////////////////////////////////////////////////
'// FASE 1: CÁLCULO DEL CUADERNO DE REPLANTEO
'//////////////////////////////////////////////////////////////////////////////

Sub calculoReplanteo(pkini As Double, _
                                          pkfin As Double, _
                                          catenaria As String)
        
        Dim i As Integer
        Dim cont_carp As Integer
        Dim nom_archivo As String
        Set a_text = CreateObject("Scripting.FileSystemObject")

        inicio = pkini
        final = pkfin
        strDB = Environ("SIRECA_HOME") & "\database\db.Accdb"

        
        i = 1
        
        While Mid(Application.ActiveWorkbook.Name, i, 1) <> "."
            i = i + 1
        Wend
        
        nom_archivo = Mid(Application.ActiveWorkbook.Name, 1, i - 1)

        i = 1
        cont_carp = 1
        While Mid(Application.ActiveWorkbook.FullName, i, 1) <> ""
            i = i + 1
            If Mid(Application.ActiveWorkbook.FullName, i, 1) = "\" Then
                cont_carp = i
            End If
        Wend
        
        dir_progress = Mid(Application.ActiveWorkbook.FullName, 1, cont_carp) & nom_archivo & ".progress"
        dir_error = Mid(Application.ActiveWorkbook.FullName, 1, cont_carp) & nom_archivo & ".error"
        Set text_progress = a_text.CreateTextFile(dir_progress)
        Set text_error = a_text.CreateTextFile(dir_error)
        text_progress.Close
        text_error.Close
        Set text_progress = Nothing
        Set text_error = Nothing
        
        Call cargar.datos_lac(catenaria)
        
        ventoso = viento.viento
        
        Call postes.ubicacion_postes(catenaria, ventoso)
        
        Call revisar_agujas.revisar_agujas

        '//Muestra la posición del poste (para autocad)
        Call CAD.posicion
        
        '//Convertir el PK lineal a PK de trazado
        Call pk_real.convertir_LT
        
        '//Muestra todos los puntos singulares del trazado
        Call comentarios.puntos_singulares

        '//Distribución de los cantones de catenaria
        Call cantonamiento.canton_final(catenaria)
        
        '//Elige el descentramiento de cada poste
        Call descentramiento.desc(catenaria)
        Call comentarios.conexiones(catenaria)
        Call comentarios.implantacion(catenaria)
        
        '//Calcula la altura del hilo de contacto
        Call altura.altura(catenaria)
        Call pendolado_2hc_VP.pendolado_columna(catenaria)
        Call momento.momento(catenaria, 10, False)
        Call eleccion.postes(catenaria, 10, False)
        Call eleccion.cimentaciones(catenaria, idioma, "desmonte", 10, False)
        Call num_postes.postes(catenaria)
        Call revision.revision(catenaria, ventoso)
        Set a_text = Nothing
        Set text = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////////////
'// FASE 3: DIBUJADO DEL PLANO DE REPLANTEO
'//////////////////////////////////////////////////////////////////////////////

Function dibujoReplanteo(pkini As Double, _
                                                  pkfin As Double, _
                                                  catenaria As String, _
                                                  geoPost As Boolean, _
                                                  etiPost As Boolean, _
                                                  datPost As Boolean, _
                                                  vanos As Boolean, _
                                                  flechas As Boolean, _
                                                  descentramientos As Boolean, _
                                                  implantacion As Boolean, _
                                                  altHilo As Boolean, _
                                                  distCant As Boolean, _
                                                  conexiones As Boolean, _
                                                  protecciones As Boolean, _
                                                  pendolado As Boolean, _
                                                  altCat As Boolean, _
                                                  puntSing As Boolean, _
                                                  cableado As Boolean, _
                                                  datTraz As Boolean, _
                                                  HDC As Boolean) As Long
        Dim ruta_autocad As String
        
        strDB = Environ("SIRECA_HOME") & "\database\db.Accdb"
        If HDC = True Then
            ruta_autocad = Environ("SIRECA_HOME") & "\core\blocks\HDC.dwg"
        Else
            i = 1
            While Mid(Application.ActiveWorkbook.FullName, i, 1) <> "."
                    i = i + 1
            Wend
            
            ruta_autocad = Mid(Application.ActiveWorkbook.FullName, 1, i - 1) & ".dwg"

        End If
        
        cadena_ruta = dibujar.seleccionar_polilinea(ruta_autocad)
        Call dibujar.Obtener_datos_Excel(pkini, pkfin)
        Call dibujar.Encontrar_coordenadas_pk
        Call cargar.datos_lac(catenaria)
        'Call dibujar.Dibujar_PK
        'Call Obtener_excel_pks
                
        If geoPost = True Then
                Call dibujar.dibujar_postes(cadena_ruta, HDC)
                Call dibujar.borrar("E-AUX")
        End If
        
        If etiPost = True Then
                Call dibujar.dibujar_etiquetas(cadena_ruta, HDC)
        End If
        
        If datPost = True Then
                Call dibujar.dibujar_datos_poste(cadena_ruta, HDC)
        End If
        
        If vanos = True Then
                Call dibujar.dibujar_vanos(cadena_ruta, HDC)
        End If
        
        If flechas = True Then
                Call dibujar.dibujar_flechas(cadena_ruta, HDC)
        End If
        
        If descentramientos = True Then
                Call dibujar.dibujar_descentramientos(cadena_ruta, HDC)
        End If
        
        If implantacion = True Then
                Call dibujar.dibujar_implantacion(cadena_ruta, HDC)
        End If
        
        If altHilo = True Then
                Call dibujar.dibujar_alturaHC(cadena_ruta, HDC)
        End If
        
        If distCant = True Then
                Call dibujar.dibujar_cantones(cadena_ruta, HDC)
                Call dibujar.borrar("E-AUX")
        End If
        
        If conexiones = True Then
                Call dibujar.dibujar_conexion(cadena_ruta, HDC)
        End If
        
        If protecciones = True Then
                Call dibujar.dibujar_proteccion(cadena_ruta, HDC)
        End If
        
        If pendolado = True Then
                Call dibujar.dibujar_pendola(cadena_ruta, HDC)
        End If
        
        If altCat = True Then
                Call dibujar.dibujar_alt_cat(cadena_ruta, HDC)
        End If
        
        If puntSing = True Then
                Call dibujar.dibujar_singular(cadena_ruta, HDC)
        End If
        
        If cableado = True Then
                Call dibujar.dibujar_linea
        End If
        
        If datTraz = True Then
                Call dibujar.dibujar_datos_trazado
        End If
               
End Function

'//////////////////////////////////////////////////////////////////////////////
'// FASE 4: DIBUJADO
'//////////////////////////////////////////////////////////////////////////////

Function dibujoMontaje(pk_ini As Double, _
                                                  pk_fin As Double, _
                                                  nombre_cat As String, _
                                                  print_pdf As Boolean, _
                                                  print_cad As Boolean) As Long

Dim accapa As AcadLayer, accapa1 As AcadLayer, acCapa2 As AcadLayer, acCapa3 As AcadLayer, acCapa4 As AcadLayer, acCapa5 As AcadLayer
strDB = Environ("SIRECA_HOME") & "\database\db.Accdb"
Call cargar.datos_lac(nombre_cat)
On Error Resume Next
'///
'/// se intenta abrir el autocad sin haberlo declarado, si da error es que el archivo aun no está abierto y se fuerza su apertura
'///
Set acaddoc = GetObject(, "AutoCAD.Application")
num_pag = 1
'AcadDoc.Visible = False
If Err Then
    Err.Clear
    Set acaddoc = AcadApplication
    acaddoc.Visible = True
    '///
    '/// por defecto al abrir el autocad se abre un fichero en blanco, se elimina y además se declara el fichero de la cartografia como topo
    '///
    acaddoc.Documents.Close
    acaddoc.Documents.Open "C:\Users\23370\Documents\Proyectos\D223041 - SiReCa\DR_PLANOS\Ejes\eje_Fez_Taza_3D.dwg"
    Set topo = acaddoc.Documents.Item(0)
    '///
    '/// si no se puede abrir el archivo especificado saldrá una pantala de error
    '///
    If Err Then
        MsgBox "Error opening AutoCAD"
        Exit Function
    End If
Else
    '///
    '/// autocad abierto previamente y solamente se debe declarar la cartografia como topo
    '///
    If acaddoc.Documents.Item(0).FullName <> "C:\Users\23370\Documents\Proyectos\D223041 - SiReCa\DR_PLANOS\Ejes\eje_Fez_Taza_3D.dwg" Then
        acaddoc.Documents.Close
        acaddoc.Documents.Open "C:\Users\23370\Documents\Proyectos\D223041 - SiReCa\DR_PLANOS\Ejes\eje_Fez_Taza_3D.dwg"
        Set topo = acaddoc.Documents.Item(0)
    Else
        Set topo = acaddoc.Documents.Item(0)
        topo.Activate
        topo.WindowState = acMax
    End If
End If

With Sheets("Replanteo")

'///
'/// se llama a la rutina Leer_Polilinea para recoger los datos del eje de la cartografía
'///

Call Leer_Polilinea
'///
'/// directorio donde se guardan los bloques a utilizar y directorio donde se guardará el resultado de la aplicación
'///
cadena_carnet = "C:\Users\23370\Documents\Proyectos\D223041 - SiReCa\DR_PLANOS\" & nombre_cat & "\"
cadena_general = "C:\Users\23370\Desktop\D50\"
Sheets("Material").Range("E2:E600").ClearContents
Sheets("Material").Range("K16:K20").ClearContents
Sheets("Material").Range("K14:K14").ClearContents
PDFfijo = cadena_general & Sheets("Replanteo").Cells(fila_ini, 1).Value & ".pdf"
count = 0
'///
'/// se realizarán todas las fichas de postes dentro del rango introducido por el usuario
'///!!!!!!!!! se debe cambiar fila por pk, como se realiza en todas partes
'///

fila = 10
While .Cells(fila, 33).Value < pk_ini
    fila = fila + 2
Wend
While .Cells(fila, 33).Value < pk_fin And Not IsEmpty(.Cells(fila, 33))

'For fila = fila_ini To fila_fin
Call codificacion.codificacion("montage", fila, cadena_general)
inicio:
'///
'/// Recoger datos
'///
If Sheets("Replanteo").Cells(fila, 16).Value = semi_eje_sla & " + " & anc_aguj Then
    tip_1 = semi_eje_sla
    tip_pf_1 = anc_aguj
ElseIf Sheets("Replanteo").Cells(fila, 16).Value = anc_sla_con & " + " & semi_eje_aguj Then
    tip_1 = anc_sla_con
    tip_pf_1 = semi_eje_aguj
ElseIf Len(Sheets("Replanteo").Cells(fila, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(fila, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(fila, 16).Value = anc_sm_sin) Then
    tip_1 = Mid(Sheets("Replanteo").Cells(fila, 16).Value, 15)
    tip_pf_1 = Mid(Sheets("Replanteo").Cells(fila, 16).Value, 1, 11)
Else
    tip_1 = Sheets("Replanteo").Cells(fila, 16).Value
    tip_pf_1 = Sheets("Replanteo").Cells(fila, 16).Value
End If
If Sheets("Replanteo").Cells(fila - 2, 16).Value = semi_eje_sla & " + " & anc_aguj Then
    tip_0 = semi_eje_sla
    tip_pf_0 = anc_aguj
ElseIf Sheets("Replanteo").Cells(fila - 2, 16).Value = anc_sla_con & " + " & semi_eje_aguj Then
    tip_0 = anc_sla_con
    tip_pf_0 = semi_eje_aguj
ElseIf Len(Sheets("Replanteo").Cells(fila - 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(fila - 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(fila - 2, 16).Value = anc_sm_sin) Then
    tip_0 = Mid(Sheets("Replanteo").Cells(fila - 2, 16).Value, 15)
    tip_pf_0 = Mid(Sheets("Replanteo").Cells(fila - 2, 16).Value, 1, 11)
Else
    tip_0 = Sheets("Replanteo").Cells(fila - 2, 16).Value
    tip_pf_0 = Sheets("Replanteo").Cells(fila - 2, 16).Value
End If
If Sheets("Replanteo").Cells(fila + 2, 16).Value = semi_eje_sla & " + " & anc_aguj Then
    tip_2 = semi_eje_sla
    tip_pf_2 = anc_aguj
ElseIf Sheets("Replanteo").Cells(fila + 2, 16).Value = anc_sla_con & " + " & semi_eje_aguj Then
    tip_2 = anc_sla_con
    tip_pf_2 = semi_eje_aguj
ElseIf Len(Sheets("Replanteo").Cells(fila + 2, 16).Value) > 14 And (Not Sheets("Replanteo").Cells(fila + 2, 16).Value = anc_sla_sin) And (Not Sheets("Replanteo").Cells(fila + 2, 16).Value = anc_sm_sin) Then
    tip_2 = Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 15)
    tip_pf_2 = Mid(Sheets("Replanteo").Cells(fila + 2, 16).Value, 1, 11)
Else
    tip_2 = Sheets("Replanteo").Cells(fila + 2, 16).Value
    tip_pf_2 = Sheets("Replanteo").Cells(fila + 2, 16).Value
End If
        '///
        '/// se abre una ficha concreta dependiendo de si el poste se debe instalar en la derecha o en la izquierda
        '///
conti:
        On Error Resume Next
        If Cells(fila, 38).Value = "Tunel" Then
            acaddoc.Documents.Open (cadena_carnet & "Carnet_montage_T.dwg")
            lado = Cells(fila, 30).Value
                
        ElseIf Cells(fila, 30).Value = "G" Then
            'Set AcadDoc = AcadApplication
            acaddoc.Documents.Open (cadena_carnet & "Carnet_montage_G.dwg")
            lado = "G"
        ElseIf Cells(fila, 30).Value = "D" Then
            'Set acadDoc = AcadApplication
            acaddoc.Documents.Open (cadena_carnet & "Carnet_montage_D.dwg")
            lado = "D"


        End If
        If Err Then
            Err.Clear
            acaddoc.Quit
            Set acaddoc = Nothing
            Set acaddoc = GetObject(, "AutoCAD.Application")
            Set acaddoc = AcadApplication
            acaddoc.Visible = True
            acaddoc.Documents.Open "C:\Users\23370\Documents\Proyectos\D223041 - SiReCa\DR_PLANOS\Ejes\eje_Fez_Taza_3D.dwg"
            Set topo = acaddoc.Documents.Item(0)
            GoTo conti
        End If
        topo.Activate
        '///
        '/// se declara este nuevo archivo dwg como carnet
        '///
        Set carnet = acaddoc.Documents.Item(1)
        'carnet.SaveAs (cadena_general & nombre_archivo & ".dwg")

        Set accapa = carnet.Layers.Add("E-MENSULA1")
        accapa.LineWeight = acLnWt000
        Set accapa1 = carnet.Layers.Add("E-MENSULA2")
        carnet.Linetypes.Load "LÍNEAS_OCULTAS", "acad.lin"
        accapa1.LineType = "LÍNEAS_OCULTAS"
        accapa1.LineWeight = acLnWt000
        
        Set acCapa2 = carnet.Layers.Add("E-COTAS")
        acCapa2.LineWeight = acLnWt000
        Set acCapa3 = carnet.Layers.Add("E-TERRENO")
        acCapa3.LineWeight = acLnWt000
        '///
        '/// se llama a la rutina Obtener_perfil para obtener
        '///
        If .Cells(fila, 38).Value <> "Tunel" And .Cells(fila, 38).Value <> "Marquesina" Then
        
            Call Dibujar_Seccion.Obtener_perfil
        End If
        '///
        '/// se llama a la rutina para dibujar las curvas de nivel circundantes a la vía
        '///
        Call Dibujar_Seccion.Dibujar_Seccion
      
            '///
            '/// se llama a la rutina Decidir_cimentacion para elegir si la cimentación debe ir en desmonte o en terraplén
            '///
            Call Dibujar_Seccion.Decidir_cimentacion(nombre_catVB, fila)
            '///
            '/// se llama a la rutina Dibujar elementos para dibujar el poste y la cimentación adecuada
            '///
             Call Dibujar_Seccion.Dibujar_elementos(fila)
         
        Call Dibujar_Seccion.dibujar_mensulas(fila)
        '///
        '/// se llama a la rutina Obtener_perfil para obtener
        '///
        Call Dibujar_Seccion.Escribir_Textos(fila, nombre_archivo, "", num_pag, tip_1)

        'End If
        '///
        '/// se llama a la rutina para guardar o no los archivos en el formato deseado
        '///
        Call Dibujar_Seccion.guardar_archivo(print_pdf, print_cad, nombre_archivo, nombre_archivo_fijo)
        '///
        '/// se comprueba si es necesario dibujar el poste de anclaje antes
        '///
        If (tip_1 = anc_sla_con Or tip_1 = anc_sm_con Or tip_1 = anc_sla_sin Or tip_1 = anc_sm_sin Or tip_1 = anc_aguj Or tip_pf_1 = anc_pf Or tip_pf_1 = anc_aguj) And (tip_2 = semi_eje_sm Or tip_2 = semi_eje_aguj Or tip_2 = semi_eje_sla Or tip_pf_2 = eje_pf Or tip_pf_2 = semi_eje_aguj) _
        Or Sheets("Replanteo").Cells(fila, 17).Value = "Anc. Feeder Alim." Or Sheets("Replanteo").Cells(fila, 17).Value = "Anc. CdPA et Feeder" Then
            normal = Sheets("Replanteo").Cells(fila, 47).Value
            num_pag = num_pag + 1
            acaddoc.Documents.Open (cadena_carnet & "Carnet_anclaje_i.dwg")
            Set carnet = acaddoc.Documents.Item(1)
            carnet.Linetypes.Load "CDPA", "acadiso.lin"
            carnet.Linetypes.Load "ACAD_ISO06W100", "acadiso.lin"
            Set accapa = carnet.Layers.Add("E-MENSULA1")
            accapa.LineWeight = acLnWt000
            Set accapa = carnet.Layers.Add("E-MENSULA2")
            accapa.LineWeight = acLnWt000
            Set acCapa2 = carnet.Layers.Add("E-COTAS")
            acCapa2.LineWeight = acLnWt000
            Set acCapa3 = carnet.Layers.Add("E-TERRENO")
            acCapa3.LineWeight = acLnWt000
            Set acCapa4 = carnet.Layers.Add("E-CDPA")
            acCapa4.LineWeight = acLnWt000
            acCapa4.LineType = "CDPA"
            acCapa4.Color = acBlue
            Set acCapa5 = carnet.Layers.Add("E-FEEDER")
            acCapa5.LineWeight = acLnWt000
            acCapa5.LineType = "ACAD_ISO06W100"
            alt_anc_hc = Sheets("Replanteo").Cells(fila + 1, 46).Value
            Call Dibujar_Seccion.dibujar_anclaje(fila, "i", alt_anc_hc)
            If Sheets("Replanteo").Cells(fila, 17).Value = "Anc. Feeder Alim." Or Sheets("Replanteo").Cells(fila, 17).Value = "Anc. CdPA et Feeder" Then
                GoTo otro_anclaje
            End If
            Call Dibujar_Seccion.Escribir_Textos(fila, nombre_archivo, "A", num_pag, tip_1)
            Call Dibujar_Seccion.guardar_archivo(print_pdf, print_cad, nombre_archivo & "A", nombre_archivo_fijo)
        End If
        '///
        '/// se comprueba si es necesario dibujar el poste de anclaje
        '///
        If (tip_1 = anc_sla_con Or tip_1 = anc_sm_con Or tip_1 = anc_aguj Or tip_1 = anc_sla_sin Or tip_1 = anc_sm_sin Or tip_pf_1 = anc_pf) And (tip_0 = semi_eje_sla Or tip_0 = semi_eje_sm Or tip_0 = semi_eje_aguj Or tip_pf_0 = eje_pf) _
        Or Sheets("Replanteo").Cells(fila, 17).Value = "Anc. Feeder Alim." Or Sheets("Replanteo").Cells(fila, 17).Value = "Anc. CdPA et Feeder" Then
            num_pag = num_pag + 1
            acaddoc.Documents.Open (cadena_carnet & "Carnet_anclaje_d.dwg")
            Set carnet = acaddoc.Documents.Item(1)
            carnet.Linetypes.Load "CDPA", "acadiso.lin"
            carnet.Linetypes.Load "ACAD_ISO06W100", "acadiso.lin"
            Set accapa = carnet.Layers.Add("E-MENSULA1")
            accapa.LineWeight = acLnWt000
            Set acCapa2 = carnet.Layers.Add("E-COTAS")
            acCapa2.LineWeight = acLnWt000
            Set acCapa3 = carnet.Layers.Add("E-TERRENO")
            acCapa3.LineWeight = acLnWt000
            Set acCapa4 = carnet.Layers.Add("E-CDPA")
            acCapa4.LineWeight = acLnWt000
            acCapa4.LineType = "CDPA"
            acCapa4.Color = acBlue
            Set acCapa5 = carnet.Layers.Add("E-FEEDER")
            acCapa5.LineWeight = acLnWt000
            acCapa5.LineType = "ACAD_ISO06W100"
otro_anclaje:
            alt_anc_hc = Sheets("Replanteo").Cells(fila - 1, 48).Value
            Call Dibujar_Seccion.dibujar_anclaje(fila, "d", alt_anc_hc)
            Call Dibujar_Seccion.Escribir_Textos(fila, nombre_archivo, "A", num_pag, tip_1)
            Call Dibujar_Seccion.guardar_archivo(print_pdf, print_cad, nombre_archivo & "-A", nombre_archivo_fijo)
            
        End If
              
fila = fila + 2
num_pag = num_pag + 1
'Next
Wend
End With

End Function

