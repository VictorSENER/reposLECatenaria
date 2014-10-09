Attribute VB_Name = "principal"
'//
'// declaración de variables publicas
'//
Public dist_va_max As Double, dist_max_canton As Double, va_max_sm As Double, va_max_tunel As Double
Public inc_norm_va As Double, va_max As Double, inicio As Double, r_re As Double
Public pol As Integer
'Public Const strDB = "C:\Users\23370\Documents\Proyectos\D223041 - SiReCa\DR_PLANOS\BBDD\Base de datos.accdb"
Public strDB
Public nombre_cat As String
Sub principal(pk_inicio, pk_final, nombre_cat)
Dim texto As String
Dim fichero As String
Dim text As Scripting.TextStream
strDB = Environ("SIRECA_HOME") & "\database\db.Accdb"
Set a_text = CreateObject("Scripting.fileSystemObject")
i = 1
While Mid(Application.ActiveWorkbook.FullName, i, 1) <> "."
    i = i + 1
Wend



Call cargar.datos_lac(nombre_cat)
ventoso = viento.viento
pol = pol + 3
b = 0
Call formato.lenguaje(idioma)
a = 4
h = 10
polpri = 3
VB = replanteo(pk_inicio, h, a, r_re, dist_va_max, inc_norm_va, va_max_tunel, dist_max_canton, va_max_sm, ventoso, polpri, nombre_cat, pol)
While VB(3) < pk_final

    
VB = replanteo(VB(0), VB(1), VB(2), r_re, dist_va_max, inc_norm_va, va_max_tunel, dist_max_canton, va_max_sm, ventoso, polpri, nombre_cat, pol)
    Set text = a_text.CreateTextFile(Mid(Application.ActiveWorkbook.FullName, 1, i) & "proges")
    text.WriteLine VB(3) & " / " & pk_final
    text.Close
    'principal = sheets("Replanteo").Cells(h, 33).Value
Wend


Call revisar_agujas.revisar_agujas


Call CAD.posicion                                          ' muestra la posición del poste (para autocad)
Call pk_real.convertir_LT                                      ' convertir el PK lineal a PK de trazado
Call comentarios.puntos_singulares                                         ' muestra todos los puntos singulares del trazado
'Call cad.esfuerzo
'If va_max = va_max_sm Then
    Call cantonamiento.canton_final(nombre_cat, pk_final)                                            ' distribución de los cantones de catenaria
'End If
Call descentramiento.desc(nombre_cat)                                              ' elige el descentramiento de cada poste
Call comentarios.conexiones(nombre_cat)
Call comentarios.implantacion(nombre_cat)
Call altura.altura(nombre_cat)                                            ' calcula la altura del hilo de contacto
Call pendolado_2hc_VP.pendolado_columna(nombre_cat)

Call momento.momento(nombre_cat, 10, False)

Call eleccion.postes(nombre_cat, 10, False)
Call eleccion.cimentaciones(nombre_cat, idioma, "desmonte", 10, False)
'Call im_pend(fin)

Call num_postes.postes(nombre_cat)
'Call pendolado.pendolado(nombre_cat, ruta_replanteo) ' numerar los postes segun PK trazado
Call revision.revision(nombre_cat, ventoso)
'Call Formato.Formato(idioma)

End Sub
'//
'// Función principal. Es la responsable de la rutina general y de la comunicación con VB Studio
'//
Function replanteo(inicioVB, hVB, aVB, r_reVB, dist_va_maxVB, inc_norm_vaVB, va_max_tunelVB, dist_max_cantonVB, va_max_smVB, ventosoVB, poliVB, nombre_catVB, pol) As Long()
'//
'// Recolección de datos
'//
inicio = inicioVB
h = hVB
a = aVB
r_re = r_reVB
dist_va_max = dist_va_maxVB
inc_norm_va = inc_norm_vaVB
va_max_tunel = va_max_tunelVB
va_max = va_maxVB
dist_max_canton = dist_max_cantonVB
va_max_sm = va_max_smVB
ventoso = ventosoVB
pol = poliVB
nombre_cat = nombre_catVB
marcador = 0
On Error Resume Next
While inicio > ventoso(pol - 1)

Wend

If ventoso(pol - 1) >= Sheets("Replanteo").Cells(h, 33).Value Then
        Sheets("Vano").Range("A3:E20").ClearContents
        Call tabla_vanos.tabla_vanos(nombre_cat, pol, ventoso)
Else
        pol = pol + 3
        Sheets("Vano").Range("A3:E20").ClearContents
        Call tabla_vanos.tabla_vanos(nombre_cat, pol, ventoso)
End If

aqui:
va_max = Sheets("Vano").Cells(3, 1).Value
'//
'// Inicializar variable al inicio de la rutina
'//
If h = 10 Then
    Sheets("Replanteo").Cells(10, 33) = inicio
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

'If h > 16 Then
       Call regulacion.regulacion(h, a)
End If
Call punto_singular.sing(h, a, k)
Call punto_singular.sing1(h, a, marcador, 0)


h = h + 2
Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h - 1, 4) + Sheets("Replanteo").Cells(h - 2, 33)
'Call radio.radio1(h)
'//
'// Declaración de variable y comunicación con VB Studio
'//
Dim X(3) As Long
    X(0) = inicio
    X(1) = h
    X(2) = a
    X(3) = Sheets("Replanteo").Cells(h, 33).Value
replanteo = X
End Function

Function AutoCAD(pos, eti, va, fle, des, impl, alt, can, dat, con, pro, pen, alt_cat, sin, lin, fila_ini, fila_fin, ruta_autocad) As Long
nombre_catVB = Sheets("Replanteo").Cells(1, 1).Value
cadena_ruta = dibujar.seleccionar_polilinea(ruta_autocad)
Call dibujar.Obtener_datos_Excel(fila_ini, fila_fin)
Call dibujar.Encontrar_coordenadas_pk
Call cargar.datos_lac(nombre_catVB)
Call dibujar_trazado.Dibujar_PK
Call Obtener_excel_pks
If pos = True Then
    Call dibujar.dibujar_postes(cadena_ruta)
    Call dibujar.borrar("E-AUX")
End If
If eti = True Then
    Call dibujar.dibujar_etiquetas(cadena_ruta)
End If
If va = True Then
    Call dibujar.dibujar_vanos(cadena_ruta)
End If
If fle = True Then
    Call dibujar.dibujar_vanos(cadena_ruta)
End If
If fle = True Then
    Call dibujar.dibujar_flechas(cadena_ruta)
End If
If des = True Then
    Call dibujar.dibujar_descentramientos(cadena_ruta)
End If
If impl = True Then
    Call dibujar.dibujar_implantacion(cadena_ruta)
End If
If alt = True Then
    Call dibujar.dibujar_alturaHC(cadena_ruta)
End If
If can = True Then
    Call dibujar.dibujar_cantones(cadena_ruta)
    Call dibujar.borrar("E-AUX")
End If
If dat = True Then
    Call dibujar.dibujar_datos_poste(cadena_ruta)
End If
If con = True Then
    Call dibujar.dibujar_conexion(cadena_ruta)
End If
If pro = True Then
    Call dibujar.teccion(cadena_ruta)
End If
If pen = True Then
    Call dibujar.dibujar_pendola(cadena_ruta)
End If
If alt_cat = True Then
    Call dibujar.dibujar_alt_cat(cadena_ruta)
End If
If sin = True Then
    Call dibujar.dibujar_singular(cadena_ruta)
End If
If lin = True Then
    Call dibujar.dibujar_linea
End If

acaddoc.Documents.Close
acaddoc.Quit
'AcadDoc = Nothing
End Function
