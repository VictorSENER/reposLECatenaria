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
Public pol As Integer
Public strDB As String
'Public catenaria As String

'//////////////////////////////////////////////////////////////////////////////
'// FASE 1: CÁLCULO DEL CUADERNO DE REPLANTEO
'//////////////////////////////////////////////////////////////////////////////

Sub calculoReplanteo( pkIni As Long, pkFin As Long, catenaria As String )
	
	Dim i As Integer
	Dim cont_ As Integer
	Dim cont_carp As Integer
	Dim b As Integer
	Dim a As Integer
	Dim h As Integer
	Dim polpri As Integer

	Dim texto As String
	Dim fichero As String
	Dim nom_archivo As String
	Dim dir_archivo As String
	
	Dim text As Scripting.TextStream
	
	strDB = Environ("SIRECA_HOME") & "\database\db.Accdb"
	Set a_text = CreateObject("Scripting.fileSystemObject")
	
	i  = 1
	cont_ = 1
	
	While Mid(Application.ActiveWorkbook.Name, i, 1) <> "_" Or cont_ <= 2
		i = i + 1
		If Mid(Application.ActiveWorkbook.Name, i, 1) = "_" Then
			cont_ = cont_ + 1
		End If
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
	
	dir_archivo = Mid(Application.ActiveWorkbook.FullName, 1, cont_carp) & nom_archivo & ".progress"
	
	Call cargar.datos_lac(catenaria)
	
	ventoso = viento.viento
	pol = pol + 3
	b = 0
	Call formato.lenguaje(idioma)
	a = 4
	h = 10
	polpri = 3
	VB = replanteo(pkIni, h, a, r_re, dist_va_max, inc_norm_va, va_max_tunel, dist_max_canton, va_max_sm, ventoso, polpri, catenaria, pol)
	While VB(3) < pkFin

	VB = replanteo(VB(0), VB(1), VB(2), r_re, dist_va_max, inc_norm_va, va_max_tunel, dist_max_canton, va_max_sm, ventoso, polpri, catenaria, pol)
		Set text = a_text.CreateTextFile(dir_archivo)
		text.WriteLine "1" & "/" & "14" & "/" & "Replanteo de los postes" & "/" & VB(3) & "/" & pkFin
		text.Close
	Wend


	Call revisar_agujas.revisar_agujas

	'//Muestra la posición del poste (para autocad)
	Call CAD.posicion
	
	'//Convertir el PK lineal a PK de trazado
	Call pk_real.convertir_LT
	
	'//Muestra todos los puntos singulares del trazado
	Call comentarios.puntos_singulares                                         

	'//Distribución de los cantones de catenaria
	Call cantonamiento.canton_final(catenaria, pkFin)                                            
	
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

End Sub

'//
'// Función principal. Es la responsable de la rutina general y de la comunicación con VB Studio
'//
Function replanteo(inicioVB, hVB, aVB, r_reVB, dist_va_maxVB, inc_norm_vaVB, va_max_tunelVB, dist_max_cantonVB, va_max_smVB, ventosoVB, poliVB, catenariaVB, pol) As Long()
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
	catenaria = catenariaVB
	marcador = 0
	On Error Resume Next
	While inicio > ventoso(pol - 1)

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
		Call regulacion.regulacion(h, a)
	End If
	Call punto_singular.sing(h, a, k)
	Call punto_singular.sing1(h, a, marcador, 0)


	h = h + 2
	Sheets("Replanteo").Cells(h, 33).Value = Sheets("Replanteo").Cells(h - 1, 4) + Sheets("Replanteo").Cells(h - 2, 33)

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

'//////////////////////////////////////////////////////////////////////////////
'// FASE 2: DIBUJADO DEL PLANO DE REPLANTEO
'//////////////////////////////////////////////////////////////////////////////

Function AutoCAD(geoPost As Boolean, etiPost As Boolean, datPost As Boolean, vanos As Boolean, flechas As Boolean, descentramientos As Boolean, implantacion As Boolean, altHilo As Boolean, distCant As Boolean, conexiones As Boolean, protecciones As Boolean,  pendolado As Boolean, altCat As Boolean,  puntSing As Boolean,  cableado As Boolean,  datTraz As Boolean, ruta_autocad As String ) As Long
				  
	catenariaVB = Sheets("Replanteo").Cells(1, 1).Value
	cadena_ruta = dibujar.seleccionar_polilinea(ruta_autocad)
	Call dibujar.Obtener_datos_Excel(fila_ini, fila_fin)
	Call dibujar.Encontrar_coordenadas_pk
	Call cargar.datos_lac(catenariaVB)
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
	
End Function
