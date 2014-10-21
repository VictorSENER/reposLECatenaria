Attribute VB_Name = "eleccion"
Sub postes(nombre_catVB)
    
    Sheets(8).Range("A1:AH10001").ClearContents
    
'//
'//LECTURA BASE DE DATOS
'//
    
    Call cargar.datos_acces(nombre_catVB)
 
    
'//
'//INSERCI�N DATOS EN HOJA ANEXA DE REPLANTEO
'//
    
    Dim oConn As ADODB.Connection
    Dim oRead As ADODB.Recordset
    Dim strDB, strSQL As String
    Dim strTabla As String
    Dim lngTablas As Long
    Dim i As Long
    'elegir uno de estas dos rutas al archivo Access
    strDB = "W:\223\D\D223041\CC_CALCULOS\SiReCa\Base de datos.accdb"
    'nombre de la tabla del archivo Access
    strTabla = "Postes"
    'crear la conexi�n
    Set oConn = New ADODB.Connection
    Set oRead = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oRead.Open strSQL, oConn
    'copiar datos a la hoja
    
    j = 1
    'mientras hayan registros
    While Not oRead.EOF
 
    If oRead.Fields(1).Value = adm_lin_poste And oRead.Fields(0).Value = tip_poste Then
    
    lngCampos = oRead.Fields.count
    For i = 0 To lngCampos - 1
    Sheets(8).Cells(j + 1, i + 1).Value = oRead.Fields(i).Value
    Next
  
    j = j + 1
    End If
    'saltar al siguiente registro
    oRead.MoveNext
    Wend
    
    'copiar r�tulos
    j = 1
    lngCampos = oRead.Fields.count
    For i = 0 To lngCampos - 1
    Sheets(8).Cells(j, i + 1).Value = oRead.Fields(i).Name
    Next
    'desconectar
    oRead.Close: Set oRead = Nothing
    oConn.Close: Set oConn = Nothing
    
'//
'//ELECCI�N POSTE EN FUNCI�N MOMENTO CALCULADO
'//
    
    z = 10
    i = 2

    While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
        If Sheets(1).Cells(z, 25).Value = "Tunnel" Or Sheets(1).Cells(z, 25).Value = "Pilier Viaducto" Then
        Else
        mom_poste_calc = Sheets(1).Cells(z, 19)
        alt_nenc_poste = Sheets(8).Cells(i, 9)
        If IsEmpty(Sheets(1).Cells(z, 16).Value) Then
            While Sheets(8).Cells(i, 17).Value < mom_poste_calc
                i = i + 1
            Wend
        Else
            While Sheets(8).Cells(i, 17).Value < mom_poste_calc Or alt_nenc_poste < 8
                i = i + 1
                alt_nenc_poste = Sheets(8).Cells(i, 9)
            Wend
        End If

        mom_poste_var = Sheets(8).Cells(i, 17)
        alt_nenc_poste = Sheets(8).Cells(i, 9)
        tip_poste = Sheets(8).Cells(i, 3)
x:
        i = 2
        While Sheets(8).Cells(i, 17) <> 0
            If Sheets(8).Cells(i, 17) < mom_poste_var And Sheets(8).Cells(i, 17) > mom_poste_calc Then
                mom_poste_var = Sheets(8).Cells(i, 17)
                alt_nenc_poste = Sheets(8).Cells(i, 10)
                tip_poste = Sheets(8).Cells(i, 3)
                GoTo x
            Else: i = i + 1
            End If
        Wend
        
'//
'//INSERECI�N POSTE EN REPLANTEO
'//
        
        Sheets(1).Cells(z, 35) = mom_poste_var
        Sheets(1).Cells(z, 36) = alt_nenc_poste
        Sheets(1).Cells(z, 18) = tip_poste
        End If
        z = z + 2
        i = 2
    Wend

End Sub
Sub cimentaciones(nombre_catVB)

    Sheets(9).Range("A1:AH10001").ClearContents
    
'//
'//LECTURA BASE DE DATOS
'//
    
    Call cargar.datos_acces(nombre_catVB)

'//
'//INSERCI�N DATOS EN HOJA ANEXA DE REPLANTEO
'//
    
    Dim oConn As ADODB.Connection
    Dim oRead As ADODB.Recordset
    Dim strDB, strSQL As String
    Dim strTabla As String
    Dim lngTablas As Long
    Dim i As Long
    'elegir uno de estas dos rutas al archivo Access
    strDB = "W:\223\D\D223041\CC_CALCULOS\SiReCa\Base de datos.accdb"
    'nombre de la tabla del archivo Access
    strTabla = "Macizos"
    'crear la conexi�n
    Set oConn = New ADODB.Connection
    Set oRead = New ADODB.Recordset
    oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source =" & strDB & ";"
    'consulta SQL
    strSQL = "SELECT * FROM " & strTabla & ""
    oRead.Open strSQL, oConn
    'copiar datos a la hoja
        
    j = 1
    'mientras hayan registros
    While Not oRead.EOF
 
    'se debe leer del programa
    If oRead.Fields(0).Value = tip_mac And oRead.Fields(1).Value = adm_lin_mac Then 'And oRead.Fields(2).Value = "desmonte" then
    
    lngCampos = oRead.Fields.count
    For i = 0 To lngCampos - 1
    Sheets(9).Cells(j + 1, i + 1).Value = oRead.Fields(i).Value
    Next
  
    j = j + 1
    End If
    'saltar al siguiente registro
    oRead.MoveNext
    Wend
    
    'copiar r�tulos
    j = 1
    lngCampos = oRead.Fields.count
    For i = 0 To lngCampos - 1
    Sheets(9).Cells(j, i + 1).Value = oRead.Fields(i).Name
    Next
    'desconectar
    oRead.Close: Set oRead = Nothing
    oConn.Close: Set oConn = Nothing


'//
'//C�LCULO CIMENTACI�N
'//

Dim desm_terrap_mac As String
Dim res_lat As Double
Dim res_arr As Double
Dim alt_tot_mac As Double
Dim mom_vuelco_pos As Double
Dim mom_vuelco_neg As Double

dist_ad_mac = 4 'metros desde el extremo del macizo hasta el final del terreno (se deber� recoger del perfil)
desn_ad_mac = 0.3 'metros de diferencia entre terrenos (se debe recoger del perfil)
incl_ad_mac = 10 'grados (o pendiente, 19�->1:3, 33�->1:1.5) de inclinaci�n del terreno (se debe recoger del perfil)

'valores impuestos por ADIF a variar si las condiciones son peores
coef_compres_base_desm = 6 ' valor a leer del sireca
alt_base_desm = 2 'valor a leer del sireca
coef_compres_base_terrap = 4 'valor a leer del sireca
alt_base_terrap = 2 'valor a leer del sireca
p_terr = 1400 'kg/m3 peso propio del terreno, valor a leer del sireca
and_roz_terr = 22 '�ngulo de rozamiento interno del terreno, valor a leer del sireca
tg_alpha = 0.005 'Valor m�ximo de inclinaci�n del macizo
ang_roz = 14 'Valor del SiReCa '�ngulo de rozamiento interno del terreno
p_esp_horm = 2200 'Valor obtenido del SiReCa (kg/m3)
p_esp_terr = 1400 'Valor obtenido del SiReCa (kg/m3)
cap_lat = 10000 'kg/m2 capacidad lateral del terreno, obtenido del SiReCa
'no se tiene en cuenta el peso del poste, puesto que lo que har�a es incrementar el mom_base (mom_vuelco), la elecci�n de la cimentaci�n se hace partiendo del momento del poste calculado, por lo que la cimentaci�n que se escojer� ser� la que soporte ese momento, si aplicasemos el peso del poste, el mom_vuelco incrementar�a.

i = 2
z = 10

Sheets(9).Cells(1, 16) = "mom_vuelco_mac"
Sheets(9).Cells(1, 17) = "fuerza_anc_mac"

While Sheets(9).Cells(i, 1) <> ""

mac_mac = Sheets(9).Cells(i, 1)

desm_terrap_mac = Sheets(9).Cells(i, 3)

    If mac_mac = "Paralelep�pedo" Then
                     
            'ACTIVAR CUANDO SE VINCULE CON EL PERFIL PARA ESCOGER DESMONTE O TERRAPL�N
            'If dist_ad_mac >= 3.5 And desn_ad_mac < 0.5 And incl_ad_mac <= 19 Then '
            '    desm_terrap_mac = "desmonte"
                
            '    Else: desm_terrap_mac = "terrapl�n"
            'End If
             
                If desm_terrap_mac = "desmonte" Then
                    l_ent_mac = Sheets(9).Cells(i, 5) 'ancho cimentaci�n en sentido perpendicular a la v�a
                    ancho_ent_mac = Sheets(9).Cells(i, 6) 'ancho cimentaci�n en sentido paralelo a la v�a
                                           
                    alt_ent_mac = Sheets(9).Cells(i, 8) ' altura enterrada del macizo
                    alt_nent_mac = Sheets(9).Cells(i, 12) 'altura no enterrada del macizo
                                                               
                    alt_tot_mac = alt_ent_mac + alt_nent_mac 'altura total del macizo
                                                                      
                    p_tot_mac = l_ent_mac * ancho_ent_mac * alt_tot_mac * p_esp_horm 'peso total macizo
                        
                        If alt_ent_mac <= 2 Then
                            coef_compres_var = coef_compres_base_desm * (alt_ent_mac / alt_base_desm)
                        Else: coef_compres_var = coef_compres_base_desm
                        End If
                        
                    mom_lat = (1000000 / 36) * tg_alpha * ancho_ent_mac * coef_compres_var * (alt_ent_mac ^ 3)
                    mom_base = p_tot_mac * ((l_ent_mac / 2) - (1 / 3000) * Sqr((2 * p_tot_mac) / (ancho_ent_mac * coef_compres_var * tg_alpha)))
            
                    coef_pond = mom_lat / mom_base
            
                        If coef_pond <= 1 Then
                            coef_pond = 0.4167 * ((mom_lat / mom_base) ^ 2) - 0.9167 * (mom_lat / mom_base) + 1.5
                        Else: coef_pond = 1
                        End If
            
                    mom_vuelco = (mom_lat + mom_base) / coef_pond
                    
                    If Sheets(9).Cells(i, 16) = "" Then
                                                                   
                        Sheets(9).Cells(i, 16) = mom_vuelco
                        
                    End If
                    
                
                ElseIf desm_terrap_mac = "terrapl�n" Then
                    l_ent_mac = Sheets(9).Cells(i, 5) 'ancho cimentaci�n en sentido perpendicular a la v�a
                    ancho_ent_mac = Sheets(9).Cells(i, 6) 'ancho cimentaci�n en sentido paralelo a la v�a
                    l_tot_mac = Sheets(9).Cells(i, 7) 'ancho cimentaci�n total en sentido perpendicular a la v�a
                    
                    alt_ent_mac = Sheets(9).Cells(i, 8) ' altura enterrada del macizo
                    alt_nent_mac = Sheets(9).Cells(i, 12) 'altura no enterrada del macizo
                                               
                    alt_tot_mac = alt_ent_mac + alt_nent_mac 'altura total del macizo
                                
                    p_tot_mac_1 = l_ent_mac * ancho_ent_mac * alt_tot_mac * p_esp_horm 'peso total macizo
                    p_tot_mac_2 = 0.5 * (l_tot_mac - l_ent_mac) * ancho_ent_mac * alt_tot_mac * p_esp_horm
                                
                        If alt_ent_mac <= 2 Then
                            coef_compres_var = coef_compres_base_terrap * (alt_ent_mac / alt_base_terrap)
                        Else: coef_compres_var = coef_compres_base_terrap
                        End If
                        
                    mom_lat_pos = (4000000 / 243) * tg_alpha * ancho_ent_mac * coef_compres_var * (alt_ent_mac ^ 3)
                    mom_base_pos = p_tot_mac_1 * (l_ent_mac / 2) + p_tot_mac_2 * (l_tot_mac - (2 / 3) * (l_tot_mac - l_ent_mac))
                    
                    coef_pond = mom_lat_pos / mom_base_pos
            
                        If coef_pond <= 1 Then
                            coef_pond = 0.4167 * ((mom_lat_pos / mom_base_pos) ^ 2) - 0.9167 * (mom_lat_pos / mom_base_pos) + 1.5
                        Else: coef_pond = 1
                        End If
            
                    mom_vuelco_pos = (mom_lat_pos + mom_base_pos) / coef_pond
                                       
                    param_z = 0.001 * (Sqr((2 * (p_tot_mac_1 + p_tot_mac_2)) / (tg_alpha * coef_compres_var * ancho_ent_mac)))
                    
                    mom_lat_neg = (4000000 / 243) * tg_alpha * ancho_ent_mac * coef_compres_var * (alt_ent_mac ^ 3)
                    mom_base_neg = p_tot_mac_1 * (l_tot_mac - (l_ent_mac / 2)) + p_tot_mac_2 * (((2 / 3) * (l_tot_mac - l_ent_mac)) - param_z) + (p_tot_mac_1 + p_tot_mac_2) * (1 / 3) * param_z
                    
                                    coef_pond = mom_lat_neg / mom_base_neg
            
                        If coef_pond <= 1 Then
                            coef_pond = 0.4167 * ((mom_lat_neg / mom_base_neg) ^ 2) - 0.9167 * (mom_lat_neg / mom_base_neg) + 1.5
                        Else: coef_pond = 1
                        End If
                    
                    mom_vuelco_neg = mom_base_neg / coef_pond
                    
                    mom_vuelco_neg = mom_vuelco_neg / 1.5
                    
                    mom_vuelco = MAX(mom_vuelco_pos, mom_vuelco_neg)
                    
                    If Sheets(9).Cells(i, 16) = "" Then
                                                                   
                        Sheets(9).Cells(i, 16) = mom_vuelco
                        
                    End If
                                       
                ElseIf desm_terrap_mac = "anclaje" Then
                    
                    l_ent_mac = Sheets(9).Cells(i, 5) 'ancho cimentaci�n en sentido perpendicular a la v�a
                    ancho_ent_mac = Sheets(9).Cells(i, 6) 'ancho cimentaci�n en sentido paralelo a la v�a
                                    
                    alt_ent_mac = Sheets(9).Cells(i, 8) ' altura enterrada del macizo
                    alt_nent_mac = Sheets(9).Cells(i, 12) 'altura no enterrada del macizo
                    
                    v_tot_mac = Sheets(9).Cells(i, 14)
                    v_terr = 2 * (alt_ent_mac * Tan(ang_roz * (3.1416 / 180)) * (l_ent_mac + ancho_ent_mac)) * alt_ent_mac + (4 / 3) * ((alt_ent_mac * Tan(ang_roz * (3.1416 / 180))) ^ 2) * alt_ent_mac
                    
                    res_lat = cap_lat * l_ent_mac * alt_ent_mac
                    res_arr = v_tot_mac * p_esp_horm + v_terr * p_esp_terr
                    
                    fuerza_anc = Sqr(2) * MIN(res_lat, res_arr)
                    
                    mom_vuelco = 0
                    Sheets(9).Cells(i, 17) = fuerza_anc
                    
                End If
            
            'a partir de este momento hay que desplazar el momento a la base del poste para saber cuanto vale, para ello es necesario conocer el tipo de poste y su altura.
                    
    ElseIf mac_mac = "Cil�ndrico" Then
    
        If incl_ad_mac > 33 Then
            'no se debe usar los macizos cil�ndricos
        End If
    
        'ACTIVAR CUANDO SE VINCULE CON EL PERFIL PARA ESCOGER DESMONTE O TERRAPL�N
            'If dist_ad_mac >= 3.5 And desn_ad_mac < 0.5 And incl_ad_mac <= 19 Then '
            '    desm_terrap_mac = "desmonte"
                
            '    Else: desm_terrap_mac = "terrapl�n"
            'End If
             
                If desm_terrap_mac = "desmonte" Then
                    diam_mac = Sheets(9).Cells(i, 15) 'ancho cimentaci�n en sentido perpendicular a la v�a
                        
                    alt_ent_mac = Sheets(9).Cells(i, 8) ' altura enterrada del macizo
                    alt_nent_mac = Sheets(9).Cells(i, 12) 'altura no enterrada del macizo
                    p_esp_horm = 2200 'Valor obtenido del SiReCa (kg/m3)
                                               
                    alt_tot_mac = alt_ent_mac + alt_nent_mac 'altura total del macizo
                                                  
                        If alt_ent_mac <= 2 Then
                            coef_compres_var = coef_compres_base_desm * (alt_ent_mac / alt_base_desm)
                        Else: coef_compres_var = coef_compres_base_desm
                        End If
                        
                    mom_lat = (1000000 / 36) * tg_alpha * coef_compres_var * diam_mac * (alt_ent_mac ^ 3)
                          
                    coef_pond = 1
            
                    mom_vuelco = (mom_lat) / coef_pond
                    
                    If Sheets(9).Cells(i, 16) = "" Then
                                                                   
                        Sheets(9).Cells(i, 16) = mom_vuelco
                        
                    End If
                
                ElseIf desm_terrap_mac = "terrapl�n" Then
                    diam_mac = Sheets(9).Cells(i, 15) 'ancho cimentaci�n en sentido perpendicular a la v�a
                        
                    alt_ent_mac = Sheets(9).Cells(i, 8) ' altura enterrada del macizo
                    alt_nent_mac = Sheets(9).Cells(i, 12) 'altura no enterrada del macizo
                    p_esp_horm = 2200 'Valor obtenido del SiReCa (kg/m3)
                                               
                    alt_tot_mac = alt_ent_mac + alt_nent_mac 'altura total del macizo
                                                  
                        If alt_ent_mac <= 2 Then
                            coef_compres_var = coef_compres_base_terrap * (alt_ent_mac / alt_base_terrap)
                        Else: coef_compres_var = coef_compres_base_terrap
                        End If
                        
                    mom_lat = (1000000 / 36) * tg_alpha * coef_compres_var * diam_mac * (alt_ent_mac ^ 3)
                          
                    coef_pond = 1
            
                    mom_vuelco = (mom_lat) / coef_pond
                    
                    If Sheets(9).Cells(i, 16) = "" Then
                                                                   
                        Sheets(9).Cells(i, 16) = mom_vuelco
                        
                    End If
                    
                ElseIf desm_terrap_mac = "anclaje" Then
                    
                    diam_mac = Sheets(9).Cells(i, 15) 'ancho cimentaci�n en sentido perpendicular a la v�a
                   
                    alt_ent_mac = Sheets(9).Cells(i, 8) ' altura enterrada del macizo
                    alt_nent_mac = Sheets(9).Cells(i, 12) 'altura no enterrada del macizo
                    
                    v_tot_mac = Sheets(9).Cells(i, 14)
                    v_terr = 3.1416 * (((alt_ent_mac ^ 3) / 3) * Tan(ang_roz * (3.1416 / 180)) + (diam_mac / 2) * (alt_ent_mac ^ 2) * Tan(ang_roz * (3.1416 / 180)))
                    
                    'Hay que revisar la resistencia lateral
                    'res_lat = cap_lat * diam_mac * 3.1416 * diam_mac * alt_ent_mac
                    res_arr = v_tot_mac * p_esp_horm + v_terr * p_esp_terr
                    
                    fuerza_anc = Sqr(2) * MIN(res_lat, res_arr)
                    
                    mom_vuelco = 0
                    Sheets(9).Cells(i, 17) = fuerza_anc
                    
                End If
            
            'a partir de este momento hay que desplazar el momento a la base del poste para saber cuanto vale, para ello es necesario conocer el tipo de poste y su altura.
    
    
    End If
    
    If Sheets(9).Cells(i, 16) = "" Then
                                                                                
    Sheets(9).Cells(i, 19) = mom_vuelco * (Sheets(1).Cells(z, 36) / (Sheets(1).Cells(z, 36) + alt_nent_mac + (2 / 3) * alt_ent_mac))
     
    Else
    
    Sheets(9).Cells(i, 19) = Sheets(9).Cells(i, 16)
     
    End If
      
    i = i + 1
   
    Wend
    
  
'//
'//ELECCI�N CIMENTACI�N EN FUNCI�N MOMENTO POSTE
'//
  
    z = 10
    i = 2
    
    While Not IsEmpty(Sheets(1).Cells(z, 33).Value)
        If Sheets(1).Cells(z, 25).Value = "Tunnel" Or Sheets(1).Cells(z, 25).Value = "Pilier Viaducto" Then
        Else
        mom_poste_var = Sheets(1).Cells(z, 19)
        
        While Sheets(9).Cells(i, 19) < mom_poste_var
            i = i + 1
        Wend
        
        mom_cim_var = Sheets(9).Cells(i, 19)
        tip_mac = Sheets(9).Cells(i, 4)
        vol_tot_mac = Sheets(9).Cells(i, 14)
        
x:
        i = 2
        
        While Sheets(9).Cells(i, 19) <> 0
            If Sheets(9).Cells(i, 19) < mom_cim_var And Sheets(9).Cells(i, 19) > mom_poste_var Then
                mom_cim_var = Sheets(9).Cells(i, 19)
                tip_mac = Sheets(9).Cells(i, 4)
                vol_tot_mac = Sheets(9).Cells(i, 14)
                GoTo x
            Else: i = i + 1
            End If
        Wend
        
'//
'//INSERCI�N CIMENTACI�N EN REPLANTEO
'//
        
        Sheets(1).Cells(z, 37) = mom_cim_var
        Sheets(1).Cells(z, 22) = tip_mac
        Sheets(1).Cells(z, 23) = vol_tot_mac
        End If
        z = z + 2
        i = 2
        
    Wend
        
  End Sub
Public Function MAX(x As Double, y As Double) As Double
Dim Resul As Double
If x > y Then
Resul = x
ElseIf x <= y Then
Resul = y
End If
MAX = Resul
End Function
Public Function MIN(x As Double, y As Double) As Double
Dim Resul As Double
If x > y Then
Resul = y
ElseIf x <= y Then
Resul = x
End If
MIN = Resul
End Function

 



