Imports System.Data.OleDb
Module nueva_lac
Sub introducir 
        Dim oConn As New OleDbConnection
        Dim oComm As OleDbCommand
        Dim oComm2 As OleDbCommand
        Dim oRead As OleDbDataReader
        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")
        'INTRODUCIR NUEVA CATENARIA FALTA VER QUE EL NOMBRE ESCRITO NO COINCIDA

        oComm = New OleDbCommand("insert into Datos(Nombre_Catenaria, Sistema, Alimentación, Altura_nominal, Altura_mínima, Altura_máxima, Altura_catenaria, Distancia_máx_entre_vanos, Distancia_máx_del_cantón, Vano_máximo, Vano_máx_en_sec_mecánico, Vano_máx_en_sec_eléctrico, Vano_máx_en_túnel, Incr_normalizado_de_vano, Incr_máx_altura_HC, Núm_mín_vanos_en_sec_mec, Núm_mín_vanos_en_sec_eléct, Ancho_vía, Descentramiento_máx_recta, Descentramiento_máx_curva, Radio_considerable_como_recta, Zona_trabajo_pantógrafo, Elevación_máx_pantógrafo, Velocidad_viento, Flecha_máx_centro_vano, Distancia_carril_poste, Distancia_base_poste_PMR, Distancia_eléct_sec_mecánico, Distancia_eléct_sec_eléctrico, Long_zona_común_máx, Long_zona_común_mín, Long_Zona_Neutra, Hilo_de_Contacto, Sustentador, C_de_Protección_Aérea, Cable_de_Tierra, Feeder_positivo, Feeder_negativo, Punto_fijo, Péndola, Anclaje, Posición_Feeder_positivo, Posición_Feeder_negativo, Núm_HC, Núm_CdPA, Núm_Feeder_positivo, Núm_Feeder_negativo, Tensión_HC, Tensión_sustentador, Tensión_CdPA, Tensión_Feeder_positivo, Tensión_Feeder_negativo, Tensión_punto_fijo, Adm_Línea, Tipo, Numeración, Adm_Línea_macizo, Tipo_macizo, Tubo_de_ménsula, Tubo_tirante, Cola_de_anclaje, Aislador_Feeder_positivo, Aislador_Feeder_negativo, Distancia_apoyo_y_1ª_péndola, Distancia_1ª_y_2ª_péndola, Distancia_máx_entre_péndolas) values(@Nombre_Catenaria, @Sistema, @Alimentación, @Altura_nominal, @Altura_mínima, @Altura_máxima, @Altura_catenaria, @Distancia_máx_entre_vanos, @Distancia_máx_del_cantón, @Vano_máximo, @Vano_máx_en_sec_mecánico, @Vano_máx_en_sec_eléctrico, @Vano_máx_en_túnel, Incr_normalizado_de_vano, @Incr_máx_altura_HC, @Núm_mín_vanos_en_sec_mec, @Núm_mín_vanos_en_sec_eléct, @Ancho_vía, @Descentramiento_máx_recta, @Descentramiento_máx_curva, @Radio_considerable_como_recta, @Zona_trabajo_pantógrafo, @Elevación_máx_pantógrafo, @Velocidad_viento, @Flecha_máx_centro_vano, @Distancia_carril_poste, @Distancia_base_poste_PMR, @Distancia_eléct_sec_mecánico, @Distancia_eléct_sec_eléctrico, @Long_zona_común_máx, @Long_zona_común_mín, @Long_Zona_Neutra, @Hilo_de_Contacto, @Sustentador, @C_de_Protección_Aérea, @Cable_de_Tierra, @Feeder_positivo, Feeder_negativo, @Punto_fijo, @Péndola, @Anclaje, @Posición_Feeder_positivo, @Posición_Feeder_negativo, @Núm_HC, @Núm_CdPA, @Núm_Feeder_positivo, @Núm_Feeder_negativo, @Tensión_HC, @Tensión_sustentador, @Tensión_CdPA, @Tensión_Feeder_positivo, @Tensión_Feeder_negativo, @Tensión_punto_fijo, @Adm_Línea, @Tipo, @Numeración, @Adm_Línea_macizo, @Tipo_macizo, @Tubo_de_ménsula, @Tubo_tirante, @Cola_de_anclaje, @Aislador_Feeder_positivo, @Aislador_Feeder_negativo, @Distancia_apoyo_y_1ª_péndola, @Distancia_1ª_y_2ª_péndola, @Distancia_máx_entre_péndolas)", oConn)
        oComm2 = New OleDbCommand("select * from Datos", oConn)
        oConn.Open()

        oRead = oComm2.ExecuteReader

        If Pantalla_datos.TextNombrecatenaria.Visible Then

            While oRead.Read
                If (oRead("Nombre_Catenaria") = Pantalla_datos.TextNombrecatenaria.Text) Then
                    Pantalla_datos.TextNombrecatenaria.BackColor = Color.Red
                    Pantalla_datos.Label2.ForeColor = Color.Red
                    Pantalla_datos.TextNombrecatenaria.SelectAll()
                    MsgBox("NOMBRE REPETIDO", 48)
                    GoTo x
                End If

            End While

            oComm.Parameters.Add(New OleDbParameter("@Nombre_Catenaria", OleDbType.VarChar))
            oComm.Parameters("@Nombre_Catenaria").Value = Pantalla_datos.TextNombrecatenaria.Text

        Else

            oComm.Parameters.Add(New OleDbParameter("@Nombre_Catenaria", OleDbType.VarChar))
            oComm.Parameters("@Nombre_Catenaria").Value = Pantalla_principal.nueva_lac

        End If


        oComm.Parameters.Add(New OleDbParameter("@Sistema", OleDbType.VarChar))
        oComm.Parameters("@Sistema").Value = Pantalla_datos.ComboSistema.Text

        oComm.Parameters.Add(New OleDbParameter("@Alimentación", OleDbType.VarChar))
        oComm.Parameters("@Alimentación").Value = Pantalla_datos.TextAlimentacion.Text

        oComm.Parameters.Add(New OleDbParameter("@Altura_nominal", OleDbType.VarChar))
        oComm.Parameters("@Altura_nominal").Value = Pantalla_datos.TextAlturanominal.Text

        oComm.Parameters.Add(New OleDbParameter("@Altura_mínima", OleDbType.VarChar))
        oComm.Parameters("@Altura_mínima").Value = Pantalla_datos.TextAlturaminima.Text

        oComm.Parameters.Add(New OleDbParameter("@Altura_máxima", OleDbType.VarChar))
        oComm.Parameters("@Altura_máxima").Value = Pantalla_datos.TextAlturamaxima.Text

        oComm.Parameters.Add(New OleDbParameter("@Altura_catenaria", OleDbType.VarChar))
        oComm.Parameters("@Altura_catenaria").Value = Pantalla_datos.TextAlturacatenaria.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_máx_entre_vanos", OleDbType.VarChar))
        oComm.Parameters("@Distancia_máx_entre_vanos").Value = Pantalla_datos.TextDistanciamaxentrevanos.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_máx_del_cantón", OleDbType.VarChar))
        oComm.Parameters("@Distancia_máx_del_cantón").Value = Pantalla_datos.TextDistanciamaxdelcanton.Text

        oComm.Parameters.Add(New OleDbParameter("@Vano_máximo", OleDbType.VarChar))
        oComm.Parameters("@Vano_máximo").Value = Pantalla_datos.TextVanomaximo.Text

        oComm.Parameters.Add(New OleDbParameter("@Vano_máx_en_sec_mecánico", OleDbType.VarChar))
        oComm.Parameters("@Vano_máx_en_sec_mecánico").Value = Pantalla_datos.TextVanomaxensecmecanico.Text

        oComm.Parameters.Add(New OleDbParameter("@Vano_máx_en_sec_eléctrico", OleDbType.VarChar))
        oComm.Parameters("@Vano_máx_en_sec_eléctrico").Value = Pantalla_datos.TextVanomaxensecelectrico.Text

        oComm.Parameters.Add(New OleDbParameter("@Vano_máx_en_túnel", OleDbType.VarChar))
        oComm.Parameters("@Vano_máx_en_túnel").Value = Pantalla_datos.TextVanomaxentunel.Text

        oComm.Parameters.Add(New OleDbParameter("@Incr_normalizado_de_vano", OleDbType.VarChar))
        oComm.Parameters("@Incr_normalizado_de_vano").Value = Pantalla_datos.TextIncrnormalizadodevano.Text

        oComm.Parameters.Add(New OleDbParameter("@Incr_máx_altura_HC", OleDbType.VarChar))
        oComm.Parameters("@Incr_máx_altura_HC").Value = Pantalla_datos.TextIncrmaxalturahc.Text

        oComm.Parameters.Add(New OleDbParameter("@Núm_mín_vanos_en_sec_mec", OleDbType.VarChar))
        oComm.Parameters("@Núm_mín_vanos_en_sec_mec").Value = Pantalla_datos.TextNumminvanosensecmec.Text

        oComm.Parameters.Add(New OleDbParameter("@Núm_mín_vanos_en_sec_eléct", OleDbType.VarChar))
        oComm.Parameters("@Núm_mín_vanos_en_sec_eléct").Value = Pantalla_datos.TextNumminvanosensecelect.Text

        oComm.Parameters.Add(New OleDbParameter("@Ancho_vía", OleDbType.VarChar))
        oComm.Parameters("@Ancho_vía").Value = Pantalla_datos.TextAnchovia.Text

        oComm.Parameters.Add(New OleDbParameter("@Descentramiento_máx_recta", OleDbType.VarChar))
        oComm.Parameters("@Descentramiento_máx_recta").Value = Pantalla_datos.TextDescentramientomaxrecta.Text

        oComm.Parameters.Add(New OleDbParameter("@Descentramiento_máx_curva", OleDbType.VarChar))
        oComm.Parameters("@Descentramiento_máx_curva").Value = Pantalla_datos.TextDescentramientomaxcurva.Text

        oComm.Parameters.Add(New OleDbParameter("@Radio_considerable_como_recta", OleDbType.VarChar))
        oComm.Parameters("@Radio_considerable_como_recta").Value = Pantalla_datos.TextRadioconsiderablecomorecta.Text

        oComm.Parameters.Add(New OleDbParameter("@Zona_trabajo_pantógrafo", OleDbType.VarChar))
        oComm.Parameters("@Zona_trabajo_pantógrafo").Value = Pantalla_datos.TextZonatrabajopantografo.Text

        oComm.Parameters.Add(New OleDbParameter("@Elevación_máx_pantógrafo", OleDbType.VarChar))
        oComm.Parameters("@Elevación_máx_pantógrafo").Value = Pantalla_datos.TextElevacionmaxpantografo.Text

        oComm.Parameters.Add(New OleDbParameter("@Velocidad_viento", OleDbType.VarChar))
        oComm.Parameters("@Velocidad_viento").Value = Pantalla_datos.TextVelocidadviento.Text

        oComm.Parameters.Add(New OleDbParameter("@Flecha_máx_centro_vano", OleDbType.VarChar))
        oComm.Parameters("@Flecha_máx_centro_vano").Value = Pantalla_datos.TextFlechamaxcentrovano.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_carril_poste", OleDbType.VarChar))
        oComm.Parameters("@Distancia_carril_poste").Value = Pantalla_datos.TextDistanciacarrilposte.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_base_poste_PMR", OleDbType.VarChar))
        oComm.Parameters("@Distancia_base_poste_PMR").Value = Pantalla_datos.TextDistanciabasepostepmr.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_eléct_sec_mecánico", OleDbType.VarChar))
        oComm.Parameters("@Distancia_eléct_sec_mecánico").Value = Pantalla_datos.TextDistanciaelectsecmecanico.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_eléct_sec_eléctrico", OleDbType.VarChar))
        oComm.Parameters("@Distancia_eléct_sec_eléctrico").Value = Pantalla_datos.TextDistanciaelectsecelectrico.Text

        oComm.Parameters.Add(New OleDbParameter("@Long_zona_común_máx", OleDbType.VarChar))
        oComm.Parameters("@Long_zona_común_máx").Value = Pantalla_datos.TextLongzonacomunmax.Text

        oComm.Parameters.Add(New OleDbParameter("@Long_zona_común_mín", OleDbType.VarChar))
        oComm.Parameters("@Long_zona_común_mín").Value = Pantalla_datos.TextLongzonacomunmin.Text

        oComm.Parameters.Add(New OleDbParameter("@Long_Zona_Neutra", OleDbType.VarChar))
        oComm.Parameters("@Long_Zona_Neutra").Value = Pantalla_datos.TextLongzonaneutra.Text

        oComm.Parameters.Add(New OleDbParameter("@Hilo_de_contacto", OleDbType.VarChar))
        oComm.Parameters("@Hilo_de_contacto").Value = Pantalla_datos.ComboHilodecontacto.Text

        oComm.Parameters.Add(New OleDbParameter("@Sustentador", OleDbType.VarChar))
        oComm.Parameters("@Sustentador").Value = Pantalla_datos.ComboSustentador.Text

        oComm.Parameters.Add(New OleDbParameter("@C_de_Protección_Aérea", OleDbType.VarChar))
        oComm.Parameters("@C_de_Protección_Aérea").Value = Pantalla_datos.ComboCdeproteccionaerea.Text

        oComm.Parameters.Add(New OleDbParameter("@Cable_de_Tierra", OleDbType.VarChar))
        oComm.Parameters("@Cable_de_Tierra").Value = Pantalla_datos.ComboCabledetierra.Text

        oComm.Parameters.Add(New OleDbParameter("@Feeder_positivo", OleDbType.VarChar))
        oComm.Parameters("@Feeder_positivo").Value = Pantalla_datos.ComboFeederpositivo.Text

        oComm.Parameters.Add(New OleDbParameter("@Feeder_negativo", OleDbType.VarChar))
        oComm.Parameters("@Feeder_negativo").Value = Pantalla_datos.ComboFeedernegativo.Text

        oComm.Parameters.Add(New OleDbParameter("@Punto_fijo", OleDbType.VarChar))
        oComm.Parameters("@Punto_fijo").Value = Pantalla_datos.ComboPuntofijo.Text

        oComm.Parameters.Add(New OleDbParameter("@Péndola", OleDbType.VarChar))
        oComm.Parameters("@Péndola").Value = Pantalla_datos.ComboPendola.Text

        oComm.Parameters.Add(New OleDbParameter("@Anclaje", OleDbType.VarChar))
        oComm.Parameters("@Anclaje").Value = Pantalla_datos.ComboAnclaje.Text

        oComm.Parameters.Add(New OleDbParameter("@Posición_Feeder_positivo", OleDbType.VarChar))
        oComm.Parameters("@Posición_Feeder_positivo").Value = Pantalla_datos.ComboPosicionfeederpositivo.Text

        oComm.Parameters.Add(New OleDbParameter("@Posición_Feeder_negativo", OleDbType.VarChar))
        oComm.Parameters("@Posición_Feeder_negativo").Value = Pantalla_datos.ComboPosicionfeedernegativo.Text

        oComm.Parameters.Add(New OleDbParameter("@Núm_HC", OleDbType.VarChar))
        oComm.Parameters("@Núm_HC").Value = Pantalla_datos.TextNumhc.Text

        oComm.Parameters.Add(New OleDbParameter("@Núm_CdPA", OleDbType.VarChar))
        oComm.Parameters("@Núm_CdPA").Value = Pantalla_datos.TextNumcdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@Núm_Feeder_positivo", OleDbType.VarChar))
        oComm.Parameters("@Núm_Feeder_positivo").Value = Pantalla_datos.TextNumfeederpositivo.Text

        oComm.Parameters.Add(New OleDbParameter("@Núm_Feeder_negativo", OleDbType.VarChar))
        oComm.Parameters("@Núm_Feeder_negativo").Value = Pantalla_datos.TextNumfeedernegativo.Text

        oComm.Parameters.Add(New OleDbParameter("@Tensión_HC", OleDbType.VarChar))
        oComm.Parameters("@Tensión_HC").Value = Pantalla_datos.TextTensionhc.Text

        oComm.Parameters.Add(New OleDbParameter("@Tensión_sustentador", OleDbType.VarChar))
        oComm.Parameters("@Tensión_sustentador").Value = Pantalla_datos.TextTensionsustentador.Text

        oComm.Parameters.Add(New OleDbParameter("@Tensión_CdPA", OleDbType.VarChar))
        oComm.Parameters("@Tensión_CdPA").Value = Pantalla_datos.TextTensioncdpa.Text

        oComm.Parameters.Add(New OleDbParameter("@Tensión_Feeder_positivo", OleDbType.VarChar))
        oComm.Parameters("@Tensión_Feeder_positivo").Value = Pantalla_datos.TextTensionfeederpositivo.Text

        oComm.Parameters.Add(New OleDbParameter("@Tensión_Feeder_negativo", OleDbType.VarChar))
        oComm.Parameters("@Tensión_Feeder_negativo").Value = Pantalla_datos.TextTensionfeedernegativo.Text

        oComm.Parameters.Add(New OleDbParameter("@Tensión_punto_fijo", OleDbType.VarChar))
        oComm.Parameters("@Tensión_punto_fijo").Value = Pantalla_datos.TextTensionpuntofijo.Text

        oComm.Parameters.Add(New OleDbParameter("@Adm_Línea", OleDbType.VarChar))
        oComm.Parameters("@Adm_Línea").Value = Pantalla_datos.ComboAdmlinea.Text

        oComm.Parameters.Add(New OleDbParameter("@Tipo", OleDbType.VarChar))
        oComm.Parameters("@Tipo").Value = Pantalla_datos.TextTipo.Text

        oComm.Parameters.Add(New OleDbParameter("@Numeración", OleDbType.VarChar))
        oComm.Parameters("@Numeración").Value = Pantalla_datos.ComboNumeración.Text

        oComm.Parameters.Add(New OleDbParameter("@Adm_Línea_macizo", OleDbType.VarChar))
        oComm.Parameters("@Adm_Línea_macizo").Value = Pantalla_datos.ComboAdmlineamacizo.Text

        oComm.Parameters.Add(New OleDbParameter("@Tipo_macizo", OleDbType.VarChar))
        oComm.Parameters("@Tipo_macizo").Value = Pantalla_datos.TextTipomacizo.Text

        oComm.Parameters.Add(New OleDbParameter("@Tubo_de_ménsula", OleDbType.VarChar))
        oComm.Parameters("@Tubo_de_ménsula").Value = Pantalla_datos.ComboTubodemensula.Text

        oComm.Parameters.Add(New OleDbParameter("@Tubo_tirante", OleDbType.VarChar))
        oComm.Parameters("@Tubo_tirante").Value = Pantalla_datos.ComboTubotirante.Text

        oComm.Parameters.Add(New OleDbParameter("@Cola_de_anclaje", OleDbType.VarChar))
        oComm.Parameters("@Cola_de_anclaje").Value = Pantalla_datos.ComboColadeanclaje.Text

        oComm.Parameters.Add(New OleDbParameter("@Aislador_Feeder_positivo", OleDbType.VarChar))
        oComm.Parameters("@Aislador_Feeder_positivo").Value = Pantalla_datos.ComboAisladorfeederpositivo.Text

        oComm.Parameters.Add(New OleDbParameter("@Aislador_Feeder_negativo", OleDbType.VarChar))
        oComm.Parameters("@Aislador_Feeder_negativo").Value = Pantalla_datos.ComboAisladorfeedernegativo.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_apoyo_y_1ª_péndola", OleDbType.VarChar))
        oComm.Parameters("@Distancia_apoyo_y_1ª_péndola").Value = Pantalla_datos.TextDistanciaapoyoyprimerapendola.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_1ª_y_2ª_péndola", OleDbType.VarChar))
        oComm.Parameters("@Distancia_1ª_y_2ª_péndola").Value = Pantalla_datos.TextDistanciaprimeraysegundapendola.Text

        oComm.Parameters.Add(New OleDbParameter("@Distancia_máx_entre_péndolas", OleDbType.VarChar))
        oComm.Parameters("@Distancia_máx_entre_péndolas").Value = Pantalla_datos.TextDistanciamaxentrependolas.Text

        oComm2.Connection.Close()
        oComm.Connection.Open()
        On Error GoTo mserror
        oComm.ExecuteNonQuery()
        oComm.Connection.Close()
        MsgBox("RESGITRO AÑADIDO")
        Pantalla_datos.Close()
        Pantalla_principal.Label1.Hide()
        Pantalla_principal.Label2.Hide()
        Pantalla_principal.TextBox1.Hide()
        Pantalla_principal.ComboBox1.Hide()
        Pantalla_principal.Button1.Hide()
        Pantalla_principal.RadioButton1.Hide()
        Pantalla_principal.RadioButton2.Hide()
        Pantalla_principal.GroupBox1.Text = "Datos de catenaria introducidos"
        Pantalla_principal.GroupBox2.ForeColor = Color.Green

        Exit Sub
mserror:
        oComm.Connection.Close()
        MsgBox("FALTAN CAMPOS POR COMPLETAR")
x:

    End Sub

End Module
