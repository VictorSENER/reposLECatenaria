Imports System.Data.OleDb

Public Class Pantalla_datos

    Dim oConn As New OleDbConnection

    Private Sub Pantalla_datos_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Anclaje' Puede moverla o quitarla según sea necesario.
        Me.Conductor_AnclajeTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Anclaje)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Pendola' Puede moverla o quitarla según sea necesario.
        Me.Conductor_PendolaTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Pendola)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_punto_fijo' Puede moverla o quitarla según sea necesario.
        Me.Conductor_punto_fijoTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_punto_fijo)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet._Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Feeder__TableAdapter.Fill(Me.Base_de_datosDataSet._Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.__Conductor_Feeder__' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Feeder__TableAdapter1.Fill(Me.Base_de_datosDataSet.__Conductor_Feeder__)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Cable_de_Tierra' Puede moverla o quitarla según sea necesario.
        Me.Conductor_Cable_de_TierraTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Cable_de_Tierra)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_CDPA' Puede moverla o quitarla según sea necesario.
        Me.Conductor_CDPATableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_CDPA)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_Sustentador' Puede moverla o quitarla según sea necesario.
        Me.Conductor_SustentadorTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_Sustentador)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Conductor_HC' Puede moverla o quitarla según sea necesario.
        Me.Conductor_HCTableAdapter.Fill(Me.Base_de_datosDataSet.Conductor_HC)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Postes_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Postes_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Postes_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Macizos_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Macizos_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Macizos_Consulta)
        'TODO: esta línea de código carga datos en la tabla 'Base_de_datosDataSet.Electrificación_Consulta' Puede moverla o quitarla según sea necesario.
        Me.Electrificación_ConsultaTableAdapter.Fill(Me.Base_de_datosDataSet.Electrificación_Consulta)

        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\29289\Escritorio\SIRECA\reposLECatenaria\Nueva carpeta\SiReCa\SiReCa\Base de datos.accdb")

    End Sub

    Private Sub Pantalla_datos_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        oConn = Nothing

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click


        Dim oComm As OleDbCommand
        Dim oRead As OleDbDataReader

        If RadioCatenariaCargar.Checked = True Then

            'LEER NOMBRE CATENARIA Y CARGAR

            oConn.Open()
            oComm = New OleDbCommand("select * from Datos", oConn)
            oRead = oComm.ExecuteReader

            Dim Ncat As String

            While oRead.Read

                'El DataReader se situa sobre el registro

                If (TextNombrecatenaria.Text = oRead("Nombre_Catenaria")) Then

                    ComboSistema.Text = oRead("sistema") 'Lee los campos que se requieran ya situado sobre el registro correspondiente
                    TextAlimentacion.Text = oRead("Alimentación")
                    TextAlturanominal.Text = oRead("Altura_nominal")
                    TextAlturaminima.Text = oRead("Altura_mínima")
                    TextAlturamaxima.Text = oRead("Altura_máxima")
                    TextAlturacatenaria.Text = oRead("Altura_catenaria")
                    TextDistanciamaxentrevanos.Text = oRead("Distancia_máx_entre_vanos")
                    TextDistanciamaxdelcanton.Text = oRead("Distancia_máx_del_cantón")
                    TextVanomaximo.Text = oRead("Vano_máximo")
                    TextVanomaxensecmecanico.Text = oRead("Vano_máx_en_sec_mecánico")
                    TextVanomaxensecelectrico.Text = oRead("Vano_máx_en_sec_eléctrico")
                    TextVanomaxentunel.Text = oRead("Vano_máx_en_túnel")
                    TextIncrnormalizadodevano.Text = oRead("Incr_normalizado_de_vano")
                    TextIncrmaxalturahc.Text = oRead("Incr_máx_altura_hc")
                    TextNumminvanosensecmec.Text = oRead("Núm_mín_vanos_en_sec_mec")
                    TextNumminvanosensecelect.Text = oRead("Núm_mín_vanos_en_sec_eléct")
                    TextAnchovia.Text = oRead("Ancho_vía")
                    TextDescentramientomaxrecta.Text = oRead("Descentramiento_máx_recta")
                    TextDescentramientomaxcurva.Text = oRead("Descentramiento_máx_curva")
                    TextRadioconsiderablecomorecta.Text = oRead("Radio_considerable_como_recta")
                    TextZonatrabajopantografo.Text = oRead("Zona_trabajo_pantógrafo")
                    TextElevacionmaxpantografo.Text = oRead("Elevación_máx_pantógrafo")
                    TextVelocidadviento.Text = oRead("Velocidad_viento")
                    TextFlechamaxcentrovano.Text = oRead("Flecha_máx_centro_vano")
                    TextDistanciacarrilposte.Text = oRead("Distancia_carril_poste")
                    TextDistanciabasepostepmr.Text = oRead("Distancia_base_poste_pmr")
                    TextDistanciaelectsecmecanico.Text = oRead("Distancia_eléct_sec_mecánico")
                    TextDistanciaelectsecelectrico.Text = oRead("Distancia_eléct_sec_eléctrico")
                    TextLongzonacomunmax.Text = oRead("Long_zona_común_máx")
                    TextLongzonacomunmin.Text = oRead("Long_zona_común_mín")
                    TextLongzonaneutra.Text = oRead("Long_zona_neutra")
                    ComboHilodecontacto.Text = oRead("Hilo_de_contacto")
                    ComboSustentador.Text = oRead("Sustentador")
                    ComboCdeproteccionaerea.Text = oRead("C_de_protección_aérea")
                    ComboCabledetierra.Text = oRead("Cable_de_tierra")
                    ComboFeederpositivo.Text = oRead("Feeder_positivo")
                    ComboFeedernegativo.Text = oRead("Feeder_negativo")
                    ComboPuntofijo.Text = oRead("Punto_fijo")
                    ComboPendola.Text = oRead("Péndola")
                    ComboAnclaje.Text = oRead("Anclaje")
                    ComboPosicionfeederpositivo.Text = oRead("Posición_feeder_positivo")
                    ComboPosicionfeedernegativo.Text = oRead("Posición_feeder_negativo")
                    TextNumhc.Text = oRead("Núm_hc")
                    TextNumcdpa.Text = oRead("Núm_cdpa")
                    TextNumfeederpositivo.Text = oRead("Núm_feeder_positivo")
                    TextNumfeedernegativo.Text = oRead("Núm_feeder_negativo")
                    TextTensionhc.Text = oRead("Tensión_hc")
                    TextTensionsustentador.Text = oRead("Tensión_sustentador")
                    TextTensioncdpa.Text = oRead("Tensión_cdpa")
                    TextTensionfeederpositivo.Text = oRead("Tensión_feeder_positivo")
                    TextTensionfeedernegativo.Text = oRead("Tensión_feeder_negativo")
                    TextTensionpuntofijo.Text = oRead("Tensión_punto_fijo")
                    ComboAdmlinea.Text = oRead("Adm_línea")
                    TextTipo.Text = oRead("Tipo")
                    ComboNumeración.Text = oRead("Numeración")
                    ComboAdmlineamacizo.Text = oRead("Adm_Línea_macizo")
                    TextTipomacizo.Text = oRead("Tipo_macizo")
                    ComboTubodemensula.Text = oRead("Tubo_de_ménsula")
                    ComboTubotirante.Text = oRead("Tubo_tirante")
                    ComboColadeanclaje.Text = oRead("Cola_de_anclaje")
                    ComboFeederpositivo.Text = oRead("Feeder_positivo")
                    ComboFeedernegativo.Text = oRead("Feeder_negativo")
                    TextDistanciaapoyoyprimerapendola.Text = oRead("Distancia_apoyo_y_1ª_péndola")
                    TextDistanciaprimeraysegundapendola.Text = oRead("Distancia_1ª_y_2ª_péndola")
                    TextDistanciamaxentrependolas.Text = oRead("Distancia_máx_entre_péndolas")



                End If

            End While

            oRead.Close()
            oConn.Close()

        ElseIf RadioCatenarianueva.Checked = True Then

            'INTRODUCIR NUEVA CATENARIA FALTA VER QUE EL NOMBRE ESCRITO NO COINCIDA

            oComm = New OleDbCommand("insert into Datos(Nombre_Catenaria, Sistema, Alimentación, Altura_nominal, Altura_mínima, Altura_máxima, Altura_catenaria, Distancia_máx_entre_vanos, Distancia_máx_del_cantón, Vano_máximo, Vano_máx_en_sec_mecánico, Vano_máx_en_sec_eléctrico, Vano_máx_en_túnel, Incr_normalizado_de_vano, Incr_máx_altura_HC, Núm_mín_vanos_en_sec_mec, Núm_mín_vanos_en_sec_eléct, Ancho_vía, Descentramiento_máx_recta, Descentramiento_máx_curva, Radio_considerable_como_recta, Zona_trabajo_pantógrafo, Elevación_máx_pantógrafo, Velocidad_viento, Flecha_máx_centro_vano, Distancia_carril_poste, Distancia_base_poste_PMR, Distancia_eléct_sec_mecánico, Distancia_eléct_sec_eléctrico, Long_zona_común_máx, Long_zona_común_mín, Long_Zona_Neutra, Hilo_de_Contacto, Sustentador, C_de_Protección_Aérea, Cable_de_Tierra, Feeder_positivo, Feeder_negativo, Punto_fijo, Péndola, Anclaje, Posición_Feeder_positivo, Posición_Feeder_negativo, Núm_HC, Núm_CdPA, Núm_Feeder_positivo, Núm_Feeder_negativo, Tensión_HC, Tensión_sustentador, Tensión_CdPA, Tensión_Feeder_positivo, Tensión_Feeder_negativo, Tensión_punto_fijo, Adm_Línea, Tipo, Numeración, Adm_Línea_macizo, Tipo_macizo, Tubo_de_ménsula, Tubo_tirante, Cola_de_anclaje, Aislador_Feeder_positivo, Aislador_Feeder_negativo, Distancia_apoyo_y_1ª_péndola, Distancia_1ª_y_2ª_péndola, Distancia_máx_entre_péndolas) values(@Nombre_Catenaria, @Sistema, @Alimentación, @Altura_nominal, @Altura_mínima, @Altura_máxima, @Altura_catenaria, @Distancia_máx_entre_vanos, @Distancia_máx_del_cantón, @Vano_máximo, @Vano_máx_en_sec_mecánico, @Vano_máx_en_sec_eléctrico, @Vano_máx_en_túnel, Incr_normalizado_de_vano, @Incr_máx_altura_HC, @Núm_mín_vanos_en_sec_mec, @Núm_mín_vanos_en_sec_eléct, @Ancho_vía, @Descentramiento_máx_recta, @Descentramiento_máx_curva, @Radio_considerable_como_recta, @Zona_trabajo_pantógrafo, @Elevación_máx_pantógrafo, @Velocidad_viento, @Flecha_máx_centro_vano, @Distancia_carril_poste, @Distancia_base_poste_PMR, @Distancia_eléct_sec_mecánico, @Distancia_eléct_sec_eléctrico, @Long_zona_común_máx, @Long_zona_común_mín, @Long_Zona_Neutra, @Hilo_de_Contacto, @Sustentador, @C_de_Protección_Aérea, @Cable_de_Tierra, @Feeder_positivo, Feeder_negativo, @Punto_fijo, @Péndola, @Anclaje, @Posición_Feeder_positivo, @Posición_Feeder_negativo, @Núm_HC, @Núm_CdPA, @Núm_Feeder_positivo, @Núm_Feeder_negativo, @Tensión_HC, @Tensión_sustentador, @Tensión_CdPA, @Tensión_Feeder_positivo, @Tensión_Feeder_negativo, @Tensión_punto_fijo, @Adm_Línea, @Tipo, @Numeración, @Adm_Línea_macizo, @Tipo_macizo, @Tubo_de_ménsula, @Tubo_tirante, @Cola_de_anclaje, @Aislador_Feeder_positivo, @Aislador_Feeder_negativo, @Distancia_apoyo_y_1ª_péndola, @Distancia_1ª_y_2ª_péndola, @Distancia_máx_entre_péndolas)", oConn)

            oComm.Parameters.Add(New OleDbParameter("@Nombre_Catenaria", OleDbType.VarChar))
            oComm.Parameters("@Nombre_Catenaria").Value = ComboSistema.Text

            oComm.Parameters.Add(New OleDbParameter("@Sistema", OleDbType.VarChar))
            oComm.Parameters("@Sistema").Value = ComboSistema.Text

            oComm.Parameters.Add(New OleDbParameter("@Alimentación", OleDbType.VarChar))
            oComm.Parameters("@Alimentación").Value = TextAlimentacion.Text

            oComm.Parameters.Add(New OleDbParameter("@Altura_nominal", OleDbType.VarChar))
            oComm.Parameters("@Altura_nominal").Value = TextAlturanominal.Text

            oComm.Parameters.Add(New OleDbParameter("@Altura_mínima", OleDbType.VarChar))
            oComm.Parameters("@Altura_mínima").Value = TextAlturaminima.Text

            oComm.Parameters.Add(New OleDbParameter("@Altura_máxima", OleDbType.VarChar))
            oComm.Parameters("@Altura_máxima").Value = TextAlturamaxima.Text

            oComm.Parameters.Add(New OleDbParameter("@Altura_catenaria", OleDbType.VarChar))
            oComm.Parameters("@Altura_catenaria").Value = TextAlturacatenaria.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_máx_entre_vanos", OleDbType.VarChar))
            oComm.Parameters("@Distancia_máx_entre_vanos").Value = TextDistanciamaxentrevanos.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_máx_del_cantón", OleDbType.VarChar))
            oComm.Parameters("@Distancia_máx_del_cantón").Value = TextDistanciamaxdelcanton.Text

            oComm.Parameters.Add(New OleDbParameter("@Vano_máximo", OleDbType.VarChar))
            oComm.Parameters("@Vano_máximo").Value = TextVanomaximo.Text

            oComm.Parameters.Add(New OleDbParameter("@Vano_máx_en_sec_mecánico", OleDbType.VarChar))
            oComm.Parameters("@Vano_máx_en_sec_mecánico").Value = TextVanomaxensecmecanico.Text

            oComm.Parameters.Add(New OleDbParameter("@Vano_máx_en_sec_eléctrico", OleDbType.VarChar))
            oComm.Parameters("@Vano_máx_en_sec_eléctrico").Value = TextVanomaxensecelectrico.Text

            oComm.Parameters.Add(New OleDbParameter("@Vano_máx_en_túnel", OleDbType.VarChar))
            oComm.Parameters("@Vano_máx_en_túnel").Value = TextVanomaxentunel.Text

            oComm.Parameters.Add(New OleDbParameter("@Incr_normalizado_de_vano", OleDbType.VarChar))
            oComm.Parameters("@Incr_normalizado_de_vano").Value = TextIncrnormalizadodevano.Text

            oComm.Parameters.Add(New OleDbParameter("@Incr_máx_altura_HC", OleDbType.VarChar))
            oComm.Parameters("@Incr_máx_altura_HC").Value = TextIncrmaxalturahc.Text

            oComm.Parameters.Add(New OleDbParameter("@Núm_mín_vanos_en_sec_mec", OleDbType.VarChar))
            oComm.Parameters("@Núm_mín_vanos_en_sec_mec").Value = TextNumminvanosensecmec.Text

            oComm.Parameters.Add(New OleDbParameter("@Núm_mín_vanos_en_sec_eléct", OleDbType.VarChar))
            oComm.Parameters("@Núm_mín_vanos_en_sec_eléct").Value = TextNumminvanosensecelect.Text

            oComm.Parameters.Add(New OleDbParameter("@Ancho_vía", OleDbType.VarChar))
            oComm.Parameters("@Ancho_vía").Value = TextAnchovia.Text

            oComm.Parameters.Add(New OleDbParameter("@Descentramiento_máx_recta", OleDbType.VarChar))
            oComm.Parameters("@Descentramiento_máx_recta").Value = TextDescentramientomaxrecta.Text

            oComm.Parameters.Add(New OleDbParameter("@Descentramiento_máx_curva", OleDbType.VarChar))
            oComm.Parameters("@Descentramiento_máx_curva").Value = TextDescentramientomaxcurva.Text

            oComm.Parameters.Add(New OleDbParameter("@Radio_considerable_como_recta", OleDbType.VarChar))
            oComm.Parameters("@Radio_considerable_como_recta").Value = TextRadioconsiderablecomorecta.Text

            oComm.Parameters.Add(New OleDbParameter("@Zona_trabajo_pantógrafo", OleDbType.VarChar))
            oComm.Parameters("@Zona_trabajo_pantógrafo").Value = TextZonatrabajopantografo.Text

            oComm.Parameters.Add(New OleDbParameter("@Elevación_máx_pantógrafo", OleDbType.VarChar))
            oComm.Parameters("@Elevación_máx_pantógrafo").Value = TextElevacionmaxpantografo.Text

            oComm.Parameters.Add(New OleDbParameter("@Velocidad_viento", OleDbType.VarChar))
            oComm.Parameters("@Velocidad_viento").Value = TextVelocidadviento.Text

            oComm.Parameters.Add(New OleDbParameter("@Flecha_máx_centro_vano", OleDbType.VarChar))
            oComm.Parameters("@Flecha_máx_centro_vano").Value = TextFlechamaxcentrovano.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_carril_poste", OleDbType.VarChar))
            oComm.Parameters("@Distancia_carril_poste").Value = TextDistanciacarrilposte.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_base_poste_PMR", OleDbType.VarChar))
            oComm.Parameters("@Distancia_base_poste_PMR").Value = TextDistanciabasepostepmr.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_eléct_sec_mecánico", OleDbType.VarChar))
            oComm.Parameters("@Distancia_eléct_sec_mecánico").Value = TextDistanciaelectsecmecanico.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_eléct_sec_eléctrico", OleDbType.VarChar))
            oComm.Parameters("@Distancia_eléct_sec_eléctrico").Value = TextDistanciaelectsecelectrico.Text

            oComm.Parameters.Add(New OleDbParameter("@Long_zona_común_máx", OleDbType.VarChar))
            oComm.Parameters("@Long_zona_común_máx").Value = TextLongzonacomunmax.Text

            oComm.Parameters.Add(New OleDbParameter("@Long_zona_común_mín", OleDbType.VarChar))
            oComm.Parameters("@Long_zona_común_mín").Value = TextLongzonacomunmin.Text

            oComm.Parameters.Add(New OleDbParameter("@Long_Zona_Neutra", OleDbType.VarChar))
            oComm.Parameters("@Long_Zona_Neutra").Value = TextLongzonaneutra.Text

            oComm.Parameters.Add(New OleDbParameter("@Hilo_de_contacto", OleDbType.VarChar))
            oComm.Parameters("@Hilo_de_contacto").Value = ComboHilodecontacto.Text

            oComm.Parameters.Add(New OleDbParameter("@Sustentador", OleDbType.VarChar))
            oComm.Parameters("@Sustentador").Value = ComboSustentador.Text

            oComm.Parameters.Add(New OleDbParameter("@C_de_Protección_Aérea", OleDbType.VarChar))
            oComm.Parameters("@C_de_Protección_Aérea").Value = ComboCdeproteccionaerea.Text

            oComm.Parameters.Add(New OleDbParameter("@Cable_de_Tierra", OleDbType.VarChar))
            oComm.Parameters("@Cable_de_Tierra").Value = ComboCabledetierra.Text

            oComm.Parameters.Add(New OleDbParameter("@Feeder_positivo", OleDbType.VarChar))
            oComm.Parameters("@Feeder_positivo").Value = ComboFeederpositivo.Text

            oComm.Parameters.Add(New OleDbParameter("@Feeder_negativo", OleDbType.VarChar))
            oComm.Parameters("@Feeder_negativo").Value = ComboFeedernegativo.Text

            oComm.Parameters.Add(New OleDbParameter("@Punto_fijo", OleDbType.VarChar))
            oComm.Parameters("@Punto_fijo").Value = ComboPuntofijo.Text

            oComm.Parameters.Add(New OleDbParameter("@Péndola", OleDbType.VarChar))
            oComm.Parameters("@Péndola").Value = ComboPendola.Text

            oComm.Parameters.Add(New OleDbParameter("@Anclaje", OleDbType.VarChar))
            oComm.Parameters("@Anclaje").Value = ComboAnclaje.Text

            oComm.Parameters.Add(New OleDbParameter("@Posición_Feeder_positivo", OleDbType.VarChar))
            oComm.Parameters("@Posición_Feeder_positivo").Value = ComboPosicionfeederpositivo.Text

            oComm.Parameters.Add(New OleDbParameter("@Posición_Feeder_negativo", OleDbType.VarChar))
            oComm.Parameters("@Posición_Feeder_negativo").Value = ComboPosicionfeedernegativo.Text

            oComm.Parameters.Add(New OleDbParameter("@Núm_HC", OleDbType.VarChar))
            oComm.Parameters("@Núm_HC").Value = TextNumhc.Text

            oComm.Parameters.Add(New OleDbParameter("@Núm_CdPA", OleDbType.VarChar))
            oComm.Parameters("@Núm_CdPA").Value = TextNumcdpa.Text

            oComm.Parameters.Add(New OleDbParameter("@Núm_Feeder_positivo", OleDbType.VarChar))
            oComm.Parameters("@Núm_Feeder_positivo").Value = TextNumfeederpositivo.Text

            oComm.Parameters.Add(New OleDbParameter("@Núm_Feeder_negativo", OleDbType.VarChar))
            oComm.Parameters("@Núm_Feeder_negativo").Value = TextNumfeedernegativo.Text

            oComm.Parameters.Add(New OleDbParameter("@Tensión_HC", OleDbType.VarChar))
            oComm.Parameters("@Tensión_HC").Value = TextTensionhc.Text

            oComm.Parameters.Add(New OleDbParameter("@Tensión_sustentador", OleDbType.VarChar))
            oComm.Parameters("@Tensión_sustentador").Value = TextTensionsustentador.Text

            oComm.Parameters.Add(New OleDbParameter("@Tensión_CdPA", OleDbType.VarChar))
            oComm.Parameters("@Tensión_CdPA").Value = TextTensioncdpa.Text

            oComm.Parameters.Add(New OleDbParameter("@Tensión_Feeder_positivo", OleDbType.VarChar))
            oComm.Parameters("@Tensión_Feeder_positivo").Value = TextTensionfeederpositivo.Text

            oComm.Parameters.Add(New OleDbParameter("@Tensión_Feeder_negativo", OleDbType.VarChar))
            oComm.Parameters("@Tensión_Feeder_negativo").Value = TextTensionfeedernegativo.Text

            oComm.Parameters.Add(New OleDbParameter("@Tensión_punto_fijo", OleDbType.VarChar))
            oComm.Parameters("@Tensión_punto_fijo").Value = TextTensionpuntofijo.Text

            oComm.Parameters.Add(New OleDbParameter("@Adm_Línea", OleDbType.VarChar))
            oComm.Parameters("@Adm_Línea").Value = ComboAdmlinea.Text

            oComm.Parameters.Add(New OleDbParameter("@Tipo", OleDbType.VarChar))
            oComm.Parameters("@Tipo").Value = TextTipo.Text

            oComm.Parameters.Add(New OleDbParameter("@Numeración", OleDbType.VarChar))
            oComm.Parameters("@Numeración").Value = ComboNumeración.Text

            oComm.Parameters.Add(New OleDbParameter("@Adm_Línea_macizo", OleDbType.VarChar))
            oComm.Parameters("@Adm_Línea_macizo").Value = ComboAdmlineamacizo.Text

            oComm.Parameters.Add(New OleDbParameter("@Tipo_macizo", OleDbType.VarChar))
            oComm.Parameters("@Tipo_macizo").Value = TextTipomacizo.Text

            oComm.Parameters.Add(New OleDbParameter("@Tubo_de_ménsula", OleDbType.VarChar))
            oComm.Parameters("@Tubo_de_ménsula").Value = ComboTubodemensula.Text

            oComm.Parameters.Add(New OleDbParameter("@Tubo_tirante", OleDbType.VarChar))
            oComm.Parameters("@Tubo_tirante").Value = ComboTubotirante.Text

            oComm.Parameters.Add(New OleDbParameter("@Cola_de_anclaje", OleDbType.VarChar))
            oComm.Parameters("@Cola_de_anclaje").Value = ComboColadeanclaje.Text

            oComm.Parameters.Add(New OleDbParameter("@Aislador_Feeder_positivo", OleDbType.VarChar))
            oComm.Parameters("@Aislador_Feeder_positivo").Value = ComboAisladorfeederpositivo.Text

            oComm.Parameters.Add(New OleDbParameter("@Aislador_Feeder_negativo", OleDbType.VarChar))
            oComm.Parameters("@Aislador_Feeder_negativo").Value = ComboAisladorfeedernegativo.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_apoyo_y_1ª_péndola", OleDbType.VarChar))
            oComm.Parameters("@Distancia_apoyo_y_1ª_péndola").Value = TextDistanciaapoyoyprimerapendola.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_1ª_y_2ª_péndola", OleDbType.VarChar))
            oComm.Parameters("@Distancia_1ª_y_2ª_péndola").Value = TextDistanciaprimeraysegundapendola.Text

            oComm.Parameters.Add(New OleDbParameter("@Distancia_máx_entre_péndolas", OleDbType.VarChar))
            oComm.Parameters("@Distancia_máx_entre_péndolas").Value = TextDistanciamaxentrependolas.Text

            oComm.Connection.Open()
            On Error GoTo mserror
            oComm.ExecuteNonQuery()
            oComm.Connection.Close()
            MsgBox("Resgistro añadido")

            Exit Sub

mserror:
            oComm.Connection.Close()
            MsgBox("Usuario repetido")
            'falta el textNombrecatenaria.Focus()
            'falta textnombrecatenaria.SelectAll()

        End If

    End Sub


End Class