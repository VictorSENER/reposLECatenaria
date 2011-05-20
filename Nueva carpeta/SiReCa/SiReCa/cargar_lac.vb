Imports System.Data.OleDb
Module cargar_lac
    Sub cargar_lac()
        Dim oConn As New OleDbConnection
        Dim oComm As OleDbCommand
        Dim oRead As OleDbDataReader
        'LEER NOMBRE CATENARIA Y CARGAR
        oConn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Documents and Settings\23370\Escritorio\Nueva carpeta\SiReCa\Base de datos.accdb")
        oConn.Open()
        oComm = New OleDbCommand("select * from Datos", oConn)
        oRead = oComm.ExecuteReader

        While oRead.Read

            'El DataReader se situa sobre el registro

            If (Pantalla_principal.ComboBox1.Text = oRead("Nombre_Catenaria")) Then

                Pantalla_datos.ComboSistema.Text = oRead("sistema") 'Lee los campos que se requieran ya situado sobre el registro correspondiente
                Pantalla_datos.TextAlimentacion.Text = oRead("Alimentación")
                Pantalla_datos.TextAlturanominal.Text = oRead("Altura_nominal")
                Pantalla_datos.TextAlturaminima.Text = oRead("Altura_mínima")
                Pantalla_datos.TextAlturamaxima.Text = oRead("Altura_máxima")
                Pantalla_datos.TextAlturacatenaria.Text = oRead("Altura_catenaria")
                Pantalla_datos.TextDistanciamaxentrevanos.Text = oRead("Distancia_máx_entre_vanos")
                Pantalla_datos.TextDistanciamaxdelcanton.Text = oRead("Distancia_máx_del_cantón")
                Pantalla_datos.TextVanomaximo.Text = oRead("Vano_máximo")
                Pantalla_datos.TextVanomaxensecmecanico.Text = oRead("Vano_máx_en_sec_mecánico")
                Pantalla_datos.TextVanomaxensecelectrico.Text = oRead("Vano_máx_en_sec_eléctrico")
                Pantalla_datos.TextVanomaxentunel.Text = oRead("Vano_máx_en_túnel")
                Pantalla_datos.TextIncrnormalizadodevano.Text = oRead("Incr_normalizado_de_vano")
                Pantalla_datos.TextIncrmaxalturahc.Text = oRead("Incr_máx_altura_hc")
                Pantalla_datos.TextNumminvanosensecmec.Text = oRead("Núm_mín_vanos_en_sec_mec")
                Pantalla_datos.TextNumminvanosensecelect.Text = oRead("Núm_mín_vanos_en_sec_eléct")
                Pantalla_datos.TextAnchovia.Text = oRead("Ancho_vía")
                Pantalla_datos.TextDescentramientomaxrecta.Text = oRead("Descentramiento_máx_recta")
                Pantalla_datos.TextDescentramientomaxcurva.Text = oRead("Descentramiento_máx_curva")
                Pantalla_datos.TextRadioconsiderablecomorecta.Text = oRead("Radio_considerable_como_recta")
                Pantalla_datos.TextZonatrabajopantografo.Text = oRead("Zona_trabajo_pantógrafo")
                Pantalla_datos.TextElevacionmaxpantografo.Text = oRead("Elevación_máx_pantógrafo")
                Pantalla_datos.TextVelocidadviento.Text = oRead("Velocidad_viento")
                Pantalla_datos.TextFlechamaxcentrovano.Text = oRead("Flecha_máx_centro_vano")
                Pantalla_datos.TextDistanciacarrilposte.Text = oRead("Distancia_carril_poste")
                Pantalla_datos.TextDistanciabasepostepmr.Text = oRead("Distancia_base_poste_pmr")
                Pantalla_datos.TextDistanciaelectsecmecanico.Text = oRead("Distancia_eléct_sec_mecánico")
                Pantalla_datos.TextDistanciaelectsecelectrico.Text = oRead("Distancia_eléct_sec_eléctrico")
                Pantalla_datos.TextLongzonacomunmax.Text = oRead("Long_zona_común_máx")
                Pantalla_datos.TextLongzonacomunmin.Text = oRead("Long_zona_común_mín")
                Pantalla_datos.TextLongzonaneutra.Text = oRead("Long_zona_neutra")
                Pantalla_datos.ComboHilodecontacto.Text = oRead("Hilo_de_contacto")
                Pantalla_datos.ComboSustentador.Text = oRead("Sustentador")
                Pantalla_datos.ComboCdeproteccionaerea.Text = oRead("C_de_protección_aérea")
                Pantalla_datos.ComboCabledetierra.Text = oRead("Cable_de_tierra")
                Pantalla_datos.ComboFeederpositivo.Text = oRead("Feeder_positivo")
                Pantalla_datos.ComboFeedernegativo.Text = oRead("Feeder_negativo")
                Pantalla_datos.ComboPuntofijo.Text = oRead("Punto_fijo")
                Pantalla_datos.ComboPendola.Text = oRead("Péndola")
                Pantalla_datos.ComboAnclaje.Text = oRead("Anclaje")
                Pantalla_datos.ComboPosicionfeederpositivo.Text = oRead("Posición_feeder_positivo")
                Pantalla_datos.ComboPosicionfeedernegativo.Text = oRead("Posición_feeder_negativo")
                Pantalla_datos.TextNumhc.Text = oRead("Núm_hc")
                Pantalla_datos.TextNumcdpa.Text = oRead("Núm_cdpa")
                Pantalla_datos.TextNumfeederpositivo.Text = oRead("Núm_feeder_positivo")
                Pantalla_datos.TextNumfeedernegativo.Text = oRead("Núm_feeder_negativo")
                Pantalla_datos.TextTensionhc.Text = oRead("Tensión_hc")
                Pantalla_datos.TextTensionsustentador.Text = oRead("Tensión_sustentador")
                Pantalla_datos.TextTensioncdpa.Text = oRead("Tensión_cdpa")
                Pantalla_datos.TextTensionfeederpositivo.Text = oRead("Tensión_feeder_positivo")
                Pantalla_datos.TextTensionfeedernegativo.Text = oRead("Tensión_feeder_negativo")
                Pantalla_datos.TextTensionpuntofijo.Text = oRead("Tensión_punto_fijo")
                Pantalla_datos.ComboAdmlinea.Text = oRead("Adm_línea")
                Pantalla_datos.TextTipo.Text = oRead("Tipo")
                Pantalla_datos.ComboNumeración.Text = oRead("Numeración")
                Pantalla_datos.ComboAdmlineamacizo.Text = oRead("Adm_Línea_macizo")
                Pantalla_datos.TextTipomacizo.Text = oRead("Tipo_macizo")
                Pantalla_datos.ComboTubodemensula.Text = oRead("Tubo_de_ménsula")
                Pantalla_datos.ComboTubotirante.Text = oRead("Tubo_tirante")
                Pantalla_datos.ComboColadeanclaje.Text = oRead("Cola_de_anclaje")
                Pantalla_datos.ComboFeederpositivo.Text = oRead("Feeder_positivo")
                Pantalla_datos.ComboFeedernegativo.Text = oRead("Feeder_negativo")
                Pantalla_datos.TextDistanciaapoyoyprimerapendola.Text = oRead("Distancia_apoyo_y_1ª_péndola")
                Pantalla_datos.TextDistanciaprimeraysegundapendola.Text = oRead("Distancia_1ª_y_2ª_péndola")
                Pantalla_datos.TextDistanciamaxentrependolas.Text = oRead("Distancia_máx_entre_péndolas")



            End If

        End While

        oRead.Close()
        oConn.Close()
        Pantalla_datos.Show()


    End Sub
End Module
