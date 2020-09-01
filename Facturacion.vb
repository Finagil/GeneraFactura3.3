Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient

Module GneraFactura
    Dim ErrorControl As New EventLog
    Dim OpcionCompraAF As String
    Dim TipoCredito As String
    Public ProductDS As New ProduccionDS
    Dim CFDI_H As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
    Dim CFDI_D As New ProduccionDSTableAdapters.CFDI_DetalleTableAdapter
    Dim CFDI_P As New ProduccionDSTableAdapters.CFDI_ComplementoPagoTableAdapter
    Dim CUENTAS As New ProduccionDSTableAdapters.DatosCuentasTableAdapter
    Dim ROWheader As ProduccionDS.CFDI_EncabezadoRow
    Dim ROWdetail As ProduccionDS.CFDI_DetalleRow
    Dim SE_PROCESARON_FACTURAS As Boolean = False
    Dim Arg() As String

    Sub Main()
        'LecturaPreviaAUX()

        Dim mf As Date = Date.Now.AddHours(-72)
        Console.WriteLine("Inicia proceso")
        ErrorControl = New EventLog("Application", System.Net.Dns.GetHostName(), "GeneracionCFDI33")
        Dim D As New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        Dim F As FileInfo() = D.GetFiles("*.txt")

        Arg = Environment.GetCommandLineArgs()
        If Arg.Length > 1 Then
            Select Case UCase(Arg(1))
                Case "ACUSES_CAN"
                    NotificaCANF()
                    NotificaCANA()
                Case "ENVIA_RECIBOS"
                    Envia_RecibosPAGO()
                    '////Comentario temporal
                    'Case "FOLIOS"
                    'Console.WriteLine("Leyendo Folios CFDI ...")
                    'CFDI33.LeeFoliosFiscales()
                Case "AVISOS"
                    Console.WriteLine("Generando Avisos CFDI ...")

                    CFDI33.FacturarCFDI_AV(Date.Now.Date)

                    CFDI33.FacturarCFDI("PORVENCER")
                    CFDI33.FacturarCFDI("ANTERIORES")
                    CFDI33.FacturarCFDI("PREPAGO")
                    CFDI33.FacturarCFDI("DIA")
                    CFDI33.FacturarCFDI("MANUAL")
                Case "FACTURAS"
                    Console.WriteLine("Generando CFDI Avio...")
                    GeneraArchivosAvio()

                    Console.WriteLine("Generando CFDI Externas...")
                    GeneraArchivosEXternas()
                    Console.WriteLine("Generando CFDI Finagil...")
                    GeneraArchivos(True)
                Case "PAGOS"
                    Console.WriteLine("Generando CFDI Pago...")
                    GeneraArchivos(False) 'COMPLEMENTOS
                Case "FACTURAS_EKO"
                    Console.WriteLine("Generando CFDI Facturas EKomercio...")
                    CFDI33.GeneraFacturaEkomercio()
                Case "NOMINA_EKO"
                    Console.WriteLine("Generando CFDI Nomina EKomercio...")
                    CFDI33.GeneraRNominaekomercio()
                Case "PAGOS_EKO"
                    Console.WriteLine("Generando CFDI Pagos EKomercio...")
                    CFDI33.GeneraComplementoEkomercio()
                Case "WSKN"
                    Console.WriteLine("Subiendo Archivos Nomina EKomercio...")
                    CFDI33.SubeWSN()
                Case "WSK"
                    Console.WriteLine("Subiendo Archivos EKomercio...")
                    CFDI33.SubeWS()
                '////Comentario temporal
                'Case "FTP"
                'Console.WriteLine("Subiendo Archivos EKomercio...")
                'CFDI33.SubeFTP()
                Case "TODO_FTP"
                    Console.WriteLine("Leyendo Folios CFDI ...")
                    CFDI33.LeeFoliosFiscales()
                    If Date.Now.Hour >= 20 Or Date.Now.Hour <= 9 Then 'se ocupa despues de las 8pm y antes de las 9 am
                        Console.WriteLine("Generando Avisos CFDI ...")
                        CFDI33.FacturarCFDI_AV(Date.Now.Date)
                        CFDI33.FacturarCFDI("PORVENCER")
                        CFDI33.FacturarCFDI("ANTERIORES")
                        CFDI33.FacturarCFDI("PREPAGO")
                        CFDI33.FacturarCFDI("DIA")

                        Console.WriteLine("Generando CFDI Avio...")
                        GeneraArchivosAvio()
                    End If



                    Console.WriteLine("Generando CFDI Externas...")
                    GeneraArchivosEXternas()

                    Console.WriteLine("Generando CFDI Finagil...")
                    GeneraArchivos(True)

                    Console.WriteLine("Generando CFDI Pago...")
                    GeneraArchivos(False) 'COMPLEMENTOS

                    Console.WriteLine("Generando CFDI Facturas EKomercio...")
                    CFDI33.GeneraFacturaEkomercio()

                    Console.WriteLine("Generando CFDI Pagos EKomercio...")
                    CFDI33.GeneraComplementoEkomercio()

                    Console.WriteLine("Subiendo Archivos EKomercio...")
                    CFDI33.SubeFTP()

                    Console.WriteLine("Notificaciones de cancelación...")
                    CFDI33.NotificaCANF()
                    CFDI33.NotificaCANA()
                    FacturasSinSERIE()
                Case "TODO_WS"

                    'Console.WriteLine("Leyendo Folios CFDI ...")
                    'CFDI33.LeeFoliosFiscales()
                    If Date.Now.Hour >= 20 Or Date.Now.Hour <= 9 Then 'se ocupa despues de las 8pm y antes de las 9 am
                        Console.WriteLine("Generando Avisos CFDI ...")
                        CFDI33.FacturarCFDI_AV(Date.Now.Date)
                        CFDI33.FacturarCFDI("PORVENCER")
                        CFDI33.FacturarCFDI("ANTERIORES")
                        CFDI33.FacturarCFDI("PREPAGO")
                        CFDI33.FacturarCFDI("DIA")

                        Console.WriteLine("Generando CFDI Avio...")
                        GeneraArchivosAvio()
                    End If



                    Console.WriteLine("Generando CFDI Externas...")
                    GeneraArchivosEXternas()

                    Console.WriteLine("Generando CFDI Finagil...")
                    GeneraArchivos(True)

                    Console.WriteLine("Generando CFDI Pago...")
                    GeneraArchivos(False) 'COMPLEMENTOS

                    Console.WriteLine("Generando CFDI Facturas EKomercio...")
                    CFDI33.GeneraFacturaEkomercio()

                    Console.WriteLine("Generando CFDI Pagos EKomercio...")
                    CFDI33.GeneraComplementoEkomercio()

                    Console.WriteLine("Generando CFDI Nomina EKomercio...")
                    CFDI33.GeneraRNominaekomercio()

                    Console.WriteLine("Subiendo Archivos EKomercio...")
                    CFDI33.SubeWS()

                    Console.WriteLine("Subiendo Archivos Nomina EKomercio...")
                    CFDI33.SubeWSN()

                    Console.WriteLine("Notificaciones de cancelación...")
                    CFDI33.NotificaCANF()
                    CFDI33.NotificaCANA()

                    FacturasSinSERIE()
                Case "ORGANIZA_BACKUP"
                    Console.WriteLine("Orgaizando...")
                    Organiza_Backup(My.Settings.RutaBackup)
                    Organiza_Backup(My.Settings.BackupEKO)
            End Select
        End If
        Console.WriteLine("Terminado...")
    End Sub

    Sub GeneraArchivosAvio()
        Dim TasaIVACliente As Decimal
        Dim RFC As String
        Dim SubTT As Double
        Dim IVA As Double
        Dim Fega As Double
        Dim Razon As String
        Dim fecha As New DateTime
        Dim taCli As New GeneraFactura.ProduccionDSTableAdapters.ClientesTableAdapter
        Dim Facturas As New GeneraFactura.ProduccionDSTableAdapters.FacturasAvioTableAdapter
        Dim FAC As New GeneraFactura.ProduccionDS.FacturasAvioDataTable
        Dim Detalles As New GeneraFactura.ProduccionDSTableAdapters.FacturasAvioDetalleTableAdapter
        Dim DET As New GeneraFactura.ProduccionDS.FacturasAvioDetalleDataTable
        Dim Folios As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter

        Dim RegAfec As Integer
        Dim Concep As String
        Dim ContLin As Integer

        Dim ProducDS As New ProduccionDS
        Dim CFDI_H As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim CFDI_D As New ProduccionDSTableAdapters.CFDI_DetalleTableAdapter
        Dim ROWheader As ProduccionDS.CFDI_EncabezadoRow
        Dim ROWdetail As ProduccionDS.CFDI_DetalleRow

        If Date.Now.Day <= 3 Then
            fecha = Date.Now.AddDays(Date.Now.Day * -1)
        Else
            fecha = Date.Now
        End If
        Detalles.QuitarPagosEfectivo()
        '***************************************************************
        'quita seguros que nos sean de guanajuato y michoacan haste que se hagan dos conceptos de seguros
        Detalles.FillBySeguro(DET)
        For Each rr As GeneraFactura.ProduccionDS.FacturasAvioDetalleRow In DET.Rows
            Detalles.Facturar(False, "N/A", rr.Anexo, rr.Ciclo, rr.FechaFinal, rr.Concepto)
        Next
        '***************************************************************
        Detalles.QuitarPagosEfectivo()
        Facturas.Fill(FAC, fecha.ToString("yyyyMMdd"))

        For Each r As GeneraFactura.ProduccionDS.FacturasAvioRow In FAC.Rows
            TasaIVACliente = r.IvaAnexo
            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
            ROWheader._1_Folio = Folios.SerieGVA
            Console.WriteLine("Generando CFDI AVIO..." & r.Anexo & " " & ROWheader._1_Folio)

            SubTT = 0
            IVA = 0
            Fega = 0
            ContLin = 0

            ROWheader._2_Nombre_Emisor = "FINAGIL S.A. DE C.V, SOFOM E.N.R"
            ROWheader._3_RFC_Emisor = "FIN940905AX7"
            ROWheader._4_Dom_Emisor_calle = "Leandro Valle"
            ROWheader._5_Dom_Emisor_noExterior = "402"
            ROWheader._6_Dom_Emisor_noInterior = ""
            ROWheader._7_Dom_Emisor_colonia = "Reforma y F.F.C.C"
            ROWheader._8_Dom_Emisor_localidad = "Toluca"
            ROWheader._9_Dom_Emisor_referencia = ""
            ROWheader._10_Dom_Emisor_municipio = "Toluca"
            ROWheader._11_Dom_Emisor_estado = "Estado de México"
            ROWheader._12_Dom_Emisor_pais = "México"
            ROWheader._13_Dom_Emisor_codigoPostal = "50070"

            ROWheader._26_Version = "3.3"
            ROWheader._27_Serie_Comprobante = "AV"
            ROWheader._29_FormaPago = "27" '27 A satisfacción del acreedor
            ROWheader._30_Fecha = fecha.Date
            ROWheader.Fecha = fecha.Date
            ROWheader._31_Hora = fecha.ToString("HH:mm:ss")
            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

            RFC = Trim(r.RFC)
            RFC = ValidaRFC(RFC, r.Tipo)
            If RFC = "SDA070613KU6" Then
                Razon = """SERVICIO DAYCO"" SA DE CV"
            Else
                Razon = r.Nombre
            End If


            ROWheader._42_Nombre_Receptor = Razon.Trim
            ROWheader._43_RFC_Receptor = RFC.Trim
            ROWheader._44_Dom_Receptor_calle = r.Calle.Trim
            ROWheader._45_Dom_Receptor_noExterior = ""
            ROWheader._46_Dom_Receptor_noInterior = ""
            ROWheader._47_Dom_Receptor_colonia = r.Colonia.Trim
            ROWheader._48_Dom_Receptor_localidad = ""
            ROWheader._49_Dom_Receptor_referencia = ""
            ROWheader._50_Dom_Receptor_municipio = r.Delegacion.Trim
            ROWheader._51_Dom_Receptor_estado = r.Estado.Trim
            ROWheader._52_Dom_Receptor_pais = "México"
            ROWheader._53_Dom_Receptor_codigoPostal = r.Copos.Trim
            ROWheader._57_Estado = 1
            ROWheader._58_TipoCFD = "FA"
            ROWheader._83_Cod_Moneda = r.Moneda.Trim
            ROWheader._97_Condiciones_Pago = "Contado"
            ROWheader._144_Misc32 = "G03" 'claves del SAT P01=por definir, G03=Gastos generales CATALOGO USO DE COMPROBANTE hoja excel= c_UsoCFDI, a solicitud del cliente
            ROWheader._167_RegimentFiscal = 601
            If ROWheader._83_Cod_Moneda = "MXN" Then
                ROWheader._177_Tasa_Divisa = 0
            Else
                ROWheader._177_Tasa_Divisa = taCli.SacaTipoCambio(fecha, ROWheader._83_Cod_Moneda)
            End If
            ROWheader._180_LugarExpedicion = "50070"
            ROWheader._190_Metodo_Pago = "PUE"
            ROWheader._191_Efecto_Comprobante = "I"

            Detalles.QuitarPagosEfectivo()
            Detalles.Fill(DET, r.Anexo, r.Ciclo)
            Dim TasaIVA As Decimal = 0
            Dim TipoImpuesto As String

            Fega = 0
            For Each rr As GeneraFactura.ProduccionDS.FacturasAvioDetalleRow In DET.Rows
                ROWdetail = ProducDS.CFDI_Detalle.NewCFDI_DetalleRow
                TasaIVA = TasaIVACliente / 100
                TipoImpuesto = TasaIVACliente
                If Trim(rr.Concepto) = "EFECTIVO" _
                    Or Trim(rr.Concepto) = "AGROQUIMICOS" _
                        Or Trim(rr.Concepto) = "EFECTIVO2" _
                    Or Trim(rr.Concepto) = "EFECTIVO 2" _
                    Or Trim(rr.Concepto) = "VALE" _
                    Or Trim(rr.Concepto) = "ASISTENCIA" _
                    Or Trim(rr.Concepto) = "INTEGRACION" _
                    Or Trim(rr.Concepto) = "ANALISIS DE SUELOS" _
                    Or Trim(rr.Factura) = "N/A" Then
                    Fega += rr.FEGA
                Else
                    ContLin += 1
                    Fega += rr.FEGA
                    Concep = Trim(rr.Concepto)
                    Select Case UCase(Concep)
                        Case "NOTARIO"
                            Concep = "GASTOS DE NOTARIO"
                        Case "RPP"
                            Concep = "REGISTRO PÚBLICO DE LA PROPIEDAD"
                        Case "GASTOS"
                            Concep = "GASTOS ADMINISTRATIVOS"
                        Case "ASISTENCIA"
                            Concep = "ASISTENCIA TÉCNICA"
                        Case "SEGURO"
                            If r.Tipar <> "C" Then
                                If r.Tipo <> "F" Then
                                    TipoImpuesto = "Exento"
                                    Concep = "SEGURO AGRÍCOLA EXENTOS DE IVA"
                                Else
                                    Concep = "SEGURO AGRÍCOLA"
                                End If
                            End If
                        Case "SEGURO DE VIDA"
                            TipoImpuesto = "Exento"
                            Concep = "SEGURO DE VIDA EXENTOS DE IVA (" & UCase(MonthName(CInt(Mid(r.FechaFinal, 5, 2)), True)) & "-" & Mid(r.FechaFinal, 1, 4) & ")"
                        Case "INTERESES", "INTERESES POR PREPAGO", "INTERESES POR PREPAGO SEGURO", "INTERESES POR PREPAGO OTROS"
                            If r.Tipo <> "F" Then
                                TipoImpuesto = "Exento"
                                Concep = Concep & " EXENTOS DE IVA"
                            End If
                        Case "INTERESES MORATORIOS"
                            If r.Tipo <> "F" Then
                                TipoImpuesto = "Exento"
                                Concep = Concep & " EXENTOS DE IVA"
                            End If
                    End Select

                    ROWdetail._1_Linea_Descripcion = Concep
                    ROWdetail._2_Linea_Cantidad = 1
                    ROWdetail._3_Linea_Unidad = "E48"
                    ROWdetail._4_Linea_PrecioUnitario = rr.Importe.ToString("n2")
                    ROWdetail._5_Linea_Importe = rr.Importe.ToString("n2")
                    ROWdetail._16_Linea_Cod_Articulo = "84101700" ' Manejo de deuda
                    ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                    ROWdetail._2_Impuesto_Descripcion = "TR"
                    ROWdetail._3_Impuesto_Monto_base = rr.Importe.ToString("n2")
                    ROWdetail._5_Impuesto_Clave = "002"
                    ROWdetail._6_Impuesto_Tasa = "Tasa"
                    If TipoImpuesto = "Exento" Then
                        ROWdetail._7_Impuesto_Porcentaje = ""
                        ROWdetail._4_Impuesto_Monto_Impuesto = ""
                        ROWdetail._6_Impuesto_Tasa = "Exento"
                    Else
                        ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                        ROWdetail._4_Impuesto_Monto_Impuesto = TruncarDecimales((ROWdetail._5_Linea_Importe * TasaIVA))
                    End If

                    SubTT += ROWdetail._3_Impuesto_Monto_base
                    If IsNumeric(ROWdetail._4_Impuesto_Monto_Impuesto) Then
                        IVA += CDec(ROWdetail._4_Impuesto_Monto_Impuesto).ToString("n2")
                    End If

                    ROWdetail._11_Linea_Notas = "SER"
                    ROWdetail._53_Linea_Misc22 = "SER"
                    ROWdetail.Detalle_Folio = ROWheader._1_Folio
                    ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante

                    ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                End If

                RegAfec = Detalles.Facturar(True, "AV" & ROWheader._1_Folio, r.Anexo, r.Ciclo, rr.FechaFinal, Trim(rr.Concepto))
                If RegAfec = 0 Then
                    'EnviaError(GeneraFactura.My.Settings.MailError, "Error Factura sin Afectar", "Error Factura sin Afectar" & r.Anexo)
                End If

            Next
            If Fega > 0 Then
                ContLin += 1
                ROWdetail = ProducDS.CFDI_Detalle.NewRow
                Concep = "GARANTIA FEGA"
                TasaIVA = TasaIVACliente / 100
                ROWdetail._1_Linea_Descripcion = Concep
                ROWdetail._2_Linea_Cantidad = 1
                ROWdetail._3_Linea_Unidad = "E48"
                ROWdetail._4_Linea_PrecioUnitario = Fega.ToString("n2")
                ROWdetail._5_Linea_Importe = Fega.ToString("n2")
                ROWdetail._16_Linea_Cod_Articulo = "84101700"
                ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                ROWdetail._2_Impuesto_Descripcion = "TR"
                ROWdetail._3_Impuesto_Monto_base = Fega.ToString("n2")
                ROWdetail._4_Impuesto_Monto_Impuesto = TruncarDecimales((ROWdetail._5_Linea_Importe * TasaIVA))
                ROWdetail._5_Impuesto_Clave = "002" ' Clave IVA
                ROWdetail._6_Impuesto_Tasa = "Tasa"
                ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                ROWdetail._11_Linea_Notas = "SER"
                ROWdetail._53_Linea_Misc22 = "SER"
                ROWdetail.Detalle_Folio = ROWheader._1_Folio
                ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante

                SubTT += ROWdetail._3_Impuesto_Monto_base
                IVA += CDec(ROWdetail._4_Impuesto_Monto_Impuesto).ToString("n2")

                ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
            End If


            ROWheader._90_Cantidad_LineasFactura = ContLin
            ROWheader._54_Monto_SubTotal = SubTT
            ROWheader._55_Monto_IVA = IVA
            ROWheader._56_Monto_Total = ROWheader._54_Monto_SubTotal + ROWheader._55_Monto_IVA
            ROWheader._193_Monto_TotalImp_Trasladados = ROWheader._55_Monto_IVA
            ROWheader._100_Letras_Monto_Total = Letras(ROWheader._56_Monto_Total, "MXN")
            ROWheader._114_Misc02 = r.AnexoCon & " " & r.CicloPagare
            ROWheader._115_Misc03 = r.Anexo & "-" & r.Ciclo & "-" & r.FechaFinal
            ROWheader._162_Misc50 = ""
            ROWheader.Encabezado_Procesado = False

            If r.Cliente = "05978 " Then
                ROWheader._162_Misc50 = Trim(r.EMail1) & ";" & Trim(r.EMail2) & ";flen.estrada@ciasaconstruccion.com.mx;administacion@ciasaconstruccion.com.mx"
            End If

            If r.Tipar <> "C" Then
                Select Case Trim(r.Sucursal)
                    Case "MEXICALI"
                        ROWheader._162_Misc50 += ";sduarte@finagil.com.mx"
                    Case "NAVOJOA"
                        ROWheader._162_Misc50 += ";mlopezb@finagil.com.mx"
                    Case "IRAPUATO"
                        ROWheader._162_Misc50 += ";vtezcuc@finagil.com.mx"
                    Case Else
                        ROWheader._162_Misc50 += ";vcruz@finagil.com.mx"
                End Select
            Else
                ROWheader._162_Misc50 += ";vcruz@finagil.com.mx"
            End If
            ProducDS.CFDI_Encabezado.AddCFDI_EncabezadoRow(ROWheader)
            ProducDS.CFDI_Encabezado.GetChanges()
            ProducDS.CFDI_Detalle.GetChanges()
            CFDI_H.Update(ProducDS.CFDI_Encabezado)
            CFDI_D.Update(ProducDS.CFDI_Detalle)
            Folios.ConsumeFolio()

            ''Catch ex As Exception
            ''    EnviaError(GeneraFactura.My.Settings.MailError, ex.Message, "Error CFDI AVIO" & r.Anexo)
            ''End Try
            'End If
        Next

    End Sub

    Sub GeneraArchivos(EsFactura As Boolean)
        Dim EsFacturaAux As Boolean = EsFactura
        Dim NoFactError As Integer
        Dim Folio, FolioORG As Integer
        Dim Serie As String = ""
        Dim GUID As String = ""
        Dim taCli As New GeneraFactura.ProduccionDSTableAdapters.ClientesTableAdapter
        Dim Facturas As New GeneraFactura.ProduccionDSTableAdapters.FacturasAvioTableAdapter
        Dim FAC As New GeneraFactura.ProduccionDS.FacturasAvioDataTable
        Dim Detalles As New GeneraFactura.ProduccionDSTableAdapters.FacturasAvioDetalleTableAdapter
        Dim DET As New GeneraFactura.ProduccionDS.FacturasAvioDetalleDataTable
        Dim Folios As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim ProducDS As New ProduccionDS
        Dim TasaIVACliente As Decimal
        Dim IVACapital As String
        Dim SubTT, IVA, MontoBaseIVA As Decimal
        Dim NoLineas As Integer
        Dim EsNotaCredito As Boolean = False
        Dim EnviarGisela As Boolean = False
        Dim ta As New GeneraFactura.ProduccionDSTableAdapters.ClientesTableAdapter
        Dim t As New GeneraFactura.ProduccionDS.ClientesDataTable
        Dim taEnc As New GeneraFactura.ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim tEnc As New GeneraFactura.ProduccionDS.CFDI_EncabezadoDataTable
        Dim taMail As New ProduccionDSTableAdapters.CorreosAnexosTableAdapter
        Dim tMail As New ProduccionDS.CorreosAnexosDataTable
        Dim Rmail As ProduccionDS.CorreosAnexosRow
        Dim taCodigo As New ProduccionDSTableAdapters.CodigosSATTableAdapter
        Dim tCodigo As New ProduccionDS.CodigosSATDataTable
        Dim rCod As ProduccionDS.CodigosSATRow
        Dim Linea, Mail As String
        Dim suma As Double
        Dim Codigo, Unidad, UsoCFDI, Concepto As String
        Dim Adenda, Errores As Boolean
        Dim D As System.IO.DirectoryInfo
        Dim F As System.IO.FileInfo()
        If EsFactura = True Then
            D = New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
            F = D.GetFiles("*.txt").OrderBy(Function(fi) fi.Name).ToArray()
        Else
            D = New System.IO.DirectoryInfo(GeneraFactura.My.Settings.Complementos)
            F = D.GetFiles("*.txt").OrderBy(Function(fi) fi.Name).ToArray()
        End If
        Dim Aviso, SerieORG As String
        Dim Datos() As String
        Dim DatosFinagil() As String
        Dim f2 As System.IO.StreamReader
        Dim fecha As New DateTime
        Dim fecha_pago As New DateTime
        fecha_pago = Nothing
        Dim horas As Integer
        Dim IvaAux As Decimal
        Dim Tipar As String = ""
        Dim TipoPersona As String = ""
        Dim Moneda As String = ""
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim cAnexo, cAnexoAux As String
        Dim LeyendaCapital, Metodo_Pago, FormaPago, Referencia As String
        Dim RFC_BancoCliente, NombreBancoCliente, NoCuentaCliente As String
        Dim RFC_BancoFinagil, CuentaFinagil As String
        Dim Spei, SpeiCert, SpeiCadOrg, SpeiSello As String
        Dim EsPago As Boolean
        Dim TipoCambioSTR, ErrorMSG As String
        Dim SinFolio As String = ""
        'Try
        NoFactError = 0
        For i = 0 To F.Length - 1
            'Try
            EsFactura = EsFacturaAux
            Console.WriteLine("Generando CFDI..." & F(i).Name)
            NoLineas = 0
            suma = 0
            Mail = ""
            EsPago = False
            EsNotaCredito = False
            SubTT = 0
            IVA = 0
            EnviarGisela = False
            Adenda = False
            LeyendaCapital = ""
            FolioORG = 0
            SerieORG = ""
            Errores = False
            ReDim Datos(1)
            Datos(0) = "X"
            cAnexoAux = ""
            RFC_BancoCliente = ""
            NombreBancoCliente = ""
            NoCuentaCliente = ""
            RFC_BancoFinagil = ""
            CuentaFinagil = ""
            TipoCambioSTR = ""
            Spei = ""
            SpeiCert = ""
            SpeiCadOrg = ""
            SpeiSello = ""
            LecturaPrevia(F(i).FullName, F(i).Name, Moneda, Tipar, Folio, Serie, EsFactura, EsPago, SerieORG, FolioORG, GUID, Referencia, Aviso, IVACapital)
            'If LecturaPrevia(F(i).FullName, F(i).Name, Moneda, Tipar, Folio, Serie, EsPago) Then
            '    File.Copy(F(i).FullName, GeneraFactura.My.Settings.Raiz & F(i).Name, True)
            '    File.Delete(F(i).FullName)
            '    Continue For
            'Else
            '    Continue For
            'End If
            If EsPago = True And EsFactura = False And Serie <> "C" Then
#Region "Espago"
                If GUID = "SIN FOLIO FISCAL" Then
                    SinFolio += "Tipo de Cambio : 1 Concepto: " & "SIN FOLIO FISCAL" & vbCrLf & " TipoCredito : " & Tipar & " Anexo : " & cAnexo & " Factura sin Procesar " & Serie & Folio & "<BR>"
                    Continue For
                End If
                f2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
                If Mid(F(i).Name, 1, 3) <> "FIN" And Mid(F(i).Name, 1, 3) <> "XXA" And IsNumeric(Mid(F(i).Name, 1, 4)) = True Then
                    fecha = New DateTime(Mid(F(i).Name, 1, 4), Mid(F(i).Name, 5, 2), Mid(F(i).Name, 7, 2), Mid(F(i).Name, 9, 2), Mid(F(i).Name, 11, 2), Mid(F(i).Name, 13, 2))
                    horas = DateDiff(DateInterval.Hour, fecha, Date.Now)
                    If horas >= 72 Then
                        fecha = fecha.AddHours(horas - 71)
                    End If
                End If

                While Not f2.EndOfStream
                    Linea = f2.ReadLine
                    If UCase(Linea) = "X" Then
                        EnviarGisela = True
                        Linea = f2.ReadLine
                    End If
                    Datos = Linea.Split("|")
                    If Datos.Length > 4 Then
                        cAnexoAux = Datos(2)
                        If Datos(2) = "03282/0002" Then Datos(2) = "2885803-001"
                        If Datos(2) = "01350/0012" Then Datos(2) = "10318141001"
                    End If

                    Select Case Datos(0)
                        Case "M1"
                            fecha = Datos(6)
                            Mail = Datos(5)
                        Case "H1"
                            fecha = Datos(1)
                            If DateDiff(DateInterval.Hour, fecha, Date.Now) > 72 Then
                                fecha = Date.Now.AddDays(-3)
                                fecha = fecha.AddHours(2)
                            Else
                                'pone la hora a la fecha del archivo
                                fecha = fecha.Date
                                fecha = fecha.AddHours(Date.Now.Hour + 1)
                                fecha = fecha.AddMinutes(Date.Now.Minute)
                                fecha = fecha.AddSeconds(Date.Now.Second)
                            End If
                            Metodo_Pago = Datos(2)
                            FormaPago = Datos(3)
                            If Datos.Length > 5 Then
                                fecha_pago = Datos(5)
                            End If
                        Case "H3"
                            'Aviso = CFDI_H.SacaAvisoCFDI(SerieORG, FolioORG)
                            If Datos(2).Length <> 10 Then
                                cAnexo = Mid(cAnexoAux, 1, 5) & Mid(cAnexoAux, 7, 4)
                            Else
                                cAnexo = Mid(Datos(2), 1, 5) & Mid(Datos(2), 7, 4)
                            End If

                            TipoPersona = taTipar.TipoPersona(Datos(1))
                            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
                            TasaIVACliente = taCli.SacaTasaIvaAnexo(cAnexo)
                            ROWheader._1_Folio = Folio
                            ROWheader._2_Nombre_Emisor = "FINAGIL S.A. DE C.V, SOFOM E.N.R"
                            ROWheader._3_RFC_Emisor = "FIN940905AX7"
                            ROWheader._4_Dom_Emisor_calle = "Leandro Valle"
                            ROWheader._5_Dom_Emisor_noExterior = "402"
                            ROWheader._6_Dom_Emisor_noInterior = ""
                            ROWheader._7_Dom_Emisor_colonia = "Reforma y F.F.C.C"
                            ROWheader._8_Dom_Emisor_localidad = "Toluca"
                            ROWheader._9_Dom_Emisor_referencia = ""
                            ROWheader._10_Dom_Emisor_municipio = "Toluca"
                            ROWheader._11_Dom_Emisor_estado = "Estado de México"
                            ROWheader._12_Dom_Emisor_pais = "México"
                            ROWheader._13_Dom_Emisor_codigoPostal = "50070"

                            ROWheader._26_Version = "3.3"
                            ROWheader._27_Serie_Comprobante = Serie
                            ROWheader._29_FormaPago = "" 'FormaPago '"27" '27 A satisfacción del acreedor
                            ROWheader._30_Fecha = Today.Date.ToString("yyyy-MM-dd") ' se manda con fecha y hora de ejecucion
                            ROWheader._31_Hora = Today.ToString("HH:mm:ss") ' se manda con fecha y hora de ejecucion
                            ROWheader.Fecha = fecha.Date
                            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

                            If Moneda = "WWW" Then
                                If Datos(17) = "M.N." Then Datos(17) = "MXN"
                                If Datos(17) = "MXP" Then Datos(17) = "MXN"
                                Moneda = Datos(17)
                            End If

                            If Moneda <> "MXN" Then
                                TipoCambioSTR = CFDI_P.SacaTipoCambio(fecha.Date, Moneda).ToString
                            End If

                            Datos(16) = ValidaRFC(Datos(16), TipoPersona)
                            If Trim(Datos(16)) = "SDA070613KU6" Then
                                Datos(5) = """SERVICIO DAYCO"" SA DE CV"
                            End If
                            If Trim(Datos(16)) = "CARD840606LEA" Then
                                Datos(5) = "DANIEL CADENA RUVALCABA"
                            End If
                            If Trim(Datos(16)) = "GET090828K63" Then
                                Errores = True
                                ErrorMSG = "!!No se Factura de Grupo empresarian transforma!!"
                            End If

                            If FormaPago = "03" Or FormaPago = "02" Then ' tranferencias
                                RFC_BancoFinagil = CUENTAS.DatosBancoFinagil(Datos(3), Datos(4))
                                If Not IsNothing(RFC_BancoFinagil) Then
                                    DatosFinagil = RFC_BancoFinagil.Split("|")
                                    RFC_BancoFinagil = DatosFinagil(0)
                                    CuentaFinagil = DatosFinagil(1)
                                End If

                                CUENTAS.FillporCliente(ProductDS.DatosCuentas, Datos(16).ToUpper)
                                If ProductDS.DatosCuentas.Rows.Count > 0 Then
                                    RFC_BancoCliente = ProductDS.DatosCuentas.Rows(0).Item("RFC_Banco")
                                    NombreBancoCliente = ProductDS.DatosCuentas.Rows(0).Item("Nombre")
                                    NoCuentaCliente = ProductDS.DatosCuentas.Rows(0).Item("NoCuenta")
                                    If FormaPago = "03" And (NoCuentaCliente.Length <> 10 Or NoCuentaCliente.Length <> 18) Then
                                        NombreBancoCliente = ""
                                        NoCuentaCliente = ""
                                        RFC_BancoCliente = ""
                                    End If
                                    If FormaPago = "02" And (NoCuentaCliente.Length <> 11 Or NoCuentaCliente.Length <> 18) Then
                                        NombreBancoCliente = ""
                                        NoCuentaCliente = ""
                                        RFC_BancoCliente = ""
                                    End If
                                End If
                            End If


                            ROWheader._42_Nombre_Receptor = Datos(5).Trim
                            ROWheader._43_RFC_Receptor = Datos(16).ToUpper
                            ROWheader._44_Dom_Receptor_calle = Datos(6).Trim
                            ROWheader._45_Dom_Receptor_noExterior = ""
                            ROWheader._46_Dom_Receptor_noInterior = ""
                            ROWheader._47_Dom_Receptor_colonia = Datos(9).Trim
                            ROWheader._48_Dom_Receptor_localidad = ""
                            ROWheader._49_Dom_Receptor_referencia = ""
                            ROWheader._50_Dom_Receptor_municipio = Datos(10).Trim
                            ROWheader._51_Dom_Receptor_estado = Datos(11).Trim
                            ROWheader._52_Dom_Receptor_pais = Datos(15).Trim
                            ROWheader._53_Dom_Receptor_codigoPostal = Datos(12).Trim
                            ROWheader._57_Estado = 1

                            ROWheader._83_Cod_Moneda = Moneda
                            ROWheader._97_Condiciones_Pago = ""
                            ROWheader._144_Misc32 = "P01" 'UsoCFDI 'claves del SAT P01=por definir, G03=Gastos generales CATALOGO USO DE COMPROBANTE hoja excel= c_UsoCFDI, a solicitud del cliente
                            ROWheader._167_RegimentFiscal = 601
                            If ROWheader._83_Cod_Moneda = "MXN" Then
                                ROWheader._177_Tasa_Divisa = 0
                            Else
                                ROWheader._177_Tasa_Divisa = Nothing ' taCli.SacaTipoCambio(fecha.Date, ROWheader._83_Cod_Moneda)
                                If ROWheader._177_Tasa_Divisa = 1 Then
                                    Errores = True
                                    EnviaCorreoFASE("Contabilidad", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ErrorMSG & "Tipo de Cambio : 1 Concepto: " & Concepto & vbCrLf & " TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo)
                                    EnviaCorreoFASE("Desarrollo", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ErrorMSG & "Tipo de Cambio : 1 Concepto: " & Concepto & vbCrLf & " TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo)
                                End If
                            End If
                            ROWheader._180_LugarExpedicion = "50070"
                            ROWheader._190_Metodo_Pago = "" 'PPD pago en parcialidades PUE pago en una sola exhibision
                            ROWheader._191_Efecto_Comprobante = "P"
                            ROWheader._58_TipoCFD = "FA"
                            ROWheader._83_Cod_Moneda = "XXX"
                        Case "D1"
                            Dim TasaIVA As Decimal = 0
                            Dim TipoImpuesto As String
                            If InStr(UCase(Datos(8)), "IVA") <= 0 Then
                                Datos(8) = Trim(Datos(8))
                                Concepto = LimpiarConcepto(Datos(8), ROWheader._27_Serie_Comprobante)
                                taCodigo.Fill(tCodigo, Tipar, Concepto)
                                If tCodigo.Rows.Count > 0 Then
                                    rCod = tCodigo.Rows(0)
                                    If rCod.Adenda = True Then
                                        LeyendaCapital += Concepto & " " & CDec(Datos(10)).ToString("n2") & vbCrLf
                                        Continue While
                                    End If
                                End If
                                TasaIVA = TasaIVACliente / 100
                                TipoImpuesto = TasaIVACliente
                                NoLineas += 1
                                If NoLineas = 1 Then
                                    ROWdetail = ProducDS.CFDI_Detalle.NewCFDI_DetalleRow
                                End If
                                If Tipar <> "F" And Tipar <> "P" Then
                                    Select Case Datos(8)
                                        Case "ADELANTO CAPITAL EQUIPO"
                                            Datos(8) = "ADELANTO CAPITAL"
                                        Case "SALDO INSOLUTO EQUIPO"
                                            Datos(8) = "SALDO INSOLUTO"
                                        Case "SALDO INSOLUTO DEL EQUIPO"
                                            Datos(8) = "SALDO INSOLUTO"
                                        Case "CAPITAL EQUIPO"
                                            Datos(8) = "CAPITAL"
                                    End Select
                                    If InStr(Datos(8), "CAPITAL EQUIPO VEN") > 0 Then
                                        Datos(8) = "CAPITAL VENCIMIENTO" '& Right(Datos(8), 7)
                                    End If
                                End If
                                If (Tipar = "F") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "INTERES OTRO", "INTERES SEGU"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If (Tipar = "R") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "MORATORIOS V"
                                            TipoImpuesto = "Exento"
                                        Case "INTERESES VE", "INTERES OTRO", "INTERES SEGU", "INTERESES PO"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If (Tipar = "S") And TipoPersona <> "F" Then
                                    If TipoPersona = "E" Then 'debe contar con autorizacion
                                        If CFDI_H.AutorizarIVA_Interes(cAnexo, "") > 0 Then
                                            Select Case Mid(Datos(8), 1, 12)
                                                Case "MORATORIOS V"
                                                    TipoImpuesto = "Exento"
                                                Case "INTERESES VE", "INTERES OTRO", "INTERES SEGU", "INTERESES PO"
                                                    TipoImpuesto = "Exento"
                                            End Select
                                        End If
                                    Else
                                        Select Case Mid(Datos(8), 1, 12)
                                            Case "MORATORIOS V"
                                                TipoImpuesto = "Exento"
                                            Case "INTERESES VE", "INTERES OTRO", "INTERES SEGU", "INTERESES PO"
                                                TipoImpuesto = "Exento"
                                        End Select
                                    End If

                                End If
                                If Tipar = "F" And cAnexo = "038240001" Then
                                    TipoImpuesto = "Exento"
                                End If
                                If Tipar = "F" And cAnexo = "025620003" Then '#ECT Solicitado por Valentin 24/09/2015
                                    Select Case Datos(8)
                                        Case "ADELANTO CAPITAL EQUIPO"
                                            Datos(8) = Datos(8) '& " A TASA IVA 0%"
                                            TasaIVA = 0
                                        Case "CAPITAL EQUIPO"
                                            Datos(8) = Datos(8) '& " A TASA IVA 0%"
                                            TasaIVA = 0
                                        Case "AMORTIZACION INICIAL"
                                            Datos(8) = Datos(8) '& " A TASA IVA 0%"
                                            TasaIVA = 0
                                    End Select
                                    If Mid(Datos(8), 1, 9) = "INTERESES" Then
                                        TipoImpuesto = "Exento"
                                    End If
                                End If

                                If (Tipar = "A" Or Tipar = "H" Or Tipar = "C") Then
                                    If Tipar = "C" Then
                                        Select Case Mid(Datos(8), 1, 12)
                                            Case "INTERESES AV"
                                                Datos(8) = "INTERESES CUENTA CORRIENTE"
                                            Case "INTERESES MO"
                                                Datos(8) = "INTERESES MORATORIO CUENTA CORRIENTE"
                                            Case "PAGO CREDITO"
                                                Datos(8) = "PAGO CREDITO EN CUENTA CORRIENTE"
                                        End Select
                                    End If

                                    If TipoPersona <> "F" And Mid(Datos(8), 1, 9) = "INTERESES" Then
                                        TipoImpuesto = "Exento"
                                    End If
                                End If

                                If InStr(Datos(8), "SEGURO DE VI") > 0 Then
                                    TipoImpuesto = "Exento"
                                End If

                                If Tipar = "P" Then
                                    Select Case Datos(8)
                                        Case "AMORTIZACION INICIAL"
                                            Datos(8) = "RENTA INICIAL"
                                    End Select
                                    If CDec(Datos(11)) = 0 Then
                                        TipoImpuesto = "Exento"
                                    End If
                                End If

                                If Tipar = "B" Then
                                    Select Case Mid(Datos(8), 1, 11)
                                        Case "MENSUALIDAD"
                                            Datos(8) = "SERVICIO DE TRANSPORTE EJECUTIVO EMPRESARIAL, " & Datos(8)
                                        Case "MORATORIOS "
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If

                                IvaAux = Math.Round(CDec(Datos(11)) / CDec(Datos(10)), 3)
                                If IvaAux > TasaIVA Then
                                    Datos(11) = Math.Round(CDec(Datos(10)) * TasaIVA, 2)
                                End If

                                ROWdetail._1_Linea_Descripcion = "Pago"
                                ROWdetail._2_Linea_Cantidad = 1
                                ROWdetail._3_Linea_Unidad = "ACT"
                                ROWdetail._4_Linea_PrecioUnitario = 0
                                ROWdetail._5_Linea_Importe = 0
                                ROWdetail._16_Linea_Cod_Articulo = "84111506" 'Codigo ' Manejo de deuda
                                If Datos(11) = "" Then Datos(11) = 0
                                SubTT += CDec(Datos(10)).ToString("n2")
                                IVA += CDec(Datos(11)).ToString("n2")

                                ROWdetail.Detalle_Folio = ROWheader._1_Folio
                                ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante
                                Try
                                    If NoLineas = 1 Then
                                        ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                                    End If
                                Catch ex As Exception
                                    Errores = False
                                    EnviaCorreoFASE("Desarrollo", "Error de Factura " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ex.Message & " " & ErrorMSG)
                                End Try

                            End If
                    End Select
                End While
                If Datos(0) <> "X" Then
                    ROWheader._90_Cantidad_LineasFactura = 0
                    ROWheader._54_Monto_SubTotal = 0 ' SubTT
                    ROWheader._55_Monto_IVA = 0 'IVA
                    ROWheader._56_Monto_Total = 0 'ROWheader._54_Monto_SubTotal + ROWheader._55_Monto_IVA
                    ROWheader._193_Monto_TotalImp_Trasladados = 0 'ROWheader._55_Monto_IVA
                    ROWheader._100_Letras_Monto_Total = "" ' Letras(ROWheader._56_Monto_Total, Moneda)
                    ROWheader._113_Misc01 = "[CPG_FINAGIL]"
                    ROWheader._114_Misc02 = Datos(2)
                    ROWheader._115_Misc03 = Datos(1)
                    ROWheader._132_Misc20 = "[CPG]"
                    ROWheader._158_Misc46 = TipoCredito.Trim
                    ROWheader._159_Misc47 = Aviso
                    ROWheader._162_Misc50 = ""
                    'ROWheader._161_Misc49 = ""
                    If OpcionCompraAF.Trim.Length > 0 Then
                        ROWheader._157_Misc45 = LeyendaCapital.Trim & " " & OpcionCompraAF.Trim
                    Else
                        ROWheader._157_Misc45 = LeyendaCapital.Trim
                    End If


                    ROWheader.Encabezado_Procesado = Errores

                    If EnviarGisela = False Then
                        'CORREOS ADICIONALES++++++++++++++++++++++++++++
                        taMail.Fill(tMail, cAnexo)
                        If tMail.Rows.Count > 0 Then
                            For Each Rmail In tMail.Rows
                                If Rmail.Correo1 > "" Then
                                    ROWheader._162_Misc50 += ";" & Rmail.Correo1
                                End If
                                If Rmail.Correo2 > "" Then
                                    ROWheader._162_Misc50 += ";" & Rmail.Correo2
                                End If
                            Next
                            If InStr(Mail, "@") Then ROWheader._162_Misc50 += ";" & Mail
                        Else
                            ta.Fill(t, Datos(1))
                            If t.Rows.Count > 0 Then
                                If InStr(Trim(t.Rows(0).Item("EMail1")), "@") Then ROWheader._162_Misc50 += ";" & Trim(t.Rows(0).Item("EMail1"))
                                If InStr(Trim(t.Rows(0).Item("EMail2")), "@") Then ROWheader._162_Misc50 += ";" & Trim(t.Rows(0).Item("EMail2"))
                                If InStr(Mail, "@") Then ROWheader._162_Misc50 += ";" & Mail
                            Else
                                If InStr(Mail, "@") Then ROWheader._162_Misc50 += ";" & Mail
                            End If
                        End If
                    Else
                        'CORREOS ADICIONALES++++++++++++++++++++++++++++
                        ROWheader._162_Misc50 += ";lhernandez@finagil.com.mx"
                    End If

                    If Datos(1).Trim = "05978 " Then
                        ROWheader._162_Misc50 += ";flen.estrada@ciasaconstruccion.com.mx;administacion@ciasaconstruccion.com.mx;"
                    End If
                    ProducDS.CFDI_Encabezado.AddCFDI_EncabezadoRow(ROWheader)
                End If
                f2.Close()
                If Errores = False Then
                    Try
                        Dim Total As Decimal = Math.Round(IVA + SubTT, 2)
                        Dim SaldoFactura, SaldoInsolFactura As Decimal
                        Dim NoPago As Integer = 0
                        SaldoFactura = CFDI_P.SaldoFactura(FolioORG, SerieORG)
                        'SaldoFactura = Total 'LINEA DE PRUEBAS
                        SaldoInsolFactura = SaldoFactura - Total
                        If SaldoInsolFactura < 0 Then
                            SaldoInsolFactura = 0
                        End If
                        ProducDS.CFDI_Encabezado.GetChanges()
                        ProducDS.CFDI_Detalle.GetChanges()
                        CFDI_D.Update(ProducDS.CFDI_Detalle)
                        CFDI_H.Update(ProducDS.CFDI_Encabezado)
                        NoPago = CFDI_H.NoPago(GUID)

                        'CFDI_P.Insert("CPG", "Pagos", "HD", "1.0", fecha.ToString("yyyy/MM/ddThh:mm:ss"), FormaPago, Moneda, TipoCambioSTR, Total, "REF_NO OPRACION", "RFC CEUNTA CLIENTE", "banco cliente", "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        If fecha_pago <> Nothing Then
                            CFDI_P.Insert("CPG", "Pagos", "HD", "1.0", fecha_pago.ToString("yyyy/MM/dd") + "T12:00:00", FormaPago, Moneda, TipoCambioSTR, Total, Referencia, RFC_BancoCliente, NombreBancoCliente, "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        Else
                            CFDI_P.Insert("CPG", "Pagos", "HD", "1.0", fecha.ToString("yyyy/MM/ddThh:mm:ss"), FormaPago, Moneda, TipoCambioSTR, Total, Referencia, RFC_BancoCliente, NombreBancoCliente, "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        End If
                        'CFDI_P.Insert("CPG", "Pago", "HD", "no cuenta cliete", "rfcBancoCeuntaFinagil", "cunetafinagil", "tipo de adena de pago 01 spei", "certificadopago", "cadenaorg", "sello", "", "", "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        CFDI_P.Insert("CPG", "Pago", "HD", NoCuentaCliente, RFC_BancoFinagil, CuentaFinagil, Spei, SpeiCert, SpeiCadOrg, SpeiSello, "", "", "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        'CFDI_P.Insert("CPG", "DoctoRelacionado", "HD", GUID, Serie, Folio, Moneda, TipoCambioSTR, "PPD", NoPago, SaldoFactura, Total, SaldoInsolFactura, "", "", "", "", "", Folio, Serie, Folio, Serie)
                        CFDI_P.Insert("CPG", "DoctoRelacionado", "HD", GUID, SerieORG, FolioORG, Moneda, "", "PPD", NoPago, SaldoFactura, Total, SaldoInsolFactura, "", "", "", "", "", Folio, Serie, Folio, Serie)
                            'CFDI_H.ConsumeFolio()

                            ProducDS.CFDI_Encabezado.Clear()
                        ProducDS.CFDI_Detalle.Clear()
                        File.Copy(F(i).FullName, GeneraFactura.My.Settings.RutaBackup & F(i).Name, True)
                        File.Delete(F(i).FullName)

                    Catch ex As Exception
                        EnviaCorreoFASE("Desarrollo", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "Error Factura TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo & " " & ex.Message)
                        ProducDS.CFDI_Encabezado.Clear()
                        ProducDS.CFDI_Detalle.Clear()
                    End Try
                Else
                    ProducDS.CFDI_Encabezado.Clear()
                    ProducDS.CFDI_Detalle.Clear()
                    File.Copy(F(i).FullName, GeneraFactura.My.Settings.Raiz & F(i).Name, True)
                    File.Delete(F(i).FullName)
                    NoFactError += 1
                End If
#End Region
            ElseIf EsPago = False And EsFactura = True And Serie <> "C" Then
                SE_PROCESARON_FACTURAS = True
#Region "Factura"
                f2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
                If Mid(F(i).Name, 1, 3) <> "FIN" And Mid(F(i).Name, 1, 3) <> "XXA" And IsNumeric(Mid(F(i).Name, 1, 4)) = True Then
                    fecha = New DateTime(Mid(F(i).Name, 1, 4), Mid(F(i).Name, 5, 2), Mid(F(i).Name, 7, 2), Mid(F(i).Name, 9, 2), Mid(F(i).Name, 11, 2), Mid(F(i).Name, 13, 2))
                    horas = DateDiff(DateInterval.Hour, fecha, Date.Now)
                    If horas >= 72 Then
                        fecha = fecha.AddHours(horas - 71)
                    End If
                End If
                ReDim Datos(1)
                Datos(0) = "X"
                While Not f2.EndOfStream
                    Linea = f2.ReadLine
                    If UCase(Linea) = "X" Then
                        EnviarGisela = True
                        Linea = f2.ReadLine
                    End If
                    Datos = Linea.Split("|")
                    If Datos.Length > 4 Then
                        cAnexoAux = Datos(2)
                        If Datos(2) = "03282/0002" Then Datos(2) = "2885803-001"
                        If Datos(2) = "01350/0012" Then Datos(2) = "10318141001"
                    End If

                    Select Case Datos(0)
                        Case "M1"
                            fecha = Datos(6)
                            Mail = Datos(5)
                        Case "H1"
                            fecha = Datos(1)
                            fecha = fecha.AddHours(Date.Now.Hour + 1)
                            fecha = fecha.AddMinutes(Date.Now.Minute)
                            fecha = fecha.AddSeconds(Date.Now.Second)
                            If DateDiff(DateInterval.Hour, fecha, Date.Now) > 72 Then
                                fecha = Date.Now.AddDays(-3)
                                fecha = fecha.AddHours(2)
                            Else
                                'pone la hora a la fecha del archivo
                                fecha = fecha.Date
                                fecha = fecha.AddHours(Date.Now.Hour)
                                fecha = fecha.AddMinutes(Date.Now.Minute)
                                fecha = fecha.AddSeconds(Date.Now.Second)
                            End If
                            Metodo_Pago = Datos(2)
                            FormaPago = Datos(3)
                            If SerieORG = "PUE" Then
                                Metodo_Pago = SerieORG
                            End If
                            If Datos.Length > 5 Then
                                fecha_pago = Datos(5)
                            End If
                        Case "H3"
                            If Datos(2).Length <> 10 Then
                                cAnexo = Mid(cAnexoAux, 1, 5) & Mid(cAnexoAux, 7, 4)
                            Else
                                cAnexo = Mid(Datos(2), 1, 5) & Mid(Datos(2), 7, 4)
                            End If

                            'If Datos.Length > 28 And Serie <> "F" Then
                            '    If Datos(29).Trim.Length = 2 Then
                            '        Aviso = CFDI_H.SacaAvisoAV(Datos(28), Datos(29)).Trim
                            '    Else
                            '        Aviso = CFDI_H.SacaAviso(Datos(28), Datos(29))
                            '    End If
                            '    If Aviso = "0" Then Aviso = ""
                            'Else
                            '    Aviso = ""
                            'End If

                            If Serie = "F" Then
                                Metodo_Pago = "PUE"
                                FormaPago = "17"
                            End If

                            If Metodo_Pago = "PPD" Then FormaPago = "99"
                            TipoPersona = taTipar.TipoPersona(Datos(1))
                            If IsNothing(TipoPersona) And Serie = "F" Then
                                TipoPersona = "M"
                            End If
                            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
                            TasaIVACliente = taCli.SacaTasaIvaAnexo(cAnexo)
                            ROWheader._1_Folio = Val(Datos(4))
                            ROWheader._2_Nombre_Emisor = "FINAGIL S.A. DE C.V, SOFOM E.N.R"
                            ROWheader._3_RFC_Emisor = "FIN940905AX7"
                            ROWheader._4_Dom_Emisor_calle = "Leandro Valle"
                            ROWheader._5_Dom_Emisor_noExterior = "402"
                            ROWheader._6_Dom_Emisor_noInterior = ""
                            ROWheader._7_Dom_Emisor_colonia = "Reforma y F.F.C.C"
                            ROWheader._8_Dom_Emisor_localidad = "Toluca"
                            ROWheader._9_Dom_Emisor_referencia = ""
                            ROWheader._10_Dom_Emisor_municipio = "Toluca"
                            ROWheader._11_Dom_Emisor_estado = "Estado de México"
                            ROWheader._12_Dom_Emisor_pais = "México"
                            ROWheader._13_Dom_Emisor_codigoPostal = "50070"

                            ROWheader._26_Version = "3.3"
                            ROWheader._27_Serie_Comprobante = Left(Serie, 8)
                            ROWheader._29_FormaPago = FormaPago '"27" '27 A satisfacción del acreedor
                            ROWheader._30_Fecha = fecha.Date
                            ROWheader.Fecha = fecha.Date
                            ROWheader._31_Hora = fecha.ToString("HH:mm:ss")
                            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

                            If Mid(Serie, 1, 1) = "C" Then
                                EsNotaCredito = True
                            Else
                                EsNotaCredito = False
                            End If

                            If Moneda = "WWW" Then
                                Moneda = Datos(17)
                            End If
                            If Moneda = "M.N." Then Moneda = "MXN"
                            If Moneda = "MXP" Then Moneda = "MXN"

                            If Mid(F(i).Name, 1, 3) = "XXA" Then
                                Serie = "DV"
                            End If

                            Datos(16) = ValidaRFC(Datos(16), TipoPersona)
                            If Trim(Datos(16)) = "SDA070613KU6" Then
                                Datos(5) = """SERVICIO DAYCO"" SA DE CV"
                            End If
                            If Trim(Datos(16)) = "CARD840606LEA" Then
                                Datos(5) = "DANIEL CADENA RUVALCABA"
                            End If
                            If Trim(Datos(16)) = "GET090828K63" Then
                                Errores = True
                                ErrorMSG = "!!No se Factura de Grupo empresarian transforma!!"
                            End If

                            ROWheader._42_Nombre_Receptor = Datos(5).Trim
                            ROWheader._43_RFC_Receptor = Datos(16).ToUpper
                            ROWheader._44_Dom_Receptor_calle = Datos(6).Trim
                            ROWheader._45_Dom_Receptor_noExterior = ""
                            ROWheader._46_Dom_Receptor_noInterior = ""
                            ROWheader._47_Dom_Receptor_colonia = Datos(9).Trim
                            ROWheader._48_Dom_Receptor_localidad = ""
                            ROWheader._49_Dom_Receptor_referencia = ""
                            ROWheader._50_Dom_Receptor_municipio = Datos(10).Trim
                            ROWheader._51_Dom_Receptor_estado = Datos(11).Trim
                            ROWheader._52_Dom_Receptor_pais = Datos(15).Trim
                            ROWheader._53_Dom_Receptor_codigoPostal = Datos(12).Trim
                            ROWheader._57_Estado = 1

                            ROWheader._83_Cod_Moneda = Moneda
                            ROWheader._97_Condiciones_Pago = "Contado"
                            UsoCFDI = taCodigo.SacaUsoCFDI(cAnexo)
                            If UsoCFDI.Trim = "" Then
                                UsoCFDI = taCodigo.SacaUsoCFDI(Tipar)
                            End If
                            ROWheader._144_Misc32 = UsoCFDI 'claves del SAT P01=por definir, G03=Gastos generales CATALOGO USO DE COMPROBANTE hoja excel= c_UsoCFDI, a solicitud del cliente
                            ROWheader._167_RegimentFiscal = 601
                            If ROWheader._83_Cod_Moneda = "MXN" Then
                                ROWheader._177_Tasa_Divisa = 0
                            Else
                                ROWheader._177_Tasa_Divisa = taCli.SacaTipoCambio(fecha.Date, ROWheader._83_Cod_Moneda)
                                If ROWheader._177_Tasa_Divisa = 1 Then
                                    Errores = True
                                    EnviaCorreoFASE("Contabilidad", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ErrorMSG & "Tipo de Cambio : 1 Concepto: " & Concepto & " TipoCredito : " & Tipar & " Anexo : " & cAnexo)
                                    EnviaCorreoFASE("Desarrollo", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ErrorMSG & "Tipo de Cambio : 1 Concepto: " & Concepto & " TipoCredito : " & Tipar & " Anexo : " & cAnexo)
                                End If
                            End If
                            ROWheader._180_LugarExpedicion = "50070"
                            ROWheader._190_Metodo_Pago = Metodo_Pago 'PPD pago en parcialidades PUE pago en una sola exhibision

                            If Serie = "C" Then
                                ROWheader._191_Efecto_Comprobante = "E"
                                ROWheader._58_TipoCFD = "NC"
                            Else
                                ROWheader._191_Efecto_Comprobante = "I"
                                ROWheader._58_TipoCFD = "FA"
                            End If
                        Case "D1"
                            Dim TasaIVA As Decimal = 0.16
                            Dim TipoImpuesto As String
                            If InStr(UCase(Datos(8)), "IVA") <= 0 Then
                                Datos(8) = Trim(Datos(8))
                                Concepto = LimpiarConcepto(Datos(8), ROWheader._27_Serie_Comprobante)
                                taCodigo.Fill(tCodigo, Tipar, Concepto)
                                If tCodigo.Rows.Count > 0 Then
                                    rCod = tCodigo.Rows(0)
                                    If rCod.Adenda = True Then
                                        If Tipar <> "F" And Tipar <> "P" Then
                                            Select Case Concepto
                                                Case "ADELANTO CAPITAL EQUIPO"
                                                    Concepto = "ADELANTO CAPITAL"
                                                Case "SALDO INSOLUTO EQUIPO"
                                                    Concepto = "SALDO INSOLUTO"
                                                Case "SALDO INSOLUTO DEL EQUIPO"
                                                    Concepto = "SALDO INSOLUTO"
                                                Case "CAPITAL EQUIPO"
                                                    Concepto = "CAPITAL"
                                                Case "PAGO CREDITO DE AVIO"
                                                    Concepto = "CREDITO DE AVIO"
                                            End Select
                                            If InStr(Concepto, "CAPITAL EQUIPO VEN") > 0 Then
                                                Concepto = "CAPITAL VENCIMIENTO" '& Right(Concepto, 7)
                                            End If
                                        End If
                                        LeyendaCapital += "* " & Concepto & " " & CDec(Datos(10)).ToString("n2") & vbCrLf
                                        Continue While
                                    End If
                                End If
                                ROWdetail = ProducDS.CFDI_Detalle.NewCFDI_DetalleRow
                                If ROWheader._27_Serie_Comprobante = "B" Then
                                    Dim TasaIvaCapital As Decimal
                                    Select Case IVACapital
                                        Case "0%", "EXE"
                                            TasaIvaCapital = 0
                                        Case "8%"
                                            TasaIvaCapital = 8
                                        Case "16%"
                                            TasaIvaCapital = 16
                                        Case Else
                                            TasaIvaCapital = TasaIVACliente
                                    End Select
                                    TasaIVA = TasaIvaCapital / 100
                                    TipoImpuesto = TasaIvaCapital
                                Else
                                    TasaIVA = TasaIVACliente / 100
                                    TipoImpuesto = TasaIVACliente
                                End If


                                NoLineas += 1

                                If InStr(Datos(8), "Comisión del") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "Comision del") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "COMISION POR APERTURA") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "COMISION DE APERTURA") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "GASTOS DE RATIFICACION") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "COMISION FEGA") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "GASTOS NOTARIALES") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = Math.Round(Math.Round(CDec(Datos(10)), 2) * TasaIVA, 2)
                                End If

                                If InStr(Datos(8), "Int. Ord.") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = 0
                                    TipoImpuesto = "Exento"
                                End If
                                If InStr(Datos(8), "INTERES MENSUAL") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = 0
                                    TipoImpuesto = "Exento"
                                End If
                                If (Datos(8) = "INTERESES AL VENCIMIENTO" Or Datos(8) = "INTERESES MORATORIOS" Or Datos(8) = "INTERES MORATORIO") And ROWheader._27_Serie_Comprobante = "F" Then
                                    ReDim Preserve Datos(11)
                                    Datos(11) = 0
                                    TipoImpuesto = "Exento"
                                    ROWheader._29_FormaPago = "03"
                                End If


                                'If Tipar <> "F" And Tipar <> "P" Then 'se puede borrar
                                '    Select Case Datos(8)
                                '        Case "ADELANTO CAPITAL EQUIPO"
                                '            Datos(8) = "ADELANTO CAPITAL"
                                '        Case "SALDO INSOLUTO EQUIPO"
                                '            Datos(8) = "SALDO INSOLUTO"
                                '        Case "SALDO INSOLUTO DEL EQUIPO"
                                '            Datos(8) = "SALDO INSOLUTO"
                                '        Case "CAPITAL EQUIPO"
                                '            Datos(8) = "CAPITAL"
                                '    End Select
                                '    If InStr(Datos(8), "CAPITAL EQUIPO VEN") > 0 Then
                                '        Datos(8) = "CAPITAL VENCIMIENTO" '& Right(Datos(8), 7)
                                '    End If
                                'End If

                                If (Tipar = "F") And TipoPersona <> "F" Then
                                    Select Case Concepto.Trim
                                        Case "INTERES OTROS ADEUDOS", "INTERES SEGURO", "INTERESES POR PREPAGO SEGURO", "INTERES SEGURO VENCIMIENTO", "INTERES OTROS ADEUDOS VENCIMIENTO"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If (Tipar = "R" Or Tipar = "S") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "MORATORIOS V"
                                            TipoImpuesto = "Exento"
                                        Case "INTERESES VE", "INTERES OTRO", "INTERES SEGU", "INTERESES PO"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If (Tipar = "R" Or Tipar = "S" Or Tipar = "F") And TipoPersona = "F" And CDec(Datos(11)) = 0 Then ' para fisicas sin iva por inflacion
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "MORATORIOS V"
                                            TipoImpuesto = "No Objeto"
                                        Case "INTERESES VE", "INTERES OTRO", "INTERES SEGU", "INTERESES PO"
                                            TipoImpuesto = "No Objeto"
                                    End Select
                                    If Datos(8) = "INTERESES" Then
                                        TipoImpuesto = "No Objeto"
                                    End If
                                End If
                                If Tipar = "F" And cAnexo = "038240001" Then
                                    TipoImpuesto = "Exento"
                                ElseIf Tipar = "F" And cAnexo = "025620003" Then '#ECT Solicitado por Valentin 24/09/2015
                                    Select Case Datos(8)
                                        Case "ADELANTO CAPITAL EQUIPO"
                                            Datos(8) = Datos(8) '& " A TASA IVA 0%"
                                            TasaIVA = 0
                                        Case "CAPITAL EQUIPO"
                                            Datos(8) = Datos(8) '& " A TASA IVA 0%"
                                            TasaIVA = 0
                                        Case "AMORTIZACION INICIAL"
                                            Datos(8) = Datos(8) '& " A TASA IVA 0%"
                                            TasaIVA = 0
                                    End Select
                                    If Mid(Datos(8), 1, 9) = "INTERESES" Then
                                        TipoImpuesto = "Exento"
                                    End If
                                ElseIf Tipar = "F" And InStr(Datos(8), "CAPITAL") Then
                                    Select Case IVACapital
                                        Case "NOO"
                                            TipoImpuesto = "No Objeto"
                                        Case "EXE"
                                            TipoImpuesto = "Exento"
                                        Case "16%"
                                        Case ""
                                        Case "0%"
                                            TasaIVA = 0
                                    End Select
                                End If


                                If (Tipar = "A" Or Tipar = "H" Or Tipar = "C") Then
                                    If Tipar = "C" Then
                                        Select Case Mid(Datos(8), 1, 12)
                                            Case "INTERESES AV"
                                                Datos(8) = "INTERESES CUENTA CORRIENTE"
                                            Case "INTERESES MO"
                                                Datos(8) = "INTERESES MORATORIO CUENTA CORRIENTE"
                                            Case "PAGO CREDITO"
                                                Datos(8) = "PAGO CREDITO EN CUENTA CORRIENTE"
                                        End Select
                                    End If

                                    If TipoPersona <> "F" And Mid(Datos(8), 1, 9) = "INTERESES" Then
                                        TipoImpuesto = "Exento"
                                    End If
                                End If

                                If InStr(Datos(8), "SEGURO DE VI") > 0 Then
                                    TipoImpuesto = "Exento"
                                End If

                                If Tipar = "P" Then
                                    Select Case Datos(8)
                                        Case "AMORTIZACION INICIAL"
                                            Datos(8) = "RENTA INICIAL"
                                        Case "INTERESES POR PREPAGO"
                                            Datos(8) = "PAGO DE RENTA VENCIMIENTO"
                                    End Select
                                    If CDec(Datos(11)) = 0 Then
                                        TipoImpuesto = "Exento"
                                    End If

                                    If Concepto = "PAGO DE RENTA VENCIMIENTO" Then
                                        Datos(8) = Mid(Datos(8), 9, Datos(8).Length)
                                    End If
                                End If

                                If Tipar = "B" Then
                                        Select Case Mid(Datos(8), 1, 11)
                                            Case "MENSUALIDAD"
                                                Datos(8) = "SERVICIO DE TRANSPORTE EJECUTIVO EMPRESARIAL, " & Datos(8)
                                            Case "MORATORIOS "
                                                TipoImpuesto = "Exento"
                                        End Select
                                    End If

                                    If Datos(8) = "AJUSTE INTERES" Then
                                        TipoImpuesto = "Exento"
                                    End If

                                    taCodigo.Fill(tCodigo, Tipar, Concepto)
                                    If tCodigo.Rows.Count > 0 Then
                                        rCod = tCodigo.Rows(0)
                                        Unidad = rCod.Unidad
                                        Codigo = rCod.Codigo
                                        If Codigo = "" Then
                                            If Tipar = "P" Then
                                                Codigo = taCodigo.SacaCodigoAnexo(cAnexo)
                                            End If
                                            If Codigo = "" Then
                                                Codigo = "84101700"
                                                Errores = True
                                                ErrorMSG = "Falta codigo "
                                            End If
                                        End If
                                        If Unidad = "" Then
                                            If Tipar = "P" Then
                                                Unidad = "E48"
                                            Else
                                                Unidad = "E48"
                                                Errores = True
                                                ErrorMSG = "Falta Unidad "
                                            End If

                                        End If
                                        If Errores = True And (Tipar = "F" Or Tipar = "S") And Concepto = "CAPITAL EQUIPO VENCIMIENTO" Then
                                            'Errores = False 'SE QUITA CUANDO ESTE CONFIGURADOS LOS ARTICULOS Y UNIDADES
                                        End If
                                        If Errores = True And Tipar = "P" And Concepto = "PAGO DE RENTA VENCIMIENTO" Then
                                            'Errores = False 'SE QUITA CUANDO ESTE CONFIGURADOS LOS ARTICULOS Y UNIDADES
                                        End If

                                    Else
                                        If Tipar = "P" Then
                                            Unidad = "E48"
                                            If Codigo = "" Then
                                                If taCodigo.ExisteConcepto(Tipar, Concepto) <= 0 And ROWheader._27_Serie_Comprobante <> "B" Then
                                                    taCodigo.Insert(Tipar, Concepto, "", "", False)
                                                End If
                                                Errores = True
                                                ErrorMSG = "Falta Concepto "
                                            End If
                                        Else

                                        End If
                                        If taCodigo.ExisteConcepto(Tipar, Concepto) <= 0 And ROWheader._27_Serie_Comprobante <> "B" Then
                                            taCodigo.Insert(Tipar, Concepto, "", "", False)
                                        End If
                                        Unidad = "E48"
                                        Codigo = "84101700"
                                        Errores = True
                                        ErrorMSG = "Falta codigo "
                                        If ROWheader._27_Serie_Comprobante = "B" Then
                                            Errores = False 'quitamos el error
                                            Codigo = Datos(12)
                                            ROWheader._144_Misc32 = Datos(13)
                                            Unidad = Datos(14)
                                        End If
                                        If Serie = "F" Then
                                            Errores = True ' quitamos el error de factoraje
                                        End If

                                    End If
                                If Errores = True Then
                                    EnviaCorreoFASE("Contabilidad", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ErrorMSG & "Concepto: " & Concepto & " TipoCredito : " & Tipar & " Anexo : " & cAnexo)
                                    EnviaCorreoFASE("Desarrollo", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ErrorMSG & "Concepto: " & Concepto & " TipoCredito : " & Tipar & " Anexo : " & cAnexo)
                                End If

                                ROWdetail._1_Linea_Descripcion = Datos(8).Trim
                                    ROWdetail._2_Linea_Cantidad = 1
                                    ROWdetail._3_Linea_Unidad = Unidad
                                    ROWdetail._4_Linea_PrecioUnitario = CDec(Datos(10)).ToString("n2")
                                    ROWdetail._5_Linea_Importe = CDec(Datos(10)).ToString("n2")
                                    ROWdetail._16_Linea_Cod_Articulo = Codigo ' Manejo de deuda
                                    ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                                    ROWdetail._2_Impuesto_Descripcion = "TR"
                                    ROWdetail._3_Impuesto_Monto_base = CDec(Datos(10)).ToString("n2")
                                    ROWdetail._5_Impuesto_Clave = "002"
                                    ROWdetail._6_Impuesto_Tasa = "Tasa"
                                If Datos(6).Trim = "" Or Serie = "F" Then
                                    Datos(6) = "SER"
                                End If
                                ROWdetail._11_Linea_Notas = Datos(6)
                                ROWdetail._53_Linea_Misc22 = Datos(6)
                                    Try
                                        If TipoImpuesto = "Exento" Then
                                            ROWdetail._7_Impuesto_Porcentaje = ""
                                            ROWdetail._4_Impuesto_Monto_Impuesto = ""
                                            ROWdetail._6_Impuesto_Tasa = TipoImpuesto
                                        ElseIf TipoImpuesto = "No Objeto" Then
                                            ROWdetail._7_Impuesto_Porcentaje = ""
                                            ROWdetail._4_Impuesto_Monto_Impuesto = ""
                                            ROWdetail._6_Impuesto_Tasa = TipoImpuesto
                                        Else
                                            ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                                            If TasaIVA = 0 Or CDec(Datos(11)) = 0 Then
                                                ROWdetail._4_Impuesto_Monto_Impuesto = 0
                                            Else
                                                ROWdetail._4_Impuesto_Monto_Impuesto = Math.Round(CDec(Datos(11)), 2)
                                                MontoBaseIVA = CDec(Datos(11)) / TasaIVA
                                                If MontoBaseIVA < ROWdetail._3_Impuesto_Monto_base And (Mid(Datos(8), 1, 7) = "INTERES" Or InStr(Datos(8), "MORATORIOS VENCIMIENTO")) Then
                                                    ROWdetail._3_Impuesto_Monto_base = Math.Round(MontoBaseIVA, 2)
                                                End If
                                                IvaAux = Math.Round(ROWdetail._4_Impuesto_Monto_Impuesto / ROWdetail._3_Impuesto_Monto_base, 3)
                                                If IvaAux > TasaIVA Then
                                                    ROWdetail._4_Impuesto_Monto_Impuesto = Math.Round(ROWdetail._3_Impuesto_Monto_base * TasaIVA, 2)
                                                End If
                                            End If
                                        End If

                                        SubTT += ROWdetail._5_Linea_Importe
                                        If IsNumeric(ROWdetail._4_Impuesto_Monto_Impuesto) Then
                                            IVA += CDec(ROWdetail._4_Impuesto_Monto_Impuesto).ToString("n2")
                                        End If

                                        ROWdetail.Detalle_Folio = ROWheader._1_Folio
                                        ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante

                                        ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                                    Catch ex As Exception
                                        Errores = True
                                    EnviaCorreoFASE("Desarrollo", "Error de Factura " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ex.Message & " " & ErrorMSG)
                                End Try

                                End If
                    End Select
                End While
                If Datos(0) <> "X" Then
                    ROWheader._90_Cantidad_LineasFactura = NoLineas
                    ROWheader._54_Monto_SubTotal = SubTT
                    ROWheader._55_Monto_IVA = IVA
                    ROWheader._56_Monto_Total = ROWheader._54_Monto_SubTotal + ROWheader._55_Monto_IVA
                    ROWheader._193_Monto_TotalImp_Trasladados = ROWheader._55_Monto_IVA
                    ROWheader._100_Letras_Monto_Total = Letras(ROWheader._56_Monto_Total, Moneda)
                    ROWheader._114_Misc02 = Datos(2)
                    ROWheader._115_Misc03 = Datos(1)
                    ROWheader._158_Misc46 = TipoCredito.Trim
                    ROWheader._159_Misc47 = Aviso
                    ROWheader._162_Misc50 = ""
                    'ROWheader._161_Misc49 = ""
                    If OpcionCompraAF.Trim.Length > 0 Then
                        ROWheader._157_Misc45 = LeyendaCapital.Trim & " * " & OpcionCompraAF.Trim
                    Else
                        ROWheader._157_Misc45 = LeyendaCapital.Trim
                    End If


                    ROWheader.Encabezado_Procesado = Errores

                    If EnviarGisela = False Then
                        'CORREOS ADICIONALES++++++++++++++++++++++++++++
                        taMail.Fill(tMail, cAnexo)
                        If tMail.Rows.Count > 0 Then
                            For Each Rmail In tMail.Rows
                                If Rmail.Correo1 > "" Then
                                    ROWheader._162_Misc50 += ";" & Rmail.Correo1
                                End If
                                If Rmail.Correo2 > "" Then
                                    ROWheader._162_Misc50 += ";" & Rmail.Correo2
                                End If
                            Next
                            If InStr(Mail, "@") Then ROWheader._162_Misc50 += ";" & Mail
                        Else
                            ta.Fill(t, Datos(1))
                            If t.Rows.Count > 0 Then
                                If InStr(Trim(t.Rows(0).Item("EMail1")), "@") Then ROWheader._162_Misc50 += ";" & Trim(t.Rows(0).Item("EMail1"))
                                If InStr(Trim(t.Rows(0).Item("EMail2")), "@") Then ROWheader._162_Misc50 += ";" & Trim(t.Rows(0).Item("EMail2"))
                                If InStr(Mail, "@") Then ROWheader._162_Misc50 += ";" & Mail
                            Else
                                If InStr(Mail, "@") Then ROWheader._162_Misc50 += ";" & Mail
                            End If
                        End If
                    Else
                        'CORREOS ADICIONALES++++++++++++++++++++++++++++
                        ROWheader._162_Misc50 += ";lhernandez@finagil.com.mx"
                    End If

                    If Datos(1).Trim = "05978 " Then
                        ROWheader._162_Misc50 += ";flen.estrada@ciasaconstruccion.com.mx;administacion@ciasaconstruccion.com.mx;"
                    End If
                    ProducDS.CFDI_Encabezado.AddCFDI_EncabezadoRow(ROWheader)
                End If
                f2.Close()
                If Errores = False Then
                    Try
                        ProducDS.CFDI_Encabezado.GetChanges()
                        ProducDS.CFDI_Detalle.GetChanges()
                        CFDI_D.Update(ProducDS.CFDI_Detalle)
                        CFDI_H.Update(ProducDS.CFDI_Encabezado)


                        ProducDS.CFDI_Encabezado.Clear()
                        ProducDS.CFDI_Detalle.Clear()
                        File.Copy(F(i).FullName, GeneraFactura.My.Settings.RutaBackup & F(i).Name, True)
                        File.Delete(F(i).FullName)

                    Catch ex As Exception
                        EnviaCorreoFASE("desarrollo", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, ex.Message & ErrorMSG & " Error Factura TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo)
                        ProducDS.CFDI_Encabezado.Clear()
                        ProducDS.CFDI_Detalle.Clear()
                    End Try
                Else
                    ProducDS.CFDI_Encabezado.Clear()
                    ProducDS.CFDI_Detalle.Clear()
                    File.Copy(F(i).FullName, GeneraFactura.My.Settings.NoPasa & F(i).Name, True)
                    File.Delete(F(i).FullName)
                    NoFactError += 1
                End If
#End Region
            Else
                If EsFactura = True And EsPago = True Then
                    File.Copy(F(i).FullName, GeneraFactura.My.Settings.Complementos & F(i).Name, True)
                    File.Delete(F(i).FullName)
                End If
                If EsFactura = False Or Serie = "C" Then
                    'File.Copy(F(i).FullName, GeneraFactura.My.Settings.NoPasa & F(i).Name, True)
                    'File.Delete(F(i).FullName)
                End If
            End If
        Next
        If SinFolio.Length > 0 Then
            EnviaCorreoFASE("Desarrollo", "Factura sin Procesar " & Serie & Folio, SinFolio)
        End If
        If NoFactError > 0 Then
            EnviaCorreoFASE("desarrollo", "Error  de Facturas sin procesar:  " & NoFactError, "Error  de Facturas sin procesar:  " & NoFactError)
        End If
    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, "No se ha encontrado la ruta de acceso de la red") = 0 Then
            Dim Mensage As New MailMessage("InternoBI2008@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient(My.Settings.SMTP, My.Settings.SMTP_port)
            Try
                Dim Credenciales As String() = My.Settings.SMTP_creden.Split(",")
                Cliente.Credentials = New System.Net.NetworkCredential(Credenciales(0), Credenciales(1), Credenciales(2))
                Cliente.Send(Mensage)
            Catch ex As Exception
                ReportError(ex)
            End Try
        Else
            Console.WriteLine("No se ha encontrado la ruta de acceso de la red")
        End If
    End Sub

    Private Sub ReportError(ByVal ex As Exception)
        ErrorControl.WriteEntry(ex.Message, EventLogEntryType.Error)
    End Sub

    Sub GeneraArchivosEXternas()
        Dim CFDI_H As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim CFDI_D As New ProduccionDSTableAdapters.CFDI_DetalleTableAdapter
        Dim taImprAdic As New ProduccionDSTableAdapters.CFDI_Impuestos_AdicionalesTableAdapter
        Dim ROWheader As ProduccionDS.CFDI_EncabezadoRow
        Dim ROWdetail As ProduccionDS.CFDI_DetalleRow
        Dim ProducDS As New ProduccionDS
        Dim TasaIVACliente, MontoBaseIVA As Decimal
        Dim NoLineas As Integer
        Dim RFC As String = ""
        Dim Razon As String = ""
        Dim SubTT As Double
        Dim RetencionT As Double
        Dim IVA As Double
        Dim TOt As Double
        Dim Fec As Date
        Dim Errores As Boolean
        Dim Facturas As New GeneraFactura.ProduccionDSTableAdapters.FacturasExternasTableAdapter
        Dim FAC As New GeneraFactura.ProduccionDS.FacturasExternasDataTable
        Dim Detalles As New GeneraFactura.ProduccionDSTableAdapters.FacturasExternasDETTableAdapter
        Dim DET As New GeneraFactura.ProduccionDS.FacturasExternasDETDataTable
        Dim TipoPersona As String = "F"
        'Dim MetodoPago As String = ""
        'Dim taMetodo As New ProduccionDSTableAdapters.LlavesTableAdapter


        Facturas.Fill(FAC)
        For Each r As GeneraFactura.ProduccionDS.FacturasExternasRow In FAC.Rows
            Fec = r.fecha
            Fec = Fec.AddHours(Today.Now.AddHours(1).Hour)
            Fec = Fec.AddMinutes(Today.Now.Minute)
            Console.WriteLine("Generando CFDI Facturas Externas..." & r.Factura)
            SubTT = 0
            TOt = 0
            IVA = 0
            RetencionT = 0
            Errores = False
            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
            ROWheader._1_Folio = Val(r.Factura)
            If r.Finagil = True Then
                ROWheader._2_Nombre_Emisor = "FINAGIL S.A. DE C.V, SOFOM E.N.R"
                ROWheader._3_RFC_Emisor = "FIN940905AX7"
                ROWheader._6_Dom_Emisor_noInterior = ""
            Else
                ROWheader._2_Nombre_Emisor = "SERVICIOS ARFIN S.A. DE C.V."
                ROWheader._3_RFC_Emisor = "SAR951230N5A"
                ROWheader._6_Dom_Emisor_noInterior = "2do PISO"
            End If

            ROWheader._4_Dom_Emisor_calle = "Leandro Valle"
            ROWheader._5_Dom_Emisor_noExterior = "402"
            ROWheader._7_Dom_Emisor_colonia = "Reforma y F.F.C.C"
            ROWheader._8_Dom_Emisor_localidad = "Toluca"
            ROWheader._9_Dom_Emisor_referencia = ""
            ROWheader._10_Dom_Emisor_municipio = "Toluca"
            ROWheader._11_Dom_Emisor_estado = "Estado de México"
            ROWheader._12_Dom_Emisor_pais = "México"
            ROWheader._13_Dom_Emisor_codigoPostal = "50070"

            ROWheader._26_Version = "3.3"
            ROWheader._27_Serie_Comprobante = r.Serie.Trim
            ROWheader._29_FormaPago = r.MetodoPago.Trim '"27" '27 A satisfacción del acreedor
            ROWheader._30_Fecha = Fec.Date
            ROWheader.Fecha = Fec.Date
            ROWheader._31_Hora = Fec.ToString("HH:mm:ss")
            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

            If IsNumeric(Mid(r.RFC, 4, 1)) Then
                TipoPersona = "M"
            Else
                TipoPersona = "F"
            End If
            RFC = ValidaRFC(r.RFC, TipoPersona)

            If RFC = "SDA070613KU6" Then
                Razon = """SERVICIO DAYCO"" SA DE CV"
            Else
                Razon = r.Nombre
            End If

            ROWheader._42_Nombre_Receptor = Razon.Trim
            ROWheader._43_RFC_Receptor = RFC.Trim
            ROWheader._44_Dom_Receptor_calle = r.Calle.Trim
            ROWheader._45_Dom_Receptor_noExterior = ""
            ROWheader._46_Dom_Receptor_noInterior = ""
            ROWheader._47_Dom_Receptor_colonia = r.Colonia.Trim
            ROWheader._48_Dom_Receptor_localidad = ""
            ROWheader._49_Dom_Receptor_referencia = ""
            ROWheader._50_Dom_Receptor_municipio = r.Municipio.Trim
            ROWheader._51_Dom_Receptor_estado = r.Estado.Trim
            ROWheader._52_Dom_Receptor_pais = "México"
            ROWheader._53_Dom_Receptor_codigoPostal = r.CP
            ROWheader._57_Estado = 1

            ROWheader._83_Cod_Moneda = r.Moneda
            ROWheader._97_Condiciones_Pago = "Contado"
            If Not IsNothing(r.FormatoImp) Then
                ROWheader._113_Misc01 = r.FormatoImp
            Else
                ROWheader._113_Misc01 = ""
            End If
            ROWheader._144_Misc32 = r.UsoCFDI '"G03" 'claves del SAT P01=por definir, G03=Gastos generales CATALOGO USO DE COMPROBANTE hoja excel= c_UsoCFDI, a solicitud del cliente
            ROWheader._167_RegimentFiscal = 601
            If ROWheader._83_Cod_Moneda = "MXN" Then
                ROWheader._177_Tasa_Divisa = 0
            Else
                ROWheader._177_Tasa_Divisa = Facturas.SacaTipoCambio(Today.Date, ROWheader._83_Cod_Moneda)
            End If
            ROWheader._180_LugarExpedicion = "50070"
            ROWheader._190_Metodo_Pago = r.MetodoPagoSAT 'PPD pago en parcialidades PUE pago en una sola exhibision
            If ROWheader._27_Serie_Comprobante = "C" Then
                ROWheader._191_Efecto_Comprobante = "E"
            Else
                ROWheader._191_Efecto_Comprobante = "I"
            End If
            ROWheader._58_TipoCFD = "FA"

            Dim var_TotalTrasladosExcentos As String = "NA"
            Dim contadorTraslados As Integer = 0

            Detalles.Fill(DET, r.Serie, r.Factura)
            For Each rr As GeneraFactura.ProduccionDS.FacturasExternasDETRow In DET.Rows
                NoLineas += 1
                ROWdetail = ProducDS.CFDI_Detalle.NewCFDI_DetalleRow
                ROWdetail._1_Linea_Descripcion = rr.Detalle
                ROWdetail._2_Linea_Cantidad = rr.Cantidad
                ROWdetail._3_Linea_Unidad = rr.Unidad
                ROWdetail._4_Linea_PrecioUnitario = Math.Round(rr.Unitario, 4)
                ROWdetail._5_Linea_Importe = Math.Round(rr.Importe, 4)
                ROWdetail._16_Linea_Cod_Articulo = rr.CodigoART ' Manejo de deuda
                ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                ROWdetail._2_Impuesto_Descripcion = "TR"
                ROWdetail._3_Impuesto_Monto_base = Math.Round(rr.Importe, 4)
                ROWdetail._5_Impuesto_Clave = "002"
                ROWdetail._6_Impuesto_Tasa = "Tasa"
                ROWdetail._11_Linea_Notas = rr.UnidadInterna
                ROWdetail._53_Linea_Misc22 = rr.UnidadInterna

                If rr.TasaIva = "Exento" Then
                    ROWdetail._7_Impuesto_Porcentaje = ""
                    ROWdetail._4_Impuesto_Monto_Impuesto = ""
                    ROWdetail._6_Impuesto_Tasa = "Exento"
                    var_TotalTrasladosExcentos = "SI"
                ElseIf rr.TasaIva = "No Objeto" Then
                    ROWdetail._7_Impuesto_Porcentaje = ""
                    ROWdetail._4_Impuesto_Monto_Impuesto = ""
                    ROWdetail._6_Impuesto_Tasa = rr.TasaIva
                Else
                    contadorTraslados += 1
                    TasaIVACliente = Val(rr.TasaIva.Substring(0, 2)) / 100
                    ROWdetail._7_Impuesto_Porcentaje = TasaIVACliente
                    If TasaIVACliente = 0 Then
                        ROWdetail._4_Impuesto_Monto_Impuesto = 0
                    Else
                        ROWdetail._4_Impuesto_Monto_Impuesto = Math.Round(rr.Iva, 4)
                        MontoBaseIVA = CDec(rr.Iva / TasaIVACliente)
                        If MontoBaseIVA < ROWdetail._3_Impuesto_Monto_base Then
                            ROWdetail._3_Impuesto_Monto_base = MontoBaseIVA
                        End If
                        ROWdetail._8_Retencion_Tasa = rr.RetencionTasa
                        ROWdetail._9_Retencion_Monto_Base = rr.RetencionBase
                        ROWdetail._10_Retencion_Monto = rr.RetencionMonto
                        RetencionT += rr.RetencionMonto
                    End If
                End If

                SubTT += ROWdetail._5_Linea_Importe
                If IsNumeric(ROWdetail._4_Impuesto_Monto_Impuesto) Then
                    IVA += CDec(ROWdetail._4_Impuesto_Monto_Impuesto)
                End If

                ROWdetail.Detalle_Folio = ROWheader._1_Folio
                ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante
                Try
                    ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                Catch ex As Exception
                    Errores = True
                    EnviaCorreoFASE("SISTEMAS_CFDI", "Error Factura TipoCredito : EXTERNA", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante)
                End Try
                Facturas.Facturar(r.Serie, r.Factura, rr.Consec)
            Next

            ROWheader._192_Monto_TotalImp_Retenidos = RetencionT
            ROWheader._90_Cantidad_LineasFactura = NoLineas
            ROWheader._54_Monto_SubTotal = SubTT
            ROWheader._55_Monto_IVA = IVA
            ROWheader._193_Monto_TotalImp_Trasladados = IVA
            If taImprAdic.Obt_Iporte_Ret_ScalarQuery(r.Serie, r.Factura) > 0 Then
                ROWheader._192_Monto_TotalImp_Retenidos += taImprAdic.Obt_Iporte_Ret_ScalarQuery(r.Serie, r.Factura) ' Acumula Retencion de ISR + IVA
            End If
            ROWheader._56_Monto_Total = ROWheader._54_Monto_SubTotal + ROWheader._55_Monto_IVA - ROWheader._192_Monto_TotalImp_Retenidos
            ROWheader._100_Letras_Monto_Total = Letras(ROWheader._56_Monto_Total, r.Moneda)
            ROWheader._114_Misc02 = "" ' contrato
            ROWheader._115_Misc03 = "" ' Cliente
            ROWheader._162_Misc50 = ""
            ROWheader._157_Misc45 = "" ' Adenda
            ROWheader._161_Misc49 = ""
            ROWheader.Encabezado_Procesado = False
            ROWheader._162_Misc50 = r.Mail1.Trim

            If var_TotalTrasladosExcentos = "SI" And contadorTraslados = 0 Then
                ROWheader._193_Monto_TotalImp_Trasladados = -1
            End If

            If r.Mail2.Trim.Length > 3 Then ROWheader._162_Misc50 += ";" & r.Mail2.Trim
            ROWheader._162_Misc50 += ";lhernandez@finagil.com.mx"
            ProducDS.CFDI_Encabezado.AddCFDI_EncabezadoRow(ROWheader)

            If Errores = False Then
                Try
                    ProducDS.CFDI_Encabezado.GetChanges()
                    ProducDS.CFDI_Detalle.GetChanges()
                    CFDI_D.Update(ProducDS.CFDI_Detalle)
                    CFDI_H.Update(ProducDS.CFDI_Encabezado)
                    ProducDS.CFDI_Encabezado.Clear()
                    ProducDS.CFDI_Detalle.Clear()

                Catch ex As Exception
                    EnviaCorreoFASE("SISTEMAS_CFDI", "Error Factura TipoCredito : EXTERNA", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante)
                    ProducDS.CFDI_Encabezado.Clear()
                    ProducDS.CFDI_Detalle.Clear()
                End Try
            Else
                ProducDS.CFDI_Encabezado.Clear()
                ProducDS.CFDI_Detalle.Clear()
            End If
        Next

    End Sub

    Function LecturaPrevia(RutaArchivo As String, NombreArchivo As String, ByRef Moneda As String, ByRef Tipar As String, ByRef Folio As Integer, ByRef Serie As String,
                           ByRef EsFactura As Boolean, ByRef EsPAgo As Boolean, ByRef SerieORG As String, ByRef FolioORG As Integer, ByRef GUID As String,
                           ByRef Referencia As String, Optional ByRef aviso As String = "", Optional ByRef IVACapital As String = "") As Boolean
        OpcionCompraAF = ""
        Dim Numero As Integer = 1
        Dim f2 As System.IO.StreamReader
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim Linea As String
        Dim Datos() As String
        Dim Anexo As String = ""
        Dim FechaAviso As Date = "01/01/1900"

        f2 = New System.IO.StreamReader(RutaArchivo, Text.Encoding.GetEncoding(1252))
        Try
            While Not f2.EndOfStream
                Linea = f2.ReadLine
                Datos = Linea.Split("|")

                Select Case Datos(0)
                    Case "H1"
                        If Datos.Length >= 5 Then
                            Referencia = Datos(4)
                        Else
                            Referencia = ""
                        End If
                    Case "H3"
                        Anexo = Mid(Datos(2), 1, 5) & Mid(Datos(2), 7, 4)
                        Folio = Val(Datos(4))
                        Serie = Datos(3)
                        If Serie = "REP" Or Serie = "REPP" Then
                            EsPAgo = True
                        Else
                            EsPAgo = False
                        End If
                        If Serie = "F" Then
                            Moneda = Datos(17)
                            Tipar = "X"
                            TipoCredito = "FACTORAJE"
                            EsFactura = True
                        Else
                            Moneda = taTipar.SacaMoneda(Datos(2))
                            Tipar = taTipar.TipaR(Datos(2))
                            TipoCredito = taTipar.TipoCredito(Tipar)
                            IVACapital = taTipar.SacaIvaCapital(Anexo)
                        End If

                        If Datos.Length > 28 And Serie <> "F" Then
                            If Datos(29).Trim.Length = 2 Then
                                aviso = CFDI_H.SacaAvisoAV(Datos(28), Datos(29)).Trim
                            Else
                                aviso = CFDI_H.SacaAviso(Datos(28), Datos(29))
                            End If
                            If aviso = "0" Then
                                aviso = ""
                            Else
                                If EsPAgo = True Then
                                    SerieORG = CFDI_H.SacaSerieORG(aviso, Datos(2))
                                    FolioORG = CFDI_H.SacaFolioORG(aviso, Datos(2))
                                    FechaAviso = CFDI_H.SacaAvisoFecha(Datos(28), Datos(29))
                                    If SerieORG = "XX" Then
                                        EsPAgo = False
                                        EsFactura = False
                                    End If
                                    'Folio = CFDI_H.SacaFolioPago()
                                    'Serie = "REP"
                                    GUID = CFDI_H.sacaGUID(aviso, Datos(2))
                                    GUID = GUID.ToUpper
                                End If
                            End If
                        Else
                            aviso = ""
                        End If
                        If Tipar <> "F" Then 'YA NO SEGUIR
                            Exit While
                        End If
                    Case "Z1"
                        If Tipar = "F" Then
                            OpcionCompraAF = Datos(7)
                            Numero = 0
                        End If
                End Select
            End While
        Catch ex As Exception
            EnviaError(GeneraFactura.My.Settings.MailError, ex.Message, "Error CFDI " & NombreArchivo)
            Console.WriteLine(GeneraFactura.My.Settings.NoPasa & NombreArchivo)
            Console.WriteLine(RutaArchivo)
            File.Copy(NombreArchivo, GeneraFactura.My.Settings.NoPasa & NombreArchivo, True)
        Finally
            f2.Close()
            If Serie = "B" Then
                JuntaDescripcion(RutaArchivo, NombreArchivo)
            End If
        End Try
    End Function

    Function ValidaRFC(rfc As String, tipo As String)
        If IsNumeric(Mid(rfc, 4, 1)) Then
            tipo = "M"
        End If
        If tipo = "F" Or tipo = "E" Then
            If rfc.Length < 13 Then
                rfc = "XAXX010101000"
            Else
                If Microsoft.VisualBasic.Right(rfc, 3) = "000" Then
                    rfc = "XAXX010101000"
                End If
            End If
        End If
        Return rfc
    End Function

    Function TruncarDecimales(Numero As Decimal) As Decimal
        TruncarDecimales = Math.Truncate(Numero) + (Math.Truncate((Numero - Math.Truncate(Numero)) * 100) / 100)
    End Function

    Function LimpiarConcepto(ByVal Concepto As String, Serie As String) As String
        Dim Cad As String = ""
        If Serie <> "B" Then
            For X = 1 To Concepto.Length
                If Not IsNumeric(Mid(Concepto, X, 1)) And Mid(Concepto, X, 1) <> "/" And Mid(Concepto, X, 1) <> "," And Mid(Concepto, X, 1) <> "." Then
                    Cad += Mid(Concepto, X, 1)
                End If
            Next
        Else
            Cad = Concepto
        End If
        Return Cad.Trim
    End Function

    Sub JuntaDescripcion(RutaArchivo As String, NombreArchivo As String)
        Dim f1 As System.IO.StreamWriter
        Dim f2 As System.IO.StreamReader
        Dim Cont As Integer = 0
        Dim Linea As String = ""
        Dim Linea1 As String = ""
        Dim Datos() As String
        Dim Descrip As String = ""

        f1 = New System.IO.StreamWriter(RutaArchivo & ".tmp")
        f2 = New System.IO.StreamReader(RutaArchivo, Text.Encoding.GetEncoding(1252))

        Try
            While Not f2.EndOfStream
                Linea = f2.ReadLine
                Datos = Linea.Split("|")
                If Datos(0) = "D1" Then
                    Cont += 1
                    If Cont = 1 And Val(Datos(10)) > 0 Then
                        Linea1 = Linea
                        Descrip = Datos(8).Trim
                    Else
                        Descrip += " " & Datos(8).Trim
                    End If
                Else
                    If Cont > 1 Then
                        Datos = Linea1.Split("|")
                        Linea1 = ""
                        For x = 0 To Datos.Length - 1
                            If x = 0 Then
                                Linea1 += Datos(x)
                            ElseIf x = 8 Then
                                Linea1 += "|" & Descrip
                            Else
                                Linea1 += "|" & Datos(x)
                            End If
                        Next
                        f1.WriteLine(Linea1)
                        Cont = 0
                    Else
                        f1.WriteLine(Linea)
                    End If
                End If
            End While
            f1.Close()
            f2.Close()
            File.Delete(RutaArchivo)
            My.Computer.FileSystem.RenameFile(RutaArchivo & ".tmp", NombreArchivo)
        Catch ex As Exception

        End Try

    End Sub

    Function LecturaPreviaAUX() As Boolean

        Dim D As New System.IO.DirectoryInfo("R:\CFDIbackup\")
        Dim F As FileInfo() = D.GetFiles("*.txt")


        Dim Numero As Integer = 1
        Dim f2 As System.IO.StreamReader
        Dim f1 As System.IO.StreamWriter
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim Linea As String
        Dim Datos() As String
        Dim Anexo As String = ""
        Dim FechaAviso As Date = "01/01/1900"
        Dim Iva As Decimal

        f1 = New System.IO.StreamWriter("c:\Temp\ErroresIVA.txt")
        For i = 0 To F.Length - 1
            f2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
            Console.WriteLine(F(i).FullName)
            While Not f2.EndOfStream
                Linea = f2.ReadLine
                Datos = Linea.Split("|")
                Select Case Datos(0)
                    Case "D1"
                        If Datos(11).Length > 0 Then
                            Iva = Math.Round(CDec(Datos(11)) / CDec(Datos(10)), 2)
                            If Iva > 0.16 Then
                                f1.WriteLine(Linea)
                            End If
                        End If
                End Select
            End While
            f2.Close()
        Next
        f1.Close()
    End Function

    Sub FacturasSinSERIE()
        Dim ta As New ProduccionDSTableAdapters.FacturasSinSerieTableAdapter
        ta.Fill(ProductDS.FacturasSinSerie, Date.Now.AddDays(-10).ToString("yyyyMMdd"))
        For Each r As ProduccionDS.FacturasSinSerieRow In ProductDS.FacturasSinSerie
            EnviaError("ecacerest@finagil.com.mx", "Factura sin Serie", "Aviso:" & r.Anexo & " Aviso: " & r.Factura)
            EnviaError("viapolo@finagil.com.mx", "Factura sin Serie", "Aviso:" & r.Anexo & " Aviso: " & r.Factura)
            EnviaError("denise.gonzalez@finagil.com.mx", "Factura sin Serie", "Aviso:" & r.Anexo & " Aviso: " & r.Factura)
        Next

    End Sub

End Module
