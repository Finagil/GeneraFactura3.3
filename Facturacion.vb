Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient

Module GneraFactura
    Dim ErrorControl As New EventLog
    Dim OpcionCompraAF As String

    Sub Main()
        Dim mf As Date = Date.Now.AddHours(-72)
        Console.WriteLine("Inicia proceso")
        ErrorControl = New EventLog("Application", System.Net.Dns.GetHostName(), "GeneracionCFDI33")
        Dim D As New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        Dim F As FileInfo() = D.GetFiles("*.txt")

        Console.WriteLine("Generando CFDI Avio...")
        'GeneraArchivosAvio()
        Console.WriteLine("Generando CFDI Facturas Externas...")
        GeneraArchivosEXternas()
        Console.WriteLine("leyendo " & GeneraFactura.My.Settings.RutaOrigen)
        Console.WriteLine("Generando CFDI...")
        GeneraArchivos()

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

        fecha = Date.Now.AddDays(-1)
        Facturas.QuitarPagosEfectivo()
        '***************************************************************
        'quita seguros que nos sean de guanajuato y michoacan haste que se hagan dos conceptos de seguros
        Detalles.FillBySeguro(DET)
        For Each rr As GeneraFactura.ProduccionDS.FacturasAvioDetalleRow In DET.Rows
            Facturas.Facturar("N/A", rr.Anexo, rr.Ciclo, rr.FechaFinal, rr.Concepto)
        Next
        '***************************************************************
        Facturas.QuitarPagosEfectivo()
        Facturas.Fill(FAC, fecha.ToString("yyyyMMdd"))

        For Each r As GeneraFactura.ProduccionDS.FacturasAvioRow In FAC.Rows
            TasaIVACliente = taCli.SacaTasaIVACliente(r.Cliente)
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

            Facturas.QuitarPagosEfectivo()
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
                    Or Trim(rr.Concepto) = "ANALISIS DE SUELOS" Then
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
                        Case "INTERESES"
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
                        ROWdetail._7_Impuesto_Porcentaje = "EXE"
                        ROWdetail._4_Impuesto_Monto_Impuesto = 0
                    Else
                        ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                        ROWdetail._4_Impuesto_Monto_Impuesto = TruncarDecimales((ROWdetail._5_Linea_Importe * TasaIVA))
                    End If

                    SubTT += ROWdetail._3_Impuesto_Monto_base
                    IVA += ROWdetail._4_Impuesto_Monto_Impuesto

                    ROWdetail.Detalle_Folio = ROWheader._1_Folio
                    ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante
                    ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                End If

                RegAfec = Facturas.Facturar("AV" & ROWheader._1_Folio, r.Anexo, r.Ciclo, rr.FechaFinal, Trim(rr.Concepto))
                If RegAfec = 0 Then
                    EnviaError(GeneraFactura.My.Settings.MailError, "Error Factura sin Afectar", "Error Factura sin Afectar" & r.Anexo)
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
                ROWdetail.Detalle_Folio = ROWheader._1_Folio
                ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante

                SubTT += ROWdetail._3_Impuesto_Monto_base
                IVA += ROWdetail._4_Impuesto_Monto_Impuesto

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

    Sub GeneraArchivos()
        Dim NoFactError As Integer
        Dim Folio As Integer
        Dim Serie As String = ""
        Dim GUID As String = ""
        Dim taCli As New GeneraFactura.ProduccionDSTableAdapters.ClientesTableAdapter
        Dim Facturas As New GeneraFactura.ProduccionDSTableAdapters.FacturasAvioTableAdapter
        Dim FAC As New GeneraFactura.ProduccionDS.FacturasAvioDataTable
        Dim Detalles As New GeneraFactura.ProduccionDSTableAdapters.FacturasAvioDetalleTableAdapter
        Dim DET As New GeneraFactura.ProduccionDS.FacturasAvioDetalleDataTable
        Dim Folios As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim ProducDS As New ProduccionDS
        Dim CFDI_H As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim CFDI_D As New ProduccionDSTableAdapters.CFDI_DetalleTableAdapter
        Dim CFDI_P As New ProduccionDSTableAdapters.CFDI_ComplementoPagoTableAdapter
        Dim ROWheader As ProduccionDS.CFDI_EncabezadoRow
        Dim ROWdetail As ProduccionDS.CFDI_DetalleRow
        Dim TasaIVACliente As Decimal
        Dim SubTT, IVA, MontoBaseIVA As Decimal
        Dim NoLineas As Integer
        Dim EsNotaCredito As Boolean = False
        Dim EnviarGisela As Boolean = False
        Dim ta As New GeneraFactura.ProduccionDSTableAdapters.ClientesTableAdapter
        Dim t As New GeneraFactura.ProduccionDS.ClientesDataTable
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
        Dim D As New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        Dim F As System.IO.FileInfo() = D.GetFiles("*.txt").OrderBy(Function(fi) fi.CreationTime).ToArray()
        Dim Aviso As String

        Dim Datos() As String
        Dim f2 As System.IO.StreamReader
        Dim fecha As New DateTime
        Dim horas As Integer
        Dim Tipar As String = ""
        Dim TipoPersona As String = ""
        Dim Moneda As String = ""
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim cAnexo, cAnexoAux As String
        Dim LeyendaCapital, Metodo_Pago, FormaPago As String
        Dim EsPago As Boolean
        'Try
        NoFactError = 0
        For i = 0 To F.Length - 1
            'Try
            Console.WriteLine("Generando CFDI..." & F(i).Name)
            NoLineas = 0
            suma = 0
            EsPago = False
            LecturaPrevia(F(i).FullName, F(i).Name, Moneda, Tipar, Folio, Serie, EsPago)
            If EsPago = True Then
#Region "Espago"
                f2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
                If Mid(F(i).Name, 1, 3) <> "FIN" And Mid(F(i).Name, 1, 3) <> "XXA" And IsNumeric(Mid(F(i).Name, 1, 4)) = True Then
                    fecha = New DateTime(Mid(F(i).Name, 1, 4), Mid(F(i).Name, 5, 2), Mid(F(i).Name, 7, 2), Mid(F(i).Name, 9, 2), Mid(F(i).Name, 11, 2), Mid(F(i).Name, 13, 2))
                    horas = DateDiff(DateInterval.Hour, fecha, Date.Now)
                    If horas >= 72 Then
                        fecha = fecha.AddHours(horas - 71)
                    End If
                End If
                EsNotaCredito = False
                SubTT = 0
                IVA = 0
                EnviarGisela = False
                Adenda = False
                LeyendaCapital = ""
                Errores = False
                ReDim Datos(1)
                Datos(0) = "X"
                cAnexoAux = ""
                While Not f2.EndOfStream
                    Linea = f2.ReadLine
                    If UCase(Linea) = "X" Then
                        EnviarGisela = True
                        Linea = f2.ReadLine
                    End If
                    Datos = Linea.Split("|")
                    If Datos.Length > 4 Then
                        cAnexoAux = Datos(2)
                        If Datos(2) = "03284/0001" Then Datos(2) = "29320141001-001"
                        If Datos(2) = "03285/0001" Then Datos(2) = "29477141001-001"
                        If Datos(2) = "03286/0001" Then Datos(2) = "29248141001-001"
                        If Datos(2) = "03287/0001" Then Datos(2) = "29291141001-001"
                        If Datos(2) = "03288/0001" Then Datos(2) = "29478141001-001"
                        If Datos(2) = "03289/0001" Then Datos(2) = "29475141001-001"
                        If Datos(2) = "02541/0023" Then Datos(2) = "10375101001-001"
                        If Datos(2) = "01966/0002" Then Datos(2) = "19177101001-001"
                        If Datos(2) = "01476/0003" Then Datos(2) = "01858101001-002"
                        If Datos(2) = "03291/0001" Then Datos(2) = "29484001"
                        If Datos(2) = "03282/0002" Then Datos(2) = "2885803-001"
                        If Datos(2) = "03292/0001" Then Datos(2) = "29290101001-001"
                        If Datos(2) = "08386/0006" Then Datos(2) = "04495150001-001"
                        If Datos(2) = "00223/0036" Then Datos(2) = "10284121001"
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
                            Metodo_Pago = Datos(2)
                            FormaPago = Datos(3)
                        Case "H3"
                            If Datos.Length > 28 Then
                                Aviso = CFDI_H.SacaAviso(Datos(28), Datos(29))
                                If Aviso = "0" Then Aviso = ""
                            Else
                                Aviso = ""
                            End If
                            If Metodo_Pago = "PAGO" Then
                                Folio = CFDI_H.SacaFolio()
                                Serie = "COMP"
                                GUID = CFDI_H.sacaGUID(Aviso)
                            End If

                            If Datos(2).Length <> 10 Then
                                cAnexo = Mid(cAnexoAux, 1, 5) & Mid(cAnexoAux, 7, 4)
                            Else
                                cAnexo = Mid(Datos(2), 1, 5) & Mid(Datos(2), 7, 4)
                            End If

                            TipoPersona = taTipar.TipoPersona(Datos(1))
                            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
                            TasaIVACliente = taCli.SacaTasaIVACliente(Datos(1))
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
                            ROWheader._29_FormaPago = FormaPago '"27" '27 A satisfacción del acreedor
                            ROWheader._30_Fecha = fecha.Date
                            ROWheader._31_Hora = fecha.ToString("HH:mm:ss")
                            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

                            If Moneda = "WWW" Then
                                If Datos(17) = "M.N." Then Datos(17) = "MXN"
                                If Datos(17) = "MXP" Then Datos(17) = "MXN"
                                Moneda = Datos(17)
                            End If

                            Datos(16) = ValidaRFC(Datos(16), TipoPersona)
                            If Trim(Datos(16)) = "SDA070613KU6" Then
                                Datos(5) = """SERVICIO DAYCO"" SA DE CV"
                            End If
                            If Trim(Datos(16)) = "CARD840606LEA" Then
                                Datos(5) = "DANIEL CADENA RUVALCABA"
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
                            UsoCFDI = taCodigo.SacaUsoCFDI(cAnexo)
                            If UsoCFDI.Trim = "" Then
                                UsoCFDI = taCodigo.SacaUsoCFDI(Tipar)
                            End If
                            ROWheader._144_Misc32 = UsoCFDI 'claves del SAT P01=por definir, G03=Gastos generales CATALOGO USO DE COMPROBANTE hoja excel= c_UsoCFDI, a solicitud del cliente
                            ROWheader._167_RegimentFiscal = 601
                            If ROWheader._83_Cod_Moneda = "MXN" Then
                                ROWheader._177_Tasa_Divisa = 0
                            Else
                                ROWheader._177_Tasa_Divisa = taCli.SacaTipoCambio(fecha, ROWheader._83_Cod_Moneda)
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
                                        Datos(8) = "CAPITAL VENCIMIENTO" & Right(Datos(8), 7)
                                    End If
                                End If
                                If (Tipar = "F") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "INTERES OTRO", "INTERES SEGU"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If (Tipar = "R" Or Tipar = "S") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "MORATORIOS V"
                                            TipoImpuesto = "Exento"
                                        Case "INTERESES VE", "INTERES OTRO"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If Tipar = "F" And cAnexo = "02562/0003" Then '#ECT Solicitado por Valentin 24/09/2015
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
                                End If

                                If Tipar = "B" Then
                                    Select Case Mid(Datos(8), 1, 11)
                                        Case "MENSUALIDAD"
                                            Datos(8) = "SERVICIO DE TRANSPORTE EJECUTIVO EMPRESARIAL, " & Datos(8)
                                    End Select
                                End If

                                taCodigo.Fill(tCodigo, Tipar, Concepto)
                                If tCodigo.Rows.Count > 0 Then
                                    rCod = tCodigo.Rows(0)
                                    Unidad = rCod.Unidad
                                    Codigo = rCod.Codigo
                                    If Codigo = "" Then
                                        Codigo = "84101700"
                                        Errores = True
                                    End If
                                    If Unidad = "" Then
                                        Unidad = "E48"
                                        Errores = True
                                    End If
                                    If Errores = True And (Tipar = "F" Or Tipar = "S") And Concepto = "CAPITAL EQUIPO VENCIMIENTO" Then
                                        Errores = False 'SE QUITA CUANDO ESTE CONFIGURADOS LOS ARTICULOS Y UNIDADES
                                    End If
                                    If Errores = True And Tipar = "P" And Concepto = "PAGO DE RENTA VENCIMIENTO" Then
                                        Errores = False 'SE QUITA CUANDO ESTE CONFIGURADOS LOS ARTICULOS Y UNIDADES
                                    End If

                                Else
                                    If taCodigo.ExisteConcepto(Tipar, Concepto) <= 0 And ROWheader._27_Serie_Comprobante <> "B" Then
                                        taCodigo.Insert(Tipar, Concepto, "", "", False)
                                    End If
                                    Unidad = "E48"
                                    Codigo = "84101700"
                                    Errores = True
                                End If
                                If Errores = True Then
                                    'EnviacORREO("vcruz@finagil.com.mx", "Concepto: " & Concepto & " TipoCredito : " & Tipar & " Anexo : " & cAnexo, "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
                                    'EnviacORREO("ecacerest@finagil.com.mx", "Concepto: " & Concepto & vbCrLf & " TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo, "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
                                End If

                                ROWdetail._1_Linea_Descripcion = "Pago"
                                ROWdetail._2_Linea_Cantidad = 1
                                ROWdetail._3_Linea_Unidad = "ACT"
                                ROWdetail._4_Linea_PrecioUnitario = 0
                                ROWdetail._5_Linea_Importe = 0
                                ROWdetail._16_Linea_Cod_Articulo = Codigo ' Manejo de deuda
                                'ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                                'ROWdetail._2_Impuesto_Descripcion = "TR"
                                'ROWdetail._3_Impuesto_Monto_base = CDec(Datos(10)).ToString("n2")
                                'ROWdetail._5_Impuesto_Clave = "002"
                                'ROWdetail._6_Impuesto_Tasa = "Tasa"
                                'If TipoImpuesto = "Exento" Then
                                '    'ROWdetail._7_Impuesto_Porcentaje = "EXE"
                                '    'ROWdetail._4_Impuesto_Monto_Impuesto = 0
                                'Else
                                '    'ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                                '    If TasaIVA = 0 Or CDec(Datos(11)) = 0 Then
                                '        'ROWdetail._4_Impuesto_Monto_Impuesto = 0
                                '    Else
                                '        'ROWdetail._4_Impuesto_Monto_Impuesto = CDec(Datos(11)).ToString("n2")
                                '        'MontoBaseIVA = CDec(Datos(11)) / TasaIVA
                                '        'If MontoBaseIVA < ROWdetail._3_Impuesto_Monto_base Then
                                '        'ROWdetail._3_Impuesto_Monto_base = MontoBaseIVA
                                '        'End If
                                '    End If
                                'End If
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
                                    EnviacORREO("ecacerest@finagil.com.mx", ex.Message, "Error de Factura " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
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
                    ROWheader._113_Misc01 = "CPG"
                    ROWheader._114_Misc02 = Datos(2)
                    ROWheader._115_Misc03 = Datos(1)
                    ROWheader._132_Misc20 = "CPG"
                    ROWheader._159_Misc47 = Aviso
                    ROWheader._162_Misc50 = ""
                    'ROWheader._161_Misc49 = ""
                    If OpcionCompraAF.Trim.Length > 0 Then
                        ROWheader._160_Misc48 = LeyendaCapital.Trim & " " & OpcionCompraAF.Trim
                    Else
                        ROWheader._160_Misc48 = LeyendaCapital.Trim
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
                        Dim SaldoFactura As Decimal = Total
                        ProducDS.CFDI_Encabezado.GetChanges()
                        ProducDS.CFDI_Detalle.GetChanges()
                        CFDI_D.Update(ProducDS.CFDI_Detalle)
                        CFDI_H.Update(ProducDS.CFDI_Encabezado)
                        CFDI_P.Insert("CPG", "Pagos", "HD", "1.0", fecha.ToString("yyyy/MM/ddThh:mm:ss"), "03", Moneda, "", Total, "", "", "", "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        CFDI_P.Insert("CPG", "Pago", "HD", "", "", "", Moneda, "", "", "", "", "", "", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        CFDI_P.Insert("CPG", "DoctoRelacionado", "HD", GUID, Serie, Folio, Moneda, "", "PPD", "1", SaldoFactura, Total, "0.0", "", "", "", "", "", Folio, Serie, Folio, Serie)
                        CFDI_H.ConsumeFolio()

                        ProducDS.CFDI_Encabezado.Clear()
                        ProducDS.CFDI_Detalle.Clear()
                        File.Copy(F(i).FullName, GeneraFactura.My.Settings.RutaBackup & F(i).Name, True)
                        File.Delete(F(i).FullName)

                    Catch ex As Exception
                        EnviacORREO("ecacerest@finagil.com.mx", "Error Factura TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo, "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
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
            Else
#Region "Factura"
                f2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
                If Mid(F(i).Name, 1, 3) <> "FIN" And Mid(F(i).Name, 1, 3) <> "XXA" And IsNumeric(Mid(F(i).Name, 1, 4)) = True Then
                    fecha = New DateTime(Mid(F(i).Name, 1, 4), Mid(F(i).Name, 5, 2), Mid(F(i).Name, 7, 2), Mid(F(i).Name, 9, 2), Mid(F(i).Name, 11, 2), Mid(F(i).Name, 13, 2))
                    horas = DateDiff(DateInterval.Hour, fecha, Date.Now)
                    If horas >= 72 Then
                        fecha = fecha.AddHours(horas - 71)
                    End If
                End If
                EsNotaCredito = False
                SubTT = 0
                IVA = 0
                EnviarGisela = False
                Adenda = False
                LeyendaCapital = ""
                Errores = False
                ReDim Datos(1)
                Datos(0) = "X"
                cAnexoAux = ""
                While Not f2.EndOfStream
                    Linea = f2.ReadLine
                    If UCase(Linea) = "X" Then
                        EnviarGisela = True
                        Linea = f2.ReadLine
                    End If
                    Datos = Linea.Split("|")
                    If Datos.Length > 4 Then
                        cAnexoAux = Datos(2)
                        If Datos(2) = "03284/0001" Then Datos(2) = "29320141001-001"
                        If Datos(2) = "03285/0001" Then Datos(2) = "29477141001-001"
                        If Datos(2) = "03286/0001" Then Datos(2) = "29248141001-001"
                        If Datos(2) = "03287/0001" Then Datos(2) = "29291141001-001"
                        If Datos(2) = "03288/0001" Then Datos(2) = "29478141001-001"
                        If Datos(2) = "03289/0001" Then Datos(2) = "29475141001-001"
                        If Datos(2) = "02541/0023" Then Datos(2) = "10375101001-001"
                        If Datos(2) = "01966/0002" Then Datos(2) = "19177101001-001"
                        If Datos(2) = "01476/0003" Then Datos(2) = "01858101001-002"
                        If Datos(2) = "03291/0001" Then Datos(2) = "29484001"
                        If Datos(2) = "03282/0002" Then Datos(2) = "2885803-001"
                        If Datos(2) = "03292/0001" Then Datos(2) = "29290101001-001"
                        If Datos(2) = "08386/0006" Then Datos(2) = "04495150001-001"
                        If Datos(2) = "00223/0036" Then Datos(2) = "10284121001"
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
                            Metodo_Pago = Datos(2)
                            FormaPago = Datos(3)
                        Case "H3"
                            If Datos.Length > 28 Then
                                Aviso = CFDI_H.SacaAviso(Datos(28), Datos(29))
                                If Aviso = "0" Then Aviso = ""
                            Else
                                Aviso = ""
                            End If

                            If Datos(2).Length <> 10 Then
                                cAnexo = Mid(cAnexoAux, 1, 5) & Mid(cAnexoAux, 7, 4)
                            Else
                                cAnexo = Mid(Datos(2), 1, 5) & Mid(Datos(2), 7, 4)
                            End If

                            TipoPersona = taTipar.TipoPersona(Datos(1))
                            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
                            TasaIVACliente = taCli.SacaTasaIVACliente(Datos(1))
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
                            ROWheader._31_Hora = fecha.ToString("HH:mm:ss")
                            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

                            If Mid(Serie, 1, 1) = "C" Then
                                EsNotaCredito = True
                            Else
                                EsNotaCredito = False
                            End If

                            If Moneda = "WWW" Then
                                If Datos(17) = "M.N." Then Datos(17) = "MXN"
                                If Datos(17) = "MXP" Then Datos(17) = "MXN"
                                Moneda = Datos(17)
                            End If

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
                                ROWheader._177_Tasa_Divisa = taCli.SacaTipoCambio(fecha, ROWheader._83_Cod_Moneda)
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
                                ROWdetail = ProducDS.CFDI_Detalle.NewCFDI_DetalleRow
                                TasaIVA = TasaIVACliente / 100
                                TipoImpuesto = TasaIVACliente
                                NoLineas += 1
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
                                        Datos(8) = "CAPITAL VENCIMIENTO" & Right(Datos(8), 7)
                                    End If
                                End If
                                If (Tipar = "F") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "INTERES OTRO", "INTERES SEGU"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If (Tipar = "R" Or Tipar = "S") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "MORATORIOS V"
                                            TipoImpuesto = "Exento"
                                        Case "INTERESES VE", "INTERES OTRO"
                                            TipoImpuesto = "Exento"
                                    End Select
                                End If
                                If Tipar = "F" And cAnexo = "02562/0003" Then '#ECT Solicitado por Valentin 24/09/2015
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
                                End If

                                If Tipar = "B" Then
                                    Select Case Mid(Datos(8), 1, 11)
                                        Case "MENSUALIDAD"
                                            Datos(8) = "SERVICIO DE TRANSPORTE EJECUTIVO EMPRESARIAL, " & Datos(8)
                                    End Select
                                End If

                                taCodigo.Fill(tCodigo, Tipar, Concepto)
                                If tCodigo.Rows.Count > 0 Then
                                    rCod = tCodigo.Rows(0)
                                    Unidad = rCod.Unidad
                                    Codigo = rCod.Codigo
                                    If Codigo = "" Then
                                        Codigo = "84101700"
                                        Errores = True
                                    End If
                                    If Unidad = "" Then
                                        Unidad = "E48"
                                        Errores = True
                                    End If
                                    If Errores = True And (Tipar = "F" Or Tipar = "S") And Concepto = "CAPITAL EQUIPO VENCIMIENTO" Then
                                        Errores = False 'SE QUITA CUANDO ESTE CONFIGURADOS LOS ARTICULOS Y UNIDADES
                                    End If
                                    If Errores = True And Tipar = "P" And Concepto = "PAGO DE RENTA VENCIMIENTO" Then
                                        Errores = False 'SE QUITA CUANDO ESTE CONFIGURADOS LOS ARTICULOS Y UNIDADES
                                    End If

                                Else
                                    If taCodigo.ExisteConcepto(Tipar, Concepto) <= 0 And ROWheader._27_Serie_Comprobante <> "B" Then
                                        taCodigo.Insert(Tipar, Concepto, "", "", False)
                                    End If
                                    Unidad = "E48"
                                    Codigo = "84101700"
                                    Errores = True
                                End If
                                If Errores = True Then
                                    'EnviacORREO("vcruz@finagil.com.mx", "Concepto: " & Concepto & " TipoCredito : " & Tipar & " Anexo : " & cAnexo, "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
                                    'EnviacORREO("ecacerest@finagil.com.mx", "Concepto: " & Concepto & vbCrLf & " TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo, "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
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
                                If Datos(6).Trim = "" Then
                                    Datos(6) = "SER"
                                End If
                                ROWdetail._53_Linea_Misc22 = Datos(6)
                                If TipoImpuesto = "Exento" Then
                                    ROWdetail._7_Impuesto_Porcentaje = "EXE"
                                    ROWdetail._4_Impuesto_Monto_Impuesto = 0
                                Else
                                    ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                                    If TasaIVA = 0 Or CDec(Datos(11)) = 0 Then
                                        ROWdetail._4_Impuesto_Monto_Impuesto = 0
                                    Else
                                        ROWdetail._4_Impuesto_Monto_Impuesto = CDec(Datos(11)).ToString("n2")
                                        MontoBaseIVA = CDec(Datos(11)) / TasaIVA
                                        If MontoBaseIVA < ROWdetail._3_Impuesto_Monto_base Then
                                            ROWdetail._3_Impuesto_Monto_base = MontoBaseIVA
                                        End If
                                    End If
                                End If

                                SubTT += ROWdetail._5_Linea_Importe
                                IVA += ROWdetail._4_Impuesto_Monto_Impuesto

                                ROWdetail.Detalle_Folio = ROWheader._1_Folio
                                ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante
                                Try
                                    ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                                Catch ex As Exception
                                    Errores = False
                                    EnviacORREO("ecacerest@finagil.com.mx", ex.Message, "Error de Factura " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
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
                    ROWheader._159_Misc47 = Aviso
                    ROWheader._162_Misc50 = ""
                    'ROWheader._161_Misc49 = ""
                    If OpcionCompraAF.Trim.Length > 0 Then
                        ROWheader._160_Misc48 = LeyendaCapital.Trim & " " & OpcionCompraAF.Trim
                    Else
                        ROWheader._160_Misc48 = LeyendaCapital.Trim
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
                        EnviacORREO("ecacerest@finagil.com.mx", "Error Factura TipoCredito : " & Tipar & vbCrLf & " Anexo : " & cAnexo, "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
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
            End If
        Next
        If NoFactError > 0 Then
            EnviacORREO("ecacerest@finagil.com.mx", "Error  de Facturas sin procesar:  " & NoFactError, "Error  de Facturas sin procesar:  " & NoFactError, "CFDI33@finagil.com.mx")
        End If
    End Sub

    Private Sub EnviaError(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String)
        If InStr(Mensaje, "No se ha encontrado la ruta de acceso de la red") = 0 Then
            Dim Mensage As New MailMessage("InternoBI2008@cmoderna.com", Trim(Para), Trim(Asunto), Mensaje)
            Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
            Try
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
        Dim ROWheader As ProduccionDS.CFDI_EncabezadoRow
        Dim ROWdetail As ProduccionDS.CFDI_DetalleRow
        Dim ProducDS As New ProduccionDS
        Dim TasaIVACliente, MontoBaseIVA As Decimal
        Dim NoLineas As Integer
        Dim RFC As String = ""
        Dim Razon As String = ""
        Dim SubTT As Double
        Dim IVA As Double
        Dim TOt As Double
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
            Console.WriteLine("Generando CFDI Facturas Externas..." & r.Factura)
            SubTT = 0
            TOt = 0
            IVA = 0
            Errores = False
            ROWheader = ProducDS.CFDI_Encabezado.NewCFDI_EncabezadoRow
            ROWheader._1_Folio = Val(r.Factura)
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
            ROWheader._27_Serie_Comprobante = r.Serie.Trim
            ROWheader._29_FormaPago = r.MetodoPago.Trim '"27" '27 A satisfacción del acreedor
            ROWheader._30_Fecha = Today.Date
            ROWheader._31_Hora = Today.ToString("HH:mm:ss")
            ROWheader._41_Dom_LugarExpide_codigoPostal = "50070"

            If IsNumeric(Mid(r.RFC, 4, 1)) Then
                TipoPersona = "M"
            Else
                TipoPersona = "F"
            End If
            RFC = ValidaRFC(RFC, TipoPersona)

            If RFC = "SDA070613KU6" Then
                Razon = """SERVICIO DAYCO"" SA DE CV"
            Else
                Razon = r.Nombre
            End If

            ROWheader._42_Nombre_Receptor = Razon.Trim
            ROWheader._43_RFC_Receptor = RFC
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
            ROWheader._144_Misc32 = "G03" 'claves del SAT P01=por definir, G03=Gastos generales CATALOGO USO DE COMPROBANTE hoja excel= c_UsoCFDI, a solicitud del cliente
            ROWheader._167_RegimentFiscal = 601
            If ROWheader._83_Cod_Moneda = "MXN" Then
                ROWheader._177_Tasa_Divisa = 0
            Else
                ROWheader._177_Tasa_Divisa = Facturas.SacaTipoCambio(Today, ROWheader._83_Cod_Moneda)
            End If
            ROWheader._180_LugarExpedicion = "50070"
            ROWheader._190_Metodo_Pago = r.MetodoPagoSAT 'PPD pago en parcialidades PUE pago en una sola exhibision
            ROWheader._191_Efecto_Comprobante = "I"
            ROWheader._58_TipoCFD = "FA"

            Detalles.Fill(DET, r.Serie, r.Factura)
            For Each rr As GeneraFactura.ProduccionDS.FacturasExternasDETRow In DET.Rows
                NoLineas += 1
                ROWdetail = ProducDS.CFDI_Detalle.NewCFDI_DetalleRow
                ROWdetail._1_Linea_Descripcion = rr.Detalle
                ROWdetail._2_Linea_Cantidad = 1
                ROWdetail._3_Linea_Unidad = rr.Unidad
                ROWdetail._4_Linea_PrecioUnitario = Math.Round(rr.Unitario, 4)
                ROWdetail._5_Linea_Importe = Math.Round(rr.Importe, 4)
                ROWdetail._16_Linea_Cod_Articulo = rr.CodigoART ' Manejo de deuda
                ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                ROWdetail._2_Impuesto_Descripcion = "TR"
                ROWdetail._3_Impuesto_Monto_base = Math.Round(rr.Importe, 4)
                ROWdetail._5_Impuesto_Clave = "002"
                ROWdetail._6_Impuesto_Tasa = "Tasa"

                If rr.TasaIva = "Exento" Then
                    ROWdetail._7_Impuesto_Porcentaje = "EXE"
                    ROWdetail._4_Impuesto_Monto_Impuesto = 0
                Else
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
                    End If
                End If

                SubTT += ROWdetail._5_Linea_Importe
                IVA += ROWdetail._4_Impuesto_Monto_Impuesto

                ROWdetail.Detalle_Folio = ROWheader._1_Folio
                ROWdetail.Detalle_Serie = ROWheader._27_Serie_Comprobante
                Try
                    ProducDS.CFDI_Detalle.AddCFDI_DetalleRow(ROWdetail)
                Catch ex As Exception
                    Errores = True
                    EnviacORREO("ecacerest@finagil.com.mx", ex.Message, "Error de Factura " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
                End Try
                Facturas.Facturar(r.Serie, r.Factura, rr.Consec)
            Next


            ROWheader._90_Cantidad_LineasFactura = NoLineas
            ROWheader._54_Monto_SubTotal = SubTT
            ROWheader._55_Monto_IVA = IVA
            ROWheader._56_Monto_Total = ROWheader._54_Monto_SubTotal + ROWheader._55_Monto_IVA
            ROWheader._193_Monto_TotalImp_Trasladados = ROWheader._55_Monto_IVA
            ROWheader._100_Letras_Monto_Total = Letras(ROWheader._56_Monto_Total, r.Moneda)
            ROWheader._114_Misc02 = "" ' contrato
            ROWheader._115_Misc03 = "" ' Cliente
            ROWheader._162_Misc50 = ""
            ROWheader._160_Misc48 = "" ' Adenda
            ROWheader._161_Misc49 = ""
            ROWheader.Encabezado_Procesado = False
            ROWheader._162_Misc50 = r.Mail1.Trim
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
                    EnviacORREO("ecacerest@finagil.com.mx", "Error Factura TipoCredito : EXTERNA", "Factura sin Procesar " & ROWheader._1_Folio & ROWheader._27_Serie_Comprobante, "CFDI33@finagil.com.mx")
                    ProducDS.CFDI_Encabezado.Clear()
                    ProducDS.CFDI_Detalle.Clear()
                End Try
            Else
                ProducDS.CFDI_Encabezado.Clear()
                ProducDS.CFDI_Detalle.Clear()
            End If
        Next

    End Sub

    Sub LecturaPrevia(RutaArchivo As String, NombreArchivo As String, ByRef Moneda As String, ByRef Tipar As String, ByRef Folio As Integer, ByRef Serie As String, ByRef EsPAgo As Boolean)
        OpcionCompraAF = ""
        Dim Numero As Integer = 1
        Dim f2 As System.IO.StreamReader
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim Linea As String
        Dim Datos() As String
        f2 = New System.IO.StreamReader(RutaArchivo, Text.Encoding.GetEncoding(1252))
        Try
            While Not f2.EndOfStream
                Linea = f2.ReadLine
                Datos = Linea.Split("|")
                If Numero = 1 And Datos.Length > 4 Then
                    Folio = Val(Datos(4))
                    Serie = Datos(3)
                    Moneda = taTipar.SacaMoneda(Datos(2))
                    Tipar = taTipar.TipaR(Datos(2))
                    If Tipar = "F" Then
                        Select Case Datos(0)
                            Case "Z1"
                                If Tipar = "F" Then
                                    OpcionCompraAF = Datos(7)
                                    Numero = 0
                                End If
                                Continue While
                            Case Else
                                Continue While
                        End Select
                    Else
                        Exit While
                    End If
                Else
                    If Datos(2) = "PAGO" Then
                        EsPAgo = True
                    Else
                        EsPAgo = False
                    End If
                End If
            End While

        Catch ex As Exception
            EnviaError(GeneraFactura.My.Settings.MailError, ex.Message, "Error CFDI " & NombreArchivo)
            Console.WriteLine(GeneraFactura.My.Settings.NoPasa & NombreArchivo)
            Console.WriteLine(RutaArchivo)
            File.Copy(NombreArchivo, GeneraFactura.My.Settings.NoPasa & NombreArchivo, True)
        Finally
            f2.Close()
        End Try
    End Sub

    Public Function Estado_de_Cuenta_Avio(ByVal cAnexo As String, ByVal cCiclo As String, ByVal Proyectado As Integer, ByVal Usuario As String, Mensual As String)
        Dim cnAgil As New SqlConnection("Server=SERVER-RAID; DataBase=Production; User ID=User_PRO; pwd=User_PRO2015")
        Dim Res As Object
        Dim cm1 As New SqlCommand()
        With cm1
            .CommandType = CommandType.StoredProcedure
            If Mensual.ToUpper = "SI" Then
                .CommandText = "dbo.EstadoCuentaAvio_MENSUAL"
            Else
                .CommandText = "dbo.EstadoCuentaAvio"
            End If

            .CommandTimeout = 50
            .Parameters.AddWithValue("Anexo", cAnexo)
            .Parameters.AddWithValue("Ciclo", cCiclo)
            .Parameters.AddWithValue("Proyectado", Proyectado)
            .Parameters.AddWithValue("usuario", Usuario)
            .Connection = cnAgil
        End With
        cnAgil.Open()
        Res = cm1.ExecuteScalar()
        cnAgil.Close()
        cnAgil.Dispose()
        cm1.Dispose()
        Return (Res)
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

    Sub LeerConceptos()
        Dim F2 As StreamReader
        Dim taCodigo As New ProduccionDSTableAdapters.CodigosSATTableAdapter
        Dim tCodigo As New ProduccionDS.CodigosSATDataTable
        Dim Linea, Concepto As String
        Dim D As New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        Dim F As System.IO.FileInfo() = D.GetFiles("*.txt")
        Dim Datos() As String
        Dim Tipar As String = ""
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        'Try

        For i = 0 To F.Length - 1
            'Try
            Console.WriteLine("Leyendo conceptos..." & F(i).Name)
            F2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
            Console.WriteLine(F(i).Name)
            While Not F2.EndOfStream
                Linea = f2.ReadLine
                Datos = Linea.Split("|")
                Tipar = taTipar.TipaR(Datos(2))
                Select Case Datos(0)
                    Case "D1"
                        If InStr(UCase(Datos(8)), "IVA") <= 0 Then
                            Datos(8) = Trim(Datos(8))
                            Concepto = LimpiarConcepto(Datos(8), Tipar)
                            taCodigo.Fill(tCodigo, Tipar, Concepto)
                            Console.WriteLine(Concepto)
                            If tCodigo.Rows.Count > 0 Then
                            Else
                                If taCodigo.ExisteConcepto(Tipar, Concepto) <= 0 And Tipar <> "B" And Concepto.Length <= 50 Then
                                    taCodigo.Insert(Tipar, Concepto.Substring(0, Concepto.Length - 1), "", "", False)
                                End If
                            End If
                        End If
                End Select
            End While
            f2.Close()
        Next
    End Sub

End Module
