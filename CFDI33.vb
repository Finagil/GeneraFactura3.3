﻿Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Math
Imports System.WeakReference
Imports System.Xml
Imports System.Text
Module CFDI33
    Dim drUdis As DataRowCollection
    'Dim nIDSerieA As Decimal = 0
    'Dim nIDSerieMXL As Decimal = 0
    Dim cSerie As String = ""
    Dim cSucursal As String = ""
    Dim nTasaIVACliente As Decimal = 0
    Dim Production_AUXDataSet As New ProduccionDS
    Dim CFDI_EncabezadoTableAdapter As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
    Dim CFDI_DetalleTableAdapter As New ProduccionDSTableAdapters.CFDI_DetalleTableAdapter
    Dim CFDI_ComplementoPagoTableAdapter As New ProduccionDSTableAdapters.CFDI_ComplementoPagoTableAdapter
    Dim CFDI_EncabezadoNominaTableAdapter As New ProduccionDSTableAdapters.CFDI_Encabezado_NominaTableAdapter
    Dim Tafolios As New ProduccionDSTableAdapters.TraspasosAvioCCTableAdapter

    Sub FacturarCFDI(Tipo As String)
        Dim TaAvisos As New ProduccionDSTableAdapters.AvisosCFDITableAdapter
        Dim TaUdis As New ProduccionDSTableAdapters.TraeUdisTableAdapter
        Dim ProdDS As New ProduccionDS
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionStringFACTURA)
        Dim cm1 As New SqlCommand()
        Dim cm2 As New SqlCommand()
        Dim cm3 As New SqlCommand()
        Dim cm4 As New SqlCommand()
        Dim daSeries As New SqlDataAdapter(cm1)
        Dim daFacturas As New SqlDataAdapter(cm2)

        Dim dsAgil As New DataSet()
        Dim dtMovimientos As New DataTable("Movimientos")
        Dim drSerie As DataRow
        Dim strUpdate As String = ""
        Dim strInsert As String = ""
        Dim InstrumentoMonetario As String = ""
        Dim MetodoPago As String

        ' Declaración de variables de datos

        Dim cBanco As String = ""
        Dim cCheque As String = ""
        Dim cAnexo As String = ""
        Dim cReferencia As String = ""
        Dim cLetra As String = ""
        Dim cTipar As String = ""
        Dim cTipo As String = ""
        Dim nImporte As Decimal = 0
        Dim nSaldo As Decimal = 0
        Dim nDiasMoratorios As Decimal = 0
        Dim nTasaMoratoria As Decimal = 0
        Dim nMoratorios As Decimal = 0
        Dim nIvaMoratorios As Decimal = 0
        Dim nMontoPago As Decimal = 0
        Dim cFeven As String = ""
        Dim cFepag As String = ""
        Dim cFechaPago As String = ""
        Dim i As Integer = 0
        Dim nRecibo As Decimal = 0
        Dim FechaProc As Date = TaAvisos.ScalarFechaAplicacion

        ' Primero creo la tabla Movimientos que contendrá los registros contables de la cobranza

        dtMovimientos.Columns.Add("Anexo", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Letra", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Tipos", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Fepag", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Cve", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Imp", Type.GetType("System.Decimal"))
        dtMovimientos.Columns.Add("Tip", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Catal", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Esp", Type.GetType("System.Decimal"))
        dtMovimientos.Columns.Add("Coa", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Tipmon", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Banco", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Concepto", Type.GetType("System.String"))
        dtMovimientos.Columns.Add("Factura", Type.GetType("System.String"))

        ' El siguiente Command trae los consecutivos de cada Serie

        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT IDSerieA, IDSerieMXL FROM Llaves"
            .Connection = cnAgil
        End With

        ' Llenar el DataSet lo cual abre y cierra la conexión

        daSeries.Fill(dsAgil, "Series")

        ' Toma el número consecutivo de facturas de pago -que depende de la Serie- y lo incrementa en uno

        drSerie = dsAgil.Tables("Series").Rows(0)
        'nIDSerieA = drSerie("IDSerieA")
        'nIDSerieMXL = drSerie("IDSerieMXL")

        ' Solo necesito saber el número de elementos que tiene el DataGridView1
        Select Case Tipo.ToUpper
            Case "MANUAL" ' prepagos antes de su fecha de vencimiento
                TaAvisos.FillByManual(ProdDS.AvisosCFDI)'Facturadas por el usuario
            Case "PREPAGO" ' prepagos antes de su fecha de vencimiento
                cFechaPago = FechaProc.ToString("yyyyMMdd")
                TaAvisos.FillByPrepagos(ProdDS.AvisosCFDI, cFechaPago, "20171201")'Fecha de Salida a Producion
            Case "DIA" 'avisos de vencimiento del dia
                cFechaPago = FechaProc.AddDays(-1).ToString("yyyyMMdd") 'se factura todo lo que resta y no se aplico nada de un dia antes de la fecha de trabajo
                TaAvisos.FillporDIA(ProdDS.AvisosCFDI, cFechaPago)
            Case "ANTERIORES" ' avisos generados despues de su vencimiento
                If Date.Now.Hour >= 21 Then 'se factura todo lo que resta y no se aplico nada
                    cFechaPago = Date.Now.AddHours(-72).ToString("yyyyMMdd")
                Else
                    cFechaPago = FechaProc.AddDays(-1).ToString("yyyyMMdd")
                End If
                TaAvisos.FillByAnteriores(ProdDS.AvisosCFDI, cFechaPago)
        End Select

        'TaAvisos.FillHastaFecha(ProdDS.AvisosCFDI, cFechaPago)
        TaUdis.Fill(ProdDS.TraeUdis)
        drUdis = ProdDS.TraeUdis.Rows
        For Each r As ProduccionDS.AvisosCFDIRow In ProdDS.AvisosCFDI.Rows
            Console.WriteLine("Aviso:" & r.Factura & " " & cFechaPago)

            If TaAvisos.AnexosNoFacturables(r.Anexo) > 0 Then
                TaAvisos.FacturarAviso(True, "", 0, r.Factura, r.Anexo)
                Continue For
            End If
            If r.SaldoFac = 0 Then ' con esto no generamos facturas pagadas en una sola exhibicion
                If TaAvisos.NumeroDePagos(r.Anexo, r.Letra) = 1 And Tipo.ToUpper <> "MANUAL" Then
                    Continue For
                End If
            End If

            cAnexo = r.Anexo
            If r.Fepag.Trim.Length > 0 And r.Fepag < r.Feven Then
                cFechaPago = r.Fepag
            Else
                cFechaPago = r.Feven
            End If

            cBanco = ""
            cReferencia = ""
            nImporte = r.SaldoFac
            cCheque = "Facturacion CFDI"
            cBanco = "02" 'bancomer
            nDiasMoratorios = 0
            nTasaMoratoria = 0
            nMoratorios = 0
            nIvaMoratorios = 0
            cFeven = r.Feven
            cFepag = r.Feven

            ' Traigo la Sucursal y la Tasa de IVA que aplica al cliente a efecto de poder determinar la Serie a utilizar

            cSucursal = r.Sucursal
            nTasaIVACliente = r.TasaIVA


            If cSucursal = "04" Or cSucursal = "08" Or nTasaIVACliente = 11 Then
                cSerie = "MXL"
            Else
                cSerie = "A"
            End If

            If r.Tipar <> "B" Then
                nMontoPago = r.ImporteFac * 2
            Else
                nMontoPago = (r.IvaCapital + r.RenPr) * 2
            End If

            If nMontoPago > 3 Then
                If cSerie = "A" Then
                    'nIDSerieA = nIDSerieA + 1
                    nRecibo = Tafolios.SerieA
                ElseIf cSerie = "MXL" Then
                    'nIDSerieMXL = nIDSerieMXL + 1
                    nRecibo = Tafolios.SerieMXL
                End If
                MetodoPago = "PPD"
                cLetra = r.Letra
                Acepagov(cAnexo, cLetra, nMontoPago, nMoratorios, nIvaMoratorios, cBanco, cCheque, dtMovimientos, cFechaPago, cFechaPago, cSerie, nRecibo, InstrumentoMonetario, FechaProc, MetodoPago)

                If cSerie = "A" And nRecibo <> 0 Then
                    'strUpdate = "UPDATE Llaves SET IDSerieA = " & nRecibo
                    Tafolios.ConsumeSerieA()
                ElseIf cSerie = "MXL" And nRecibo <> 0 Then
                    'strUpdate = "UPDATE Llaves SET IDSerieMXL = " & nRecibo
                    Tafolios.ConsumeSerieMXL()
                End If
                TaAvisos.FacturarAviso(True, cSerie.Trim, nRecibo, r.Factura, r.Anexo)
                'cm1 = New SqlCommand(strUpdate, cnAgil)
                'cnAgil.Open()
                'cm1.ExecuteNonQuery()
                'cnAgil.Close()
            End If
        Next

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()
        cm3.Dispose()
        cm4.Dispose()
    End Sub

    Sub FacturarCFDI_AV(FechaProc As Date)

        Dim t As New ProduccionDS.TraspasosAvioCCDataTable
        Dim nRecibo As Integer
        Dim cRenglon As String
        Dim FechaS As String = FechaProc.ToString("yyyyMMdd")

        Tafolios.Fill(t, FechaS)
        For Each r As ProduccionDS.TraspasosAvioCCRow In t.Rows
            cSerie = "IAV"
            nRecibo = Tafolios.SerieIAV

            Dim stmWriter As New StreamWriter(My.Settings.RutaOrigen & "FACTURA_" & cSerie & "_" & nRecibo & ".txt")

            stmWriter.WriteLine("H1|" & FechaProc.ToShortDateString & "|PPD|99|")

            cRenglon = "H3|" & r.Cliente & "|" & Mid(r.Anexo, 1, 5) & "/" & Mid(r.Anexo, 6, 4) & "|" & cSerie & "|" & nRecibo & "|" & Trim(r.Descr) & "|" &
            Trim(r.Calle) & "|||" & Trim(r.Colonia) & "|" & Trim(r.Delegacion) & "|" & Trim(r.Estado) & "|" & r.Copos & "|||MEXICO|" & Trim(r.RFC) & "|M.N.|" &
            "|FACTURA|" & r.Cliente & "|LEANDRO VALLE 402||REFORMA Y FFCCNN|TOLUCA|ESTADO DE MEXICO|50070|MEXICO|" & r.Anexo & "|" & r.Ciclo & "|"

            cRenglon = cRenglon.Replace("Ñ", Chr(165))
            cRenglon = cRenglon.Replace("ñ", Chr(164))
            cRenglon = cRenglon.Replace("á", Chr(160))
            cRenglon = cRenglon.Replace("é", Chr(130))
            cRenglon = cRenglon.Replace("í", Chr(161))
            cRenglon = cRenglon.Replace("ó", Chr(162))
            cRenglon = cRenglon.Replace("ú", Chr(163))
            cRenglon = cRenglon.Replace("Á", Chr(181))
            cRenglon = cRenglon.Replace("É", Chr(144))
            cRenglon = cRenglon.Replace("Ó", Chr(224))
            cRenglon = cRenglon.Replace("Ú", Chr(233))
            cRenglon = cRenglon.Replace("°", Chr(167))
            stmWriter.WriteLine(cRenglon)


            cRenglon = "D1|" & r.Cliente & "|" & Mid(r.Anexo, 1, 5) & "/" & Mid(r.Anexo, 6, 4) & "|" & cSerie & "|" & nRecibo & "|1|||INTERESES AVIO||" & r.InteresesDias & "|0"
            cRenglon = cRenglon.Replace("Ñ", Chr(165))
            cRenglon = cRenglon.Replace("ñ", Chr(164))
            cRenglon = cRenglon.Replace("á", Chr(160))
            cRenglon = cRenglon.Replace("é", Chr(130))
            cRenglon = cRenglon.Replace("í", Chr(161))
            cRenglon = cRenglon.Replace("ó", Chr(162))
            cRenglon = cRenglon.Replace("ú", Chr(163))
            cRenglon = cRenglon.Replace("Á", Chr(181))
            cRenglon = cRenglon.Replace("É", Chr(144))
            cRenglon = cRenglon.Replace("Ó", Chr(224))
            cRenglon = cRenglon.Replace("Ú", Chr(233))
            cRenglon = cRenglon.Replace("°", Chr(167))
            stmWriter.WriteLine(cRenglon)

            cRenglon = "D1|" & r.Cliente & "|" & Mid(r.Anexo, 1, 5) & "/" & Mid(r.Anexo, 6, 4) & "|" & cSerie & "|" & nRecibo & "|1|||CAPITAL CREDITO DE AVIO||" & r.Importe + r.Fega & "|0"
            cRenglon = cRenglon.Replace("Ñ", Chr(165))
            cRenglon = cRenglon.Replace("ñ", Chr(164))
            cRenglon = cRenglon.Replace("á", Chr(160))
            cRenglon = cRenglon.Replace("é", Chr(130))
            cRenglon = cRenglon.Replace("í", Chr(161))
            cRenglon = cRenglon.Replace("ó", Chr(162))
            cRenglon = cRenglon.Replace("ú", Chr(163))
            cRenglon = cRenglon.Replace("Á", Chr(181))
            cRenglon = cRenglon.Replace("É", Chr(144))
            cRenglon = cRenglon.Replace("Ó", Chr(224))
            cRenglon = cRenglon.Replace("Ú", Chr(233))
            cRenglon = cRenglon.Replace("°", Chr(167))
            stmWriter.WriteLine(cRenglon)
            stmWriter.Close()

            Tafolios.ConsumeSerieIAV()
            Tafolios.FacturarTraspaso(True, cSerie, nRecibo, r.id_Traspaso)
            Tafolios.FacturaIAV(cSerie & nRecibo, r.Anexo, r.Ciclo)
        Next
    End Sub

    Sub LeeFoliosFiscales()
        Dim NombreLOG As String
        Dim Lectura As StreamReader
        Dim Linea As String
        Dim Datos() As String
        Dim contador As Integer
        Dim taFact As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter

        If Directory.Exists(My.Settings.RutaFolios) = False Then
            Directory.CreateDirectory(My.Settings.RutaFolios)
        End If
        For x = 0 To 5
            Console.WriteLine("Leyendo folios Fiscales " & x)
            NombreLOG = "Ejecutor" & Date.Now.AddDays(x * -1).ToString("yyyyMMdd") & ".log"
            If File.Exists(My.Settings.RutaFolios & NombreLOG) Then
                Lectura = New StreamReader(My.Settings.RutaFolios & NombreLOG)
                While Not Lectura.EndOfStream
                    Linea = Lectura.ReadLine
                    Datos = Linea.Split("|")
                    If Datos.Length > 4 Then
                        If IsNumeric(Datos(2)) Then
                            taFact.UpdateGUID(Datos(3), Datos(2), Datos(1))
                            contador += 1
                        End If
                    End If
                End While
                Lectura.Close()
            End If
        Next
        If Not IsNothing(Lectura) Then
            Lectura.Dispose()
        End If

        Dim tAREC As New ProduccionDSTableAdapters.CFDI_RecibosPAGOTableAdapter
        Dim tREC As New ProduccionDS.CFDI_RecibosPAGODataTable
        tAREC.Fill(tREC)
        For Each r As ProduccionDS.CFDI_RecibosPAGORow In tREC.Rows
            taFact.UpdateGUID("Recibo de Pago", r._1_Folio, r._27_Serie_Comprobante)
        Next

    End Sub

    Sub NotificaCANA()
        Dim taFact As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim taMail As New ProduccionDSTableAdapters.GEN_Correos_SistemaFinagilTableAdapter
        Dim dsMail As New ProduccionDS

        Dim D As System.IO.DirectoryInfo
        Dim F As System.IO.FileInfo()

        D = New System.IO.DirectoryInfo(My.Settings.RutaFolios & "\Cancelados\SAR951230N5A\")
        If Not Directory.Exists(D.Name) Then
            Exit Sub
        End If

        F = D.GetFiles("*.xml")
        For i As Integer = 0 To F.Length - 1
            Dim retorno(5) As String
            Dim cadena As StreamReader
            cadena = New StreamReader(My.Settings.RutaFolios & "\Cancelados\SAR951230N5A\" & F(i).Name)
            retorno = Lee_XML_Cancelacion(cadena.ReadToEnd)
            Dim mails() As String = taFact.Obtiene_Mail(retorno(2)).ToString.Split(";")
            cadena.Close()

            Dim rowMail As ProduccionDS.GEN_Correos_SistemaFinagilRow
            rowMail = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()


            For m As Integer = 1 To mails.Length - 1
                rowMail.De = "CFDI@Finagil.com.mx"
                rowMail.Para = mails(m)
                rowMail.Asunto = "Acuse de cancelación SAT CFDI " & retorno(0) & " - " & taFact.Obtiene_Serie(retorno(2)) & "-" & taFact.Obtiene_Folio(retorno(2)) & " UUID " & retorno(2).ToString
                rowMail.Mensaje = Crea_Mensaje(retorno(0), taFact.Obtiene_Serie(retorno(2)), taFact.Obtiene_Folio(retorno(2)), retorno(2), taFact.Obtiene_RFC_Rec(retorno(2)), taFact.Obtiene_RS_Rec(retorno(2)), taFact.Obtiene_FechaEmi(retorno(2)), retorno(1), retorno(3), retorno(5), retorno(4))
                rowMail.Enviado = False
                rowMail.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                rowMail.Attach = ""

                dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail)
                taMail.Update(dsMail.GEN_Correos_SistemaFinagil)
            Next

            File.Copy(F(i).FullName, My.Settings.RutaFolios & "Cancelados\Backup\" & F(i).Name, True)
            File.Delete(F(i).FullName)
        Next
    End Sub

    Sub NotificaCANF()
        Dim taFact As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim taMail As New ProduccionDSTableAdapters.GEN_Correos_SistemaFinagilTableAdapter
        Dim dsMail As New ProduccionDS

        Dim D As System.IO.DirectoryInfo
        Dim F As System.IO.FileInfo()

        D = New System.IO.DirectoryInfo(My.Settings.RutaFolios & "\Cancelados\FIN940905AX7\")
        If Not Directory.Exists(D.Name) Then
            Exit Sub
        End If
        F = D.GetFiles("*.xml")
        For i As Integer = 0 To F.Length - 1
            Dim retorno(5) As String
            Dim cadena As StreamReader
            cadena = New StreamReader(My.Settings.RutaFolios & "\Cancelados\FIN940905AX7\" & F(i).Name)
            retorno = Lee_XML_Cancelacion(cadena.ReadToEnd)
            Dim mails() As String = taFact.Obtiene_Mail(retorno(2)).ToString.Split(";")
            cadena.Close()

            For m As Integer = 1 To mails.Length - 1
                Dim rowMail As ProduccionDS.GEN_Correos_SistemaFinagilRow
                rowMail = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()
                rowMail.De = "CFDI@Finagil.com.mx"
                rowMail.Para = mails(m)
                rowMail.Asunto = "Acuse de cancelación SAT CFDI " & retorno(0) & " - " & taFact.Obtiene_Serie(retorno(2)) & "-" & taFact.Obtiene_Folio(retorno(2)) & " UUID " & retorno(2).ToString
                rowMail.Mensaje = Crea_Mensaje(retorno(0), taFact.Obtiene_Serie(retorno(2)), taFact.Obtiene_Folio(retorno(2)), retorno(2), taFact.Obtiene_RFC_Rec(retorno(2)), taFact.Obtiene_RS_Rec(retorno(2)), taFact.Obtiene_FechaEmi(retorno(2)), retorno(1), retorno(3), retorno(5), retorno(4))
                rowMail.Enviado = False
                rowMail.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                rowMail.Attach = ""

                dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail)
                taMail.Update(dsMail.GEN_Correos_SistemaFinagil)
            Next

            File.Copy(F(i).FullName, My.Settings.RutaFolios & "Cancelados\Backup\" & F(i).Name, True)
            File.Delete(F(i).FullName)
        Next
    End Sub

    Sub SubeFTP()
        Dim D As System.IO.DirectoryInfo
        Dim F As System.IO.FileInfo()
        Dim wrUpload As FtpWebRequest
        Dim btfile() As Byte
        Dim strFile As Stream
        Dim contador As Integer

        If Directory.Exists(My.Settings.RutaFTP) = False Then
            Directory.CreateDirectory(My.Settings.RutaFTP)
        End If

        D = New System.IO.DirectoryInfo(My.Settings.RutaFTP)
        F = D.GetFiles("*.txt")
        For i As Integer = 0 To F.Length - 1

            Console.WriteLine("Subiendo " & F(i).Name)
            wrUpload = DirectCast(WebRequest.Create("ftp://ftplamoderna.ekomercio.com/TXT_Entrada/" & F(i).Name), FtpWebRequest)
            wrUpload.Credentials = New NetworkCredential("lmoderna", "Ekomercio.1")
            wrUpload.Method = WebRequestMethods.Ftp.UploadFile
            btfile = File.ReadAllBytes(My.Settings.RutaFTP & F(i).Name)
            strFile = wrUpload.GetRequestStream()
            strFile.Write(btfile, 0, btfile.Length)
            strFile.Close()
            File.Copy(F(i).FullName, My.Settings.RutaFTP & "Backup\" & F(i).Name, True)
            File.Delete(F(i).FullName)
            contador += 1
        Next
        Console.WriteLine("Subieron: " + contador.ToString + " CFDI txt ")
    End Sub

    Sub SubeWSN()

        Dim taFact As New ProduccionDSTableAdapters.CFDI_Encabezado_NominaTableAdapter
        Dim taMail As New ProduccionDSTableAdapters.GEN_Correos_SistemaFinagilTableAdapter
        Dim taCtrlUUID As New ProduccionDSTableAdapters.CFDI_ControlTimbresTableAdapter
        Dim dsMail As New ProduccionDS

        Dim D As System.IO.DirectoryInfo
        Dim F As System.IO.FileInfo()
        Dim contador As Integer


        If Directory.Exists(My.Settings.RutaNomina) = False Then
            Directory.CreateDirectory(My.Settings.RutaNomina)
        End If

        D = New System.IO.DirectoryInfo(My.Settings.RutaNomina)
        F = D.GetFiles("*.txt")
        For i As Integer = 0 To F.Length - 1

            Console.WriteLine("Subiendo " & F(i).Name)
            Dim cadena As StreamReader
            cadena = New StreamReader(My.Settings.RutaNomina & F(i).Name)
            Dim cadena2 As String = ""
            cadena2 = cadena.ReadToEnd
            Dim serv As WebReferenceNomina_Ek.WSCFDBuilderPlus
            serv = New WebReferenceNomina_Ek.WSCFDBuilderPlus
            Dim resultado As String = ""
            Dim nombre_a() As String = F(i).Name.ToString.Split("_")
            cadena.Close()


            Try
                resultado = serv.procesarTextoPlano("NOMCMO0616", "@NOMCMO0616", nombre_a(1), cadena2)
                If leeXML(resultado, "Valida").ToString = "SinError" Then
                    taFact.UpdateGUID(leeXML(resultado, "UUID"), leeXML(resultado, "Folio"), leeXML(resultado, "Serie"))
                    taCtrlUUID.Insert(leeXML(resultado, "Serie").ToString, leeXML(resultado, "Folio").ToString, leeXML(resultado, "UUID").ToString, leeXML(resultado, "Fecha").ToString, leeXML(resultado, "FechaTimbrado").ToString, leeXML(resultado, "RFCE").ToString, leeXML(resultado, "RFCR").ToString, resultado.ToString)

                ElseIf leeXML(resultado, "Valida").ToString = "Error" Then
                    Dim rowMail As ProduccionDS.GEN_Correos_SistemaFinagilRow
                    rowMail = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                    rowMail.De = "CFDI@Finagil.com.mx"
                    rowMail.Para = "viapolo@finagil.com.mx;denise.gonzalez@finagil.com.mx"
                    rowMail.Asunto = "Error al certificar comprobante de nómina" + F(i).Name
                    If leeXML(resultado, "Err").ToString.Length > 900 Then
                        rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                    Else
                        rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                    End If
                    rowMail.Enviado = False
                    rowMail.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                    rowMail.Attach = ""

                    dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail)
                    taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                    Dim rowMail2 As ProduccionDS.GEN_Correos_SistemaFinagilRow
                    rowMail2 = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                    rowMail2.De = "CFDI@Finagil.com.mx"
                    rowMail2.Para = "ecacerest@finagil.com.mx"
                    rowMail2.Asunto = "Error al certificar comprobante de nómina" + F(i).Name
                    If leeXML(resultado, "Err").ToString.Length > 900 Then
                        rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                    Else
                        rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                    End If
                    rowMail2.Enviado = False
                    rowMail2.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                    rowMail2.Attach = ""

                    dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail2)
                    taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                    File.Copy(F(i).FullName, My.Settings.NoPasa & "Errores\" & F(i).Name, True)
                End If
            Catch ex As Exception
                Dim rowMail As ProduccionDS.GEN_Correos_SistemaFinagilRow
                rowMail = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                rowMail.De = "CFDI@Finagil.com.mx"
                rowMail.Para = "viapolo@finagil.com.mx;denise.gonzalez@finagil.com.mx"
                rowMail.Asunto = "Error al certificar comprobante de nómina" + F(i).Name
                SysLog(ex.ToString, (leeXML(resultado, "Serie").ToString + leeXML(resultado, "Folio").ToString))
                If leeXML(resultado, "Err").ToString.Length > 900 Then
                    rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                    SysLog(leeXML(resultado, "Err").ToString.Substring(0, 900), (leeXML(resultado, "Serie").ToString + leeXML(resultado, "Folio").ToString))
                Else
                    rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                    SysLog(leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1), (leeXML(resultado, "Serie").ToString + leeXML(resultado, "Folio").ToString))
                End If
                rowMail.Enviado = False
                rowMail.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                rowMail.Attach = ""

                dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail)
                taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                Dim rowMail2 As ProduccionDS.GEN_Correos_SistemaFinagilRow
                rowMail2 = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                rowMail2.De = "CFDI@Finagil.com.mx"
                rowMail2.Para = "ecacerest@finagil.com.mx"
                rowMail2.Asunto = "Error al certificar comprobante de nómina" + F(i).Name
                If leeXML(resultado, "Err").ToString.Length > 900 Then
                    rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                Else
                    rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                End If
                rowMail2.Enviado = False
                rowMail2.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                rowMail2.Attach = ""

                dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail2)
                taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                File.Copy(F(i).FullName, My.Settings.NoPasa & "Errores\" & F(i).Name, True)
            End Try
            File.Copy(F(i).FullName, My.Settings.RutaNomina & "Backup\" & F(i).Name, True)
            File.Delete(F(i).FullName)
            contador += 1
        Next
        Console.WriteLine("Subieron: " + contador.ToString + " CFDI txt ")


        '****Try
        'resultado = serv.procesarTextoPlano("ekomercio", "aserri", nombre_a(1), encode(cadena2))
        'resultado = serv.procesarTextoPlano("ekomercio", "aserri", "FGM790801SD7", cadena2)
        'resultado = serv.procesarTextoPlano("ekomercio", "aserri", nombre_a(1), cadena2)
        'Catch e As Exception
        'End Try
        'Next
    End Sub

    Public Function encode(ByVal str As String)
        Dim utf8Encoding As New System.Text.UTF8Encoding
        Dim encodedString() As Byte

        encodedString = utf8Encoding.GetBytes(str)

        Return encodedString.ToString()
    End Function

    Sub SubeWS()
        Dim taFact As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim taMail As New ProduccionDSTableAdapters.GEN_Correos_SistemaFinagilTableAdapter
        Dim taCtrlUUID As New ProduccionDSTableAdapters.CFDI_ControlTimbresTableAdapter
        Dim taInfo100n As New Info100DSTableAdapters.CFDI_EXTERNOTableAdapter
        Dim dsMail As New ProduccionDS

        Dim D As System.IO.DirectoryInfo
        Dim F As System.IO.FileInfo()
        Dim contador As Integer


        If Directory.Exists(My.Settings.RutaFTP) = False Then
            Directory.CreateDirectory(My.Settings.RutaFTP)
        End If

        D = New System.IO.DirectoryInfo(My.Settings.RutaFTP)
        F = D.GetFiles("*.txt")
        For i As Integer = 0 To F.Length - 1

            Console.WriteLine("Subiendo " & F(i).Name)
            Dim cadena As StreamReader
            cadena = New StreamReader(My.Settings.RutaFTP & F(i).Name)
            Dim cadena2 As String = ""
            cadena2 = cadena.ReadToEnd
            Dim serv As WebReference_Ek.WSCFDBuilderPlus
            serv = New WebReference_Ek.WSCFDBuilderPlus
            Dim resultado As String = ""
            Dim nombre_a() As String = F(i).Name.ToString.Split("_")
            Dim time_inicio, time_final As DateTime
            Dim file_log As New System.IO.FileStream(My.Settings.RutaFTP & "eventos_ws.log", FileMode.Append, FileAccess.Write)
            Dim objStream As New StreamWriter(file_log)
            cadena.Close()

            Try
                time_inicio = Date.Now.ToLongTimeString
                resultado = serv.procesarTextoPlano("CFDICMO0617", "@CFDICMO0617", nombre_a(1), cadena2)
                time_final = Date.Now.ToLongTimeString

                If leeXML(resultado, "Valida").ToString = "SinError" Then
                    taFact.UpdateGUID(leeXML(resultado, "UUID"), leeXML(resultado, "Folio"), leeXML(resultado, "Serie"))
                    If leeXML(resultado, "Serie") = "F" Then
                        taInfo100n.ActualizaUUID_UpdateQuery(leeXML(resultado, "UUID"), leeXML(resultado, "Serie"), leeXML(resultado, "Folio"))
                    End If
                    taCtrlUUID.Insert(leeXML(resultado, "Serie").ToString, leeXML(resultado, "Folio").ToString, leeXML(resultado, "UUID").ToString, leeXML(resultado, "Fecha").ToString, leeXML(resultado, "FechaTimbrado").ToString, leeXML(resultado, "RFCE").ToString, leeXML(resultado, "RFCR").ToString, resultado.ToString)
                    If leeXML(resultado, "RFCE").ToString = "FIN940905AX7" Then
                        Dim fecha As String = leeXML(resultado, "Fecha")
                        Dim anio As Integer = CDate(fecha).Year
                        Dim mes As Integer = CDate(fecha).Month
                        Dim dia As Integer = CDate(fecha).Day
                        Dim files As StreamWriter = Nothing

                        If Directory.Exists(My.Settings.RutaXML & anio & "\" & mes & "\" & dia) Then
                            files = New StreamWriter(My.Settings.RutaXML & anio & "\" & mes & "\" & dia & "\" & leeXML(resultado, "Folio") & "-" & leeXML(resultado, "Serie") & "-" & leeXML(resultado, "UUID") & ".xml", False, Encoding.UTF8)
                            files.Write(resultado)
                            files.Close()
                        Else
                            System.IO.Directory.CreateDirectory(My.Settings.RutaXML & anio & "\" & mes & "\" & dia)
                            files = New StreamWriter(My.Settings.RutaXML & anio & "\" & mes & "\" & dia & "\" & leeXML(resultado, "Folio") & "-" & leeXML(resultado, "Serie") & "-" & leeXML(resultado, "UUID") & ".xml", False, Encoding.UTF8)
                            files.Write(resultado)
                            files.Close()
                        End If
                    End If

                    objStream.WriteLine("Archivo procesado: " & F(i).Name & " Inicio: " & time_inicio.ToString & " Fin: " & time_final.ToString)
                    objStream.Close()
                    file_log.Close()

                ElseIf leeXML(resultado, "Valida").ToString = "Error" Then
                    Dim rowMail As ProduccionDS.GEN_Correos_SistemaFinagilRow
                    rowMail = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                    rowMail.De = "CFDI@Finagil.com.mx"
                    rowMail.Para = "viapolo@finagil.com.mx;denise.gonzalez@finagil.com.mx"
                    rowMail.Asunto = "Error al certificar comprobante" + F(i).Name
                    If leeXML(resultado, "Err").ToString.Length > 900 Then
                        rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                    Else
                        rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                    End If
                    rowMail.Enviado = False
                    rowMail.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                    rowMail.Attach = ""

                    dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail)
                    taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                    Dim rowMail2 As ProduccionDS.GEN_Correos_SistemaFinagilRow
                    rowMail2 = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                    rowMail2.De = "CFDI@Finagil.com.mx"
                    rowMail2.Para = "ecacerest@finagil.com.mx"
                    rowMail2.Asunto = "Error al certificar comprobante" + F(i).Name
                    If leeXML(resultado, "Err").ToString.Length > 900 Then
                        rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                    Else
                        rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                    End If
                    rowMail2.Enviado = False
                    rowMail2.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                    rowMail2.Attach = ""

                    dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail2)
                    taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                    objStream.WriteLine("Archivo procesado: " & F(i).Name & " Inicio: " & time_inicio.ToString & " Fin: " & time_final.ToString & " Error: " & resultado)
                    objStream.Close()
                    file_log.Close()

                    File.Copy(F(i).FullName, My.Settings.NoPasa & "Errores\" & F(i).Name, True)
                End If
            Catch ex As Exception
                Dim rowMail As ProduccionDS.GEN_Correos_SistemaFinagilRow
                rowMail = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                rowMail.De = "CFDI@Finagil.com.mx"
                rowMail.Para = "viapolo@finagil.com.mx;denise.gonzalez@finagil.com.mx"
                rowMail.Asunto = "Error al certificar comprobante" + F(i).Name
                SysLog(ex.ToString, (leeXML(resultado, "Serie").ToString + leeXML(resultado, "Folio").ToString))
                If leeXML(resultado, "Err").ToString.Length > 900 Then
                    rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                    SysLog(leeXML(resultado, "Err").ToString.Substring(0, 900), (leeXML(resultado, "Serie").ToString + leeXML(resultado, "Folio").ToString))
                Else
                    rowMail.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                    SysLog(leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1), (leeXML(resultado, "Serie").ToString + leeXML(resultado, "Folio").ToString))
                End If
                rowMail.Enviado = False
                rowMail.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                rowMail.Attach = ""

                dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail)
                taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                Dim rowMail2 As ProduccionDS.GEN_Correos_SistemaFinagilRow
                rowMail2 = dsMail.GEN_Correos_SistemaFinagil.NewGEN_Correos_SistemaFinagilRow()

                rowMail2.De = "CFDI@Finagil.com.mx"
                rowMail2.Para = "ecacerest@finagil.com.mx"
                rowMail2.Asunto = "Error al certificar comprobante" + F(i).Name
                If leeXML(resultado, "Err").ToString.Length > 900 Then
                    rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, 900)
                Else
                    rowMail2.Mensaje = leeXML(resultado, "Err").ToString.Substring(0, leeXML(resultado, "Err").ToString.Length - 1)
                End If
                rowMail2.Enviado = False
                rowMail2.fecha = Date.Now.Date.ToString("yyyy-MM-dd hh:mm:ss.fff")
                rowMail2.Attach = ""

                dsMail.GEN_Correos_SistemaFinagil.Rows.Add(rowMail2)
                taMail.Update(dsMail.GEN_Correos_SistemaFinagil)

                File.Copy(F(i).FullName, My.Settings.NoPasa & "Errores\" & F(i).Name, True)
            End Try
            File.Copy(F(i).FullName, My.Settings.RutaFTP & "Backup\" & F(i).Name, True)
            File.Delete(F(i).FullName)
            contador += 1
        Next
        Console.WriteLine("Subieron: " + contador.ToString + " CFDI txt ")
    End Sub

    Private Function SysLog(TextLogP As String, name As String)
        Dim LogFile As String = "\" + DateTime.Now.ToString("dd-MM-yyyy") + "-" + DateTime.Now.ToString("hhmmss") + name + ".log"
        Using outputFile As New StreamWriter(My.Settings.NoPasa & "Errores\" & Convert.ToString(LogFile), True)
            outputFile.WriteLine(DateTime.Now.ToString("dd/MM/yyyy") + " " + DateTime.Now.ToString("hh:mm:ss") + " - " + TextLogP)
        End Using
        Return False
    End Function

    Function Crea_Mensaje(RFC_Emisor As String, serie As String, folio As String, UUIDG As String, Receptor As String, RSocial As String, FechaEmision As String, FechaCancelacion As String, Estatus_UUID As String, DigestValue As String, SignatureValue As String)
        Dim retorno_mensaje As String = ""
        Try
            retorno_mensaje = "<font size=5 face=" + Chr(34) + "Arial" + Chr(34) + ">Acuse de cancelaci&oacute;n SAT... " + vbNewLine + "RFC Emisor: " + RFC_Emisor +
                                    "<br><br/>" +
                                    "<table  align=" + Chr(34) + "center" + Chr(34) + " border=1 cellspacing=0 cellpadding=2>" +
                                        "<tr>" +
                                            "<th scope=" + Chr(34) + "col" + Chr(34) + "> - </th>" +
                                            "<th scope=" + Chr(34) + "col" + Chr(34) + ">Descripci&oacute;n</th>" +
                                        "<tr>" +
                                            "<td>Serie: </td>" +
                                            "<td>" + serie + "</td>" +
                                        "<tr>" +
                                            "<td>Folio: </td>" +
                                            "<td>" + folio + "</td>" +
                                        "<tr>" +
                                            "<td>Folio Fiscal: </td>" +
                                            "<td>" + UUIDG + "</td>" +
                                        "<tr>" +
                                            "<td>RFC Receptor: </td>" +
                                            "<td>" + Receptor + "</td>" +
                                        "<tr>" +
                                            "<td>Razón Social: </td>" +
                                            "<td>" + RSocial + "</td>" +
                                        "<tr>" +
                                            "<td>Fecha de Emisi&oacute;n: </td>" +
                                            "<td>" + FechaEmision + "</td>" +
                                        "<tr>" +
                                            "<td>Fecha de Cancelaci&oacute;n: </td>" +
                                            "<td>" + FechaCancelacion + "</td>" +
                                        "<tr>" +
                                            "<td>Estatus de Cancelaci&oacute;n: </td>" +
                                            "<td>" + Estatus_UUID + "</td>" +
                                        "<tr>" +
                                            "<td>Digest Value: </td>" +
                                            "<td>" + DigestValue + "</td>" +
                                        "<tr>" +
                                            "<td>SignatureValue: </td>" +
                                            "<td>" + SignatureValue + "</td>" +
                                        "<tr>" +
                                    "</table>" +
                                    "</font>"
        Catch
        End Try
        Return retorno_mensaje
    End Function

    Function Lee_XML_Cancelacion(ByVal docXmlCAN As String)
        Dim m_xmld As XmlDocument
        Dim m_nodelist As XmlNodeList
        Dim m_node As XmlNode
        Dim m_attn_d As XmlAttribute
        Dim m_node_b As XmlNode
        Dim m_node_c As XmlNode
        Dim m_node_d As XmlNode
        Dim m_node_e As XmlNode
        Dim m_node_f As XmlNode
        Dim retorno(5) As String


        m_xmld = New XmlDataDocument
        m_xmld.LoadXml(docXmlCAN)


        m_nodelist = m_xmld.SelectNodes("/Acuse")
        For Each m_node In m_nodelist
            For Each m_attn_d In m_node.Attributes
                Select Case m_attn_d.Name
                    Case "RfcEmisor"
                        'RFC_Emisor
                        retorno(0) = m_attn_d.Value
                    Case "Fecha"
                        'FechaCancelacion
                        retorno(1) = m_attn_d.Value
                End Select
            Next

            For Each m_node_b In m_node.ChildNodes
                For Each m_node_c In m_node_b.ChildNodes
                    Select Case m_node_c.Name
                        Case "UUID"
                            'UUID
                            retorno(2) = m_node_c.InnerText
                        Case "EstatusUUID"
                            'Estatus_UUID
                            retorno(3) = m_node_c.InnerText
                        Case "SignatureValue"
                            'SignatureValue
                            retorno(4) = m_node_c.InnerText
                    End Select
                Next
                If m_node_b.Name = "Signature" Then
                    For Each m_node_d In m_node_b.ChildNodes
                        If m_node_d.Name = "SignedInfo" Then
                            For Each m_node_e In m_node_d.ChildNodes
                                If m_node_e.Name = "Reference" Then
                                    For Each m_node_f In m_node_e.ChildNodes
                                        Select Case m_node_f.Name
                                            Case "DigestValue"
                                                'DigestValue
                                                retorno(5) = m_node_f.InnerText
                                        End Select
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Next
        Return retorno
    End Function

    Public Function leeXML(docXML As String, nodo As String)
        Dim doc As XmlDataDocument
        doc = New XmlDataDocument
        doc.LoadXml(docXML)
        Dim CFDI As XmlNode
        Dim retorno As String = ""

        If nodo = "Valida" Then
            If doc.LastChild.Name = "Error" Then
                retorno = "Error"
                Return retorno
                Exit Function
            Else
                retorno = "SinError"
                Return retorno
                Exit Function
            End If
        End If

        CFDI = doc.DocumentElement

        If nodo = "Err" Then
            For Each Err As XmlNode In CFDI.ChildNodes
                If Err.Name = "ErrorMessage" And nodo = "Err" Then
                    retorno = Err.InnerText
                    Return retorno
                    Exit Function
                End If
            Next
        End If

        If nodo = "UUID" Then
            For Each ComprobanteP1 As XmlNode In CFDI.ChildNodes
                If ComprobanteP1.Name = "XMLTimbrado" Then
                    For Each ComprobanteP2 As XmlNode In ComprobanteP1.ChildNodes
                        If ComprobanteP2.Name = "cfdi:Comprobante" Then
                            For Each Comprobante As XmlNode In ComprobanteP2.ChildNodes
                                If Comprobante.Name = "cfdi:Complemento" And nodo = "UUID" Then
                                    For Each Complemento As XmlNode In Comprobante.ChildNodes
                                        If Complemento.Name = "tfd:TimbreFiscalDigital" Then
                                            For Each TimbreFiscalDigital As XmlNode In Complemento.Attributes
                                                If TimbreFiscalDigital.Name = "UUID" Then
                                                    retorno = TimbreFiscalDigital.Value.ToString
                                                    Return retorno
                                                    Exit Function
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If


                If ComprobanteP1.Name = "cfdi:Complemento" And nodo = "UUID" Then
                    For Each Complemento As XmlNode In ComprobanteP1.ChildNodes
                        If Complemento.Name = "tfd:TimbreFiscalDigital" Then
                            For Each TimbreFiscalDigital As XmlNode In Complemento.Attributes
                                If TimbreFiscalDigital.Name = "UUID" Then
                                    retorno = TimbreFiscalDigital.Value.ToString
                                    Return retorno
                                    Exit Function
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If


        If nodo = "RFCE" Or nodo = "RFCR" Then
            For Each ComprobanteP1 As XmlNode In CFDI.ChildNodes
                If ComprobanteP1.Name = "XMLTimbrado" Then
                    For Each ComprobanteP2 As XmlNode In ComprobanteP1.ChildNodes
                        If ComprobanteP2.Name = "cfdi:Comprobante" Then
                            For Each Comprobante As XmlNode In ComprobanteP2.ChildNodes
                                If Comprobante.Name = "cfdi:Emisor" And nodo = "RFCE" Then
                                    For Each Emisor As XmlNode In Comprobante.Attributes
                                        'For Each Atributos As XmlNode In Emisor.Attributes
                                        If Emisor.Name = "Rfc" Then
                                            retorno = Emisor.Value.ToString
                                            Return retorno
                                            Exit Function
                                        End If
                                        'Next
                                    Next
                                End If
                                If Comprobante.Name = "cfdi:Receptor" And nodo = "RFCR" Then
                                    For Each Emisor As XmlNode In Comprobante.Attributes
                                        'For Each Atributos As XmlNode In Emisor.Attributes
                                        If Emisor.Name = "Rfc" Then
                                            retorno = Emisor.Value.ToString
                                            Return retorno
                                            Exit Function
                                        End If
                                        'Next
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If

                If ComprobanteP1.Name = "cfdi:Emisor" And nodo = "RFCE" Then
                    For Each Emisor As XmlNode In ComprobanteP1.Attributes
                        'For Each Atributos As XmlNode In Emisor.Attributes
                        If Emisor.Name = "Rfc" Then
                            retorno = Emisor.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                        'Next
                    Next
                End If
                If ComprobanteP1.Name = "cfdi:Receptor" And nodo = "RFCR" Then
                    For Each Emisor As XmlNode In ComprobanteP1.Attributes
                        'For Each Atributos As XmlNode In Emisor.Attributes
                        If Emisor.Name = "Rfc" Then
                            retorno = Emisor.Value.ToString
                            Return retorno
                            Exit Function
                        End If
                        'Next
                    Next
                End If
            Next
        End If

        If nodo = "FechaTimbrado" Then
            For Each ComprobanteP1 As XmlNode In CFDI.ChildNodes
                If ComprobanteP1.Name = "XMLTimbrado" Then
                    For Each ComprobanteP2 As XmlNode In ComprobanteP1.ChildNodes
                        If ComprobanteP2.Name = "cfdi:Comprobante" Then
                            For Each Comprobante As XmlNode In ComprobanteP2.ChildNodes
                                If Comprobante.Name = "cfdi:Complemento" And nodo = "FechaTimbrado" Then
                                    For Each Complemento As XmlNode In Comprobante.ChildNodes
                                        If Complemento.Name = "tfd:TimbreFiscalDigital" Then
                                            For Each TimbreFiscalDigital As XmlNode In Complemento.Attributes
                                                If TimbreFiscalDigital.Name = "FechaTimbrado" Then
                                                    retorno = TimbreFiscalDigital.Value.ToString
                                                    Return retorno
                                                    Exit Function
                                                End If
                                            Next
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If

                If ComprobanteP1.Name = "cfdi:Complemento" And nodo = "FechaTimbrado" Then
                    For Each Complemento As XmlNode In ComprobanteP1.ChildNodes
                        If Complemento.Name = "tfd:TimbreFiscalDigital" Then
                            For Each TimbreFiscalDigital As XmlNode In Complemento.Attributes
                                If TimbreFiscalDigital.Name = "FechaTimbrado" Then
                                    retorno = TimbreFiscalDigital.Value.ToString
                                    Return retorno
                                    Exit Function
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If


        For Each Comprobante As XmlNode In CFDI.Attributes
            If Comprobante.Name = "Moneda" And nodo = "Moneda" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            ElseIf Comprobante.Name = "TipoCambio" And nodo = "TipoCambio" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            ElseIf (Comprobante.Name = "Total" Or Comprobante.Name = "total") And nodo = "Total" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            ElseIf (Comprobante.Name = "MetodoPago" Or Comprobante.Name = "metodoDePago") And nodo = "MetodoPago" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            ElseIf (Comprobante.Name = "Serie" Or Comprobante.Name = "serie") And nodo = "Serie" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            ElseIf (Comprobante.Name = "Folio" Or Comprobante.Name = "folio") And nodo = "Folio" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            ElseIf Comprobante.Name = "Fecha" And nodo = "Fecha" Then
                retorno = Comprobante.Value.ToString
                Return retorno
                Exit Function
            End If
        Next

        For Each ComprobanteP1 As XmlNode In CFDI.ChildNodes
            If ComprobanteP1.Name = "XMLTimbrado" Then
                For Each ComprobanteP2 As XmlNode In ComprobanteP1.ChildNodes
                    If ComprobanteP2.Name = "cfdi:Comprobante" Then
                        For Each Comprobante As XmlNode In ComprobanteP2.Attributes
                            If Comprobante.Name = "Moneda" And nodo = "Moneda" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            ElseIf Comprobante.Name = "TipoCambio" And nodo = "TipoCambio" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            ElseIf (Comprobante.Name = "Total" Or Comprobante.Name = "total") And nodo = "Total" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            ElseIf (Comprobante.Name = "MetodoPago" Or Comprobante.Name = "metodoDePago") And nodo = "MetodoPago" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            ElseIf (Comprobante.Name = "Serie" Or Comprobante.Name = "serie") And nodo = "Serie" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            ElseIf (Comprobante.Name = "Folio" Or Comprobante.Name = "folio") And nodo = "Folio" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            ElseIf Comprobante.Name = "Fecha" And nodo = "Fecha" Then
                                retorno = Comprobante.Value.ToString
                                Return retorno
                                Exit Function
                            End If
                        Next
                    End If
                Next
            End If


        Next
    End Function

    Sub GeneraRNominaekomercio()
        Dim Encabezado As ProduccionDS.CFDI_Encabezado_NominaRow
        Dim Complemento As ProduccionDS.CFDI_Complemento_NominaRow

        Dim col As DataColumn
        Dim f As StreamWriter

        Dim taEncabezado As New ProduccionDSTableAdapters.CFDI_Encabezado_NominaTableAdapter
        Dim taComplemento As New ProduccionDSTableAdapters.CFDI_Complemento_NominaTableAdapter

        taEncabezado.Facturas_No_Procesadas_FillBy(Production_AUXDataSet.CFDI_Encabezado_Nomina)

        For Each Encabezado In Production_AUXDataSet.CFDI_Encabezado_Nomina.Rows
            Dim Cad_Nom As String = "~"
            Dim subtotal As Double = 0
            Dim descuento As Double = 0
            f = New StreamWriter(My.Settings.RutaNomina & "eKomercio_" & Encabezado._3_RFC_Emisor & "_" & Encabezado._27_Serie_Comprobante & Encabezado._1_Folio & ".txt", False)
            For Each col In Production_AUXDataSet.CFDI_Encabezado_Nomina.Columns
                If col.ColumnName <> "Encabezado_Procesado" And col.ColumnName <> "Fecha" Then

                    If col.ColumnName = "31_Hora" Then
                        'Encabezado(col) = "[Hora]"
                    ElseIf col.ColumnName = "55_Monto_IVA" Then
                        Encabezado(col) = "0"
                    ElseIf col.ColumnName = "144_Misc32" Then
                        Encabezado(col) = "P01"
                    End If

                    If col.ColumnName <> "193_Monto_TotalImp_Trasladados" Then
                        If col.ColumnName <> "Guid" Then
                            Cad_Nom += Encabezado(col).ToString + "|"
                        End If
                    Else
                        If col.ColumnName <> "Guid" Then
                            Cad_Nom += Encabezado(col).ToString
                        End If
                    End If
                    If col.ColumnName = "54_Monto_SubTotal" Then
                        subtotal = Encabezado(col).ToString
                    ElseIf col.ColumnName = "89_Monto_Descuento" Then
                        descuento = Encabezado(col).ToString
                    End If
                End If
            Next
            'Cad_Nom = Cad_Nom + "[@@CRLF]¬Pago de nómina|1|ACT|" + subtotal.ToString + "|" + subtotal.ToString + "|||||||||||84111505|||" + descuento.ToString + "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||[@@CRLF]"
            Cad_Nom = Cad_Nom + "[@@CRLF]¬Pago de nómina|1|ACT|" + subtotal.ToString + "|" + subtotal.ToString + "|||||||||||84111505|||" + descuento.ToString + "||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||[@@CRLF]"
            taComplemento.Obtien_FAC_FillBy(Production_AUXDataSet.CFDI_Complemento_Nomina, Encabezado._27_Serie_Comprobante, Encabezado._1_Folio)
            Dim cont As Integer = 0
            For Each Complemento In Production_AUXDataSet.CFDI_Complemento_Nomina.Rows
                cont += 1
                For Each col In Production_AUXDataSet.CFDI_Complemento_Nomina.Columns
                    If col.ColumnName <> "id" And col.ColumnName <> "serie" And col.ColumnName <> "folio" Then
                        If col.ColumnName <> "Comp_18" Then
                            Cad_Nom += Complemento(col).ToString + "|"
                        Else
                            Cad_Nom += Complemento(col).ToString
                        End If
                    End If
                Next

                If Production_AUXDataSet.CFDI_Complemento_Nomina.Rows.Count <> cont Then
                    Cad_Nom = Cad_Nom + "[@@CRLF]"
                Else
                    Cad_Nom = Cad_Nom
                End If
            Next
            taEncabezado.PrcesadoNomina(True, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante)
            f.WriteLine(Cad_Nom)
            f.Close()
        Next
    End Sub

    Sub GeneraFacturaEkomercio()
        Dim Cad As String = "~"
        Dim Cad_Retencion As String = ""
        Dim TotalImpuesto16 As Decimal = 0.0
        Dim TotalImpuesto0 As Decimal = 0
        Dim TotalImpuestoEXE As Decimal = 0
        Dim Encabezado As ProduccionDS.CFDI_EncabezadoRow
        Dim Detalle As ProduccionDS.CFDI_DetalleRow
        Dim Imp_Adic As ProduccionDS.CFDI_Impuestos_AdicionalesRow
        Dim taAnexo As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
        Dim f As StreamWriter
        Dim Col As DataColumn
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim cfilas As Integer = 0
        Dim cexento As Integer = 0
        Dim ctasa As Integer = 0
        Dim cpcero As Integer = 0
        Dim contador As Integer = 0
        Dim vTipoImpuesto As String = ""
        Dim vExento As String = ""
        Dim vLimpia As String = ""
        Dim cad_imp_adic_total As String = ""
        Dim Tasa_RET As String = ""
        Dim T16 As Decimal = 0
        Dim T00 As Decimal = 0
        Dim T08 As Decimal = 0
        Dim bT16 As String = "NA"
        Dim bT00 As String = "NA"
        Dim bT08 As String = "NA"

        CFDI_EncabezadoTableAdapter.FillByNoProcesadosFACT(Production_AUXDataSet.CFDI_Encabezado) 'LLENO ENCABEZADO

        ' Recorrido de Renglones de Tabla Encabezado 
        For Each Encabezado In Production_AUXDataSet.CFDI_Encabezado.Rows() 'RECORRO FACTURAS SIN PROCESAR
            T00 = 0
            T16 = 0
            T08 = 0
            bT00 = "NA"
            bT08 = "NA"
            bT16 = "NA"
            CFDI_DetalleTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_Detalle, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE

            If Production_AUXDataSet.CFDI_Detalle.Rows.Count > 0 Then
                f = New StreamWriter(My.Settings.RutaFTP & "eKomercio_" & Encabezado._3_RFC_Emisor & "_" & Encabezado._27_Serie_Comprobante & Encabezado._1_Folio & ".txt", False)
                If CDate(Encabezado._30_Fecha).Date = Date.Now.Date Then
                    Encabezado._31_Hora = Date.Now.ToString("HH:mm:ss")
                ElseIf CDate(Encabezado._30_Fecha) < Date.Now.Date Then
                    Encabezado._31_Hora = Date.Now.AddHours(2).ToString("HH:mm:ss")
                End If

                If Not IsNothing(Encabezado._157_Misc45) Then
                    If Encabezado._157_Misc45.ToString <> "" Then
                        'Encabezado._155_Misc43 = "[Addenda_Finagil]"
                    End If
                End If
                'Formato para activo fijo
                If Encabezado._3_RFC_Emisor = "FIN940905AX7" And Encabezado._27_Serie_Comprobante = "B" Then
                    Encabezado._113_Misc01 = "[AFIN]"
                End If
                'Cambia CP por incetivo del 8% Zona fronteriza
                If Encabezado._114_Misc02 <> "" Then
                    'MsgBox(Encabezado._114_Misc02.Substring(0, 10))
                    'MsgBox(taAnexo.ObtieneImpEst8pc_ScalarQuery(Encabezado._114_Misc02.Substring(0, 10)).ToString)
                    If taAnexo.ObtieneImpEst8pc_ScalarQuery(Encabezado._114_Misc02.Substring(0, 10)) = "8.0000" Then
                        Encabezado._180_LugarExpedicion = "21384"
                    End If
                End If

                Cad = "~"
                i += 1

                ' Recorrido de Columnas o Campos de Tabla Encabezado 
                Dim val As String = "SR"
                For Each Col In Production_AUXDataSet.CFDI_Encabezado.Columns ' CONCATENO EL RENGLON DEL ENCABEZADO

                    If Col.ColumnName = "192_Monto_TotalImp_Retenidos" Then
                        If IsNothing(Encabezado(Col)) Then
                            val = "CR"
                        Else
                            Cad += "*"
                        End If
                    End If

                    If Col.ColumnName <> "Encabezado_Procesado" And Col.ColumnName <> "Fecha" Then
                        If Col.ColumnName <> "193_Monto_TotalImp_Trasladados" Then
                            ' 25 Octubre 2017
                            ' 6 de Noviembre se Agrego un Campo
                            If Col.ColumnName <> "Guid" Then
                                vLimpia = Encabezado(Col).ToString   ' Para quitar Salto de linea 25Octubre2017
                                Cad += vLimpia.Replace(vbCrLf, " ") & "|"   ' Para quitar Salto de linea 25Octubre2017
                                ' Cad += Encabezado(Col) & "|"     '   LINEA ORIGINAL SIN PIMPIAR 
                            End If
                        Else
                            If val = "SR" Then
                                If Not IsNothing(Encabezado(Col)) Then
                                    TotalImpuesto16 = Encabezado(Col)
                                    If CDec(Encabezado(Col) = -1) Then
                                        Cad += ""
                                    Else
                                        Cad += Encabezado(Col).ToString
                                    End If
                                Else
                                    Cad += ""
                                End If
                            Else
                                'TotalImpuesto16 = Encabezado(Col)
                                Cad += ""
                            End If
                        End If
                        j += 1
                    End If
                Next
                f.WriteLine(Cad.Replace("*0.0000", "").Replace("|*|", "||").Replace("|*", "|"))

                Cad = "¬" ' PREPARO PARA DETALLES
                CFDI_DetalleTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_Detalle, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE
                cfilas += 1
                ctasa = 0
                cpcero = 0
                cexento = 0
                Tasa_RET = ""
                For Each Detalle In Production_AUXDataSet.CFDI_Detalle.Rows 'RECORRO DETALLE DE LA FACTURA EN CUESTION

                    ' Impuestos adicionales
                    Dim taImpAdic As New ProduccionDSTableAdapters.CFDI_Impuestos_AdicionalesTableAdapter
                    taImpAdic.Obt_Imp_Adic_FillBy(Production_AUXDataSet.CFDI_Impuestos_Adicionales, Encabezado._27_Serie_Comprobante, Encabezado._1_Folio)
                    Dim cad_imp_adic As String

                    If Production_AUXDataSet.CFDI_Impuestos_Adicionales.Rows.Count > 0 Then
                        For Each Imp_Adic In Production_AUXDataSet.CFDI_Impuestos_Adicionales.Rows
                            cad_imp_adic = "\" & Imp_Adic._1_Impuesto_TipoImpuesto & "|" & Imp_Adic._2_Impuesto_Descripcion & "|" & Imp_Adic._3_Impuesto_Monto_base & "|" & Imp_Adic._4_Impuesto_Monto_Impuesto & "|" & Imp_Adic._5_Impuesto_Clave & "|" & Imp_Adic._6_Impuesto_Tasa & "|" & Imp_Adic._7_Impuesto_Porcentaje.Trim & "||"
                            cad_imp_adic_total = "RE|" & Imp_Adic._5_Impuesto_Clave & "|" & Imp_Adic._4_Impuesto_Monto_Impuesto & "||"
                        Next
                    End If

                    Dim val_ImpTraslados As String = "NA"

                    For Each Col In Production_AUXDataSet.CFDI_Detalle.Columns ' CONCATENO EL RENGLON DE DETALLE CON IMPUESTOS

                        If Col.ColumnName = "1_Impuesto_TipoImpuesto" Then
                            If Detalle.Item("6_Impuesto_Tasa") = "No Objeto" Then
                                Exit For
                            End If
                            Cad += "\Impuesto|" ' DIVIDO SESSION DE IMPUESTOS EN DETALLE
                            If Detalle(Col) = "EXE" Then
                                vTipoImpuesto = "EXE"
                            End If
                        Else
                            If Col.ColumnName <> "Detalle_Folio" And Col.ColumnName <> "Detalle_Serie" And Col.ColumnName <> "id_Detalle" Then
                                If Col.ColumnName <> "99_Linea_NoIdentificacion" Then
                                    If Col.ColumnName = "4_Impuesto_Monto_Impuesto" Then
                                        If vTipoImpuesto = "EXE" Then
                                            Cad += "|"
                                        Else
                                            Cad += Detalle(Col).ToString & "|"
                                        End If
                                    ElseIf Col.ColumnName = "8_Retencion_Tasa" Then
                                        If Not Detalle.Is_8_Retencion_TasaNull Then ' Retenciones de IVA
                                            If Detalle._10_Retencion_Monto > 0 Then
                                                Cad_Retencion = "\Impuesto|RE|" & Detalle._9_Retencion_Monto_Base & "|" & Detalle._10_Retencion_Monto & "|" & Detalle._5_Impuesto_Clave & "|Tasa|" & Detalle._8_Retencion_Tasa & "||"
                                                Tasa_RET = Detalle._8_Retencion_Tasa
                                            Else
                                                Cad_Retencion = ""
                                            End If
                                        End If
                                    ElseIf Col.ColumnName = "9_Retencion_Monto_Base" Then
                                        'campo no considerado en el layout de salida
                                    ElseIf Col.ColumnName = "10_Retencion_Monto" Then
                                        'campo no considerado en el layout de salida
                                    Else
                                        ' 21 Noviembre
                                        If Col.ColumnName = "6_Impuesto_Tasa" Then
                                            If Detalle(Col).ToString = "Exento" Then
                                                val_ImpTraslados = "NA"
                                                cexento += 1
                                            End If
                                            If Detalle(Col).ToString = "Tasa" Then
                                                ctasa += 1
                                            End If
                                        Else
                                            If Col.ColumnName = "7_Impuesto_Porcentaje" Then
                                                If Detalle(Col).ToString = "0.0000" Or Detalle(Col).ToString = "0" Then
                                                    cpcero += 1
                                                    T00 += CDec(Detalle._4_Impuesto_Monto_Impuesto)
                                                    bT00 = "SA"
                                                ElseIf Detalle(Col).ToString = "0.1600" Or Detalle(Col).ToString = "0.16" Or Detalle(Col).ToString = "0.160000" Then
                                                    T16 += CDec(Detalle._4_Impuesto_Monto_Impuesto)
                                                    bT16 = "SA"
                                                ElseIf Detalle(Col).ToString = "0.0800" Then
                                                    T08 += CDec(Detalle._4_Impuesto_Monto_Impuesto)
                                                    bT08 = "SA"
                                                End If
                                            End If
                                        End If
                                        ' 25 Octubre 2017
                                        vLimpia = ""
                                        vLimpia = Detalle(Col).ToString   ' Para quitar Salto de linea 25Octubre2017
                                        Cad += vLimpia.Replace(vbCrLf, " ") & "|"   ' Para quitar Salto de linea 25Octubre2017
                                    End If
                                End If
                            Else
                                If Col.ColumnName = "id_Detalle" Then
                                    Cad += "|"
                                End If
                            End If
                        End If
                    Next
                    Cad = Cad + cad_imp_adic + Cad_Retencion
                    f.WriteLine(Cad)
                    Cad = "" 'LIPIO PARA SIGUIENTE LINEA
                Next
                'MsgBox(" Filas: " + cfilas.ToString + " Exentas: " + cexento.ToString)
                Dim SeccionImpuesto As String = "¬"
                If ctasa > 0 Then
                    If bT00 = "SA" Then
                        f.WriteLine(SeccionImpuesto & "TR|002|" & T00 & "|0.000000|Tasa")
                    ElseIf bT08 = "SA" Then
                        f.WriteLine(SeccionImpuesto & "TR|002|" & T08 & "|0.080000|Tasa")
                    ElseIf bT16 = "SA" Then
                        f.WriteLine(SeccionImpuesto & "TR|002|" & T16 & "|0.160000|Tasa")
                    End If
                    SeccionImpuesto = ""
                ElseIf cpcero > 0 Then
                    f.WriteLine(SeccionImpuesto & "TR|002|0.0000|0.000000|Tasa")
                    SeccionImpuesto = ""
                ElseIf cexento > 0 Then
                    'f.WriteLine("¬TR|002|0.0000|0.000000|Exento")
                    'SeccionImpuesto = ""
                End If

                If Not Encabezado.Is_192_Monto_TotalImp_RetenidosNull Then ' retenciones
                    If Encabezado._192_Monto_TotalImp_Retenidos > 0 Then
                        f.WriteLine(SeccionImpuesto & "RE|002|" & Encabezado._192_Monto_TotalImp_Retenidos & "|" & Tasa_RET & "|Tasa")
                        SeccionImpuesto = ""
                    End If
                End If

                If cad_imp_adic_total <> "" Then
                    f.WriteLine(SeccionImpuesto & cad_imp_adic_total)
                    SeccionImpuesto = ""
                End If


                CFDI_ComplementoPagoTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_ComplementoPago, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE
                    If Production_AUXDataSet.CFDI_ComplementoPago.Rows.Count > 0 Then
                        Cad = "¬*" ' PREPARO PARA DETALLES
                        For Each Complemento As ProduccionDS.CFDI_ComplementoPagoRow In Production_AUXDataSet.CFDI_ComplementoPago.Rows 'RECORRO DETALLE DE LA FACTURA EN CUESTION
                            For Each Col In Production_AUXDataSet.CFDI_ComplementoPago.Columns
                                If Col.ColumnName = "18_DetalleAux_Misc16" Then
                                    Cad += Complemento(Col).ToString.Trim
                                    Exit For
                                Else
                                    Cad += Complemento(Col).ToString.Trim & "|"
                                End If
                            Next
                            f.WriteLine(Cad)
                            Cad = "" 'LIPIO PARA SIGUIENTE LINEA
                        Next
                    End If

                    TotalImpuesto16 = 0
                    f.Close()
                    CFDI_EncabezadoTableAdapter.ProcesarFactura(True, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante)
                    contador += 1
                    If contador = 1 Then
                        'Exit For
                    End If
                End If
        Next

        Console.WriteLine("Proceso Terminado, se Generaron: " + contador.ToString + " CFDI txt ")
    End Sub

    Sub GeneraComplementoEkomercio()

        Dim Cad As String = "~"
        Dim vLimpia As String = ""

        Dim TotalImpuesto16 As Decimal = 0.0
        Dim TotalImpuesto0 As Decimal = 0
        Dim TotalImpuestoEXE As Decimal = 0

        Dim Encabezado As ProduccionDS.CFDI_EncabezadoRow
        Dim Detalle As ProduccionDS.CFDI_DetalleRow
        Dim Complemento As ProduccionDS.CFDI_ComplementoPagoRow


        CFDI_EncabezadoTableAdapter.FillByNoProcesadosREP(Production_AUXDataSet.CFDI_Encabezado) 'LLENO ENCABEZADO

        Dim f As StreamWriter
        Dim Col As DataColumn
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim contador As Integer = 0

        Dim vTipoImpuesto As String = ""
        Dim vExento As String = ""

        ' Recorrido de Renglones de Tabla Encabezado 
        For Each Encabezado In Production_AUXDataSet.CFDI_Encabezado.Rows() 'RECORRO FACTURAS SIN PROCESAR
            f = New StreamWriter(My.Settings.RutaFTP & "eKomplemento_" & Encabezado._3_RFC_Emisor & "_" & Encabezado._27_Serie_Comprobante & Encabezado._1_Folio & ".txt", False)
            Cad = "~"
            i += 1

            If Encabezado._83_Cod_Moneda = "XXX" And Encabezado._191_Efecto_Comprobante = "P" Then
                ' Recorrido de Columnas o Campos de Tabla Encabezado 
                For Each Col In Production_AUXDataSet.CFDI_Encabezado.Columns ' CONCATENO EL RENGLON DEL ENCABEZADO
                    If Col.ColumnName <> "Encabezado_Procesado" And Col.ColumnName <> "Fecha" Then
                        If Col.ColumnName <> "193_Monto_TotalImp_Trasladados" Then
                            ' 25 Octubre 2017
                            ' 6 de Noviembre se Agrego un Campo
                            If Col.ColumnName <> "Guid" Then
                                vLimpia = Encabezado(Col).ToString   ' Para quitar Salto de linea 25Octubre2017
                                Cad += vLimpia.Replace(vbCrLf, " ") & "|"   ' Para quitar Salto de linea 25Octubre2017
                                ' Cad += Encabezado(Col) & "|"     '   LINEA ORIGINAL SIN PIMPIAR 
                            End If
                        Else
                            'TotalImpuesto16 = Encabezado(Col)
                            Cad += Encabezado(Col).ToString
                        End If
                        j += 1
                    End If
                Next
                f.WriteLine(Cad)

                Cad = "¬" ' PREPARO PARA DETALLES
                CFDI_DetalleTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_Detalle, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE

                For Each Detalle In Production_AUXDataSet.CFDI_Detalle.Rows 'RECORRO DETALLE DE LA FACTURA EN CUESTION
                    Cad += Detalle._1_Linea_Descripcion
                    Cad += StrDup(1, "|")
                    Cad += Detalle._2_Linea_Cantidad.ToString
                    Cad += StrDup(1, "|")
                    Cad += Detalle._3_Linea_Unidad
                    Cad += StrDup(1, "|")
                    Cad += Detalle._4_Linea_PrecioUnitario.ToString
                    Cad += "|0|"
                    Cad += StrDup(10, "|")
                    Cad += Detalle._16_Linea_Cod_Articulo
                    Cad += StrDup(83, "|")
                    f.WriteLine(Cad)
                    Cad = "" 'LIPIO PARA SIGUIENTE LINEA

                    ' Proceso para llenar Tabla de COMPLEMENTO

                    Cad = "¬*" ' PREPARO PARA DETALLES
                    CFDI_ComplementoPagoTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_ComplementoPago, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE

                    For Each Complemento In Production_AUXDataSet.CFDI_ComplementoPago.Rows 'RECORRO DETALLE DE LA FACTURA EN CUESTION
                        j = j + 1
                        'MsgBox("Entra a Complemento")
                        If Complemento._1_DetalleAux_Tipo = "DR" Then
                            Cad = "¬*" ' PREPARO PARA DETALLES
                        End If
                        Cad += Complemento._1_DetalleAux_Tipo
                        Cad += StrDup(1, "|")
                        Cad += Complemento._2_DetalleAux_DescTipo

                        Cad += StrDup(1, "|")
                        Cad += Complemento._3_DetalleAux_Misc01

                        Cad += StrDup(1, "|")
                        Cad += Complemento._4_DetalleAux_Misc02

                        Cad += StrDup(1, "|")
                        Cad += Complemento._5_DetalleAux_Misc03

                        Cad += StrDup(1, "|")
                        Cad += Complemento._6_DetalleAux_Misc04

                        Cad += StrDup(1, "|")
                        Cad += Complemento._7_DetalleAux_Misc05

                        Cad += StrDup(1, "|")
                        Cad += Complemento._8_DetalleAux_Misc06

                        Cad += StrDup(1, "|")
                        Cad += Complemento._9_DetalleAux_Misc07

                        Cad += StrDup(1, "|")
                        Cad += Complemento._10_DetalleAux_Misc08

                        Cad += StrDup(1, "|")
                        Cad += Complemento._11_DetalleAux_Misc09

                        Cad += StrDup(1, "|")
                        Cad += Complemento._12_DetalleAux_Misc10

                        Cad += StrDup(1, "|")
                        Cad += Complemento._13_DetalleAux_Misc11

                        Cad += StrDup(5, "|")
                        f.WriteLine(Cad)
                        Cad = "" 'LIPIO PARA SIGUIENTE LINEA
                    Next
                Next
                contador += 1
                CFDI_EncabezadoTableAdapter.ProcesarFactura(True, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante)
            End If
            f.Close()
        Next
        Console.WriteLine("Proceso Terminado, se Generaron: " + contador.ToString + " Complementos de Pago, CFDI txt ")

    End Sub

    Sub Envia_RecibosPAGO()

        Dim Guid As Guid
        Dim Servidor As New Mail.SmtpClient
        Dim Mensaje As Mail.MailMessage
        Dim Adjunto As Mail.Attachment
        Dim CadenaGUID, Archivo As String
        Dim TaRec As New ProduccionDSTableAdapters.RecibosDePagoTableAdapter

        Dim t As New ProduccionDS.RecibosDePagoDataTable
        Dim t2 As New ProduccionDS.RecibosDePagoDataTable
        Dim crDiskFileDestinationOptions As New DiskFileDestinationOptions()

        Servidor.Host = My.Settings.SMTP
        Servidor.Port = My.Settings.SMTP_port
        Dim Cred() As String = My.Settings.SMTP_creden.Split(",")
        Servidor.Credentials = New System.Net.NetworkCredential(Cred(0), Cred(1), Cred(2))
        TaRec.MarcarComoRecibos()
        TaRec.Fill_Recibos(t)

        For Each r As ProduccionDS.RecibosDePagoRow In t.Rows
            Dim NewRPT As New GeneraFactura.CR_Recibo
            Dim ds As New ProduccionDS
            Dim TaRec2 As New ProduccionDSTableAdapters.RecibosDePagoTableAdapter
            Try
                CadenaGUID = Guid.NewGuid.ToString.ToUpper
                Console.WriteLine("Recibo de pago : " & CadenaGUID)
                Mensaje = New Mail.MailMessage
                Mensaje.IsBodyHtml = True
                Mensaje.From = New Mail.MailAddress("CFDI@cmoderna.com", "FINAGIL envíos automáticos")
                Mensaje.ReplyTo = New Mail.MailAddress("maria.vidal@finagil.com.mx", "Maria Vidal    (Finagil)")

                Mensaje.To.Add("ecacerest@finagil.com.mx")
                Mensaje.To.Add("maria.vidal@finagil.com.mx")
                Mensaje.To.Add("viapolo@finagil.com.mx")
                Mensaje.To.Add("denise.gonzalez@finagil.com.mx")
                If r.EMail1.ToString.Trim.Length > 3 Then
                    If validar_Mail(r.EMail1) Then
                        Mensaje.To.Add(r.EMail1)

                    End If
                End If
                If r.EMail2.ToString.Trim.Length > 3 Then
                    If validar_Mail(r.EMail2) Then
                        Mensaje.To.Add(r.EMail2)
                    End If
                End If

                Mensaje.Subject = "Comprobante de Pago Finagil -" & r._27_Serie_Comprobante.Trim & r._1_Folio & "(Sin valor Fiscal)"
                Mensaje.Body = "Estimado Cliente: " & r._42_Nombre_Receptor & "<br>" &
                        "Por este medio le hacemos el envio de su recibo de pago sin valor fiscal del contrato " & r._114_Misc02.Trim &
                        " por concepto de " & r._157_Misc45.Trim & "<br><br>Sin más por el momento agradecemos su atención y nos ponemos a su disposición en el teléfono 01 722 214 5533 ext. 1010 o al 01 800 727 7100, en caso de cualquier duda o comentario al respecto."


                TaRec2.FillByGUID(ds.RecibosDePago, r._1_Folio, r._27_Serie_Comprobante)
                ds.WriteXml("C:\Files\dsRec.xml", XmlWriteMode.WriteSchema)
                NewRPT.SetDataSource(ds)
                ds.Dispose()
                NewRPT.Refresh()
                NewRPT.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                NewRPT.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
                Archivo = "C:\Files\Recibo_" & CStr(r._1_Folio) & r._27_Serie_Comprobante.Trim & ".pdf"
                crDiskFileDestinationOptions.DiskFileName = Archivo
                NewRPT.ExportOptions.DestinationOptions = crDiskFileDestinationOptions
                NewRPT.Export()
                NewRPT.Dispose()



                TaRec.ReciboEnviado(CadenaGUID, "Recibo de Pago", r._1_Folio, r._27_Serie_Comprobante)

                If Directory.Exists(GeneraFactura.My.Settings.RutaRecPago & Date.Now.Year.ToString & "-" & MonthName(Date.Now.Month) & "-" & String.Format("{0:00}", Date.Now.Day)) = False Then
                    Directory.CreateDirectory(GeneraFactura.My.Settings.RutaRecPago & Date.Now.Year.ToString & "-" & MonthName(Date.Now.Month) & "-" & String.Format("{0:00}", Date.Now.Day))
                End If

                System.IO.File.Copy(Archivo, GeneraFactura.My.Settings.RutaRecPago & Date.Now.Year.ToString & "-" & MonthName(Date.Now.Month) & "-" & String.Format("{0:00}", Date.Now.Day) & "\Recibo_" & CStr(r._1_Folio) & r._27_Serie_Comprobante.Trim & ".pdf")

                Adjunto = New Mail.Attachment(Archivo, "PDF/pdf")
                Mensaje.Attachments.Add(Adjunto)
                Servidor.Send(Mensaje)
                Adjunto.Dispose()

                Console.WriteLine("Envio Exitoso :" & Archivo)
                System.IO.File.Delete(Archivo)
            Catch ex As Exception
                Console.WriteLine("error:" & ex.Message)
                Adjunto.Dispose()
                System.IO.File.Delete(Archivo)
            End Try
        Next
        'NewRPT.Dispose()
    End Sub

End Module
