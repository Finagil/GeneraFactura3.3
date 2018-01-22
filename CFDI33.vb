﻿Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Math
Module CFDI33
    Dim drUdis As DataRowCollection
    Dim nIDSerieA As Decimal = 0
    Dim nIDSerieMXL As Decimal = 0
    Dim cSerie As String = ""
    Dim cSucursal As String = ""
    Dim nTasaIVACliente As Decimal = 0
    Dim Production_AUXDataSet As New ProduccionDS
    Dim CFDI_EncabezadoTableAdapter As New ProduccionDSTableAdapters.CFDI_EncabezadoTableAdapter
    Dim CFDI_DetalleTableAdapter As New ProduccionDSTableAdapters.CFDI_DetalleTableAdapter
    Dim CFDI_ComplementoPagoTableAdapter As New ProduccionDSTableAdapters.CFDI_ComplementoPagoTableAdapter

    Sub FacturarCFDI(FechaProc As Date, Tipo As String)
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
        Dim cFechaAplicacion As String = ""
        Dim i As Integer = 0
        Dim nRecibo As Decimal = 0


        cFechaAplicacion = FechaProc.ToString("yyyyMMdd")
        cFechaPago = cFechaAplicacion
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
        nIDSerieA = drSerie("IDSerieA")
        nIDSerieMXL = drSerie("IDSerieMXL")

        ' Solo necesito saber el número de elementos que tiene el DataGridView1
        Select Case Tipo.ToUpper
            Case "PREPAGO" ' prepagos antes de su fecha de vencimiento
                TaAvisos.FillByPrepagos(ProdDS.AvisosCFDI, cFechaPago, "20171201")'Fecha de Salida a Producion
            Case "DIA" 'avisos de vencimiento del dia
                TaAvisos.FillporDia(ProdDS.AvisosCFDI, cFechaPago)
            Case "ANTERIORES" ' avisos generados despues de su vencimiento
                TaAvisos.FillByAnteriores(ProdDS.AvisosCFDI, cFechaPago)
                'Case "PENDIENTES"
                'TaAvisos.FillHastaFecha(ProdDS.AvisosCFDI, cFechaPago)
        End Select

        'TaAvisos.FillHastaFecha(ProdDS.AvisosCFDI, cFechaPago)
        TaUdis.Fill(ProdDS.TraeUdis)
        drUdis = ProdDS.TraeUdis.Rows
        For Each r As ProduccionDS.AvisosCFDIRow In ProdDS.AvisosCFDI.Rows
            Console.WriteLine("Aviso:" & r.Factura)

            If TaAvisos.AnexosNoFacturables(r.Anexo) > 0 Then
                TaAvisos.FacturarAviso(True, "", 0, r.Factura, r.Anexo)
                Continue For
            End If

            cAnexo = r.Anexo
            If r.Fepag.Trim.Length > 0 Then
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
            nTasaIVACliente = r.TasaIVACliente


            If cSucursal = "04" Or nTasaIVACliente = 11 Then
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
                    nIDSerieA = nIDSerieA + 1
                    nRecibo = nIDSerieA
                ElseIf cSerie = "MXL" Then
                    nIDSerieMXL = nIDSerieMXL + 1
                    nRecibo = nIDSerieMXL
                End If
                MetodoPago = "PPD"
                cLetra = r.Letra
                Acepagov(cAnexo, cLetra, nMontoPago, nMoratorios, nIvaMoratorios, cBanco, cCheque, dtMovimientos, cFechaAplicacion, cFechaPago, cSerie, nRecibo, InstrumentoMonetario, FechaProc, MetodoPago)

                If cSerie = "A" And nRecibo <> 0 Then
                    strUpdate = "UPDATE Llaves SET IDSerieA = " & nRecibo
                ElseIf cSerie = "MXL" And nRecibo <> 0 Then
                    strUpdate = "UPDATE Llaves SET IDSerieMXL = " & nRecibo
                End If
                TaAvisos.FacturarAviso(True, cSerie.Trim, nRecibo, r.Factura, r.Anexo)
                cm1 = New SqlCommand(strUpdate, cnAgil)
                cnAgil.Open()
                cm1.ExecuteNonQuery()
                cnAgil.Close()
            End If
        Next

        cnAgil.Dispose()
        cm1.Dispose()
        cm2.Dispose()
        cm3.Dispose()
        cm4.Dispose()
    End Sub

    Sub FacturarCFDI_AV(FechaProc As Date)
        Dim Ta As New ProduccionDSTableAdapters.TraspasosAvioCCTableAdapter
        Dim t As New ProduccionDS.TraspasosAvioCCDataTable
        Dim nRecibo As Integer
        Dim cRenglon As String
        Dim FechaS As String = FechaProc.ToString("yyyyMMdd")

        Ta.Fill(t, FechaS)
        For Each r As ProduccionDS.TraspasosAvioCCRow In t.Rows
            If r.Sucursal = "04" Then
                cSerie = "MXL"
                nRecibo = Ta.SerieMXL
            Else
                cSerie = "A"
                nRecibo = Ta.SerieA
            End If


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


            cRenglon = "D1|" & r.Cliente & "|" & Mid(r.Anexo, 1, 5) & "/" & Mid(r.Anexo, 6, 4) & "|" & cSerie & "|" & nRecibo & "|1|||INTERESES AVIO||" & r.Intereses + r.InteresesDias & "|0"
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

            If r.Sucursal = "04" Then
                Ta.ConsumeSerieMXL()
            Else
                Ta.ConsumeSerieA()
            End If
            Ta.FacturarTraspaso(True, cSerie, nRecibo, r.id_Traspaso)
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
        Lectura.Dispose()

        Dim tAREC As New ProduccionDSTableAdapters.CFDI_RecibosPAGOTableAdapter
        Dim tREC As New ProduccionDS.CFDI_RecibosPAGODataTable
        tAREC.Fill(tREC)
        For Each r As ProduccionDS.CFDI_RecibosPAGORow In tREC.Rows
            taFact.UpdateGUID("Recibo de Pago", r._1_Folio, r._27_Serie_Comprobante)
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

    Sub GeneraFacturaEkomercio()
        Dim Cad As String = "~"
        Dim TotalImpuesto16 As Decimal = 0.0
        Dim TotalImpuesto0 As Decimal = 0
        Dim TotalImpuestoEXE As Decimal = 0
        Dim Encabezado As ProduccionDS.CFDI_EncabezadoRow
        Dim Detalle As ProduccionDS.CFDI_DetalleRow
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

        CFDI_EncabezadoTableAdapter.FillByNoProcesadosFACT(Production_AUXDataSet.CFDI_Encabezado) 'LLENO ENCABEZADO

        ' Recorrido de Renglones de Tabla Encabezado 
        For Each Encabezado In Production_AUXDataSet.CFDI_Encabezado.Rows() 'RECORRO FACTURAS SIN PROCESAR

            CFDI_DetalleTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_Detalle, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE

            If Production_AUXDataSet.CFDI_Detalle.Rows.Count > 0 Then
                f = New StreamWriter(My.Settings.RutaFTP & "eKomercio_" & Encabezado._27_Serie_Comprobante & Encabezado._1_Folio & ".txt", False)
                If CDate(Encabezado._30_Fecha) < Date.Now.Date Then
                    Encabezado._31_Hora = Date.Now.AddHours(2).ToString("HH:mm:ss")
                End If

                Cad = "~"
                i += 1

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
                            TotalImpuesto16 = Encabezado(Col)
                            Cad += Encabezado(Col).ToString
                        End If
                        j += 1
                    End If
                Next
                f.WriteLine(Cad)

                Cad = "¬" ' PREPARO PARA DETALLES
                CFDI_DetalleTableAdapter.FillByFactura(Production_AUXDataSet.CFDI_Detalle, Encabezado._1_Folio, Encabezado._27_Serie_Comprobante) 'LLENO DETALLE
                cfilas += 1
                ctasa = 0
                cpcero = 0
                cexento = 0
                For Each Detalle In Production_AUXDataSet.CFDI_Detalle.Rows 'RECORRO DETALLE DE LA FACTURA EN CUESTION

                    For Each Col In Production_AUXDataSet.CFDI_Detalle.Columns ' CONCATENO EL RENGLON DE DETALLE CON IMPUESTOS

                        If Col.ColumnName = "1_Impuesto_TipoImpuesto" Then
                            If Detalle.Item("6_Impuesto_Tasa") = "No Objeto" Then
                                Exit For
                            End If
                            Cad += "\Impuesto|" ' DIVIDO SESSION DE IMPOUESTOS EN DETALLE
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
                                    Else
                                        ' 21 Noviembre
                                        If Col.ColumnName = "6_Impuesto_Tasa" Then
                                            If Detalle(Col).ToString = "Exento" Then
                                                cexento += 1
                                            End If
                                            If Detalle(Col).ToString = "Tasa" Then
                                                ctasa += 1
                                            End If
                                        Else
                                            If Col.ColumnName = "7_Impuesto_Porcentaje" Then
                                                If Detalle(Col).ToString = "0.0000" Then
                                                    cpcero += 1
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
                    f.WriteLine(Cad)
                    Cad = "" 'LIPIO PARA SIGUIENTE LINEA
                Next
                'MsgBox(" Filas: " + cfilas.ToString + " Exentas: " + cexento.ToString)

                If ctasa > 0 Then
                    f.WriteLine("¬TR|002|" & TotalImpuesto16 & "|0.160000|Tasa")
                End If

                If cpcero > 0 Then
                    f.WriteLine("¬TR|002|0.0000|0.000000|Tasa")
                End If

                If cexento > 0 Then
                    'f.WriteLine("¬TR|002|0.0000|0.000000|Exento")
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
            f = New StreamWriter(My.Settings.RutaFTP & "eKomplemento_" & Encabezado._27_Serie_Comprobante & Encabezado._1_Folio & ".txt", False)
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
        Dim NewRPT As New GeneraFactura.CR_Recibo
        Dim Guid As Guid
        Dim Servidor As New Mail.SmtpClient
        Dim Mensaje As Mail.MailMessage
        Dim Adjunto As Mail.Attachment
        Dim CadenaGUID, Archivo As String
        Dim TaRec As New ProduccionDSTableAdapters.RecibosDePagoTableAdapter
        Dim ds As New ProduccionDS
        Dim t As New ProduccionDS.RecibosDePagoDataTable
        Dim crDiskFileDestinationOptions As New DiskFileDestinationOptions()

        Servidor.Host = "smtp01.cmoderna.com"
        Servidor.Port = "26"
        TaRec.RecibosProcesados()
        TaRec.Fill_Recibos(t)


        For Each r As ProduccionDS.RecibosDePagoRow In t.Rows
            Try
                CadenaGUID = Guid.NewGuid.ToString.ToUpper
                Mensaje = New Mail.MailMessage
                Mensaje.IsBodyHtml = True
                Mensaje.From = New Mail.MailAddress("CFDI@Finagil.com.mx", "FINAGIL envíos automáticos")
                Mensaje.ReplyTo = New Mail.MailAddress("maria.vidal@finagil.com.mx", "Maria Vidal    (Finagil)")

                Mensaje.To.Add("ecacerest@finagil.com.mx")
                If r.EMail1.Length > 3 Then
                    'Mensaje.To.Add(r.EMail1)
                End If
                If r.EMail2.Length > 3 Then
                    'Mensaje.To.Add(r.EMail2)
                End If

                Mensaje.Subject = "Recibo de Pago Finagil -" & r._27_Serie_Comprobante.Trim & r._1_Folio & "(Sin valor Fiscal)"
                Mensaje.Body = "Estimado Cliente: " & r._42_Nombre_Receptor & "<br>" &
                        "Por este medio le hacemos el envio de su recibo de pago sin valor fiscal del contrato " & r._114_Misc02.Trim &
                        " por concepto de " & r._157_Misc45.Trim & "<br><br>Sin más por el momento agradecemos su atención y nos ponemos a su disposición en el teléfono 01 722 214 5533 ext. 1010 o al 01 800 727 7100, en caso de cualquier duda o comentario al respecto."
                TaRec.ReciboEnviado(CadenaGUID, "Recibo de PAgo", r._1_Folio, r._27_Serie_Comprobante)

                TaRec.FillByGUID(ds.RecibosDePago, CadenaGUID)
                NewRPT.SetDataSource(ds)
                NewRPT.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                NewRPT.ExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
                Archivo = "C:\FILES\Recibo_" & CStr(r._1_Folio) & r._27_Serie_Comprobante.Trim & ".pdf"
                crDiskFileDestinationOptions.DiskFileName = Archivo
                NewRPT.ExportOptions.DestinationOptions = crDiskFileDestinationOptions
                NewRPT.Export()
                NewRPT.Dispose()

                Adjunto = New Mail.Attachment(Archivo, "PDF/pdf")
                Mensaje.Attachments.Add(Adjunto)
                Servidor.Send(Mensaje)
                Console.WriteLine("Envio Exsitoso :" & Archivo)
            Catch ex As Exception
                Console.WriteLine("error:" & ex.Message)
            End Try
        Next
    End Sub

End Module