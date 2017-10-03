Imports System.IO
Imports System.Net.Mail
Imports System.Data.SqlClient

Module GneraFactura
    Dim ErrorControl As New EventLog
    Dim OpcionCompraAF As String

    Sub Main()
        'Try
        Dim mf As Date = Date.Now.AddHours(-72)
        Console.WriteLine("Inicia proceso")
        ErrorControl = New EventLog("Application", System.Net.Dns.GetHostName(), "GeneracionCFDI33")
        Dim D As New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        ' procesa y genera txt-CFDI
        Dim F As FileInfo() = D.GetFiles("*.txt")

        Console.WriteLine("Generando CFDI Avio...")
        GeneraArchivosAvio()
        Console.WriteLine("Generando CFDI Facturas Externas...")
        GeneraArchivosEXternas()
        Console.WriteLine("leyendo " & GeneraFactura.My.Settings.RutaOrigen)
        Console.WriteLine("Generando CFDI...")
        GeneraArchivos()
        ' borra archivos leidos
        D = New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        F = D.GetFiles("*.txt")
        Console.WriteLine("Borrando procesados...")
        For i As Integer = 0 To F.Length - 1
            Console.WriteLine(GeneraFactura.My.Settings.RutaOrigen & F(i).Name)
            File.Copy(F(i).FullName, GeneraFactura.My.Settings.RutaBackup & F(i).Name, True)
            File.Delete(F(i).FullName)
        Next
        ' mandamos archivos el servidor CFDI
        D = New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaCFDI)
        F = D.GetFiles("*.txt")
        Console.WriteLine("Enviando a Servidor CFDI...")
        Dim Tope As Integer = F.Length - 1
        If Tope > 50 Then Tope = 50
        For i As Integer = 0 To Tope
            Console.WriteLine(GeneraFactura.My.Settings.ServerCFDIFin & F(i).Name)
            File.Copy(F(i).FullName, GeneraFactura.My.Settings.RutaBackup & F(i).Name & Date.Now.ToString("yyMMddHHmmss") & ".txt", True)
            If InStr(UCase(F(i).Name), "ARFIN", ) Then
                File.Copy(F(i).FullName, GeneraFactura.My.Settings.ServerCFDIArf & F(i).Name, True)
            Else
                File.Copy(F(i).FullName, GeneraFactura.My.Settings.ServerCFDIFin & F(i).Name, True)
            End If

            File.Delete(F(i).FullName)
        Next
        Console.WriteLine("Terminado...")
        'Catch ex As Exception
        '    EnviaError(GeneraFactura.My.Settings.MailError, ex.Message, "Error " & Now.Date)
        'End Try
    End Sub

    Sub GeneraArchivos()
        Const MontAux As String = "0"
        Dim x As Integer
        Dim i As Integer
        Dim EsNotaCredito As Boolean = False
        Dim EsArfin As Boolean = False
        Dim EnviarGisela As Boolean = False
        Dim EsFinagil As Boolean = False
        Dim ta As New GeneraFactura.ProduccionDSTableAdapters.ClientesTableAdapter
        Dim t As New GeneraFactura.ProduccionDS.ClientesDataTable
        Dim taMail As New ProduccionDSTableAdapters.CorreosAnexosTableAdapter
        Dim tMail As New ProduccionDS.CorreosAnexosDataTable
        Dim Rmail As ProduccionDS.CorreosAnexosRow
        Dim Linea As String = ""
        Dim Mail As String = ""
        Dim SinDet As String = ""
        Dim suma As Double
        Dim Arre(2, 20) As String
        Dim cad As String = ""
        Dim cadX As String = ""
        Dim cadY As String = ""
        Dim M1 As String = ""
        Dim M2 As String = ""
        Dim Datos() As String
        Dim D As New System.IO.DirectoryInfo(GeneraFactura.My.Settings.RutaOrigen)
        Dim F As System.IO.FileInfo() = D.GetFiles("*.txt")
        Dim f1 As System.IO.StreamWriter
        Dim f2 As System.IO.StreamReader
        Dim fecha As New DateTime
        Dim Entero As Integer
        Dim Importe As Double
        Dim ISR As Double = 0
        Dim horas As Integer
        Dim sig As String = ""
        Dim Tipar As String = ""
        Dim TipoPersona As String = ""
        Dim Moneda As String = ""
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim BanImpuesto As Boolean = False
        Dim MetodoPago As String
        Dim cAnexo As String
        Dim cMail1 As String
        Dim cMail2 As String
        'Try

        For i = 0 To F.Length - 1
            'Try

            Console.WriteLine("Generando CFDI..." & F(i).Name)
                For x = 0 To 20
                    Arre(0, x) = ""
                    Arre(1, x) = ""
                    Arre(2, x) = ""
                Next
                x = 0
                suma = 0
                LecturaPrevia(F(i).FullName, F(i).Name)
                f2 = New System.IO.StreamReader(F(i).FullName, Text.Encoding.GetEncoding(1252))
                If Mid(F(i).Name, 1, 3) <> "FIN" And Mid(F(i).Name, 1, 3) <> "XXA" And IsNumeric(Mid(F(i).Name, 1, 4)) = True Then
                    fecha = New DateTime(Mid(F(i).Name, 1, 4), Mid(F(i).Name, 5, 2), Mid(F(i).Name, 7, 2), Mid(F(i).Name, 9, 2), Mid(F(i).Name, 11, 2), Mid(F(i).Name, 13, 2))
                    horas = DateDiff(DateInterval.Hour, fecha, Date.Now)
                    If horas >= 72 Then
                        fecha = fecha.AddHours(horas - 71)
                    End If
                    f1 = New System.IO.StreamWriter(GeneraFactura.My.Settings.RutaCFDI & "CFDI-" & Mid(F(i).Name, 16, 100), False, Text.Encoding.GetEncoding(1252))
                Else
                If F(i).Name.Length < 13 + 14 Then
                    f1 = New System.IO.StreamWriter(GeneraFactura.My.Settings.RutaCFDI & "CFDI-" & F(i).Name, False, Text.Encoding.GetEncoding(1252))
                Else
                    f1 = New System.IO.StreamWriter(GeneraFactura.My.Settings.RutaCFDI & "CFDI-" & Mid(F(i).Name, 13, 100), False, Text.Encoding.GetEncoding(1252))
                End If


            End If
                'fecha = CDate("28/12/2016")
                EsNotaCredito = False
                EsArfin = False
                EsFinagil = True
                EnviarGisela = False
                BanImpuesto = False
                SinDet = "SinDetalle"
                MetodoPago = "99"
                ISR = 0
                While Not f2.EndOfStream
                    Linea = f2.ReadLine
                    If UCase(Linea) = "X" Then
                        EnviarGisela = True
                        Linea = f2.ReadLine
                    End If
                    Datos = Linea.Split("|")
                    If Datos.Length > 3 Then
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
                    cAnexo = Mid(Datos(2), 1, 5) & Mid(Datos(2), 7, 4)
                    Tipar = taTipar.Tipar(Datos(2))
                    TipoPersona = taTipar.TipoPersona(Datos(1))
                    Select Case Datos(0)
                        Case "M1"
                        EsArfin = False
                        fecha = Datos(6)
                        ''fecha = "06/09/2017 14:00:00" 'NO DEJAR ESTA LINEA
                        Mail = Datos(5)
                            If UCase(Datos(7)) = "ARFIN" Then
                                EsFinagil = False
                                EsArfin = True
                            End If
                        Case "H1"
                            fecha = Datos(1)
                            fecha = fecha.AddHours(Date.Now.Hour + 1)
                            fecha = fecha.AddMinutes(Date.Now.Minute)
                        fecha = fecha.AddSeconds(Date.Now.Second)
                        'fecha = "06/09/2017 14:00:00" 'NO DEJAR ESTA LINEA
                    Case "H3"
                            f1.WriteLine("#InicioComprobante")
                            If Mid(Datos(3), 1, 1) = "C" Then
                                f1.WriteLine("idn:documento                   =        NCREDITO")
                                EsNotaCredito = True
                            Else
                                f1.WriteLine("idn:documento                   =        FACTURA")
                                EsNotaCredito = False
                            End If
                            f1.WriteLine("idn:tipofactura                 =        FINANZAS")
                            f1.WriteLine("idn:planta                      =        TOLUCA")
                            f1.WriteLine("idn:tipodocto                   =        NACIONAL")
                            If EsFinagil = True Then
                                f1.WriteLine("idn:documanager                 =        FINAGIL")
                            Else
                                f1.WriteLine("idn:documanager                 =        ARFIN")
                            End If
                        If Datos(17) = "M.N." Then Datos(17) = "MXN"
                        If Datos(17) = "MXP" Then Datos(17) = "MXN"


                        If Datos(2) = "03473/0001" Or Datos(2) = "04235/0001" Then
                            Datos(17) = "USD"
                            Moneda = "USD"
                        Else
                            Moneda = Datos(17)
                        End If

                        f1.WriteLine()
                        If Mid(F(i).Name, 1, 3) = "XXA" Then Datos(3) = "DV"
                        'If Datos(3) = "F" Then Datos(3) = "FF" 'NO DEJAR ESTA LINEA
                        f1.WriteLine("fsc:serie                       =        " & Left(Datos(3), 8))
                        f1.WriteLine("fsc:folio                       =        " & Val(Datos(4)))
                        f1.WriteLine("fsc:fecha                       =        " & fecha.ToString("yyyy-MM-dd") & "T" & fecha.ToString("HH:mm:ss"))

                        f1.WriteLine("fsc:formaDePago                 =        PAGO EN UNA SOLA EXHIBICION")
                        f1.WriteLine("fsc:noCertificado               =        00001000000202240016")
                        f1.WriteLine("fsc:condicionesDePago           =        ")
                        f1.WriteLine("fsc:motivoDescuento             =	       ")

                        f1.WriteLine("fsc:TipoCambio                  =        " & ta.SacaTipoCambio(fecha.Date, Moneda))
                        f1.WriteLine("fsc:Moneda                      =        " & Moneda)
                        If Datos(3) = "C" Then
                                f1.WriteLine("fsc:tipoDeComprobante           =        Egreso")
                            Else
                                f1.WriteLine("fsc:tipoDeComprobante           =        Ingreso")
                            End If

                            Select Case (Trim(Datos(14)))
                                Case "01", "02", "03", "04", "05", "06", "07", "08", "28", "29", "99"
                                    MetodoPago = Trim(Datos(14))
                                Case "Transferencia de fondos"
                                    MetodoPago = "03"
                                Case Else
                                    If InStr(Datos(14).Trim, "Transfe") Then
                                        MetodoPago = "03"
                                    ElseIf InStr(Datos(14).Trim, "Cheque") Then
                                        MetodoPago = "02"
                                    Else
                                        MetodoPago = taTipar.SacaID_Metodo(Trim(Datos(14)))
                                    End If

                            End Select

                            f1.WriteLine("fsc:metodoDePago                =        " & MetodoPago)
                            f1.WriteLine("fsc:LugarExpedicion             =        Toluca, México")
                            If Not IsNumeric(Datos(13)) Then Datos(13) = ""
                            f1.WriteLine("fsc:NumCtaPago                  =        " & Datos(13))
                            f1.WriteLine()
                            f1.WriteLine("#Emisor")
                            If EsFinagil = True Then
                                f1.WriteLine("fem:rfc                         =        FIN940905AX7")
                                f1.WriteLine("fem:nombre                      =        FINAGIL S.A. DE C.V, SOFOM E.N.R")
                            Else
                                f1.WriteLine("fem:rfc                         =        SAR951230N5A")
                                f1.WriteLine("fem:nombre                      =        SERVICIOS ARFIN S.A. DE C.V.")
                            End If

                            f1.WriteLine("fed:calle                       =        Leandro Valle")
                            f1.WriteLine("fed:noExterior                  =        402")
                            f1.WriteLine("fed:noInterior                  =        ")
                            f1.WriteLine("fed:colonia                     =        Reforma y F.F.C.C")
                            f1.WriteLine("fed:localidad                   =        Toluca")
                            f1.WriteLine("fed:municipio                   =        Toluca")
                            f1.WriteLine("fed:estado                      =        Estado de Mexico")
                            f1.WriteLine("fed:pais                        =        Mexico")
                            f1.WriteLine("fed:codigoPostal                =        50070")
                            f1.WriteLine()
                            f1.WriteLine("#ExpendidoEn")
                            f1.WriteLine("fee:calle                       =        Leandro Valle")
                            f1.WriteLine("fee:noExterior                  =        402")
                            f1.WriteLine("fee:noInterior                  =        ")
                            f1.WriteLine("fee:colonia                     =        Reforma y F.F.C.C")
                            f1.WriteLine("fee:localidad                   =        Toluca")
                            f1.WriteLine("fee:municipio                   =        Toluca")
                            f1.WriteLine("fee:estado                      =        Estado de Mexico")
                            f1.WriteLine("fee:pais                        =        Mexico")
                            f1.WriteLine("fee:codigopostal                =        50070")
                            f1.WriteLine("fer:regimen                     =        REGIMEN GENERAL DE LEY PERSONAS MORALES")
                            f1.WriteLine()
                            f1.WriteLine("#Receptor")
                            'If Len(Datos(16)) = 10 Then Datos(16) = Datos(16) & "000"
                            'If Len(Datos(16)) < 10 Then Datos(16) = "XAXX010101000"
                            'If Len(Datos(16)) > 13 Then Datos(16) = "XAXX010101000"
                            Datos(16) = ValidaRFC(Datos(16), TipoPersona)
                            f1.WriteLine("fre:rfc                         =        " & Datos(16).ToUpper)
                            If Trim(Datos(16)) = "SDA070613KU6" Then
                                Datos(5) = """SERVICIO DAYCO"" SA DE CV"
                            End If
                            If Trim(Datos(16)) = "CARD840606LEA" Then
                                Datos(5) = "DANIEL CADENA RUVALCABA"
                            End If

                            f1.WriteLine("fre:nombre                      =        " & Datos(5))
                            f1.WriteLine("frd:calle                       =        " & Datos(6))
                            f1.WriteLine("frd:noExterior                  =        ")
                            f1.WriteLine("frd:noInterior                  =        ")
                            f1.WriteLine("frd:colonia                     =        " & Datos(9))
                            f1.WriteLine("frd:localidad                   =        ")
                            f1.WriteLine("frd:municipio                   =        " & Datos(10))
                            f1.WriteLine("frd:estado                      =        " & Datos(11))
                            f1.WriteLine("frd:pais                        =        " & Datos(15))
                            f1.WriteLine("frd:codigopostal                =        " & Datos(12))
                            f1.WriteLine()
                            f1.WriteLine("#Detalle")
                            'f1.WriteLine(" dco:cant  dco:unit   dco:noId         dco:desc                       dco:vUni      dco:impo       dcd:line dcd:unid dcd:kg     dcd:prom  dcd:coAl")
                            f1.WriteLine(" dco:cant  dco:unit   dco:noId         dco:desc                                                              dco:vUni      dco:impo       dcd:line dcd:unid dcd:kg     dcd:prom  dcd:coAl")
                            f1.WriteLine()
                        Case "D1"
                            If InStr(UCase(Datos(8)), "IVA") <= 0 Then
                                SinDet = ""
                                x += 1
                                Datos(8) = Trim(Datos(8))
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
                                If (Tipar = "R" Or Tipar = "S") And TipoPersona <> "F" Then
                                    Select Case Mid(Datos(8), 1, 12)
                                        Case "MORATORIOS V"
                                            Datos(8) = Datos(8) & " EXENTOS DE IVA"
                                        Case "INTERESES VE"
                                            Datos(8) = Datos(8) & " EXENTOS DE IVA"
                                    End Select
                                End If
                                If Tipar = "F" And Datos(2) = "02562/0003" Then '#ECT Solicitado por Valentin 24/09/2015
                                    Select Case Datos(8)
                                        Case "ADELANTO CAPITAL EQUIPO"
                                            Datos(8) = Datos(8) & " A TASA IVA 0%"
                                        Case "CAPITAL EQUIPO"
                                            Datos(8) = Datos(8) & " A TASA IVA 0%"
                                        Case "AMORTIZACION INICIAL"
                                            Datos(8) = Datos(8) & " A TASA IVA 0%"
                                    End Select
                                    If Mid(Datos(8), 1, 9) = "INTERESES" Then
                                        Datos(8) = Datos(8) & " EXENTOS DE IVA"
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
                                        Datos(8) = Datos(8) & " EXENTOS DE IVA"
                                    End If
                                End If

                                If InStr(Datos(8), "SEGURO DE VI") > 0 Then
                                    Datos(8) = Datos(8) & " EXENTOS DE IVA"
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

                                Arre(1, x) = Datos(8)
                                Datos(10) = Math.Round(CDbl(Datos(10)), 4)
                                If Datos(7) = "" Then Datos(7) = Datos(10)
                                Datos(7) = Math.Round(CDbl(Datos(7)), 2) ' valor unitario
                                If Datos(6) = "" Then
                                    Datos(6) = "NAP"
                                End If
                                Datos(6) = Mid(Datos(6), 1, 3)
                                Arre(0, x) = Datos(6)

                                If Datos(8).Length <= 70 Then
                                    Datos(8) = Mid(Trim(Datos(8)), 1, 70)
                                    f1.WriteLine(Space(9 - Datos(5).Length) & Datos(5) & "   " & Datos(6) & Space(24) & Datos(8) & Space(70 - Trim(Datos(8).Length)) & Space(13 - Datos(7).Length) & Datos(7) & Space(15 - Trim(Datos(10).Length)) & Trim(Datos(10) & Space(20 - Trim(Datos(10).Length))) & "  ")
                                Else
                                    Dim Cero As String = "0"
                                    Dim Largo As Integer = Datos(8).Length
                                    Dim posss As Integer = 70
                                    While Largo <> 0
                                        If Largo = Datos(8).Length Then
                                            f1.WriteLine(Space(9 - Datos(5).Length) & Datos(5) & "   " & Datos(6) & Space(24) & Left(Datos(8), 70) & Space(13 - Datos(7).Length) & Datos(7) & Space(15 - Trim(Datos(10).Length)) & Trim(Datos(10) & Space(20 - Trim(Datos(10).Length))) & "  ")
                                            Largo -= 70
                                        Else
                                            f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), posss + 1, 70) & Space(70 - Mid(Datos(8), posss + 1, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                            posss += 70
                                            Largo -= 70
                                            If Largo < 0 Then
                                                Largo = 0
                                            End If
                                        End If
                                    End While
                                    'Select Case Datos(8).Length
                                    '    Case 71 To 140
                                    '        f1.WriteLine(Space(9 - Datos(5).Length) & Datos(5) & "   " & Datos(6) & Space(24) & Left(Datos(8), 70) & Space(13 - Datos(7).Length) & Datos(7) & Space(15 - Trim(Datos(10).Length)) & Trim(Datos(10) & Space(20 - Trim(Datos(10).Length))) & "  ")
                                    '        f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), 71, 70) & Space(70 - Mid(Datos(8), 71, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                    '    Case 141 To 210
                                    '        f1.WriteLine(Space(9 - Datos(5).Length) & Datos(5) & "   " & Datos(6) & Space(24) & Left(Datos(8), 70) & Space(13 - Datos(7).Length) & Datos(7) & Space(15 - Trim(Datos(10).Length)) & Trim(Datos(10) & Space(20 - Trim(Datos(10).Length))) & "  ")
                                    '        f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), 71, 70) & Space(70 - Mid(Datos(8), 71, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                    '        f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), 141, 70) & Space(70 - Mid(Datos(8), 141, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                    '    Case 211 To 280
                                    '        f1.WriteLine(Space(9 - Datos(5).Length) & Datos(5) & "   " & Datos(6) & Space(24) & Left(Datos(8), 70) & Space(13 - Datos(7).Length) & Datos(7) & Space(15 - Trim(Datos(10).Length)) & Trim(Datos(10) & Space(20 - Trim(Datos(10).Length))) & "  ")
                                    '        f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), 71, 70) & Space(70 - Mid(Datos(8), 71, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                    '        f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), 141, 70) & Space(70 - Mid(Datos(8), 141, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                    '        f1.WriteLine(Space(9 - Cero.Length) & Cero & "   " & Datos(6) & Space(24) & Mid(Datos(8), 141, 70) & Space(70 - Mid(Datos(8), 141, 70).Length) & Space(13 - Cero.Length) & Cero & Space(15 - Trim(MontAux.Length)) & Trim(MontAux & Space(20 - Trim(MontAux.Length))) & "  ")
                                    'End Select

                                End If
                                suma = suma + Val(Datos(10))
                                cad = "*"
                                Const Letras As Integer = 86
                                If Arre(1, x).Length > Letras And EsArfin = False Then
                                    Entero = Fix(Arre(1, x).Length / Letras)
                                    cadY = Datos(8)
                                    Importe = Datos(10)
                                    For ww = 1 To Entero + 1
                                        cadX = ""
                                        While cadX.Length + sig.Length <= Letras
                                            cadX = cadX & Mid(cadY, 1, InStr(cadY, " "))
                                            cadY = Mid(cadY, InStr(cadY, " ") + 1, cadY.Length)
                                            If InStr(cadY, " ") <= 0 Then cadY = cadY & " "
                                            sig = Mid(cadY, 1, InStr(cadY, " "))
                                        End While

                                        Arre(1, x) = Mid(Trim(cadX), 1, Letras)
                                        If Importe <> 0 Then
                                            Arre(2, x) = cad & Format(Importe, "#,##0.00")
                                        Else
                                            Arre(2, x) = cad
                                        End If

                                        x += 1
                                        Importe = 0
                                        cadX = Mid(cadX, Letras + 1, cadX.Length)
                                    Next
                                    x -= 1
                                Else
                                    If CDbl(Datos(10)) <> 0 Then
                                        Arre(2, x) = cad & Format(CDbl(Datos(10)), "#,##0.00")
                                    Else
                                        Arre(2, x) = cad
                                    End If

                                End If
                            End If
                        Case "S0"
                            ISR = Math.Round(Val(Datos(6)), 4)
                        Case "S1"
                            Datos(5) = Math.Round(Val(Datos(5)), 4)
                            Datos(6) = Math.Round(Val(Datos(6)), 4)
                            Datos(7) = Math.Round(Val(Datos(7)), 4)

                            f1.WriteLine()
                            f1.WriteLine("#finDetalle")
                            f1.WriteLine("fsc:descuento                   =                  0.0000")
                            f1.WriteLine("fsc:subTotal                    =" & Space(26 - Datos(5).Length) & Datos(5))
                            f1.WriteLine("fsc:total                       =" & Space(26 - Datos(7).Length) & Datos(7))
                            f1.WriteLine()
                            If (Val(Datos(6)) > 0 Or EsFinagil = False) Then
                                If Val(Datos(11)) <> 100 Then
                                    BanImpuesto = True
                                    f1.WriteLine("#Impuestos")
                                    f1.WriteLine("iim:totalImpuestosTrasladados   =" & Space(26 - Datos(6).Length) & Datos(6))
                                    f1.WriteLine("iit:impuesto                    =        IVA")
                                    f1.WriteLine("iit:tasa                        =        " & Datos(11))
                                    f1.WriteLine("iit:importe                     =" & Space(26 - Datos(6).Length) & Datos(6))
                                    f1.WriteLine()
                                End If
                            End If
                            If ISR > 0 Then
                                If BanImpuesto = False Then f1.WriteLine("#Impuestos")
                                f1.WriteLine("iim:totalImpuestosRetenidos     =" & Space(26 - ISR.ToString.Length) & ISR)
                                f1.WriteLine("iir:impuesto                    =        ISR")
                                f1.WriteLine("iir:importe                     =" & Space(26 - ISR.ToString.Length) & ISR)
                                f1.WriteLine()
                            End If

                            f1.WriteLine("#EntregaEn")
                            f1.WriteLine("aen:nombre                      =")
                            f1.WriteLine("aen:calle                       =")
                            f1.WriteLine("aen:noExterior                  =")
                            f1.WriteLine("aen:noInterior                  =")
                            f1.WriteLine("aen:colonia                     =")
                            f1.WriteLine("aen:localidad                   =")
                            f1.WriteLine("aen:referencia                  =")
                            f1.WriteLine("aen:municipio                   =")
                            f1.WriteLine("aen:estado                      =")
                            f1.WriteLine("aen:pais                        =")
                            f1.WriteLine("aen:codigopostal                =")
                            f1.WriteLine()
                            f1.WriteLine("#Totales")
                            f1.WriteLine("ato:subtotalSinDescuentoSinIva  =" & Space(26 - Datos(5).Length) & Datos(5))
                        If Moneda = "USD" Then
                            Datos(8) = Letras(Datos(7), Moneda)
                        End If
                        f1.WriteLine("ato:cantidadConLetra            =       " & Datos(8))
                            f1.WriteLine()
                            f1.WriteLine("agr:noCliente                   =       " & Datos(1))
                            f1.WriteLine("agr:fechaOrdenCompra            =")
                            f1.WriteLine("agr:fechaDeContraReciboMercancia=")
                            f1.WriteLine("agr:tipoMoneda                  =")
                            f1.WriteLine("agr:totalKilos                  =")
                            f1.WriteLine("agr:chofer                      =        " & OpcionCompraAF)
                            f1.WriteLine("agr:comentariosLeyenda          =        " & Datos(2))
                            f1.WriteLine("agr:LeyendaP                    =        EL PAGO DE ESTE DOCUMENTO SE HACE EN UNA SOLA EXHIBICION")
                            f1.WriteLine()
                            f1.WriteLine("adi:impresora                   =")
                            f1.WriteLine("adi:email                       =")
                            f1.WriteLine("adi:mailagente                  =")
                            f1.WriteLine("adi:ImpresoraLocal              =")

                            If MetodoPago = "NA" Then MetodoPago = ""
                            MetodoPago += " " & taTipar.SacaMetodo(MetodoPago)

                            f1.WriteLine("adi:Condicion                   =         " & Trim(MetodoPago))
                            If EnviarGisela = False Then
                                If EsArfin = True Then
                                    f1.WriteLine("adi:Mail1                       =         vcruz@finagil.com.mx")
                                    f1.WriteLine("adi:Mail2                       =         lgarcia@finagil.com.mx")
                                Else
                                    'CORREOS ADICIONALES++++++++++++++++++++++++++++
                                    taMail.Fill(tMail, cAnexo)
                                If tMail.Rows.Count > 0 Then
                                    cMail1 = ""
                                    cMail2 = ""
                                    For Each Rmail In tMail.Rows
                                        If Rmail.Correo1 > "" Then
                                            cMail1 += ";" & Rmail.Correo1
                                        End If
                                        If Rmail.Correo2 > "" Then
                                            cMail2 += ";" & Rmail.Correo2
                                        End If
                                    Next
                                    f1.WriteLine("adi:Mail1                       =         " & cMail1)
                                    f1.WriteLine("adi:Mail2                       =         " & cMail2)
                                    If InStr(Mail, "@") Then
                                        f1.WriteLine("adi:Mail3                       =         " & Mail)
                                    End If
                                Else
                                    ta.Fill(t, Datos(1))
                                    If t.Rows.Count > 0 Then
                                        f1.WriteLine("adi:Mail1                       =         " & Trim(t.Rows(0).Item("EMail1")))
                                        f1.WriteLine("adi:Mail2                       =         " & Trim(t.Rows(0).Item("EMail2")))
                                        If InStr(Mail, "@") Then
                                            f1.WriteLine("adi:Mail3                       =         " & Mail)
                                        Else
                                            f1.WriteLine("adi:Mail3                       =")
                                        End If
                                    Else
                                        If InStr(Mail, "@") Then
                                            f1.WriteLine("adi:Mail1                       =         " & Mail)
                                            f1.WriteLine("adi:Mail2                       =         ")
                                            f1.WriteLine("adi:Mail3                       =")
                                        Else
                                            f1.WriteLine("adi:Mail1                       =         ")
                                            f1.WriteLine("adi:Mail2                       =         ")
                                            f1.WriteLine("adi:Mail3                       =")
                                        End If
                                    End If
                                End If
                                    'CORREOS ADICIONALES++++++++++++++++++++++++++++
                                End If
                            Else
                            f1.WriteLine("adi:Mail1                       =         lhernandez@finagil.com.mx")
                            f1.WriteLine("adi:Mail2                       =         ")
                            If InStr(Mail, "@") Then
                                f1.WriteLine("adi:Mail3                       =         " & Mail)
                            Else
                                f1.WriteLine("adi:Mail3                       =")
                            End If
                        End If
                        If Datos(1).Trim = "05978 " Then
                            f1.WriteLine("adi:Mail4                       =         flen.estrada@ciasaconstruccion.com.mx")
                            f1.WriteLine("adi:Mail5                       =         administacion@ciasaconstruccion.com.mx")
                            f1.WriteLine("adi:Mail6                       =         ecacerest@lamoderna.com.mx")
                        Else
                            f1.WriteLine("adi:Mail4                       =")
                            f1.WriteLine("adi:Mail5                       =")
                            f1.WriteLine("adi:Mail6                       =")
                        End If


                            f1.WriteLine()
                            f1.WriteLine("agr:lineatexto1                 =        " & Arre(1, 1) & Arre(2, 1)) '& "*" & Arre(0, x))
                            f1.WriteLine("agr:lineatexto2                 =        " & Arre(1, 2) & Arre(2, 2)) '& "*" & Arre(0, x))
                            f1.WriteLine("agr:lineatexto3                 =        " & Arre(1, 3) & Arre(2, 3)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:lineatexto4                 =        " & Arre(1, 4) & Arre(2, 4)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:lineatexto5                 =        " & Arre(1, 5) & Arre(2, 5)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:lineatexto6                 =        " & Arre(1, 6) & Arre(2, 6)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:lineatexto7                 =        " & Arre(1, 7) & Arre(2, 7)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:lineatexto8                 =        " & Arre(1, 8) & Arre(2, 8)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:GalleT                      =        " & Arre(1, 9) & Arre(2, 9)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:GalleD                      =        " & Arre(1, 10) & Arre(2, 10)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:HarinT                      =        " & Arre(1, 11) & Arre(2, 11)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:HarinD                      =        " & Arre(1, 12) & Arre(2, 12)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:InstaT                      =        " & Arre(1, 13) & Arre(2, 13)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:InstaD                      =        " & Arre(1, 14) & Arre(2, 14)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:OtrosT                      =        " & Arre(1, 15) & Arre(2, 15)) '& "*" & Arre(0, x))
                            f1.WriteLine("aex:OtrosD                      =        " & Arre(1, 16) & Arre(2, 16)) '& "*" & Arre(0, x))
                        Case "Z1"
                            If EsNotaCredito = True And (Mid(Datos(3), 1, 1) <> "C") Then
                                f1.WriteLine()
                                f1.WriteLine("#finDetalle")
                                f1.WriteLine("fsc:descuento                   =                  0.0000")
                                f1.WriteLine("fsc:subTotal                    =" & Space(26 - suma.ToString.Length) & suma.ToString)
                                f1.WriteLine("fsc:total                       =" & Space(26 - suma.ToString.Length) & suma.ToString)
                                f1.WriteLine()
                                f1.WriteLine("#Impuestos")
                                f1.WriteLine("iim:totalImpuestosRetenidos     =")
                                f1.WriteLine("iim:totalImpuestosTrasladados   =                  0.0000")
                                f1.WriteLine()

                                If Val(Datos(6)) > 0 Or EsFinagil = False Then
                                    f1.WriteLine("#Impuestos")
                                    f1.WriteLine("iir:impuesto                    =        IVA")
                                    f1.WriteLine("iir:importe                     =        0.0000")
                                    f1.WriteLine("iit:impuesto                    =        IVA")
                                    f1.WriteLine("iit:tasa                        =        " & Datos(11))
                                    f1.WriteLine("iit:importe                     =" & Space(26 - Datos(6).Length) & Datos(6))
                                    f1.WriteLine()
                                End If

                                'f1.WriteLine("iir:impuesto                    =        IVA")
                                'f1.WriteLine("iir:importe                     =        0.0000")
                                'f1.WriteLine("iit:impuesto                    =        IVA")
                                'f1.WriteLine("iit:tasa                        =        ")
                                'f1.WriteLine("iit:importe                     =" & Space(26 - suma.ToString.Length) & suma.ToString)
                                'f1.WriteLine()

                                f1.WriteLine()
                                f1.WriteLine("#EntregaEn")
                                f1.WriteLine("aen:nombre                      =")
                                f1.WriteLine("aen:calle                       =")
                                f1.WriteLine("aen:noExterior                  =")
                                f1.WriteLine("aen:noInterior                  =")
                                f1.WriteLine("aen:colonia                     =")
                                f1.WriteLine("aen:localidad                   =")
                                f1.WriteLine("aen:referencia                  =")
                                f1.WriteLine("aen:municipio                   =")
                                f1.WriteLine("aen:estado                      =")
                                f1.WriteLine("aen:pais                        =")
                                f1.WriteLine("aen:codigopostal                =")
                                f1.WriteLine()
                                f1.WriteLine("#Totales")
                                f1.WriteLine("ato:subtotalSinDescuentoSinIva  =" & Space(26 - suma.ToString.Length) & suma.ToString)
                                f1.WriteLine("ato:cantidadConLetra            =       ") '& Datos(8))
                                f1.WriteLine()
                                f1.WriteLine("agr:noCliente                   =       " & Datos(1))
                                f1.WriteLine("agr:fechaOrdenCompra            =")
                                f1.WriteLine("agr:fechaDeContraReciboMercancia=")
                                f1.WriteLine("agr:tipoMoneda                  =")
                                f1.WriteLine("agr:totalKilos                  =")
                                f1.WriteLine("agr:telefonoCliente             =")
                                f1.WriteLine("agr:comentariosLeyenda          =        " & Datos(2))
                                f1.WriteLine("agr:LeyendaP                    =        EL PAGO DE ESTE DOCUMENTO SE HACE EN UNA SOLA EXHIBICION")
                                f1.WriteLine()
                                f1.WriteLine("adi:impresora                   =")
                                f1.WriteLine("adi:email                       =")
                                f1.WriteLine("adi:mailagente                  =")
                                f1.WriteLine("adi:ImpresoraLocal              =")
                                MetodoPago = taTipar.SacaMetodo(MetodoPago)
                                f1.WriteLine("adi:Condicion                   =         " & MetodoPago)
                                If EnviarGisela = False Then
                                    If EsArfin = True Then
                                        f1.WriteLine("adi:Mail1                       =         " & Mail)
                                        f1.WriteLine("adi:Mail2                       =         ")
                                    Else
                                        ta.Fill(t, Datos(1))

                                        If Trim(t.Rows(0).Item("EMail1")).Length < 3 Then t.Rows(0).Item("EMail1") = ""
                                        If Trim(t.Rows(0).Item("EMail2")).Length < 3 Then t.Rows(0).Item("EMail2") = ""

                                        f1.WriteLine("adi:Mail1                       =         " & Trim(t.Rows(0).Item("EMail1")))
                                        f1.WriteLine("adi:Mail2                       =         " & Trim(t.Rows(0).Item("EMail2")))
                                    End If
                                Else
                                    f1.WriteLine("adi:Mail1                       =         lhernandez@finagil.com.mx")
                                    f1.WriteLine("adi:Mail2                       =         ")
                                End If
                                f1.WriteLine("adi:Mail3                       =")
                                f1.WriteLine("adi:Mail4                       =")
                                f1.WriteLine("adi:Mail5                       =")
                                f1.WriteLine("adi:Mail6                       =")
                                f1.WriteLine()
                                f1.WriteLine("agr:lineatexto1                 =        " & Arre(1, 1) & Arre(2, 1))
                                f1.WriteLine("agr:lineatexto2                 =        " & Arre(1, 2) & Arre(2, 2))
                                f1.WriteLine("agr:lineatexto3                 =        " & Arre(1, 3) & Arre(2, 3))
                                f1.WriteLine("aex:lineatexto4                 =        " & Arre(1, 4) & Arre(2, 4))
                                f1.WriteLine("aex:lineatexto5                 =        " & Arre(1, 5) & Arre(2, 5))
                                f1.WriteLine("aex:lineatexto6                 =        " & Arre(1, 6) & Arre(2, 6))
                                f1.WriteLine("aex:lineatexto7                 =        " & Arre(1, 7) & Arre(2, 7))
                                f1.WriteLine("aex:lineatexto8                 =        " & Arre(1, 8) & Arre(2, 8))
                                f1.WriteLine("aex:GalleT                      =        " & Arre(1, 9) & Arre(2, 9))
                                f1.WriteLine("aex:GalleD                      =        " & Arre(1, 10) & Arre(2, 10))
                                f1.WriteLine("aex:HarinT                      =        " & Arre(1, 11) & Arre(2, 11))
                                f1.WriteLine("aex:HarinD                      =        " & Arre(1, 12) & Arre(2, 12))
                                f1.WriteLine("aex:InstaT                      =        " & Arre(1, 13) & Arre(2, 13))
                                f1.WriteLine("aex:InstaD                      =        " & Arre(1, 14) & Arre(2, 14))
                                f1.WriteLine("aex:OtrosT                      =        " & Arre(1, 15) & Arre(2, 15))
                                f1.WriteLine("aex:OtrosD                      =        " & Arre(1, 16) & Arre(2, 16))
                                f1.WriteLine("aex:refFacturaNo                =        " & Datos(3) & Datos(4))
                            Else 'no es nota de credito
                                If Tipar = "F" Then
                                    Datos(7) = Datos(7)
                                End If
                            End If
                    End Select
                End While
                f1.Close()
                f2.Close()
                If SinDet = "SinDetalle" Then
                    EnviaError(GeneraFactura.My.Settings.MailError, "Sin Detalle CFDI " & F(i).Name, "Sin Detalle CFDI " & F(i).Name)
                End If
            'Catch ex As Exception
            '    EnviaError(GeneraFactura.My.Settings.MailError, ex.Message, "Error CFDI " & F(i).Name)
            '    Console.WriteLine(GeneraFactura.My.Settings.NoPasa & F(i).Name)
            '    Console.WriteLine(F(i).FullName)
            '    File.Copy(F(i).FullName, GeneraFactura.My.Settings.NoPasa & F(i).Name, True)
            '    f2.Close()
            'End Try
        Next
        'Catch ex As Exception
        '    EnviaError(GeneraFactura.My.Settings.MailError, ex.Message, "Error CFDI " & F(i).Name)
        'End Try
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
            ROWheader._29_FormaPago = "99"
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
                TasaIVA = TasaIVACliente
                TipoImpuesto = TasaIVACliente
                If Trim(rr.Concepto) = "EFECTIVO" _
                Or Trim(rr.Concepto) = "AGROQUIMICOS" _
                    Or Trim(rr.Concepto) = "EFECTIVO2" _
                Or Trim(rr.Concepto) = "EFECTIVO 2" _
                Or Trim(rr.Concepto) = "VALE" _
                Or Trim(rr.Concepto) = "ASISTENCIA" _
                Or Trim(rr.Concepto) = "INTEGRACION" _
                Or Trim(rr.Concepto) = "ANALISIS DE SUELOS" Then
                    Fega += Math.Round(rr.FEGA, 4)
                Else
                    ContLin += 1
                    Fega += Math.Round(rr.FEGA, 4)
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
                    ROWdetail._16_Linea_Cod_Articulo = "84101700"
                    ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                    ROWdetail._2_Impuesto_Descripcion = "TR"
                    ROWdetail._3_Impuesto_Monto_base = rr.Importe.ToString("n2")
                    ROWdetail._5_Impuesto_Clave = "002"
                    ROWdetail._6_Impuesto_Tasa = "Tasa"
                    If TipoImpuesto = "Excento" Then
                        ROWdetail._7_Impuesto_Porcentaje = "EXE"
                        ROWdetail._4_Impuesto_Monto_Impuesto = 0
                    Else
                        ROWdetail._7_Impuesto_Porcentaje = TasaIVA
                        ROWdetail._4_Impuesto_Monto_Impuesto = Math.Round(rr.Importe * (TasaIVA / 100), 2)
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
                TasaIVA = TasaIVACliente
                ROWdetail._1_Linea_Descripcion = Concep
                ROWdetail._2_Linea_Cantidad = 1
                ROWdetail._3_Linea_Unidad = "E48"
                ROWdetail._4_Linea_PrecioUnitario = Fega.ToString("n2")
                ROWdetail._5_Linea_Importe = Fega.ToString("n2")
                ROWdetail._16_Linea_Cod_Articulo = "84101700"
                ROWdetail._1_Impuesto_TipoImpuesto = "Impuesto"
                ROWdetail._2_Impuesto_Descripcion = "TR"
                ROWdetail._3_Impuesto_Monto_base = Fega.ToString("n2")
                ROWdetail._4_Impuesto_Monto_Impuesto = Math.Round(Fega * (TasaIVA / 100), 2)
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
            ROWheader._54_Monto_SubTotal = SubTT.ToString("n2")
            ROWheader._55_Monto_IVA = IVA.ToString("n2")
            ROWheader._56_Monto_Total = Math.Round(SubTT + IVA, 2)
            ROWheader._193_Monto_TotalImp_Trasladados = IVA.ToString("n2")
            ROWheader._100_Letras_Monto_Total = Letras(Math.Round(ROWheader._56_Monto_Total, 2), "MXN")
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

    Sub GeneraArchivosEXternas()
        Dim x As Integer
        Dim EsNotaCredito As Boolean = False
        Dim Impo As String
        Dim Arre(2, 20) As String
        Dim cad As String = ""
        Dim RFC As String = ""
        Dim Razon As String = ""
        Dim SubTT As Double
        Dim IVA As Double
        Dim TOt As Double
        Dim PUnitario As Double

        Dim f1 As System.IO.StreamWriter
        Dim Facturas As New GeneraFactura.ProduccionDSTableAdapters.FacturasExternasTableAdapter
        Dim FAC As New GeneraFactura.ProduccionDS.FacturasExternasDataTable
        Dim Detalles As New GeneraFactura.ProduccionDSTableAdapters.FacturasExternasDETTableAdapter
        Dim DET As New GeneraFactura.ProduccionDS.FacturasExternasDETDataTable
        Dim tasa As Double
        Dim SubTTAux As Double
        Dim IVAAux As Double
        Dim Concep As String
        Dim RFC_Cli As String = ""
        Dim TipoPersona As String = "F"
        Dim MetodoPago As String = ""
        Dim taMetodo As New ProduccionDSTableAdapters.LlavesTableAdapter


        Facturas.Fill(FAC)
        For Each r As GeneraFactura.ProduccionDS.FacturasExternasRow In FAC.Rows
            Console.WriteLine("Generando CFDI Facturas Externas..." & r.Factura)
            For x = 0 To 20
                Arre(1, x) = ""
                Arre(2, x) = ""
            Next
            x = 0
            SubTT = 0
            TOt = 0
            IVA = 0
            SubTTAux = 0
            IVAAux = 0
            RFC_Cli = ""
            f1 = New System.IO.StreamWriter(GeneraFactura.My.Settings.RutaCFDI & "CFDI-" & r.Serie & "-" & r.Factura & ".txt", False, Text.Encoding.GetEncoding(1252))


            f1.WriteLine("#InicioComprobante")
            If r.Serie = "C" Then
                EsNotaCredito = True
                f1.WriteLine("idn:documento                   =        NCREDITO")
            Else
                EsNotaCredito = False
                f1.WriteLine("idn:documento                   =        FACTURA")
            End If



            f1.WriteLine("idn:tipofactura                 =        FINANZAS")
            f1.WriteLine("idn:planta                      =        TOLUCA")
            f1.WriteLine("idn:tipodocto                   =        NACIONAL")
            f1.WriteLine("idn:documanager                 =        FINAGIL")
            f1.WriteLine()
            f1.WriteLine("fsc:serie                       =        " & r.Serie)
            f1.WriteLine("fsc:folio                       =        " & r.Factura)
            f1.WriteLine("fsc:fecha                       =        " & r.fecha.ToString("yyyy-MM-dd") & "T" & r.fecha.ToString("HH:mm:ss"))
            f1.WriteLine("fsc:formaDePago                 =        PAGO EN UNA SOLA EXHIBICION")
            f1.WriteLine("fsc:noCertificado               =        00001000000202240016")
            f1.WriteLine("fsc:condicionesDePago           =        Contado")
            f1.WriteLine("fsc:motivoDescuento             =	       ")
            f1.WriteLine("fsc:TipoCambio                  =        1")
            f1.WriteLine("fsc:Moneda                      =        " & r.Moneda)
            If EsNotaCredito = True Then
                f1.WriteLine("fsc:tipoDeComprobante           =        Egreso")
            Else
                f1.WriteLine("fsc:tipoDeComprobante           =        Ingreso")
            End If

            MetodoPago = taMetodo.SacaID_Metodo(Trim(r.MetodoPago))
            f1.WriteLine("fsc:metodoDePago                =        " & MetodoPago)
            If MetodoPago = "NA" Then MetodoPago = ""

            f1.WriteLine("fsc:LugarExpedicion             =        Toluca, México")
            f1.WriteLine("fsc:NumCtaPago                  =        " & Trim(r.Cuenta))
            f1.WriteLine()
            f1.WriteLine("#Emisor")
            f1.WriteLine("fem:rfc                         =        FIN940905AX7")
            f1.WriteLine("fem:nombre                      =        FINAGIL S.A. DE C.V, SOFOM E.N.R")
            f1.WriteLine("fed:calle                       =        Leandro Valle")
            f1.WriteLine("fed:noExterior                  =        402")
            f1.WriteLine("fed:noInterior                  =        ")
            f1.WriteLine("fed:colonia                     =        Reforma y F.F.C.C")
            f1.WriteLine("fed:localidad                   =        Toluca")
            f1.WriteLine("fed:municipio                   =        Toluca")
            f1.WriteLine("fed:estado                      =        Estado de Mexico")
            f1.WriteLine("fed:pais                        =        Mexico")
            f1.WriteLine("fed:codigoPostal                =        50070")
            f1.WriteLine()
            f1.WriteLine("#ExpendidoEn")
            f1.WriteLine("fee:calle                       =        Leandro Valle")
            f1.WriteLine("fee:noExterior                  =        402")
            f1.WriteLine("fee:noInterior                  =        ")
            f1.WriteLine("fee:colonia                     =        Reforma y F.F.C.C")
            f1.WriteLine("fee:localidad                   =        Toluca")
            f1.WriteLine("fee:municipio                   =        Toluca")
            f1.WriteLine("fee:estado                      =        Estado de Mexico")
            f1.WriteLine("fee:pais                        =        Mexico")
            f1.WriteLine("fee:codigopostal                =        50070")
            f1.WriteLine("fer:regimen                     =        REGIMEN GENERAL DE LEY PERSONAS MORALES")
            f1.WriteLine()
            f1.WriteLine("#Receptor")
            RFC = Trim(r.RFC)
            'If RFC.Length = 10 Then RFC += "000"
            'If RFC.Length < 10 Then RFC = "XAXX010101000"
            'If RFC.Length > 13 Then RFC = "XAXX010101000"
            If IsNumeric(Mid(RFC, 4, 1)) Then
                TipoPersona = "M"
            Else
                TipoPersona = "F"
            End If
            RFC = ValidaRFC(RFC, TipoPersona)
            f1.WriteLine("fre:rfc                         =        " & RFC)
            If RFC = "SDA070613KU6" Then
                Razon = """SERVICIO DAYCO"" SA DE CV"
            Else
                Razon = r.Nombre
            End If
            f1.WriteLine("fre:nombre                      =        " & Razon)
            f1.WriteLine("frd:calle                       =        " & r.Calle)
            f1.WriteLine("frd:noExterior                  =        ")
            f1.WriteLine("frd:noInterior                  =        ")
            f1.WriteLine("frd:colonia                     =        " & r.Colonia)
            f1.WriteLine("frd:localidad                   =        ")
            f1.WriteLine("frd:municipio                   =        " & r.Municipio)
            f1.WriteLine("frd:estado                      =        " & r.Estado)
            f1.WriteLine("frd:pais                        =        México")
            f1.WriteLine("frd:codigopostal                =        " & r.CP)
            f1.WriteLine()
            f1.WriteLine("#Detalle")
            'f1.WriteLine(" dco:cant  dco:unit   dco:noId         dco:desc                       dco:vUni      dco:impo       dcd:line dcd:unid dcd:kg     dcd:prom  dcd:coAl")
            f1.WriteLine(" dco:cant  dco:unit   dco:noId         dco:desc                                                              dco:vUni      dco:impo       dcd:line dcd:unid dcd:kg     dcd:prom  dcd:coAl")
            f1.WriteLine()

            Detalles.Fill(DET, r.Serie, r.Factura)
            For Each rr As GeneraFactura.ProduccionDS.FacturasExternasDETRow In DET.Rows
                x += 1
                SubTTAux = Math.Round(rr.Importe, 4)
                IVAAux = Math.Round(rr.Iva, 4)
                Concep = Trim(rr.Detalle)
                PUnitario = Math.Round(rr.Unitario, 4)
                Arre(1, x) = Mid(Trim(Concep), 1, 70)
                Impo = SubTTAux
                'f1.WriteLine(Space(8 - rr.Cantidad.ToString.Length) & rr.Cantidad.ToString() & "   UNI " & Space(24) & Mid(Trim(Concep), 1, 30) & Space(30 - Trim(Mid(Concep, 1, 30)).Length) & Space(13 - PUnitario.ToString.Length) & PUnitario & Space(15 - Impo.Length) & Impo & Space(20 - Impo.Length) & "  ")
                f1.WriteLine(Space(8 - rr.Cantidad.ToString.Length) & rr.Cantidad.ToString() & "   " & rr.Unidad.Trim & " " & Space(24) & Mid(Trim(Concep), 1, 70) & Space(70 - Trim(Mid(Concep, 1, 70)).Length) & Space(13 - PUnitario.ToString.Length) & PUnitario & Space(15 - Impo.Length) & Impo & Space(20 - Impo.Length) & "  ")
                IVA += IVAAux
                TOt += SubTTAux + IVAAux
                SubTT += SubTTAux
                cad = "*"
                If SubTT <> 0 Then
                    Arre(2, x) = cad & Format(SubTTAux, "#,##0.00")
                Else
                    Arre(2, x) = cad
                End If

                Facturas.Facturar(r.Serie, r.Factura, rr.Consec)
                Select Case UCase(Trim(rr.TasaIva))
                    Case "16 %"
                        tasa = 16
                    Case "0 %"
                        tasa = 0
                    Case "EXCENTO"
                        tasa = -1
                End Select
            Next

            If IVA > 0 Then tasa = 16

            f1.WriteLine()
            f1.WriteLine("#finDetalle")
            f1.WriteLine("fsc:descuento                   =                  0.0000")
            f1.WriteLine("fsc:subTotal                    =" & Space(26 - SubTT.ToString.Length) & SubTT.ToString)
            f1.WriteLine("fsc:total                       =" & Space(26 - TOt.ToString.Length) & TOt.ToString)
            f1.WriteLine()
            If IVA > 0 Or tasa = 0 Then
                f1.WriteLine("#Impuestos")
                f1.WriteLine("iim:totalImpuestosRetenidos     =")
                f1.WriteLine("iim:totalImpuestosTrasladados   =" & Space(26 - IVA.ToString.Length) & IVA.ToString)
                f1.WriteLine()
                f1.WriteLine("iir:impuesto                    =        IVA")
                f1.WriteLine("iir:importe                     =        0.0000")
                f1.WriteLine("iit:impuesto                    =        IVA")
                f1.WriteLine("iit:tasa                        =        " & tasa)
                f1.WriteLine("iit:importe                     =" & Space(26 - IVA.ToString.Length) & IVA.ToString)
                f1.WriteLine()
            End If
            f1.WriteLine()
            f1.WriteLine("#EntregaEn")
            f1.WriteLine("aen:nombre                      =")
            f1.WriteLine("aen:calle                       =")
            f1.WriteLine("aen:noExterior                  =")
            f1.WriteLine("aen:noInterior                  =")
            f1.WriteLine("aen:colonia                     =")
            f1.WriteLine("aen:localidad                   =")
            f1.WriteLine("aen:referencia                  =")
            f1.WriteLine("aen:municipio                   =")
            f1.WriteLine("aen:estado                      =")
            f1.WriteLine("aen:pais                        =")
            f1.WriteLine("aen:codigopostal                =")
            f1.WriteLine()
            f1.WriteLine("#Totales")
            f1.WriteLine("ato:subtotalSinDescuentoSinIva  =" & Space(26 - SubTT.ToString.Length) & SubTT.ToString)
            f1.WriteLine("ato:cantidadConLetra            =       " & Letras(TOt.ToString, "MXN"))
            f1.WriteLine()
            f1.WriteLine("agr:noCliente                   =       " & r.RFC)
            f1.WriteLine("agr:fechaOrdenCompra            =")
            f1.WriteLine("agr:fechaDeContraReciboMercancia=       ")
            f1.WriteLine("agr:tipoMoneda                  =")
            f1.WriteLine("agr:totalKilos                  =")
            f1.WriteLine("agr:telefonoCliente             =")
            If EsNotaCredito = True Then
                f1.WriteLine("agr:comentariosLeyenda          =        Nota de Crédito")
            Else
                f1.WriteLine("agr:comentariosLeyenda          =        Factura Manual")
            End If

            f1.WriteLine("agr:LeyendaP                    =        EL PAGO DE ESTE DOCUMENTO SE HACE EN UNA SOLA EXHIBICION")
            f1.WriteLine()
            f1.WriteLine("adi:impresora                   =")
            f1.WriteLine("adi:email                       =")
            f1.WriteLine("adi:mailagente                  =")
            f1.WriteLine("adi:ImpresoraLocal              =")
            f1.WriteLine("adi:Condicion                   =        " & MetodoPago & Trim(r.MetodoPago))

            f1.WriteLine("adi:Mail1                       =         " & Trim(r.Mail1))
            f1.WriteLine("adi:Mail2                       =         " & Trim(r.Mail2))
            f1.WriteLine("adi:Mail3                       =         vcruz@finagil.com.mx;lhernandez@finagil.com.mx")
            f1.WriteLine("adi:Mail4                       =")
            f1.WriteLine("adi:Mail5                       =")
            f1.WriteLine("adi:Mail6                       =")
            f1.WriteLine()
            f1.WriteLine("agr:lineatexto1                 =        " & Arre(1, 1) & Arre(2, 1))
            f1.WriteLine("agr:lineatexto2                 =        " & Arre(1, 2) & Arre(2, 2))
            f1.WriteLine("agr:lineatexto3                 =        " & Arre(1, 3) & Arre(2, 3))
            f1.WriteLine("aex:lineatexto4                 =        " & Arre(1, 4) & Arre(2, 4))
            f1.WriteLine("aex:lineatexto5                 =        " & Arre(1, 5) & Arre(2, 5))
            f1.WriteLine("aex:lineatexto6                 =        " & Arre(1, 6) & Arre(2, 6))
            f1.WriteLine("aex:lineatexto7                 =        " & Arre(1, 7) & Arre(2, 7))
            f1.WriteLine("aex:lineatexto8                 =        " & Arre(1, 8) & Arre(2, 8))
            f1.WriteLine("aex:GalleT                      =        " & Arre(1, 9) & Arre(2, 9))
            f1.WriteLine("aex:GalleD                      =        " & Arre(1, 10) & Arre(2, 10))
            f1.WriteLine("aex:HarinT                      =        " & Arre(1, 11) & Arre(2, 11))
            f1.WriteLine("aex:HarinD                      =        " & Arre(1, 12) & Arre(2, 12))
            f1.WriteLine("aex:InstaT                      =        " & Arre(1, 13) & Arre(2, 13))
            f1.WriteLine("aex:InstaD                      =        " & Arre(1, 14) & Arre(2, 14))
            f1.WriteLine("aex:OtrosT                      =        " & Arre(1, 15) & Arre(2, 15))
            f1.WriteLine("aex:OtrosD                      =        " & Arre(1, 16) & Arre(2, 16))
            f1.Close()

        Next

    End Sub

    Sub LecturaPrevia(RutaArchivo As String, NombreArchivo As String)
        OpcionCompraAF = ""
        Dim Numero As Integer = 1
        Dim f2 As System.IO.StreamReader
        Dim taTipar As New GeneraFactura.ProduccionDSTableAdapters.LlavesTableAdapter
        Dim Linea As String
        Dim Datos() As String
        Dim Tipar As String = ""
        f2 = New System.IO.StreamReader(RutaArchivo, Text.Encoding.GetEncoding(1252))
        Try

            While Not f2.EndOfStream
                Linea = f2.ReadLine
                Datos = Linea.Split("|")
                If Numero = 1 And Datos.Length > 3 Then
                    Tipar = taTipar.Tipar(Datos(2))
                    If Tipar = "F" Then
                        Select Case Datos(0)
                            Case "M1"
                            Case "H3"
                            Case "D1"
                            Case "S1"
                            Case "Z1"
                                If Tipar = "F" Then
                                    OpcionCompraAF = Datos(7)
                                    Numero = 0
                                End If
                        End Select
                    Else
                        Exit While
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

End Module
