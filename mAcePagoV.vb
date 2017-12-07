Option Explicit On

Imports System.Data.SqlClient
Imports System.Math
Imports System.IO

' Esta función realiza el desglose de la aplicación de un pago a una factura determinada.
' Es llamada por la forma frmAcepagoIVF la cual le envía como argumentos:
' cAnexo (el número de anexo)
' cLetra (la letra del vencimiento)
' nMontoPago (el monto pagado)
' nMoratorios
' nIvaMoratorios
' cBanco
' cCheque
' dtMovimientos (la cual contendrá los movimientos contables del pago y es pasada por referencia, por lo que se ve afectada por este proceso)
' cFechaAplicacion (que es la fecha con la que se generará la factura electrónica y con la que se registrará el ingreso)
' cFechaPago (esta es la fecha en la que el cliente realizó el pago y determina si proceden o no intereses moratorios)
' ¿Cuál fecha se registra en FACTURAS como último pago?
' cSerie
' nRecibo
' Es importante comentar que el monto del pago NO es el importe total del pago, sino la parte del pago 
' que será aplicado a esta factura.

Module mAcepagov

    ' Declaración de variables de alcance privado

    Dim cFeven As String = ""
    Dim cPlazo As String = ""
    Dim cTipar As String = ""
    Dim nBaseFEGA As Decimal = 0
    Dim nImporteFac As Decimal = 0
    Dim nImporteFEGA As Decimal = 0
    Dim nIvaFEGA As Decimal = 0
    Dim nPagado As Decimal = 0
    Dim nSaldoFac As Decimal = 0
    Dim nTasaIVA As Decimal = 0         ' Tasa de IVA que le corresponde al cliente de acuerdo a su domicilio fiscal

    Private Structure Conceptos
        Public Concepto As String
        Public Importe As Decimal
        Public Porcentaje As Decimal
        Public Iva As Decimal
    End Structure

    Dim aConcepto As New Conceptos()

    Public Sub Acepagov(ByVal cAnexo As String, ByVal cLetra As String, ByVal nMontoPago As Decimal, ByVal nMoratorios As Decimal, ByVal nIvaMoratorios As Decimal, ByVal cBanco As String, ByVal cCheque As String, ByRef dtMovimientos As DataTable, ByVal cFecha As String, ByVal cFechaPago As String, ByVal cSerie As String, ByVal nRecibo As Decimal, InstrumentoMonetario As String, FechaProc As Date, MetodoPago As String)

        ' Declaración de variables de conexión ADO .NET

        Dim cnAgil As New SqlConnection(My.Settings.ConnectionStringFACTURA)
        Dim cm1 As New SqlCommand()
        Dim daFacturas As New SqlDataAdapter(cm1)

        Dim dsAgil As New DataSet
        Dim dtPagos As New DataTable("Pagos")
        Dim drPago As DataRow
        Dim drFactura As DataRow
        Dim drMovimiento As DataRow

        'Dim strInsert As String
        'Dim strUpdate As String

        ' Declaración de variables de datos

        Dim aConceptos As New ArrayList()
        Dim cCalle As String = ""
        Dim cCatal As String = ""
        Dim cCliente As String = ""
        Dim cColonia As String = ""
        Dim cConcepto As String = ""
        Dim cCopos As String = ""
        Dim cCuentaPago As String = ""
        Dim cDelegacion As String = ""
        Dim cEstado As String = ""
        Dim cFepag As String = ""
        Dim cFormaPago As String = ""
        Dim cNombre As String = ""
        Dim cObserva As String = ""
        Dim cPrevio As String = ""
        Dim cRenglon As String = ""
        Dim cRfc As String = ""
        Dim cTipmon As String = ""
        Dim cTipos As String = ""
        Dim i As Integer
        'Dim lCredito As Boolean
        Dim nAbonoCartera As Decimal = 0
        Dim nAbonoOtros As Decimal = 0
        Dim nBonifica As Decimal = 0
        Dim nCapitalEquipo As Decimal = 0
        Dim nCapitalOtros As Decimal = 0
        Dim nCapitalSeguro As Decimal = 0
        Dim nEsp As Decimal = 0
        Dim nFactura As Decimal = 0
        Dim nImporte As Decimal = 0
        Dim nInteres As Decimal = 0
        Dim nInteresOtros As Decimal = 0
        Dim nInteresSEG As Decimal = 0
        Dim nIva As Decimal = 0
        Dim nIvaCapital As Decimal = 0
        Dim nIvaInteres As Decimal = 0
        Dim nIvaInteresSeg As Decimal = 0
        Dim nIvaOtros As Decimal = 0
        Dim nPagoConcepto As Decimal = 0
        Dim nPlazo As Byte = 0
        Dim nPrevioCartera As Decimal = 0
        Dim nPrevioOtros As Decimal = 0
        Dim nSeguroVida As Decimal = 0
        Dim nSubTotal As Decimal = 0
        Dim nTotal As Decimal = 0

        ' Luego creo la tabla dtPagos

        dtPagos.Columns.Add("Anexo", Type.GetType("System.String"))
        dtPagos.Columns.Add("Letra", Type.GetType("System.String"))
        dtPagos.Columns.Add("Tipos", Type.GetType("System.String"))
        dtPagos.Columns.Add("Fepag", Type.GetType("System.String"))
        dtPagos.Columns.Add("Concepto", Type.GetType("System.String"))
        dtPagos.Columns.Add("Importe", Type.GetType("System.Decimal"))
        dtPagos.Columns.Add("Tipmon", Type.GetType("System.String"))
        dtPagos.Columns.Add("Banco", Type.GetType("System.String"))
        dtPagos.Columns.Add("IVA", Type.GetType("System.Decimal"))
        dtPagos.Clear()

        ' El siguiente Stored Procedure trae todos los atributos de la factura correspondiente al anexo y al vencimiento dados

        With cm1
            .CommandType = CommandType.StoredProcedure
            .CommandText = "Acepagov1"
            .Connection = cnAgil
            .Parameters.Add("@Anexo", SqlDbType.NVarChar)
            .Parameters(0).Value = cAnexo
            .Parameters.Add("@Letra", SqlDbType.NVarChar)
            .Parameters(1).Value = cLetra
        End With

        ' Llenar el DataSet a través del DataAdapter, lo cual abre y cierra la conexión

        daFacturas.Fill(dsAgil, "Facturas")

        ' Teóricamente, la tabla Facturas del dataset debe contener un solo registro

        drFactura = dsAgil.Tables("Facturas").Rows(0)

        ' Datos del Cliente

        cNombre = drFactura("Descr")
        cRfc = drFactura("Rfc")
        cCalle = RTrim(drFactura("Calle"))
        cColonia = RTrim(drFactura("Colonia"))
        cDelegacion = RTrim(drFactura("Delegacion"))
        cEstado = RTrim(drFactura("Estado"))
        cCopos = RTrim(drFactura("Copos"))
        cCliente = drFactura("Cliente")

        For i = 1 To 5
            Select Case i
                Case 1
                    If RTrim(drFactura("CuentadePago1")) <> "0" And RTrim(drFactura("FormadePago1")) <> "EFECTIVO" Then
                        cCuentaPago = drFactura("CuentadePago1")
                        cFormaPago = RTrim(drFactura("FormadePago1"))
                    ElseIf RTrim(drFactura("CuentadePago1")) = "0" And RTrim(drFactura("FormadePago1")) = "EFECTIVO" Then
                        cCuentaPago = "NO IDENTIFICABLE"
                        cFormaPago = RTrim(drFactura("FormadePago1"))
                    End If
                Case 2
                    If RTrim(drFactura("CuentadePago2")) <> "0" And RTrim(drFactura("FormadePago2")) <> "EFECTIVO" Then
                        cCuentaPago = cCuentaPago & "," & drFactura("CuentadePago2")
                        cFormaPago = cFormaPago & "," & RTrim(drFactura("FormadePago2"))
                    ElseIf RTrim(drFactura("CuentadePago2")) = "0" And RTrim(drFactura("FormadePago2")) = "EFECTIVO" Then
                        cCuentaPago = cCuentaPago & "," & "NO IDENTIFICABLE"
                        cFormaPago = cFormaPago & "," & RTrim(drFactura("FormadePago2"))
                    End If
                Case 3
                    If RTrim(drFactura("CuentadePago3")) <> "0" And RTrim(drFactura("FormadePago3")) <> "EFECTIVO" Then
                        cCuentaPago = cCuentaPago & "," & drFactura("CuentadePago3")
                        cFormaPago = cFormaPago & "," & RTrim(drFactura("FormadePago3"))
                    ElseIf RTrim(drFactura("CuentadePago3")) = "0" And RTrim(drFactura("FormadePago3")) = "EFECTIVO" Then
                        cCuentaPago = cCuentaPago & "," & "NO IDENTIFICABLE"
                        cFormaPago = cFormaPago & "," & RTrim(drFactura("FormadePago3"))
                    End If
                Case 4
                    If RTrim(drFactura("CuentadePago4")) <> "0" And RTrim(drFactura("FormadePago4")) <> "EFECTIVO" Then
                        cCuentaPago = cCuentaPago & "," & drFactura("CuentadePago4")
                        cFormaPago = cFormaPago & "," & RTrim(drFactura("FormadePago4"))
                    ElseIf RTrim(drFactura("CuentadePago4")) = "0" And RTrim(drFactura("FormadePago4")) = "EFECTIVO" Then
                        cCuentaPago = cCuentaPago & "," & "NO IDENTIFICABLE"
                        cFormaPago = cFormaPago & "," & RTrim(drFactura("FormadePago4"))
                    End If
                Case 5
                    If cCuentaPago = "" And cFormaPago = "" Then
                        cCuentaPago = "NO IDENTIFICABLE"
                        cFormaPago = "NO IDENTIFICABLE"
                    End If
            End Select
        Next

        ' Esta es una nueva forma de calcular el plazo que implementé a partir del 17 de octubre de 2011
        ' para determinar correctamente el plazo para frecuencias de pago diferentes a mensual

        nPlazo = 0
        CuentaPagos(cAnexo, nPlazo)

        cPlazo = CStr(nPlazo)

        ' Datos de la Factura

        cTipar = drFactura("Tipar")
        nTasaIVA = Round(drFactura("TasaIVA") / 100, 2)

        nFactura = drFactura("Factura")
        cFeven = drFactura("Feven")
        cFepag = drFactura("Fepag")
        nImporteFac = drFactura("ImporteFac")
        nSaldoFac = drFactura("SaldoFac")
        nPagado = nImporteFac - nSaldoFac

        nImporteFEGA = drFactura("ImporteFEGA")
        nBaseFEGA = Round(nImporteFEGA / (1 + nTasaIVA), 2)
        nIvaFEGA = Round(nImporteFEGA - nBaseFEGA, 2)

        nSeguroVida = drFactura("SeguroVida")

        nIvaOtros = drFactura("IvaOt")
        nInteresOtros = drFactura("InteresOt") + drFactura("VarOt")
        nCapitalOtros = drFactura("CapitalOt")

        nIvaInteres = drFactura("IvaPr")
        nIvaInteresSeg = drFactura("IvaSe")
        nInteres = drFactura("IntPr") + drFactura("VarPr")
        nInteresSEG = drFactura("IntSe") + drFactura("VarSe")
        nCapitalSeguro = drFactura("RenSe")
        nCapitalEquipo = drFactura("RenPr") - drFactura("IntPr")
        nIvaCapital = drFactura("IvaCapital")
        nBonifica = drFactura("Bonifica")

        ' El siguiente arreglo de estructuras debe ser inicializado por cada factura que se pague ya que de ello depende la jerarquización del pago.
        ' Además su conformación depende del tipo de producto (es diferente para Arrendamiento Puro).

        If cFeven >= "20131001" Or cTipar = "P" Then

            ' Nueva forma de prelación

            If nImporteFEGA > 0 Then
                aConcepto.Concepto = "FEGA"
                aConcepto.Importe = nBaseFEGA
                aConcepto.Porcentaje = nBaseFEGA / (nBaseFEGA + nIvaFEGA)
                aConcepto.Iva = nIvaFEGA
                aConceptos.Add(aConcepto)
            End If

            If nSeguroVida > 0 Then
                aConcepto.Concepto = "SEGURO DE VIDA"
                aConcepto.Importe = nSeguroVida
                aConcepto.Porcentaje = 1
                aConcepto.Iva = 0
                aConceptos.Add(aConcepto)
            End If

            If nInteresOtros > 0 Then
                aConcepto.Concepto = "INTERES OTROS ADEUDOS"
                aConcepto.Importe = nInteresOtros
                aConcepto.Porcentaje = nInteresOtros / (nInteresOtros + nIvaOtros)
                aConcepto.Iva = nIvaOtros
                aConceptos.Add(aConcepto)
            End If

            If nCapitalOtros > 0 Then
                aConcepto.Concepto = "CAPITAL OTROS ADEUDOS"
                aConcepto.Importe = nCapitalOtros
                aConcepto.Porcentaje = 1
                aConcepto.Iva = 0
                aConceptos.Add(aConcepto)
            End If

            If cTipar = "P" Then

                If drFactura("IntSe") + drFactura("VarSe") > 0 Then
                    aConcepto.Concepto = "INTERES SEGURO"
                    aConcepto.Importe = drFactura("IntSe") + drFactura("VarSe")
                    aConcepto.Porcentaje = (drFactura("IntSe") + drFactura("VarSe")) / (drFactura("IntSe") + drFactura("VarSe") + drFactura("IvaSe"))
                    aConcepto.Iva = drFactura("IvaSe")
                    aConceptos.Add(aConcepto)
                End If

                If drFactura("Rense") > 0 Then
                    aConcepto.Concepto = "CAPITAL SEGURO"
                    aConcepto.Importe = drFactura("Rense")
                    aConcepto.Iva = 0
                    aConcepto.Porcentaje = 1
                    aConceptos.Add(aConcepto)
                End If

                If drFactura("RenPr") + drFactura("VarPr") > 0 Then
                    aConcepto.Concepto = "PAGO DE RENTA"
                    aConcepto.Importe = drFactura("RenPr") + drFactura("VarPr")
                    aConcepto.Porcentaje = (drFactura("RenPr") + drFactura("VarPr")) / (drFactura("RenPr") + drFactura("VarPr") + drFactura("IvaCapital") + drFactura("IvaPr"))
                    aConcepto.Iva = drFactura("IvaCapital") + drFactura("IvaPr")
                    aConceptos.Add(aConcepto)
                End If

            ElseIf cTipar = "B" Then


                If nCapitalEquipo > 0 Then
                    aConcepto.Concepto = "MENSUALIDAD"
                    aConcepto.Importe = nCapitalEquipo
                    aConcepto.Porcentaje = nCapitalEquipo / (nCapitalEquipo + nIvaCapital)
                    aConcepto.Iva = nIvaCapital
                    aConceptos.Add(aConcepto)
                End If
            Else


                If nInteres > 0 Then
                    aConcepto.Concepto = "INTERESES"
                    aConcepto.Importe = nInteres
                    aConcepto.Porcentaje = nInteres / (nInteres + 0)
                    aConcepto.Iva = nIvaInteres
                    aConceptos.Add(aConcepto)
                End If

                If drFactura("IntSe") + drFactura("VarSe") > 0 Then 'ECT separacion de interes Seguro
                    aConcepto.Concepto = "INTERES SEGURO"
                    aConcepto.Importe = nInteresSEG
                    aConcepto.Porcentaje = (drFactura("IntSe") + drFactura("VarSe")) / (drFactura("IntSe") + drFactura("VarSe") + drFactura("IvaSe"))
                    aConcepto.Iva = nIvaInteresSeg
                    aConceptos.Add(aConcepto)
                End If

                If nCapitalSeguro > 0 Then
                    aConcepto.Concepto = "CAPITAL SEGURO"
                    aConcepto.Importe = nCapitalSeguro
                    aConcepto.Porcentaje = 1
                    aConcepto.Iva = 0
                    aConceptos.Add(aConcepto)
                End If

                If nCapitalEquipo > 0 Then
                    aConcepto.Concepto = "CAPITAL EQUIPO"
                    aConcepto.Importe = nCapitalEquipo
                    aConcepto.Porcentaje = nCapitalEquipo / (nCapitalEquipo + nIvaCapital)
                    aConcepto.Iva = nIvaCapital
                    aConceptos.Add(aConcepto)
                End If

            End If

            For Each aConcepto In aConceptos

                If nMontoPago > 0 And aConcepto.Importe > 0 Then

                    cConcepto = aConcepto.Concepto
                    If nMontoPago >= aConcepto.Importe / aConcepto.Porcentaje Then
                        ' Pago completo del importe
                        nPagoConcepto = aConcepto.Importe
                        nMontoPago = nMontoPago - aConcepto.Importe
                        aConcepto.Importe = 0
                    End If

                    drPago = dtPagos.NewRow()
                    drPago("Anexo") = cAnexo
                    drPago("Letra") = cLetra
                    drPago("Tipos") = "2"
                    drPago("Fepag") = cFecha
                    drPago("Tipmon") = "01"
                    drPago("Banco") = cBanco
                    If cTipar = "B" Then
                        drPago("Concepto") = cConcepto + " No. " + cLetra + "/0" + cPlazo
                    Else
                        drPago("Concepto") = cConcepto + " VENCIMIENTO " + cLetra + "/0" + cPlazo
                    End If

                    drPago("Importe") = nPagoConcepto
                    drPago("IVA") = aConcepto.Iva
                    dtPagos.Rows.Add(drPago)

                End If

            Next
        End If


        cPrevio = "N"
        If Mid(cFeven, 1, 6) > Mid(cFecha, 1, 6) Then
            cPrevio = "S"
        End If

        ' Aquí se descuentan los pagos que se hayan hecho a la factura exceptuando Moratorios y el IVA de los Moratorios
        ' a fin de determinar el saldo de la factura

        For Each drPago In dtPagos.Rows

            If cTipar = "B" Then
                If InStr(drPago("Concepto"), "No.") > 0 Then
                    cObserva = Mid(drPago("Concepto"), 1, InStr(drPago("Concepto"), "No.", CompareMethod.Text) - 2)
                End If
            Else
                cObserva = Mid(drPago("Concepto"), 1, InStr(drPago("Concepto"), "VENCIMIENTO", CompareMethod.Text) - 2)
            End If

            nImporte = drPago("Importe")

            If nImporte <> 0 And InStr(cObserva, "MORATORIOS", CompareMethod.Text) = 0 Then

                nSaldoFac = Round(nSaldoFac - nImporte, 2)

            End If

        Next

        nPrevioCartera = 0
        nPrevioOtros = 0
        nAbonoCartera = 0
        nAbonoOtros = 0

        ' En este punto está hecha la aplicación del pago, por lo que procedo a actualizar la historia de pagos y a generar los asientos contables 

        nSubTotal = 0
        nIva = 0
        nTotal = 0

        dsAgil.Tables.Remove("Facturas")

        Dim stmFactura As New FileStream("\\server-nas\FacturasCFDI\ORGtxt\FACTURA_" & cSerie & "_" & nRecibo & ".txt", FileMode.Create, FileAccess.Write, FileShare.None)
        Dim stmWriter As New StreamWriter(stmFactura, System.Text.Encoding.Default)

        stmWriter.WriteLine("H1|" & FechaProc.ToShortDateString & "|" & MetodoPago & "|99|")

        cRenglon = "H3|" & cCliente & "|" & Mid(cAnexo, 1, 5) & "/" & Mid(cAnexo, 6, 4) & "|" & cSerie & "|" & nRecibo & "|" & Trim(cNombre) & "|" &
        Trim(cCalle) & "|||" & Trim(cColonia) & "|" & Trim(cDelegacion) & "|" & Trim(cEstado) & "|" & cCopos & "|" & cCuentaPago & "|" & cFormaPago & "|MEXICO|" & Trim(cRfc) & "|M.N.|" &
        "|FACTURA|" & cCliente & "|LEANDRO VALLE 402||REFORMA Y FFCCNN|TOLUCA|ESTADO DE MEXICO|50070|MEXICO|" & cAnexo & "|" & cLetra & "|"

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

        For Each drPago In dtPagos.Rows
            If InStr(Trim(drPago("Concepto")), "IVA ") Then
                Continue For
            End If
            cRenglon = "D1|" & cCliente & "|" & Mid(cAnexo, 1, 5) & "/" & Mid(cAnexo, 6, 4) & "|" & cSerie & "|" & nRecibo & "|1|||" & Trim(drPago("Concepto")) & "||" & drPago("Importe") & "|" & drPago("Iva")
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
        Next

        stmWriter.Flush()
        stmFactura.Flush()
        stmFactura.Close()

        Try

        Catch eException As Exception
            MsgBox(eException.Message, MsgBoxStyle.Critical, "Mensaje de Error")
        End Try

    End Sub

    Private Sub PPTotal(ByRef dtPagos As DataTable, ByVal aConceptos As ArrayList, ByVal drFactura As DataRow, ByVal cFecha As String, ByVal cBanco As String)

        ' Declaración de variables de datos

        Dim cAnexo As String = ""
        Dim cLetra As String = ""
        Dim drPago As DataRow
        Dim nBonifica As Decimal = 0
        Dim nCapitalEquipo As Decimal = 0
        Dim nIvaCapital As Decimal = 0
        Dim nPorieq As Decimal = 0

        cAnexo = drFactura("Anexo")
        nCapitalEquipo = drFactura("RenPr") - drFactura("IntPr")
        nIvaCapital = drFactura("IvaCapital")
        nBonifica = drFactura("Bonifica")
        cLetra = drFactura("Letra")
        nPorieq = drFactura("Porieq")
        nImporteFEGA = drFactura("ImporteFEGA")


        For Each aConcepto In aConceptos

            If aConcepto.Importe > 0 Then

                If aConcepto.Concepto = "CAPITAL EQUIPO" Then
                    drPago = dtPagos.NewRow()
                    drPago("Anexo") = cAnexo
                    drPago("Letra") = cLetra
                    drPago("Tipos") = "2"
                    drPago("Fepag") = cFecha
                    drPago("Tipmon") = "01"
                    drPago("Banco") = cBanco
                    drPago("Concepto") = "CAPITAL EQUIPO VENCIMIENTO " + cLetra + "/0" + cPlazo
                    drPago("Importe") = nCapitalEquipo
                    dtPagos.Rows.Add(drPago)
                    If nBonifica > 0 Then
                        drPago = dtPagos.NewRow()
                        drPago("Anexo") = cAnexo
                        drPago("Letra") = cLetra
                        drPago("Tipos") = "2"
                        drPago("Fepag") = cFecha
                        drPago("Tipmon") = "01"
                        drPago("Banco") = cBanco
                        drPago("Concepto") = "APLICACION DEPOSITO vs CAPITAL VENCIMIENTO " + cLetra + "/0" + cPlazo
                        drPago("Importe") = Round(-nBonifica / (1 + (nPorieq / 100)), 2)
                        dtPagos.Rows.Add(drPago)
                        nBonifica = Round(nBonifica - Round(nBonifica / (1 + (nPorieq / 100)), 2), 2)
                    End If
                    If nIvaCapital > 0 Then
                        drPago = dtPagos.NewRow()
                        drPago("Anexo") = cAnexo
                        drPago("Letra") = cLetra
                        drPago("Tipos") = "2"
                        drPago("Fepag") = cFecha
                        drPago("Tipmon") = "01"
                        drPago("Banco") = cBanco
                        drPago("Concepto") = "IVA CAPITAL VENCIMIENTO " + cLetra + "/0" + cPlazo
                        drPago("Importe") = nIvaCapital
                        dtPagos.Rows.Add(drPago)
                        If nBonifica > 0 Then
                            drPago = dtPagos.NewRow()
                            drPago("Anexo") = cAnexo
                            drPago("Letra") = cLetra
                            drPago("Tipos") = "2"
                            drPago("Fepag") = cFecha
                            drPago("Tipmon") = "01"
                            drPago("Banco") = cBanco
                            drPago("Concepto") = "APLICACION DEPOSITO vs IVA CAPITAL VENCIMIENTO " + cLetra + "/0" + cPlazo
                            drPago("Importe") = -nBonifica
                            dtPagos.Rows.Add(drPago)
                            nBonifica = 0
                        End If
                    End If

                ElseIf aConcepto.Concepto = "FEGA" Then

                    drPago = dtPagos.NewRow()
                    drPago("Anexo") = cAnexo
                    drPago("Letra") = cLetra
                    drPago("Tipos") = "2"
                    drPago("Fepag") = cFecha
                    drPago("Tipmon") = "01"
                    drPago("Banco") = cBanco
                    drPago("Concepto") = "FEGA VENCIMIENTO " + cLetra + "/0" + cPlazo
                    drPago("Importe") = nBaseFEGA
                    dtPagos.Rows.Add(drPago)

                    drPago = dtPagos.NewRow()
                    drPago("Anexo") = cAnexo
                    drPago("Letra") = cLetra
                    drPago("Tipos") = "2"
                    drPago("Fepag") = cFecha
                    drPago("Tipmon") = "01"
                    drPago("Banco") = cBanco
                    drPago("Concepto") = "IVA FEGA VENCIMIENTO " + cLetra + "/0" + cPlazo
                    drPago("Importe") = nIvaFEGA
                    dtPagos.Rows.Add(drPago)

                Else

                    drPago = dtPagos.NewRow()
                    drPago("Anexo") = cAnexo
                    drPago("Letra") = cLetra
                    drPago("Tipos") = "2"
                    drPago("Fepag") = cFecha
                    drPago("Tipmon") = "01"
                    drPago("Banco") = cBanco
                    If cTipar = "B" Then
                        drPago("Concepto") = aConcepto.Concepto + " No. " + cLetra + "/0" + cPlazo
                    Else
                        drPago("Concepto") = aConcepto.Concepto + " VENCIMIENTO " + cLetra + "/0" + cPlazo
                    End If

                    drPago("Importe") = aConcepto.Importe
                    dtPagos.Rows.Add(drPago)

                End If

            End If

        Next

    End Sub

    Public Sub CuentaPagos(ByVal cAnexo As String, ByRef nPagos As Byte)
        Dim cnAgil As New SqlConnection(My.Settings.ConnectionStringFACTURA)
        Dim cm1 As New SqlCommand()
        nPagos = 0
        With cm1
            .CommandType = CommandType.Text
            .CommandText = "SELECT COUNT(*) FROM Edoctav WHERE Anexo = '" & cAnexo & "' AND IndRec = 'S' AND Nufac <> 9999999 AND Nufac <> 7777777"
            .Connection = cnAgil
        End With

        cnAgil.Open()
        nPagos = CInt(cm1.ExecuteScalar())
        cnAgil.Close()
        cm1.Dispose()
        cnAgil.Dispose()

    End Sub

End Module
