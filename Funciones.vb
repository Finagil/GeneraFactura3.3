Imports System.Net.Mail
Module Funciones

    Public Function Letras(ByVal numero As String, ByVal moneda As String) As String

        'Declaración de variables de datos
        Dim NumAux As Double = numero
        numero = NumAux.ToString("f2")
        Dim entero As String
        Dim cMillones As String
        Dim cMiles As String
        Dim cCentenas As String
        Dim cCant_Cent As String = ""
        Dim cCant_Mil As String = ""
        Dim cCant_Mill As String = ""
        Dim dec As String
        Dim cCant As String
        Dim flag2 As String = "N"
        Dim x As Integer

        'Dividir parte entera y decimal

        For x = 1 To Len(numero)
            If Mid(numero, x, 1) = "." Then
                flag2 = "S"
            Else
                If flag2 = "N" Then
                    entero = entero + Mid(numero, x, 1)
                Else
                    dec = dec + Mid(numero, x, 1)
                End If
            End If
        Next x

        If Len(dec) = 1 Then dec = dec & "0"

        If Len(entero) > 6 Then
            cCentenas = Mid(entero, (Len(entero) + 1) - 3, 3)
            cMiles = Mid(entero, (Len(entero) + 1) - 6, 3)
            cMillones = Mid(entero, 1, Len(entero) - 6)
        ElseIf Len(entero) <= 6 And Len(entero) > 3 Then
            cCentenas = Mid(entero, (Len(entero) + 1) - 3, 3)
            cMiles = Mid(entero, 1, Len(entero) - 3)
        ElseIf Len(entero) <= 3 Then
            cCentenas = Mid(entero, 1, Len(entero))
        End If

        'proceso de conversión

        cCant_Cent = Cambio(cCentenas)
        cCant_Mil = Cambio(cMiles)
        cCant_Mill = Cambio(cMillones)

        'Asigna la palabra millón

        If Trim(cCant_Mill) <> "" And Trim(cCant_Mill) <> "CERO" Then
            If Trim(cCant_Mill) = "UN" Then
                cCant = cCant_Mill & "MILLON "
            Else
                cCant = cCant_Mill & " MILLONES "
            End If
        End If

        'Asigna la palabra mil

        If Trim(cCant_Mil) <> "" And Trim(cCant_Mil) <> "CERO" Then
            If Trim(cCant_Mil) = "UN" And Trim(cCant) <> "" Then
                cCant = cCant & " MIL "
            ElseIf Trim(cCant_Mil) = "UN" And Trim(cCant) = "" Then
                cCant = "MIL "
            Else
                cCant = cCant & cCant_Mil & "MIL "
            End If
        End If

        'Asigna la palabra correspondiente a als unidades

        If Trim(cCant) <> "" And Trim(cCant_Cent) <> "CERO" Then
            cCant = cCant & cCant_Cent
        ElseIf Trim(cCant) = "" And Trim(cCant_Cent) <> "CERO" Then
            cCant = cCant_Cent
        ElseIf Trim(cCant) = "" And Trim(cCant_Cent) = "CERO" Then
            cCant = cCant_Cent
        End If

        'Se une la parte entera y la parte decimal
        If moneda <> "USD" Then
            If dec <> "" Then
                If Trim(cCant_Mill) <> "" And Trim(cCant_Mil) = "" Or Trim(cCant_Mil) = "CERO" Then
                    Letras = "(" & cCant & "DE PESOS " & dec & "/100 M.N.)"
                Else
                    Letras = "(" & cCant & "PESOS " & dec & "/100 M.N.)"
                End If
            Else
                If Trim(cCant_Mill) <> "" And Trim(cCant_Mil) = "" Or Trim(cCant_Mil) = "CERO" Then
                    Letras = "(" & cCant & "DE PESOS 00/100 M.N.)"
                Else
                    Letras = "(" & cCant & "PESOS 00/100 M.N.)"
                End If
            End If
        ElseIf moneda = "USD" Then
            If dec <> "" Then
                If Trim(cCant_Mill) <> "" And Trim(cCant_Mil) = "" Or Trim(cCant_Mil) = "CERO" Then
                    Letras = "(" & cCant & "DE DOLAES " & dec & "/100 " & moneda & ")"
                Else
                    Letras = "(" & cCant & "DOLARES " & dec & "/100 " & moneda & ")"
                End If
            Else
                If Trim(cCant_Mill) <> "" And Trim(cCant_Mil) = "" Or Trim(cCant_Mil) = "CERO" Then
                    Letras = "(" & cCant & "DE DOLARES 00/100 " & moneda & ")"
                Else
                    Letras = "(" & cCant & "DOLARES 00/100 " & moneda & ")"
                End If
            End If
        End If

    End Function

    Function Cambio(ByVal Cantidad As String) As String

        Dim y As Integer
        Dim num As Integer
        Dim flag As String = "N"
        Dim palabras As String = ""

        For y = Len(Cantidad) To 1 Step -1

            num = Len(Cantidad) - (y - 1)

            Select Case y

                Case 3

                    'Asigna las palabras para las centenas

                    Select Case Mid(Cantidad, num, 1)
                        Case "1"
                            If Mid(Cantidad, num + 1, 1) = "0" And Mid(Cantidad, num + 2, 1) = "0" Then
                                palabras = palabras & "CIEN "
                            Else
                                palabras = palabras & "CIENTO "
                            End If
                        Case "2"
                            palabras = palabras & "DOSCIENTOS "
                        Case "3"
                            palabras = palabras & "TRESCIENTOS "
                        Case "4"
                            palabras = palabras & "CUATROCIENTOS "
                        Case "5"
                            palabras = palabras & "QUINIENTOS "
                        Case "6"
                            palabras = palabras & "SEISCIENTOS "
                        Case "7"
                            palabras = palabras & "SETECIENTOS "
                        Case "8"
                            palabras = palabras & "OCHOCIENTOS "
                        Case "9"
                            palabras = palabras & "NOVECIENTOS "
                    End Select
                Case 2

                    'Asigna las palabras para las decenas 

                    Select Case Mid(Cantidad, num, 1)

                        Case "0"
                            flag = "N"
                        Case "1"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                flag = "S"
                                palabras = palabras & "DIEZ "
                            End If
                            If Mid(Cantidad, num + 1, 1) = "1" Then
                                flag = "S"
                                palabras = palabras & "ONCE "
                            End If
                            If Mid(Cantidad, num + 1, 1) = "2" Then
                                flag = "S"
                                palabras = palabras & "DOCE "
                            End If
                            If Mid(Cantidad, num + 1, 1) = "3" Then
                                flag = "S"
                                palabras = palabras & "TRECE "
                            End If
                            If Mid(Cantidad, num + 1, 1) = "4" Then
                                flag = "S"
                                palabras = palabras & "CATORCE "
                            End If
                            If Mid(Cantidad, num + 1, 1) = "5" Then
                                flag = "S"
                                palabras = palabras & "QUINCE "
                            End If
                            If Mid(Cantidad, num + 1, 1) > "5" Then
                                flag = "N"
                                palabras = palabras & "DIECI"
                            End If
                        Case "2"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "VEINTE "
                                flag = "S"
                            Else
                                palabras = palabras & "VEINTI"
                                flag = "N"
                            End If
                        Case "3"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "TREINTA "
                                flag = "S"
                            Else
                                palabras = palabras & "TREINTA Y "
                                flag = "N"
                            End If
                        Case "4"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "CUARENTA "
                                flag = "S"
                            Else
                                palabras = palabras & "CUARENTA Y "
                                flag = "N"
                            End If
                        Case "5"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "CINCUENTA "
                                flag = "S"
                            Else
                                palabras = palabras & "CINCUENTA Y "
                                flag = "N"
                            End If
                        Case "6"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "SESENTA "
                                flag = "S"
                            Else
                                palabras = palabras & "SESENTA Y "
                                flag = "N"
                            End If
                        Case "7"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "SETENTA "
                                flag = "S"
                            Else
                                palabras = palabras & "SETENTA Y "
                                flag = "N"
                            End If
                        Case "8"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "OCHENTA "
                                flag = "S"
                            Else
                                palabras = palabras & "OCHENTA Y "
                                flag = "N"
                            End If
                        Case "9"
                            If Mid(Cantidad, num + 1, 1) = "0" Then
                                palabras = palabras & "NOVENTA "
                                flag = "S"
                            Else
                                palabras = palabras & "NOVENTA Y "
                                flag = "N"
                            End If
                    End Select
                Case 1

                    'Asigna las palabras para las unidades

                    Select Case Mid(Cantidad, num, 1)
                        Case "1"
                            If flag = "N" Then
                                If y = 1 Then
                                    palabras = palabras & "UN "
                                Else
                                    palabras = palabras & "UN "
                                End If
                            End If
                        Case "2"
                            If flag = "N" Then palabras = palabras & "DOS "
                        Case "3"
                            If flag = "N" Then palabras = palabras & "TRES "
                        Case "4"
                            If flag = "N" Then palabras = palabras & "CUATRO "
                        Case "5"
                            If flag = "N" Then palabras = palabras & "CINCO "
                        Case "6"
                            If flag = "N" Then palabras = palabras & "SEIS "
                        Case "7"
                            If flag = "N" Then palabras = palabras & "SIETE "
                        Case "8"
                            If flag = "N" Then palabras = palabras & "OCHO "
                        Case "9"
                            If flag = "N" Then palabras = palabras & "NUEVE "
                        Case "0"
                            If Trim(palabras) = "" Then
                                If flag = "N" Then palabras = palabras & "CERO "
                            End If
                    End Select
            End Select

        Next y
        Cambio = palabras

    End Function

    Public Sub EnviacORREO(ByVal Para As String, ByVal Mensaje As String, ByVal Asunto As String, de As String, Optional Attach As String = "")
        Dim Mensage As New MailMessage(Trim(de), Trim(Para), Trim(Asunto), Mensaje)
        Dim Cliente As New SmtpClient("smtp01.cmoderna.com", 26)
        Try
            Mensage.IsBodyHtml = True
            If Attach.Trim.Length > 0 Then
                Dim Att As New Attachment("\\server-raid\TmpFinagil\" & Attach)
                Mensage.Attachments.Add(Att)
            End If
            Cliente.Send(Mensage)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            EscribeLOG(ex.Message)
        End Try
    End Sub

    Public Sub EscribeLOG(Mensaje)
        Dim f As New IO.StreamWriter("C:\TMP\Log_Correos.txt", True)
        f.WriteLine(Mensaje)
        f.Close()
    End Sub

End Module
