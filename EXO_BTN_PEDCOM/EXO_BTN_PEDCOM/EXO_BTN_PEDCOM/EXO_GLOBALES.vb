Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Public Class EXO_GLOBALES
#Region "Funciones formateos datos"
    Public Shared Function TextToDbl(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As Double
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"

        TextToDbl = 0

        Try
            sValorAux = sValor

            If oObjGlobal.SBOApp.ClientType = BoClientType.ct_Desktop Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            TextToDbl = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function DblNumberToText(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByVal sValor As String) As String
        Dim sNumberDouble As String = "0"

        DblNumberToText = "0"

        Try
            If sValor <> "" Then
                If oObjGlobal.refDi.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.refDi.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = sValor
                Else 'Decimales USA
                    sNumberDouble = sValor.Replace(",", ".")
                End If
            End If

            DblNumberToText = sNumberDouble


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function FormateaString(ByVal dato As Object, ByVal tam As Integer) As String
        Dim retorno As String = String.Empty

        If dato IsNot Nothing Then
            retorno = dato.ToString
        End If

        If retorno.Length > tam Then
            retorno = retorno.Substring(0, tam)
        End If

        Return retorno.PadRight(tam, CChar(" "))
    End Function
    Public Shared Function FormateaNumero(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
    Public Shared Function FormateaNumeroSinPunto(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
    Public Shared Function FormateaNumeroconSigno(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        Else
            retorno = " " & retorno
        End If
        Return retorno
    End Function
#End Region
#Region "REPORTS"
    Public Shared Sub GetCrystalReportFile(ByVal oCompany As SAPbobsCOM.Company, ByVal sFormatoImp As String, ByVal sOutFileName As String)
        Dim oBlobParams As SAPbobsCOM.BlobParams = Nothing
        Dim oKeySegment As SAPbobsCOM.BlobTableKeySegment = Nothing
        Dim oBlob As SAPbobsCOM.Blob = Nothing
        Dim sContent As String = ""
        Dim obuff() As Byte = Nothing

        Try
            oBlobParams = CType(oCompany.GetCompanyService().GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), SAPbobsCOM.BlobParams)

            oBlobParams.Table = "RDOC"
            oBlobParams.Field = "Template"

            oKeySegment = oBlobParams.BlobTableKeySegments.Add()
            oKeySegment.Name = "DocCode"

            oKeySegment.Value = sFormatoImp
            oBlob = oCompany.GetCompanyService().GetBlob(oBlobParams)
            sContent = oBlob.Content

            obuff = Convert.FromBase64String(sContent)

            Using oFile As New System.IO.FileStream(sOutFileName, System.IO.FileMode.Create)
                oFile.Write(obuff, 0, obuff.Length)

                oFile.Close()
            End Using

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlobParams, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oKeySegment, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oBlob, Object))
        End Try
    End Sub
    Public Shared Function GenerarCrystal(ByVal oobjglobal As EXO_UIAPI.EXO_UIAPI, ByVal strRutaInforme As String, ByVal sFileCrystal As String,
                                          ByVal sPath As String, sDocEntry As String, ByVal sDocNum As String, ByVal sDocDate As String, ByVal sTipo As String) As String
        Dim oCRReport As ReportDocument = Nothing
        'objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fecha:" & sDocDate & " - .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Dim dFecha As Date = sDocDate
        Dim sFilePDF As String = sTipo & "_" & sDocNum '& "_" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")


        'Guardar ficheros 
        If Not System.IO.Directory.Exists(sPath) Then
            System.IO.Directory.CreateDirectory(sPath)
        End If
        Try
            oobjglobal.SBOApp.StatusBar.SetText("(EXO) - Preparando fichero - " & sFilePDF & " - .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'generar el rpt
            GenerarCrystal = ""
            'objGlobal.SBOApp.StatusBar.SetText("(EXO) - 1", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oCRReport = New ReportDocument
            oobjglobal.SBOApp.StatusBar.SetText("(EXO) - Cargando Crystal - " & strRutaInforme & sFileCrystal & " - .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oCRReport.Load(strRutaInforme & sFileCrystal)

            oCRReport.SetParameterValue("DocKey@", sDocEntry)

            'PONER USUARIO Y CONTRASEÑA

            Dim conrepor As CrystalDecisions.Shared.DataSourceConnections = oCRReport.DataSourceConnections
            conrepor(0).SetConnection(oobjglobal.compañia.Server, oobjglobal.compañia.CompanyDB, oobjglobal.refDi.SQL.usuarioSQL, oobjglobal.refDi.SQL.claveSQL)
            'objGlobal.SBOApp.StatusBar.SetText("(EXO) - 2", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sFilePDF = sPath & sFilePDF & ".pdf"
            oCRReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sFilePDF)
            oobjglobal.SBOApp.StatusBar.SetText("(EXO) - Pdf creado : " & sFilePDF, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            GenerarCrystal = sFilePDF
        Catch ex As Exception
            Throw ex
        Finally
            oCRReport.Close()
            oCRReport.Dispose()
        End Try
    End Function
    Public Shared Function Enviarmail(ByVal oobjglobal As EXO_UIAPI.EXO_UIAPI, sCuerpo As String, sRutahtml As String, dirmail As String, sProveedor As String, sProveedorNom As String, sFichero As String, sficheroCompra As String) As Boolean
        Dim correo As New System.Net.Mail.MailMessage()
        Dim adjunto As System.Net.Mail.Attachment
        Dim StrFirma As String = ""
        Dim htmbody As New System.Text.StringBuilder()
        Enviarmail = False
        Dim sMail As String = oobjglobal.funcionesUI.refDi.OGEN.valorVariable("COM_Mail")
        Dim sCMail As String = oobjglobal.funcionesUI.refDi.OGEN.valorVariable("COM_CMail")
        Dim sMail_Usuario As String = oobjglobal.funcionesUI.refDi.OGEN.valorVariable("COM_US")
        Dim sMail_PS As String = oobjglobal.funcionesUI.refDi.OGEN.valorVariable("COM_PS")
        Dim sMail_SMTP As String = oobjglobal.funcionesUI.refDi.OGEN.valorVariable("COM_SMTP")
        Dim sMail_PORT As String = oobjglobal.funcionesUI.refDi.OGEN.valorVariable("COM_PORT")
        Dim oCC As New Net.Mail.MailAddressCollection
        correo.From = New System.Net.Mail.MailAddress(sMail, "Brompton House")
        If sCMail.Trim <> "" Then
            correo.CC.Add(sCMail.Trim)
        End If

        If dirmail <> "" Then
            Dim delimitadores() As String = {";", "+", "-", ":"}
            Dim vectoraux() As String = dirmail.Split(delimitadores, StringSplitOptions.None)
            For Each item As String In vectoraux
                correo.To.Add(item)
            Next
        End If
        correo.Subject = "Nuevo Pedido Procesado"

        If sFichero <> "" Then
            adjunto = New System.Net.Mail.Attachment(sFichero)
            correo.Attachments.Add(adjunto)
        End If

        If sficheroCompra <> "" Then
            adjunto = New System.Net.Mail.Attachment(sficheroCompra)
            correo.Attachments.Add(adjunto)
        End If


        Dim FicheroCab As String = ""

        correo.Body = sCuerpo
        correo.IsBodyHtml = True
        correo.Priority = System.Net.Mail.MailPriority.Normal

        Dim smtp As New System.Net.Mail.SmtpClient
        smtp.Host = sMail_SMTP
        smtp.Port = sMail_PORT
        smtp.UseDefaultCredentials = True
        smtp.Credentials = New System.Net.NetworkCredential(sMail_Usuario, sMail_PS)
        smtp.EnableSsl = True

        'smtp.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network

        Try
            smtp.Send(correo)

            correo.Dispose()
            oobjglobal.SBOApp.StatusBar.SetText("Correo enviado a " & sProveedorNom & " con mail: " & dirmail & ", adjuntando ficheros:" & sFichero & ", " & sficheroCompra, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Enviarmail = True
        Catch ex As Exception
            oobjglobal.SBOApp.StatusBar.SetText("No se ha podido envial mail a " & sProveedorNom & " con mail: " & dirmail & ", adjuntando ficheros:" & sFichero & ", " & sficheroCompra & ". Error: " & ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Enviarmail = False
        Finally
        End Try
        Return True
    End Function
#End Region
End Class
