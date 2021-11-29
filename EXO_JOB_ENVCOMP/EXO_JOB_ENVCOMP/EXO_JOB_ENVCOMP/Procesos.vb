Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml.Serialization
Imports CrystalDecisions.CrystalReports.Engine
Imports SAPbobsCOM

Public Class Procesos

#Region "Pedidos al Proveedor"
    Public Shared Sub Pedidos_Proveedor(ByRef oLog As EXO_Log.EXO_Log)
        Dim sError As String = ""
        Dim sPath As String = My.Application.Info.DirectoryPath.ToString : Dim oFiles() As String = Nothing
        Dim oDBSAP As SqlConnection = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim sSQL As String = ""
        Dim oDtProveedores As System.Data.DataTable = New System.Data.DataTable

        Dim ns As XmlSerializerNamespaces = Nothing
        Dim sw As StringWriter = Nothing
        Dim xmlSerializador As XmlSerializer = Nothing
        Dim strRutaInforme As String = ""
        Dim sFicheroPDF As String = ""
        Dim sRutaFicheroPdf As String = ""
        Try
            oLog.escribeMensaje("Conectando a SQL...", EXO_Log.EXO_Log.Tipo.advertencia)
            Conexiones.Connect_SQLServer(oDBSAP, "SQL", oLog)
            oLog.escribeMensaje("Conectando a la compañia...", EXO_Log.EXO_Log.Tipo.advertencia)
            Conexiones.Connect_Company(oCompany, oLog, "DI")
            oLog.escribeMensaje("Ha conectado con la compañia " & oCompany.CompanyName & ".", EXO_Log.EXO_Log.Tipo.advertencia)

            strRutaInforme = My.Application.Info.DirectoryPath.ToString & "\03.RPT\"
            sSQL = "SELECT DISTINCT T1.""DocEntry"", T1.""DocNum"", T4.""CardCode"", T4.""CardName"", T5.""E_MailL"" ""Email"" "
            sSQL &= " From ""ORDR"" T1 "
            sSQL &= " LEFT JOIN ""OWOR"" T2 ON T1.""DocEntry""=T2.""OriginAbs"" "
            sSQL &= " INNER JOIN (SELECT T0.""DocEntry"", Case when T0.""U_ECI_FAB"" is null then T1.""CardCode"" else T0.""U_ECI_FAB"" end as ""CodFab""  "
            sSQL &= " From ""OWOR"" T0 Left JOIN  ""OITM"" T1 ON T0.""ItemCode""=T1.""ItemCode"" "
            sSQL &= " WHERE T0.""STATUS""<>'C') T3 ON T2.""DocEntry""= T3.""DocEntry"" "
            sSQL &= " INNER JOIN ""OCRD"" T4 ON T3.""CodFab""= T4.""CardCode"" "
            sSQL &= " INNER JOIN  ""OCPR"" T5 On  T4.""CardCode""= T5.""CardCode"" And T4.""CntctPrsn""= T5.""Name""  "
            sSQL &= " WHERE ""U_EXO_EMAIL""='N' "
            sSQL &= " Order by  T1.""DocEntry"", T4.""CardCode"" "

            Conexiones.FillDtDB(oDBSAP, oDtProveedores, sSQL)
            If oDtProveedores.Rows.Count > 0 Then
                For i As Integer = 0 To oDtProveedores.Rows.Count - 1
                    oLog.escribeMensaje("Creando report de Pedidos Al proveedor " & oDtProveedores.Rows.Item(i).Item("CardName").ToString, EXO_Log.EXO_Log.Tipo.advertencia)
                    sRutaFicheroPdf = GenerarCrystal(strRutaInforme, "Pedido.rpt", sPath, oDtProveedores.Rows.Item(i).Item("DocEntry").ToString,
                                                     oDtProveedores.Rows.Item(i).Item("DocNum").ToString, oDtProveedores.Rows.Item(i).Item("CardCode").ToString,
                                                     oDtProveedores.Rows.Item(i).Item("CardName").ToString, oLog)

                    If sRutaFicheroPdf <> "" Then
                        'Como se ha generado el crystal, procedemos a su envío 
                        Enviarmail(sPath & "\02.HTML\", oDtProveedores.Rows.Item(i).Item("Email").ToString,
                                   oDtProveedores.Rows.Item(i).Item("CardCode").ToString, oDtProveedores.Rows.Item(i).Item("CardName").ToString, sRutaFicheroPdf, oLog)
#Region "Actualizar a enviado"
                        'Actualizamos el documento a enviado
                        sSQL = "UPDATE ""ORDR"" "
                        sSQL &= " SET ""U_EXO_EMAIL""='Y' "
                        sSQL &= " WHERE ""DocEntry""='" & oDtProveedores.Rows.Item(i).Item("DocEntry").ToString & "' "
                        Conexiones.ExecuteSQLDB(oDBSAP, sSQL)
                        oLog.escribeMensaje(" Se actualiza el Pedido de compra Nº" & oDtProveedores.Rows.Item(i).Item("DocEntry").ToString & " indicando Enviado.")
#End Region
#Region "Guardar en Hco el fichero creado"
                        'Guardar fichero en Hcos.
                        If Not System.IO.Directory.Exists(sPath & "\" & Conexiones._sEmpresa & "\ENVIADOS\HCOS") Then
                            System.IO.Directory.CreateDirectory(sPath & "\" & Conexiones._sEmpresa & "\ENVIADOS\HCOS")
                        End If
                        sFicheroPDF = Path.GetFileName(sRutaFicheroPdf)
                        'oLog.escribeMensaje("Guardando Archivo en Hcos....", EXO_Log.EXO_Log.Tipo.advertencia)
                        FicheroaHistorico(sRutaFicheroPdf, sPath & "\" & Conexiones._sEmpresa & "\ENVIADOS\HCOS", sFicheroPDF)
                        oLog.escribeMensaje("Archivo guardado en Hcos - " & sPath & "\" & Conexiones._sEmpresa & "\ENVIADOS\HCOS - .", EXO_Log.EXO_Log.Tipo.advertencia)
#End Region

                    End If

                Next
            Else
                oLog.escribeMensaje("No hay agentes para generar el report de Pedidos Procesados.", EXO_Log.EXO_Log.Tipo.advertencia)
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            If oDtProveedores IsNot Nothing Then oDtProveedores.Dispose()
            Conexiones.Disconnect_SQLServer(oDBSAP)
            Conexiones.Disconnect_Company(oCompany)
        End Try
    End Sub


#End Region

#Region "Comunes"
    Public Shared Sub FicheroaHistorico(ByVal folderPathOri As String, ByVal folderPathDes As String, ByVal file As String)
        Try
            Dim FileDestino As String = file
            My.Computer.FileSystem.MoveFile(folderPathOri, folderPathDes & "\" & FileDestino, True)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Shared Function GenerarCrystal(ByVal strRutaInforme As String, ByVal sFileCrystal As String, ByVal sPath As String, sDocEntry As String, sDocNum As String,
                                          sProveedor As String, ByVal sProveedorNombre As String, ByRef oLog As EXO_Log.EXO_Log) As String
        Dim oCRReport As ReportDocument = Nothing
        Dim sFilePDF As String = sProveedorNombre & "_" & sDocNum

        sPath = sPath & "\" & Conexiones._sEmpresa & "\ENVIADOS\"
        'Guardar ficheros enviados
        If Not System.IO.Directory.Exists(sPath) Then
            System.IO.Directory.CreateDirectory(sPath)
        End If
        Try
            'generar el rpt
            GenerarCrystal = ""
            oCRReport = New ReportDocument
            oLog.escribeMensaje("Cargando Crystal - " & strRutaInforme & sFileCrystal & " - .", EXO_Log.EXO_Log.Tipo.advertencia)
            oCRReport.Load(strRutaInforme & sFileCrystal)

            oCRReport.SetParameterValue("DocKey@", sDocEntry)
            oCRReport.SetParameterValue("Proveedor", sProveedor)

            'PONER USUARIO Y CONTRASEÑA

            Dim conrepor As CrystalDecisions.Shared.DataSourceConnections = oCRReport.DataSourceConnections
            conrepor(0).SetConnection(Conexiones._sServer, Conexiones._sEmpresa, Conexiones._sUserBD, Conexiones._sPassBD)

            sFilePDF = sPath & sFilePDF & ".pdf"

            oCRReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sFilePDF)

            oLog.escribeMensaje("Pdf creado : " & sFilePDF, EXO_Log.EXO_Log.Tipo.informacion)

            GenerarCrystal = sFilePDF
        Catch ex As Exception
            Throw ex
        Finally
            oCRReport.Close()
            oCRReport.Dispose()
            GC.Collect()
        End Try
    End Function
    Private Shared Function Enviarmail(sRutahtml As String, dirmail As String, sProveedor As String, sProveedorNom As String, sFichero As String, ByRef oLog As EXO_Log.EXO_Log) As Boolean

        Dim correo As New System.Net.Mail.MailMessage()
        Dim adjunto As System.Net.Mail.Attachment
        Dim StrFirma As String = ""
        Dim htmbody As New System.Text.StringBuilder()
        Enviarmail = False
        Dim sMail As String = Conexiones.Datos_Confi("DATOS_MAIL", "MAIL")
        Dim sCMail As String = Conexiones.Datos_Confi("DATOS_MAIL", "CMAIL")
        Dim sMail_Usuario As String = Conexiones.Datos_Confi("DATOS_MAIL", "USUARIO")
        Dim sMail_PS As String = Conexiones.Datos_Confi("DATOS_MAIL", "PS")
        Dim sMail_SMTP As String = Conexiones.Datos_Confi("DATOS_MAIL", "SMTP")
        Dim sMail_PORT As String = Conexiones.Datos_Confi("DATOS_MAIL", "PORT")
        Dim oCC As New Net.Mail.MailAddressCollection
        correo.From = New System.Net.Mail.MailAddress(sMail, "Brompton House")
        If sCMail.Trim <> "" Then
            correo.CC.Add(sCMail.Trim)
        End If

        If dirmail <> "" Then
            'dirmail = "omartinez@expertone.es"
        correo.To.Add(dirmail)
        End If
        correo.Subject = "Nuevo Pedido Procesado"

        If sFichero <> "" Then
            adjunto = New System.Net.Mail.Attachment(sFichero)
            correo.Attachments.Add(adjunto)
        End If

        Dim cuerpo As String = ""

        Dim FicheroCab As String = ""

        Select Case sProveedor
            Case "S0012" : FicheroCab = sRutahtml & "Mail_TIPO1.htm"
            Case "S0004" : FicheroCab = sRutahtml & "Mail_TIPO2.htm"
            Case "S0359" : FicheroCab = sRutahtml & "Mail_TIPO3.htm"
            Case "S0013" : FicheroCab = sRutahtml & "Mail_TIPO4.htm"
            Case Else : FicheroCab = sRutahtml & "mail.htm"
        End Select

        Dim srCAB As StreamReader = New StreamReader(FicheroCab)

        cuerpo = srCAB.ReadToEnd()

        correo.Body = cuerpo
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
            oLog.escribeMensaje("Correo enviado a " & sProveedorNom & " con mail: " & dirmail & ", adjuntando fichero:" & sFichero, EXO_Log.EXO_Log.Tipo.informacion)
            Enviarmail = True
        Catch ex As Exception
            oLog.escribeMensaje("No se ha podido envial mail a " & sProveedorNom & " con mail: " & dirmail & ", adjuntando fichero:" & sFichero & ". Error: " & ex.Message, EXO_Log.EXO_Log.Tipo.informacion)
            Enviarmail = False
        Finally



        End Try

        Return True
    End Function
#End Region
#Region "Actualizar campos"
    Public Shared Sub Actualizar_Campos(ByRef oLog As EXO_Log.EXO_Log)
        Dim oDBSAPSQL As SqlConnection = Nothing
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim sError As String = ""
        Dim sSQL As String = ""
        Dim OdtDatos As System.Data.DataTable = Nothing
        Dim sPass As String = "" : Dim sVSQL As String = ""
        Dim oXML As String = ""
        Dim sDir As String = Application.StartupPath

        Dim tipoServidor As SAPbobsCOM.BoDataServerTypes = Nothing
        Dim sBBDD As String = "" : Dim sUsuario As String = "" : Dim sPassword As String = ""
        Dim sServidor As String = "" : Dim sSRVLicencia As String = "" : Dim sUsBBDD As String = "" : Dim sPwdBBDD As String = ""

        Try
            OdtDatos = New System.Data.DataTable

            sBBDD = Conexiones.Datos_Confi("DI", "CompanyDB")
            sUsuario = Conexiones.Datos_Confi("DI", "UserName")
            sPassword = Conexiones.Datos_Confi("DI", "Password")
            sServidor = Conexiones.Datos_Confi("DI", "Server")
            sSRVLicencia = Conexiones.Datos_Confi("DI", "LicenseServer")
            sUsBBDD = Conexiones.Datos_Confi("DI", "DbUserName")
            sPwdBBDD = Conexiones.Datos_Confi("DI", "DbPassword")
            sVSQL = Conexiones.Datos_Confi("DI", "SQLV")
            Select Case sVSQL
                Case "2012" : tipoServidor = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                Case "2014" : tipoServidor = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                Case "2016" : tipoServidor = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
            End Select
            Conexiones.Connect_SQLServer(oDBSAPSQL, "SQL", oLog)

#Region "CREAR EN SAP EN BBDD"
            Dim sRuta As String = sDir & "\01.XML\XML_BD\UDFs_ORDR.xml"
            If sRuta <> "" Then
                Dim compañia As SAPbobsCOM.Company = New SAPbobsCOM.Company
                Dim errorSBO As Boolean = False
                Dim i As Integer = 4000
                Dim elementos As Integer
                Dim codError As Integer
                'Conexion empresa origen
                compañia.DbServerType = tipoServidor
                compañia.Server = sServidor
                compañia.LicenseServer = sSRVLicencia
                compañia.UseTrusted = False
                compañia.DbUserName = sUsBBDD
                compañia.DbPassword = sPwdBBDD
                compañia.CompanyDB = sBBDD
                compañia.UserName = sUsuario
                compañia.Password = sPassword
                oLog.escribeMensaje("Conectando a: " + compañia.CompanyDB, EXO_Log.EXO_Log.Tipo.advertencia)
                If compañia.Connect <> 0 Then
                    oLog.escribeMensaje("Error de conexion: " + compañia.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                    MsgBox("Error de conexion: " + compañia.GetLastErrorDescription, MsgBoxStyle.Exclamation)
                    errorSBO = True
                End If
                If Not errorSBO Then
                    compañia.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    compañia.XMLAsString = True
                    If System.IO.File.Exists(sRuta) Then
                        Dim docXML As Xml.XmlDocument = New Xml.XmlDocument()
                        docXML.Load(sRuta)

                        elementos = compañia.GetXMLelementCount(docXML.InnerXml)
                        For i = 0 To elementos - 1
                            Select Case compañia.GetXMLobjectType(docXML.InnerXml, i)
                                Case SAPbobsCOM.BoObjectTypes.oUserFields
                                    Dim campoUsuario As SAPbobsCOM.UserFieldsMD
                                    campoUsuario = compañia.GetBusinessObjectFromXML(docXML.InnerXml, i)
                                    oLog.escribeMensaje("Campo: " + campoUsuario.Name, EXO_Log.EXO_Log.Tipo.informacion)
                                    If Not exiteCampoUsuario(campoUsuario.TableName, campoUsuario.Name, compañia) Then
                                        codError = campoUsuario.Add()
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(campoUsuario)
                                    campoUsuario = Nothing
                                Case SAPbobsCOM.BoObjectTypes.oUserTables
                                    Dim tablaUsuario As SAPbobsCOM.UserTablesMD
                                    tablaUsuario = compañia.GetBusinessObjectFromXML(docXML.InnerXml, i)
                                    oLog.escribeMensaje("Tabla: " + tablaUsuario.TableName, EXO_Log.EXO_Log.Tipo.informacion)
                                    If Not tablaUsuario.GetByKey(tablaUsuario.TableName) Then
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(tablaUsuario)
                                        tablaUsuario = Nothing
                                        tablaUsuario = compañia.GetBusinessObjectFromXML(docXML.InnerXml, i)
                                        codError = tablaUsuario.Add()
                                    End If
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tablaUsuario)
                                    tablaUsuario = Nothing
                                'UDOS
                                Case SAPbobsCOM.BoObjectTypes.oUserObjectsMD
                                    Dim oUDO As SAPbobsCOM.UserObjectsMD = compañia.GetBusinessObjectFromXML(docXML.InnerXml, i)
                                    '               gProgressBar.Value = gProgressBar.Value + 1
                                    oLog.escribeMensaje("UDO: " + oUDO.Code, EXO_Log.EXO_Log.Tipo.informacion)
                                    Dim oUDO2 As SAPbobsCOM.UserObjectsMD = compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
                                    If oUDO2.GetByKey(oUDO.Code) Then
                                        Dim xmlUDO As String = oUDO.GetAsXML
                                        codError = actualizaUDO(oUDO2, oUDO, compañia)
                                        If codError <> 0 Then
                                            oLog.escribeMensaje("Error: " + compañia.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                            System.Threading.Thread.Sleep(3000)
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO2)
                                        oUDO2 = Nothing
                                        Continue For
                                    Else
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO2)
                                        oUDO2 = Nothing
                                        GC.Collect()
                                        Dim xmlUDO As String = oUDO.GetAsXML
                                        codError = oUDO.Add
                                        If codError <> 0 And codError <> -2035 Then
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO)
                                            oUDO = Nothing
                                            oLog.escribeMensaje("Error: " + compañia.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                            System.Threading.Thread.Sleep(3000)
                                            Exit For
                                        ElseIf codError = -2035 Then
                                        End If
                                        If Not oUDO Is Nothing Then
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO)
                                            oUDO = Nothing
                                        End If
                                    End If
                                Case SAPbobsCOM.BoObjectTypes.oUserKeys
                                    Dim oKeys As SAPbobsCOM.UserKeysMD = compañia.GetBusinessObjectFromXML(docXML.InnerXml, i)
                                    codError = oKeys.Add
                                    If codError <> 0 And codError <> -1 Then
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oKeys)
                                        oKeys = Nothing
                                        oLog.escribeMensaje("Error: " + compañia.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                                        System.Threading.Thread.Sleep(3000)
                                        Exit For
                                    End If
                                    If Not oKeys Is Nothing Then
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oKeys)
                                        oKeys = Nothing
                                    End If
                            End Select
                        Next i
                    Else
                        MsgBox("No existe el fichero indicado")
                    End If
                End If

                If compañia.Connected Then
                    compañia.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(compañia)
                compañia = Nothing
            Else
                MsgBox("Debe indicar un fichero")
            End If
#End Region

        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            sError = ex.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            Conexiones.Disconnect_SQLServer(oDBSAPSQL)
        End Try

    End Sub
#End Region
#Region "SAP"
    Public Shared Function exiteCampoUsuario(ByVal tabla As String, ByVal campo As String, ByRef interfazDatos As SAPbobsCOM.Company) As Boolean
        Dim rs As SAPbobsCOM.Recordset = interfazDatos.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If interfazDatos.DbServerType = BoDataServerTypes.dst_HANADB Then
            rs.DoQuery("SELECT COUNT('A') FROM ""CUFD"" WHERE ""TableID"" = '" + tabla + "' AND ""AliasID"" = '" + campo + "'")
        Else
            rs.DoQuery("SELECT COUNT('A') FROM CUFD WHERE TableID = '" + tabla + "' AND AliasID = '" + campo + "'")
        End If
        Dim num As Integer = Int32.Parse(rs.Fields.Item(0).Value.ToString())
        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
        Return num = 1
    End Function
    Public Shared Function actualizaUDO(ByRef udo As SAPbobsCOM.UserObjectsMD, udoB1Aux As SAPbobsCOM.UserObjectsMD, ByRef compañia As Company) As Integer
        Dim res As Integer = 0
        Dim xmlasstring As Boolean = compañia.XMLAsString
        compañia.XMLAsString = True
        If udo.Code = udoB1Aux.Code Then
            udo.Name = udoB1Aux.Name
            udo.CanArchive = udoB1Aux.CanArchive
            udo.CanCancel = udoB1Aux.CanCancel
            udo.CanClose = udoB1Aux.CanClose
            udo.CanCreateDefaultForm = udoB1Aux.CanCreateDefaultForm
            udo.CanDelete = udoB1Aux.CanDelete
            udo.CanFind = udoB1Aux.CanFind
            udo.CanLog = udoB1Aux.CanLog
            udo.CanYearTransfer = udoB1Aux.CanYearTransfer
            For indiceTablas As Integer = 0 To udoB1Aux.ChildTables.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.ChildTables.SetCurrentLine(indiceTablas)
                For indiceTablasOriginales As Integer = 0 To udo.ChildTables.Count - 1
                    udo.ChildTables.SetCurrentLine(indiceTablasOriginales)
                    If udo.ChildTables.TableName = udoB1Aux.ChildTables.TableName Then
                        encontrada = True
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.ChildTables.Add()
                    udo.ChildTables.SetCurrentLine(udo.ChildTables.Count - 1)
                    udo.ChildTables.TableName = udoB1Aux.ChildTables.TableName
                    udo.ChildTables.LogTableName = udoB1Aux.ChildTables.LogTableName
                End If
            Next
            udo.EnableEnhancedForm = udoB1Aux.EnableEnhancedForm
            For indiceForm As Integer = 0 To udoB1Aux.EnhancedFormColumns.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.EnhancedFormColumns.SetCurrentLine(indiceForm)
                For indiceFormOriginal As Integer = 0 To udo.EnhancedFormColumns.Count - 1
                    udo.EnhancedFormColumns.SetCurrentLine(indiceFormOriginal)
                    If udo.EnhancedFormColumns.ColumnAlias = udoB1Aux.EnhancedFormColumns.ColumnAlias Then
                        encontrada = True
                        udo.EnhancedFormColumns.ColumnDescription = udoB1Aux.EnhancedFormColumns.ColumnDescription
                        Try
                            udo.EnhancedFormColumns.ColumnIsUsed = udoB1Aux.EnhancedFormColumns.ColumnIsUsed
                        Catch
                        End Try
                        udo.EnhancedFormColumns.ColumnNumber = udoB1Aux.EnhancedFormColumns.ColumnNumber
                        Try
                            udo.EnhancedFormColumns.Editable = udoB1Aux.EnhancedFormColumns.Editable
                        Catch
                        End Try
                        udo.EnhancedFormColumns.ChildNumber = udoB1Aux.EnhancedFormColumns.ChildNumber
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.EnhancedFormColumns.Add()
                    udo.EnhancedFormColumns.SetCurrentLine(udo.EnhancedFormColumns.Count - 1)
                    udo.EnhancedFormColumns.ColumnAlias = udoB1Aux.EnhancedFormColumns.ColumnAlias
                    udo.EnhancedFormColumns.ColumnDescription = udoB1Aux.EnhancedFormColumns.ColumnDescription
                    udo.EnhancedFormColumns.ColumnIsUsed = udoB1Aux.EnhancedFormColumns.ColumnIsUsed
                    udo.EnhancedFormColumns.ColumnNumber = udoB1Aux.EnhancedFormColumns.ColumnNumber
                    udo.EnhancedFormColumns.Editable = udoB1Aux.EnhancedFormColumns.Editable
                    udo.EnhancedFormColumns.ChildNumber = udoB1Aux.EnhancedFormColumns.ChildNumber
                End If
            Next
            udo.ExtensionName = udoB1Aux.ExtensionName
            udo.FatherMenuID = udoB1Aux.FatherMenuID
            For indiceBucar As Integer = 0 To udoB1Aux.FindColumns.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.FindColumns.SetCurrentLine(indiceBucar)
                For indiceBuscarOriginal As Integer = 0 To udo.FindColumns.Count - 1
                    udo.FindColumns.SetCurrentLine(indiceBuscarOriginal)
                    If udo.FindColumns.ColumnAlias = udoB1Aux.FindColumns.ColumnAlias Then
                        encontrada = True
                        udo.FindColumns.ColumnDescription = udoB1Aux.FindColumns.ColumnDescription
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.FindColumns.Add()
                    udo.FindColumns.SetCurrentLine(udo.FindColumns.Count - 1)
                    udo.FindColumns.ColumnAlias = udoB1Aux.FindColumns.ColumnAlias
                    udo.FindColumns.ColumnDescription = udoB1Aux.FindColumns.ColumnDescription
                End If
            Next
            For indiceFormB As Integer = 0 To udoB1Aux.FormColumns.Count - 1
                Dim encontrada As Boolean = False
                udoB1Aux.FormColumns.SetCurrentLine(indiceFormB)
                For indiceFormBOriginal As Integer = 0 To udo.FormColumns.Count - 1
                    udo.FormColumns.SetCurrentLine(indiceFormBOriginal)
                    If udo.FormColumns.FormColumnAlias = udoB1Aux.FormColumns.FormColumnAlias Then
                        encontrada = True
                        udo.FormColumns.Editable = udoB1Aux.FormColumns.Editable
                        udo.FormColumns.FormColumnDescription = udoB1Aux.FormColumns.FormColumnDescription
                        udo.FormColumns.SonNumber = udoB1Aux.FormColumns.SonNumber
                        Exit For
                    End If
                Next
                If Not encontrada Then
                    udo.FormColumns.Add()
                    udo.FormColumns.SetCurrentLine(udo.FormColumns.Count - 1)
                    udo.FormColumns.FormColumnAlias = udoB1Aux.FormColumns.FormColumnAlias
                    udo.FormColumns.Editable = udoB1Aux.FormColumns.Editable
                    udo.FormColumns.FormColumnDescription = udoB1Aux.FormColumns.FormColumnDescription
                    udo.FormColumns.SonNumber = udoB1Aux.FormColumns.SonNumber
                End If
            Next
            udo.FormSRF = udoB1Aux.FormSRF
            udo.RebuildEnhancedForm = udoB1Aux.RebuildEnhancedForm
            udo.LogTableName = udoB1Aux.LogTableName
            udo.ManageSeries = udoB1Aux.ManageSeries
            udo.MenuCaption = udoB1Aux.MenuCaption
            udo.MenuItem = udoB1Aux.MenuItem
            udo.MenuUID = udoB1Aux.MenuUID
            udo.Name = udoB1Aux.Name
            udo.OverwriteDllfile = udoB1Aux.OverwriteDllfile
            udo.Position = udoB1Aux.Position
            udo.TableName = udoB1Aux.TableName
            udo.UseUniqueFormType = udoB1Aux.UseUniqueFormType
            System.Runtime.InteropServices.Marshal.ReleaseComObject(udoB1Aux)
            udoB1Aux = Nothing
            GC.Collect()
            res = udo.Update()
        End If
        compañia.XMLAsString = xmlasstring
        Return res
    End Function
#End Region
End Class

