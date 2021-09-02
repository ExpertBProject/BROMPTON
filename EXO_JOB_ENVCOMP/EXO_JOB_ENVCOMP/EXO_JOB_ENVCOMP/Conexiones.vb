Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient

Public Class Conexiones
#Region "Variables globales"
    Public Shared _sUserBD As String = ""
    Public Shared _sPassBD As String = ""
    Public Shared _sServer As String = ""
    Public Shared _sEmpresa As String = ""
#End Region
#Region "Connect to Company"
    Public Shared Sub Connect_Company(ByRef oCompany As SAPbobsCOM.Company, ByRef oLog As EXO_Log.EXO_Log, ByRef ConexBD As String)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing
        Dim sTipoSQL As String = ""
        Try
            'Conectar DI SAP
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            oLog.escribeMensaje("Leyendo cadena de conexión...", EXO_Log.EXO_Log.Tipo.advertencia)
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case ConexBD
                                oCompany = New SAPbobsCOM.Company
                                oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
                                oCompany.Server = Reader.GetAttribute("Server").ToString.Trim
                                oCompany.LicenseServer = Reader.GetAttribute("LicenseServer").ToString.Trim
                                oCompany.UserName = Reader.GetAttribute("UserName").ToString.Trim
                                oCompany.Password = Reader.GetAttribute("Password").ToString.Trim
                                oCompany.UseTrusted = False
                                oCompany.DbPassword = Reader.GetAttribute("DbPassword").ToString.Trim
                                oCompany.DbUserName = Reader.GetAttribute("DbUserName").ToString.Trim
                                sTipoSQL = Reader.GetAttribute("SQLV").ToString.Trim
                                Select Case sTipoSQL
                                    Case "2012" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                                    Case "2014" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                                    Case "2016" : oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
                                End Select
                                oLog.escribeMensaje("Versión SQL:" & sTipoSQL, EXO_Log.EXO_Log.Tipo.advertencia)
                                oCompany.SLDServer = Reader.GetAttribute("SLDAddress").ToString.Trim
                                oCompany.CompanyDB = Reader.GetAttribute("CompanyDB").ToString.Trim
                                oLog.escribeMensaje("database:" & oCompany.CompanyDB, EXO_Log.EXO_Log.Tipo.advertencia)
                                If oCompany.Connect <> 0 Then
                                    Throw New System.Exception("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim)
                                    oLog.escribeMensaje("Error en la conexión a la compañia:" & oCompany.GetLastErrorDescription.Trim, EXO_Log.EXO_Log.Tipo.error)
                                End If
                                Exit While
                        End Select
                End Select
            End While
            oLog.escribeMensaje("Conectado a la compañia", EXO_Log.EXO_Log.Tipo.advertencia)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
            oLog.escribeMensaje(exCOM.Message, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            Throw ex
            oLog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Public Shared Sub Disconnect_Company(ByRef oCompany As SAPbobsCOM.Company)
        Try
            If Not oCompany Is Nothing Then
                If oCompany.Connected = True Then
                    oCompany.Disconnect()
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If oCompany IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCompany)
            oCompany = Nothing
        End Try
    End Sub
#End Region

#Region "Connect to SQL Server"
    Public Shared Sub Connect_SQLServer(ByRef db As SqlConnection, ByVal sTipoSQL As String, ByRef oLog As EXO_Log.EXO_Log)
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing
        Dim sError As String = ""
        Try
            'Conectar SQL
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sTipoSQL
                                If db Is Nothing OrElse db.State = ConnectionState.Closed Then
                                    _sServer = Reader.GetAttribute("Server").ToString.Trim
                                    _sUserBD = Reader.GetAttribute("DbUser").ToString
                                    _sPassBD = Reader.GetAttribute("DbPwd").ToString
                                    _sEmpresa = Reader.GetAttribute("Empresa").ToString
                                    db = New SqlConnection
                                    db.ConnectionString = "Database=" & Reader.GetAttribute("Db").ToString.Trim.Trim & ";Data Source=" & Reader.GetAttribute("Server").ToString.Trim & ";User Id=" & Reader.GetAttribute("DbUser").ToString & ";Password=" & Reader.GetAttribute("DbPwd").ToString
                                    oLog.escribeMensaje("Cadena de conexión:" & db.ConnectionString, EXO_Log.EXO_Log.Tipo.advertencia)
                                    db.Open()
                                    oLog.escribeMensaje("Ha conectado con la BBDD", EXO_Log.EXO_Log.Tipo.advertencia)
                                End If
                                Exit While
                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            sError = exCOM.Message
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
            Throw exCOM

        Catch ex As Exception
            oLog.escribeMensaje(ex.Message, EXO_Log.EXO_Log.Tipo.error)
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Sub
    Public Shared Sub Disconnect_SQLServer(ByRef db As SqlConnection)
        Try
            If Not db Is Nothing AndAlso db.State = ConnectionState.Open Then
                db.Close()
                db.Dispose()
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub
    Public Shared Sub FillDtDB(ByRef db As SqlConnection, ByRef dt As System.Data.DataTable, ByVal strConsulta As String)
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            cmd = New SqlCommand(strConsulta, db)

            cmd.CommandTimeout = 0
            da = New SqlDataAdapter
            da.SelectCommand = cmd
            da.Fill(dt)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Sub
    Public Shared Sub ExecuteSQLDB(ByRef db As SqlConnection, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing

        Try
            cmd = New SqlCommand(sSQL, db)
            cmd.ExecuteNonQuery()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub
    Public Shared Sub ExecuteSQLDB(ByRef db As SqlConnection, ByRef oTransaction As SqlTransaction, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing

        Try
            cmd = New SqlCommand(sSQL, db)
            cmd.Transaction = oTransaction
            cmd.ExecuteNonQuery()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
        End Try
    End Sub
    Public Shared Function GetValueDB(ByRef db As SqlConnection, ByRef sTabla As String, ByRef sCampo As String, ByRef sCondicion As String) As String
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            If sCondicion = "" Then
                sSQL = "SELECT " & sCampo & " FROM " & sTabla
            Else
                sSQL = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sCondicion
            End If

            dt = New System.Data.DataTable("Tabla")

            cmd = New SqlCommand(sSQL, db)
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows.Item(0).Item(0).ToString) Then
                    Return dt.Rows.Item(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function
    Public Shared Function GetValueDB(ByRef db As SqlConnection, ByRef oTransaction As SqlTransaction, ByRef sTabla As String, ByRef sCampo As String, ByRef sCondicion As String) As String
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            If sCondicion = "" Then
                sSQL = "SELECT " & sCampo & " FROM " & sTabla
            Else
                sSQL = "SELECT " & sCampo & " FROM " & sTabla & " WHERE " & sCondicion
            End If

            dt = New System.Data.DataTable("Tabla")

            cmd = New SqlCommand(sSQL, db)
            cmd.Transaction = oTransaction
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows.Item(0).Item(0).ToString) Then
                    Return dt.Rows.Item(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
            End If
            If Not cmd Is Nothing Then
                cmd.Dispose()
            End If
            If Not da Is Nothing Then
                da.Dispose()
            End If
        End Try
    End Function
#End Region
#Region "Otros"
    Public Shared Function Datos_Confi(ByVal sTipo As String, ByVal sDato As String) As String
        Dim myStream As Stream = Nothing
        Dim Reader As XmlTextReader = Nothing

        Try
            myStream = File.OpenRead(My.Application.Info.DirectoryPath.ToString & "\Connections.xml")
            Reader = New XmlTextReader(myStream)
            myStream = Nothing
            While Reader.Read
                Select Case Reader.NodeType
                    Case XmlNodeType.Element
                        Select Case Reader.Name.ToString.Trim
                            Case sTipo
                                Datos_Confi = Reader.GetAttribute(sDato).ToString.Trim
                        End Select
                End Select
            End While

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
        End Try
    End Function
#End Region
End Class
