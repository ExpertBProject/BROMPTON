Imports System.Xml
Imports SAPbouiCOM
Imports CrystalDecisions.CrystalReports.Engine

Public Class EXO_PDF
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        cargamenu()
        If actualizar Then
            GenerarParametros()
        End If
    End Sub
    Private Sub cargamenu()
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
    End Sub
    Private Sub GenerarParametros()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim sPath As String = ""
            Dim sSQL As String = ""
            sSQL = "SELECT ""U_EXO_PATH"" FROM ""@EXO_OGEN"" "
            sPath = objGlobal.refDi.SQL.sqlStringB1(sSQL)
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("PATH_PDF") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("PATH_PDF", sPath & "\08.Historico\")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("F_PEDIDO") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("F_PEDIDO", "RDR10007")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("F_FACTURA") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("F_FACTURA", "INV20018")
            End If
        End If
    End Sub

    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnPDFV"
                        'Cargamos pantalla.
                        If CargarFormPDFV() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormPDFV() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormPDFV = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_PDFV.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            'CargaComboFormato(oForm)
            'CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Item.Enabled = False
            'If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 1 Then
            '    CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("FACVENTAS", BoSearchKey.psk_ByValue)
            'Else
            '    CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("--", BoSearchKey.psk_ByValue)
            'End If
            CargarFormPDFV = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PDFV"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PDFV"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PDFV"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PDFV"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False

        Try

            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            Dim oDataTable As DataTable
            If pVal.ItemUID = "txtPD" AndAlso pVal.ChooseFromListUID = "CFLDP" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("UDDP").Value = oDataTable.GetValue("DocEntry", 0).ToString
                        'CType(oForm.Items.Item("DIC").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                    Catch ex As Exception
                        CType(oForm.Items.Item("txtPD").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("DocEntry", 0).ToString
                    End Try
                End If
            ElseIf pVal.ItemUID = "txtPH" AndAlso pVal.ChooseFromListUID = "CFLHP" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("UDHP").Value = oDataTable.GetValue("DocEntry", 0).ToString
                    Catch ex As Exception
                        CType(oForm.Items.Item("txtPH").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("DocEntry", 0).ToString
                    End Try
                End If
            ElseIf pVal.ItemUID = "txtFD" AndAlso pVal.ChooseFromListUID = "CFLDF" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("UDDF").Value = oDataTable.GetValue("DocEntry", 0).ToString
                        'CType(oForm.Items.Item("DIC").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                    Catch ex As Exception
                        CType(oForm.Items.Item("txtFD").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("DocEntry", 0).ToString
                    End Try
                End If
            ElseIf pVal.ItemUID = "txtFH" AndAlso pVal.ChooseFromListUID = "CFLHF" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("UDHF").Value = oDataTable.GetValue("DocEntry", 0).ToString
                    Catch ex As Exception
                        CType(oForm.Items.Item("txtFH").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("DocEntry", 0).ToString
                    End Try
                End If
            End If




            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sPath As String = "" : Dim sOutFileName As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sRutaFicheroPdf As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            sPath = objGlobal.refDi.OGEN.valorVariable("PATH_PDF")

            Select Case pVal.ItemUID
                Case "btnGen"
#Region "Impresión de pedidos"
                    'Impresión de pedidos
                    If oForm.DataSources.UserDataSources.Item("UDDP").Value <> "" And oForm.DataSources.UserDataSources.Item("UDHP").Value <> "" Then
                        If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar los PDF de los pedidos seleccionados?", 1, "Sí", "No") = 1 Then
                            'Comprobamos que exista el directorio y sino, lo creamos

                            sOutFileName = sPath & "Pedido.rpt"
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Exportando RPT de PEDIDO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            GetCrystalReportFile(objGlobal.compañia, objGlobal.funcionesUI.refDi.OGEN.valorVariable("F_PEDIDO"), sOutFileName)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Utilizando formato: " & sOutFileName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'generar PDF 
                            sSQL = "SELECT ""DocEntry"",""DocNum"", ""DocDate"" FROM ""ORDR"" WHERE ""DocEntry"">=" & oForm.DataSources.UserDataSources.Item("UDDP").Value & " and ""DocEntry""<=" & oForm.DataSources.UserDataSources.Item("UDHP").Value
                            oRs.DoQuery(sSQL)
                            For i = 0 To oRs.RecordCount - 1
                                sRutaFicheroPdf = sPath & "VENTAS\PEDIDOS\"
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Pedido: " & oRs.Fields.Item("DocNum").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                sRutaFicheroPdf = GenerarCrystal(sPath, "Pedido.rpt", sRutaFicheroPdf, oRs.Fields.Item("DocEntry").Value.ToString, oRs.Fields.Item("DocNum").Value.ToString, oRs.Fields.Item("DocDate").Value.ToString)
                                oRs.MoveNext()
                            Next


                        End If
                    End If
#End Region
#Region "Impresión de Facturas"
                    'Impresión de facturas
                    If oForm.DataSources.UserDataSources.Item("UDDF").Value <> "" And oForm.DataSources.UserDataSources.Item("UDHF").Value <> "" Then
                        If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar los PDF de las facturas seleccionadas?", 1, "Sí", "No") = 1 Then
                            'Comprobamos que exista el directorio y sino, lo creamos


                            sOutFileName = sPath & "Factura.rpt"
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Exportando RPT de PEDIDO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            GetCrystalReportFile(objGlobal.compañia, objGlobal.funcionesUI.refDi.OGEN.valorVariable("F_FACTURA"), sOutFileName)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Utilizando formato: " & sOutFileName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'generar PDF 
                            sSQL = "SELECT ""DocEntry"",""DocNum"", ""DocDate"" FROM ""OINV"" WHERE ""DocEntry"">=" & oForm.DataSources.UserDataSources.Item("UDDF").Value & " and ""DocEntry""<=" & oForm.DataSources.UserDataSources.Item("UDHF").Value
                            oRs.DoQuery(sSQL)
                            For i = 0 To oRs.RecordCount - 1
                                sRutaFicheroPdf = sPath & "VENTAS\FACTURAS\"
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Factura: " & oRs.Fields.Item("DocNum").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                sRutaFicheroPdf = GenerarCrystal(sPath, "Factura.rpt", sRutaFicheroPdf, oRs.Fields.Item("DocEntry").Value.ToString, oRs.Fields.Item("DocNum").Value.ToString, oRs.Fields.Item("DocDate").Value.ToString)
                                oRs.MoveNext()
                            Next


                        End If
                    End If
#End Region
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Sub GetCrystalReportFile(ByVal oCompany As SAPbobsCOM.Company, ByVal sFormatoImp As String, ByVal sOutFileName As String)
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
    Public Function GenerarCrystal(ByVal strRutaInforme As String, ByVal sFileCrystal As String, ByVal sPath As String, sDocEntry As String, ByVal sDocNum As String, ByVal sDocDate As String) As String
        Dim oCRReport As ReportDocument = Nothing
        'objGlobal.SBOApp.StatusBar.SetText("(EXO) - Fecha:" & sDocDate & " - .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Dim dFecha As Date = sDocDate
        Dim sFilePDF As String = sDocNum & "_" & dFecha.Year.ToString("0000") & dFecha.Month.ToString("00") & dFecha.Day.ToString("00")


        'Guardar ficheros 
        If Not System.IO.Directory.Exists(sPath) Then
            System.IO.Directory.CreateDirectory(sPath)
        End If
        Try
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Preparando fichero - " & sFilePDF & " - .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'generar el rpt
            GenerarCrystal = ""
            'objGlobal.SBOApp.StatusBar.SetText("(EXO) - 1", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oCRReport = New ReportDocument
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Cargando Crystal - " & strRutaInforme & sFileCrystal & " - .", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            oCRReport.Load(strRutaInforme & sFileCrystal)

            oCRReport.SetParameterValue("DocKey@", sDocEntry)

            'PONER USUARIO Y CONTRASEÑA

            Dim conrepor As CrystalDecisions.Shared.DataSourceConnections = oCRReport.DataSourceConnections
            conrepor(0).SetConnection(objGlobal.compañia.Server, objGlobal.compañia.CompanyDB, objGlobal.refDi.SQL.usuarioSQL, objGlobal.refDi.SQL.claveSQL)
            'objGlobal.SBOApp.StatusBar.SetText("(EXO) - 2", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sFilePDF = sPath & sFilePDF & ".pdf"
            oCRReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, sFilePDF)
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Pdf creado : " & sFilePDF, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            GenerarCrystal = sFilePDF
        Catch ex As Exception
            Throw ex
        Finally
            oCRReport.Close()
            oCRReport.Dispose()
        End Try
    End Function

End Class
