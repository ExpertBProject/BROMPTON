Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_PEDCOM
    Inherits EXO_UIAPI.EXO_DLLBase
#Region "Variables públicas"
    Public Shared _sDocEntryVenta As String = ""
#End Region
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        If actualizar Then
            GenerarParametros()
        End If
    End Sub
    Private Sub GenerarParametros()
        If objGlobal.refDi.comunes.esAdministrador Then
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_ART") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_ART", "A0030")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_IMP") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_IMP", "EX-I")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_ALM") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_ALM", "01.FW")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_Mail") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_Mail", "export@bromptonhouse.co.uk")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_US") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_US", "export@bromptonhouse.co.uk")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_PS") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_PS", "vKmmvWBi@7F9")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_SMTP") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_SMTP", "mail2.cloudari.com")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_PORT") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_PORT", "587")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("COM_CMail") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("COM_CMail", "ignacio@bromptonhouse.co.uk")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("F_PED_VENTA") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("F_PED_VENTA", "RDR20011")
            End If
            If Not objGlobal.funcionesUI.refDi.OGEN.existeVariable("F_PED_COMPRA") Then
                objGlobal.funcionesUI.refDi.OGEN.fijarValorVariable("F_PED_COMPRA", "POR20006")
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
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_PEDCOM"
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
                        Case "EXO_PEDCOM"
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
                        Case "EXO_PEDCOM"
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
                        Case "EXO_PEDCOM"
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
            If pVal.ItemUID = "txtPROV" AndAlso pVal.ChooseFromListUID = "CFLIC" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("UDPROV").Value = oDataTable.GetValue("CardCode", 0).ToString
                        'CType(oForm.Items.Item("DIC").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                    Catch ex As Exception
                        CType(oForm.Items.Item("txtPROV").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
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
        EventHandler_ItemPressed_After = False
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim oDocVenta As SAPbobsCOM.Documents = Nothing
        Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sMensaje As String = "" : Dim sErrorDes As String = ""
        Dim sSQL As String = "" : Dim oDtProveedores As System.Data.DataTable = New System.Data.DataTable
        Dim sPath As String = "" : Dim sOutFileName As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sRutaFicheroPdfVenta As String = "" : Dim sRutaFicheroPdfCompra As String = ""
        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)


            Select Case pVal.ItemUID
                Case "btnGen"
                    If oForm.DataSources.UserDataSources.Item("UDFECH").Value <> "" And oForm.DataSources.UserDataSources.Item("UDGRUP").Value <> "" _
                            And oForm.DataSources.UserDataSources.Item("UDALM").Value <> "" And oForm.DataSources.UserDataSources.Item("UDPROV").Value <> "" _
                            And oForm.DataSources.UserDataSources.Item("UDPRICE").Value <> "" And oForm.DataSources.UserDataSources.Item("UDARTI").Value <> "" Then
                        If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar el pedido de compra?", 1, "Sí", "No") = 1 Then
#Region "Generación de pedido de compra"
                            oDoc = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders), SAPbobsCOM.Documents)
                            oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
                            oDoc.CardCode = oForm.DataSources.UserDataSources.Item("UDPROV").Value
                            oDoc.DocDate = oForm.DataSources.UserDataSources.Item("UDFECH").Value
                            oDoc.TaxDate = oForm.DataSources.UserDataSources.Item("UDFECH").Value
                            oDoc.Comments = oForm.DataSources.UserDataSources.Item("UDCOM").Value.ToString.Trim
                            oDoc.DocDueDate = oForm.DataSources.UserDataSources.Item("UDFECH").Value
                            oDoc.Lines.ItemCode = oForm.DataSources.UserDataSources.Item("UDARTI").Value
                            oDoc.Lines.Quantity = 1
                            oDoc.Lines.VatGroup = oForm.DataSources.UserDataSources.Item("UDGRUP").Value
                            oDoc.Lines.WarehouseCode = oForm.DataSources.UserDataSources.Item("UDALM").Value
                            Dim dPrecio As Double = EXO_GLOBALES.TextToDbl(objGlobal, oForm.DataSources.UserDataSources.Item("UDPRICE").Value)
                            oDoc.Lines.UnitPrice = dPrecio
                            If oDoc.Add() <> 0 Then 'Si ocurre un error en la grabación entra
                                sErrorDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                                objGlobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sErrorDes)
                            Else
                                sDocEntry = objGlobal.refDi.compañia.GetNewObjectKey() 'Recoge el último documento creado
                                'Buscamos el documento para crear un mensaje
                                sDocNum = objGlobal.refDi.SQL.sqlStringB1("SELECT ""DocNum"" FROM  ""OPOR""  WHERE ""DocEntry""=" & sDocEntry)
                                oForm.DataSources.UserDataSources.Item("UDPED").Value = sDocEntry
                                oForm.DataSources.UserDataSources.Item("UDNPED").Value = sDocNum


                                sMensaje = "(EXO) - Ha sido creado el pedido de compras Nº" & sDocNum
                                objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                objGlobal.SBOApp.MessageBox(sMensaje)
#Region "Asignación del documento como referencia"
                                '
                                'oDocVenta = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                'oDocVenta.GetByKey(_sDocEntryVenta)
                                'Dim iDoc As Integer = objGlobal.refDi.SQL.sqlNumericaB1("SELECT COUNT(""RefDocNum"") FROM ""RDR21"" Where ""DocEntry""=" & _sDocEntryVenta)
                                'If iDoc > 0 Then
                                '    oDoc.DocumentReferences.Add()
                                'End If
                                'objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se va a Referencia el documento " & sDocNum & " a Docentry " & _sDocEntryVenta, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                'oDocVenta.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_PurchaseOrder
                                'oDocVenta.DocumentReferences.ReferencedDocEntry = sDocEntry
                                'oDocVenta.OpeningRemarks = oForm.DataSources.UserDataSources.Item("UDCOM").Value.ToString.Trim
                                'If oDocVenta.Update() <> 0 Then
                                '    sErrorDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                                '    objGlobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '    objGlobal.SBOApp.MessageBox(sErrorDes)
                                'Else
                                '    sMensaje = "(EXO) - Se ha actualizado el pedido de venta con el documento de referencia"
                                '    objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                '    objGlobal.SBOApp.MessageBox(sMensaje)
                                'End If

#End Region
                            End If
#End Region
                        End If
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellene todos los datos por favor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btnENV"
#Region "Enviar Docs. a proveedor transportista"
                    If oForm.DataSources.UserDataSources.Item("UDPED").Value <> "" Then
                        sSQL = "SELECT T1.""DocEntry"", T4.""CardCode"", T4.""CardName"", T5.""E_MailL"" ""Email"" "
                        sSQL &= " From ""OPOR"" T1 "
                        sSQL &= " INNER JOIN ""OCRD"" T4 ON T1.""CardCode""= T4.""CardCode"" "
                        sSQL &= " INNER JOIN  ""OCPR"" T5 On  T4.""CardCode""= T5.""CardCode"" And T4.""CntctPrsn""= T5.""Name""  "
                        sSQL &= " WHERE T1.""DocEntry""='" & oForm.DataSources.UserDataSources.Item("UDPED").Value & "' "
                        sSQL &= " Order by  T1.""DocEntry"", T4.""CardCode"" "

                        oDtProveedores = objGlobal.refDi.SQL.sqlComoDataTable(sSQL)
                        If oDtProveedores.Rows.Count > 0 Then
                            sPath = objGlobal.refDi.OGEN.pathGeneral & "\08.Historico\"
#Region "Exportar el pedido de venta"
                            sOutFileName = sPath & "Pedido_Venta.rpt"
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Exportando RPT de PEDIDO VENTA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            EXO_GLOBALES.GetCrystalReportFile(objGlobal.compañia, objGlobal.funcionesUI.refDi.OGEN.valorVariable("F_PED_VENTA"), sOutFileName)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Utilizando formato: " & sOutFileName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            'generar PDF 
                            sSQL = "SELECT ""DocEntry"",""DocNum"", ""DocDate"" FROM ""ORDR"" WHERE ""DocEntry""=" & _sDocEntryVenta
                            oRs.DoQuery(sSQL)
                            For i = 0 To oRs.RecordCount - 1
                                sRutaFicheroPdfVenta = sPath
                                If Not System.IO.Directory.Exists(sPath) Then
                                    System.IO.Directory.CreateDirectory(sPath)
                                End If
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Pedido de Venta: " & oRs.Fields.Item("DocNum").Value.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                sRutaFicheroPdfVenta = EXO_GLOBALES.GenerarCrystal(objGlobal, sPath, "Pedido_Venta.rpt", sRutaFicheroPdfVenta, oRs.Fields.Item("DocEntry").Value.ToString, oRs.Fields.Item("DocNum").Value.ToString, oRs.Fields.Item("DocDate").Value.ToString)
                                oRs.MoveNext()
                            Next
#End Region
#Region "Exportar Formato pedido de compra"
                            sOutFileName = sPath & "Pedido_Compra.rpt"
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Exportando RPT de PEDIDO COMPRA", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            EXO_GLOBALES.GetCrystalReportFile(objGlobal.compañia, objGlobal.funcionesUI.refDi.OGEN.valorVariable("F_PED_COMPRA"), sOutFileName)
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Utilizando formato: " & sOutFileName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                            For i As Integer = 0 To oDtProveedores.Rows.Count - 1
#Region "Generamos PDF Compra"
                                'generar PDF 
                                sRutaFicheroPdfCompra = sPath
                                If Not System.IO.Directory.Exists(sPath) Then
                                    System.IO.Directory.CreateDirectory(sPath)
                                End If
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Pedido de compra: " & oForm.DataSources.UserDataSources.Item("UDNPED").Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                sRutaFicheroPdfCompra = EXO_GLOBALES.GenerarCrystal(objGlobal, sPath, "Pedido_Compra.rpt", sRutaFicheroPdfCompra, oForm.DataSources.UserDataSources.Item("UDPED").Value, oForm.DataSources.UserDataSources.Item("UDNPED").Value, oForm.DataSources.UserDataSources.Item("UDFECH").Value)

#End Region
#Region "Enviamos Mail"
                                'Como se ha generado el crystal, procedemos a su envío 
                                Dim cuerpo As String = objGlobal.leerEmbebido(Me.GetType(), "mail.htm")
                                EXO_GLOBALES.Enviarmail(objGlobal, cuerpo, sPath & "\02.HTML\", oDtProveedores.Rows.Item(i).Item("Email").ToString,
                                   oDtProveedores.Rows.Item(i).Item("CardCode").ToString, oDtProveedores.Rows.Item(i).Item("CardName").ToString, sRutaFicheroPdfVenta, sRutaFicheroPdfCompra)
#End Region
#Region "Borramos los ficheros de documentos"
                                If Not System.IO.File.Exists(sRutaFicheroPdfVenta) Then
                                    Try
                                        System.IO.File.Delete(sRutaFicheroPdfVenta)
                                    Catch ex As Exception

                                    End Try
                                End If
                                If Not System.IO.File.Exists(sRutaFicheroPdfCompra) Then
                                    Try
                                        System.IO.File.Delete(sRutaFicheroPdfCompra)
                                    Catch ex As Exception

                                    End Try
                                End If
#End Region
#Region "Actualizar a enviado"
                                'Actualizamos el documento a enviado
                                sSQL = "UPDATE ""ORDR"" "
                                sSQL &= " SET ""U_EXO_EMAIL""='Y' "
                                sSQL &= " WHERE ""DocEntry""='" & oDtProveedores.Rows.Item(i).Item("DocEntry").ToString & "' "
                                objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                objGlobal.SBOApp.StatusBar.SetText("Se actualiza el Pedido de compra Nº" & oDtProveedores.Rows.Item(i).Item("DocEntry").ToString & " indicando Enviado.", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region

                            Next
#Region "Borramos los ficheros de formatos"
                            If Not System.IO.File.Exists(sPath & "Pedido_Venta.rpt") Then
                                System.IO.File.Delete(sPath & "Pedido_Venta.rpt")
                            End If

                            If Not System.IO.File.Exists(sPath & "Pedido_Compra.rpt") Then
                                System.IO.File.Delete(sPath & "Pedido_Compra.rpt")
                            End If
#End Region
                        Else
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se encuentra Contacto para enviar Mail. Lo siento, deberá hacerlo manualmente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            objGlobal.SBOApp.MessageBox("No se encuentra Contacto para enviar Mail. Lo siento, deberá hacerlo manualmente.")
                        End If
                    Else
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se ha generado el pedido de compra. No puede enviar el documento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        objGlobal.SBOApp.MessageBox("No se ha generado el pedido de compra. No puede enviar el documento.")
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
End Class
