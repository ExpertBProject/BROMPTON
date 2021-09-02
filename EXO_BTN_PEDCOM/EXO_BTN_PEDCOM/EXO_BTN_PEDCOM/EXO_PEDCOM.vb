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
        Dim sDocEntry As String = "" : Dim sDocNum As String = ""
        Dim sMensaje As String = "" : Dim sErrorDes As String = ""
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
                                sDocEntry = objGlobal.compañia.GetNewObjectKey() 'Recoge el último documento creado
                                'Buscamos el documento para crear un mensaje
                                sDocNum = objGlobal.refDi.SQL.sqlStringB1("SELECT ""DocNum"" FROM  ""OPOR""  WHERE ""DocEntry""=" & sDocEntry)
                                oForm.DataSources.UserDataSources.Item("UDPED").Value = sDocEntry
                                oForm.DataSources.UserDataSources.Item("UDNPED").Value = sDocNum
#Region "Asignación del documento como referencia"
                                oDoc = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
                                oDoc.GetByKey(_sDocEntryVenta)
                                Dim iDoc As Integer = objGlobal.refDi.SQL.sqlNumericaB1("SELECT COUNT(""RefDocNum"") FROM ""RDR21"" Where ""DocEntry""=" & _sDocEntryVenta)
                                If iDoc > 0 Then
                                    oDoc.DocumentReferences.Add()
                                End If
                                oDoc.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_PurchaseOrder
                                oDoc.DocumentReferences.ReferencedDocEntry = sDocEntry
                                oDoc.OpeningRemarks = oForm.DataSources.UserDataSources.Item("UDCOM").Value.ToString.Trim
                                If oDoc.Update() <> 0 Then
                                    sErrorDes = objGlobal.compañia.GetLastErrorCode & " / " & objGlobal.compañia.GetLastErrorDescription
                                    objGlobal.SBOApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objGlobal.SBOApp.MessageBox(sErrorDes)
                                Else
                                    sMensaje = "(EXO) - Se ha actualizado el pedido de venta con el documento de referencia"
                                    objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                End If

#End Region
                                sMensaje = "(EXO) - Ha sido creado el pedido de compras Nº" & sDocNum
                                objGlobal.SBOApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                            End If
#End Region
                        End If
                    Else
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellene todos los datos por favor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If

            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
End Class
