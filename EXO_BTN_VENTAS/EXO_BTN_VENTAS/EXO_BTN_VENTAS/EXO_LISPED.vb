Imports System.Xml
Imports SAPbouiCOM


Public Class EXO_LISPED
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        cargamenu()
        If actualizar Then

        End If
    End Sub
    Private Sub cargamenu()
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
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
                    Case "EXO-MnPedV"
                        'Cargamos pantalla.
                        If CargarFormLisPed() = False Then
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
    Public Function CargarFormLisPed() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing

        CargarFormLisPed = False

        Try

            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_LISPED.srf")

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
            CargarFormLisPed = True
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
                        Case "EXO_LISPED"
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
                                    If EventHandler_FORM_RESIZE_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_LISPED"
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
                        Case "EXO_LISPED"
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
                        Case "EXO_LISPED"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
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
            If pVal.ItemUID = "txtCodCli" AndAlso pVal.ChooseFromListUID = "CFLCLI" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("dsCli").Value = oDataTable.GetValue("CardCode", 0).ToString
                        oForm.DataSources.UserDataSources.Item("dsNom").Value = oDataTable.GetValue("CardName", 0).ToString

                    Catch ex As Exception
                        CType(oForm.Items.Item("txtCodCli").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("CardCode", 0).ToString
                    End Try
                End If

            ElseIf pVal.ItemUID = "txtNumD" AndAlso pVal.ChooseFromListUID = "CFLNUMD" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("dsNumD").Value = oDataTable.GetValue("DocNum", 0).ToString

                    Catch ex As Exception
                        CType(oForm.Items.Item("txtNumD").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("DocNum", 0).ToString
                    End Try
                End If
            ElseIf pVal.ItemUID = "txtNumH" AndAlso pVal.ChooseFromListUID = "CFLNUMH" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.UserDataSources.Item("dsNumH").Value = oDataTable.GetValue("DocNum", 0).ToString

                    Catch ex As Exception
                        CType(oForm.Items.Item("txtNumH").Specific, SAPbouiCOM.EditText).Value = oDataTable.GetValue("DocNum", 0).ToString
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

    Private Function EventHandler_Choose_FromList_Before(ByRef pVal As ItemEvent) As Boolean
        Dim oCFLEvento As ItemEvent = Nothing
        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_Choose_FromList_Before = False

        Try
            If pVal.ItemUID = "txtCodCli" Then 'Cliente
                oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

                oCFLEvento = CType(pVal, ItemEvent)

                oConds = New SAPbouiCOM.Conditions

                oCond = oConds.Add
                oCond.Alias = "CardType"
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCond.CondVal = "C"

                oForm.ChooseFromLists.Item(oCFLEvento.ChooseFromListUID).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCFLEvento, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sPath As String = "" : Dim sOutFileName As String = ""
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sRutaFicheroPdf As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            Select Case pVal.ItemUID
                Case "btCon"
                    sSQL = "SELECT distinct t0.DocEntry ""IrP"", t0.DocNum ""Número Documento"", t0.DocDueDate ""Fecha de contabilización"", t0.CardCode ""Número de Cliente"" ,
                            t0.U_ECI_RINTC ""Referencia interna de carga"", t0.NumAtCard ""Número referencia Interlocutor"",T0.DocTotal ""Total Documento"",
                            CASE WHEN COALESCE(t2.DocEntry,0) <> 0 THEN  COALESCE(t8.DocEntry,0) ELSE COALESCE(t5.DocEntry,0) END ""IrF"",
                            CASE WHEN COALESCE(t2.DocEntry,0) <> 0 THEN  COALESCE(t8.DocNum,0) ELSE COALESCE(t5.DocNum,0) END ""Factura de cliente asociada"",
                            T0.U_ECI_DEST ""Destino"",
                            CASE WHEN t0.DocStatus ='O' then 'Abiertos'  
                            WHEN t0.DocStatus ='C' AND T0.CANCELED='Y' THEN 'Cancelados' 
                            WHEN t0.DocStatus ='C' AND T0.CANCELED='N' THEN 'Cerrados' 
                            end ""Status Documento""
                            FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry
                            left OUTER JOIN INV1 T2 ON T2.BaseType = t1.ObjType and t2.BaseEntry = t1.DocEntry and t2.BaseLine = t1.LineNum
                            LEFT OUTER JOIN OINV T8 ON T2.DocEntry= T8.DocEntry
                            left OUTER JOIN DLN1 T3 ON T3.BaseType = t1.ObjType and T3.BaseEntry = t1.DocEntry and T3.BaseLine = t1.LineNum
						    LEFT OUTER JOIN ODLN T6 ON T3.DocEntry = T6.DocEntry
                            left OUTER JOIN INV1 T4 ON T4.BaseType = t3.ObjType and T4.BaseEntry = t3.DocEntry and T4.BaseLine = t3.LineNum
                            LEFT OUTER JOIN OINV T5 ON T4.DocEntry= T5.DocEntry
						    INNER JOIN OCRD T7 ON T0.CardCode = T7.CardCode "
                    If oForm.DataSources.UserDataSources.Item("dsNumD").Value <> "" Then
                        sSQL = sSQL & " WHERE t0.DocNum>= '" & oForm.DataSources.UserDataSources.Item("dsNumD").Value & "' "
                    End If
                    If oForm.DataSources.UserDataSources.Item("dsNumH").Value <> "" Then
                        sSQL = sSQL & " AND  t0.DocNum<= '" & oForm.DataSources.UserDataSources.Item("dsNumH").Value & "' "
                    End If

                    If oForm.DataSources.UserDataSources.Item("dsCli").Value <> "" Then
                        sSQL = sSQL & " AND  t0.CardCode= '" & oForm.DataSources.UserDataSources.Item("dsCli").Value & "' "
                    End If


                    sSQL = sSQL & " order by t0.DocNum DESC"
                    oForm.DataSources.DataTables.Item("dtCon").ExecuteQuery(sSQL)

                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(0).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(1).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(2).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(3).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(4).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(5).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(6).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(7).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(8).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(9).TitleObject.Sortable = True
                    CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(10).TitleObject.Sortable = True

                    'columnas con choose
                    oColumnTxt = CType(CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.EditTextColumn)
                    oColumnTxt.LinkedObjectType = "17"
                    oColumnTxt.Width = 15

                    oColumnTxt = CType(CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(3), SAPbouiCOM.EditTextColumn)
                    oColumnTxt.LinkedObjectType = "2"

                    oColumnTxt = CType(CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(7), SAPbouiCOM.EditTextColumn)
                    oColumnTxt.LinkedObjectType = "13"
                    oColumnTxt.Width = 15


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

    Private Function EventHandler_FORM_RESIZE_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing

        EventHandler_FORM_RESIZE_After = False

        Try
            If oForm.Visible = True Then

                oColumnTxt = CType(CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Width = 15


                oColumnTxt = CType(CType(oForm.Items.Item("gridCon").Specific, SAPbouiCOM.Grid).Columns.Item(7), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Width = 15

            End If

            EventHandler_FORM_RESIZE_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)

        Catch ex As Exception
            oForm.Freeze(False)

        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

End Class
