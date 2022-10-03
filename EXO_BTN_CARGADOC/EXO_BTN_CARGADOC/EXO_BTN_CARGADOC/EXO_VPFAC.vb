Imports SAPbouiCOM

Public Class EXO_VPFAC
    Private objGlobal As EXO_UIAPI.EXO_UIAPI

    Public Sub New(ByRef objG As EXO_UIAPI.EXO_UIAPI)
        Me.objGlobal = objG
    End Sub
    Public Function SBOApp_MenuEvent(ByVal infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnFPPDTE"
                        'Cargamos pantalla de gestión.
                        If CargarFormVPFAC() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Function CargarFormVPFAC() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        CargarFormVPFAC = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_VPFAC.srf")

            Try
                oForm = objGlobal.SBOApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    objGlobal.SBOApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                Else
                    objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
                    Exit Function
                End If
            End Try

            CType(oForm.Items.Item("btnGen").Specific, SAPbouiCOM.Button).Item.Enabled = False
            CType(oForm.Items.Item("txtDFECHA").Specific, SAPbouiCOM.EditText).Value = Now.Date.Year.ToString("0000") & Now.Date.Month.ToString("00") & "01"
            CType(oForm.Items.Item("txtHFECHA").Specific, SAPbouiCOM.EditText).Value = Now.Date.Year.ToString("0000") & Now.Date.Month.ToString("00") & Now.Date.Day.ToString("00")
            CargarFormVPFAC = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function

    Public Function SBOApp_ItemEvent(ByVal infoEvento As ItemEvent) As Boolean
        Try
            'Apaño por un error que da EXO_Basic.dll al consultar infoEvento.FormTypeEx
            Try
                If infoEvento.FormTypeEx <> "" Then

                End If
            Catch ex As Exception
                Return False
            End Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_VPFAC"
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
                        Case "EXO_VPFAC"
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
                        Case "EXO_VPFAC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_VPFAC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return True

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos

            Select Case pVal.ItemUID
                Case "btnGen"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                        If ComprobarDOC(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btnGen").Enabled = False
                            'Generamos facturas
                            objGlobal.SBOApp.StatusBar.SetText("Creando Documentos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If CrearDocFACT(oForm, "DT_DOC", "FACTURAS", objGlobal.compañia, objGlobal.SBOApp, objGlobal) = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                            oForm.Items.Item("btnGen").Enabled = True
                        End If
                    End If
                Case "btnFiltro"
                    Cargar_Grid(oForm)
                    oForm.Items.Item("btnGen").Enabled = True
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Shared Function CrearDocFACT(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByVal sTDoc As String, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_UIAPI.EXO_UIAPI) As Boolean
        CrearDocFACT = False
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim sExiste As String = "" ' Para comprobar si existen los datos

        Dim sErrorDes As String = "" : Dim sDocAdd As String = "" : Dim sMensaje As String = ""
        Dim sModo As String = "" : Dim sTabla As String = ""

        Dim oRsCab As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLin As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Dim oRsSerie As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsSerieNumber As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim esprimeralinea As Boolean = True
        Dim esprimeraportes As Boolean = True
        Try
            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            'Company.StartTransaction()
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
#Region "Cabecera"
                    esprimeralinea = True
#Region "Tipo Documento"
                    sModo = ""
                    If sModo = "F" Then
                        oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)
                        sTabla = "OINV"
                    Else
                        oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts), SAPbobsCOM.Documents)
                        sTabla = "ODRF"
                    End If

                    oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
#End Region
#Region " Serie o Num Documento"
                    Dim sAnno As String = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString)
                    sSQL = "SELECT ""Series"" "
                    sSQL += " FROM ""NNM1"" "
                    sSQL += " WHERE ""ObjectCode""=13 and ""Indicator""='" & sAnno & "' "
                    oRsSerie = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                    oRsSerie.DoQuery(sSQL)
                    If oRsSerie.RecordCount > 0 Then
                        Dim sSerieDoc As String = oRsSerie.Fields.Item("Series").Value.ToString
                        oDoc.Series = sSerieDoc
                    Else
                        objglobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado serie para el documento.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Exit Function
                    End If

#End Region
                    oDoc.CardCode = oForm.DataSources.DataTables.Item(sData).GetValue("IC", i).ToString
                    oDoc.NumAtCard = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Referencia", i).ToString
#Region "Fechas"
                    oDoc.DocDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString)
                    oDoc.TaxDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Factura", i).ToString)
#End Region
#End Region

#Region "Líneas"
                    'Buscamos las líneas del documento
                    sSQL = "Select * FROM ""RDR1"" Where ""DocEntry""=" & oForm.DataSources.DataTables.Item(sData).GetValue("Nº Int. Pedido", i).ToString & " ORDER BY ""LineNum"" "
                    oRsLin.DoQuery(sSQL)
                    For iLin = 1 To oRsLin.RecordCount
                        If esprimeralinea = False Then
                            oDoc.Lines.Add()
                        Else
                            esprimeralinea = False
                        End If
                        oDoc.Lines.BaseEntry = oRsLin.Fields.Item("DocEntry").Value
                        oDoc.Lines.BaseLine = oRsLin.Fields.Item("LineNum").Value
                        oDoc.Lines.BaseType = "17"
                        oRsLin.MoveNext()
                    Next
#End Region

                    'grabar el documento
                    If oDoc.Add() <> 0 Then 'Si ocurre un error en la grabación entra
                        sErrorDes = oCompany.GetLastErrorCode & " / " & oCompany.GetLastErrorDescription
                        oSboApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "ERROR")
                        oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sErrorDes)
                        oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, "")
                    Else
                        esprimeralinea = True
                        esprimeraportes = True
                        sDocAdd = oCompany.GetNewObjectKey() 'Recoge el último documento creado
                        oForm.DataSources.DataTables.Item(sData).SetValue("Nº Int. Factura", i, sDocAdd)
                        'Buscamos el documento para crear un mensaje
                        sDocAdd = EXO_GLOBALES.GetValueDB(oCompany, """" & sTabla & """", """DocNum""", """DocEntry""=" & sDocAdd)
                        If sModo = "F" Then
                            sModo = ""
                        Else
                            sModo = " borrador "
                        End If
                        oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "OK")
                        oForm.DataSources.DataTables.Item(sData).SetValue("Nº Factura", i, sDocAdd)


                        sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de ventas Nº" & sDocAdd

                        oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sMensaje)
                        oSboApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                End If
            Next

            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If

            CrearDocFACT = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDoc, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCab, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLin, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsSerie, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsSerieNumber, Object))
        End Try
    End Function
    Private Sub Cargar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim sFechaDesde As String = "" : Dim dFechaDesde As Date = Now.Date
        Dim sFechaHasta As String = "" : Dim dFechaHasta As Date = Now.Date
        Try
            oForm.Freeze(True)
            If oForm.DataSources.UserDataSources.Item("UDDFECHA").Value.ToString <> "" Then
                dFechaDesde = oForm.DataSources.UserDataSources.Item("UDDFECHA").Value.ToString
                sFechaDesde = dFechaDesde.Year.ToString("0000") & dFechaDesde.Month.ToString("00") & dFechaDesde.Day.ToString("00")
            Else
                sFechaDesde = ""
            End If

            If oForm.DataSources.UserDataSources.Item("UDHFECHA").Value.ToString <> "" Then
                dFechaHasta = oForm.DataSources.UserDataSources.Item("UDHFECHA").Value.ToString
                sFechaHasta = dFechaHasta.Year.ToString("0000") & dFechaHasta.Month.ToString("00") & dFechaHasta.Day.ToString("00")
            Else
                sFechaHasta = ""
            End If



            'Ahora cargamos el Grid con los datos
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' ""Sel"", '     ' ""Estado"",CAST('' as varchar(50)) ""Nº Int. Factura"",CAST('' as varchar(50)) ""Nº Factura"", T0.DocDate ""F. Factura"", "
            sSQL &= " T0.""DocEntry"" ""Nº Int. Pedido"", T0.""DocNum"" ""Nº Pedido"", isnull(T0.""NumAtCard"",'') ""Nº Referencia"", T0.""CardCode"" ""IC"", T0.""CardName"" ""Nombre"", "
            sSQL &= " T0.""DocDate"" ""Fecha Pedido"", T0.""DocTotal"" ""Imp. Pedido EUR"", T0.""DocTotalFC"" ""Imp. Pedido"", T0.""DocCur"" ""Moneda"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""ORDR"" T0 "
            sSQL &= " WHERE T0.""DocStatus""<>'C' "
            If sFechaDesde <> "" Then
                sSQL &= " and T0.""DocDate"">='" & sFechaDesde & "' "
            End If

            If sFechaHasta <> "" Then
                sSQL &= " and T0.""DocDate""<='" & sFechaHasta & "' "
            End If

            sSQL &= " ORDER BY T0.""DocDate"", T0.""DocNum"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 14
                Select Case i
                    Case 1, 3, 6, 7, 9, 10, 13
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                    Case 2
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "112"
                        oColumnTxt.Editable = False
                    Case 4
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = True
                    Case 5
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "17"
                        oColumnTxt.Editable = False
                    Case 8
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.LinkedObjectType = "2"
                        oColumnTxt.Editable = False
                    Case 11, 12
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.RightJustified = True
                        oColumnTxt.Editable = False

                    Case Else
                        CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                        oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                        oColumnTxt.Editable = False
                End Select
            Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
        End Try
    End Sub
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                objGlobal.SBOApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
