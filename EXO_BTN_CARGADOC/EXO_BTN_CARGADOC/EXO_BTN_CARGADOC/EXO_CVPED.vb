Imports System.Xml
Imports SAPbouiCOM
Imports OfficeOpenXml
Imports System.IO
Public Class EXO_CVPED
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
                    Case "EXO-MnCVPed"
                        'Cargamos pantalla de gestión.
                        If CargarFormCDOC() = False Then
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
    Public Function CargarFormCDOC() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        CargarFormCDOC = False

        Try
            oFP = CType(objGlobal.SBOApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.leerEmbebido(Me.GetType(), "EXO_CVPED.srf")

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

            CargaComboFormato(oForm)
            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Item.Enabled = False
            If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 1 Then
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("PEDIDOVENTA", BoSearchKey.psk_ByValue)
            Else
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("--", BoSearchKey.psk_ByValue)
            End If
            CargarFormCDOC = True
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
                        Case "EXO_CVPED"
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
                        Case "EXO_CVPED"
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
                        Case "EXO_CVPED"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
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
                        Case "EXO_CVPED"
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
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False

        Try
            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = CType(pVal, SAPbouiCOM.IChooseFromListEvent)
            Dim oDataTable As DataTable
            If pVal.ItemUID = "grd_DOC" AndAlso pVal.ChooseFromListUID = "CFL_0" Then
                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Comercial", pVal.Row, oDataTable.GetValue("SlpName", 0).ToString)
                    Catch ex As Exception
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Comercial", pVal.Row, oDataTable.GetValue("SlpName", 0).ToString)
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

    Private Function EventHandler_FORM_VISIBLE(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Items.Item("btn_Carga").Enabled = False
            End If

            EventHandler_FORM_VISIBLE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim sTipoArchivo As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sArchivo As String = objGlobal.pathHistorico & "\DOC_CARGADOS\" & objGlobal.SBOApp.Company.DatabaseName & "\VENTAS\PEDIDOS\"
        Dim sNomFICH As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            If System.IO.Directory.Exists(sArchivo) = False Then
                System.IO.Directory.CreateDirectory(sArchivo)
            End If
            Select Case pVal.ItemUID
                Case "btn_Carga"
                    If objGlobal.SBOApp.MessageBox("¿Está seguro que quiere generar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                        If ComprobarDOC(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btn_Carga").Enabled = False
                            'Generamos Documentos
                            objGlobal.SBOApp.StatusBar.SetText("Creando Documentos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If EXO_GLOBALES.CrearDocumentos(oForm, "DT_DOC", "PEDIDO", objGlobal.compañia, objGlobal.SBOApp, objGlobal) = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objGlobal.SBOApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    End If
                Case "btn_Fich"
                    Limpiar_Grid(oForm)
                    'Cargar Fichero para leer
                    If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString <> "--" Then
                        If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString = "XML" Then
                            sTipoArchivo = "XML|*.xml"
                        Else
                            sSQL = "Select ""U_EXO_TEXP"" FROM ""@EXO_CFCNF""  WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                Select Case oRs.Fields.Item("U_EXO_TEXP").Value.ToString
                                    Case "1" : sTipoArchivo = "Ficheros CSV|*.csv|Texto|*.txt"
                                    Case "2" : sTipoArchivo = "Libro de Excel|*.xlsx|Excel 97-2003|*.xls"
                                    Case Else
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Error inesperado. No ha encontrado el tipo de fichero a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oForm.Items.Item("btn_Carga").Enabled = False
                                        Exit Function
                                End Select
                            End If
                        End If
                        'Tenemos que controlar que es cliente o web
                        If objGlobal.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                            sArchivoOrigen = objGlobal.SBOApp.GetFileFromBrowser() 'Modificar
                        Else
                            'Controlar el tipo de fichero que vamos a abrir según campo de formato
                            sArchivoOrigen = objGlobal.funciones.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                        End If

                        If Len(sArchivoOrigen) = 0 Then
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                            objGlobal.SBOApp.MessageBox("Debe indicar un archivo a importar.")
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            oForm.Items.Item("btn_Carga").Enabled = False
                            Exit Function
                        Else
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                            sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                            sArchivo = sArchivo & sNomFICH
                            'Hacemos copia de seguridad para tratarlo
                            Copia_Seguridad(sArchivoOrigen, sArchivo)
                            'Ahora abrimos el fichero para tratarlo
                            TratarFichero(sArchivo, CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm)
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    Else
                        objGlobal.SBOApp.MessageBox("No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Active = True
                    End If
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
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oColumnTxt, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oColumnChk, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPDOC"" where ""U_EXO_USR""='" & objGlobal.compañia.UserName & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPDOCL"" where ""U_EXO_USR""='" & objGlobal.compañia.UserName & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPDOCLT"" where ""U_EXO_USR""='" & objGlobal.compañia.UserName & "'  "
            oRs.DoQuery(sSQL)
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""Code"",""U_EXO_MODO"" as ""Modo"", '     ' as ""Estado"",""U_EXO_TIPOF"" As ""Tipo"",'      ' as ""DocEntry"", ""U_EXO_Serie"" as ""Serie"",""U_EXO_DOCNUM"" as ""Nº Documento"","
            sSQL &= " ""U_EXO_REF"" as ""Referencia"", ""U_EXO_MONEDA"" as ""Moneda"", ""U_EXO_COMER"" as ""Comercial"", ""U_EXO_CLISAP"" as ""Interlocutor SAP"", ""U_EXO_ADDID"" as ""Interlocutor Ext."", "
            sSQL &= " ""U_EXO_FCONT"" as ""F. Contable"", ""U_EXO_FDOC"" as ""F. Documento"", ""U_EXO_FVTO"" as ""F. Vto"", ""U_EXO_TDTO"" as ""T. Dto."", ""U_EXO_DTO"" as ""Dto."",  "
            sSQL &= " ""U_EXO_CPAGO"" as ""Vía Pago"", ""U_EXO_GROUPNUM"" as ""Cond. Pago"", ""U_EXO_COMENT"" as ""Comentario"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPDOC"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.compañia.UserName & "' "
            sSQL &= " ORDER BY ""U_EXO_MODO"", ""U_EXO_TIPOF"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub TratarFichero_Excel(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRCampos As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCampo As String = ""

        Dim iDoc As Integer = 0 'Contador de Cabecera de documentos
        Dim sTFac As String = "" : Dim sTFacColumna As String = "" : Dim sTipoLineas As String = "" : Dim sTDoc As String = ""
        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCodCliente As String = "" : Dim sClienteColumna As String = "" : Dim sCodClienteColumna As String = ""
        Dim sSerie As String = "" : Dim sDocNum As String = "" : Dim sManual As String = "" : Dim sSerieColumna As String = "" : Dim sDocNumColumna As String = ""
        Dim sNumAtCard As String = "" : Dim sNumAtCardColumna As String = ""
        Dim sMoneda As String = "" : Dim sMonedaColumna As String = ""
        Dim sEmpleado As String = ""
        Dim sFContable As String = "" : Dim sFecha As String = "" : Dim sFDocumento As String = "" : Dim sFVto As String = "" : Dim sFDocumentoColumna As String = "" : Dim sShipToCodeColumna As String = ""
        Dim sTipoDto As String = "" : Dim sDto As String = ""
        Dim sPeyMethod As String = "" : Dim sCondPago As String = ""
        Dim sDirFac As String = "" : Dim sDirEnv As String = ""
        Dim sComent As String = "" : Dim sComentCab As String = "" : Dim sComentPie As String = ""
        Dim sCondicion As String = ""

        Dim sExiste As String = ""
        Dim iLinea As Integer = 0 : Dim sCodCampos As String = ""

        Dim pck As ExcelPackage = Nothing
        Dim iLin As Integer = 0
        Try
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Dim excel As New FileInfo(sArchivo)
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook
                Dim worksheet = workbook.Worksheets.First()
                sSQL = "SELECT ""U_EXO_FEXCEL"",""U_EXO_CSAP"",""U_EXO_TDOC"" FROM ""@EXO_CFCNF"" WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    iLin = oRs.Fields.Item("U_EXO_FEXCEL").Value
                    sTDoc = oRs.Fields.Item("U_EXO_TDOC").Value
                    If sTDoc = "1" Then
                        sTDoc = "B"
                    Else
                        sTDoc = "F"
                    End If
                    sCodCampos = oRs.Fields.Item("U_EXO_CSAP").Value
                    sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPC"" WHERE ""Code""='" & sCodCampos & "'"
                    oRCampos.DoQuery(sSQL)
                    Dim sCamposC(oRCampos.RecordCount, 3) As String

                    If oRCampos.RecordCount > 0 Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Estructura de Cabecera...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Matrix de Cabecera"
                        For I = 1 To oRCampos.RecordCount
                            sCamposC(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                            sCamposC(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                            sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='C' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                            sCampo = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
                            If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
                                sMensaje &= "Por favor, Revise la parametrización."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Sub
                            End If
                            sCamposC(I, 3) = sCampo
                            Select Case oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                Case "ObjType" : sTFacColumna = sCampo
                                Case "CardCode" : sClienteColumna = sCampo
                                Case "ADDID" : sCodClienteColumna = sCampo
                                Case "Series" : sSerieColumna = sCampo
                                Case "DocNum" : sDocNumColumna = sCampo
                                Case "NumAtCard" : sNumAtCardColumna = sCampo
                                Case "DocCurrency" : sMonedaColumna = sCampo
                                Case "TaxDate" : sFDocumentoColumna = sCampo
                                Case "ShipToCode" : sShipToCodeColumna = sCampo
                            End Select
                            oRCampos.MoveNext()
                        Next
#End Region
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Estructura de Cabecera leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                    sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPL"" WHERE ""Code""='" & sCodCampos & "'"
                    oRCampos.DoQuery(sSQL)
                    If oRCampos.RecordCount > 0 Then
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Estructura de Lineas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Matrix de Líneas"
                        Dim sCamposL(oRCampos.RecordCount, 3) As String
                        For I = 1 To oRCampos.RecordCount
                            sCamposL(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                            sCamposL(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                            sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='L' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                            sCampo = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
                            If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
                                sMensaje &= "Por favor, Revise la parametrización."
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objGlobal.SBOApp.MessageBox(sMensaje)
                                Exit Sub
                            End If
                            sCamposL(I, 3) = sCampo
                            oRCampos.MoveNext()
                        Next
#End Region
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Estructura de Líneas leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        Do
#Region "Cabecera"
                            If sCliente <> worksheet.Cells(sClienteColumna & iLin).Text Or sDirEnv <> worksheet.Cells(sShipToCodeColumna & iLin).Text Or
                                sNumAtCard <> worksheet.Cells(sNumAtCardColumna & iLin).Text Or sFDocumento <> worksheet.Cells(sFDocumentoColumna & iLin).Text Then
                                'If sTFac <> worksheet.Cells(sTFacColumna & iLin).Text Or sCliente <> worksheet.Cells(sClienteColumna & iLin).Text Or sCodCliente <> worksheet.Cells(sCodClienteColumna & iLin).Text _
                                'Or sSerie <> worksheet.Cells(sSerieColumna & iLin).Text Or sDocNum <> worksheet.Cells(sDocNumColumna & iLin).Text Or sNumAtCard <> worksheet.Cells(sNumAtCardColumna & iLin).Text _
                                'Or sMoneda <> worksheet.Cells(sMonedaColumna & iLin).Text Or sFDocumento <> worksheet.Cells(sFDocumentoColumna & iLin).Text Then
                                'Grabamos la cabecera
                                For C = 1 To sCamposC.GetUpperBound(0)
                                    Select Case sCamposC(C, 1)
                                        Case "ObjType"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    If worksheet.Cells("A" & iLin).Text <> "" Then
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        objGlobal.SBOApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    Else
                                                        Exit Do
                                                    End If
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "DocType"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTipoLineas = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTipoLineas = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "CardCode"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "CardName"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "ADDID"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "NumAtCard"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sNumAtCard = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sNumAtCard = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "EXO_Manual"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sManual = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sManual = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "Series"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sSerie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sSerie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "DocNum"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDocNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDocNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "DocCurrency"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "SlpCode"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sEmpleado = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sEmpleado = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "DocDate"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFContable = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFContable = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                            If sFContable.Trim <> "" Then
                                                Dim Fecha() As String = sFContable.Split("-/")
                                                sFContable = ""
                                                If Fecha(2).Trim.Length = 2 Then
                                                    sFContable &= "20" & Fecha(2)
                                                Else
                                                    sFContable &= Fecha(2)
                                                End If
                                                If Fecha(1).Trim.Length = 1 Then
                                                    sFContable &= "0" & Fecha(1)
                                                Else
                                                    sFContable &= Fecha(1)
                                                End If

                                                If Fecha(0).Trim.Length = 1 Then
                                                    sFContable &= "0" & Fecha(0)
                                                Else
                                                    sFContable &= Fecha(0)
                                                End If
                                            End If
                                        Case "TaxDate"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFDocumento = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFDocumento = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                            If sFDocumento.Trim <> "" Then
                                                sFecha = sFDocumento
                                                Dim Fecha() As String = sFDocumento.Split("-/")
                                                sFDocumento = ""
                                                If Fecha(2).Trim.Length = 2 Then
                                                    sFDocumento &= "20" & Fecha(2)
                                                Else
                                                    sFDocumento &= Fecha(2)
                                                End If
                                                If Fecha(1).Trim.Length = 1 Then
                                                    sFDocumento &= "0" & Fecha(1)
                                                Else
                                                    sFDocumento &= Fecha(1)
                                                End If

                                                If Fecha(0).Trim.Length = 1 Then
                                                    sFDocumento &= "0" & Fecha(0)
                                                Else
                                                    sFDocumento &= Fecha(0)
                                                End If
                                            End If
                                        Case "DocDueDate"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFVto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFVto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                            If sFVto.Trim <> "" Then
                                                Dim Fecha() As String = sFVto.Split("-/")
                                                sFVto = ""
                                                If Fecha(2).Trim.Length = 2 Then
                                                    sFVto &= "20" & Fecha(2)
                                                Else
                                                    sFVto &= Fecha(2)
                                                End If
                                                If Fecha(1).Trim.Length = 1 Then
                                                    sFVto &= "0" & Fecha(1)
                                                Else
                                                    sFVto &= Fecha(1)
                                                End If

                                                If Fecha(0).Trim.Length = 1 Then
                                                    sFVto &= "0" & Fecha(0)
                                                Else
                                                    sFVto &= Fecha(0)
                                                End If
                                            End If
                                        Case "EXO_TDTO"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTipoDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTipoDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "EXO_DTO"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "PeyMethod"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sPeyMethod = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sPeyMethod = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "GroupNum"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCondPago = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCondPago = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "PayToCode"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDirFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDirFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "ShipToCode"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDirEnv = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sDirEnv = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "Comments"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sComent = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sComent = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "OpeningRemarks"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sComentCab = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sComentCab = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                        Case "ClosingRemarks"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sComentPie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sComentPie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
                                    End Select
                                Next
                                'Grabamos la cabecera

                                'Insertar en la tabla temporal la cabecera
                                If sNumAtCard <> "" Then
                                    iDoc += 1 : iLinea = 0
                                    If sMoneda = "" Then : sMoneda = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """Currency""", """CardCode""='" & sCliente & "' ") : End If 'En el caso de no estar indicado, se ha tomado moneda cliente por defecto
                                    If sDirFac = "" Then
                                        'Buscamos la dirección por defecto del cliente
                                        sDirFac = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """BillToDef""", """CardCode""='" & sCliente & "' ")
                                    End If
                                    If sDirEnv = "" Then
                                        'Buscamos la dirección por defecto del cliente
                                        sDirEnv = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """ShipToDef""", """CardCode""='" & sCliente & "' ")
                                    End If
                                    sSQL = "insert into ""@EXO_TMPDOC"" values('" & iDoc.ToString & "','" & iDoc.ToString & "'," & iDoc.ToString & ",'N','',0," & objGlobal.compañia.UserSignature
                                    sSQL &= ",'','',0,'',0,'','" & objGlobal.compañia.UserName & "',"
                                    sSQL &= "'" & sTDoc & "','" & sDocNum & "','17','" & sManual & "','" & sSerie & "','" & sNumAtCard & "','" & sMoneda & "','','" & sEmpleado & "',"
                                    sSQL &= "'" & sCliente & "','" & sCodCliente & "','" & sFContable & "','" & sFDocumento & "','" & sFVto & "','" & sTipoDto & "',"
                                    sSQL &= EXO_GLOBALES.DblNumberToText(objGlobal, sDto.ToString) & ",'" & sPeyMethod & "','" & sDirFac & "','" & sDirEnv & "','" & sComent.Replace("'", "") & "','"
                                    sSQL &= sComentCab.Replace("'", "") & "','" & sComentPie.Replace("'", "") & "','" & sCondPago & "') "
                                    oRs.DoQuery(sSQL)
                                    sFDocumento = sFecha
                                End If
                            End If
#End Region
                            'Ahora tratamos la línea
#Region "Líneas"
                            Dim sCuenta As String = "" : Dim sArt As String = "" : Dim sArtDes As String = ""
                            Dim sCantidad As String = "0.00" : Dim sprecio As String = "0.00" : Dim sDtoLin As String = "0.00" : Dim sTotalServicios As String = "0.00"
                            Dim sTextoAmpliado As String = "" : Dim sLinImpuestoCod As String = "" : Dim sLinRetCodigo As String = "" : Dim sPrecioBruto As String = "0.00" : Dim sReparto As String = ""

                            If oRCampos.RecordCount > 0 Then
                                For L = 1 To sCamposL.GetUpperBound(0)
                                    Select Case sCamposL(L, 1)
                                        Case "AcctCode"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sCuenta = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sCuenta = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sCuenta = ""
                                                End If
                                            End If
                                        Case "ItemCode"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sArt = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sArt = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sArt = ""
                                                End If
                                            End If
                                        Case "Dscription"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sArtDes = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sArtDes = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sArtDes = ""
                                                End If
                                            End If
                                        Case "Quantity"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sCantidad = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sCantidad = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sCantidad = "0.00"
                                                End If
                                            End If
                                        Case "UnitPrice"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sprecio = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sprecio = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sprecio = "0.00"
                                                End If
                                            End If
                                        Case "DiscPrcnt"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sDtoLin = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sDtoLin = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sDtoLin = "0.00"
                                                End If
                                            End If
                                        Case "EXO_IMPSRV"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sTotalServicios = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sTotalServicios = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sTotalServicios = "0.00"
                                                End If
                                            End If
                                        Case "EXO_TextoLin"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sTextoAmpliado = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sTextoAmpliado = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sTextoAmpliado = ""
                                                End If
                                            End If
                                        Case "EXO_IMP"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sLinImpuestoCod = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sLinImpuestoCod = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sLinImpuestoCod = ""
                                                End If
                                            End If
                                        Case "EXO_RET"
                                            If sCamposL(L, 2) = "Y" Then
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sLinRetCodigo = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    objGlobal.SBOApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                    sLinRetCodigo = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                Else
                                                    sLinRetCodigo = ""
                                                End If
                                            End If
                                        Case "GrossBuyPr"
#Region "GrossBuyPr"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sPrecioBruto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        objGlobal.SBOApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sPrecioBruto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sPrecioBruto = "0.00"
                                                    End If
                                                End If
                                            Else
                                                sPrecioBruto = "0.00"
                                            End If

#End Region
                                        Case "EXO_REPARTO"
#Region "EXO_REPARTO"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sReparto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        objGlobal.SBOApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sReparto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sReparto = ""
                                                    End If
                                                End If
                                            Else
                                                sReparto = ""
                                            End If

#End Region
                                    End Select
                                Next
#Region "Comprobar datos línea"
                                'Comprobamos que exista la cuenta
                                If sCuenta <> "" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OACT""", """AcctCode""", """AcctCode""='" & sCuenta & "'")
                                    If sExiste = "" Then
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - La Cuenta contable SAP  - " & sCuenta & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objGlobal.SBOApp.MessageBox("La Cuenta contable SAP - " & sCuenta & " - no existe.")
                                        Exit Sub
                                    End If
                                End If
                                'Comprobamos que exista el artículo
                                If sTipoLineas = "" Then
                                    sTipoLineas = "I"
                                End If
                                If sTipoLineas = "I" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITM""", """ItemCode""", """ItemCode"" like '" & sArt & "'")
                                    If sExiste = "" Then
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objGlobal.SBOApp.MessageBox("El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.")
                                        Exit Sub
                                    End If
                                ElseIf sTipoLineas = "S" Then
                                    If sCuenta = "" Then
                                        ' No puede estar la cuenta vacía si es de tipo servicio
                                        Dim sMensaje As String = " La cuenta en la línea del servicio no puede estar vacía. Por favor, Revise los datos."
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objGlobal.SBOApp.MessageBox(sMensaje)
                                        Exit Sub
                                    End If
                                End If
                                'Comprobamos que exista el impuesto si está relleno
                                If sLinImpuestoCod <> "" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OVTG""", """Code""", """Code""='" & sLinImpuestoCod & "'")
                                    If sExiste = "" Then
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objGlobal.SBOApp.MessageBox("El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.")
                                        Exit Sub
                                    End If
                                End If
                                'Comprobamos que exista la retención si está relleno
                                If sLinRetCodigo <> "" Then
                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """CRD4""", """WTCode""", """CardCode""='" & sCliente & "' and ""WTCode""='" & sLinRetCodigo & "'")
                                    If sExiste = "" Then
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objGlobal.SBOApp.MessageBox("El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.")
                                        Exit Sub
                                    End If
                                End If
#End Region
                                'Grabamos la línea
                                sSQL = "insert into ""@EXO_TMPDOCL"" values('" & iDoc.ToString & "','" & iLinea & "','',0,'" & objGlobal.compañia.UserName & "',"
                                sSQL &= "'" & sCuenta & "','" & sArt & "','" & sArtDes & "'," & EXO_GLOBALES.DblNumberToText(objGlobal, sCantidad.ToString).Replace(",", ".") & ","
                                sSQL &= EXO_GLOBALES.DblNumberToText(objGlobal, sprecio.ToString) & "," & EXO_GLOBALES.DblNumberToText(objGlobal, sDtoLin.ToString)
                                sSQL &= "," & EXO_GLOBALES.DblNumberToText(objGlobal, sTotalServicios.ToString).Replace(",", ".") & ",'" & sLinImpuestoCod & "','" & sLinRetCodigo & "','"
                                sSQL &= sTextoAmpliado & "','" & sTipoLineas & "'," & sPrecioBruto & ",'" & sReparto & "' ) "
                                oRs.DoQuery(sSQL)
                                iLin += 1 : iLinea += 1
                            End If
#End Region
                        Loop While sTFac <> ""
                    End If


                Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.")
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("No se ha encontrado el fichero excel a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            pck.Dispose()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
        End Try
    End Sub
    Private Sub TratarFichero_TXT(ByVal sArchivo As String, ByVal sDelimitador As String, ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRCampos As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCampo As String = ""

        Dim iDoc As Integer = 0 'Contador de Cabecera de documentos
        Dim sTFac As String = "" : Dim sTFacColumna As String = "" : Dim sTipoLineas As String = "" : Dim sTDoc As String = ""
        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCodCliente As String = "" : Dim sClienteColumna As String = "" : Dim sCodClienteColumna As String = ""
        Dim sSerie As String = "" : Dim sDocNum As String = "" : Dim sManual As String = "" : Dim sSerieColumna As String = "" : Dim sDocNumColumna As String = ""
        Dim sDIR As String = "" : Dim sPob As String = "" : Dim sProv As String = "" : Dim sCPos As String = ""
        Dim sNumAtCard As String = "" : Dim sNumAtCardColumna As String = ""
        Dim sMoneda As String = "EUR" : Dim sMonedaColumna As String = ""
        Dim sEmpleado As String = ""
        Dim sFContable As String = "" : Dim sFDocumento As String = "" : Dim sFVto As String = "" : Dim sFDocumentoColumna As String = ""
        Dim sTipoDto As String = "" : Dim sDto As String = ""
        Dim sPeyMethod As String = "" : Dim sCondPago As String = ""
        Dim sDirFac As String = "" : Dim sDirEnv As String = ""
        Dim sComent As String = "" : Dim sComentCab As String = "" : Dim sComentPie As String = ""
        Dim sCondicion As String = ""

        Dim sExiste As String = ""
        Dim bCrearCli As Boolean = False
        Dim iLinea As Integer = 0 : Dim sCodCampos As String = ""
        Dim bSaltaCabecera As Boolean = True
        Dim sMensaje As String = ""
        Dim sCamposC(1, 3) As String : Dim sCamposL(1, 3) As String

        ' Apuntador libre a archivo
        Dim Apunt As Integer = FreeFile()
        ' Variable donde guardamos cada línea de texto
        Dim Texto As String = ""
        Dim sValorCampo As String = ""

        Dim sDocumento As String = ""
        Try
            'Tengo que buscar en la tabla el último numero de documento
            iDoc = objGlobal.refDi.SQL.sqlNumericaB1("SELECT isnull(MAX(cast(CODE as int)),0) FROM ""@EXO_TMPDOC"" ")
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Using MyReader As New Microsoft.VisualBasic.
                        FileIO.TextFieldParser(sArchivo, System.Text.Encoding.UTF7)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    Select Case sDelimitador
                        Case "1" : MyReader.SetDelimiters(vbTab)
                        Case "2" : MyReader.SetDelimiters(";")
                        Case "3" : MyReader.SetDelimiters(",")
                        Case "4" : MyReader.SetDelimiters("-")
                        Case Else : MyReader.SetDelimiters(vbTab)
                    End Select

                    Dim currentRow As String()
                    While Not MyReader.EndOfData
                        Try
                            If bSaltaCabecera = True Then
                                currentRow = MyReader.ReadFields()
                                bSaltaCabecera = False
                            End If
                            currentRow = MyReader.ReadFields()
                            Dim currentField As String
                            Dim scampos(1) As String
                            Dim iCampo As Integer = 0
                            For Each currentField In currentRow
                                iCampo += 1
                                ReDim Preserve scampos(iCampo)
                                scampos(iCampo) = currentField
                                'SboApp.MessageBox(scampos(iCampo))
                            Next
                            'Buscamos campos para traducir 
                            sSQL = "SELECT ""U_EXO_FEXCEL"",""U_EXO_CSAP"",""U_EXO_TDOC"" FROM ""@EXO_CFCNF"" "
                            sSQL &= " WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                sTDoc = oRs.Fields.Item("U_EXO_TDOC").Value
                                If sTDoc = "1" Then
                                    sTDoc = "B"
                                Else
                                    sTDoc = "F"
                                End If
                                sCodCampos = oRs.Fields.Item("U_EXO_CSAP").Value
                                sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPC"" WHERE ""Code""='" & sCodCampos & "'"
                                oRCampos.DoQuery(sSQL)
                                If oRCampos.RecordCount > 0 Then
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Estructura de Cabecera...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Matrix de Cabecera"
                                    ReDim sCamposC(oRCampos.RecordCount, 3)
                                    For I = 1 To oRCampos.RecordCount
                                        sCamposC(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                        sCamposC(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                                        sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='C' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                                        sCampo = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_FCCFL""", """U_EXO_posTXT""", sCondicion)
                                        If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                            ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                            sMensaje = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en el fichero TXT y es obligatorio." & ChrW(13) & ChrW(10)
                                            sMensaje &= "Por favor, Revise la parametrización."
                                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objGlobal.SBOApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                        sCamposC(I, 3) = sCampo
                                        oRCampos.MoveNext()
                                    Next
#End Region
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Estructura de Cabecera leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If

                                Dim sCuenta As String = "" : Dim sArt As String = "" : Dim sArtDes As String = ""
                                Dim sCantidad As String = "0.00" : Dim sprecio As String = "0.00" : Dim sDtoLin As String = "0.00" : Dim sTotalServicios As String = "0.00" : Dim sPrecioBruto As String = "0.00"
                                Dim sTextoAmpliado As String = "" : Dim sLinImpuestoCod As String = "" : Dim sLinRetCodigo As String = ""
                                sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPL"" WHERE ""Code""='" & sCodCampos & "'"
                                oRCampos.DoQuery(sSQL)
                                If oRCampos.RecordCount > 0 Then
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Estructura de líneas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Matrix de Líneas"
                                    ReDim sCamposL(oRCampos.RecordCount, 3)
                                    For I = 1 To oRCampos.RecordCount
                                        sCamposL(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                        sCamposL(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                                        sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='L' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                                        sCampo = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_FCCFL""", """U_EXO_posTXT""", sCondicion)
                                        If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                            ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                            sMensaje = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en el TXT y es obligatorio." & ChrW(13) & ChrW(10)
                                            sMensaje &= "Por favor, Revise la parametrización."
                                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objGlobal.SBOApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                        sCamposL(I, 3) = sCampo
                                        oRCampos.MoveNext()
                                    Next
#End Region
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Estructura de Líneas leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
#Region "Lectura cabecera"
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Valores de Cabecera...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                For C = 1 To sCamposC.GetUpperBound(0)
                                    Select Case sCamposC(C, 1)
                                        Case "ObjType"
                                            If sCamposC(C, 3) <> "" Then
                                                sTFac = sCamposC(C, 3)
                                            End If
                                        Case "DocType"
                                            If sCamposC(C, 3) <> "" Then
                                                sTipoLineas = sCamposC(C, 3)
                                            End If
                                        Case "CardCode"
                                            sCliente = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sCliente = "" Then
                                                Exit Sub
                                            ElseIf sCliente = "" Then
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Al leer el archivo nos encontramos con el código vacío. No se tratarán más datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                Exit Sub
                                            End If
                                            'Por lo que se ha visto no se da el Código de SAP, Por lo que buscaremos por el código de SAP, 
                                            'Buscamos por el CODIGO DE SAP
                                            sExiste = ""
                                            sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """CardCode""", """CardCode""='" & sCliente & "'")
                                            If sExiste = "" Then
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Interlocutor  - " & sCliente & " - no existe. Se buscará por el campo ID Número 2.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OCRD""", """CardCode""", """AddID""='" & sCliente & "'")
                                                If sExiste = "" Then
                                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Interlocutor  - " & sCliente & " - no existe al buscarlo por ID Número 2.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    'bCrearCli = True
                                                Else
                                                    sCliente = sExiste
                                                End If
                                            End If
                                        Case "CardName"
                                            sCliNombre = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sCliNombre = "" Then
                                                Exit Sub
                                            End If
                                        Case "ADDID"
                                            sCodCliente = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sCodCliente = "" Then
                                                Exit Sub
                                            End If
                                        Case "NumAtCard"
                                            sNumAtCard = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sNumAtCard = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_Manual"
                                            sManual = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sManual = "" Then
                                                Exit Sub
                                            End If
                                        Case "Series"
                                            sSerie = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sSerie = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocNum"
                                            sDocNum = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sDocNum = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocCurrency"
                                            sMoneda = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sMoneda = "" Then
                                                Exit Sub
                                            End If
                                        Case "SlpCode"
                                            sEmpleado = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sEmpleado = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocDate"
                                            sFContable = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sFContable = "" Then
                                                Exit Sub
                                            End If
                                        Case "TaxDate"
                                            sFDocumento = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sFDocumento = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocDueDate"
                                            sFVto = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sFVto = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_TDTO"
                                            sTipoDto = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sTipoDto = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_DTO"
                                            sDto = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sDto = "" Then
                                                Exit Sub
                                            End If
                                        Case "PeyMethod"
                                            sPeyMethod = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sPeyMethod = "" Then
                                                Exit Sub
                                            End If
                                        Case "GroupNum"
                                            sCondPago = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sCondPago = "" Then
                                                Exit Sub
                                            End If
                                        Case "PayToCode"
                                            sDirFac = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sDirFac = "" Then
                                                Exit Sub
                                            ElseIf sDirFac = "" Then
                                                sDirFac = "Facturación"
                                            End If
                                        Case "ShipToCode"
                                            sDirEnv = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sDirEnv = "" Then
                                                Exit Sub
                                            ElseIf sDirEnv = "" Then
                                                sDirEnv = "Entrega"
                                            End If
                                        Case "Comments"
                                            sComent = "Importado a través del fichero - " & sArchivo & " - "
                                            sComent &= Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sComent = "Importado a través del fichero - " & sArchivo & " - " Then
                                                Exit Sub
                                            End If
                                        Case "OpeningRemarks"
                                            sComentCab = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sComentCab = "" Then
                                                Exit Sub
                                            End If
                                        Case "ClosingRemarks"
                                            sComentPie = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sComentPie = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_DIR"
                                            sDIR = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sDIR = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_POB"
                                            sPob = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sPob = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_PRO"
                                            sProv = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sProv = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_CPOS"
                                            sCPos = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos)
                                            If sCamposC(C, 2) = "Y" And sCPos = "" Then
                                                Exit Sub
                                            End If
                                    End Select
                                Next
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Valores de Cabecera leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
#Region "Comprobar datos cabecera"
                                If sTFac = "" Then : sTFac = "17" : End If ' En el caso de no estar indicado, se ha tomado como Pedido de Compra
                                If sTipoLineas = "" Then : sTipoLineas = "I" : End If ' En el caso de no estar indicado, se ha tomado como que son líneas de servicio
                                'Comprobamos que se haya introducido manual o con una serie
                                If sDocNum = "" Then
                                    If sSerie = "" Then
                                        sSerie = objGlobal.refDi.SQL.sqlStringB1("SELECT ""SeriesName"" FROM NNM1 WHERE ""ObjectCode""=17")
                                        sMensaje &= "No se ha indicado ni Nº de documento ni serie. Se uilizará la serie por defecto."
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
                                    sManual = "N"
                                Else
                                    sManual = "Y"
                                End If
                                If sMoneda = "" Then : sMoneda = "EUR" : End If 'En el caso de no estar indicado, se ha tomado EUR por defecto
                                If sFContable = "" Then
                                    sFContable = Now.Year.ToString("0000") & "-" & Now.Month.ToString("00") & "-" & Now.Day.ToString("00")
                                    sMensaje &= "No se ha indicado una fecha Contable para el documento. Se indica la fecha actual."
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Else
                                    'Ponemos formato para SQL
                                    sFContable = Year(sFContable) & "-" & Month(sFContable) & "-" & Day(sFContable)
                                    If sFDocumento = "" Then
                                        sMensaje &= "No se ha indicado una fecha de documento. Se actualizará con la fecha contable."
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        sFDocumento = sFContable
                                    Else
                                        sFDocumento = Year(sFDocumento) & "-" & Month(sFDocumento) & "-" & Day(sFDocumento)
                                    End If
                                End If
                                If sTipoDto = "" Then : sTipoDto = "%" : End If ' Se toma si no tiene valor que el dto va en Porcentaje
                                If sDto = "" Then : sDto = "0.00" : End If ' Se toma por defecto dto valor a 0.00

                                If bCrearCli = True Then
                                    'Creamos el cliente con los datos que nos han dado
                                    'Busco y compruebo que exista la serie que han marcado en los parametros por defecto depediento si es de venta o de compra
                                    Dim sSerieIC As String = Nothing : Dim sTipoIC As String = ""
                                    Select Case sTFac
                                        Case "13", "14"
                                            sSerieIC = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_CFCNF""", """U_EXO_SERIEV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            sTipoIC = "C"
                                        Case "18", "19", "22"
                                            sSerieIC = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_CFCNF""", """U_EXO_SERIEC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            sTipoIC = "S"
                                    End Select
                                    If sSerieIC = "" Then
                                        sMensaje &= "No se ha indicado la serie para crear el interlocutor. No se puede continuar. Revise la parametrización."
                                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    Else
                                        EXO_GLOBALES.CrearInterlocutorSencillo(CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, sCliente,
                                                                  sSerieIC, sTipoIC, sCliNombre, "ES", sCliente, sCodCliente, "1", sCondPago, sPeyMethod,
                                                                  sDirFac, sDirEnv, sDIR, sPob, sProv, sCPos, "ES", objGlobal.compañia, objGlobal.SBOApp, objGlobal)
                                    End If
                                End If
                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Datos de cabecera comprobados.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                                'Grabamos la cabecera
                                'Insertar en la tabla temporal la cabecera
                                'Antes de insertar, comprobamos las direcciones de entrega y de facturación para comprobar que si son las de defecto del desarrollo, debemos buscar las de por defecto del cliente
                                If sDirFac = "Facturación" Then
                                    sSQL = "SELECT ""BillToDef"" FROM ""OCRD"" WHERE ""CardCode""='" & sCliente & "' "
                                    Dim sDIrFacDef As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    If sDIrFacDef <> "" Then
                                        sDirFac = sDIrFacDef
                                    End If
                                End If
                                If sDirEnv = "Entrega" Then
                                    sSQL = "SELECT ""ShipToDef"" FROM ""OCRD"" WHERE ""CardCode""='" & sCliente & "' "
                                    Dim sDirEnvDef As String = objGlobal.refDi.SQL.sqlStringB1(sSQL)
                                    If sDirEnvDef <> "" Then
                                        sDirEnv = sDirEnvDef
                                    End If
                                End If

                                If sTFac <> "" Then
                                    If sDocumento.Trim <> sDocNum.Trim Or iDoc = 0 Then
                                        sDocumento = sDocNum.Trim
                                        iDoc += 1
                                        iLinea = 0
                                        sSQL = "insert into ""@EXO_TMPDOC"" values('" & iDoc.ToString & "','" & iDoc.ToString & "'," & iDoc.ToString & ",'N','',0," & objGlobal.compañia.UserSignature
                                        sSQL &= ",'','',0,'',0,'','" & objGlobal.compañia.UserName & "',"
                                        sSQL &= "'" & sTDoc & "','" & sDocNum & "','" & sTFac & "','" & sManual & "','" & sSerie & "','" & sNumAtCard & "','" & sMoneda & "','','" & sEmpleado & "',"
                                        sSQL &= "'" & sCliente & "','" & sCodCliente & "','" & sFContable & "','" & sFDocumento & "','" & sFVto & "','" & sTipoDto & "',"
                                        sSQL &= EXO_GLOBALES.DblNumberToText(objGlobal, sDto.ToString) & ",'" & sPeyMethod & "','" & sDirFac & "','" & sDirEnv & "','" & sComent.Replace("'", "") & "','"
                                        sSQL &= sComentCab.Replace("'", "") & "','" & sComentPie.Replace("'", "") & "','" & sCondPago & "') "
                                        oRs.DoQuery(sSQL)
                                    Else
                                        iLinea += 1
                                    End If

#Region "Lectura de Líneas"
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Leyendo Valores de Líneas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    'Ahora la Línea
                                    For L = 1 To sCamposL.GetUpperBound(0)
                                        Select Case sCamposL(L, 1)
                                            Case "AcctCode"
                                                sCuenta = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sCuenta = "" Then
                                                    Exit Sub
                                                End If
                                            Case "ItemCode"
                                                sArt = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sArt = "" Then
                                                    Exit Sub
                                                End If
                                            Case "Dscription"
                                                sArtDes = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sArtDes = "" Then
                                                    Exit Sub
                                                End If
                                            Case "Quantity"
                                                sCantidad = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sCantidad = "" Then
                                                    Exit Sub
                                                End If
                                            Case "UnitPrice"
                                                sprecio = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sprecio = "" Then
                                                    Exit Sub
                                                End If
                                            Case "DiscPrcnt"
                                                sDtoLin = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sDtoLin = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_IMPSRV"
                                                sTotalServicios = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sTotalServicios = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_TextoLin"
                                                sTextoAmpliado = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sTextoAmpliado = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_IMP"
                                                sLinImpuestoCod = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sLinImpuestoCod = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_RET"
                                                sLinRetCodigo = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sLinRetCodigo = "" Then
                                                    Exit Sub
                                                End If
                                            Case "GrossBuyPr"
                                                sPrecioBruto = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos)
                                                If sCamposL(L, 2) = "Y" And sPrecioBruto = "" Then
                                                    Exit Sub
                                                End If
                                        End Select
                                    Next
                                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Valores de líneas leídos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region

#Region "Comprobar datos línea"
                                    'Comprobamos que exista la cuenta                                  
                                    If sCuenta <> "" Then
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OACT""", """AcctCode""", """AcctCode""='" & sCuenta & "'")
                                        If sExiste = "" Then
                                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - La Cuenta contable SAP  - " & sCuenta & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objGlobal.SBOApp.MessageBox("La Cuenta contable SAP - " & sCuenta & " - no existe.")
                                            Exit Sub
                                        End If
                                    End If
                                    'Comprobamos que exista el artículo
                                    If sTipoLineas = "I" Then
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OITM""", """ItemCode""", """ItemCode"" like '" & sArt & "'")
                                        If sExiste = "" Then
                                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objGlobal.SBOApp.MessageBox("El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.")
                                            Exit Sub
                                        End If
                                    ElseIf sTipoLineas = "S" Then
                                        If sCuenta = "" Then
                                            ' No puede estar la cuenta vacía si es de tipo servicio
                                            sExiste = ""
                                            sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_CFCNF""", """U_EXO_CSRV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            If sExiste = "" Then
                                                sMensaje = " La cuenta en la línea del servicio no puede estar vacía. Por favor, Revise los datos de la parametrización."
                                                objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                objGlobal.SBOApp.MessageBox(sMensaje)
                                                Exit Sub
                                            Else
                                                sCuenta = sExiste
                                            End If
                                        End If
                                    End If
                                    'Comprobamos que exista el impuesto si está relleno
                                    If sLinImpuestoCod = "" Then
                                        Select Case sTFac
                                            Case "13", "14" 'Ventas
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_CFCNF""", """U_EXO_IVAV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "18", "19", "22" 'Compras
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """@EXO_CFCNF""", """U_EXO_IVAC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    Else
                                        sLinImpuestoCod = sLinImpuestoCod.Replace(",", ".")
                                        Select Case sTFac
                                            Case "13", "14" 'Ventas
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OVTG""", """Code""", """Rate""='" & sLinImpuestoCod & "' and  LENGTH(""Code"")=2 and left(""Code"",1)='R' and ""Category""='O' ")
                                            Case "18", "19", "22" 'Compras
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OVTG""", """Code""", """Rate""='" & sLinImpuestoCod & "' and  LENGTH(""Code"")=2 and left(""Code"",1)='S' and ""Category""='I' ")
                                        End Select
                                    End If
                                    If sLinImpuestoCod <> "" Then
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """OVTG""", """Code""", """Code""='" & sLinImpuestoCod & "'")
                                        If sExiste = "" Then
                                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objGlobal.SBOApp.MessageBox("El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.")
                                            Exit Sub
                                        End If
                                    End If
                                    'Comprobamos que exista la retención si está relleno
                                    If sLinRetCodigo <> "" Then
                                        sExiste = EXO_GLOBALES.GetValueDB(objGlobal.compañia, """CRD4""", """WTCode""", """CardCode""='" & sCliente & "' and ""WTCode""='" & sLinRetCodigo & "'")
                                        If sExiste = "" Then
                                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objGlobal.SBOApp.MessageBox("El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.")
                                            Exit Sub
                                        End If
                                    End If
#End Region
                                    'Grabamos la línea
                                    sSQL = "insert into ""@EXO_TMPDOCL"" values('" & iDoc.ToString & "','" & iLinea & "','',0,'" & objGlobal.compañia.UserName & "',"
                                    sSQL &= "'" & sCuenta & "','" & sArt & "','" & sArtDes & "'," & EXO_GLOBALES.DblNumberToText(objGlobal, sCantidad.ToString).Replace(",", ".") & ","
                                    sSQL &= EXO_GLOBALES.DblNumberToText(objGlobal, sprecio.ToString) & "," & EXO_GLOBALES.DblNumberToText(objGlobal, sDtoLin.ToString)
                                    sSQL &= "," & EXO_GLOBALES.DblNumberToText(objGlobal, sTotalServicios.ToString).Replace(",", ".") & ",'" & sLinImpuestoCod & "','" & sLinRetCodigo & "','"
                                    sSQL &= sTextoAmpliado & "','" & sTipoLineas & "'," & sPrecioBruto & ",'-' ) "
                                    oRs.DoQuery(sSQL)
                                End If
                            End If
                        Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Línea " & ex.Message & " no es válida y se omitirá.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objGlobal.SBOApp.MessageBox("Línea " & ex.Message & " no es válida y se omitirá.")
                        End Try
                    End While
                End Using
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero txt a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            ' Cerramos el archivo
            FileClose(Apunt)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
        End Try
    End Sub
    Private Function Leer_Campo(ByVal sCampo As String, ByVal sColumna As String, ByVal sObligatorio As String, ByRef sVCampo() As String) As String
        Leer_Campo = ""
        Dim sValor As String = ""
        Dim icampo As Integer = 0
        Try
            If sColumna <> "" Then
                icampo = CInt(sColumna)
            End If
            If sVCampo(icampo) <> "" Then
                sValor = sVCampo(icampo)
            Else
                If sObligatorio = "Y" Then
                    Mensaje_CampoObligatorio(sCampo, sColumna)
                End If
            End If
            Leer_Campo = sValor
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Private Sub Mensaje_CampoObligatorio(ByVal sCampo As String, ByVal sColumna As String)
        Dim sMensaje As String = "El campo """ & sCampo & """ es obligatorio y la columna """ & sColumna & """ está vacía." & ChrW(13) & ChrW(10)
        sMensaje &= "Por favor, Revise el documento."
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        objGlobal.SBOApp.MessageBox(sMensaje)
    End Sub
    Private Sub TratarFichero(ByVal sArchivo As String, ByVal sTipoArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = ""
        Try
            sSQL = "Select ""U_EXO_STXT"" FROM ""@EXO_CFCNF""  WHERE ""Code""='" & sTipoArchivo & "'"
            sDelimitador = objGlobal.refDi.SQL.sqlStringB1(sSQL)

            sSQL = "Select ""U_EXO_TEXP"" FROM ""@EXO_CFCNF""  WHERE ""Code""='" & sTipoArchivo & "'"
            sTipoArchivo = objGlobal.refDi.SQL.sqlStringB1(sSQL)

            Select Case sTipoArchivo
                Case "1"
#Region "TXT|CSV"
                    TratarFichero_TXT(sArchivo, sDelimitador, oForm)
#End Region
                Case "2"
#Region "EXCEL"
                    TratarFichero_Excel(sArchivo, oForm)
#End Region
                Case Else
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) -El tipo de fichero a importar no está contemplado. Avise a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    objGlobal.SBOApp.MessageBox("El tipo de fichero a importar no está contemplado. Avise a su Administrador.")
                    Exit Sub
            End Select
            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Se ha leido correctamente el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

#Region "cargar Grid con los datos leidos"
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""Code"",""U_EXO_MODO"" as ""Modo"", '     ' as ""Estado"",""U_EXO_TIPOF"" As ""Tipo"",'      ' as ""DocEntry"",""U_EXO_Serie"" as ""Serie"",""U_EXO_DOCNUM"" as ""Nº Documento"","
            sSQL &= " ""U_EXO_REF"" as ""Referencia"", ""U_EXO_MONEDA"" as ""Moneda"", ""U_EXO_COMER"" as ""Comercial"", ""U_EXO_CLISAP"" as ""Interlocutor SAP"", ""U_EXO_ADDID"" as ""Interlocutor Ext."", "
            sSQL &= " ""U_EXO_FCONT"" as ""F. Contable"", ""U_EXO_FDOC"" as ""F. Documento"", ""U_EXO_FVTO"" as ""F. Vto"", ""U_EXO_TDTO"" as ""T. Dto."", ""U_EXO_DTO"" as ""Dto."",  "
            sSQL &= " ""U_EXO_CPAGO"" as ""Vía Pago"", ""U_EXO_GROUPNUM"" as ""Cond. Pago"", ""U_EXO_COMENT"" as ""Comentario"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPDOC"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.compañia.UserName & "' "
            sSQL &= " ORDER BY ""U_EXO_MODO"", ""U_EXO_TIPOF"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
#End Region
            oForm.Freeze(True)
            objGlobal.SBOApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.SBOApp.MessageBox("Se ha leido correctamente el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            For i = 1 To 3
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                If i = 2 Then
                    oColumnTxt.Visible = False
                End If
            Next

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(4).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oColumnCb = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(4), SAPbouiCOM.ComboBoxColumn)
            oColumnCb.ValidValues.Add("13", "Factura de Ventas")
            oColumnCb.ValidValues.Add("14", "Abonos de Venta")
            oColumnCb.ValidValues.Add("18", "Factura de Compras")
            oColumnCb.ValidValues.Add("19", "Abono de Compras")
            oColumnCb.ValidValues.Add("22", "Pedido de Compras")
            oColumnCb.ValidValues.Add("17", "Pedido de ventas")
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
            oColumnCb.Editable = False

            For i = 5 To 10
                If i <> 8 Then
                    CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                    oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                    If i <> 10 Then
                        oColumnTxt.Editable = False
                    End If
                End If
                If i = 5 Then
                    oColumnTxt.LinkedObjectType = "17"
                ElseIf i = 10 Then
                    'Comercial
                    oColumnTxt.ChooseFromListUID = "CFL_0"
                    oColumnTxt.ChooseFromListAlias = "SlpName"
                End If
            Next
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).AutoResizeColumns()
            For i = 11 To 21
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                Select Case i
                    Case 21 : oColumnTxt.Width = 300
                End Select

                If i = 11 Then
                    oColumnTxt.LinkedObjectType = "2"
                End If
                Select Case i
                    Case 16, 17 : oColumnTxt.RightJustified = True
                End Select
            Next
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub

    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        objGlobal.SBOApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Private Function CargaComboFormato(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboFormato = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            If objGlobal.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sSQL = "(Select '--' as ""Code"",' ' as ""Name"" FROM ""DUMMY"") "
                sSQL &= " UNION ALL "
                sSQL &= " (Select ""Code"",""Name"" FROM ""@EXO_CFCNF"" Order by ""Name"") "
            Else
                sSQL = "SELECT * FROM ( "
                sSQL &= " (Select ""Code"",""Name"" FROM ""@EXO_CFCNF"") "
                sSQL &= " UNION ALL "
                sSQL &= "(Select '--' as ""Code"",' ' as ""Name"" ) "
                sSQL &= " ) T  Order by ""Name"" "
            End If

            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboFormato = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
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
