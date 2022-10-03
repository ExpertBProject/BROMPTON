Imports SAPbouiCOM
Imports System.Xml
Imports EXO_UIAPI.EXO_UIAPI
Public Class INICIO
    Inherits EXO_UIAPI.EXO_DLLBase
    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        cargamenu()
        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim sXML As String = ""
            Dim res As String = ""
            Dim sSQL As String = ""

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_FCCNF.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_FCCNF", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CSAP.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UDO_EXO_CSAP", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOC.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_TMPDOC", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCL.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_TMPDOCL", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults

            sXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCLT.xml")
            objGlobal.SBOApp.StatusBar.SetText("Validando: UT_EXO_TMPDOCLT", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            objGlobal.refDi.comunes.LoadBDFromXML(sXML)
            res = objGlobal.SBOApp.GetLastBatchResults


            'Introducir los datos
            CargarDatos()
        End If
    End Sub
    Private Function CargarDatos() As Boolean
        CargarDatos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsCCC As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sPeriodo As String = ""
        Dim sFPoder As String = ""
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Try
            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_CSAP") 'UDO de Campos de SAP
#Region "CAMPOSSAP"
            sSQL = "SELECT * FROM ""@EXO_CSAP"" WHERE ""Code""='CAMPOSSAP' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sCode As String = oRs.Fields.Item("Code").Value.ToString
                oDI_COM.GetByKey(sCode)
                'Comprobamos que existan campos en la tabla de la cabecera
                sSQL = "SELECT * FROM ""@EXO_CSAPC"" WHERE ""Code""='CAMPOSSAP' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Cabecera Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposCabecera(oDI_COM, "CAMPOSSAP")
                End If
                'Comprobamos que existan campos en las líneas
                sSQL = "SELECT * FROM ""@EXO_CSAPL"" WHERE ""Code""='CAMPOSSAP' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Líneas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposLíneas(oDI_COM, "CAMPOSSAP")
                End If
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tablas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "CAMPOSSAP"
                oDI_COM.SetValue("CodEntry") = "99"
                oDI_COM.SetValue("Name") = "Campos de SAP"
                CrearCamposCabecera(oDI_COM, "CAMPOSSAP")
                CrearCamposLíneas(oDI_COM, "CAMPOSSAP")
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir campos SAP. " & oDI_COM.GetLastError)
                End If
            End If
#End Region
#Region "CAMPOSSAPTXT"
            sSQL = "SELECT * FROM ""@EXO_CSAP"" WHERE ""Code""='CAMPOSSAPTXT' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sCode As String = oRs.Fields.Item("Code").Value.ToString
                oDI_COM.GetByKey(sCode)
                'Comprobamos que existan campos en la tabla de la cabecera
                sSQL = "SELECT * FROM ""@EXO_CSAPC"" WHERE ""Code""='CAMPOSSAPTXT' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Cabecera Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposCabeceratxt(oDI_COM)
                End If
                'Comprobamos que existan campos en las líneas
                sSQL = "SELECT * FROM ""@EXO_CSAPL"" WHERE ""Code""='CAMPOSSAPTXT' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Líneas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposLíneastxt(oDI_COM, "")
                End If
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tablas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "CAMPOSSAPTXT"
                oDI_COM.SetValue("CodEntry") = "97"
                oDI_COM.SetValue("Name") = "Campos de SAP para Ficheros CSV"
                CrearCamposCabeceratxt(oDI_COM)
                CrearCamposLíneastxt(oDI_COM, "")
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir campos SAP. " & oDI_COM.GetLastError)
                End If
            End If
#End Region
#Region "CAMPOSSAPEXCEL"
            sSQL = "SELECT * FROM ""@EXO_CSAP"" WHERE ""Code""='CAMPOSSAPEXCEL' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sCode As String = oRs.Fields.Item("Code").Value.ToString
                oDI_COM.GetByKey(sCode)
                'Comprobamos que existan campos en la tabla de la cabecera
                sSQL = "SELECT * FROM ""@EXO_CSAPC"" WHERE ""Code""='CAMPOSSAPEXCEL' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Cabecera Campos SAP EXCEL...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposCabecera(oDI_COM, "CAMPOSSAPEXCEL")
                End If
                'Comprobamos que existan campos en las líneas
                sSQL = "SELECT * FROM ""@EXO_CSAPL"" WHERE ""Code""='CAMPOSSAPEXCEL' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Líneas Campos SAP EXCEL...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposLíneas(oDI_COM, "CAMPOSSAPEXCEL")
                End If
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tablas Campos SAP EXCEL...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "CAMPOSSAPEXCEL"
                oDI_COM.SetValue("CodEntry") = "98"
                oDI_COM.SetValue("Name") = "Campos de SAP EXCEL "
                CrearCamposCabecera(oDI_COM, "CAMPOSSAPEXCEL")
                CrearCamposLíneas(oDI_COM, "CAMPOSSAPEXCEL")
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir campos SAP EXCEL. " & oDI_COM.GetLastError)
                End If
            End If
#End Region

            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Tablas Campos SAP cargadas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            CargarDatos = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Function
    Private Function CrearCamposCabecera(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String)
        Try
            For i = 0 To 21
                oDI_COM.GetNewChild("EXO_CSAPC")
                Select Case i
                    Case 0
                        oDI_COM.SetValueChild("U_EXO_COD") = "ObjType"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Tipo Documento"
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 1
                        oDI_COM.SetValueChild("U_EXO_COD") = "CardCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Interlocutor"
                        If sCodigo = "CAMPOSSAPEXCEL" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 2
                        oDI_COM.SetValueChild("U_EXO_COD") = "CardName"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Nombre Interlocutor"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 3

                        oDI_COM.SetValueChild("U_EXO_COD") = "ADDID"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Externo Interlocutor"
                        If sCodigo = "CAMPOSSAPEXCEL" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 4
                        oDI_COM.SetValueChild("U_EXO_COD") = "NumAtCard"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Número de referencia"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 5
                        oDI_COM.SetValueChild("U_EXO_COD") = "Series"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Serie Factura"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 6
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocNum"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Nº de Documento"
                        If sCodigo = "CAMPOSSAPEXCEL" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 7
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocCurrency"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Moneda"
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 8
                        oDI_COM.SetValueChild("U_EXO_COD") = "SlpCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Empleado"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 9
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Contable"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 10
                        oDI_COM.SetValueChild("U_EXO_COD") = "TaxDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Documento"
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 11
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocDueDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Vencimiento"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 12
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TDTO"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Tipo Dto."
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 13
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_DTO"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descuento"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 14
                        oDI_COM.SetValueChild("U_EXO_COD") = "PeyMethod"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Vía de Pago"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 15
                        oDI_COM.SetValueChild("U_EXO_COD") = "GroupNum"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Condición de Pago"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 16
                        oDI_COM.SetValueChild("U_EXO_COD") = "PayToCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Dir. Facturación"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 17
                        oDI_COM.SetValueChild("U_EXO_COD") = "ShipToCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Dirección de entrega"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 18
                        oDI_COM.SetValueChild("U_EXO_COD") = "Comments"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Comentario"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 19
                        oDI_COM.SetValueChild("U_EXO_COD") = "OpeningRemarks"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Texto en Cabecera"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 20
                        oDI_COM.SetValueChild("U_EXO_COD") = "ClosingRemarks"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Texto en pie"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 21
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocType" ' I --> Artículos o S --> Servicios
                        oDI_COM.SetValueChild("U_EXO_DES") = "Tipo de Doc."
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Private Sub CrearCamposCabeceratxt(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity)
        Try
            For i = 0 To 3
                oDI_COM.GetNewChild("EXO_CSAPC")
                Select Case i
                    Case 0
                        oDI_COM.SetValueChild("U_EXO_COD") = "CardCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Interlocutor"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 1
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Contable"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 2
                        oDI_COM.SetValueChild("U_EXO_COD") = "TaxDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Documento"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 3
                        oDI_COM.SetValueChild("U_EXO_COD") = "NumAtCard"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Número de referencia"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Function CrearCamposLíneas(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String)
        Try
            For i = 0 To 10
                oDI_COM.GetNewChild("EXO_CSAPL")
                Select Case i
                    Case 0
                        oDI_COM.SetValueChild("U_EXO_COD") = "AcctCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Cta. Mayor"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 1
                        oDI_COM.SetValueChild("U_EXO_COD") = "ItemCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Artículo"

                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 2
                        oDI_COM.SetValueChild("U_EXO_COD") = "Dscription"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descripción Artículo"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 3
                        oDI_COM.SetValueChild("U_EXO_COD") = "Quantity"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Cantidad"

                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 4
                        oDI_COM.SetValueChild("U_EXO_COD") = "UnitPrice"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Precio Unidad"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 5
                        oDI_COM.SetValueChild("U_EXO_COD") = "DiscPrcnt"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descuento %"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 6
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_IMPSRV"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Importe Servicio"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 7
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TextoLin"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Texto Ampliado de la línea"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 8
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_IMP"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Impuesto"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 9
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_RET"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Retención"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 10
                        oDI_COM.SetValueChild("U_EXO_COD") = "GrossBuyPr"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Precio Bruto"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Private Sub CrearCamposLíneastxt(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String)
        Try
            For i = 0 To 3
                Select Case i
                    Case 0
                        oDI_COM.GetNewChild("EXO_CSAPL")
                        oDI_COM.SetValueChild("U_EXO_COD") = "ItemCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Artículo"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 1
                        oDI_COM.GetNewChild("EXO_CSAPL")
                        oDI_COM.SetValueChild("U_EXO_COD") = "Dscription"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descripción Artículo"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 2
                        oDI_COM.GetNewChild("EXO_CSAPL")
                        oDI_COM.SetValueChild("U_EXO_COD") = "Quantity"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Cantidad"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 3
                        If sCodigo = "CAMPOSSAPTXTENT" Then
                            oDI_COM.GetNewChild("EXO_CSAPL")
                            oDI_COM.SetValueChild("U_EXO_COD") = "U_EXO_TIPOE"
                            oDI_COM.SetValueChild("U_EXO_DES") = "TipoEntrega"
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        End If
                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Overrides Function filtros() As Global.SAPbouiCOM.EventFilters
        Dim fXML As String = ""
        Try
            fXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROC.xml")
            Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
            filtro.LoadFromXML(fXML)
            Return filtro
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return Nothing
        Finally

        End Try
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults

        If objGlobal.SBOApp.Menus.Exists("EXO-MnCDoc") = True Then
            Path = objGlobal.path & "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnCDOC.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnCDoc").Image = Path & "\MnCDOC.png"
                End If
            End If
        End If
    End Sub

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim Clase As Object = Nothing

        Try
            If infoEvento.BeforeAction = True Then
                Select Case infoEvento.MenuUID
                    Case ""
                End Select
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCFCF"
                        Clase = New EXO_FCCNF(objGlobal)
                        Return CType(Clase, EXO_FCCNF).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnCSAP"
                        Clase = New EXO_CFRP(objGlobal)
                        Return CType(Clase, EXO_CFRP).SBOApp_MenuEvent(infoEvento)
                    Case "EXO-MnCVPed"
                        Clase = New EXO_CVPED(objGlobal)
                        Return CType(Clase, EXO_CVPED).SBOApp_MenuEvent(infoEvento)
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Dim res As Boolean = True
        Dim Clase As Object = Nothing

        Try
            Select Case infoEvento.FormTypeEx
                Case "UDO_FT_EXO_FCCNF"
                    Clase = New EXO_FCCNF(objGlobal)
                    Return CType(Clase, EXO_FCCNF).SBOApp_ItemEvent(infoEvento)
                Case "UDO_FT_EXO_CSAP"
                    Clase = New EXO_CFRP(objGlobal)
                    Return CType(Clase, EXO_CFRP).SBOApp_ItemEvent(infoEvento)
                Case "EXO_CVPED"
                    Clase = New EXO_CVPED(objGlobal)
                    Return CType(Clase, EXO_CVPED).SBOApp_ItemEvent(infoEvento)
            End Select

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_TipoMensaje.Excepcion, EXO_TipoSalidaMensaje.MessageBox, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Finally
            Clase = Nothing
        End Try
    End Function
End Class
