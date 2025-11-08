' Copyright 2025 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPCostActivityPlanning
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private par As SAPCommon.TStr
    Private cName As String = "SAPCostActivityPlanning"

    Sub New(aSapCon As SapCon, ByRef aPar As SAPCommon.TStr)
        Try
            log.Debug("New - " & "checking connection")
            sapcon = aSapCon
            par = aPar
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta_PostPrimCost(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"HEADERINFO"}
        Dim aImports As String() = {"DELTA"}
        Dim aTables As String() = {"INDEXSTRUCTURE", "COOBJECT", "PERVALUE", "TOTVALUE", "CONTRL"}
        Try
            log.Debug("getMeta_PostPrimCost - " & "creating Function BAPI_COSTACTPLN_POSTPRIMCOST")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTPRIMCOST")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_PostPrimCost - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        Finally
            log.Debug("getMeta_PostPrimCost - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_PostActivityOutput(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"HEADERINFO"}
        Dim aImports As String() = {"DELTA", "PRICE_QUANT_PLAN"}
        Dim aTables As String() = {"INDEXSTRUCTURE", "COOBJECT", "PERVALUE", "TOTVALUE", "CONTRL"}
        Try
            log.Debug("getMeta_PostActivityOutput - " & "creating Function BAPI_COSTACTPLN_POSTACTOUTPUT")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTOUTPUT")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_PostActivityOutput - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        Finally
            log.Debug("getMeta_PostActivityOutput - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_PostActivityInput(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"HEADERINFO"}
        Dim aImports As String() = {"DELTA"}
        Dim aTables As String() = {"INDEXSTRUCTURE", "COOBJECT", "PERVALUE", "TOTVALUE"}
        Try
            log.Debug("getMeta_PostActivityInput - " & "creating Function BAPI_COSTACTPLN_POSTACTINPUT")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTINPUT")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_PostActivityInput - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        Finally
            log.Debug("getMeta_PostActivityInput - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Sub getMeta_PostKeyFigure(ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Dim aStructures As String() = {"HEADERINFO"}
        Dim aImports As String() = {"DELTA"}
        Dim aTables As String() = {"INDEXSTRUCTURE", "COOBJECT", "PERVALUE", "TOTVALUE"}
        Try
            log.Debug("getMeta_PostKeyFigure - " & "creating Function BAPI_COSTACTPLN_POSTKEYFIGURE")
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTKEYFIGURE")
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To aImports.Length - 1
                addToFieldDic("I|" & aImports(s), oRfcFunction.Metadata.Item(aImports(s)), pFieldDic)
            Next
            ' Import Strcutures
            For s As Integer = 0 To aStructures.Length - 1
                oStructure = oRfcFunction.GetStructure(aStructures(s))
                addToStrucDic("S|" & aStructures(s), oStructure.Metadata, pStrucDic)
            Next
            For s As Integer = 0 To aTables.Length - 1
                oTable = oRfcFunction.GetTable(aTables(s))
                addToStrucDic("T|" & aTables(s), oTable.Metadata.LineType, pStrucDic)
            Next
        Catch Ex As System.Exception
            log.Error("getMeta_PostKeyFigure - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        Finally
            log.Debug("getMeta_PostKeyFigure - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

    Public Function PostPrimCost(pData As TSAP_Data_CoPl, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        PostPrimCost = ""
        Try
            If pCheck Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_CHECKPRIMCOST")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTPRIMCOST")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim oINDEXSTRUCTURE As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCOOBJECT As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPERVALUE As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oTOTVALUE As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oCONTRL As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oINDEXSTRUCTURE.Clear()
            oCOOBJECT.Clear()
            oPERVALUE.Clear()
            oTOTVALUE.Clear()
            oCONTRL.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="INDEXSTRUCTURE", pIRfcTable:=oINDEXSTRUCTURE)
            pData.aDataDic.to_IRfcTable(pKey:="COOBJECT", pIRfcTable:=oCOOBJECT)
            pData.aDataDic.to_IRfcTable(pKey:="PERVALUE", pIRfcTable:=oPERVALUE)
            pData.aDataDic.to_IRfcTable(pKey:="TOTVALUE", pIRfcTable:=oTOTVALUE)
            pData.aDataDic.to_IRfcTable(pKey:="CONTRL", pIRfcTable:=oCONTRL)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    PostPrimCost = PostPrimCost & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            PostPrimCost = If(PostPrimCost = "", pOKMsg, If(aErr = False, pOKMsg & PostPrimCost, "Error" & PostPrimCost))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
            PostPrimCost = "Error: Exception in PostPrimCost"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityOutput(pData As TSAP_Data_CoPl, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        PostActivityOutput = ""
        Try
            If pCheck Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_CHECKACTOUTPUT")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTOUTPUT")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim oINDEXSTRUCTURE As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCOOBJECT As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPERVALUE As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oTOTVALUE As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oCONTRL As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oINDEXSTRUCTURE.Clear()
            oCOOBJECT.Clear()
            oPERVALUE.Clear()
            oTOTVALUE.Clear()
            oCONTRL.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="INDEXSTRUCTURE", pIRfcTable:=oINDEXSTRUCTURE)
            pData.aDataDic.to_IRfcTable(pKey:="COOBJECT", pIRfcTable:=oCOOBJECT)
            pData.aDataDic.to_IRfcTable(pKey:="PERVALUE", pIRfcTable:=oPERVALUE)
            pData.aDataDic.to_IRfcTable(pKey:="TOTVALUE", pIRfcTable:=oTOTVALUE)
            pData.aDataDic.to_IRfcTable(pKey:="CONTRL", pIRfcTable:=oCONTRL)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    PostActivityOutput = PostActivityOutput & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            PostActivityOutput = If(PostActivityOutput = "", pOKMsg, If(aErr = False, pOKMsg & PostActivityOutput, "Error" & PostActivityOutput))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
            PostActivityOutput = "Error: Exception in PostActivityOutput"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityInput(pData As TSAP_Data_CoPl, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        PostActivityInput = ""
        Try
            If pCheck Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_CHECKACTINPUT")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTINPUT")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim oINDEXSTRUCTURE As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCOOBJECT As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPERVALUE As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oTOTVALUE As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oINDEXSTRUCTURE.Clear()
            oCOOBJECT.Clear()
            oPERVALUE.Clear()
            oTOTVALUE.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="INDEXSTRUCTURE", pIRfcTable:=oINDEXSTRUCTURE)
            pData.aDataDic.to_IRfcTable(pKey:="COOBJECT", pIRfcTable:=oCOOBJECT)
            pData.aDataDic.to_IRfcTable(pKey:="PERVALUE", pIRfcTable:=oPERVALUE)
            pData.aDataDic.to_IRfcTable(pKey:="TOTVALUE", pIRfcTable:=oTOTVALUE)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    PostActivityInput = PostActivityInput & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            PostActivityInput = If(PostActivityInput = "", pOKMsg, If(aErr = False, pOKMsg & PostActivityInput, "Error" & PostActivityInput))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
            PostActivityInput = "Error: Exception in PostActivityInput"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostKeyFigure(pData As TSAP_Data_CoPl, Optional pOKMsg As String = "OK", Optional pCheck As Boolean = False) As String
        PostKeyFigure = ""
        Try
            If pCheck Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_CHECKKEYFIGURE")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTKEYFIGURE")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim oINDEXSTRUCTURE As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCOOBJECT As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPERVALUE As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oTOTVALUE As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oINDEXSTRUCTURE.Clear()
            oCOOBJECT.Clear()
            oPERVALUE.Clear()
            oTOTVALUE.Clear()
            oRETURN.Clear()

            Dim aTStrRec As SAPCommon.TStrRec
            Dim oStruc As IRfcStructure
            ' set the header values
            For Each aTStrRec In pData.aHdrRec.aTDataRecCol
                If aTStrRec.Strucname <> "" Then
                    oStruc = oRfcFunction.GetStructure(aTStrRec.Strucname)
                    oStruc.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                Else
                    oRfcFunction.SetValue(aTStrRec.Fieldname, aTStrRec.formated)
                End If
            Next
            ' set the table fields
            pData.aDataDic.to_IRfcTable(pKey:="INDEXSTRUCTURE", pIRfcTable:=oINDEXSTRUCTURE)
            pData.aDataDic.to_IRfcTable(pKey:="COOBJECT", pIRfcTable:=oCOOBJECT)
            pData.aDataDic.to_IRfcTable(pKey:="PERVALUE", pIRfcTable:=oPERVALUE)
            pData.aDataDic.to_IRfcTable(pKey:="TOTVALUE", pIRfcTable:=oTOTVALUE)
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean = False
            For i As Integer = 0 To oRETURN.Count - 1
                If oRETURN(i).GetValue("TYPE") <> "I" And oRETURN(i).GetValue("TYPE") <> "W" Then
                    PostKeyFigure = PostKeyFigure & ";" & oRETURN(i).GetValue("MESSAGE")
                    If oRETURN(i).GetValue("TYPE") <> "S" And oRETURN(i).GetValue("TYPE") <> "W" Then
                        aErr = True
                    End If
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            PostKeyFigure = If(PostKeyFigure = "", pOKMsg, If(aErr = False, pOKMsg & PostKeyFigure, "Error" & PostKeyFigure))
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
            PostKeyFigure = "Error: Exception in PostKeyFigure"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostPrimCostDyn(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pData As Collection,
                             Optional pDelta As String = " ", Optional pTest As Boolean = False) As Collection
        Dim aRet As New Collection
        Try
            If pTest Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_CHECKPRIMCOST")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTPRIMCOST")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPerValue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oPerValue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)
            oRfcFunction.SetValue("DELTA", pDelta)

            Dim lCnt As Integer
            Dim aTPlanRec As New TPlanRec
            Dim aTStrRec As TStrRec

            lCnt = 0
            For Each aTPlanRec In pData
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                For Each aTStrRec In aTPlanRec.aTPlanRecCol
                    Select Case aTStrRec.STRUCNAME.Value
                        Case "COOBJECT"
                            oCoobject.SetValue(aTStrRec.FIELDNAME.Value, aTStrRec.formated())
                        Case "PERVALUE"
                            oPerValue.SetValue(aTStrRec.FIELDNAME.Value, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                aRet.Add("Success")
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    aRet.Add(oRETURN(i).GetValue("MESSAGE"))
                Next i
            End If
            PostPrimCostDyn = aRet
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            aRet.Add("Error: Exception in PostPrimCostTot")
            PostPrimCostDyn = aRet
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityInputDyn(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pData As Collection,
                             Optional pDelta As String = " ", Optional pTest As Boolean = False) As Collection
        Dim aRet As New Collection
        Try
            If pTest Then
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_CHECKACTINPUT")
            Else
                oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTINPUT")
            End If
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPerValue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oPerValue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)
            oRfcFunction.SetValue("DELTA", pDelta)

            Dim lCnt As Integer
            Dim aTPlanRec As New TPlanRec
            Dim aTStrRec As TStrRec
            lCnt = 0
            For Each aTPlanRec In pData
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                For Each aTStrRec In aTPlanRec.aTPlanRecCol
                    Select Case aTStrRec.STRUCNAME.Value
                        Case "COOBJECT"
                            oCoobject.SetValue(aTStrRec.FIELDNAME.Value, aTStrRec.formated())
                        Case "PERVALUE"
                            oPerValue.SetValue(aTStrRec.FIELDNAME.Value, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                aRet.Add("Success")
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    aRet.Add(oRETURN(i).GetValue("MESSAGE"))
                Next i
            End If
            PostActivityInputDyn = aRet
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityInputDyn = aRet
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class

