' Copyright 2017 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector

Public Class SAPCostActivityPlanning
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon

    Sub New(aSapCon As SapCon)
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
        End Try
    End Sub

    Public Function ReadActivityOutput(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadActivityOutput = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTOUTPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPervalue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oPervalue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPervalue.Append()
                oPervalue.SetValue("VALUE_INDEX", lCnt)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadActivityOutput = "Success"
                For i As Integer = 0 To oPervalue.Count - 1
                    pData.Add(oPervalue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadActivityOutput = ReadActivityOutput & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityOutput = "Error: Exception in ReadActivityOutput"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadActivityOutputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection, pContrl As Collection) As String
        ReadActivityOutputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTOUTPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oContrl As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oContrl.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oIndexstructure.SetValue("ATTRIB_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oContrl.Append()
                oContrl.SetValue("ATTRIB_INDEX", lCnt)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadActivityOutputTot = "Success"
                For i As Integer = 0 To oTotValue.Count - 1
                    pData.Add(oTotValue(i))
                Next i
                For i As Integer = 0 To oContrl.Count - 1
                    pContrl.Add(oContrl(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadActivityOutputTot = ReadActivityOutputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityOutputTot = "Error: Exception in ReadActivityOutputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadActivityOutputTotS(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, ByRef pAOSets As Collection) As String
        ReadActivityOutputTotS = ""
        pAOSets = New Collection
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTOUTPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oContrl As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oContrl.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim lKey As String
            Dim lError As Boolean
            lCnt = 0
            For Each aObjRow In pObjects
                oIndexstructure.Clear()
                oCoobject.Clear()
                oTotValue.Clear()
                oContrl.Clear()
                oRETURN.Clear()
                lCnt = lCnt + 1
                lKey = aObjRow.Costcenter & "-" & aObjRow.Acttype
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oIndexstructure.SetValue("ATTRIB_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oContrl.Append()
                oContrl.SetValue("ATTRIB_INDEX", lCnt)
                ' call the BAPI
                oRfcFunction.Invoke(destination)
                If oRETURN.Count = 0 Then
                    If oTotValue.Count = 1 And oContrl.Count = 1 Then
                        Dim aAOSet As New AOSet
                        aAOSet.Key = aObjRow
                        aAOSet.Total = oTotValue(0)
                        aAOSet.Control = oContrl(0)
                        pAOSets.Add(aAOSet)
                    End If
                Else
                    For i As Integer = 0 To oRETURN.Count - 1
                        ReadActivityOutputTotS = ReadActivityOutputTotS & ";" & oRETURN(i).GetValue("MESSAGE")
                    Next i
                    lError = True
                End If
            Next aObjRow
            If lError = False Then
                ReadActivityOutputTotS = "Success"
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityOutputTotS = "Error: Exception in ReadActivityOutputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadPrimCost(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadPrimCost = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READPRIMCOST")
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

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                oPerValue.SetValue("COST_ELEM", lSAPFormat.unpack(aObjRow.Costelem, 10))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadPrimCost = "Success"
                For i As Integer = 0 To oPerValue.Count - 1
                    pData.Add(oPerValue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadPrimCost = ReadPrimCost & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadPrimCost = "Error: Exception in ReadPrimCost"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadPrimCostTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadPrimCostTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READPRIMCOST")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("COST_ELEM", lSAPFormat.unpack(aObjRow.Costelem, 10))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadPrimCostTot = "Success"
                For i As Integer = 0 To oTotValue.Count - 1
                    pData.Add(oTotValue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadPrimCostTot = ReadPrimCostTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadPrimCostTot = "Error: Exception in ReadPrimCostTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadActivityInput(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadActivityInput = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTINPUT")
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

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                oPerValue.SetValue("SEND_CCTR", lSAPFormat.unpack(aObjRow.SCostcenter, 10))
                oPerValue.SetValue("SEND_ACTIVITY", aObjRow.SActtype)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadActivityInput = "Success"
                For i As Integer = 0 To oPerValue.Count - 1
                    pData.Add(oPerValue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadActivityInput = ReadActivityInput & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityInput = "Error: Exception in ReadActivityInput"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadActivityInputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection) As String
        ReadActivityInputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READACTINPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("SEND_CCTR", lSAPFormat.unpack(aObjRow.SCostcenter, 10))
                oTotValue.SetValue("SEND_ACTIVITY", aObjRow.SActtype)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadActivityInputTot = "Success"
                For i As Integer = 0 To oTotValue.Count - 1
                    pData.Add(oTotValue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadActivityInputTot = ReadActivityInputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadActivityInputTot = "Error: Exception in ReadActivityInputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostPrimCost(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection,
                             Optional pDelta As String = " ") As String
        PostPrimCost = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTPRIMCOST")
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
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                oPerValue.SetValue("COST_ELEM", lSAPFormat.unpack(aObjRow.Costelem, 10))
                '   move the values from the data
                aDataRow = pData(lCnt)
                Dim J As Int32
                Dim aDbl As Double
                For J = 8 To 71
                    aDbl = CDbl(aDataRow(J - 7))
                    oPerValue.SetValue(J - 1, aDbl)  'Array start at 0, so 1 less then in SAP dictionary
                Next J
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostPrimCost = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostPrimCost = PostPrimCost & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostPrimCost = "Error: Exception in PostPrimCostTot"
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

    Public Function PostPrimCostTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection,
                             Optional pDelta As String = " ") As String
        PostPrimCostTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTPRIMCOST")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)
            oRfcFunction.SetValue("DELTA", pDelta)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("COST_ELEM", lSAPFormat.unpack(aObjRow.Costelem, 10))
                '   move the values from the data
                aDataRow = pData(lCnt)
                oTotValue.SetValue("FIX_VALUE", CDbl(aDataRow(1)))
                oTotValue.SetValue("DIST_KEY_FIX_VAL", CStr(aDataRow(2)))
                oTotValue.SetValue("VAR_VALUE", CDbl(aDataRow(3)))
                oTotValue.SetValue("DIST_KEY_VAR_VAL", CStr(aDataRow(4)))
                If CStr(aDataRow(6)) <> "" Then
                    oTotValue.SetValue("FIX_QUAN", aDataRow(5))
                    oTotValue.SetValue("DIST_KEY_FIX_QUAN", CStr(aDataRow(6)))
                End If
                If CStr(aDataRow(8)) <> "" Then
                    oTotValue.SetValue("VAR_QUAN", aDataRow(7))
                    oTotValue.SetValue("DIST_KEY_VAR_QUAN", CStr(aDataRow(8)))
                End If
                If CStr(aDataRow(9)) <> "" Then
                    oTotValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(9)))
                End If
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostPrimCostTot = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostPrimCostTot = PostPrimCostTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostPrimCostTot = "Error: Exception in PostPrimCostTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityOutput(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection,
                             Optional pDelta As String = " ") As String
        Dim J As Int32
        Dim aDbl As Double
        Dim aInt As Int16
        PostActivityOutput = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTOUTPUT")
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
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                '   move the values from the data
                aDataRow = pData(lCnt)

                For J = 2 To 65
                    aDbl = CDbl(aDataRow(J - 1))
                    oPerValue.SetValue(J - 1, aDbl)        'Array start at 0, so 1 less then in SAP dictionary
                Next J
                For J = 66 To 97
                    aInt = CInt(aDataRow(J - 1))
                    oPerValue.SetValue(J - 1, aInt)        'Array start at 0, so 1 less then in SAP dictionary
                Next J
                oPerValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(97)))
                oPerValue.SetValue("CURRENCY", CStr(aDataRow(98)))
            Next aObjRow

            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostActivityOutput = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostActivityOutput = PostActivityOutput & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityOutput = "Error: Exception in PostActivityOutput"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityOutputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection, pContrl As Collection,
                             Optional pDelta As String = " ", Optional pAOCtrl As String = "") As String
        PostActivityOutputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTOUTPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oContrl As IRfcTable = oRfcFunction.GetTable("CONTRL")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oContrl.Clear()
            oRETURN.Clear()

            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)
            oRfcFunction.SetValue("DELTA", pDelta)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            Dim aCtrlRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                If pAOCtrl = "" Then
                    oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                End If
                oIndexstructure.SetValue("ATTRIB_INDEX", lCnt)
                If pAOCtrl = "" Then
                    oTotValue.Append()
                    oTotValue.SetValue("VALUE_INDEX", lCnt)
                End If
                '   move the values from the data
                aDataRow = pData(lCnt)
                If pAOCtrl = "" Then
                    oTotValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(1)))
                    oTotValue.SetValue("CURRENCY", CStr(aDataRow(2)))
                    oTotValue.SetValue("ACTVTY_QTY", CDbl(aDataRow(3)))
                    oTotValue.SetValue("DIST_KEY_QUAN", CStr(aDataRow(4)))
                    oTotValue.SetValue("ACTVTY_CAPACTY", CDbl(aDataRow(5)))
                    oTotValue.SetValue("DIST_KEY_CAPCTY", CStr(aDataRow(6)))
                    oTotValue.SetValue("PRICE_FIX", CDbl(aDataRow(7)))
                    oTotValue.SetValue("DIST_KEY_PRICE_FIX", CStr(aDataRow(8)))
                    oTotValue.SetValue("PRICE_VAR", CDbl(aDataRow(9)))
                    oTotValue.SetValue("DIST_KEY_PRICE_VAR", CStr(aDataRow(10)))
                    oTotValue.SetValue("PRICE_UNIT", CInt(aDataRow(11)))
                    oTotValue.SetValue("EQUIVALENCE_NO", CInt(aDataRow(12)))
                End If
                '   move the values from the contrl
                aCtrlRow = pContrl(lCnt)
                oContrl.Append()
                oContrl.SetValue("ATTRIB_INDEX", lCnt)
                oContrl.SetValue("PRICE_INDICATOR", CStr(lSAPFormat.unpack(aCtrlRow(1), 3)))
                oContrl.SetValue("SWITCH_LAYOUT", CStr(aCtrlRow(2)))
                oContrl.SetValue("ALLOC_COST_ELEM", CStr(lSAPFormat.unpack(aCtrlRow(3), 10)))
                oContrl.SetValue("ACT_PRICE_IND", CStr(lSAPFormat.unpack(aCtrlRow(4), 3)))
                oContrl.SetValue("PREDIS_FXD_COST", CStr(aCtrlRow(5)))
                oContrl.SetValue("ACT_CAT_ACTUAL", CStr(aCtrlRow(6)))
                oContrl.SetValue("AVERAGE_PRICE_IND", CStr(aCtrlRow(7)))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostActivityOutputTot = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostActivityOutputTot = PostActivityOutputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityOutputTot = "Error: Exception in PostActivityOutputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostActivityInput(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection,
                             Optional pDelta As String = " ") As String
        PostActivityInput = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTINPUT")
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
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPerValue.Append()
                oPerValue.SetValue("VALUE_INDEX", lCnt)
                oPerValue.SetValue("SEND_CCTR", lSAPFormat.unpack(aObjRow.SCostcenter, 10))
                oPerValue.SetValue("SEND_ACTIVITY", aObjRow.SActtype)
                '   move the values from the data
                aDataRow = pData(lCnt)
                oPerValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(1)))
                Dim J As Int32
                Dim aDbl As Double
                For J = 6 To 37
                    aDbl = CDbl(aDataRow(J - 4))
                    oPerValue.SetValue(J - 1, aDbl)  'Array start at 0, so 1 less then in SAP dictionary
                Next J
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostActivityInput = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostActivityInput = PostActivityInput & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityInput = "Error: Exception in PostActivityInput"
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

    Public Function PostActivityInputTot(pCoAre As String, pFiscy As String, pPfrom As String,
                             pPto As String, pVers As String, pCurt As String,
                             pObjects As Collection, pData As Collection,
                             Optional pDelta As String = " ") As String
        PostActivityInputTot = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTACTINPUT")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oTotValue As IRfcTable = oRfcFunction.GetTable("TOTVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oTotValue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oHeaderinfo.SetValue("PLAN_CURRTYPE", pCurt)
            oRfcFunction.SetValue("DELTA", pDelta)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oTotValue.Append()
                oTotValue.SetValue("VALUE_INDEX", lCnt)
                oTotValue.SetValue("SEND_CCTR", lSAPFormat.unpack(aObjRow.SCostcenter, 10))
                oTotValue.SetValue("SEND_ACTIVITY", aObjRow.SActtype)
                '   move the values from the data
                aDataRow = pData(lCnt)
                oTotValue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(1)))
                oTotValue.SetValue("QUANTITY_FIX", CDbl(aDataRow(2)))
                oTotValue.SetValue("DIST_KEY_QUAN_FIX", CStr(aDataRow(3)))
                oTotValue.SetValue("QUANTITY_VAR", CDbl(aDataRow(4)))
                oTotValue.SetValue("DIST_KEY_QUAN_VAR", CStr(aDataRow(5)))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostActivityInputTot = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostActivityInputTot = PostActivityInputTot & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostActivityInputTot = "Error: Exception in PostActivityInputTot"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function PostKeyFigure(pCoAre As String, pFiscy As String, pPfrom As String,
                            pPto As String, pVers As String, pCurt As String,
                            pObjects As Collection, pData As Collection,
                            Optional pDelta As String = " ") As String
        PostKeyFigure = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_POSTKEYFIGURE")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPervalue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oPervalue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))
            oRfcFunction.SetValue("DELTA", pDelta)

            Dim lCnt As Integer
            Dim aObjRow As Object
            Dim aDataRow As Collection
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPervalue.Append()
                oPervalue.SetValue("VALUE_INDEX", lCnt)
                '   move the values from the data
                aDataRow = pData(lCnt)
                oPervalue.SetValue("STATKEYFIG", lSAPFormat.unpack(aObjRow.STATKEYFIG, 6))
                oPervalue.SetValue("UNIT_OF_MEASURE", CStr(aDataRow(1)))
                oPervalue.SetValue("QUANTITY_PER01", CDbl(aDataRow(2)))
                oPervalue.SetValue("QUANTITY_PER02", CDbl(aDataRow(3)))
                oPervalue.SetValue("QUANTITY_PER03", CDbl(aDataRow(4)))
                oPervalue.SetValue("QUANTITY_PER04", CDbl(aDataRow(5)))
                oPervalue.SetValue("QUANTITY_PER05", CDbl(aDataRow(6)))
                oPervalue.SetValue("QUANTITY_PER06", CDbl(aDataRow(7)))
                oPervalue.SetValue("QUANTITY_PER07", CDbl(aDataRow(8)))
                oPervalue.SetValue("QUANTITY_PER08", CDbl(aDataRow(9)))
                oPervalue.SetValue("QUANTITY_PER09", CDbl(aDataRow(10)))
                oPervalue.SetValue("QUANTITY_PER10", CDbl(aDataRow(11)))
                oPervalue.SetValue("QUANTITY_PER11", CDbl(aDataRow(12)))
                oPervalue.SetValue("QUANTITY_PER12", CDbl(aDataRow(13)))
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                PostKeyFigure = "Success"
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit()
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    PostKeyFigure = PostKeyFigure & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            PostKeyFigure = "Error: Exception in PostKeyFigure"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function ReadKeyFigure(pCoAre As String, pFiscy As String, pPfrom As String,
                                    pPto As String, pVers As String, pCurt As String,
                                    pObjects As Collection, pData As Collection) As String
        ReadKeyFigure = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("BAPI_COSTACTPLN_READKEYFIGURE")
            RfcSessionManager.BeginContext(destination)
            Dim lSAPFormat As New SAPFormat
            Dim oHeaderinfo As IRfcStructure = oRfcFunction.GetStructure("HEADERINFO")
            Dim oIndexstructure As IRfcTable = oRfcFunction.GetTable("INDEXSTRUCTURE")
            Dim oCoobject As IRfcTable = oRfcFunction.GetTable("COOBJECT")
            Dim oPervalue As IRfcTable = oRfcFunction.GetTable("PERVALUE")
            Dim oRETURN As IRfcTable = oRfcFunction.GetTable("RETURN")
            oIndexstructure.Clear()
            oCoobject.Clear()
            oPervalue.Clear()
            oRETURN.Clear()
            oHeaderinfo.SetValue("CO_AREA", pCoAre)
            oHeaderinfo.SetValue("FISC_YEAR", pFiscy)
            oHeaderinfo.SetValue("PERIOD_FROM", lSAPFormat.unpack(pPfrom, 3))
            oHeaderinfo.SetValue("PERIOD_TO", lSAPFormat.unpack(pPto, 3))
            oHeaderinfo.SetValue("VERSION", lSAPFormat.unpack(pVers, 3))

            Dim lCnt As Integer
            Dim aObjRow As Object
            lCnt = 0
            For Each aObjRow In pObjects
                lCnt = lCnt + 1
                oCoobject.Append()
                oCoobject.SetValue("OBJECT_INDEX", lCnt)
                If aObjRow.Costcenter <> "" Then
                    oCoobject.SetValue("COSTCENTER", lSAPFormat.unpack(aObjRow.Costcenter, 10))
                End If
                If aObjRow.WBS_ELEMENT <> "" Then
                    oCoobject.SetValue("WBS_ELEMENT", aObjRow.WBS_ELEMENT)
                End If
                If aObjRow.Acttype <> "" Then
                    oCoobject.SetValue("ACTTYPE", aObjRow.Acttype)
                    '      oCoobject.SetValue("ACTTYPE", lSAPFormat.unpack(aObjRow.Acttype, 6))
                End If
                oIndexstructure.Append()
                oIndexstructure.SetValue("OBJECT_INDEX", lCnt)
                oIndexstructure.SetValue("VALUE_INDEX", lCnt)
                oPervalue.Append()
                oPervalue.SetValue("VALUE_INDEX", lCnt)
                oPervalue.SetValue("STATKEYFIG", aObjRow.STATKEYFIG)
            Next aObjRow
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            If oRETURN.Count = 0 Then
                ReadKeyFigure = "Success"
                For i As Integer = 0 To oPervalue.Count - 1
                    pData.Add(oPervalue(i))
                Next i
            Else
                For i As Integer = 0 To oRETURN.Count - 1
                    ReadKeyFigure = ReadKeyFigure & ";" & oRETURN(i).GetValue("MESSAGE")
                Next i
            End If
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPCostActivityPlanning")
            ReadKeyFigure = "Error: Exception in ReadKeyFigure"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class

