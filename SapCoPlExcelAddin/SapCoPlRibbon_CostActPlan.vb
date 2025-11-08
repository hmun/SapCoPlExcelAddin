' Copyright 2025 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports System.Windows.Forms

Public Class SapCoPlRibbon_CostActPlan

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapCoPlRibbon_CostActPlan getGenParametrs - " & "reading Parameter")
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SapCoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCoPl")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SapCoPlCostActPlan"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SapCoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCoPl")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop
        getGenParameters = True
    End Function

    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SapCoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SapCoPl")
            getIntParameters = False
            Exit Function
        End Try
        i = 2
        pIntPar = New SAPCommon.TStr
        Do
            pIntPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        ' no obligatory parameters check - we should know what we are doing
        getIntParameters = True
    End Function

    Public Sub postPC(ByRef pSapCon As SapCon, Optional pCheck As Boolean = False)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aType As String = "PC"

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aCoPlLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "DATA") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "DATA")), 4)
        Dim aHdrLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "HEAD") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "HEAD")), aCoPlLOff - 4)
        Dim aCoPlWsName As String = If(aIntPar.value(aType & "_" & "WS", "DATA") <> "", aIntPar.value(aType & "_" & "WS", "DATA"), "Data")
        Dim aCoPlWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAMSG") <> "", aIntPar.value(aType & "_" & "COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aPostClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAPOST") <> "", aIntPar.value(aType & "_" & "COL", "DATAPOST"), "INT-POST")
        Dim aPostClmnNr As Integer = 0
        Dim aCoPlClmnNr As Integer = If(aIntPar.value(aType & "_" & "COLNR", "DATACHECK") <> "", CInt(aIntPar.value(aType & "_" & "COLNR", "DATACHECK")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value(aType & "_" & "RET", "OKMSG") <> "", aIntPar.value(aType & "_" & "RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aCoPlWs = aWB.Worksheets(aCoPlWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aCoPlWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl")
            Exit Sub
        End Try
        parseHeaderLine(aCoPlWs, jMax, pMsgClmn:=aMsgClmn, pMsgClmnNr:=aMsgClmnNr, pPostClmn:=aPostClmn, pPostClmnNr:=aPostClmnNr, pHdrLine:=aHdrLOff + 1)
        Try
            log.Debug("SapCoPlRibbon_CostActPlan.postPC - " & "processing data - disabling events, screen update, cursor")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.EnableEvents = False
            '            Globals.ThisAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aCoPlLOff + 1
            Dim aKey As String
            Dim aPost As String = ""
            Dim aItems As New TData(aIntPar)
            Dim aTSAP_Data_CoPl As New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postPC")
            Do
                If Left(CStr(aCoPlWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        aPost = If(String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, aPostClmnNr).value)), "", CStr(aCoPlWs.Cells(i, aPostClmnNr).value))
                    End If
                    aKey = CStr(i)
                    ' read DATA
                    aItems.ws_parse_line_simple(aCoPlWs, aCoPlLOff, i, jMax, pHdrLine:=aHdrLOff + 1)
                    If String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i + 1, aCoPlClmnNr).value)) Or aPost.ToUpper = "X" Then
                        If aTSAP_Data_CoPl.fillHeader(aItems) And aTSAP_Data_CoPl.fillData(aItems) Then
                            log.Debug("SapCoPlRibbon_CostActPlan.postPC - " & "calling aSAPHrMaster.PostPrimCost")
                            Globals.ThisAddIn.Application.StatusBar = "Calling SAP-BAPI at line " & i
                            aRetStr = aSAPCostActivityPlanning.PostPrimCost(aTSAP_Data_CoPl, pOKMsg:=aOKMsg, pCheck:=pCheck)
                            log.Debug("SapCoPlRibbon_CostActPlan.postPC - " & "aSAPHrMaster.PostPrimCost returned, aRetStr=" & aRetStr)
                            For Each aKey In aItems.aTDataDic.Keys
                                aCoPlWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        End If
                        aItems = New TData(aIntPar)
                        aTSAP_Data_CoPl = New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postPC")
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, 1).value))
            log.Debug("SapCoPlRibbon_CostActPlan.postPC - " & "all data processed - enabling events, screen update, cursor")
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoPlRibbon_CostActPlan.postPC failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl AddIn")
            log.Error("SapCoPlRibbon_CostActPlan.postPC - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub postAO(ByRef pSapCon As SapCon, Optional pCheck As Boolean = False)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aType As String = "AO"

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aCoPlLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "DATA") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "DATA")), 4)
        Dim aHdrLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "HEAD") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "HEAD")), aCoPlLOff - 4)
        Dim aCoPlWsName As String = If(aIntPar.value(aType & "_" & "WS", "DATA") <> "", aIntPar.value(aType & "_" & "WS", "DATA"), "Data")
        Dim aCoPlWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAMSG") <> "", aIntPar.value(aType & "_" & "COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aPostClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAPOST") <> "", aIntPar.value(aType & "_" & "COL", "DATAPOST"), "INT-POST")
        Dim aPostClmnNr As Integer = 0
        Dim aCoPlClmnNr As Integer = If(aIntPar.value(aType & "_" & "COLNR", "DATACHECK") <> "", CInt(aIntPar.value(aType & "_" & "COLNR", "DATACHECK")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value(aType & "_" & "RET", "OKMSG") <> "", aIntPar.value(aType & "_" & "RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aCoPlWs = aWB.Worksheets(aCoPlWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aCoPlWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl")
            Exit Sub
        End Try
        parseHeaderLine(aCoPlWs, jMax, pMsgClmn:=aMsgClmn, pMsgClmnNr:=aMsgClmnNr, pPostClmn:=aPostClmn, pPostClmnNr:=aPostClmnNr, pHdrLine:=aHdrLOff + 1)
        Try
            log.Debug("SapCoPlRibbon_CostActPlan.postAO - " & "processing data - disabling events, screen update, cursor")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.EnableEvents = False
            '            Globals.ThisAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aCoPlLOff + 1
            Dim aKey As String
            Dim aPost As String = ""
            Dim aItems As New TData(aIntPar)
            Dim aTSAP_Data_CoPl As New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postAO")
            Do
                If Left(CStr(aCoPlWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        aPost = If(String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, aPostClmnNr).value)), "", CStr(aCoPlWs.Cells(i, aPostClmnNr).value))
                    End If
                    aKey = CStr(i)
                    ' read DATA
                    aItems.ws_parse_line_simple(aCoPlWs, aCoPlLOff, i, jMax, pHdrLine:=aHdrLOff + 1)
                    If String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i + 1, aCoPlClmnNr).value)) Or aPost.ToUpper = "X" Then
                        If aTSAP_Data_CoPl.fillHeader(aItems) And aTSAP_Data_CoPl.fillData(aItems) Then
                            log.Debug("SapCoPlRibbon_CostActPlan.postAO - " & "calling aSAPHrMaster.PostActivityOutput")
                            Globals.ThisAddIn.Application.StatusBar = "Calling SAP-BAPI at line " & i
                            aRetStr = aSAPCostActivityPlanning.PostActivityOutput(aTSAP_Data_CoPl, pOKMsg:=aOKMsg, pCheck:=pCheck)
                            log.Debug("SapCoPlRibbon_CostActPlan.postAO - " & "aSAPHrMaster.PostActivityOutput returned, aRetStr=" & aRetStr)
                            For Each aKey In aItems.aTDataDic.Keys
                                aCoPlWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        End If
                        aItems = New TData(aIntPar)
                        aTSAP_Data_CoPl = New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postAO")
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, 1).value))
            log.Debug("SapCoPlRibbon_CostActPlan.postAO - " & "all data processed - enabling events, screen update, cursor")
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoPlRibbon_CostActPlan.postAO failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl AddIn")
            log.Error("SapCoPlRibbon_CostActPlan.postAO - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub postSK(ByRef pSapCon As SapCon, Optional pCheck As Boolean = False)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aType As String = "SK"

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aCoPlLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "DATA") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "DATA")), 4)
        Dim aHdrLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "HEAD") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "HEAD")), aCoPlLOff - 4)
        Dim aCoPlWsName As String = If(aIntPar.value(aType & "_" & "WS", "DATA") <> "", aIntPar.value(aType & "_" & "WS", "DATA"), "Data")
        Dim aCoPlWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAMSG") <> "", aIntPar.value(aType & "_" & "COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aPostClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAPOST") <> "", aIntPar.value(aType & "_" & "COL", "DATAPOST"), "INT-POST")
        Dim aPostClmnNr As Integer = 0
        Dim aCoPlClmnNr As Integer = If(aIntPar.value(aType & "_" & "COLNR", "DATACHECK") <> "", CInt(aIntPar.value(aType & "_" & "COLNR", "DATACHECK")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value(aType & "_" & "RET", "OKMSG") <> "", aIntPar.value(aType & "_" & "RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aCoPlWs = aWB.Worksheets(aCoPlWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aCoPlWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl")
            Exit Sub
        End Try
        parseHeaderLine(aCoPlWs, jMax, pMsgClmn:=aMsgClmn, pMsgClmnNr:=aMsgClmnNr, pPostClmn:=aPostClmn, pPostClmnNr:=aPostClmnNr, pHdrLine:=aHdrLOff + 1)
        Try
            log.Debug("SapCoPlRibbon_CostActPlan.postSK - " & "processing data - disabling events, screen update, cursor")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.EnableEvents = False
            '            Globals.ThisAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aCoPlLOff + 1
            Dim aKey As String
            Dim aPost As String = ""
            Dim aItems As New TData(aIntPar)
            Dim aTSAP_Data_CoPl As New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postSK")
            Do
                If Left(CStr(aCoPlWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        aPost = If(String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, aPostClmnNr).value)), "", CStr(aCoPlWs.Cells(i, aPostClmnNr).value))
                    End If
                    aKey = CStr(i)
                    ' read DATA
                    aItems.ws_parse_line_simple(aCoPlWs, aCoPlLOff, i, jMax, pHdrLine:=aHdrLOff + 1)
                    If String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i + 1, aCoPlClmnNr).value)) Or aPost.ToUpper = "X" Then
                        If aTSAP_Data_CoPl.fillHeader(aItems) And aTSAP_Data_CoPl.fillData(aItems) Then
                            log.Debug("SapCoPlRibbon_CostActPlan.postSK - " & "calling aSAPHrMaster.PostKeyFigure")
                            Globals.ThisAddIn.Application.StatusBar = "Calling SAP-BAPI at line " & i
                            aRetStr = aSAPCostActivityPlanning.PostKeyFigure(aTSAP_Data_CoPl, pOKMsg:=aOKMsg, pCheck:=pCheck)
                            log.Debug("SapCoPlRibbon_CostActPlan.postSK - " & "aSAPHrMaster.PostKeyFigure returned, aRetStr=" & aRetStr)
                            For Each aKey In aItems.aTDataDic.Keys
                                aCoPlWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        End If
                        aItems = New TData(aIntPar)
                        aTSAP_Data_CoPl = New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postSK")
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, 1).value))
            log.Debug("SapCoPlRibbon_CostActPlan.postSK - " & "all data processed - enabling events, screen update, cursor")
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoPlRibbon_CostActPlan.postSK failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl AddIn")
            log.Error("SapCoPlRibbon_CostActPlan.postSK - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try

    End Sub

    Public Sub postAI(ByRef pSapCon As SapCon, Optional pCheck As Boolean = False)
        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr
        Dim aType As String = "AI"

        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(pSapCon, aIntPar)

        Dim jMax As UInt64 = 0
        Dim aCoPlLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "DATA") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "DATA")), 4)
        Dim aHdrLOff As Integer = If(aIntPar.value(aType & "_" & "LOFF", "HEAD") <> "", CInt(aIntPar.value(aType & "_" & "LOFF", "HEAD")), aCoPlLOff - 4)
        Dim aCoPlWsName As String = If(aIntPar.value(aType & "_" & "WS", "DATA") <> "", aIntPar.value(aType & "_" & "WS", "DATA"), "Data")
        Dim aCoPlWs As Excel.Worksheet
        Dim aMsgClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAMSG") <> "", aIntPar.value(aType & "_" & "COL", "DATAMSG"), "INT-MSG")
        Dim aMsgClmnNr As Integer = 0
        Dim aPostClmn As String = If(aIntPar.value(aType & "_" & "COL", "DATAPOST") <> "", aIntPar.value(aType & "_" & "COL", "DATAPOST"), "INT-POST")
        Dim aPostClmnNr As Integer = 0
        Dim aCoPlClmnNr As Integer = If(aIntPar.value(aType & "_" & "COLNR", "DATACHECK") <> "", CInt(aIntPar.value(aType & "_" & "COLNR", "DATACHECK")), 1)
        Dim aRetStr As String
        Dim aOKMsg As String = If(aIntPar.value(aType & "_" & "RET", "OKMSG") <> "", aIntPar.value(aType & "_" & "RET", "OKMSG"), "OK")

        Dim aWB As Excel.Workbook
        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
        Try
            aCoPlWs = aWB.Worksheets(aCoPlWsName)
        Catch Exc As System.Exception
            MsgBox("No " & aCoPlWsName & " Sheet in current workbook. Check if the current workbook is a valid SAP CoPl Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl")
            Exit Sub
        End Try
        parseHeaderLine(aCoPlWs, jMax, pMsgClmn:=aMsgClmn, pMsgClmnNr:=aMsgClmnNr, pPostClmn:=aPostClmn, pPostClmnNr:=aPostClmnNr, pHdrLine:=aHdrLOff + 1)
        Try
            log.Debug("SapCoPlRibbon_CostActPlan.postAI - " & "processing data - disabling events, screen update, cursor")
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.ThisAddIn.Application.EnableEvents = False
            '            Globals.ThisAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aCoPlLOff + 1
            Dim aKey As String
            Dim aPost As String = ""
            Dim aItems As New TData(aIntPar)
            Dim aTSAP_Data_CoPl As New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postAI")
            Do
                If Left(CStr(aCoPlWs.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        aPost = If(String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, aPostClmnNr).value)), "", CStr(aCoPlWs.Cells(i, aPostClmnNr).value))
                    End If
                    aKey = CStr(i)
                    ' read DATA
                    aItems.ws_parse_line_simple(aCoPlWs, aCoPlLOff, i, jMax, pHdrLine:=aHdrLOff + 1)
                    If String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i + 1, aCoPlClmnNr).value)) Or aPost.ToUpper = "X" Then
                        If aTSAP_Data_CoPl.fillHeader(aItems) And aTSAP_Data_CoPl.fillData(aItems) Then
                            log.Debug("SapCoPlRibbon_CostActPlan.postAI - " & "calling aSAPHrMaster.PostActivityInput")
                            Globals.ThisAddIn.Application.StatusBar = "Calling SAP-BAPI at line " & i
                            aRetStr = aSAPCostActivityPlanning.PostActivityInput(aTSAP_Data_CoPl, pOKMsg:=aOKMsg, pCheck:=pCheck)
                            log.Debug("SapCoPlRibbon_CostActPlan.postAI - " & "aSAPHrMaster.PostActivityInput returned, aRetStr=" & aRetStr)
                            For Each aKey In aItems.aTDataDic.Keys
                                aCoPlWs.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                            Next
                        End If
                        aItems = New TData(aIntPar)
                        aTSAP_Data_CoPl = New TSAP_Data_CoPl(aPar, aIntPar, aSAPCostActivityPlanning, "postAI")
                    End If
                End If
                i += 1
            Loop While Not String.IsNullOrEmpty(CStr(aCoPlWs.Cells(i, 1).value))
            log.Debug("SapCoPlRibbon_CostActPlan.postAI - " & "all data processed - enabling events, screen update, cursor")
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.ThisAddIn.Application.EnableEvents = True
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoPlRibbon_CostActPlan.postAI failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CoPl AddIn")
            log.Error("SapCoPlRibbon_CostActPlan.postAI - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

    Private Sub parseHeaderLine(ByRef pWs As Excel.Worksheet, ByRef pMaxJ As Integer, Optional pMsgClmn As String = "", Optional ByRef pMsgClmnNr As Integer = 0, Optional pHdrLine As Integer = 1, Optional pPostClmn As String = "", Optional ByRef pPostClmnNr As Integer = 0)
        pMaxJ = 0
        Do
            pMaxJ += 1
            If Not String.IsNullOrEmpty(pMsgClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pMsgClmn Then
                pMsgClmnNr = pMaxJ
            ElseIf Not String.IsNullOrEmpty(pPostClmn) And CStr(pWs.Cells(pHdrLine, pMaxJ).value) = pPostClmn Then
                pPostClmnNr = pMaxJ
            End If
        Loop While CStr(pWs.Cells(pHdrLine, pMaxJ + 1).value) <> ""
    End Sub

End Class
