﻿Imports Microsoft.Office.Tools.Ribbon
Imports System.Configuration
Imports System.Collections.Specialized

Imports SAP.Middleware.Connector

Public Class SapCoPlRibbon
    Private aSapCon
    Private aSapGeneral
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private aCoAre As String
    Private aFiscy As String
    Private aPfrom As String
    Private aPto As String
    Private aSVers As String
    Private aTVers As String
    Private aCurt As String
    Private aCompCodes As String
    Private aDelta As String
    Private aAOCtrl As String
    Private aAOSaveMode As String

    Private aPSPar As New TPar

    Private Sub SapCoPlRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Dim sAll As NameValueCollection
        Dim s As String
        Dim enablePS As Boolean = False
        Dim enableCOOMTotal As Boolean = False
        Dim enableCOOMPeriod As Boolean = False
        Dim enableCosting As Boolean = False
        aSapGeneral = New SapGeneral
        Try
            sAll = ConfigurationManager.AppSettings
            s = sAll("enablePS")
            enablePS = Convert.ToBoolean(s)
            s = sAll("enableCOOMTotal")
            enableCOOMTotal = Convert.ToBoolean(s)
            s = sAll("enableCOOMPeriod")
            enableCOOMPeriod = Convert.ToBoolean(s)
            s = sAll("enableCosting")
            enableCosting = Convert.ToBoolean(s)
        Catch Exc As System.Exception
            log.Error("SapAccRibbon_Load - " & "Exception=" & Exc.ToString)
        End Try
        If Not enablePS Then
            Globals.Ribbons.Ribbon1.SAPPsPlan.Visible = False
        Else
            Globals.Ribbons.Ribbon1.SAPPsPlan.Visible = True
        End If
        If Not enableCOOMTotal Then
            Globals.Ribbons.Ribbon1.SAPCoOmPlan.Visible = False
        Else
            Globals.Ribbons.Ribbon1.SAPCoOmPlan.Visible = True
        End If
        If Not enableCOOMPeriod Then
            Globals.Ribbons.Ribbon1.SAPCoOmPlanPer.Visible = False
        Else
            Globals.Ribbons.Ribbon1.SAPCoOmPlanPer.Visible = True
        End If
        If Not enableCosting Then
            Globals.Ribbons.Ribbon1.SAPCoCosting.Visible = False
        Else
            Globals.Ribbons.Ribbon1.SAPCoCosting.Visible = True
        End If

    End Sub

    Private Function checkCon() As Integer
        Dim aSapConRet As Integer
        Dim aSapVersionRet As Integer
        checkCon = False
        log.Debug("checkCon - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            Exit Function
        End If
        log.Debug("checkCon - " & "checking Connection")
        aSapConRet = 0
        If aSapCon Is Nothing Then
            Try
                aSapCon = New SapCon()
            Catch ex As SystemException
                log.Warn("checkCon-New SapCon - )" & ex.ToString)
            End Try
        End If
        Try
            aSapConRet = aSapCon.checkCon()
        Catch ex As SystemException
            log.Warn("checkCon-aSapCon.checkCon - )" & ex.ToString)
        End Try
        If aSapConRet = 0 Then
            log.Debug("checkCon - " & "checking version in SAP")
            Try
                aSapVersionRet = aSapGeneral.checkVersionInSAP(aSapCon)
            Catch ex As SystemException
                log.Warn("checkCon - )" & ex.ToString)
            End Try
            log.Debug("checkCon - " & "aSapVersionRet=" & CStr(aSapVersionRet))
            If aSapVersionRet = True Then
                log.Debug("checkCon - " & "checkCon = True")
                checkCon = True
            Else
                log.Debug("checkCon - " & "connection check failed")
            End If
        End If
    End Function


    Private Sub ButtonLogoff_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogoff.Click
        log.Debug("ButtonLogoff_Click - " & "starting logoff")
        If Not aSapCon Is Nothing Then
            log.Debug("ButtonLogoff_Click - " & "calling aSapCon.SAPlogoff()")
            aSapCon.SAPlogoff()
            aSapCon = Nothing
        End If
        log.Debug("ButtonLogoff_Click - " & "exit")
    End Sub

    Private Sub ButtonLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonLogon.Click
        Dim aConRet As Integer

        log.Debug("ButtonLogon_Click - " & "checking Version")
        If Not aSapGeneral.checkVersion() Then
            log.Debug("ButtonLogon_Click - " & "Version check failed")
            Exit Sub
        End If
        log.Debug("ButtonLogon_Click - " & "creating SapCon")
        If aSapCon Is Nothing Then
            aSapCon = New SapCon()
        End If
        log.Debug("ButtonLogon_Click - " & "calling SapCon.checkCon()")
        aConRet = aSapCon.checkCon()
        If aConRet = 0 Then
            log.Debug("ButtonLogon_Click - " & "connection successfull")
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap Accounting")
        Else
            log.Debug("ButtonLogon_Click - " & "connection failed")
            aSapCon = Nothing
        End If
    End Sub

    Private Function getParameters(pType As String) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim akey As String
        Dim aName As String

        aName = "SAPCoOmPlanning" & pType
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End Try
        akey = CStr(aPws.Cells(1, 1).Value)
        If akey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP CO-OM Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End If
        aCoAre = CStr(aPws.Cells(2, 2).Value)
        aFiscy = CStr(aPws.Cells(3, 2).Value)
        aPfrom = CStr(aPws.Cells(4, 2).Value)
        aPto = CStr(aPws.Cells(5, 2).Value)
        aSVers = CStr(aPws.Cells(6, 2).Value)
        aTVers = CStr(aPws.Cells(7, 2).Value)
        aCurt = CStr(aPws.Cells(8, 2).Value)
        aCompCodes = CStr(aPws.Cells(9, 2).Value)
        aDelta = CStr(aPws.Cells(10, 2).Value)
        aAOCtrl = CStr(aPws.Cells(11, 2).Value)
        aAOSaveMode = CStr(aPws.Cells(12, 2).Value)
        If aCoAre = "" Or
            aFiscy = "" Or
            aPfrom = "" Or
            aPto = "" Or
            aSVers = "" Or
            aTVers = "" Or
            aCurt = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            getParameters = False
            Exit Function
        End If
        getParameters = True
    End Function

    Private Sub ButtonReadAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadAO.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aContrl As New Collection
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim maxData As Integer

        If getParameters("Total") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("O", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        Dim aAOSets As New Collection
        If aAOSaveMode = "X" Or aAOSaveMode = "x" Then
            aRetStr = aSAPCostActivityPlanning.ReadActivityOutputTotS(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aAOSets)
        Else
            aRetStr = aSAPCostActivityPlanning.ReadActivityOutputTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData, aContrl)
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AOData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            Dim aSapContrlRow As Object
            i = 1
            Dim aAOSet As AOSet
            Dim aCOObject As SAPCOObject
            If aAOSaveMode = "X" Then
                For Each aAOSet In aAOSets
                    aCOObject = aAOSet.Key
                    Try
                        aSapDataRow = aAOSet.Total
                        aDws.Cells(i + 1, 1) = aCOObject.Costcenter
                        aDws.Cells(i + 1, 2) = aCOObject.Acttype
                        aDws.Cells(i + 1, 3) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                        aDws.Cells(i + 1, 4) = CStr(aSapDataRow.GetValue("CURRENCY"))
                        aDws.Cells(i + 1, 5) = CDbl(aSapDataRow.GetValue("ACTVTY_QTY"))
                        aDws.Cells(i + 1, 6) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN"))
                        aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("ACTVTY_CAPACTY"))
                        aDws.Cells(i + 1, 8) = CStr(aSapDataRow.GetValue("DIST_KEY_CAPCTY"))
                        aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("PRICE_FIX"))
                        aDws.Cells(i + 1, 10) = CStr(aSapDataRow.GetValue("DIST_KEY_PRICE_FIX"))
                        aDws.Cells(i + 1, 11) = CDbl(aSapDataRow.GetValue("PRICE_VAR"))
                        aDws.Cells(i + 1, 12) = CStr(aSapDataRow.GetValue("DIST_KEY_PRICE_VAR"))
                        aDws.Cells(i + 1, 13) = CInt(aSapDataRow.GetValue("PRICE_UNIT"))
                        aDws.Cells(i + 1, 14) = CStr(aSapDataRow.GetValue("EQUIVALENCE_NO"))
                        aDws.Cells(i + 1, 23) = CInt(aSapDataRow.GetValue("VALUE_INDEX"))
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 23) = "No DataRecord"
                    End Try
                    Try
                        aSapContrlRow = aAOSet.Control
                        aDws.Cells(i + 1, 15) = CStr(aSapContrlRow.GetValue("PRICE_INDICATOR"))
                        aDws.Cells(i + 1, 16) = CStr(aSapContrlRow.GetValue("SWITCH_LAYOUT"))
                        aDws.Cells(i + 1, 17) = CStr(aSapContrlRow.GetValue("ALLOC_COST_ELEM"))
                        aDws.Cells(i + 1, 18) = CStr(aSapContrlRow.GetValue("ACT_PRICE_IND"))
                        aDws.Cells(i + 1, 19) = CStr(aSapContrlRow.GetValue("PREDIS_FXD_COST"))
                        aDws.Cells(i + 1, 20) = CStr(aSapContrlRow.GetValue("ACT_CAT_ACTUAL"))
                        aDws.Cells(i + 1, 21) = CStr(aSapContrlRow.GetValue("AVERAGE_PRICE_IND"))
                        aDws.Cells(i + 1, 22) = CInt(aSapContrlRow.GetValue("ATTRIB_INDEX"))
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 22) = "No ControlRecord"
                    End Try
                    i += 1
                Next
            Else
                If aData.Count > 0 Then
                    Do
                        Try
                            aSapDataRow = aData(i)
                            aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                            aDws.Cells(i + 1, 2) = aObjects(i).Acttype
                            aDws.Cells(i + 1, 3) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                            aDws.Cells(i + 1, 4) = CStr(aSapDataRow.GetValue("CURRENCY"))
                            aDws.Cells(i + 1, 5) = CDbl(aSapDataRow.GetValue("ACTVTY_QTY"))
                            aDws.Cells(i + 1, 6) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN"))
                            aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("ACTVTY_CAPACTY"))
                            aDws.Cells(i + 1, 8) = CStr(aSapDataRow.GetValue("DIST_KEY_CAPCTY"))
                            aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("PRICE_FIX"))
                            aDws.Cells(i + 1, 10) = CStr(aSapDataRow.GetValue("DIST_KEY_PRICE_FIX"))
                            aDws.Cells(i + 1, 11) = CDbl(aSapDataRow.GetValue("PRICE_VAR"))
                            aDws.Cells(i + 1, 12) = CStr(aSapDataRow.GetValue("DIST_KEY_PRICE_VAR"))
                            aDws.Cells(i + 1, 13) = CInt(aSapDataRow.GetValue("PRICE_UNIT"))
                            aDws.Cells(i + 1, 14) = CStr(aSapDataRow.GetValue("EQUIVALENCE_NO"))
                            aDws.Cells(i + 1, 23) = CInt(aSapDataRow.GetValue("VALUE_INDEX"))
                        Catch Ex As System.Exception
                            aDws.Cells(i + 1, 23) = "No DataRecord"
                        End Try
                        Try
                            aSapContrlRow = aContrl(i)
                            aDws.Cells(i + 1, 15) = CStr(aSapContrlRow.GetValue("PRICE_INDICATOR"))
                            aDws.Cells(i + 1, 16) = CStr(aSapContrlRow.GetValue("SWITCH_LAYOUT"))
                            aDws.Cells(i + 1, 17) = CStr(aSapContrlRow.GetValue("ALLOC_COST_ELEM"))
                            aDws.Cells(i + 1, 18) = CStr(aSapContrlRow.GetValue("ACT_PRICE_IND"))
                            aDws.Cells(i + 1, 19) = CStr(aSapContrlRow.GetValue("PREDIS_FXD_COST"))
                            aDws.Cells(i + 1, 20) = CStr(aSapContrlRow.GetValue("ACT_CAT_ACTUAL"))
                            aDws.Cells(i + 1, 21) = CStr(aSapContrlRow.GetValue("AVERAGE_PRICE_IND"))
                            aDws.Cells(i + 1, 22) = CInt(aSapContrlRow.GetValue("ATTRIB_INDEX"))
                        Catch Ex As System.Exception
                            aDws.Cells(i + 1, 22) = "No ControlRecord"
                        End Try
                        i += 1
                    Loop While i <= aData.Count
                End If
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadAO_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAO.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aDataRow As New Collection
        Dim aContrlRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Total") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AOData")
        Catch Exc As System.Exception
            MsgBox("No AOData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 2).Value), "")
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 3 To 14
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                aContrlRow = New Collection
                For J = 15 To 21
                    aVal = aDws.Cells(i, J).Value
                    aContrlRow.Add(aVal)
                Next J
                aContrl.Add(aContrlRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityOutputTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aContrl, aDelta, aAOCtrl)
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostAO_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonReadPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPC.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters("Total") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("P", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadPrimCostTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Try
            Dim aDws As Excel.Worksheet
            Dim aWB As Excel.Workbook
            aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
            aDws = aWB.Worksheets("PData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            Dim aCells As Excel.Range
            i = 1
            If aData.Count > 0 Then
                Do
                    Try
                        aSapDataRow = aData(i)
                        aCells = aDws.Range(aDws.Cells(i, 1), aDws.Cells(i, 4))
                        aCells.NumberFormat = "@"
                        aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                        aDws.Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
                        aDws.Cells(i + 1, 3) = aObjects(i).Acttype
                        aDws.Cells(i + 1, 4) = aObjects(i).Costelem
                        aDws.Cells(i + 1, 5) = CStr(aSapDataRow.GetValue("TRANS_CURR"))
                        aDws.Cells(i + 1, 6) = CDbl(aSapDataRow.GetValue("FIX_VALUE"))
                        aDws.Cells(i + 1, 7) = CStr(aSapDataRow.GetValue("DIST_KEY_FIX_VAL"))
                        aDws.Cells(i + 1, 8) = CDbl(aSapDataRow.GetValue("VAR_VALUE"))
                        aDws.Cells(i + 1, 9) = CStr(aSapDataRow.GetValue("DIST_KEY_VAR_VAL"))
                        aDws.Cells(i + 1, 10) = CDbl(aSapDataRow.GetValue("FIX_QUAN"))
                        aDws.Cells(i + 1, 11) = CStr(aSapDataRow.GetValue("DIST_KEY_FIX_QUAN"))
                        aDws.Cells(i + 1, 12) = CDbl(aSapDataRow.GetValue("VAR_QUAN"))
                        aDws.Cells(i + 1, 13) = CStr(aSapDataRow.GetValue("DIST_KEY_VAR_QUAN"))
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 15) = "No DataRecord"
                    End Try
                    i += 1
                Loop While i <= aData.Count
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPC_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPC.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Total") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("PData")
        Catch Exc As System.Exception
            MsgBox("No PData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value),
                                               CStr(aDws.Cells(i, 3).Value),
                                               CStr(aDws.Cells(i, 4).Value), "", "",
                                               CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 6 To 14
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostPrimCostTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aDelta)
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostPC_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonReadAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadAI.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters("Total") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("I", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadActivityInputTot(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AIData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            i = 1
            If aData.Count > 0 Then
                Do
                    Try
                        aSapDataRow = aData(i)
                        aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                        aDws.Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
                        aDws.Cells(i + 1, 3) = aObjects(i).Acttype
                        aDws.Cells(i + 1, 4) = aObjects(i).SCostcenter
                        aDws.Cells(i + 1, 5) = aObjects(i).SActtype
                        aDws.Cells(i + 1, 6) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                        aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("QUANTITY_FIX"))
                        aDws.Cells(i + 1, 8) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN_FIX"))
                        aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("QUANTITY_VAR"))
                        aDws.Cells(i + 1, 10) = CStr(aSapDataRow.GetValue("DIST_KEY_QUAN_VAR"))
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 12) = "No DataRecord"
                    End Try
                    i += 1
                Loop While i <= aData.Count
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadAI_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAI.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Total") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AIData")
        Catch Exc As System.Exception
            MsgBox("No AIData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value),
                                               CStr(aDws.Cells(i, 3).Value), "",
                                               CStr(aDws.Cells(i, 4).Value), CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 6 To 10
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityInputTot(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostAI_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonReadSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadSK.Click
        ButtonReadSK_Execute("Total")
    End Sub

    Private Sub ButtonReadSK_Execute(pType As String)
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range

        If getParameters(pType) = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("S", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadKeyFigure(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("SKData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            i = 1
            If aData.Count > 0 Then
                Do
                    Try
                        aSapDataRow = aData(i)
                        aDws.Cells(i + 1, 1) = aObjects(i).Costcenter
                        aDws.Cells(i + 1, 2) = aObjects(i).WBS_ELEMENT
                        aDws.Cells(i + 1, 3) = aObjects(i).Acttype
                        aDws.Cells(i + 1, 4) = CStr(aSapDataRow.GetValue("STATKEYFIG"))
                        aDws.Cells(i + 1, 5) = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                        aDws.Cells(i + 1, 6) = CDbl(aSapDataRow.GetValue("QUANTITY_PER01"))
                        aDws.Cells(i + 1, 7) = CDbl(aSapDataRow.GetValue("QUANTITY_PER02"))
                        aDws.Cells(i + 1, 8) = CDbl(aSapDataRow.GetValue("QUANTITY_PER03"))
                        aDws.Cells(i + 1, 9) = CDbl(aSapDataRow.GetValue("QUANTITY_PER04"))
                        aDws.Cells(i + 1, 10) = CDbl(aSapDataRow.GetValue("QUANTITY_PER05"))
                        aDws.Cells(i + 1, 11) = CDbl(aSapDataRow.GetValue("QUANTITY_PER06"))
                        aDws.Cells(i + 1, 12) = CDbl(aSapDataRow.GetValue("QUANTITY_PER07"))
                        aDws.Cells(i + 1, 13) = CDbl(aSapDataRow.GetValue("QUANTITY_PER08"))
                        aDws.Cells(i + 1, 14) = CDbl(aSapDataRow.GetValue("QUANTITY_PER09"))
                        aDws.Cells(i + 1, 15) = CDbl(aSapDataRow.GetValue("QUANTITY_PER10"))
                        aDws.Cells(i + 1, 16) = CDbl(aSapDataRow.GetValue("QUANTITY_PER11"))
                        aDws.Cells(i + 1, 17) = CDbl(aSapDataRow.GetValue("QUANTITY_PER12"))
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 19) = "No DataRecord"
                    End Try
                    i += 1
                Loop While i <= aData.Count
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadSK_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostSK.Click
        ButtonPostSK_Execute("Total")
    End Sub

    Private Sub ButtonPostSK_Execute(pType As String)
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters(pType) = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("SKData")
        Catch Exc As System.Exception
            MsgBox("No SKData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 3).Value), "", "", "",
                                                   CStr(aDws.Cells(i, 2).Value), CStr(aDws.Cells(i, 4).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 5 To 17
                    aVal = aDws.Cells(i, J).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostKeyFigure(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aDelta)
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostSK_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonReadPerAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerAO.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aContrl As New Collection
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim J As Integer
        Dim aVal

        If getParameters("Periodic") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("O", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadActivityOutput(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AOData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            i = 1
            If aData.Count > 0 Then
                Do
                    Try
                        aSapDataRow = aData(i)
                        aDws.Cells(i + 1, 1).Value = aObjects(i).Costcenter
                        aDws.Cells(i + 1, 2).Value = aObjects(i).Acttype
                        aDws.Cells(i + 1, 3).Value = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                        aDws.Cells(i + 1, 4).Value = CStr(aSapDataRow.GetValue("CURRENCY"))
                        For J = 2 To 65
                            aVal = CDbl(aSapDataRow.GetValue(J - 1))
                            aDws.Cells(i + 1, J + 3).Value = aVal
                        Next J
                        For J = 66 To 97
                            aVal = CInt(aSapDataRow.GetValue(J - 1))
                            aDws.Cells(i + 1, J + 3).Value = aVal
                        Next J
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 1) = "No DataRecord"
                    End Try
                    i += 1
                Loop While i <= aData.Count
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPerAO_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPerAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerAO.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aDataRow As New Collection
        Dim aContrlRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Periodic") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AOData")
        Catch Exc As System.Exception
            MsgBox("No AOData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value), CStr(aDws.Cells(i, 2).Value), "")
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 2 To 65
                    aVal = aDws.Cells(i, J + 3).Value
                    aDataRow.Add(CDbl(aVal))
                Next J
                For J = 66 To 97
                    aVal = aDws.Cells(i, J + 3).Value
                    aDataRow.Add(CInt(aVal))
                Next J
                aDataRow.Add(CStr(aDws.Cells(i, 3).Value)) 'Unit of Measure
                aDataRow.Add(CStr(aDws.Cells(i, 4).Value)) 'Curr.
                aData.Add(aDataRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityOutput(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aDelta)
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostPerAO_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonReadPerPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerPC.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim J As Integer
        Dim aVal

        If getParameters("Periodic") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("P", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadPrimCost(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Try
            Dim aDws As Excel.Worksheet
            Dim aWB As Excel.Workbook
            aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
            aDws = aWB.Worksheets("PData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            Dim aCells As Excel.Range
            i = 1
            If aData.Count > 0 Then
                Do
                    Try
                        aSapDataRow = aData(i)
                        aCells = aDws.Range(aDws.Cells(i, 1), aDws.Cells(i, 4))
                        '                    aCells.NumberFormat = "@"
                        aDws.Cells(i + 1, 1).Value = aObjects(i).Costcenter
                        aDws.Cells(i + 1, 2).Value = aObjects(i).WBS_ELEMENT
                        aDws.Cells(i + 1, 3).Value = aObjects(i).Acttype
                        aDws.Cells(i + 1, 4).Value = aObjects(i).Costelem
                        aDws.Cells(i + 1, 5).Value = CStr(aSapDataRow.GetValue("TRANS_CURR"))
                        For J = 8 To 71
                            aVal = CDbl(aSapDataRow.GetValue(J - 1))
                            aDws.Cells(i + 1, J - 2).Value = aVal
                        Next J
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 1) = "No DataRecord"
                    End Try
                    i += 1
                Loop While i <= aData.Count
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPerPC_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPerPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerPC.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Periodic") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("PData")
        Catch Exc As System.Exception
            MsgBox("No PData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value),
                                               CStr(aDws.Cells(i, 3).Value),
                                               CStr(aDws.Cells(i, 4).Value), "", "",
                                               CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                For J = 8 To 71
                    aVal = aDws.Cells(i, J - 2).Value
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostPrimCost(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData, aDelta)
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostPerPC_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonReadPerAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerAI.Click
        Dim aSAPCOObject As New SAPCOObject
        Dim aCompCodeSplit
        Dim aCompCode
        Dim aData As New Collection
        Dim aContrl As New Collection
        Dim aObjects As New Collection
        Dim aRetStr As String
        Dim i As Integer
        Dim aRange As Excel.Range
        Dim J As Integer
        Dim aVal

        If getParameters("Periodic") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        aCompCodeSplit = Split(aCompCodes, ";")
        Dim aSAPGetCOObject As New SAPGetCOObject(aSapCon)
        Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
        Globals.SapCoPlExcelAddin.Application.EnableEvents = False
        Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
        For Each aCompCode In aCompCodeSplit
            aSAPGetCOObject.GetCoObjects("I", aFiscy, aSVers, aCoAre, CStr(aCompCode), aObjects)
        Next aCompCode
        If aObjects.Count = 0 Then
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Exit Sub
        End If
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        aRetStr = aSAPCostActivityPlanning.ReadActivityInput(aCoAre, aFiscy, aPfrom, aPto, aSVers, aCurt, aObjects, aData)
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AIData")
            aDws.Activate()
            If CStr(aDws.Cells(2, 1).Value) <> "" Then
                aRange = aDws.Range("A2")
                i = 2
                Do
                    i += 1
                Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> "" Or CStr(aDws.Cells(i, 3).value) <> ""
                aRange = aDws.Range(aRange, aDws.Cells(i, 1))
                aRange.EntireRow.Delete()
            End If
            Dim aSapDataRow As Object
            i = 1
            If aData.Count > 0 Then
                Do
                    Try
                        aSapDataRow = aData(i)
                        aDws.Cells(i + 1, 1).Value = aObjects(i).Costcenter
                        aDws.Cells(i + 1, 2).Value = aObjects(i).WBS_ELEMENT
                        aDws.Cells(i + 1, 3).Value = aObjects(i).Acttype
                        aDws.Cells(i + 1, 4).Value = aObjects(i).SCostcenter
                        aDws.Cells(i + 1, 5).Value = aObjects(i).SActtype
                        aDws.Cells(i + 1, 6).Value = CStr(aSapDataRow.GetValue("UNIT_OF_MEASURE"))
                        For J = 6 To 37
                            aVal = CDbl(aSapDataRow.GetValue(J - 1))
                            aDws.Cells(i + 1, J + 1).Value = aVal
                        Next J
                    Catch Ex As System.Exception
                        aDws.Cells(i + 1, 1) = "No DataRecord"
                    End Try
                    i += 1
                Loop While i <= aData.Count
            End If
            ' aDws.Cells(i + 1, 3) = aRetStr
            Dim aRetStrSplit
            If aRetStr <> "" Then
                i += 1
                aRetStrSplit = Split(aRetStr, ";")
                For Each aRetStr In aRetStrSplit
                    If aRetStr <> "" Then
                        aDws.Cells(i, 3) = aRetStr
                        i += 1
                    End If
                Next aRetStr
            End If
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonReadPerAI_Click")
            Exit Sub
        End Try
    End Sub

    Private Sub ButtonPostPerAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerAI.Click
        Dim i As Integer
        Dim aData As New Collection
        Dim aDataRow As New Collection
        Dim aObjects As New Collection
        Dim aVal
        Dim aRetStr As String

        If getParameters("Periodic") = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If
        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aDws = aWB.Worksheets("AIData")
        Catch Exc As System.Exception
            MsgBox("No AIData Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO-OM")
            Exit Sub
        End Try
        aRetStr = ""
        aDws.Activate()
        Try
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            i = 2
            Do
                Dim aSAPCOObject = New SAPCOObject
                aSAPCOObject = aSAPCOObject.create(CStr(aDws.Cells(i, 1).Value),
                                               CStr(aDws.Cells(i, 3).Value), "",
                                               CStr(aDws.Cells(i, 4).Value), CStr(aDws.Cells(i, 5).Value), CStr(aDws.Cells(i, 2).Value))
                aObjects.Add(aSAPCOObject)
                aDataRow = New Collection
                aDataRow.Add(CStr(aDws.Cells(i, 6).Value)) ' Unit of Meassure
                For J = 6 To 37
                    aVal = CDbl(aDws.Cells(i, J + 1).Value)
                    aDataRow.Add(aVal)
                Next J
                aData.Add(aDataRow)
                i += 1
            Loop While CStr(aDws.Cells(i, 1).value) <> "" Or CStr(aDws.Cells(i, 2).value) <> ""
            Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
            aRetStr = aSAPCostActivityPlanning.PostActivityInput(aCoAre, aFiscy, aPfrom, aPto, aTVers, aCurt, aObjects, aData)
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch Ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "ButtonPostAI_Click")
        End Try
        ' aDws.Cells(i, 3) = aRetStr
        Dim aRetStrSplit
        If aRetStr <> "" Then
            aRetStrSplit = Split(aRetStr, ";")
            For Each aRetStr In aRetStrSplit
                If aRetStr <> "" Then
                    aDws.Cells(i, 3) = aRetStr
                    i += 1
                End If
            Next aRetStr
        End If
    End Sub

    Private Sub ButtonPostPerSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPerSK.Click
        ButtonPostSK_Execute("Periodic")
    End Sub

    Private Sub ButtonReadPerSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonReadPerSK.Click
        ButtonReadSK_Execute("Periodic")
    End Sub

    Private Function getPsParameters() As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim akey As String
        Dim aName As String
        Dim i As Integer

        aName = "SAPPsPlanning"
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP PS Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            getPsParameters = False
            Exit Function
        End Try
        akey = CStr(aPws.Cells(1, 1).Value)
        If akey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP PS Planning Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            getPsParameters = False
            Exit Function
        End If
        i = 2
        aPSPar = New TPar
        Do
            aPSPar.addPar(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""

        If aPSPar.value("HEADERINFO", "CO_AREA") = "" Or aPSPar.value("HEADERINFO", "VERSION") = "" Then
            MsgBox("Please fill all obligatory fields in the parameter sheet!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            getPsParameters = False
            Exit Function
        End If

        getPsParameters = True
    End Function

    Private Sub ButtonPsUpdCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPsUpdCheck.Click
        log.Debug("ButtonPsUpdCheck_Click - " & "calling SAP_PsUpd_execute")
        SAP_PsUpd_execute(pTest:=True)
    End Sub

    Private Sub ButtonPsUpdPost_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPsUpdPost.Click
        log.Debug("ButtonPsUpdCheck_Click - " & "calling SAP_PsUpd_execute")
        SAP_PsUpd_execute(pTest:=False)
    End Sub

    Private Sub SAP_PsUpd_execute(pTest As Boolean)
        Dim aR As Excel.Range
        Dim aTop, aBottom, aLeft, aRight As Integer
        If getPsParameters() = False Then
            Exit Sub
        End If
        If checkCon() = False Then
            Exit Sub
        End If

        Dim aWB As Excel.Workbook
        Dim aDwsName As String = If(aPSPar.value("WS", "PSIMPORT") <> "", aPSPar.value("WS", "PSIMPORT"), "PS-Import")
        Dim aDws As Excel.Worksheet
        Dim aWsMsgName As String = If(aPSPar.value("WS", "MSG") <> "", aPSPar.value("WS", "MSG"), "Messages")
        Dim aWsMsg As Excel.Worksheet
        Dim aWsWBSName As String = If(aPSPar.value("WS", "WBS") <> "", aPSPar.value("WS", "WBS"), "WBS")
        Dim aWsWBS As Excel.Worksheet
        Dim aWsLimName As String = If(aPSPar.value("WS", "WBSLIMIT") <> "", aPSPar.value("WS", "WBSLIMIT"), "WBS-Limit")
        Dim aWsLim As Excel.Worksheet
        Dim aNoWsLim As Boolean = False
        Dim aWsCEName As String = If(aPSPar.value("WS", "CEMAPPING") <> "", aPSPar.value("WS", "CEMAPPING"), "CE-Mapping")
        Dim aWsCE As Excel.Worksheet
        Dim aNoWsCE As Boolean = False

        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        ' PS-Import
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid PS Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            Exit Sub
        End Try
        ' Messages
        Try
            aWsMsg = aWB.Worksheets(aWsMsgName)
        Catch Exc As System.Exception
            MsgBox("No " & aWsMsgName & " Sheet in current workbook. Check if the current workbook is a valid PS Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            Exit Sub
        End Try
        ' WBS-Info
        Try
            aWsWBS = aWB.Worksheets(aWsWBSName)
        Catch Exc As System.Exception
            MsgBox("No " & aWsWBSName & " Sheet in current workbook. Check if the current workbook is a valid PS Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            Exit Sub
        End Try
        ' WBS-Limit
        Try
            aWsLim = aWB.Worksheets(aWsLimName)
        Catch Exc As System.Exception
            aWsLim = Nothing
            aNoWsLim = True
        End Try
        ' CE-Mapping
        Try
            aWsCE = aWB.Worksheets(aWsCEName)
        Catch Exc As System.Exception
            aWsCE = Nothing
            aNoWsCE = True
        End Try

        '   collect the WBSLimits
        Dim aWBSLimit As New WBSLimit
        Dim i As Integer = 2
        If Not aNoWsLim Then
            While CStr(aWsLim.Cells(i, 1).Value) <> ""
                aWBSLimit.addWBSLimit(CStr(aWsLim.Cells(i, 1).Value), CInt(aWsLim.Cells(i, 2).Value), CInt(aWsLim.Cells(i, 3).Value),
                                  CInt(aWsLim.Cells(i, 4).Value), CInt(aWsLim.Cells(i, 5).Value))
                i += 1
            End While
        End If

        '   collect the CeMap
        Dim aCeMap As New CeMap
        i = 2
        If Not aNoWsCE Then
            While CStr(aWsCE.Cells(i, 1).Value) <> ""
                aCeMap.addCeMap(CStr(aWsCE.Cells(i, 1).Value), CStr(aWsCE.Cells(i, 2).Value))
                i += 1
            End While
        End If

        '   collect the WBSInfo
        Dim aWBSInfo As New WBSInfo(aPSPar)
        Dim aRaWBSName As String = If(aPSPar.value("RA", "WBS") <> "", aPSPar.value("RA", "WBS"), "SAPPS_WBS")
        Dim aLOff As Integer = If(aPSPar.value("LOFF", "WBS") <> "", CInt(aPSPar.value("LOFF", "WBS")), 1)
        Try
            aR = aWsWBS.Range(aRaWBSName)
        Catch Exc As System.Exception
            MsgBox("No " & aRaWBSName & " Range in " & aWsWBSName & " Sheet in current workbook. Check if the current workbook is a valid PS Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            Exit Sub
        End Try
        aTop = aR.Row
        aBottom = aR.Row + aR.Rows.Count - 1
        aLeft = aR.Column
        aRight = aLeft + aR.Columns.Count - 1
        For i = aTop + 1 To aBottom
            ' Columns should be dynamic -> later
            aWBSInfo.addWBSInfo(CStr(aWsWBS.Cells(i, aLeft).Value), CStr(aWsWBS.Cells(i, aLeft + 1).Value),
                            CStr(aWsWBS.Cells(i, aLeft + 5).Value))
        Next i

        '   collect the plan records
        Dim aRaDName As String = If(aPSPar.value("RA", "PSIMPORT") <> "", aPSPar.value("RA", "PSIMPORT"), "SAPPS_Import")
        aLOff = If(aPSPar.value("LOFF", "PSIMPORT") <> "", CInt(aPSPar.value("LOFF", "PSIMPORT")), 1)
        Dim aKey As String
        Dim aTPlan As New TPlan(aPSPar)
        Dim aTPlanRec As New TPlanRec
        Dim aWBSClmn As String = If(aPSPar.value("COL", "PSIMPORTWBS") <> "", aPSPar.value("COL", "PSIMPORTWBS"), "COOBJECT-WBS_ELEMENT")
        Dim aWBS As String

        Try
            aR = aDws.Range(aRaDName)
        Catch Exc As System.Exception
            MsgBox("No " & aRaDName & " Range in " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid PS Planning Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP PS")
            Exit Sub
        End Try

        aTop = aR.Row
        aBottom = aR.Row + aR.Rows.Count - 1
        aLeft = aR.Column
        aRight = aLeft + aR.Columns.Count - 1
        For i = aTop + aLOff To aBottom
            aKey = CStr(i)
            aWBS = ""
            For j = aLeft To aRight
                If CStr(aDws.Cells(aTop - 3, j).value) <> "N/A" And CStr(aDws.Cells(aTop - 3, j).value) <> "" Then
                    If CStr(aDws.Cells(aTop - 3, j).value) = aWBSClmn Then
                        aWBS = CStr(aDws.Cells(i, j).value)
                    End If
                    aTPlan.addValue(aKey, CStr(aDws.Cells(aTop - 3, j).value), CStr(aDws.Cells(i, j).value), CStr(aDws.Cells(aTop - 2, j).value), CStr(aDws.Cells(aTop - 1, j).value))
                End If
            Next
            ' remove plan records for WBS that are closed or not planning elements
            If aWBSInfo.isClosed(aWBS) Or Not aWBSInfo.isPlanningAllowed(aWBS) Then
                aTPlan.delPlan(aKey)
            Else
                If aWBSLimit.aWBSLimitCol.Count > 0 Then
                    aTPlan.checkWBSLimit(aWBSLimit, aKey)
                End If
                If aCeMap.aCeMapCol.Count > 0 Then
                    aTPlan.mapCE(aCeMap, aKey)
                End If
                aTPlan.buildValFields(aKey)
            End If
        Next
        ' we need plan records per year -> transform to collection of SAP_Plan
        ' seperate primary cost from activities
        Dim aTSAP_Plan_PC As New TSAP_Plan(aPSPar)
        aTSAP_Plan_PC.PC_fromTPlan(aTPlan)
        Dim aTSAP_Plan_AI As New TSAP_Plan(aPSPar)
        aTSAP_Plan_AI.AI_fromTPlan(aTPlan)

        '   Clear message area
        aWsMsg.Activate()
        If CStr(aWsMsg.Cells(3, 1).Value) <> "" Then
            aWsMsg.Range("A3").Select()
            Dim aRange = aWsMsg.Range("A3")
            i = 3
            Do
                i += 1
            Loop While CStr(aWsMsg.Cells(i, 1).Value) <> ""
            aRange = aWsMsg.Range(aRange, aWsMsg.Cells(i, 1))
            aRange.EntireRow.Delete()
        End If
        '   Post the primary cost
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon)
        Dim aMsgCnt = 1
        Dim aSAP_PC_Plan As Collection
        Dim aYear As String
        Dim aSapRetCol As Collection
        For Each aSAP_PC_Plan In aTSAP_Plan_PC.aTSAP_PlanCol
            aYear = ""
            If aSAP_PC_Plan.Count > 0 Then
                aTPlanRec = aSAP_PC_Plan(1)
                aYear = aTPlanRec.getYear(aPSPar)
                aWsMsg.Cells(aMsgCnt + 2, 1) = "Primary Cost - Messages for Year: " & aYear
                aMsgCnt += 1
                aSapRetCol = aSAPCostActivityPlanning.PostPrimCostDyn(aPSPar.value("HEADERINFO", "CO_AREA"),
                            aYear,
                            "01",
                            "12",
                            aPSPar.value("HEADERINFO", "VERSION"),
                            aPSPar.value("HEADERINFO", "PLAN_CURRTYPE"),
                            aSAP_PC_Plan, pTest:=pTest)
                For Each aSapRet In aSapRetCol
                    aWsMsg.Cells(aMsgCnt + 2, 1) = aSapRet
                    aMsgCnt = aMsgCnt + 1
                Next
            End If
        Next
        Dim aSAP_AI_Plan As Collection
        For Each aSAP_AI_Plan In aTSAP_Plan_AI.aTSAP_PlanCol
            aYear = ""
            If aSAP_AI_Plan.Count > 0 Then
                aTPlanRec = aSAP_AI_Plan(1)
                aYear = aTPlanRec.getYear(aPSPar)
                aWsMsg.Cells(aMsgCnt + 2, 1) = "Activity Input - Messages for Year: " & aYear
                aMsgCnt += 1
                aSapRetCol = aSAPCostActivityPlanning.PostActivityInputDyn(aPSPar.value("HEADERINFO", "CO_AREA"),
                            aYear,
                            "01",
                            "12",
                            aPSPar.value("HEADERINFO", "VERSION"),
                            aPSPar.value("HEADERINFO", "PLAN_CURRTYPE"),
                            aSAP_AI_Plan, pTest:=pTest)
                For Each aSapRet In aSapRetCol
                    aWsMsg.Cells(aMsgCnt + 2, 1) = aSapRet
                    aMsgCnt = aMsgCnt + 1
                Next
            End If
        Next

    End Sub

    Private Sub ButtonCostingCreate_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCostingCreate.Click
        Dim aSapCoPlRibbonCosting As New SapCoPlRibbonCosting
        If checkCon() = True Then
            aSapCoPlRibbonCosting.exec(pSapCon:=aSapCon, pMode:="Create")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCostingCreate_Click")
        End If
    End Sub

    Private Sub ButtonCostingChange_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCostingChange.Click
        Dim aSapCoPlRibbonCosting As New SapCoPlRibbonCosting
        If checkCon() = True Then
            aSapCoPlRibbonCosting.exec(pSapCon:=aSapCon, pMode:="Change")
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCostingChange_Click")
        End If
    End Sub
End Class
