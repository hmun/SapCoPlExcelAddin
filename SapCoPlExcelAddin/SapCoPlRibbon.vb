Imports Microsoft.Office.Tools.Ribbon
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
            MsgBox("SAP-Logon successful! ", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Sap LTP")
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
        aWB = Globals.ThisAddin.Application.ActiveWorkbook
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

    Dim aPsPar As TPar
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
        aPsPar = New TPar
        Do
            aPsPar.addPar(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""

        If aPsPar.value("HEADERINFO", "CO_AREA") = "" Or aPsPar.value("HEADERINFO", "VERSION") = "" Then
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

        aWB = Globals.ThisAddIn.Application.ActiveWorkbook
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

        ''   collect the plan records
        Dim aRaDName As String = If(aPsPar.value("RA", "PSIMPORT") <> "", aPsPar.value("RA", "PSIMPORT"), "SAPPS_Import")
        aLOff = If(aPsPar.value("LOFF", "PSIMPORT") <> "", CInt(aPsPar.value("LOFF", "PSIMPORT")), 1)
        Dim aKey As String
        Dim aTPlan As New TPlan(aPsPar)
        Dim aTPlanRec As New TPlanRec
        Dim aWBSClmn As String = If(aPsPar.value("COL", "PSIMPORTWBS") <> "", aPsPar.value("COL", "PSIMPORTWBS"), "COOBJECT-WBS_ELEMENT")
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
        Dim aTSAP_Plan_PC As New TSAP_Plan(aPsPar)
        aTSAP_Plan_PC.PC_fromTPlan(aTPlan)
        Dim aTSAP_Plan_AI As New TSAP_Plan(aPsPar)
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
        Dim aPar As New SAPCommon.TStr 'need for SAPCostActivityPlanning but can be empty
        Dim aSAPCostActivityPlanning As New SAPCostActivityPlanning(aSapCon, aPar)
        Dim aMsgCnt = 1
        Dim aSAP_PC_Plan As Collection
        Dim aYear As String
        Dim aSapRetCol As Collection
        For Each aSAP_PC_Plan In aTSAP_Plan_PC.aTSAP_PlanCol
            aYear = ""
            If aSAP_PC_Plan.Count > 0 Then
                aTPlanRec = aSAP_PC_Plan(1)
                aYear = aTPlanRec.getYear(aPsPar)
                aWsMsg.Cells(aMsgCnt + 2, 1) = "Primary Cost - Messages for Year: " & aYear
                aMsgCnt += 1
                aSapRetCol = aSAPCostActivityPlanning.PostPrimCostDyn(aPsPar.value("HEADERINFO", "CO_AREA"),
                            aYear,
                            "01",
                            "12",
                            aPsPar.value("HEADERINFO", "VERSION"),
                            aPsPar.value("HEADERINFO", "PLAN_CURRTYPE"),
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
                aYear = aTPlanRec.getYear(aPsPar)
                aWsMsg.Cells(aMsgCnt + 2, 1) = "Activity Input - Messages for Year: " & aYear
                aMsgCnt += 1
                aSapRetCol = aSAPCostActivityPlanning.PostActivityInputDyn(aPsPar.value("HEADERINFO", "CO_AREA"),
                            aYear,
                            "01",
                            "12",
                            aPsPar.value("HEADERINFO", "VERSION"),
                            aPsPar.value("HEADERINFO", "PLAN_CURRTYPE"),
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

    Private Sub ButtonCheckAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckAO.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postAO(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCheckAO")
        End If
    End Sub

    Private Sub ButtonPostAO_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAO.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postAO(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostAO")
        End If
    End Sub

    Private Sub ButtonCheckPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckPC.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postPC(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCheckPC")
        End If
    End Sub

    Private Sub ButtonPostPC_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostPC.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postPC(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostPC")
        End If
    End Sub

    Private Sub ButtonCheckAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckAI.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postAI(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCheckAI")
        End If
    End Sub

    Private Sub ButtonPostAI_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostAI.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postAI(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostAI")
        End If
    End Sub

    Private Sub ButtonCheckSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonCheckSK.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postSK(pSapCon:=aSapCon, pCheck:=True)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonCheckSK")
        End If
    End Sub

    Private Sub ButtonPostSK_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonPostSK.Click
        Dim aSapClass As New SapCoPlRibbon_CostActPlan
        If checkCon() = True Then
            aSapClass.postSK(pSapCon:=aSapCon, pCheck:=False)
        Else
            MsgBox("Checking SAP-Connection failed!", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap ButtonPostSK")
        End If
    End Sub

    Private Sub ButtonGenData_Click(sender As Object, e As RibbonControlEventArgs) Handles ButtonGenData.Click
        Dim aRibbon_Generate As New Ribbon_Generate
        aRibbon_Generate.GenerateData()
    End Sub
End Class
