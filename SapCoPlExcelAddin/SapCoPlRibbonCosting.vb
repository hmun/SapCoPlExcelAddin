' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Public Class SapCoPlRibbonCosting

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Function getGenParameters(ByRef pPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim aKey As String
        Dim aName As String
        Dim i As Integer
        log.Debug("SapCoPlRibbonCosting getGenParametrs - " & "reading Parameter")
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter")
        Catch Exc As System.Exception
            MsgBox("No Parameter Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CO-OM")
            getGenParameters = False
            Exit Function
        End Try
        aName = "SAPCoPlCosting"
        aKey = CStr(aPws.Cells(1, 1).Value)
        If aKey <> aName Then
            MsgBox("Cell A1 of the parameter sheet does not contain the key " & aName & ". Check if the current workbook is a valid SAP CO-OM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CO-OM")
            getGenParameters = False
            Exit Function
        End If
        i = 2
        pPar = New SAPCommon.TStr
        Do
            pPar.add(CStr(aPws.Cells(i, 2).value), CStr(aPws.Cells(i, 4).value), pFORMAT:=CStr(aPws.Cells(i, 3).value))
            i += 1
        Loop While CStr(aPws.Cells(i, 2).value) <> "" Or CStr(aPws.Cells(i, 2).value) <> ""
        getGenParameters = True
    End Function
    Private Function getIntParameters(ByRef pIntPar As SAPCommon.TStr) As Integer
        Dim aPws As Excel.Worksheet
        Dim aWB As Excel.Workbook
        Dim i As Integer

        log.Debug("getIntParameters - " & "reading Parameter")
        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Try
            aPws = aWB.Worksheets("Parameter_Int")
        Catch Exc As System.Exception
            MsgBox("No Parameter_Int Sheet in current workbook. Check if the current workbook is a valid SAP CO-OM Template",
                   MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAP CO-OM")
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

    Public Sub exec(ByRef pSapCon As SapCon, Optional pMode As String = "Create")

        Dim aPar As New SAPCommon.TStr
        Dim aIntPar As New SAPCommon.TStr

        Dim aWB As Excel.Workbook
        Dim aDws As Excel.Worksheet
        Dim aData As Collection

        Dim aRetStr As String
        ' get general parameters
        If getGenParameters(aPar) = False Then
            Exit Sub
        End If
        ' get internal parameters
        If Not getIntParameters(aIntPar) Then
            Exit Sub
        End If

        Dim aSAPUnitCosting As New SapUnitCosting(pSapCon, aIntPar)

        aWB = Globals.SapCoPlExcelAddin.Application.ActiveWorkbook
        Dim aDwsName As String = If(aIntPar.value("WS", "DATA") <> "", aIntPar.value("WS", "DATA"), "Costing")
        Try
            aDws = aWB.Worksheets(aDwsName)
        Catch Exc As System.Exception
            MsgBox("No " & aDwsName & " Sheet in current workbook. Check if the current workbook is a valid SAP Costing Template",
                       MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Pl")
            Exit Sub
        End Try
        ' Read the Items
        Try
            log.Debug("SapCoPlRibbonCosting.exec - " & "processing data - disabling events, screen update, cursor")
            aDws.Activate()
            Dim aItems As New TData(aIntPar)
            Dim aItem As New TDataRec
            Dim aKey As String
            Dim j As UInt64
            Dim jMax As UInt64 = 0
            Dim aObjNr As UInt64 = 0
            Dim aLOff As Integer = If(aIntPar.value("LOFF", "DATA") <> "", CInt(aIntPar.value("LOFF", "DATA")), 4)
            Dim aDumpObjNr As UInt64 = If(aIntPar.value("DBG", "DUMPOBJNR") <> "", CInt(aIntPar.value("DBG", "DUMPOBJNR")), 0)
            Dim aMsgClmn As String = If(aIntPar.value("COL", "DATAMSG") <> "", aIntPar.value("COL", "DATAMSG"), "INT-MSG")
            Dim aMsgClmnNr As Integer = 0
            Dim aPostClmn As String = If(aIntPar.value("COL", "DATAPOST") <> "", aIntPar.value("COL", "DATAPOST"), "INT-POST")
            Dim aPostClmnNr As Integer = 0
            Dim aOKMsg As String = If(aIntPar.value("RET", "OKMSG") <> "", aIntPar.value("RET", "OKMSG"), "OK")

            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait
            Globals.SapCoPlExcelAddin.Application.EnableEvents = False
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = False
            Dim i As UInt64 = aLOff + 1
            ' determine the last column and create the fieldlist
            Do
                jMax += 1
                If CStr(aDws.Cells(1, jMax).value) = aMsgClmn Then
                    aMsgClmnNr = jMax
                ElseIf CStr(aDws.Cells(1, jMax).value) = aPostClmn Then
                    aPostClmnNr = jMax
                End If
            Loop While CStr(aDws.Cells(aLOff - 3, jMax + 1).value) <> ""
            Dim aPost As String = ""
            aData = New Collection
            j = 1
            Do
                aObjNr += 1
                ' Unit costings are handled in packages based on the posting indicator
                If Left(CStr(aDws.Cells(i, aMsgClmnNr).Value), Len(aOKMsg)) <> aOKMsg Then
                    If aPostClmnNr <> 0 Then
                        If String.IsNullOrEmpty(CStr(aDws.Cells(i, aPostClmnNr).value)) Then
                            aPost = ""
                        Else
                            aPost = CStr(aDws.Cells(i, aPostClmnNr).value)
                        End If
                    End If
                    aKey = CStr(i)
                    For j = 1 To jMax
                        If CStr(aDws.Cells(1, j).value) <> "N/A" And CStr(aDws.Cells(1, j).value) <> "" And CStr(aDws.Cells(1, j).value) <> aMsgClmn Then
                            aItems.addValue(aKey, CStr(aDws.Cells(aLOff - 3, j).value), CStr(aDws.Cells(i, j).value),
                                    CStr(aDws.Cells(aLOff - 2, j).value), CStr(aDws.Cells(aLOff - 1, j).value), pEmty:=False,
                                    pEmptyChar:="")
                        End If
                    Next
                    ' aItem = aItems.aTDataDic(aKey)
                    ' if the posting indicator is set, or this is the last line -> call the sap BAPI
                    If String.IsNullOrEmpty(CStr(aDws.Cells(i + 1, 1).value)) Or aPost.ToUpper = "X" Then
                        Dim aTSAP_CostingData As New TSAP_CostingData(aPar, aIntPar)
                        If aTSAP_CostingData.fillHeader(aItems) And aTSAP_CostingData.fillData(aItems) Then
                            ' check if we should dump this document
                            If aObjNr = aDumpObjNr Then
                                log.Debug("SapCoPlRibbonCosting.exec - " & "dumping Object Nr " & CStr(aObjNr))
                                aTSAP_CostingData.dumpHeader()
                                aTSAP_CostingData.dumpData()
                            End If
                            ' post the object here
                            If pMode = "Create" Then
                                log.Debug("SapCoPlRibbonCosting.exec - " & "calling aSAPUnitCosting.createSingle")
                                aRetStr = aSAPUnitCosting.createSingle(aTSAP_CostingData, aOKMsg)
                                log.Debug("SapCoPlRibbonCosting.exec - " & "aSAPUnitCosting.createSingle returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            ElseIf pMode = "Change" Then
                                log.Debug("SapCoPlRibbonCosting.exec - " & "calling aSAPUnitCostingt.changeSingle")
                                aRetStr = aSAPUnitCosting.changeSingle(aTSAP_CostingData, aOKMsg)
                                log.Debug("SapCoPlRibbonCosting.exec - " & "aSAPUnitCosting.changeSingle returned, aRetStr=" & aRetStr)
                                ' message has to be written in all lines that where processed in items
                                For Each aKey In aItems.aTDataDic.Keys
                                    aDws.Cells(CInt(aKey), aMsgClmnNr) = CStr(aRetStr)
                                Next
                            End If
                        Else
                            log.Warn("SapCoPlRibbonCosting.exec - " & "filling Header or Data in aTSAP_CostingData failed!")
                            aDws.Cells(i, aMsgClmnNr) = "Filling Header or Data in aTSAP_CostingData failed!"
                        End If
                        aItems = New TData(aIntPar)
                    End If
                Else
                    aDws.Cells(i, aMsgClmnNr + 1).Value = "ignored - already processed"
                End If
                i += 1
            Loop While CStr(aDws.Cells(i, 1).Value) <> ""
            log.Debug("SapCoPlRibbonCosting.exec - " & "all data processed - enabling events, screen update, cursor")
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
        Catch ex As System.Exception
            Globals.SapCoPlExcelAddin.Application.EnableEvents = True
            Globals.SapCoPlExcelAddin.Application.ScreenUpdating = True
            Globals.SapCoPlExcelAddin.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
            MsgBox("SapCoPlRibbonCosting.exec failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Sap CO Pl")
            log.Error("SapCoPlRibbonCosting.exec - " & "Exception=" & ex.ToString)
            Exit Sub
        End Try
    End Sub

End Class
