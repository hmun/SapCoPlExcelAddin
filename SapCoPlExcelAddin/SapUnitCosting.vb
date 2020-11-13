' Copyright 2020 Hermann Mundprecht
' This file is licensed under the terms of the license 'CC BY 4.0'. 
' For a human readable version of the license, see https://creativecommons.org/licenses/by/4.0/

Imports SAP.Middleware.Connector
Public Class SapUnitCosting
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private aIntPar As SAPCommon.TStr

    Sub New(aSapCon As SapCon, ByRef pIntPar As SAPCommon.TStr)
        aIntPar = pIntPar
        Try
            sapcon = aSapCon
            aSapCon.getDestination(destination)
            sapcon.checkCon()
        Catch ex As System.Exception
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPUnitCosting")
        End Try
    End Sub

    Public Function createSingle(pData As TSAP_CostingData, Optional pOKMsg As String = "OK") As String
        createSingle = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("API_UNITCOSTING_CREATE")
            RfcSessionManager.BeginContext(destination)
            Dim oEX_MESSAGES As IRfcTable = oRfcFunction.GetTable("EX_MESSAGES")
            Dim oIM_POSITIONS As IRfcTable = oRfcFunction.GetTable("IM_POSITIONS")
            oEX_MESSAGES.Clear()
            oIM_POSITIONS.Clear()

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
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIM_POSITIONSAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IM_POSITIONS"
                            If Not oIM_POSITIONSAppended Then
                                oIM_POSITIONS.Append()
                                oIM_POSITIONSAppended = True
                            End If
                            oIM_POSITIONS.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oEX_MESSAGES.Count - 1
                createSingle = createSingle & "; " & oEX_MESSAGES(i).GetValue("MSGTYP") & "-" & oEX_MESSAGES(i).GetValue("MSGID") & "-" & oEX_MESSAGES(i).GetValue("MSGNR")
                If oEX_MESSAGES(i).GetValue("MSGTYP") <> "S" And oEX_MESSAGES(i).GetValue("MSGTYP") <> "I" And oEX_MESSAGES(i).GetValue("MSGTYP") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            createSingle = If(createSingle = "", pOKMsg, If(aErr = False, pOKMsg & createSingle, "Error" & createSingle))
        Catch SapEx As SAP.Middleware.Connector.RfcAbapException
            createSingle = "Error: Exception in createSingle: " & SapEx.Message & " Type: " & SapEx.AbapMessageType & " Class: " & SapEx.AbapMessageClass & " MessageNumber: " & SapEx.AbapMessageNumber & " "
            For Each param As String In SapEx.AbapMessageParameters
                createSingle = createSingle & ";" & param
            Next
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPUnitCosting")
            createSingle = "Error: Exception in createSingle"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

    Public Function changeSingle(pData As TSAP_CostingData, Optional pOKMsg As String = "OK") As String
        changeSingle = ""
        Try
            oRfcFunction = destination.Repository.CreateFunction("API_UNITCOSTING_CHANGE")
            Dim oEX_MESSAGES As IRfcTable = oRfcFunction.GetTable("EX_MESSAGES")
            Dim oIM_POSITIONS As IRfcTable = oRfcFunction.GetTable("IM_POSITIONS")
            oEX_MESSAGES.Clear()
            oIM_POSITIONS.Clear()

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
            ' set the data values
            Dim aKvP As KeyValuePair(Of String, TDataRec)
            Dim aTDataRec As TDataRec
            For Each aKvP In pData.aData.aTDataDic
                Dim oIM_POSITIONSAppended As Boolean = False
                aTDataRec = aKvP.Value
                For Each aTStrRec In aTDataRec.aTDataRecCol
                    Select Case aTStrRec.Strucname
                        Case "IM_POSITIONS"
                            If Not oIM_POSITIONSAppended Then
                                oIM_POSITIONS.Append()
                                oIM_POSITIONSAppended = True
                            End If
                            oIM_POSITIONS.SetValue(aTStrRec.Fieldname, aTStrRec.formated())
                    End Select
                Next
            Next
            ' call the BAPI
            oRfcFunction.Invoke(destination)
            Dim aErr As Boolean
            aErr = False
            For i As Integer = 0 To oEX_MESSAGES.Count - 1
                changeSingle = changeSingle & "; " & oEX_MESSAGES(i).GetValue("MSGTYP") & "-" & oEX_MESSAGES(i).GetValue("MSGID") & "-" & oEX_MESSAGES(i).GetValue("MSGNR")
                If oEX_MESSAGES(i).GetValue("MSGTYP") <> "S" And oEX_MESSAGES(i).GetValue("MSGTYP") <> "I" And oEX_MESSAGES(i).GetValue("MSGTYP") <> "W" Then
                    aErr = True
                End If
            Next i
            If aErr = False Then
                Dim aSAPBapiTranctionCommit As New SAPBapiTranctionCommit(sapcon)
                aSAPBapiTranctionCommit.commit(pWait:="X")
            End If
            changeSingle = If(changeSingle = "", pOKMsg, If(aErr = False, pOKMsg & changeSingle, "Error" & changeSingle))
        Catch SapEx As SAP.Middleware.Connector.RfcAbapException
            changeSingle = "Error: Exception in changeSingle: " & SapEx.Message & " Type: " & SapEx.AbapMessageType & " Class: " & SapEx.AbapMessageClass & " MessageNumber: " & SapEx.AbapMessageNumber & " "
            For Each param As String In SapEx.AbapMessageParameters
                changeSingle = changeSingle & ";" & param
            Next
        Catch Ex As System.Exception
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "SAPUnitCosting")
            changeSingle = "Error: Exception in changeSingle"
        Finally
            RfcSessionManager.EndContext(destination)
        End Try
    End Function

End Class
